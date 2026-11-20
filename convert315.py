"""3.15 — converting any kind of file: documents, slides, spreadsheets, pictures, e-books, audio and video.

The pairs Academic-Home converted before (Word, PowerPoint, PDF and Excel to each other) keep their own converters in
converters.py; everything else is here, done with what the server already has (LibreOffice, Poppler, Tesseract, Pillow,
openpyxl, pdfplumber) and FFmpeg for sound and video. One table (FORMATS, PAIRS) says what turns into what: the app is
built from the same table, so it never offers a pair the server can't do.

A result with more than one file (a PDF's pages as pictures, a workbook's sheets as CSV) comes back as one .zip.
"""
import csv
import html as _html
import io
import json
import os
import re
import shutil
import signal
import subprocess
import zipfile

try:
    from converters import ConversionError, _soffice_slot, _lo_profile_dir
except ImportError:          # the app's build reads the table below without the server's modules
    class ConversionError(Exception):
        pass
    _soffice_slot = _lo_profile_dir = None

# code: (name people know, kind, the file endings it comes with)
FORMATS = {
    "docx": ("Word", "docs", ["docx"]),
    "doc": ("Word 97-2003", "docs", ["doc"]),
    "odt": ("OpenDocument Text", "docs", ["odt"]),
    "rtf": ("Rich Text", "docs", ["rtf"]),
    "txt": ("Text", "docs", ["txt"]),
    "md": ("Markdown", "docs", ["md", "markdown"]),
    "html": ("Web page", "docs", ["html", "htm"]),
    "epub": ("E-book", "docs", ["epub"]),
    "pdf": ("PDF", "docs", ["pdf"]),
    "pptx": ("PowerPoint", "slides", ["pptx"]),
    "ppt": ("PowerPoint 97-2003", "slides", ["ppt"]),
    "odp": ("OpenDocument Slides", "slides", ["odp"]),
    "xlsx": ("Excel", "sheets", ["xlsx"]),
    "xls": ("Excel 97-2003", "sheets", ["xls"]),
    "ods": ("OpenDocument Sheet", "sheets", ["ods"]),
    "csv": ("CSV", "sheets", ["csv"]),
    "json": ("JSON", "sheets", ["json"]),
    "jpg": ("JPG", "images", ["jpg", "jpeg"]),
    "png": ("PNG", "images", ["png"]),
    "webp": ("WebP", "images", ["webp"]),
    "gif": ("GIF", "images", ["gif"]),
    "heic": ("HEIC (iPhone photo)", "images", ["heic", "heif"]),
    "avif": ("AVIF", "images", ["avif"]),
    "svg": ("SVG", "images", ["svg"]),
    "tiff": ("TIFF", "images", ["tiff", "tif"]),
    "bmp": ("BMP", "images", ["bmp"]),
    "ico": ("Icon", "images", ["ico"]),
    "mp3": ("MP3", "audio", ["mp3"]),
    "wav": ("WAV", "audio", ["wav"]),
    "m4a": ("M4A", "audio", ["m4a"]),
    "aac": ("AAC", "audio", ["aac"]),
    "ogg": ("OGG", "audio", ["ogg", "oga"]),
    "opus": ("Opus", "audio", ["opus"]),
    "flac": ("FLAC", "audio", ["flac"]),
    "mp4": ("MP4", "video", ["mp4", "m4v"]),
    "mov": ("MOV", "video", ["mov"]),
    "webm": ("WebM", "video", ["webm"]),
    "mkv": ("MKV", "video", ["mkv"]),
    "avi": ("AVI", "video", ["avi"]),
}
# the pairs converters.py has always done (kept there, with their own options)
OLD_PAIRS = {("docx", "pdf"), ("docx", "pptx"), ("pptx", "docx"), ("pdf", "docx"), ("pdf", "pptx"), ("xlsx", "docx"),
             ("docx", "xlsx"), ("pptx", "pptx"), ("pptx", "pdf"), ("xlsx", "pdf")}

_WRITER_IN = ["docx", "doc", "odt", "rtf", "txt", "md", "html", "epub"]
_WRITER_OUT = ["docx", "pdf", "doc", "odt", "rtf", "txt", "md", "html", "epub"]
_IMAGES_IN = ["jpg", "png", "webp", "gif", "heic", "avif", "svg", "tiff", "bmp", "ico"]
_IMAGES_OUT = ["jpg", "png", "webp", "gif", "avif", "tiff", "bmp", "ico", "pdf"]
_AUDIO = ["mp3", "wav", "m4a", "aac", "ogg", "opus", "flac"]
_VIDEO = ["mp4", "mov", "webm", "mkv", "avi"]


def _pairs():
    p = {}

    def add(a, bs):
        for b in bs:
            if b != a:
                p.setdefault(a, [])
                if b not in p[a]:
                    p[a].append(b)
    for a in _WRITER_IN:
        add(a, _WRITER_OUT + ["jpg", "png"])
    add("docx", ["pptx", "xlsx"])
    add("pdf", ["docx", "pptx", "xlsx", "txt", "md", "html", "epub", "odt", "rtf", "jpg", "png"])
    for a in ("pptx", "ppt", "odp"):
        add(a, ["pptx", "pdf", "docx", "ppt", "odp", "jpg", "png"])
    for a in ("xlsx", "xls", "ods", "csv"):
        add(a, ["xlsx", "pdf", "xls", "ods", "csv", "html", "json"])
    add("xlsx", ["docx"])
    for a in _IMAGES_IN:
        add(a, _IMAGES_OUT + (["txt", "docx"] if a != "ico" else []))     # an icon has no words worth reading
    for a in _AUDIO:
        add(a, _AUDIO)
    for a in _VIDEO:
        add(a, ["mp4", "webm", "gif", "mp3", "m4a", "wav"])
    return p


PAIRS = _pairs()
PAIRS["pptx"] = ["pdf", "docx", "pptx"] + [x for x in PAIRS["pptx"] if x not in ("pdf", "docx", "pptx")]   # Fix slides stays
ALIASES = {e: code for code, (_, _, exts) in FORMATS.items() for e in exts}
MIME = {"docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "doc": "application/msword",
        "odt": "application/vnd.oasis.opendocument.text", "rtf": "application/rtf", "txt": "text/plain; charset=utf-8",
        "md": "text/markdown; charset=utf-8", "html": "text/html; charset=utf-8", "epub": "application/epub+zip", "pdf": "application/pdf",
        "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation", "ppt": "application/vnd.ms-powerpoint",
        "odp": "application/vnd.oasis.opendocument.presentation", "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "xls": "application/vnd.ms-excel", "ods": "application/vnd.oasis.opendocument.spreadsheet", "csv": "text/csv; charset=utf-8",
        "json": "application/json", "jpg": "image/jpeg", "png": "image/png", "webp": "image/webp", "gif": "image/gif", "avif": "image/avif",
        "tiff": "image/tiff", "bmp": "image/bmp", "ico": "image/x-icon", "mp3": "audio/mpeg", "wav": "audio/wav", "m4a": "audio/mp4",
        "aac": "audio/aac", "ogg": "audio/ogg", "opus": "audio/ogg", "flac": "audio/flac", "mp4": "video/mp4", "webm": "video/webm",
        "zip": "application/zip"}


def can(a, b):
    return b in PAIRS.get(a, [])


def catalog():
    """For the app: every format (its name and kind) and what each turns into."""
    return {"formats": {k: {"name": v[0], "kind": v[1]} for k, v in FORMATS.items()}, "pairs": PAIRS}


# ---------------------------------------------------------------- the tools
def _run(args, timeout=200, what="the converter"):
    proc = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, start_new_session=True)
    try:
        out, err = proc.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        try:
            os.killpg(os.getpgid(proc.pid), signal.SIGKILL)
        except ProcessLookupError:
            pass
        proc.communicate()
        raise ConversionError(f"That file took too long for {what}. A shorter or smaller file will work.")
    if proc.returncode != 0:
        raise ConversionError(f"{what[0].upper() + what[1:]} couldn't read this file. It may be damaged, or not really the kind of file its name says.")
    return out


def _lo(src, ext, out_dir, filt=None, infilter=None):
    """LibreOffice, with a filter of our choosing (one conversion at a time, in its own profile)."""
    slot = _soffice_slot()
    profile = _lo_profile_dir()
    try:
        os.makedirs(out_dir, exist_ok=True)
        args = ["soffice", "--headless", "--nologo", "--nofirststartwizard", f"-env:UserInstallation=file://{profile}"]
        if infilter:
            args.append(f"--infilter={infilter}")
        args += ["--convert-to", ext + (":" + filt if filt else ""), "--outdir", out_dir, src]
        try:
            _run(args, timeout=180, what="the document converter")
        except ConversionError:
            shutil.rmtree(profile, ignore_errors=True)
            raise
    finally:
        slot.close()
    out = os.path.join(out_dir, os.path.splitext(os.path.basename(src))[0] + "." + ext)
    if not os.path.exists(out) or os.path.getsize(out) == 0:
        raise ConversionError("The document converter couldn't make that file. The original may be damaged.")
    return out


_LO_OUT = {"docx": ("docx", "MS Word 2007 XML"), "doc": ("doc", "MS Word 97"), "odt": ("odt", "writer8"), "rtf": ("rtf", "Rich Text Format"),
           "txt": ("txt", "Text (encoded):UTF8"), "html": ("html", "XHTML Writer File:UTF8"), "epub": ("epub", "EPUB"),
           "pdf": ("pdf", "writer_pdf_Export")}
_LO_IN = {"html": "HTML (StarWriter)", "txt": "Text (encoded):UTF8"}


def _named(path, out_dir, stem, ext):
    """The result under the file's own name (inputs are copied under one fixed name while they're worked on)."""
    want = os.path.join(out_dir, stem + "." + ext)
    if os.path.abspath(path) != os.path.abspath(want):
        shutil.move(path, want)
    return want


def _work(out_dir, name):
    d = os.path.join(out_dir, name)
    os.makedirs(d, exist_ok=True)
    return d


# ---------------------------------------------------------------- documents
def _md_to_html(text):
    import markdown
    body = markdown.markdown(text, extensions=["tables", "fenced_code", "sane_lists"])
    return _page(body)


def _page(body, title=""):
    # 6.4: the page is given no title of ours (it said "Document", an e-book's said "Book"): the owner, 19 Nov 2026,
    # "never add anything I mean anything at all to converted files"
    return ("<!DOCTYPE html><html><head><meta charset=\"utf-8\"><title>" + _html.escape(title) + "</title></head><body>" + body + "</body></html>")


_PIC_MIME = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif", "svg": "image/svg+xml", "webp": "image/webp", "bmp": "image/bmp"}


def _inline_pictures(body, base_dir, root):
    """6.4: the book's pictures are written into the page itself. Left as pointers to files, they were pointers in the
    Word file too (LibreOffice links a web page's pictures, it does not take them in), and the person got a Word file,
    a web page or a Markdown file whose pictures were gone: the files they pointed to never left the server."""
    import base64
    import urllib.parse

    def one(m):
        src = urllib.parse.unquote(m.group(2).split("#")[0].split("?")[0])
        p = os.path.normpath(os.path.join(base_dir, src))
        ext = os.path.splitext(p)[1].lower().lstrip(".")
        if not p.startswith(root) or not os.path.isfile(p) or ext not in _PIC_MIME or os.path.getsize(p) > 12 * 1024 * 1024:
            return m.group(0)
        with open(p, "rb") as fh:
            return m.group(1) + "data:" + _PIC_MIME[ext] + ";base64," + base64.b64encode(fh.read()).decode("ascii") + '"'
    return re.sub(r'(<img\b[^>]*\bsrc=")(?!https?:|data:)([^"]+)"', one, body)


def _epub_to_html(src, work):
    """An e-book's chapters, in reading order, as one web page (its pictures alongside)."""
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        for n in names:
            if n.endswith("/") or ".." in n.split("/") or n.startswith("/"):
                continue
            z.extract(n, work)
        try:
            cont = z.read("META-INF/container.xml").decode("utf-8", "ignore")
            opf_path = re.search(r'full-path="([^"]+)"', cont).group(1)
        except Exception:
            opf_path = next((n for n in names if n.endswith(".opf")), None)
        if not opf_path:
            raise ConversionError("This e-book has no table of its chapters, so it can't be read.")
        opf = z.read(opf_path).decode("utf-8", "ignore")
    base = os.path.dirname(opf_path)
    items = {m.group(1): m.group(2) for m in re.finditer(r'<item\b[^>]*\bid="([^"]+)"[^>]*\bhref="([^"]+)"', opf)}
    items.update({m.group(2): m.group(1) for m in re.finditer(r'<item\b[^>]*\bhref="([^"]+)"[^>]*\bid="([^"]+)"', opf)})
    order = [items[m.group(1)] for m in re.finditer(r'<itemref\b[^>]*\bidref="([^"]+)"', opf) if m.group(1) in items]
    parts = []
    for href in order:
        p = os.path.normpath(os.path.join(work, base, href))
        if not p.startswith(os.path.abspath(work)) or not os.path.isfile(p):
            continue
        t = open(p, encoding="utf-8", errors="ignore").read()
        m = re.search(r"<body[^>]*>(.*)</body>", t, re.S | re.I)
        body = m.group(1) if m else t
        rel = os.path.relpath(os.path.dirname(p), os.path.join(work, base))
        if rel != ".":   # pictures relative to the chapter, from where the page is written
            body = re.sub(r'(<img\b[^>]*\bsrc=")(?!https?:|data:)([^"]+)"', lambda mm: mm.group(1) + os.path.normpath(os.path.join(rel, mm.group(2))) + '"', body)
        body = _inline_pictures(body, os.path.join(work, base), os.path.abspath(work))
        parts.append(body)
    if not parts:
        raise ConversionError("This e-book's chapters couldn't be read.")
    out = os.path.join(work, base, "book.html")
    m = re.search(r"<dc:title[^>]*>([^<]*)</dc:title>", opf)
    title = _html.unescape(m.group(1)).strip() if m else ""
    if title.lower() in ("unknown title", "unknown"):
        title = ""
    # 6.4: a chapter starts on a new page, as it does in the book (a line used to be drawn between chapters: a line the
    # book does not have). The mark goes on the chapter's first heading, paragraph, list or table: LibreOffice does not
    # read it from a box round the chapter.
    def new_page(chapter):
        def mark(m):
            tag, rest = m.group(1), m.group(2)
            if re.search(r'\bstyle="', rest):
                return "<" + tag + re.sub(r'\bstyle="', 'style="page-break-before: always; ', rest, count=1) + ">"
            return "<" + tag + ' style="page-break-before: always"' + rest + ">"
        return re.sub(r"<(h[1-6]|p|ul|ol|table|blockquote|pre)\b([^>]*)>", mark, chapter, count=1, flags=re.I)
    body = parts[0] + "".join("\n" + new_page(p) for p in parts[1:])
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(_page(body, title))
    return out


def _html_to_md(html_text):
    from markdownify import markdownify
    m = re.search(r"<body[^>]*>(.*)</body>", html_text, re.S | re.I)      # the page itself (not its head or XML line)
    try:        # (a table's first row is its header, as it is in the file: no empty row of ours above it)
        md = markdownify(m.group(1) if m else html_text, heading_style="ATX", bullets="-", table_infer_header=True)
    except TypeError:
        md = markdownify(m.group(1) if m else html_text, heading_style="ATX", bullets="-")
    # a table needs its divider under the first row
    lines, out = md.split("\n"), []
    for k, line in enumerate(lines):
        out.append(line)
        is_row = line.strip().startswith("|") and line.strip().endswith("|")
        first = is_row and not (k and lines[k - 1].strip().startswith("|"))
        nxt = lines[k + 1].strip() if k + 1 < len(lines) else ""
        if first and not re.match(r"^\|?\s*:?-{3,}", nxt):
            out.append("|" + "|".join(" --- " for _ in range(max(1, line.strip().strip("|").count("|") + 1))) + "|")
    md = re.sub(r"\n{3,}", "\n\n", "\n".join(out))
    return md.strip() + "\n"


# 6.4: Markdown is made from LibreOffice's plain web page, which says a heading with <h1>, bold with <b> and a list
# with <ol><li>. (Its XHTML page, used until now, says all of these with CSS classes that Markdown cannot read: bold and
# italic were lost, a numbered list came out as "1. 1.Mark the plot corners", and a heading that was not Word's own
# heading style was a plain line.)
_MD_FILTER = "HTML (StarWriter):EmbedImages"
_MD_BULLETS = "\u2022\u25e6\u25aa\u25ab\u25cf\u25cb\u25a0\u25a1\u2023\u2043\u00b7\uf0b7\uf0a7"


def _docx_base_size(path):
    """The size (points) a Word file's plain text is written in: its Normal style's, or the document's default."""
    try:
        with zipfile.ZipFile(path) as z:
            st = z.read("word/styles.xml").decode("utf-8", "ignore")
        for m in re.finditer(r"<w:style\b[^>]*>.*?</w:style>", st, re.S):
            head = m.group(0)[:m.group(0).index(">") + 1]
            if 'w:default="1"' in head and 'w:type="paragraph"' in head:
                z_ = re.search(r'<w:sz w:val="(\d+)"', m.group(0))
                if z_:
                    return int(z_.group(1)) / 2.0
        m = re.search(r'<w:docDefaults>.*?<w:sz w:val="(\d+)".*?</w:docDefaults>', st, re.S)
        return int(m.group(1)) / 2.0 if m else 10.0
    except Exception:
        return None


def _md_ready(html_text, base=None):
    """LibreOffice's plain web page, readied for Markdown. Nothing is added and no word is changed:
      - the page's own line ends inside a paragraph are not the document's (they made hard-wrapped Markdown);
      - page numbers and page counts go (a Markdown file has no pages), and with them a page head or foot that held
        nothing else;
      - a line that only LOOKS like a heading (all of it bold and bigger than the plain text: a PDF has no heading
        styles, and many Word files are written this way) is a heading, as it is to the reader's eye; only in a
        document that has no real headings of its own;
      - a paragraph that begins with a bullet sign is a list item (a PDF's bullets are signs typed before the words).
    base: the size, in points, of the document's plain text, when it is known."""
    from bs4 import BeautifulSoup, NavigableString
    m = re.search(r"<body[^>]*>(.*)</body>", html_text, re.S | re.I)
    body = m.group(1) if m else html_text
    parts = re.split(r"(<pre\b.*?</pre>)", body, flags=re.S | re.I)
    body = "".join(p if k % 2 else re.sub(r"[ \t]*\r?\n[ \t]*", " ", p) for k, p in enumerate(parts))
    soup = BeautifulSoup(body, "html.parser")
    for f in soup.find_all("sdfield"):
        if (f.get("type") or "").upper() in ("PAGE", "DOCSTAT"):
            f.decompose()
        else:
            f.unwrap()
    for d in soup.find_all("div", attrs={"title": re.compile(r"^(header|footer)$", re.I)}):
        if not re.search(r"\w", re.sub(r"(?i)\b(page|of)\b", "", d.get_text(" "))):
            d.decompose()
        else:
            d.unwrap()

    def size_of(node):                                   # the size a piece of text is written in, where it is said
        el = node.parent
        while el is not None and getattr(el, "name", None) not in (None, "p", "td", "th", "li", "body", "[document]"):
            z = re.search(r"font-size:\s*([0-9.]+)pt", el.get("style") or "")
            if z:
                return float(z.group(1))
            el = el.parent
        return None

    def bold(node):
        el = node.parent
        while el is not None and getattr(el, "name", None) not in (None, "p", "td", "th", "li", "body", "[document]"):
            if el.name in ("b", "strong") or re.search(r"font-weight:\s*(bold|[6-9]00)", el.get("style") or ""):
                return True
            el = el.parent
        return False

    paras = [q for q in soup.find_all("p") if q.find_parent(["table", "li", "ul", "ol", "blockquote"]) is None]
    # ---- bullets typed as signs
    run = None
    for q in paras:
        text = q.get_text()
        z = re.match(r"^\s*[" + _MD_BULLETS + r"][\s\u00a0]+(?=\S)", text)
        if not z:
            run = None
            continue
        left = len(z.group(0))
        for t in list(q.find_all(string=True)):
            if left <= 0:
                break
            cut = min(left, len(t))
            left -= cut
            t.replace_with(NavigableString(t[cut:]))
        q.name = "li"
        q.attrs = {}
        nxt = run.next_sibling if run is not None else None
        while isinstance(nxt, NavigableString) and not nxt.strip():
            nxt = nxt.next_sibling
        if run is not None and nxt is q:
            run.append(q.extract())
        else:
            run = soup.new_tag("ul")
            q.replace_with(run)
            run.append(q)
    # ---- lines that look like headings
    if not soup.find(re.compile(r"^h[1-6]$")):
        if base is None:
            tally = {}
            for q in paras:
                for t in q.find_all(string=True):
                    if t.strip():
                        k = size_of(t)
                        tally[k] = tally.get(k, 0) + len(t.strip())
            best = max(tally, key=tally.get) if tally else None
            base = best if best is not None else 12.0
        cands = []
        for q in paras:
            if q.name != "p" or q.find("img") is not None:
                continue
            texts = [t for t in q.find_all(string=True) if t.strip()]
            words = " ".join(t.strip() for t in texts)
            if not texts or len(words) > 150:
                continue
            sizes = [size_of(t) for t in texts]
            if any(z is None for z in sizes) or len(set(sizes)) != 1:
                continue
            z = sizes[0]
            if (all(bold(t) for t in texts) and z >= base + 1.0) or z >= base + 5.0:
                cands.append((q, z))
        levels = sorted(set(z for _, z in cands), reverse=True)
        for q, z in cands:
            q.name = "h%d" % min(3, levels.index(z) + 1)
            q.attrs = {}
            for b in q.find_all(["b", "strong"]):
                b.unwrap()                                # (a heading is not written **bold** as well)
    return str(soup)


def _writer_source(src, a, work):
    """A file LibreOffice Writer can open, and the filter to open it with."""
    if a == "md":
        p = os.path.join(work, "doc.html")
        open(p, "w", encoding="utf-8").write(_md_to_html(open(src, encoding="utf-8", errors="ignore").read()))
        return p, _LO_IN["html"]
    if a == "epub":
        return _epub_to_html(src, _work(work, "book")), _LO_IN["html"]
    return src, _LO_IN.get(a)


PAGES_MAX = 60            # pages made into pictures at a time
OCR_MAX = 20              # scanned pages whose words are read at a time...
OCR_SECONDS = 110         # ...and the time their reading may take
TABLE_PAGES_MAX = 80      # pages read for their tables at a time


def _say(text):
    """A word for the person about the file that was made (shown beside it in the app, never written into it)."""
    try:
        from converters import _note
        _note(text)
    except Exception:
        pass


def _pdf_count(pdf):
    try:
        from pypdf import PdfReader
        rd = PdfReader(pdf)
        if rd.is_encrypted:
            rd.decrypt("")
        return len(rd.pages)
    except Exception:
        return None


def _pdf_pages(pdf, b, out_dir, stem, own=True):
    """A PDF's pages as pictures: one picture, or a .zip of them. own: the PDF is the person's own file (not one made
    here on the way from a document or slides).
    6.4: a file with more than 60 pages gave its first 60 and said nothing of the rest. Now it is said, before any work."""
    n = _pdf_count(pdf)
    if n is not None and n > PAGES_MAX:
        raise ConversionError(("This PDF has %d pages. Up to %d can be made into pictures at a time: split it first (PDF tools, "
                               "Split a PDF), then convert each part." % (n, PAGES_MAX)) if own else
                              ("This file has %d pages. Up to %d can be made into pictures at a time." % (n, PAGES_MAX)))
    work = _work(out_dir, "pages")
    fmt = "-jpeg" if b == "jpg" else "-png"
    _run(["pdftoppm", "-r", "150", fmt, "-l", str(PAGES_MAX), pdf, os.path.join(work, "p")], what="the page renderer")
    pics = sorted(f for f in os.listdir(work) if f.endswith("." + ("jpg" if b == "jpg" else "png")))
    if not pics:
        raise ConversionError("No pages came out of that file.")
    if len(pics) == 1:
        return _named(os.path.join(work, pics[0]), out_dir, stem, b)
    out = os.path.join(out_dir, stem + ".zip")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for k, f in enumerate(pics, 1):
            z.write(os.path.join(work, f), f"{stem} page {k}.{b}")
    return out


def _pdf_text(src, out_dir):
    """The words of a PDF, page after page. A page that is a scan (a picture of a page) is read with OCR.
    6.4: before, only a PDF with no words AT ALL was read with OCR, and then its first 15 pages, silently: a PDF of
    typed pages with a few scanned ones lost the scanned pages' words without a word, and so did page 16 onwards of a
    scan. Now every page without words is looked at; what was read from pictures, and what could not be read, is said
    beside the file. The page-end sign of the PDF reader (a form feed, shown as a box by many phones) is not left in."""
    import time as _time
    txt = os.path.join(out_dir, "words.txt")
    _run(["pdftotext", "-enc", "UTF-8", src, txt], what="the PDF reader")
    raw = open(txt, encoding="utf-8", errors="ignore").read()
    pages = raw.split("\f")
    if pages and not pages[-1].strip():
        pages.pop()
    empty = [k for k, t in enumerate(pages) if len(t.strip()) < 20]
    read = unread = 0
    if empty:
        work = _work(out_dir, "ocr")
        began = _time.time()
        for k in empty:
            if read >= OCR_MAX or _time.time() - began > OCR_SECONDS:
                unread += 1
                continue
            root = os.path.join(work, "p%d" % k)
            try:
                subprocess.run(["pdftoppm", "-r", "200", "-png", "-f", str(k + 1), "-l", str(k + 1), "-singlefile", src, root],
                               capture_output=True, timeout=60)
                r = subprocess.run(["tesseract", root + ".png", "-", "--psm", "3"], capture_output=True, text=True, timeout=60)
                got = (r.stdout or "").strip()
            except Exception:
                got = ""
            finally:
                try:
                    os.remove(root + ".png")
                except OSError:
                    pass
            if len(got) >= 20:
                pages[k] = got + "\n"
                read += 1
    t = "\n".join(p.rstrip("\n") + "\n" for p in pages if p.strip())
    if not t.strip():
        raise ConversionError("No words could be found in this PDF.")
    if read:
        _say("%s scanned: the words were read from the picture, so check them against the original."
             % ("1 page was" if read == 1 else "%d pages were" % read))
    if unread:
        _say("%s could not be read this time: split the PDF (PDF tools, Split a PDF) and convert the rest."
             % ("1 more scanned page" if unread == 1 else "%d more scanned pages" % unread))
    return t


def _gap_rows(page, tops=False):
    """A page's lines, each cut into cells where the gap between two words is wider than a couple of spaces (a table
    drawn without lines keeps its columns; a paragraph's line stays one cell). tops: each line with how far down the
    page it stands, [(top, cells)]."""
    words = page.extract_words(x_tolerance=1.5, y_tolerance=2.5) or []
    lines = []
    for w in sorted(words, key=lambda w: (round(w["top"]), w["x0"])):
        if lines and abs(lines[-1][0] - w["top"]) <= 3:
            lines[-1][1].append(w)
        else:
            lines.append([w["top"], [w]])
    rows = []
    for top, ws in lines:
        ws.sort(key=lambda w: w["x0"])
        cells, cur = [], [ws[0]]
        for prev, w in zip(ws, ws[1:]):
            if w["x0"] - prev["x1"] > max(6.0, (prev["bottom"] - prev["top"]) * 1.2):
                cells.append(cur)
                cur = [w]
            else:
                cur.append(w)
        cells.append(cur)
        rows.append((top, [" ".join(x["text"] for x in c) for c in cells]))
    return rows if tops else [r for _, r in rows]


def _pdf_tables(src, out_dir, stem):
    import pdfplumber
    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)
    wb.properties.creator = ""
    wrote = 0
    with pdfplumber.open(src) as pdf:
        if len(pdf.pages) > TABLE_PAGES_MAX:       # (6.4: page 81 onwards used to be left out without a word)
            raise ConversionError("This PDF has %d pages. Up to %d can be turned into Excel at a time: split it first (PDF tools, "
                                  "Split a PDF), then convert each part." % (len(pdf.pages), TABLE_PAGES_MAX))
        for n, page in enumerate(pdf.pages, 1):
            ws = wb.create_sheet(f"Sheet{n}")          # Excel's own name for a tab (6.4: it was "Page 1": a word of ours)
            # 6.4: the page's tables AND the lines around them, in the page's own order. (Where a page had a table,
            # only the table was written: every heading and paragraph of that page was left out without a word.)
            found = []
            try:
                for tb in page.find_tables() or []:
                    rows = tb.extract() or []
                    if rows and any(any(c for c in row) for row in rows):
                        found.append((tb.bbox, rows))
            except Exception:
                found = []
            rest = page
            if found:
                boxes = [b for b, _ in found]

                def outside(o, boxes=boxes):
                    if o.get("object_type") != "char":
                        return True
                    x, y = (o["x0"] + o["x1"]) / 2.0, (o["top"] + o["bottom"]) / 2.0
                    return not any(b[0] - 1 <= x <= b[2] + 1 and b[1] - 1 <= y <= b[3] + 1 for b in boxes)
                try:
                    rest = page.filter(outside)
                except Exception:
                    rest = None
            blocks = [(b[1], rows) for b, rows in found]
            if rest is not None:
                blocks += [(top, [row]) for top, row in _gap_rows(rest, tops=True)]
            blocks.sort(key=lambda x: x[0])
            r, was_table = 1, False
            for _top, rows in blocks:
                table = len(rows) > 1 or any(rows is fr for _, fr in found)
                if (table or was_table) and r > 1:
                    r += 1                              # an empty row sets a table apart from what is round it
                for row in rows:
                    for c, v in enumerate(row, 1):
                        ws.cell(row=r, column=c, value=_cell(v))
                        wrote += 1 if v not in (None, "") else 0
                    r += 1
                was_table = table
    if not wb.sheetnames:
        raise ConversionError("This PDF has no pages to read.")
    if not wrote:      # (6.4: a scan used to give a workbook of empty sheets, and nothing was said)
        raise ConversionError("No words could be found in this PDF: its pages are pictures (a scan). PDF to Word reads the "
                              "words of a scan; convert it to Word first.")
    out = os.path.join(out_dir, stem + ".xlsx")
    wb.save(out)
    return out


def _cell(v):
    if v is None:
        return None
    # a bullet drawn with the Symbol or Wingdings font is a private sign in the PDF: outside that font it shows as a box
    s = str(v).strip().replace("\uf0b7", "\u2022").replace("\uf0a7", "\u25aa")
    if re.fullmatch(r"-?\d+", s) and len(s) < 15:
        return int(s)
    if re.fullmatch(r"-?\d+\.\d+", s) and len(s) < 18:
        return float(s)
    return s


def convert_doc(src, a, b, out_dir, stem):
    from converters import pdf_to_docx
    if a == "pdf":
        # 6.4: a PDF with a password is said to be one, with the way on (the PDF reader said "it may be damaged")
        try:
            from pdf2word import pdf_lock
            why = pdf_lock(src)
        except ImportError:
            why = None
        if why:
            raise ConversionError(why)
        if b in ("jpg", "png"):
            return _pdf_pages(src, b, out_dir, stem)
        if b == "txt":
            p = os.path.join(out_dir, stem + ".txt")
            open(p, "w", encoding="utf-8").write(_pdf_text(src, out_dir))
            return p
        if b == "xlsx":
            return _pdf_tables(src, out_dir, stem)
        docx = pdf_to_docx(src, _work(out_dir, "pdfdocx"))
        return convert_doc(docx, "docx", b, out_dir, stem) if b != "docx" else _named(docx, out_dir, stem, "docx")
    if b in ("jpg", "png"):
        pdf = convert_doc(src, a, "pdf", _work(out_dir, "topdf"), stem)
        return _pdf_pages(pdf, b, out_dir, stem, own=False)
    work = _work(out_dir, "w")
    if b == "md":
        if a == "html":
            h = open(src, encoding="utf-8", errors="ignore").read()
        elif a == "txt":
            return _named(shutil.copy(src, os.path.join(work, "t.md")), out_dir, stem, "md")
        elif a == "epub":      # the book's own page: its headings, lists and pictures as the book has them
            h = open(_epub_to_html(src, _work(work, "book")), encoding="utf-8", errors="ignore").read()
        else:
            s, inf = _writer_source(src, a, work)
            h = _md_ready(open(_lo(s, "html", work, _MD_FILTER, inf), encoding="utf-8", errors="ignore").read(),
                          _docx_base_size(s) if a == "docx" else None)
        p = os.path.join(out_dir, stem + ".md")
        open(p, "w", encoding="utf-8").write(_html_to_md(h))
        return p
    if a == "md" and b == "html":
        p = os.path.join(out_dir, stem + ".html")
        open(p, "w", encoding="utf-8").write(_md_to_html(open(src, encoding="utf-8", errors="ignore").read()))
        return p
    s, inf = _writer_source(src, a, work)
    ext, filt = _LO_OUT[b]
    if b == "epub":
        s, inf = _epub_ready(s, inf, work), None
    out = _lo(s, ext, work, filt, inf)
    if b == "txt":     # no byte-order mark at the start
        raw = open(out, "rb").read()
        if raw.startswith(b"\xef\xbb\xbf"):
            open(out, "wb").write(raw[3:])
    return _named(out, out_dir, stem, b)


def _epub_ready(src, infilter, work):
    """6.4: LibreOffice's e-book writer LEAVES OUT the heading rows of a table (the rows Word repeats at the top of each
    page: the first row of most tables in a report). The book then had the table's figures and not what they were
    figures of. The document is first written as an OpenDocument file in which those rows are ordinary rows; the
    e-book is made from that. Nothing else in the document is touched."""
    odt_dir = _work(work, "odt")
    odt = _lo(src, "odt", odt_dir, _LO_OUT["odt"][1], infilter)
    try:
        with zipfile.ZipFile(odt) as z:
            items = [(i, z.read(i.filename)) for i in z.infolist()]
        out, changed = [], False
        for info, data in items:
            if info.filename == "content.xml":
                t = data.decode("utf-8")
                u = re.sub(r"</?table:table-header-rows>", "", t)
                if u != t:
                    changed, data = True, u.encode("utf-8")
            out.append((info, data))
        if changed:
            tmp = odt + ".tmp"
            with zipfile.ZipFile(tmp, "w") as z:
                for info, data in out:
                    z.writestr(info, data, compress_type=zipfile.ZIP_STORED if info.filename == "mimetype" else zipfile.ZIP_DEFLATED)
            os.replace(tmp, odt)
    except Exception:
        pass
    return odt


# ---------------------------------------------------------------- slides
def convert_slides(src, a, b, out_dir, stem):
    from converters import pptx_to_docx
    work = _work(out_dir, "s")
    if b in ("jpg", "png"):
        pdf = _lo(src, "pdf", work, "impress_pdf_Export")
        return _pdf_pages(pdf, b, out_dir, stem, own=False)
    if b == "docx":
        # 6.4: pages that look like the slides, as for a .pptx (it was the plain handout: a look of the app's own)
        pptx = src if a == "pptx" else _lo(src, "pptx", work, "Impress MS PowerPoint 2007 XML")
        return _named(pptx_to_docx(pptx, _work(out_dir, "pd"), style="same"), out_dir, stem, "docx")
    filt = {"pptx": "Impress MS PowerPoint 2007 XML", "ppt": "MS PowerPoint 97", "odp": "impress8", "pdf": "impress_pdf_Export"}[b]
    return _named(_lo(src, b, work, filt), out_dir, stem, b)


# ---------------------------------------------------------------- spreadsheets
def _csv_to_xlsx(src, out):
    from openpyxl import Workbook
    raw = open(src, "rb").read()
    text = raw.decode("utf-8-sig", "ignore") if raw[:3] == b"\xef\xbb\xbf" else raw.decode("utf-8", "ignore")
    try:
        dialect = csv.Sniffer().sniff(text[:5000], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for r, row in enumerate(csv.reader(io.StringIO(text), dialect), 1):
        for c, v in enumerate(row, 1):
            ws.cell(row=r, column=c, value=_cell(v))
        if r > 200000:
            break
    wb.properties.creator = ""            # 6.4: not the tool's name ("openpyxl") as the file's author
    wb.save(out)
    return out


def _sheets(xlsx):
    from openpyxl import load_workbook
    wb = load_workbook(xlsx, data_only=True, read_only=True)
    out = []
    for ws in wb.worksheets:
        rows = [[("" if v is None else v) for v in row] for row in ws.iter_rows(values_only=True)]
        while rows and not any(str(v).strip() for v in rows[-1]):
            rows.pop()
        if rows:
            out.append((ws.title, rows))
    return out


def convert_sheet(src, a, b, out_dir, stem):
    from converters import xlsx_to_docx
    work = _work(out_dir, "x")
    xlsx = src
    if a == "csv":
        xlsx = _csv_to_xlsx(src, os.path.join(work, "sheet.xlsx"))
    elif a in ("xls", "ods"):
        xlsx = _lo(src, "xlsx", work, "Calc MS Excel 2007 XML")
    if b == "xlsx":
        return _named(xlsx if xlsx != src else shutil.copy(src, os.path.join(work, "copy.xlsx")), out_dir, stem, "xlsx")
    if b == "docx":
        return _named(xlsx_to_docx(xlsx, _work(out_dir, "xd")), out_dir, stem, "docx")
    if b in ("csv", "json"):
        sheets = _sheets(xlsx)
        if not sheets:
            raise ConversionError("This spreadsheet has nothing in it.")
        files = []
        for title, rows in sheets:
            buf = io.StringIO()
            if b == "csv":
                w = csv.writer(buf)
                for row in rows:
                    w.writerow([str(v) for v in row])
            else:
                head = [str(h).strip() or f"Column {k + 1}" for k, h in enumerate(rows[0])]
                data = [{head[k] if k < len(head) else f"Column {k + 1}": v for k, v in enumerate(row)} for row in rows[1:]]
                buf.write(json.dumps(data, ensure_ascii=False, indent=2, default=str))
            files.append((title, buf.getvalue()))
        if len(files) == 1:
            p = os.path.join(out_dir, stem + "." + b)
            open(p, "w", encoding="utf-8").write(files[0][1])
            return p
        if b == "json":       # every sheet in one file: {"Sheet": [rows…], …}
            p = os.path.join(out_dir, stem + ".json")
            open(p, "w", encoding="utf-8").write("{\n" + ",\n".join(json.dumps(t, ensure_ascii=False) + ": " + d for t, d in files) + "\n}\n")
            return p
        out = os.path.join(out_dir, stem + ".zip")
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            for title, data in files:
                z.writestr(re.sub(r'[\\/:*?"<>|]', "_", f"{stem} - {title}.{b}"), data)
        return out
    filt = {"pdf": "calc_pdf_Export", "xls": "MS Excel 97", "ods": "calc8", "html": "HTML (StarCalc)"}[b]
    return _named(_lo(xlsx, b, work, filt), out_dir, stem, b)


# ---------------------------------------------------------------- pictures
def _open_image(src, a, work):
    from PIL import Image
    if a == "heic":
        try:
            import pillow_heif
            pillow_heif.register_heif_opener()
        except ImportError:
            raise ConversionError("iPhone photos (HEIC) can't be opened on the server yet.")
    if a == "svg":
        png = _lo(src, "png", work, "draw_png_Export")
        return Image.open(png)
    try:
        im = Image.open(src)
        im.load()
        return im
    except Exception:
        raise ConversionError("That picture couldn't be opened. It may be damaged, or not really a picture of that kind.")


def _svg_words(src):
    """The words a drawing (SVG) has as text, line by line (no outside files or entities read)."""
    from lxml import etree
    try:
        root = etree.parse(src, etree.XMLParser(resolve_entities=False, no_network=True, load_dtd=False)).getroot()
    except Exception:
        return ""
    lines = []
    for el in root.iter("{http://www.w3.org/2000/svg}text", "text"):
        t = " ".join("".join(el.itertext()).split())
        if t:
            lines.append(t)
    return "\n".join(lines)


def _flat(im, bg=(255, 255, 255)):
    from PIL import Image
    if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
        im = im.convert("RGBA")
        base = Image.new("RGB", im.size, bg)
        base.paste(im, mask=im.split()[-1])
        return base
    return im.convert("RGB")


def convert_image(src, a, b, out_dir, stem):
    from PIL import Image, ImageOps, ImageSequence
    work = _work(out_dir, "i")
    im = _open_image(src, a, work)
    if b in ("txt", "docx"):
        words = _svg_words(src) if a == "svg" else ""
        if not words:
            png = os.path.join(work, "page.png")
            _flat(ImageOps.exif_transpose(im)).save(png)
            r = subprocess.run(["tesseract", png, "-", "--psm", "3"], capture_output=True, text=True, timeout=90)
            words = (r.stdout or "").strip()
        if not words:
            raise ConversionError("No words could be found in that picture.")
        if b == "txt":
            p = os.path.join(out_dir, stem + ".txt")
            open(p, "w", encoding="utf-8").write(words + "\n")
            return p
        from docx import Document
        d = Document()
        for para in re.split(r"\n\s*\n", words):
            d.add_paragraph(re.sub(r"\s*\n\s*", " ", para).strip())
        p = os.path.join(out_dir, stem + ".docx")
        d.save(p)
        return p
    frames = [f.copy() for f in ImageSequence.Iterator(im)] if getattr(im, "is_animated", False) else [im]
    frames = [ImageOps.exif_transpose(f) for f in frames]
    out = os.path.join(out_dir, stem + "." + b)
    if b == "jpg":
        _flat(frames[0]).save(out, "JPEG", quality=90, optimize=True, progressive=True)
    elif b == "bmp":
        _flat(frames[0]).save(out, "BMP")
    elif b == "png":
        f = frames[0]
        (f if f.mode in ("RGB", "RGBA", "L", "LA", "P") else f.convert("RGBA")).save(out, "PNG", optimize=True)
    elif b == "pdf":
        pages = [_flat(f) for f in frames]
        pages[0].save(out, "PDF", resolution=150.0, save_all=len(pages) > 1, append_images=pages[1:], title="")
    elif b == "ico":
        f = frames[0].convert("RGBA")
        f.thumbnail((256, 256))
        f.save(out, "ICO", sizes=[s for s in [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)] if s[0] <= max(f.size)] or [(16, 16)])
    elif b in ("gif", "webp", "tiff", "avif"):
        fmt = {"gif": "GIF", "webp": "WEBP", "tiff": "TIFF", "avif": "AVIF"}[b]
        conv = [(f.convert("RGBA") if b != "gif" else f.convert("RGBA").convert("P", palette=Image.ADAPTIVE)) for f in frames]
        kw = {"save_all": len(conv) > 1, "append_images": conv[1:]} if len(conv) > 1 else {}
        if b in ("gif", "webp") and len(conv) > 1:
            kw.update(duration=im.info.get("duration", 100), loop=0)
        if b == "webp":
            kw["quality"] = 88
        if b == "avif":
            kw["quality"] = 70
        if b == "tiff":
            kw["compression"] = "tiff_lzw"
        try:
            conv[0].save(out, fmt, **kw)
        except (KeyError, OSError, ValueError):
            raise ConversionError(f"{FORMATS[b][0]} pictures can't be made on the server yet.")
    else:
        raise ConversionError("Unsupported picture conversion.")
    return out


# ---------------------------------------------------------------- sound and video
_A_CODEC = {"mp3": ["-c:a", "libmp3lame", "-q:a", "3"], "wav": ["-c:a", "pcm_s16le"], "m4a": ["-c:a", "aac", "-b:a", "160k"],
            "aac": ["-c:a", "aac", "-b:a", "160k", "-f", "adts"], "ogg": ["-c:a", "libvorbis", "-q:a", "5"],
            "opus": ["-c:a", "libopus", "-b:a", "96k"], "flac": ["-c:a", "flac"]}
MEDIA_MAX_MIN = {"audio": 60, "video_audio": 60, "video": 3, "gif": 0.25}


def _seconds(src):
    try:
        r = subprocess.run(["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "default=nw=1:nk=1", src],
                           capture_output=True, text=True, timeout=30)
        return float((r.stdout or "0").strip() or 0)
    except Exception:
        return 0.0


def convert_media(src, a, b, out_dir, stem):
    if not shutil.which("ffmpeg"):
        raise ConversionError("Sound and video conversions aren't set up on the server yet.")
    secs = _seconds(src)
    kind = "audio" if a in _AUDIO else ("gif" if b == "gif" else "video_audio" if b in _AUDIO else "video")
    limit = MEDIA_MAX_MIN[kind] * 60
    if secs and secs > limit and kind != "gif":
        raise ConversionError(f"That's {round(secs / 60)} minutes long. Up to {MEDIA_MAX_MIN[kind]} minutes can be converted at a time.")
    out = os.path.join(out_dir, stem + "." + b)
    base = ["ffmpeg", "-hide_banner", "-loglevel", "error", "-y", "-i", src]
    if b in _AUDIO:
        args = base + ["-vn"] + _A_CODEC[b] + [out]
    elif b == "gif":
        args = base + ["-t", "15", "-vf", "fps=10,scale='min(480,iw)':-2:flags=lanczos,split[s0][s1];[s0]palettegen=max_colors=128[p];[s1][p]paletteuse=dither=bayer",
                       "-loop", "0", out]
    elif b == "mp4":
        args = base + ["-vf", "scale='min(1280,iw)':-2", "-c:v", "libx264", "-preset", "veryfast", "-crf", "26", "-pix_fmt", "yuv420p",
                       "-c:a", "aac", "-b:a", "128k", "-movflags", "+faststart", out]
    elif b == "webm":
        args = base + ["-vf", "scale='min(854,iw)':-2", "-c:v", "libvpx", "-deadline", "realtime", "-cpu-used", "8", "-b:v", "1M",
                       "-c:a", "libvorbis", "-q:a", "4", out]
    else:
        raise ConversionError("Unsupported sound or video conversion.")
    _run(args, timeout=210, what="the sound and video converter")
    if not os.path.exists(out) or os.path.getsize(out) == 0:
        raise ConversionError("Nothing came out of that file. It may have no sound or picture in it.")
    return out


# ---------------------------------------------------------------- one door for all
def kind_of(fmt):
    return FORMATS.get(fmt, (None, None, None))[1]


def _tidy_html(path):
    """A web page LibreOffice wrote. It gives a page without a title the title "- no title specified", and a workbook of
    several sheets a list headed "Overview", a line between sheets and "Sheet 1:" before each sheet's name. None of
    that is in the person's file: it goes (a sheet keeps its own name as its heading)."""
    try:
        t = open(path, encoding="utf-8", errors="ignore").read()
    except OSError:
        return
    u = re.sub(r"(<title\b[^>]*>)\s*-?\s*no title specified\s*(</title>)", r"\1\2", t, flags=re.I)
    # a picture's description (its alt text) was also written out beside the picture, as words on the page
    u = re.sub(r'(<img\b[^>]*\balt="([^"]+)"[^>]*/?>)\s*\2\s*(?=</span>)', r"\1", u)
    # LibreOffice's own notes in the page's source, and its word about itself
    u = re.sub(r"<!--\s*(?:This file was converted to xhtml by LibreOffice|Next '[a-z:]+' (?:is|was) a)[^>]*-->\s*", "", u)
    u = re.sub(r'<meta\s+name="(?:DCTERMS\.source|xsl:vendor|generator)"[^>]*>\s*', "", u, flags=re.I)
    if re.search(r"<h1>\s*Overview\s*</h1>", u, flags=re.I) and re.search(r'<a\s+name="table0"', u, flags=re.I):
        u = re.sub(r"<hr>\s*<p>\s*<center>\s*<h1>\s*Overview\s*</h1>.*?</center>\s*</p>\s*", "", u, count=1, flags=re.I | re.S)
        u = re.sub(r"<hr>\s*(?=<a\s+name=\"table\d+\")", "", u, flags=re.I)
        u = re.sub(r"(<a\s+name=\"table\d+\">\s*<h1>)\s*Sheet\s+\d+:\s*<em>(.*?)</em>\s*(</h1>)", r"\1\2\3", u, flags=re.I | re.S)
        u = re.sub(r"<!--\s*\*+\s*-->\s*", "", u)
    if u != t:
        open(path, "w", encoding="utf-8").write(u)


def _tidy_epub(path, stem):
    """An e-book LibreOffice wrote calls a book without a title "Unknown Title" (and its writer "Unknown Author"). A book
    must have a title: it takes the name the person gave the file."""
    try:
        with zipfile.ZipFile(path) as z:
            items = [(i, z.read(i.filename)) for i in z.infolist()]
    except Exception:
        return
    changed = False
    name = _html.escape(stem)
    out = []
    for info, data in items:
        if info.filename.endswith((".opf", ".ncx", ".xhtml")):
            t = data.decode("utf-8", "ignore")
            u = re.sub(r"(<dc:title[^>]*>)\s*Unknown Title\s*(</dc:title>)", lambda m: m.group(1) + name + m.group(2), t)
            u = re.sub(r"<dc:creator[^>]*>\s*Unknown Author\s*</dc:creator>", "", u)
            u = re.sub(r"(<docTitle>\s*<text>)\s*Unknown Title\s*(</text>)", lambda m: m.group(1) + name + m.group(2), u)
            u = re.sub(r"<docAuthor>\s*<text>\s*Unknown Author\s*</text>\s*</docAuthor>", "", u)
            u = re.sub(r"(<title>)\s*Unknown Title\s*(</title>)", lambda m: m.group(1) + name + m.group(2), u)
            u = re.sub(r'<meta\s+content="[^"]*"\s+name="generator"\s*/>|<meta\s+name="generator"\s+content="[^"]*"\s*/>', "", u)
            # LibreOffice describes every picture by the place it keeps it in ("OEBPS/images/image0001.png"): read
            # aloud to someone who cannot see the picture, and shown where the picture cannot be. No description then.
            u = re.sub(r'(<img\b[^>]*\balt=")OEBPS/images/[^"]*(")', r"\1\2", u)
            if len(re.findall(r">\s*Section \d+\s*<", u)) == 1:      # a book without headings: its one chapter is the book
                u = re.sub(r">\s*Section 1\s*<", ">" + name + "<", u)
            if u != t:
                changed = True
                data = u.encode("utf-8")
        out.append((info, data))
    if not changed:
        return
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w") as z:
        for info, data in out:        # "mimetype" stays first and unpacked, as an e-book must have it
            z.writestr(info, data, compress_type=zipfile.ZIP_STORED if info.filename == "mimetype" else zipfile.ZIP_DEFLATED)
    os.replace(tmp, path)


def _tidy(path, b, stem):
    """6.4: a converted file holds only what the person's own file holds (the owner, 19 Nov 2026: "never add anything I
    mean anything at all to converted files"). What the tools that made it wrote into it of their own goes."""
    try:
        ext = os.path.splitext(path)[1].lower().lstrip(".")
        if ext == "html":
            _tidy_html(path)
        elif ext == "epub":
            _tidy_epub(path, stem)
        elif ext in ("docx", "pptx", "xlsx"):
            from converters import strip_tool_marks
            strip_tool_marks(path)
    except Exception:
        pass
    return path


def convert_any(src, a, b, out_dir, stem="converted"):
    """Turns src (a file of kind a) into kind b. Returns the result's path (.zip when it's more than one file)."""
    stem0 = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", stem)[:120] or "converted"
    return _tidy(_convert_any(src, a, b, out_dir, stem), b, stem0)


def _convert_any(src, a, b, out_dir, stem="converted"):
    if not can(a, b):
        raise ConversionError(f"{FORMATS.get(a, (a,))[0]} can't be turned into {FORMATS.get(b, (b,))[0]} here.")
    if os.path.getsize(src) == 0:
        raise ConversionError("This file appears to be empty: there's nothing to convert.")
    os.makedirs(out_dir, exist_ok=True)
    stem = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", stem)[:120] or "converted"
    k = kind_of(a)
    if k == "docs":
        return convert_doc(src, a, b, out_dir, stem)
    if k == "slides":
        return convert_slides(src, a, b, out_dir, stem)
    if k == "sheets":
        return convert_sheet(src, a, b, out_dir, stem)
    if k == "images":
        return convert_image(src, a, b, out_dir, stem)
    if k in ("audio", "video"):
        return convert_media(src, a, b, out_dir, stem)
    raise ConversionError("Unsupported conversion.")
