"""
MasterConvert conversion engine.
Real conversions using LibreOffice (rendering-accurate formats) and
structural rebuilding (python-docx/pptx/openpyxl) for formats that have
no direct renderer path between them.
"""
import datetime
import io
import hashlib
import os
import re
import copy
import subprocess
import signal
import threading
import tempfile
import shutil
import statistics
from collections import Counter
from docx import Document
from docx.shared import Inches as DocxInches, Pt as DocxPt, RGBColor as DocxRGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from pptx.oxml.ns import qn as pptx_qn
from docx.oxml import OxmlElement
from docx.enum.style import WD_STYLE_TYPE
import json as _json
from openpyxl.comments import Comment as XlsxComment
from openpyxl.styles import Alignment as XlsxAlignment
from openpyxl.packaging.custom import StringProperty as XlsxStringProperty
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
from docx.text.run import Run as DocxRun
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt, Length
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor as PptxRGBColor
from pptx.enum.dml import MSO_FILL_TYPE, MSO_COLOR_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.chart.data import CategoryChartData
from pptx.text.text import _Run
from pptx.opc.constants import RELATIONSHIP_TYPE as _PPTX_RT
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font as XlsxFont, PatternFill as XlsxPatternFill
from lxml import etree
import pypdf
import pdfplumber


class ConversionError(Exception):
    pass


def _safe_load(fn, *args, **kwargs):
    """Calls a file-loading function (Document(), Presentation(),
    openpyxl.load_workbook(), pypdf.PdfReader()), converting any failure
    into a clear ConversionError instead of letting a raw library
    exception — a corrupted zip, malformed XML, a structure nested deep
    enough that lxml's parser refuses it outright — propagate with an
    unhelpful, implementation-exposing message."""
    try:
        return fn(*args, **kwargs)
    except ConversionError:
        raise
    except Exception as e:
        raise ConversionError(
            f"Couldn't open this file — it may be corrupted, password-protected, "
            f"or in an unexpected format ({type(e).__name__})."
        )


def _iter_block_items(doc):
    """Yields each paragraph and table in a python-docx Document in actual
    document order. python-docx's own .paragraphs and .tables are separate
    flat lists with no ordering between them — without this, a table
    embedded between two paragraphs silently gets processed out of order
    (or, in the old docx_to_pptx, not at all, since it only ever walked
    .paragraphs). This is the standard recipe for this since python-docx
    doesn't expose a unified iterator itself."""
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield DocxParagraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield DocxTable(child, doc)


_BULLET_RE = re.compile(
    r"^([\u2022\u25CF\u25AA\u25E6\uf0b7\uf0a7\uf076\uf0d8\x7f\-\*]|\d{1,3}[\.\)])\s+"
)
_SENTENCE_END_RE = re.compile(r"[.!?][\'\")\]]*$")
_SHORT_LINE_LEN = 70  # a full wrapped line in a standard document usually runs
                      # much longer than this; a short line ending a sentence
                      # is the most reliable signal pypdf gives us that a
                      # paragraph has actually finished, since it never
                      # exposes real layout/whitespace metadata to check instead.


def _reconstruct_paragraphs(raw_text, font_headings=None, allow_heading_fallback=True):
    """pypdf's extract_text() reflects the PDF's internal line-by-line visual
    layout, not its semantic paragraph structure — a single paragraph comes
    back as many short lines with no reliable blank-line separator, and two
    genuinely distinct paragraphs can butt up against each other with no gap
    at all. Dumping each raw line as its own Word paragraph (the old
    behavior) produces a broken, hard-wrapped look with a line break after
    every few words. This rejoins wrapped lines into real flowing
    paragraphs: lines keep accumulating into the same paragraph until one
    both ends a sentence AND is short (the two together are the closest
    signal available that this was genuinely the last line of a paragraph,
    rather than a mid-sentence wrap that happens to land near a period).
    Bullet/numbered list items and short heading-like lines are kept on
    their own line instead, since those usually aren't meant to be joined
    with what follows. font_headings, if given, is this page's {line_text:
    level} map from _extract_pdf_font_size_headings — real typographic
    signal takes precedence over the text-shape-only heuristic below when
    a line matches. allow_heading_fallback disables that text-shape guess
    entirely when there's no font_headings data to fall back on — used for
    OCR'd text specifically, confirmed directly to need this: OCR output
    line-wraps according to the scanned image's visual layout rather than
    sentence structure, so an ordinary line that happens to lack ending
    punctuation (because the sentence continues on the next OCR line) is
    otherwise easy to misclassify as a heading. Returns a list of (text,
    kind) tuples, kind in {"para", "bullet", "heading1", "heading2",
    "heading3"}."""
    font_headings = font_headings or {}
    lines = [l.strip() for l in raw_text.split("\n")]
    items = []
    current = []

    def flush():
        if current:
            items.append((" ".join(current).strip(), "para"))
            current.clear()

    for line in lines:
        if not line:
            flush()
            continue
        font_level = font_headings.get(line)
        if font_level is not None:
            flush()
            items.append((line, f"heading{font_level}"))
            continue
        if _BULLET_RE.match(line):
            flush()
            _m = _BULLET_RE.match(line)
            # "1." / "2)" markers are numbered steps, not bullets — keeping
            # them apart is what lets Word number them again (they used to
            # come out as plain bullets with the numbers stripped).
            items.append((_BULLET_RE.sub("", line).strip(), "number" if (_m and _m.group(1)[:1].isdigit()) else "bullet"))
            continue
        ends_sentence = bool(_SENTENCE_END_RE.search(line))
        if (
            allow_heading_fallback and not font_headings
            and not current and not ends_sentence and len(line) < 80 and line[:1].isupper()
        ):
            # A short line, sitting on its own, that doesn't end mid-sentence
            # and starts capitalized reads as a heading far more often than
            # not — e.g. "Section One", "PDF Test Document" — this no longer
            # relies on line.istitle()-style checks, which break on any line
            # containing an all-caps acronym. Used as a fallback only when
            # font_headings has no real typographic data for this page at
            # all (a pdfplumber parse failure) — real font-size signal
            # always wins over this text-shape guess when both are present.
            items.append((line, "heading3"))
            continue
        current.append(line)
        if ends_sentence and len(line) < _SHORT_LINE_LEN:
            flush()
    flush()
    return items




def _extract_pdf_font_size_headings(pdfplumber_pdf):
    """Uses pdfplumber's real per-character font size and weight data
    (pypdf's extract_text() returns plain strings with no typographic
    info at all) to detect which lines across the document are genuine
    headings — text visibly larger than the document's own body text —
    rather than relying purely on the text-shape heuristic in
    _reconstruct_paragraphs (short, capitalized, doesn't end
    mid-sentence), which has no way to see an actual font-size jump.
    Distinct heading sizes are ranked into levels (the largest text in
    the document becomes level 1, the next distinct size level 2, and
    so on, capped at 3) rather than using a fixed size-ratio cutoff —
    confirmed directly that a fixed ratio threshold conflates a 24pt
    document title and a 16pt section heading into the same level,
    since both comfortably clear a single "big enough" bar. Returns
    (headings_by_page, body_size) where headings_by_page maps
    page_index -> {line_text: level}.

    Takes an already-open pdfplumber PDF object, not a path — this and
    _extract_pdf_ruled_tables used to each open their own independent
    pdfplumber.open() on the same file, the latter doing so once per
    page. Profiled directly against a realistic 10-page document and
    confirmed pdfplumber's own page-parsing machinery (page.objects and
    everything it triggers) accounted for the overwhelming majority of
    this function's total time -- re-opening and re-parsing the entire
    file body eleven separate times for one document's worth of work.
    pdf_to_docx now opens the file exactly once and passes that same
    object to both functions.
    """
    all_sizes = Counter()
    page_lines = []
    try:
        for page_idx, page in enumerate(pdfplumber_pdf.pages):
            try:
                words = page.extract_words(extra_attrs=["size", "fontname"])
            except Exception:
                page_lines.append((page_idx, {}))
                continue
            lines = {}
            for w in words:
                key = round(w["top"])
                lines.setdefault(key, []).append(w)
            line_data = {}
            for top, line_words in lines.items():
                text = " ".join(w["text"] for w in line_words).strip()
                if not text:
                    continue
                size = max(w["size"] for w in line_words)
                bold = any("bold" in w["fontname"].lower() for w in line_words)
                line_data[top] = (text, size, bold)
                all_sizes[round(size)] += len(text)  # weighted by character count, so
                # body text (far more total characters than any heading) dominates
            page_lines.append((page_idx, line_data))
    except Exception:
        return {}, None  # a malformed/unusual PDF structure pdfplumber can't parse falls
        # back to the existing text-pattern-only heading heuristic, not a hard failure

    if not all_sizes:
        return {}, None
    body_size = all_sizes.most_common(1)[0][0]

    candidate_sizes = sorted({s for s in all_sizes if s >= body_size * 1.15}, reverse=True)
    size_to_level = {}
    level = 1
    prev_size = None
    for s in candidate_sizes:
        if prev_size is not None and prev_size - s >= 1:
            level += 1     # each distinct heading size is its own level (Word's Heading 1/2 can differ by a single point)
        size_to_level[s] = level
        prev_size = s
        if level > 3:
            break

    headings_by_page = {}
    for page_idx, line_data in page_lines:
        page_headings = {}
        for top, (text, size, bold) in line_data.items():
            if len(text) > 120:
                continue  # a genuine heading is short; a large pull-quote or similar isn't safe to treat as one
            rounded = round(size)
            if rounded in size_to_level:
                page_headings[text] = size_to_level[rounded]
        if page_headings:
            headings_by_page[page_idx] = page_headings
    return headings_by_page, body_size


def _extract_pdf_ruled_tables(pdfplumber_pdf, page_idx):
    """Detects tables on one page via pdfplumber's default ruled-line
    strategy only — deliberately not the more permissive text-alignment
    strategy, confirmed directly to be unsafe: tested against a real
    page of ordinary prose, the text-alignment strategy swallowed the
    entire page (headings and paragraphs alike) into one garbled
    "table" based on incidental column alignment, which would corrupt
    far more content than the tables it might additionally catch are
    worth. The trade-off is real but the right one: a table with no
    visible ruling lines at all is missed and falls back to the
    existing paragraph-reconstruction text flow, rather than risking a
    false positive that corrupts a page of real prose.

    Takes an already-open pdfplumber PDF object, not a path — see
    _extract_pdf_font_size_headings for why (this was, until now, the
    more severe half of that same problem: called once per page, each
    call re-opening and re-parsing the entire file just to use one page
    of the result)."""
    try:
        if page_idx >= len(pdfplumber_pdf.pages):
            return []
        return pdfplumber_pdf.pages[page_idx].extract_tables()
    except Exception:
        return []


def _ocr_pdf_page(pdf_path, page_num, work_dir):
    """Rasterizes one PDF page (1-indexed, matching pdftoppm's -f/-l page
    numbering) and runs tesseract on it — the same rasterize-then-OCR
    pattern already used for pdf_to_pptx, scoped here to a single page
    rather than the whole document, since pdf_to_docx only ever needs
    OCR for the specific pages that had no extractable text at all
    (most PDFs are mostly or entirely text-based already). Returns the
    recognized text, or an empty string on any failure — OCR not
    working shouldn't sink the rest of the conversion, just leave that
    one page's placeholder message in place."""
    page_prefix = os.path.join(work_dir, f"ocr_page_{page_num}")
    try:
        subprocess.run(
            ["pdftoppm", "-png", "-r", "200", "-f", str(page_num), "-l", str(page_num), pdf_path, page_prefix],
            capture_output=True, text=True, timeout=30, check=True
        )
    except Exception:
        return ""
    png_candidates = [f for f in os.listdir(work_dir) if f.startswith(f"ocr_page_{page_num}") and f.endswith(".png")]
    if not png_candidates:
        return ""
    png_path = os.path.join(work_dir, png_candidates[0])
    try:
        ocr = subprocess.run(
            ["tesseract", png_path, "-", "--psm", "6"],
            capture_output=True, text=True, timeout=30
        )
        return ocr.stdout.strip()
    except Exception:
        return ""
    finally:
        try:
            os.remove(png_path)  # scratch file — not part of the delivered output
        except OSError:
            pass


# ---------------------------------------------------------- PPTX templates
PPTX_TEMPLATES = {
    # Modern versions of the four Notes -> Slides looks (same names). No bold
    # anywhere: titles stand out by size and colour. "chrome" adds each
    # look's design elements (see _modern_pptx_chrome); "marker" colours the
    # bullet markers. Uploaded corporate templates have neither and keep
    # their own design untouched.
    "minimal": {
        "bg": None, "title_fill": None, "chrome": "block", "accent": PptxRGBColor(0x4F, 0x46, 0xE5),
        "marker": PptxRGBColor(0x4F, 0x46, 0xE5),
        "title_font": "Calibri Light", "title_size": Pt(40), "title_bold": False,
        "title_color": PptxRGBColor(0x11, 0x18, 0x27),
        "body_font": "Calibri", "body_size": Pt(26),
        "body_color": PptxRGBColor(0x37, 0x41, 0x51),
        "subhead_size": Pt(30), "caption_size": Pt(15),
    },
    "academic": {
        "bg": None, "title_fill": None, "chrome": "rail", "accent": PptxRGBColor(0x1F, 0x3A, 0x5F),
        "marker": PptxRGBColor(0x1F, 0x3A, 0x5F),
        "title_font": "Calibri Light", "title_size": Pt(40), "title_bold": False,
        "title_color": PptxRGBColor(0x1F, 0x3A, 0x5F),
        "body_font": "Calibri", "body_size": Pt(26),
        "body_color": PptxRGBColor(0x22, 0x22, 0x22),
        "subhead_size": Pt(30), "caption_size": Pt(15),
    },
    "bold": {
        "bg": PptxRGBColor(0x0B, 0x0B, 0x0F), "title_fill": None, "chrome": "editorial", "accent": PptxRGBColor(0xF5, 0x9E, 0x0B),
        "marker": PptxRGBColor(0xF5, 0x9E, 0x0B),
        "title_font": "Calibri Light", "title_size": Pt(46), "title_bold": False,
        "title_color": PptxRGBColor(0xFF, 0xFF, 0xFF),
        "body_font": "Calibri", "body_size": Pt(26),
        "body_color": PptxRGBColor(0xE5, 0xE7, 0xEB),
        "subhead_size": Pt(30), "caption_size": Pt(15),
    },
    "classic": {
        "bg": None, "title_fill": None, "chrome": "band", "band": PptxRGBColor(0x1F, 0x38, 0x64), "accent": PptxRGBColor(0x38, 0xBD, 0xF8),
        "marker": PptxRGBColor(0x1F, 0x38, 0x64),
        "title_font": "Calibri Light", "title_size": Pt(38), "title_bold": False,
        "title_color": PptxRGBColor(0xFF, 0xFF, 0xFF),
        "body_font": "Calibri", "body_size": Pt(26),
        "body_color": PptxRGBColor(0x22, 0x22, 0x22),
        "subhead_size": Pt(30), "caption_size": Pt(15),
    },
}


# ------------------------------------------------ 3.0: the Plain look (R11)
# White, black words, nothing else: no bars, rails, bands or coloured bullet marks. Titles fit
# on one line; nothing sits beside a picture; text never runs off a slide — what doesn't fit
# carries on, under the same title, on the next one.
PPTX_TEMPLATES["plain"] = {
    "bg": PptxRGBColor(0xFF, 0xFF, 0xFF), "title_fill": None, "chrome": None, "plain": True,
    "accent": PptxRGBColor(0, 0, 0), "marker": PptxRGBColor(0, 0, 0),
    # One family, Calibri, for titles and words: every program has it or a font of exactly its
    # widths (Carlito), so a title measured to fit on one line does fit on one line.
    "title_font": "Calibri", "title_size": Pt(40), "title_bold": False,
    "title_color": PptxRGBColor(0, 0, 0),
    "body_font": "Calibri", "body_size": Pt(28),
    "body_color": PptxRGBColor(0, 0, 0),
    "subhead_size": Pt(30), "caption_size": Pt(16),
}
_PLAIN_SUB_PT = 24        # words one level in, and deeper
_PLAIN_GAP_PT = 10        # space above a paragraph (6 above one a level in)
_PLAIN_LINE = 1.22        # Calibri's line height, in ems
_EMU_PT = 12700
_PPTX_INSETS = 91440      # a text box's inner margins: 0.05 in above and below, 0.1 in each side
_PPTX_MARL = (342900, 742950, 1143000, 1600200, 2057400)   # each level's text indent (default layouts)
_PPTX_TABLE_CONT = "Academic-Home table (continued)"
_PPTX_TABLE_CONT_OLD = "Docente table (continued)"     # decks made before 3.4 (the app was called Docente)

# Calibri's character widths (Carlito's are the same), " " to "~", in thousandths of an em.
_CALIBRI_W = (226, 326, 401, 498, 507, 715, 682, 221, 303, 303, 498, 498, 250, 306, 252, 386,
              507, 507, 507, 507, 507, 507, 507, 507, 507, 507, 268, 268, 498, 498, 498, 463,
              894, 579, 544, 533, 615, 488, 459, 631, 623, 252, 319, 520, 420, 855, 646, 662,
              517, 673, 543, 459, 487, 642, 567, 890, 519, 487, 468, 307, 386, 307, 498, 498,
              291, 479, 525, 423, 525, 498, 305, 471, 525, 230, 239, 455, 230, 799, 525, 527,
              525, 525, 349, 391, 335, 525, 452, 715, 433, 453, 395, 314, 460, 314, 498)
_CALIBRI_BW = (226, 326, 438, 498, 507, 729, 705, 233, 312, 312, 498, 498, 258, 306, 267, 430,
               507, 507, 507, 507, 507, 507, 507, 507, 507, 507, 276, 276, 498, 498, 498, 463,
               898, 606, 561, 529, 630, 488, 459, 637, 631, 267, 331, 547, 423, 874, 659, 676,
               532, 686, 563, 473, 495, 653, 591, 906, 551, 520, 478, 325, 430, 325, 498, 498,
               300, 494, 537, 418, 537, 503, 316, 474, 537, 246, 255, 480, 246, 813, 537, 538,
               537, 537, 355, 399, 347, 537, 473, 745, 459, 474, 397, 344, 475, 344, 498)


def _calibri_pt(text, size_pt, bold=False):
    """How wide `text` is, in points, set in Calibri at `size_pt`."""
    table = _CALIBRI_BW if bold else _CALIBRI_W
    total = 0
    for ch in str(text):
        o = ord(ch)
        if 32 <= o <= 126:
            total += table[o - 32]
        elif ch in "‘’‚":
            total += 250
        elif ch in "“”„":
            total += 420
        elif ch == "—":
            total += 905
        elif ch in "–•… ":
            total += 498
        elif 0x3000 <= o <= 0x9FFF or 0xAC00 <= o <= 0xD7AF or 0xF900 <= o <= 0xFAFF:
            total += 1000
        else:
            total += 560
    return total * size_pt / 1000.0


def _calibri_lines(text, size_pt, width_pt, bold=False):
    """How many lines `text` takes in a box `width_pt` wide, wrapped at spaces as a slide wraps
    it (a word longer than a whole line breaks where it must)."""
    width_pt = max(width_pt, size_pt)
    space = _calibri_pt(" ", size_pt, bold)
    total = 0
    for seg in str(text).replace("\x0b", "\n").replace("\v", "\n").split("\n"):
        lines, cur, started = 1, 0.0, False
        for word in seg.split(" "):
            w = _calibri_pt(word, size_pt, bold)
            if started and cur + space + w <= width_pt:
                cur += space + w
                continue
            if started:
                lines += 1
            started = True
            while w > width_pt:
                lines += 1
                w -= width_pt
            cur = w
        total += lines
    return max(1, total)


def _topic_from_text(text, max_words=7):
    """A short topic in a text's own words, for a slide the document gave no heading: the
    lead-in before a colon or dash; else its first sentence — whole when short, up to its
    first pause when that comes within nine words — else its first words and an ellipsis."""
    t = " ".join(str(text or "").split()).strip()
    t = re.sub(r"^(?:[•▪●◦‣*>–—-]+|\(?\d{1,2}[.)])\s+", "", t)
    if not t:
        return ""
    m = re.match(r"^(.{3,60}?)\s*(?::|\s[–—-])\s+\S", t)
    if m and len(m.group(1).split()) <= 9:
        return m.group(1).rstrip(" .,;")
    first = re.split(r"(?<=[.!?])\s+(?=[A-Z0-9\"“(])", t, maxsplit=1)[0].rstrip(" .!?;:,")
    words = first.split()
    if len(words) <= 9:
        return first
    cut = re.match(r"^((?:\S+\s+){2,8}\S+?)[,;:]\s", first)
    if cut:
        return cut.group(1).rstrip(" ,;:")
    return " ".join(words[:max_words]).rstrip(" ,;:") + "…"


def _clean_doc_title(title):
    """A document's own title property, when it is a real title (not a template's leftover)."""
    t = " ".join(str(title or "").split())
    t = re.sub(r"^microsoft (?:word|powerpoint|excel)\s*-\s*", "", t, flags=re.I)
    t = re.sub(r"\.(?:docx?|pdf|pptx?|txt|rtf)$", "", t, flags=re.I).strip()
    if not t or len(t) > 120 or re.fullmatch(r"(?:new\s+)?(?:document|untitled|title|presentation|slide|doc|file)\s*\d*", t, re.I):
        return ""
    return t


def _plain_pptx_geometry(prs):
    """Plain: the title across the top; the words below it, down to just above the footer
    (the default layouts leave the bottom fifth of the slide empty); a title slide's title
    and subtitle in the middle."""
    inch = 914400

    def put(ph, x, y, w, h):
        ph.left, ph.top, ph.width, ph.height = int(x * inch), int(y * inch), int(w * inch), int(h * inch)
    for ph in prs.slide_master.placeholders:
        kind = ph.placeholder_format.type
        if kind == PP_PLACEHOLDER.TITLE:
            put(ph, 0.6, 0.35, 12.13, 1.05)
        elif kind == PP_PLACEHOLDER.BODY:
            put(ph, 0.6, 1.5, 12.13, 5.4)
    for ph in prs.slide_layouts[0].placeholders:
        kind = ph.placeholder_format.type
        if kind == PP_PLACEHOLDER.CENTER_TITLE:
            put(ph, 0.9, 1.45, 11.53, 2.45)
        elif kind == PP_PLACEHOLDER.SUBTITLE:
            put(ph, 0.9, 4.05, 11.53, 1.7)


def _plain_fit_title(shape, style, title_slide=False):
    """Plain: the title fits its box on one line — its size lowered (to 28 pt at the least)
    until it does; a title too long for that gets two lines (24 pt at the least) and only a
    very long one more, its box made taller to hold them. A title slide's title may take two
    or three lines. Content titles sit at the left, a title slide's in the centre. Returns
    how much taller (EMU) the box was made."""
    from pptx.enum.text import MSO_ANCHOR as _ANCHOR
    if shape is None or not shape.has_text_frame:
        return 0
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = _ANCHOR.BOTTOM if title_slide else _ANCHOR.MIDDLE
    for p in tf.paragraphs:
        p.alignment = PP_ALIGN.CENTER if title_slide else PP_ALIGN.LEFT
        p.line_spacing = 0.9
    text = "\n".join(p.text for p in tf.paragraphs).strip()
    if not text:
        return 0
    width_pt = max(72.0, (shape.width - 2 * _PPTX_INSETS) / _EMU_PT) * 0.94
    bold = bool(style.get("title_bold"))
    stages = ((1, 44, 36), (2, 40, 30), (3, 32, 24)) if title_slide else ((1, 40, 28), (2, 30, 24))
    size = lines = None
    for most, hi, lo in stages:
        for sz in range(hi, lo - 1, -2):
            n = _calibri_lines(text, sz, width_pt, bold)
            if n <= most:
                size, lines = sz, n
                break
        if size:
            break
    if size is None:
        size = 24
        lines = _calibri_lines(text, size, width_pt, bold)
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.size = Pt(size)
    need = int((lines * _PLAIN_LINE * 0.9 * size + 10) * _EMU_PT)
    grow = max(0, need - shape.height)
    if grow:
        left, top, width, height = shape.left, shape.top, shape.width, shape.height
        if title_slide:     # grows upwards: it sits on its subtitle
            lift = min(grow, max(0, top - int(0.3 * 914400)))
            shape.left, shape.top, shape.width, shape.height = left, top - lift, width, height + grow
            return 0
        shape.left, shape.top, shape.width, shape.height = left, top, width, height + grow
    return grow


def _plain_title(slide, style):
    """Plain: the slide's title fitted (see _plain_fit_title); a title slide's subtitle starts
    just under the title."""
    from pptx.enum.text import MSO_ANCHOR as _ANCHOR
    t = slide.shapes.title
    if t is None:
        return 0
    title_slide = t.is_placeholder and t.placeholder_format.type == PP_PLACEHOLDER.CENTER_TITLE
    if title_slide:
        for ph in slide.placeholders:
            if ph.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE:
                ph.text_frame.vertical_anchor = _ANCHOR.TOP
    return _plain_fit_title(t, style, title_slide)


def _pptx_restyle_title(slide, style):
    """A title whose words were replaced takes its look's font, size and colour again (setting
    a title's text drops them), and its place: at the left for the modern looks, fitted for
    Plain. Returns how much taller (EMU) a Plain title's box was made."""
    from pptx.enum.text import MSO_ANCHOR as _ANCHOR
    t = slide.shapes.title
    if t is None:
        return 0
    for para in t.text_frame.paragraphs:
        for run in para.runs:
            run.font.name = style["title_font"]
            run.font.size = style["title_size"]
            run.font.bold = style["title_bold"]
            if style["title_color"] is not None:
                run.font.color.rgb = style["title_color"]
    if style.get("chrome"):
        t.text_frame.vertical_anchor = _ANCHOR.MIDDLE
        for para in t.text_frame.paragraphs:
            para.alignment = PP_ALIGN.LEFT
    if style.get("plain"):
        return _plain_title(slide, style)
    return 0


def _pptx_body_list_style(tf, style, sizes):
    """The words' size, font and colour for each level, set on the text box itself — where
    every program reads them (a paragraph's own default isn't read by all of them)."""
    tx_body = tf._txBody
    lst = tx_body.find(pptx_qn("a:lstStyle"))
    if lst is None:
        lst = etree.Element(pptx_qn("a:lstStyle"))
        body_pr = tx_body.find(pptx_qn("a:bodyPr"))
        if body_pr is not None:
            body_pr.addnext(lst)
        else:
            tx_body.insert(0, lst)
    for old in list(lst):
        lst.remove(old)
    color = str(style["body_color"]) if style.get("body_color") is not None else None
    for lvl, size in enumerate(sizes, 1):
        ppr = etree.SubElement(lst, pptx_qn(f"a:lvl{lvl}pPr"))
        d = etree.SubElement(ppr, pptx_qn("a:defRPr"))
        d.set("sz", str(int(round(size * 100))))
        if color:
            fill = etree.SubElement(d, pptx_qn("a:solidFill"))
            etree.SubElement(fill, pptx_qn("a:srgbClr")).set("val", color)
        etree.SubElement(d, pptx_qn("a:latin")).set("typeface", style["body_font"])


def _plain_pptx_table(gtable, size_pt, font):
    """Plain table: white cells, thin black lines, black words; the first row in bold."""
    gtable.first_row = False
    gtable.horz_banding = False
    for r_idx, row in enumerate(gtable.rows):
        for cell in row.cells:
            tc_pr = cell._tc.get_or_add_tcPr()
            for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
                for old in tc_pr.findall(pptx_qn(tag)):
                    tc_pr.remove(old)
            for k, tag in enumerate(("a:lnL", "a:lnR", "a:lnT", "a:lnB")):
                ln = etree.Element(pptx_qn(tag), w="12700", cap="flat", cmpd="sng", algn="ctr")
                fill = etree.SubElement(ln, pptx_qn("a:solidFill"))
                etree.SubElement(fill, pptx_qn("a:srgbClr")).set("val", "000000")
                etree.SubElement(ln, pptx_qn("a:prstDash")).set("val", "solid")
                tc_pr.insert(k, ln)
            cell.fill.solid()
            cell.fill.fore_color.rgb = PptxRGBColor(0xFF, 0xFF, 0xFF)
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(size_pt)
                    run.font.name = font
                    run.font.bold = r_idx == 0
                    run.font.color.rgb = PptxRGBColor(0, 0, 0)


def _plain_ink_ok(hex_color):
    """A colour the author gave some words is kept on a white slide only if it can be read there."""
    try:
        r, g, b = (int(hex_color[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    except (TypeError, ValueError):
        return False
    lin = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    lum = 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b)
    return (1.05) / (lum + 0.05) >= 4.5


def _find_best_pptx_layout(prs, kind):
    """Finds the layout in an uploaded custom template closest to what's
    needed, by name and placeholder shape — a corporate .potx has its own
    layout order and naming, so assuming layout index 1 is always
    "Title and Content" (true only for python-pptx's own default
    template) would silently land content on the wrong layout entirely."""
    layouts = list(prs.slide_layouts)
    if kind == "content":
        for layout in layouts:
            name = (layout.name or "").lower()
            has_title = any(p.placeholder_format.idx == 0 for p in layout.placeholders)
            has_body = any(p.placeholder_format.idx == 1 for p in layout.placeholders)
            if has_title and has_body and ("content" in name or "title and" in name):
                return layout
        for layout in layouts:
            has_title = any(p.placeholder_format.idx == 0 for p in layout.placeholders)
            has_body = any(p.placeholder_format.idx == 1 for p in layout.placeholders)
            if has_title and has_body:
                return layout
        return layouts[min(1, len(layouts) - 1)]
    else:  # "blank"
        for layout in layouts:
            if "blank" in (layout.name or "").lower():
                return layout
        fewest = min(layouts, key=lambda l: len(list(l.placeholders)))
        return fewest


def _extract_template_style(prs, content_layout):
    """Builds a template_style dict (the same shape as PPTX_TEMPLATES'
    entries) from an uploaded custom template's own actual master and
    theme, instead of one of this file's own hardcoded style choices —
    the whole point of a custom template is that its own brand fonts,
    sizes, and colors carry through, not get overwritten by a default
    meant for when no template was provided. Colors are deliberately
    left as None (see _style_pptx_slide / _style_pptx_body_paragraph) so
    a theme-color-based brand palette stays tied to the theme rather
    than being frozen into a specific RGB snapshot."""
    master = content_layout.slide_master
    master_xml = master.element

    def read_style_size(style_tag, fallback_pt):
        el = master_xml.find(".//" + pptx_qn(f"p:{style_tag}") + "/" + qn("a:lvl1pPr") + "/" + qn("a:defRPr"))
        if el is not None and el.get("sz"):
            try:
                return Pt(int(el.get("sz")) / 100)
            except (ValueError, TypeError):
                pass
        return Pt(fallback_pt)

    def read_theme_font(major):
        tag = "majorFont" if major else "minorFont"
        try:
            theme_part = master.part.part_related_by(_PPTX_RT.THEME)
            theme_root = etree.fromstring(theme_part.blob)
            font_el = theme_root.find(".//" + qn("a:fontScheme") + "/" + qn(f"a:{tag}") + "/" + qn("a:latin"))
            if font_el is not None and font_el.get("typeface"):
                return font_el.get("typeface")
        except Exception:
            pass
        return "Calibri"

    title_size = read_style_size("titleStyle", 40)
    body_size = read_style_size("bodyStyle", 26)
    title_font = read_theme_font(major=True)
    body_font = read_theme_font(major=False)

    return {
        "bg": None, "title_fill": None,
        "title_font": title_font, "title_size": title_size, "title_bold": True,
        "title_color": None,
        "body_font": body_font, "body_size": body_size,
        "body_color": None,
        "subhead_size": Pt(min(body_size.pt + 4, title_size.pt - 4)),
        "caption_size": Pt(max(body_size.pt - 11, 12)),
    }


def _fit_layouts_to_widescreen(prs, designed_width=None):
    """python-pptx's built-in template is laid out for a 10-inch-wide (4:3)
    slide. Widening the slide to 13.33in (16:9) doesn't move the layouts'
    placeholders, so every title and text box stayed 9in wide at the left:
    3.8in (29%) of each slide sat empty on the right, titles were off-centre,
    and text wrapped early onto extra slides. This stretches every
    placeholder's horizontal position and width to the real slide width.
    Only placeholders with their own position are scaled; ones that inherit
    from the master are covered by scaling the master once."""
    designed_width = designed_width or PptxInches(10)
    f = prs.slide_width / designed_width
    if abs(f - 1) < 0.01:
        return
    def scale(ph):
        if not ph._element.xpath('./p:spPr/a:xfrm'):
            return
        ph.left, ph.width = int(ph.left * f), int(ph.width * f)
    for ph in prs.slide_master.placeholders:
        scale(ph)
    for layout in prs.slide_layouts:
        for ph in layout.placeholders:
            scale(ph)


def _style_pptx_slide(slide, style):
    if style["bg"] is not None:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = style["bg"]
    title_shape = slide.shapes.title
    if style["title_fill"] is not None:
        title_shape.fill.solid()
        title_shape.fill.fore_color.rgb = style["title_fill"]
    for para in title_shape.text_frame.paragraphs:
        for run in para.runs:
            run.font.name = style["title_font"]
            run.font.size = style["title_size"]
            run.font.bold = style["title_bold"]
            # A custom uploaded template's color may be a theme-color
            # reference rather than a fixed RGB value (a brand palette
            # that's meant to stay tied to the theme) — style["title_color"]
            # is None specifically for that case, so it's left alone here
            # rather than overwritten with an extracted snapshot color.
            if style["title_color"] is not None:
                run.font.color.rgb = style["title_color"]
    if style.get("chrome"):
        _modern_pptx_chrome(slide, style)
    if style.get("plain"):
        return _plain_title(slide, style)       # how much taller the title's box was made
    return 0


def _modern_pptx_chrome(slide, style):
    """Each built-in look's design elements, placed around the slide's real
    title: Minimal's accent block beside it, Academic's side rail, Bold
    Editorial's accent bar above it, Classic Outline's full-width band."""
    from pptx.enum.shapes import MSO_SHAPE as _SHAPE
    t = slide.shapes.title
    if t is None:
        return
    W, H = slide.part.package.presentation_part.presentation.slide_width, slide.part.package.presentation_part.presentation.slide_height

    def rect(x, y, w, h, rgb, back=False):
        shp = slide.shapes.add_shape(_SHAPE.RECTANGLE, int(x), int(y), int(w), int(h))
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb
        shp.line.fill.background()
        shp.shadow.inherit = False
        if back:
            el = shp._element
            el.getparent().remove(el)
            slide.shapes._spTree.insert(2, el)
        return shp

    kind, inch = style["chrome"], 914400
    # Modern looks read left-aligned, so each accent sits with its title.
    from pptx.enum.text import PP_ALIGN as _ALIGN, MSO_ANCHOR as _ANCHOR
    t.text_frame.vertical_anchor = _ANCHOR.MIDDLE
    for para in t.text_frame.paragraphs:
        para.alignment = _ALIGN.LEFT
    if kind == "block":
        rect(max(t.left - int(0.3 * inch), 0), t.top + t.height // 2 - int(0.38 * inch), int(0.1 * inch), int(0.76 * inch), style["accent"])
    elif kind == "rail":
        rect(0, 0, int(0.22 * inch), H, style["accent"], back=True)
    elif kind == "editorial":
        rect(t.left + int(0.08 * inch), max(t.top - int(0.05 * inch), int(0.12 * inch)), int(0.9 * inch), int(0.08 * inch), style["accent"])
    elif kind == "band":
        bottom = t.top + t.height
        rect(0, 0, W, bottom, style["band"], back=True)
        rect(0, bottom, int(2.6 * inch), int(0.06 * inch), style["accent"])
        # breathing room between the band and the text below it
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == 1 and ph.top < bottom + int(0.3 * inch):
                # Set all four: a placeholder that inherits its position gets
                # a zero left and width if only top and height are written.
                left, top, width, height = ph.left, ph.top, ph.width, ph.height
                shift = bottom + int(0.3 * inch) - top
                ph.left, ph.top, ph.width, ph.height = left, top + shift, width, max(height - shift, int(1.0 * inch))


def _style_pptx_body_paragraph(p, style, size=None):
    p.font.name = style["body_font"]
    p.font.size = size if size is not None else style["body_size"]
    if style["body_color"] is not None:
        p.font.color.rgb = style["body_color"]
    if style.get("marker") is not None:
        pPr = p._p.get_or_add_pPr()
        if pPr.find(pptx_qn("a:buNone")) is None:
            for old in pPr.findall(pptx_qn("a:buClr")):
                pPr.remove(old)
            clr = etree.SubElement(pPr, pptx_qn("a:buClr"))
            etree.SubElement(clr, pptx_qn("a:srgbClr")).set("val", str(style["marker"]))
            pPr.remove(clr)
            pPr.insert(0, clr)


# ---------------------------------------------------------- DOCX styles
DOCX_STYLES = {
    "clean": {
        "heading_font": "Calibri", "heading_color": DocxRGBColor(0x22, 0x22, 0x22),
        "body_font": "Calibri", "body_size": DocxPt(11), "uppercase_headings": False,
        "page_numbers": False, "tight_spacing": False,
    },
    "academic": {
        "heading_font": "Cambria", "heading_color": DocxRGBColor(0x1F, 0x3A, 0x5F),
        "body_font": "Cambria", "body_size": DocxPt(12), "uppercase_headings": False,
        "page_numbers": True, "tight_spacing": False,
    },
    "report": {
        "heading_font": "Calibri", "heading_color": DocxRGBColor(0x7A, 0x1F, 0x2B),
        "body_font": "Calibri", "body_size": DocxPt(11), "uppercase_headings": True,
        "page_numbers": True, "tight_spacing": False,
    },
    "compact": {
        "heading_font": "Calibri", "heading_color": DocxRGBColor(0x22, 0x22, 0x22),
        "body_font": "Calibri", "body_size": DocxPt(9), "uppercase_headings": False,
        "page_numbers": False, "tight_spacing": True,
    },
}


def apply_docx_style(doc, style_name):
    style = DOCX_STYLES.get(style_name, DOCX_STYLES["clean"])
    # A single pass over doc.paragraphs, not two: _style_docx_headings
    # and _style_docx_body previously each looped over every paragraph
    # independently, meaning the same paragraph's .style.name — a
    # property confirmed elsewhere in this file to be expensive,
    # python-docx re-walking the style's XML from scratch on every
    # access — was resolved twice. Merging into one pass halves that
    # immediately; the cache below (keyed by the cheap raw XML value
    # rather than the expensive resolved name, the same fix applied to
    # docx_to_pptx above) additionally means paragraphs sharing the same
    # style, the common case for a document's many body paragraphs,
    # only pay that cost once for the whole document rather than once
    # per paragraph. Scoped as a local, not module-level like the
    # docx_to_pptx caches, since this function is itself the natural,
    # single-call boundary — nothing to clear between conversions.
    is_heading_cache = {}
    for para in doc.paragraphs:
        cache_key = para._p.style
        if cache_key not in is_heading_cache:
            name = para.style.name
            is_heading_cache[cache_key] = name.lower().startswith("heading") or name == "Title"
        is_heading = is_heading_cache[cache_key]
        if is_heading:
            if style["uppercase_headings"]:
                for run in para.runs:
                    run.font.all_caps = True   # shown in capitals; the words themselves stay as written
            for run in para.runs:
                run.font.name = style["heading_font"]
                run.font.color.rgb = style["heading_color"]
        else:
            for run in para.runs:
                run.font.name = style["body_font"]
                if run.font.size is None:
                    run.font.size = style["body_size"]
            if style.get("tight_spacing"):
                para.paragraph_format.space_before = DocxPt(0)
                para.paragraph_format.space_after = DocxPt(2)
    # doc.paragraphs is only the top-level body — text inside tables never
    # got the document's font, so every table switched to its table style's
    # own typeface (a serif, beside sans-serif text all around it).
    def _style_table_runs(tables):
        for t in tables:
            for row in t.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.font.name = style["body_font"]
                    _style_table_runs(cell.tables)
    _style_table_runs(doc.tables)
    if style["page_numbers"]:
        _add_docx_page_number_footer(doc)


def _add_docx_page_number_footer(doc):
    section = doc.sections[0]
    footer = section.footer
    para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)


_LO_PROFILE_ROOT = os.path.join(tempfile.gettempdir(), "docente_lo_profiles")


def _lo_profile_dir():
    """This worker thread's own LibreOffice profile. Building a brand-new
    profile for every call (the earlier safeguard) added start-up time to
    every conversion and preview; one reusable profile per thread keeps the
    same safeguard — no two conversions ever share a profile, so none can
    wait on another's lock — without paying that cost each time."""
    d = os.path.join(_LO_PROFILE_ROOT, f"{os.getpid()}-{threading.get_ident()}")
    os.makedirs(d, exist_ok=True)
    return d


# 3.4: LibreOffice slots, shared by every worker process on this server (see server_soffice34.py)
_SOFFICE_SLOTS = max(1, min(8, int(os.environ.get("SOFFICE_AT_ONCE", "1") or 1)))


def _soffice_slot(wait=150):
    """Waits for a free LibreOffice slot and returns its open lock file (closing it frees the slot)."""
    import fcntl
    import time as _time
    deadline = _time.time() + wait
    while True:
        for i in range(_SOFFICE_SLOTS):
            fh = open(os.path.join(tempfile.gettempdir(), f"docente-soffice-{i}.lock"), "w")
            try:
                fcntl.flock(fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
                return fh
            except OSError:
                fh.close()
        if _time.time() > deadline:
            raise ConversionError("The converter is busy right now. Please try again in a minute.")
        _time.sleep(0.25)


def _soffice_convert(src_path, target_format, out_dir):
    slot = _soffice_slot()
    try:
        return _soffice_convert_now(src_path, target_format, out_dir)
    finally:
        slot.close()


def _soffice_convert_now(src_path, target_format, out_dir):
    """Use headless LibreOffice for true rendering-based conversions
    (docx->pdf, any supported format -> pdf for preview, recalculating
    spreadsheets). Raises ConversionError on failure.

    Runs in its own process group with this thread's own profile. If a
    conversion hangs, the whole group is killed (a plain timeout only kills
    the `soffice` launcher and leaves the `soffice.bin` renderer running),
    and that profile is deleted so a stale lock from the killed process
    can't wedge the next conversion on this thread. 180s leaves headroom
    under the app's 240s request timeout."""
    profile_dir = _lo_profile_dir()
    stdout, stderr = "", ""
    failed = True
    try:
        proc = subprocess.Popen(
            ["soffice", "--headless", "--nologo", "--nofirststartwizard",
             f"-env:UserInstallation=file://{profile_dir}",
             "--convert-to", target_format, "--outdir", out_dir, src_path],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
            start_new_session=True,
        )
        try:
            stdout, stderr = proc.communicate(timeout=180)
        except subprocess.TimeoutExpired:
            try:
                os.killpg(os.getpgid(proc.pid), signal.SIGKILL)
            except ProcessLookupError:
                pass
            proc.communicate()
            raise ConversionError(
                "That file took too long to render. Please try again — "
                "if it keeps happening, the file may be unusually large or complex."
            )
        failed = proc.returncode != 0
    finally:
        if failed:
            shutil.rmtree(profile_dir, ignore_errors=True)

    base = os.path.splitext(os.path.basename(src_path))[0]
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"{base}.{target_format}")
    if not os.path.exists(out_path):
        shutil.rmtree(profile_dir, ignore_errors=True)
        raise ConversionError(f"LibreOffice conversion failed: {stderr or stdout}")
    return out_path


def warm_up_libreoffice():
    """Runs one tiny conversion when the server starts, so LibreOffice's
    program files are already loaded into memory — otherwise the first
    student after every deploy or restart waits for them to be read from a
    cold disk. Meant for a background thread; any failure is ignored."""
    work = tempfile.mkdtemp(prefix="lo_warm_")
    try:
        src = os.path.join(work, "warm.txt")
        with open(src, "w") as fh:
            fh.write("warm-up")
        _soffice_convert(src, "pdf", work)
    except Exception:
        pass
    finally:
        shutil.rmtree(work, ignore_errors=True)

def docx_to_pdf(src_path, out_dir, style=None):
    return _soffice_convert(src_path, "pdf", out_dir)


def pptx_to_pdf(src_path, out_dir, style=None):
    """Slides as a PDF: one page per slide, as LibreOffice draws them."""
    return _soffice_convert(src_path, "pdf", out_dir)


def xlsx_to_pdf(src_path, out_dir, style=None):
    """A spreadsheet as a PDF, page by page as it would print."""
    return _soffice_convert(src_path, "pdf", out_dir)


def preview_file(src_path, out_dir, ext):
    """Produces a PDF the frontend can embed directly for a genuine,
    real preview of whatever was just generated — not a mocked-up
    stand-in, the browser's own native PDF renderer showing the actual
    converted or AI-generated file's real content. A PDF source needs
    no conversion at all; docx/pptx/xlsx all go through the same
    existing LibreOffice pipeline already proven for docx_to_pdf above,
    since --convert-to pdf isn't docx-specific — LibreOffice renders
    any format it can open the same way, faithfully reflecting the
    real file's actual layout, fonts, and content rather than an
    approximation of it."""
    ext = ext.lower().lstrip(".")
    if ext == "pdf":
        return src_path
    if ext in ("docx", "pptx", "xlsx", "doc", "ppt", "xls"):
        return _soffice_convert(src_path, "pdf", out_dir)
    raise ConversionError(f"Preview isn't supported for .{ext} files.")


# --------------------------------------------------------------- DOCX -> PPTX
def _cellis_rule_matches(rule, value):
    """Evaluates whether a value satisfies a simple CellIsRule's numeric
    comparison (the common 'red if < 0' / 'green if >= 0' traffic-light
    pattern). Only handles a literal numeric comparison, not a formula
    referencing other cells — the same proportionate scope as the
    uncalculated-formula fallback elsewhere in this file: cover the
    common case honestly rather than build a full formula evaluator."""
    if value is None or not isinstance(value, (int, float)) or isinstance(value, bool):
        return False
    try:
        formula = [float(f) for f in rule.formula]
    except (ValueError, TypeError):
        return False
    op = rule.operator
    if op == "lessThan":
        return value < formula[0]
    if op == "lessThanOrEqual":
        return value <= formula[0]
    if op == "greaterThan":
        return value > formula[0]
    if op == "greaterThanOrEqual":
        return value >= formula[0]
    if op == "equal":
        return value == formula[0]
    if op == "notEqual":
        return value != formula[0]
    if op == "between" and len(formula) == 2:
        return formula[0] <= value <= formula[1]
    if op == "notBetween" and len(formula) == 2:
        return not (formula[0] <= value <= formula[1])
    return False


def _xlsx_conditional_fill_hex(ws, cell):
    """Returns a cell's effective background color from a matching
    conditional-formatting rule, or None if none applies. Conditional
    formatting colors live entirely separately from cell.fill — a cell
    with no direct fill set can still be visibly red/yellow/green through
    a rule, which _xlsx_cell_fill_hex alone never sees."""
    try:
        applicable = []
        for cf in ws.conditional_formatting:
            if cell.coordinate not in cf.cells:
                continue
            for rule in cf.rules:
                if rule.type == "cellIs" and rule.dxf and rule.dxf.fill:
                    applicable.append(rule)
        applicable.sort(key=lambda r: r.priority)
        for rule in applicable:
            if _cellis_rule_matches(rule, cell.value):
                fg = rule.dxf.fill.fgColor
                if fg is not None and fg.type == "rgb" and fg.rgb and len(fg.rgb) == 8:
                    return fg.rgb[2:]
    except Exception:
        pass
    return None


def _xlsx_chart_title(chart):
    """Extracts a chart's title text from openpyxl's deeply nested title
    object, or None if the chart has no title or an unexpected structure.
    A chart's own underlying data is just ordinary worksheet cells,
    already captured by the normal table conversion — this caption is
    only about letting the reader know a chart existed at all, not about
    recovering data that would otherwise be lost."""
    try:
        rich_paragraphs = chart.title.tx.rich.p
        parts = [run.t for para in rich_paragraphs for run in (para.r or []) if run.t]
        text = "".join(parts).strip()
        return text or None
    except Exception:
        return None


def _xlsx_cell_fill_hex(cell):
    """Returns an openpyxl cell's solid fill color as a 6-char hex string,
    or None if it has no solid fill. openpyxl reports colors as 8-char ARGB
    (alpha + RGB), while docx/pptx RGBColor expects plain 6-char RGB."""
    fill = cell.fill
    if fill.patternType != "solid" or not fill.fgColor or fill.fgColor.type != "rgb":
        return None
    rgb = fill.fgColor.rgb
    return rgb[2:] if rgb and len(rgb) == 8 else None


def _find_docx_merges(src_rows, n_cols):
    """Detects merged cell spans in a docx table. python-docx represents a
    merge as the *same* underlying cell object repeated across the whole
    span rather than exposing merge info directly — comparing cell identity
    across the grid is how you recover the actual rectangle. Returns a list
    of (min_row, min_col, max_row, max_col) tuples, one per real merge
    (spans of exactly 1x1 are skipped, they're just an ordinary cell)."""
    seen = {}
    for r_idx, row in enumerate(src_rows):
        for c_idx in range(min(len(row), n_cols)):
            key = id(row[c_idx]._tc)
            if key not in seen:
                seen[key] = [r_idx, c_idx, r_idx, c_idx]
            else:
                span = seen[key]
                span[0] = min(span[0], r_idx)
                span[1] = min(span[1], c_idx)
                span[2] = max(span[2], r_idx)
                span[3] = max(span[3], c_idx)
    return [tuple(s) for s in seen.values() if s[0] != s[2] or s[1] != s[3]]


def _get_docx_cell_shading(cell):
    """Returns a docx table cell's background shading as a hex string, or
    None if it has none. python-docx has no high-level API for cell
    shading — this reads the raw <w:shd w:fill="..."/> element directly."""
    tcPr = cell._tc.find(qn("w:tcPr"))
    if tcPr is None:
        return None
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        return None
    fill = shd.get(qn("w:fill"))
    return fill if fill and fill.upper() != "AUTO" else None


def _set_docx_cell_shading(cell, hex_color):
    """Sets a docx table cell's background shading. No high-level API for
    this in python-docx — building the <w:shd> element by hand is the
    standard recipe."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


_ALIGN_NAMES = {"LEFT", "CENTER", "RIGHT", "JUSTIFY"}


def _docx_align_to_pptx(align):
    """Word and PowerPoint's paragraph-alignment enums share the same names
    for the alignments that matter (LEFT/CENTER/RIGHT/JUSTIFY) despite being
    two unrelated classes — translating by name is simpler and more robust
    than a hand-maintained value-to-value mapping."""
    name = getattr(align, "name", None)
    return getattr(PP_ALIGN, name) if name in _ALIGN_NAMES else None


def _pptx_align_to_docx(align):
    name = getattr(align, "name", None)
    return getattr(WD_ALIGN_PARAGRAPH, name) if name in _ALIGN_NAMES else None


def _safe_hex_color(color):
    """Returns a run's color as a hex string ('FF0000'), or None if it's
    unset or a theme color rather than an explicit RGB value. Accessing
    .rgb directly raises AttributeError in both python-docx and python-pptx
    for those cases, so this is what makes copying colors between formats
    safe instead of crashing on the first themed or default-colored run."""
    try:
        rgb = color.rgb
    except AttributeError:
        return None
    return str(rgb) if rgb is not None else None


def _get_docx_comments(doc):
    """Returns an ordered list of (author, text) for each real reviewer
    comment in the document. python-docx has no API for comments at all —
    same story as footnotes/endnotes — so this reads the raw comments.xml
    part directly. Comments are genuinely relevant content for a tutoring
    tool specifically: this is often where a teacher's actual feedback on
    a draft lives, not just document metadata to discard."""
    try:
        pkg = doc.part.package
    except Exception:
        return []
    for part in pkg.iter_parts():
        if part.partname == "/word/comments.xml":
            try:
                root = etree.fromstring(part.blob)
            except Exception:
                return []
            comments = []
            for c in root.findall(qn("w:comment")):
                text = _all_text_including_deletions(c).strip()
                if text:
                    author = (c.get(qn("w:author")) or "").strip() or "Comment"
                    comments.append((author, text))
            return comments
    return []


def _get_docx_header_footer_text(doc):
    """Returns distinct, non-empty header/footer paragraph text across all
    sections. A header/footer holding only a page-number field returns
    empty text here (the field's displayed number isn't literal text, just
    a field code), which conveniently filters out routine pagination
    boilerplate on its own — only genuine typed content (a report title, a
    confidentiality line) survives the strip-and-dedupe."""
    seen = set()
    lines = []
    for section in doc.sections:
        for container in (section.header, section.footer):
            for p in container.paragraphs:
                text = _full_paragraph_text(p).strip()
                if text and text not in seen:
                    seen.add(text)
                    lines.append(text)
    return lines


def _all_text_including_deletions(element):
    """Joins all w:t and w:delText descendant text in true document order —
    two separate findall() calls (one per tag) would incorrectly group all
    insertions before all deletions regardless of where each actually sits
    if a note/comment itself contains tracked changes."""
    return "".join(
        el.text or "" for el in element.iter()
        if el.tag in (qn("w:t"), qn("w:delText"))
    )


def _get_docx_notes(doc, part_name, note_tag):
    """Returns an ordered list of real footnote/endnote texts. python-docx
    has no API for either whatsoever — not even a way to detect they exist —
    so this is the only way to reach them: find the raw XML part directly in
    the document's OPC package and parse it by hand. Skips the
    separator/continuationSeparator placeholder entries every document with
    notes has (visual dividers, not real content)."""
    try:
        pkg = doc.part.package
    except Exception:
        return []
    for part in pkg.iter_parts():
        if part.partname == part_name:
            try:
                root = etree.fromstring(part.blob)
            except Exception:
                return []
            notes = []
            for note in root.findall(qn(note_tag)):
                if note.get(qn("w:type")) in ("separator", "continuationSeparator"):
                    continue
                text = _all_text_including_deletions(note).strip()
                if text:
                    notes.append(text)
            return notes
    return []


def _iter_docx_charts(paragraph, doc):
    """Yields (title, categories, series, chart_type) for each chart
    embedded in this paragraph's runs. python-docx has no chart API at
    all — a chart's real data lives entirely in a separate chart XML
    part, reached only via a relationship id on the run's drawing — so
    without reaching for it directly, a chart's numbers are completely
    lost rather than just unstyled, the same class of gap PowerPoint
    charts had before that was fixed. series is a list of (name, values).
    chart_type is an XL_CHART_TYPE value, detected from which plot
    element (c:barChart, c:lineChart, etc.) the source chart actually
    used, so the recreated chart is the same kind of chart, not always a
    generic bar chart regardless of what the original document had."""
    for run in paragraph.runs:
        for chart_ref in run._element.findall(".//" + qn("c:chart")):
            r_id = chart_ref.get(qn("r:id"))
            if not r_id:
                continue
            try:
                chart_part = doc.part.rels[r_id].target_part
                root = etree.fromstring(chart_part.blob)
            except Exception:
                continue
            title = None
            try:
                title_el = root.find(".//" + qn("c:title"))
                if title_el is not None:
                    parts = [t.text for t in title_el.iter(qn("a:t")) if t.text]
                    title = "".join(parts).strip() or None
            except Exception:
                pass
            categories = []
            try:
                cat_el = root.find(".//" + qn("c:cat"))
                if cat_el is not None:
                    categories = [pt.findtext(qn("c:v")) or "" for pt in cat_el.iter(qn("c:pt"))]
            except Exception:
                pass
            series_list = []
            try:
                for ser in root.iter(qn("c:ser")):
                    name_el = ser.find(".//" + qn("c:tx") + "//" + qn("c:v"))
                    name = name_el.text if name_el is not None else "Series"
                    val_el = ser.find(qn("c:val"))
                    values = [pt.findtext(qn("c:v")) or "" for pt in val_el.iter(qn("c:pt"))] if val_el is not None else []
                    series_list.append((name, values))
            except Exception:
                pass
            chart_type = _detect_docx_chart_type(root)
            if categories or series_list:
                yield title, categories, series_list, chart_type


def _detect_docx_chart_type(chart_root):
    """Maps the source chart's actual plot element to the closest
    XL_CHART_TYPE, so a recreated native chart matches the kind of chart
    the original document had (a line chart stays a line chart) instead
    of every chart defaulting to the same generic type regardless of
    source. Falls back to a clustered column chart — the most broadly
    readable default — for chart kinds not worth specifically detecting
    (3D variants, radar, stock, bubble, surface)."""
    def has(tag):
        return chart_root.find(".//" + qn(tag)) is not None

    if has("c:pieChart") or has("c:pie3DChart"):
        return XL_CHART_TYPE.PIE
    if has("c:doughnutChart"):
        return XL_CHART_TYPE.DOUGHNUT
    if has("c:lineChart"):
        return XL_CHART_TYPE.LINE_MARKERS
    if has("c:areaChart"):
        return XL_CHART_TYPE.AREA
    if has("c:scatterChart"):
        return XL_CHART_TYPE.XY_SCATTER
    if has("c:barChart"):
        bar_dir_el = chart_root.find(".//" + qn("c:barChart") + "//" + qn("c:barDir"))
        grouping_el = chart_root.find(".//" + qn("c:barChart") + "//" + qn("c:grouping"))
        is_bar = bar_dir_el is not None and bar_dir_el.get("val") == "bar"
        grouping = grouping_el.get("val") if grouping_el is not None else "clustered"
        if is_bar:
            return XL_CHART_TYPE.BAR_STACKED if grouping == "stacked" else XL_CHART_TYPE.BAR_CLUSTERED
        return XL_CHART_TYPE.COLUMN_STACKED if grouping == "stacked" else XL_CHART_TYPE.COLUMN_CLUSTERED
    return XL_CHART_TYPE.COLUMN_CLUSTERED


def _add_chart_as_table_fallback(slide, categories, series_list, slide_width, slide_height):
    """The previous, already-working behavior for every chart, kept as a
    fallback specifically for chart data python-pptx's native chart API
    can't build from — a plain table of the same numbers is still far
    more useful than losing the chart's data entirely."""
    n_rows, n_cols = len(categories) + 1, len(series_list) + 1
    if n_rows < 2 or n_cols < 2:
        return
    left, top = PptxInches(0.6), PptxInches(1.6)
    width, height = slide_width - PptxInches(1.2), slide_height - PptxInches(2.2)
    gtable = slide.shapes.add_table(n_rows, n_cols, left, top, width, height).table
    gtable.cell(0, 0).text = ""
    for s_idx, (name, _values) in enumerate(series_list, 1):
        gtable.cell(0, s_idx).text = name
    for cat_idx, cat_name in enumerate(categories, 1):
        gtable.cell(cat_idx, 0).text = cat_name
        for s_idx, (_name, values) in enumerate(series_list, 1):
            val = values[cat_idx - 1] if cat_idx - 1 < len(values) else ""
            try:
                val = _format_cell_value(float(val))
            except (ValueError, TypeError):
                pass
            gtable.cell(cat_idx, s_idx).text = val
    for cell in gtable.rows[0].cells:
        for p in cell.text_frame.paragraphs:
            for run in p.runs:
                run.font.bold = True


def _get_docx_footnotes(doc):
    return _get_docx_notes(doc, "/word/footnotes.xml", "w:footnote")


def _get_docx_endnotes(doc):
    return _get_docx_notes(doc, "/word/endnotes.xml", "w:endnote")



_paragraph_style_object_cache = {}


def _cached_paragraph_style(paragraph):
    """Returns paragraph.style, cached by the cheap raw XML value at
    paragraph._p.style rather than by calling the expensive .style
    property itself just to get a cache key (the same mistake avoided
    in _effective_run_format above, and confirmed the same way: reading
    the raw id is roughly 250x cheaper than resolving the full style
    object). Both _effective_paragraph_alignment and
    _get_paragraph_direct_list_info need a paragraph's style object, so
    sharing one cache between them means a paragraph whose style was
    already resolved for one purpose isn't re-resolved from scratch for
    the other. _paragraph_style_object_cache is cleared at the start of
    docx_to_pptx, so no id ever leaks between separate conversions.
    """
    cache_key = paragraph._p.style
    if cache_key not in _paragraph_style_object_cache:
        _paragraph_style_object_cache[cache_key] = paragraph.style
    return _paragraph_style_object_cache[cache_key]


def _effective_paragraph_alignment(paragraph):
    """Returns a paragraph's alignment, falling back to its paragraph
    style's own alignment when not set directly — the same style-carries-
    the-real-formatting gap as run-level bold/italic, just for paragraph
    alignment: a custom style (a 'Caption' or 'Quote' style that bakes in
    centering, for instance) leaves paragraph.alignment as None even
    though the paragraph visibly renders centered."""
    try:
        if paragraph.alignment is not None:
            return paragraph.alignment
    except ValueError:
        return _raw_docx_alignment(paragraph._p.pPr)
    try:
        return _cached_paragraph_style(paragraph).paragraph_format.alignment
    except ValueError:
        try:
            return _raw_docx_alignment(_cached_paragraph_style(paragraph).element.pPr)
        except Exception:
            return None
    except Exception:
        return None


def _raw_docx_alignment(p_pr):
    """A paragraph's alignment read straight from its XML, "start" and "end" included."""
    jc = p_pr.find(qn("w:jc")) if p_pr is not None else None
    val = (jc.get(qn("w:val")) or "").lower() if jc is not None else ""
    return {"left": WD_ALIGN_PARAGRAPH.LEFT, "start": WD_ALIGN_PARAGRAPH.LEFT,
            "right": WD_ALIGN_PARAGRAPH.RIGHT, "end": WD_ALIGN_PARAGRAPH.RIGHT,
            "center": WD_ALIGN_PARAGRAPH.CENTER, "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
            "distribute": WD_ALIGN_PARAGRAPH.DISTRIBUTE}.get(val)


_run_style_format_cache = {}


def _effective_run_format(run):
    """Returns (bold, italic, underline, hex_color, size_pt) for a run,
    falling back to its referenced character style when the run has no
    direct formatting of its own. Word's built-in 'Strong'/'Emphasis'
    styles — and most custom character styles — carry formatting in the
    STYLE definition rather than as direct run formatting, so reading
    run.bold/font.color/font.size directly returns None for all five of
    these even though the text visibly renders styled, silently losing
    the formatting on conversion. Direct run formatting always wins when
    present; the style is only consulted for whichever is still unset.

    The style-level fallback specifically is cached, keyed by run._r.style
    (the raw XML style-id string, or None for "no explicit style") rather
    than run.style.style_id — confirmed directly (2000 calls timed against
    each other) that run.style itself, not just the .font access after it,
    is the expensive part: python-docx re-walks the style's XML definition
    from scratch on every single access, roughly 250x slower than reading
    the raw id straight off the run's own XML. Keying the cache by the
    cheap raw value means a style already seen skips calling .style at
    all, not just the .font lookup after it. _run_style_format_cache is
    cleared at the start of docx_to_pptx (this function's only caller), so
    a cached id from one document can never leak into another conversion.
    """
    bold, italic, underline = run.bold, run.italic, run.underline
    hex_color = _safe_hex_color(run.font.color)
    size_pt = run.font.size.pt if run.font.size is not None else None
    if bold is None or italic is None or underline is None or hex_color is None or size_pt is None:
        try:
            cache_key = run._r.style
            if cache_key in _run_style_format_cache:
                s_bold, s_italic, s_underline, s_color, s_size = _run_style_format_cache[cache_key]
            else:
                style_font = run.style.font
                s_bold = style_font.bold
                s_italic = style_font.italic
                s_underline = style_font.underline
                s_color = _safe_hex_color(style_font.color)
                s_size = style_font.size.pt if style_font.size is not None else None
                _run_style_format_cache[cache_key] = (s_bold, s_italic, s_underline, s_color, s_size)
            if bold is None:
                bold = s_bold
            if italic is None:
                italic = s_italic
            if underline is None:
                underline = s_underline
            if hex_color is None:
                hex_color = s_color
            if size_pt is None:
                size_pt = s_size
        except Exception:
            pass
    return bold, italic, underline, hex_color, size_pt


def _iter_all_runs(paragraph):
    """Yields (Run, text, is_hyperlink, hyperlink_address, is_deletion) for
    every actual text-bearing run in a paragraph, in document order —
    including content python-docx's own .text, .runs, AND
    iter_inner_content() all silently miss entirely: text wrapped in a
    tracked-change insertion (<w:ins>) or deletion (<w:del>). This isn't a
    narrow edge case — a paragraph.text call on a paragraph with tracked
    changes present drops ALL of that text, and tracked-changes markup is
    extremely common in real reviewed academic and business documents
    (exactly the kind an academic tutoring tool would see uploaded).
    Deletions are included rather than skipped, because that matches what
    a reader actually sees in the source with changes displayed — a
    struck-through but still-visible deletion, not truly gone until
    someone accepts the change."""
    def run_text(r_el, deleted):
        tag = qn("w:delText") if deleted else qn("w:t")
        # .findall() searches all descendants, not just direct content —
        # a run containing a shape's <w:drawing> has that shape's own
        # internal paragraph/run/text structure nested inside it (a
        # wps:txbx text box, same as a floating text box's w:txbxContent),
        # and without excluding it here, a shape's own label text gets
        # silently absorbed into the outer paragraph's text too, on top
        # of wherever it's separately extracted from (_iter_docx_shapes).
        results = []
        for t in r_el:
            if t.tag == tag:
                if any(anc.tag in (f"{{{_WPS_NS}}}txbx", qn("w:txbxContent")) for anc in t.iterancestors()):
                    continue
                results.append(t.text or "")
            elif t.tag in (qn("w:br"), qn("w:cr")) and t.get(qn("w:type")) not in ("page", "column"):
                # A line break (Shift+Enter) keeps the words apart — it used to
                # be dropped, so "Nalerigu⏎Department" came out "NaleriguDepartment".
                results.append("\n")
            elif t.tag == qn("w:tab"):
                results.append("\t")
        return "".join(results)

    def walk(container_el, deleted, link_address):
        for child in container_el:
            if child.tag == qn("w:r"):
                text = run_text(child, deleted)
                if text:
                    yield DocxRun(child, paragraph), text, link_address is not None, link_address, deleted
            elif child.tag == qn("w:hyperlink"):
                r_id = child.get(qn("r:id"))
                address = None
                if r_id:
                    try:
                        address = paragraph.part.rels[r_id].target_ref
                    except Exception:
                        pass
                yield from walk(child, deleted, address)
            elif child.tag == qn("w:ins"):
                yield from walk(child, deleted, link_address)
            elif child.tag == qn("w:del"):
                yield from walk(child, True, link_address)

    yield from walk(paragraph._p, False, None)


def _full_paragraph_text(paragraph):
    """The tracked-changes-aware equivalent of paragraph.text — includes
    insertions and (still-visible, unaccepted) deletions that python-docx's
    own .text silently drops entirely. Use this instead of .text anywhere
    the actual content matters, which is everywhere in this file."""
    return "".join(text for _run, text, _link, _addr, _del in _iter_all_runs(paragraph))


def _full_cell_text(cell):
    """The tracked-changes-aware equivalent of a docx table cell's .text —
    joins _full_paragraph_text() across the cell's paragraphs instead of
    using the cell's own .text, which has the same silent-drop problem on
    insertions/deletions as paragraph.text does."""
    return "\n".join(_full_paragraph_text(p) for p in cell.paragraphs)


def _get_docx_numbering_formats(doc):
    """Returns a dict mapping numId -> True if that list numbering
    definition is genuinely numbered (decimal, lowerLetter, upperRoman,
    etc.) or False if it's a bullet, read directly from numbering.xml.
    Needed because Word's toolbar bullet/number buttons apply direct
    per-paragraph numPr formatting while leaving the paragraph under a
    generic style name ('List Paragraph', sometimes not even that) — for
    the overwhelming majority of real-world Word lists (anything created
    via the ribbon rather than by explicitly picking a legacy style named
    'List Bullet'/'List Number'), a style-name check alone can't tell a
    bulleted list from a numbered one, or see indent levels at all."""
    try:
        pkg = doc.part.package
    except Exception:
        return {}
    numbering_root = None
    for part in pkg.iter_parts():
        if part.partname == "/word/numbering.xml":
            try:
                numbering_root = etree.fromstring(part.blob)
            except Exception:
                return {}
            break
    if numbering_root is None:
        return {}

    abstract_is_numbered = {}
    for abstract_num in numbering_root.findall(qn("w:abstractNum")):
        abstract_id = abstract_num.get(qn("w:abstractNumId"))
        lvl0 = abstract_num.find(qn("w:lvl"))
        fmt_val = None
        if lvl0 is not None:
            numFmt = lvl0.find(qn("w:numFmt"))
            fmt_val = numFmt.get(qn("w:val")) if numFmt is not None else None
        abstract_is_numbered[abstract_id] = fmt_val not in (None, "bullet", "none")

    result = {}
    for num in numbering_root.findall(qn("w:num")):
        num_id = num.get(qn("w:numId"))
        abstract_ref = num.find(qn("w:abstractNumId"))
        abstract_id = abstract_ref.get(qn("w:val")) if abstract_ref is not None else None
        if abstract_id in abstract_is_numbered:
            result[num_id] = abstract_is_numbered[abstract_id]
    return result


def _get_paragraph_direct_list_info(paragraph, numbering_formats):
    """Returns (level, is_numbered) from real numPr formatting — the
    paragraph's own direct numPr when present, else its paragraph style's
    own numPr. Checking only the paragraph itself misses a real, common
    case: a custom paragraph style (anything not literally named 'List
    Bullet'/'List Number') that carries the numbering reference in the
    STYLE's own definition rather than as direct per-paragraph formatting
    — the style-name heuristic can't catch this either, since the style
    might be named anything at all. Returns None only when neither the
    paragraph nor its style has any numPr, in which case a caller should
    fall back to the style-*name* guess for the legacy 'List Bullet'/'List
    Number' built-in styles, which carry numbering a third way (through
    Word's separate style-to-list linkage rather than an explicit numPr
    in the style's own pPr)."""
    def numpr_from(pPr):
        if pPr is None:
            return None
        numPr = pPr.find(qn("w:numPr"))
        if numPr is None:
            return None
        ilvl_el = numPr.find(qn("w:ilvl"))
        level = int(ilvl_el.get(qn("w:val"))) if ilvl_el is not None else 0
        numId_el = numPr.find(qn("w:numId"))
        num_id = numId_el.get(qn("w:val")) if numId_el is not None else None
        is_numbered = numbering_formats.get(num_id, False) if num_id else False
        return (level, is_numbered)

    direct = numpr_from(paragraph._p.find(qn("w:pPr")))
    if direct is not None:
        return direct
    try:
        return numpr_from(_cached_paragraph_style(paragraph).element.find(qn("w:pPr")))
    except Exception:
        return None


def _iter_textbox_paragraphs(paragraph):
    """Yields each paragraph nested inside a floating text box embedded in
    this paragraph's runs. A text box's own paragraphs live inside a nested
    <w:txbxContent> deep within the run's drawing XML — a completely
    separate tree from the main document body — so normal paragraph
    iteration (doc.paragraphs, iter_block_items) never sees them at all,
    silently dropping the text box's entire content.

    A modern text box is typically wrapped in <mc:AlternateContent>: a
    <mc:Choice> branch with the real DrawingML shape, and a <mc:Fallback>
    branch holding a legacy VML re-representation of that exact same
    content for older software — confirmed directly by rendering one and
    checking LibreOffice uses the Choice branch, not the Fallback. Both
    branches use the identical <w:txbxContent> element name, so a plain
    descendant search finds the same text twice; this skips anything
    inside a Fallback branch specifically to avoid that duplication."""
    for run in paragraph.runs:
        for txbx_content in run._element.findall(".//" + qn("w:txbxContent")):
            if any(anc.tag == f"{{{_MC_NS}}}Fallback" for anc in txbx_content.iterancestors()):
                continue
            for p_el in txbx_content.findall(qn("w:p")):
                yield DocxParagraph(p_el, paragraph)


def _iter_inline_images(paragraph, doc):
    """Yields raw image bytes for each inline picture embedded in this
    paragraph, in document order — including images inside a hyperlink.
    python-docx's own paragraph.runs deliberately excludes hyperlink-wrapped
    runs, so walking iter_inner_content() (which covers both) is required
    here, not just runs, or a clickable image would be silently skipped.
    Also excludes any a:blip found inside an mc:Fallback branch — the same
    duplication risk confirmed for text boxes applies here: if a picture
    is ever wrapped in mc:AlternateContent (some newer image effects use
    this for compatibility), both the real and fallback branches could
    reference what's logically the same image, and this search wouldn't
    otherwise tell them apart."""
    for item in paragraph.iter_inner_content():
        runs = item.runs if type(item).__name__ == "Hyperlink" else [item]
        for run in runs:
            for blip in run._element.findall(".//" + qn("a:blip")):
                if any(anc.tag == f"{{{_MC_NS}}}Fallback" for anc in blip.iterancestors()):
                    continue
                r_id = blip.get(qn("r:embed"))
                if not r_id:
                    continue
                try:
                    yield doc.part.related_parts[r_id].blob
                except KeyError:
                    continue


_WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"

# OOXML's preset geometry names (the a:prstGeom "prst" attribute) are the
# same standardized vocabulary in both Word and PowerPoint's DrawingML —
# this maps the ones most likely to actually appear in a real document
# (flowchart boxes, callouts, arrows) to the closest MSO_SHAPE for
# recreating them as real, editable PowerPoint shapes rather than losing
# them or falling back to a plain rectangle for everything.
_DOCX_SHAPE_PRESET_MAP = {
    "rect": MSO_SHAPE.RECTANGLE,
    "roundRect": MSO_SHAPE.ROUNDED_RECTANGLE,
    "ellipse": MSO_SHAPE.OVAL,
    "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
    "rtTriangle": MSO_SHAPE.RIGHT_TRIANGLE,
    "diamond": MSO_SHAPE.DIAMOND,
    "pentagon": MSO_SHAPE.REGULAR_PENTAGON,
    "hexagon": MSO_SHAPE.HEXAGON,
    "chevron": MSO_SHAPE.CHEVRON,
    "rightArrow": MSO_SHAPE.RIGHT_ARROW,
    "leftArrow": MSO_SHAPE.LEFT_ARROW,
    "upArrow": MSO_SHAPE.UP_ARROW,
    "downArrow": MSO_SHAPE.DOWN_ARROW,
    "leftRightArrow": MSO_SHAPE.LEFT_RIGHT_ARROW,
    "upDownArrow": MSO_SHAPE.UP_DOWN_ARROW,
    "star4": MSO_SHAPE.STAR_4_POINT,
    "star5": MSO_SHAPE.STAR_5_POINT,
    "star6": MSO_SHAPE.STAR_6_POINT,
    "cloud": MSO_SHAPE.CLOUD,
    "heart": MSO_SHAPE.HEART,
    "lightningBolt": MSO_SHAPE.LIGHTNING_BOLT,
    "smileyFace": MSO_SHAPE.SMILEY_FACE,
    "wedgeRectCallout": MSO_SHAPE.RECTANGULAR_CALLOUT,
    "wedgeRoundRectCallout": MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT,
    "wedgeEllipseCallout": MSO_SHAPE.OVAL_CALLOUT,
}


def _iter_docx_shapes(paragraph):
    """Yields (mso_shape, fill_hex, text, width_emu, height_emu) for each
    DrawingML shape (a Word AutoShape — a callout, an arrow, a flowchart
    box) found in this paragraph. These are structurally distinct from
    a:blip picture references (_iter_inline_images) — a shape has no
    embedded image at all, just a preset outline filled with color and
    optionally holding its own text box, so the picture-extraction path
    never sees them and they'd otherwise be silently dropped entirely."""
    for run in paragraph.runs:
        for wsp in run._element.findall(".//{%s}wsp" % _WPS_NS):
            prst_el = wsp.find(".//" + qn("a:prstGeom"))
            prst = prst_el.get("prst") if prst_el is not None else None
            mso_shape = _DOCX_SHAPE_PRESET_MAP.get(prst, MSO_SHAPE.RECTANGLE)

            fill_hex = None
            fill_el = wsp.find(".//{%s}spPr/" % _WPS_NS + qn("a:solidFill") + "/" + qn("a:srgbClr"))
            if fill_el is not None:
                fill_hex = fill_el.get("val")

            text_parts = []
            for t in wsp.findall(".//{%s}txbx//" % _WPS_NS + qn("w:t")):
                if t.text:
                    text_parts.append(t.text)
            text = "".join(text_parts).strip()

            width_emu = height_emu = None
            ext_el = wsp.find(".//{%s}spPr/" % _WPS_NS + qn("a:xfrm") + "/" + qn("a:ext"))
            if ext_el is not None:
                try:
                    width_emu = int(ext_el.get("cx"))
                    height_emu = int(ext_el.get("cy"))
                except (TypeError, ValueError):
                    pass

            yield mso_shape, fill_hex, text, width_emu, height_emu


# ------------------------------------------------ PPTX auto-fix / cleanup
# Anchor point for the dynamic budget formulas below: empirically measured
# on a real uploaded deck (not assumed) — a 4.76"-tall, 11.5"-wide content
# placeholder at 26pt body text fits 7 short bullets cleanly, and visibly
# overflows at 9. Other decks' placeholders are scaled proportionally from
# this measured point rather than from an untested theoretical formula.
_PPTX_FIX_CAL_WIDTH_IN = 11.5
_PPTX_FIX_CAL_HEIGHT_IN = 4.76
_PPTX_FIX_CAL_SIZE_PT = 26
_PPTX_FIX_CAL_CHARS_PER_LINE = 48
_PPTX_FIX_CAL_MAX_EFFECTIVE_LINES = 11.2  # 7 bullets * ~1.6 effective lines each (text + spacing overhead)
_PPTX_FIX_TITLE_SIZE = Pt(40)
_PPTX_FIX_BODY_SIZE = Pt(26)


def _pptx_fix_chars_per_line(width_emu, size_pt):
    width_in = (width_emu / 914400) if width_emu else _PPTX_FIX_CAL_WIDTH_IN
    size = size_pt.pt if hasattr(size_pt, "pt") else size_pt
    return max(10, _PPTX_FIX_CAL_CHARS_PER_LINE * (width_in / _PPTX_FIX_CAL_WIDTH_IN) * (_PPTX_FIX_CAL_SIZE_PT / size))


def _pptx_fix_max_effective_lines(height_emu, size_pt):
    height_in = (height_emu / 914400) if height_emu else _PPTX_FIX_CAL_HEIGHT_IN
    size = size_pt.pt if hasattr(size_pt, "pt") else size_pt
    return max(3, _PPTX_FIX_CAL_MAX_EFFECTIVE_LINES * (height_in / _PPTX_FIX_CAL_HEIGHT_IN) * (_PPTX_FIX_CAL_SIZE_PT / size))


def _pptx_fix_effective_lines(text, chars_per_line):
    return max(1, -(-len(text) // int(chars_per_line))) + 0.6  # ceiling division + spacing overhead


def _pptx_fix_is_simple_text_shape(shape, title_shape):
    """A plain bullet-list placeholder — not a table, chart, picture, or
    the title itself. Only slides built entirely from these are eligible
    for splitting; anything else (an image, an embedded chart, a table)
    means the slide is left structurally alone so that element is never
    at risk of being orphaned from the text that refers to it."""
    if shape == title_shape:
        return False
    if not shape.has_text_frame:
        return False
    if shape.has_table or shape.has_chart:
        return False
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return False
    return True


def _pptx_fix_extract_bullets(shape):
    bullets = []
    for p in shape.text_frame.paragraphs:
        if not p.text.strip():
            continue
        runs = [(r.text, bool(r.font.bold), bool(r.font.italic)) for r in p.runs if r.text]
        if not runs:
            runs = [(p.text, False, False)]
        pPr = p._p.find(pptx_qn("a:pPr"))
        is_numbered = pPr is not None and pPr.find(pptx_qn("a:buAutoNum")) is not None
        bullets.append((p.level, runs, is_numbered))
    return bullets


def _pptx_fix_split_bullet_if_long(level, runs, numbered):
    if numbered:
        return [(level, runs, numbered)]  # never split a numbered step — it would break the sequence's meaning
    full_text = "".join(t for t, b, i in runs)
    if len(full_text.split()) <= _SENTENCE_SPLIT_WORD_THRESHOLD:
        return [(level, runs, numbered)]
    runs_data = [(t, b, i, False, False, None, None, None, False) for t, b, i in runs]
    groups = _split_runs_at_sentence_boundaries(runs_data)
    if len(groups) <= 1:
        return [(level, runs, numbered)]
    return [(level, [(t, bold, italic) for t, bold, italic, *_ in g], False) for g in groups]


def _pptx_fix_set_run_style(run, size, bold=None, italic=None):
    run.font.size = size
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic


_PPTX_FIX_ALIGN_MAP = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}


def _pptx_fix_apply_sizing_in_place(shape, size, bold_title=False, font_name=None, font_color=None, text_align=None):
    """Fixes font size and disables autofit on an existing shape without
    restructuring it — the safe path for any slide too complex to split
    (has an image, chart, table, or more than one text box), so a table
    or picture on that slide is never touched. font_name is only passed
    when a custom template's brand font should replace whatever the
    source file's own text used. font_color, when given, is a "RRGGBB"
    hex string (the natural form a color-picker input sends) — left
    None to leave the source file's existing text color untouched.
    text_align is one of "left"/"center"/"right", a paragraph-level
    property in PowerPoint's object model, so it's set once per
    paragraph rather than per run."""
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    align_value = _PPTX_FIX_ALIGN_MAP.get(text_align)
    for p in tf.paragraphs:
        if align_value is not None:
            p.alignment = align_value
        for r in p.runs:
            r.font.size = size
            if bold_title:
                r.font.bold = True
            if font_name:
                r.font.name = font_name
            if font_color:
                r.font.color.rgb = PptxRGBColor.from_string(font_color)


def _pptx_fix_move_slide_to(prs, slide, position):
    """Moves a slide (already added via add_slide, currently at the end)
    to a specific position in the deck — python-pptx's add_slide always
    appends, so continuation slides created by a split need to be
    relocated to immediately follow the slide they split from, rather
    than landing at the very end of the presentation."""
    xml_slides = prs.slides._sldIdLst
    slide_elements = list(xml_slides)
    xml_slides.remove(slide_elements[-1])
    xml_slides.insert(position, slide_elements[-1])


def _relative_luminance(r, g, b):
    """WCAG 2.x relative luminance from sRGB channel values (0-255)."""
    def channel(c):
        c = c / 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b)


def _contrast_ratio(rgb1, rgb2):
    l1, l2 = _relative_luminance(*rgb1), _relative_luminance(*rgb2)
    lighter, darker = max(l1, l2), min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


def _pptx_fix_get_slide_bg_rgb(slide):
    """Resolves a slide's effective background color for contrast
    checking, walking slide -> layout -> master -> a plain white
    fallback. Only handles an explicit solid RGB fill at each level
    (not gradients, pictures, or theme-color references) — those are
    left alone rather than guessed at, since contrast-fixing text
    against a background color that's actually a guess could easily
    make things worse, not better."""
    for source in (slide, slide.slide_layout, slide.slide_layout.slide_master):
        try:
            fill = source.background.fill
            if fill.type is not None and fill.fore_color.type == MSO_COLOR_TYPE.RGB:
                rgb = fill.fore_color.rgb
                return (rgb[0], rgb[1], rgb[2])
        except (AttributeError, TypeError, KeyError):
            continue
    return (255, 255, 255)  # the overwhelming majority default


def _pptx_fix_apply_contrast_fixes(prs):
    """Checks each text run with an explicit color against its slide's
    resolved background and snaps it to black or white (whichever gives
    stronger contrast) when it falls below WCAG AA's 4.5:1 threshold for
    normal text — a real, common accessibility failure most decks have
    without the author ever noticing, since it's invisible until someone
    actually needs the contrast. Runs with no explicit color (inheriting
    from the layout/theme instead) are left alone rather than guessed at."""
    for slide in prs.slides:
        bg_rgb = _pptx_fix_get_slide_bg_rgb(slide)
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    try:
                        if run.font.color.type != MSO_COLOR_TYPE.RGB:  # only an explicit RGB color is safe to evaluate
                            continue
                        rgb = run.font.color.rgb
                        text_rgb = (rgb[0], rgb[1], rgb[2])
                    except (AttributeError, TypeError):
                        continue
                    if _contrast_ratio(text_rgb, bg_rgb) < 4.5:
                        black_ratio = _contrast_ratio((0, 0, 0), bg_rgb)
                        white_ratio = _contrast_ratio((255, 255, 255), bg_rgb)
                        run.font.color.rgb = PptxRGBColor(0, 0, 0) if black_ratio >= white_ratio else PptxRGBColor(255, 255, 255)


def _pptx_fix_add_agenda_slide(prs, title_size, body_size, title_font, body_font,
                                title_color=None, body_color=None, title_align=None):
    """Inserts an agenda slide right after the title slide, listing each
    distinct section's own title verbatim — built entirely from titles
    the deck already has, nothing invented. Only added once the deck has
    enough distinct sections to actually be worth summarizing, and
    skipped (rather than spilling onto a second agenda slide) once there
    are more sections than a single overview slide can usefully list —
    an agenda that itself needs "(cont.)" defeats its own purpose."""
    # Conversions mirror the original: Academic-Home never adds a slide — not
    # even an agenda built from the deck's own titles.
    return
    slides = list(prs.slides)
    if len(slides) < 2:
        return
    section_titles = []
    for slide in slides[1:]:
        title_shape = slide.shapes.title
        if title_shape is None or not title_shape.has_text_frame:
            continue
        text = title_shape.text.strip()
        if not text or text.endswith("(cont.)"):
            continue
        section_titles.append(text)

    MIN_SECTIONS_FOR_AGENDA = 4
    MAX_AGENDA_ITEMS = 10
    if not (MIN_SECTIONS_FOR_AGENDA <= len(section_titles) <= MAX_AGENDA_ITEMS):
        return

    layout = _find_best_pptx_layout(prs, "content")
    agenda_slide = prs.slides.add_slide(layout)
    align_value = _PPTX_FIX_ALIGN_MAP.get(title_align)
    if agenda_slide.shapes.title is not None:
        agenda_slide.shapes.title.text = "Agenda"
        for p in agenda_slide.shapes.title.text_frame.paragraphs:
            if align_value is not None:
                p.alignment = align_value
            for r in p.runs:
                r.font.size = title_size
                r.font.bold = True
                if title_font:
                    r.font.name = title_font
                if title_color:
                    r.font.color.rgb = PptxRGBColor.from_string(title_color)
        agenda_slide.shapes.title.text_frame.auto_size = MSO_AUTO_SIZE.NONE

    body_ph = [s for s in agenda_slide.placeholders if s != agenda_slide.shapes.title and s.has_text_frame]
    if not body_ph:
        return
    tf = body_ph[0].text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.clear()
    for i, title_text in enumerate(section_titles):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run()
        r.text = title_text
        r.font.size = body_size
        if body_font:
            r.font.name = body_font
        if body_color:
            r.font.color.rgb = PptxRGBColor.from_string(body_color)
    _pptx_fix_move_slide_to(prs, agenda_slide, 1)


_SMART_EMPHASIS_RE = re.compile(
    r"\d+(?:\.\d+)?\s*(?:°\s*[CF]\b|%|(?:seconds?|minutes?|hours?|days?)\b)",
    re.IGNORECASE,
)
_SMART_EMPHASIS_COLOR = PptxRGBColor(0xC0, 0x39, 0x2B)  # a clear, warm red-orange — reads as "pay attention to this number"


def _pptx_fix_split_run_for_emphasis(run, paragraph):
    """Splits a single run's text at each key-data-point match (a
    temperature, duration, or percentage — the kind of number a reader's
    eye should catch immediately) into separate runs, bolding and
    color-accenting just the matched portion while the surrounding text
    keeps its original formatting untouched. The same visual "pop"
    modern AI-generated decks give numbers, applied to text the deck
    already has — nothing is reworded or invented, only re-emphasized."""
    text = run.text
    matches = list(_SMART_EMPHASIS_RE.finditer(text))
    if not matches:
        return
    insert_after_el = run._r
    pos = 0
    for m in matches:
        if m.start() > pos:
            plain_el = copy.deepcopy(run._r)
            insert_after_el.addnext(plain_el)
            plain_run = _Run(plain_el, paragraph)
            plain_run.text = text[pos:m.start()]
            insert_after_el = plain_el
        emph_el = copy.deepcopy(run._r)
        insert_after_el.addnext(emph_el)
        emph_run = _Run(emph_el, paragraph)
        emph_run.text = text[m.start():m.end()]
        emph_run.font.bold = True
        emph_run.font.color.rgb = _SMART_EMPHASIS_COLOR
        insert_after_el = emph_el
        pos = m.end()
    if pos < len(text):
        plain_el = copy.deepcopy(run._r)
        insert_after_el.addnext(plain_el)
        plain_run = _Run(plain_el, paragraph)
        plain_run.text = text[pos:]
    run._r.getparent().remove(run._r)


def _pptx_fix_apply_smart_emphasis(prs):
    """Applies smart auto-emphasis to every plain-text bullet placeholder
    across the deck — table cells, chart data, and titles are left
    alone, since this is specifically about making a key number stand
    out within otherwise-plain body prose, not decorating every text
    element in the file."""
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape == slide.shapes.title or not shape.has_text_frame:
                continue
            if shape.has_table or shape.has_chart:
                continue
            for para in shape.text_frame.paragraphs:
                # snapshot first - the loop body inserts new sibling runs,
                # which would otherwise be picked up again mid-iteration
                for run in list(para.runs):
                    if not run.font.bold:  # don't re-emphasize text already emphasized
                        _pptx_fix_split_run_for_emphasis(run, para)


def _pptx_fix_apply_transitions(prs):
    """Applies a uniform, subtle fade transition across every slide — a
    deck presented with no transitions at all (the overwhelming default
    for anything not built directly in a template gallery) reads as
    noticeably less polished than one with even a simple, consistent
    fade between every slide. Kept to a single, unobtrusive transition
    for the whole deck rather than mixing flashy ones per-slide, since
    a consistent, quiet transition reads as intentional design and a
    different one every slide reads as distracting default-clicking."""
    p_ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
    for slide in prs.slides:
        sld_el = slide._element
        existing = sld_el.find(f"{{{p_ns}}}transition")
        if existing is not None:
            sld_el.remove(existing)
        csld_el = sld_el.find(f"{{{p_ns}}}cSld")
        clrmapovr_el = sld_el.find(f"{{{p_ns}}}clrMapOvr")
        transition_el = sld_el.makeelement(f"{{{p_ns}}}transition", {"spd": "med"})
        fade_el = transition_el.makeelement(f"{{{p_ns}}}fade", {})
        transition_el.append(fade_el)
        anchor = clrmapovr_el if clrmapovr_el is not None else csld_el
        anchor.addnext(transition_el)


def _pptx_fix_try_numeric(text):
    """Parses a cell's text as a float, tolerating a leading currency
    symbol, thousands separators, and a trailing percent sign — the
    common ways a number shows up in a real table cell. Returns None if
    the text isn't recognizably numeric."""
    if not text:
        return None
    cleaned = text.strip().replace(",", "").replace("$", "").replace("%", "")
    try:
        return float(cleaned)
    except ValueError:
        return None


def _pptx_fix_detect_chart_shaped_table(grid):
    """Determines whether an existing table's data is actually a numeric
    chart in table form — a header row of series labels, a header
    column of categories, and a body that's uniformly numeric — the
    classic Excel-paste-into-PowerPoint pattern. Deliberately strict
    (every body cell must parse as a number, not just "most") to avoid
    misfiring on an ordinary table that happens to contain a few numbers
    (a schedule with times, a roster with ages) — a wrongly "smart"
    conversion that throws away real tabular content would be a worse
    outcome than leaving a genuine chart-shaped table exactly as it
    already is. Returns (categories, series_list) or None."""
    if len(grid) < 3 or len(grid[0]) < 2:
        return None  # need a header row plus at least 2 data rows, and at least one series
    n_cols = len(grid[0])
    if any(len(row) != n_cols for row in grid):
        return None  # ragged table (merged cells etc.) - not safe to reinterpret as chart data
    if n_cols > 7 or len(grid) > 11:
        return None  # too many series/categories to make a readable chart anyway

    header_row, body_rows = grid[0], grid[1:]
    if any(_pptx_fix_try_numeric(c) is not None for c in header_row[1:]):
        return None  # a numeric "label" means this probably isn't a labeled header row at all
    if any(_pptx_fix_try_numeric(row[0]) is not None for row in body_rows):
        return None  # a numeric "category" means this column isn't acting as row labels

    numeric_grid = []
    for row in body_rows:
        numeric_row = []
        for cell in row[1:]:
            val = _pptx_fix_try_numeric(cell)
            if val is None:
                return None  # strict: even one non-numeric body cell means this isn't chart data
            numeric_row.append(val)
        numeric_grid.append(numeric_row)

    categories = [row[0] for row in body_rows]
    series_names = header_row[1:]
    series_list = [
        (series_names[i], [numeric_grid[r][i] for r in range(len(body_rows))])
        for i in range(len(series_names))
    ]
    return categories, series_list


def _pptx_fix_convert_table_to_chart(slide, table_shape):
    """Replaces a chart-shaped table with a real, native, editable chart
    in the same position and size — the "modernized" version of pasting
    numbers into a table instead of building an actual chart. Leaves the
    table completely untouched on any failure (a shape that isn't
    actually a table, data that doesn't pass the strict chart-shape
    check, or a chart the API can't build from it), since the original
    table is always a safe, working fallback."""
    if not table_shape.has_table:
        return False
    grid = [[cell.text.strip() for cell in row.cells] for row in table_shape.table.rows]
    detected = _pptx_fix_detect_chart_shaped_table(grid)
    if detected is None:
        return False
    categories, series_list = detected

    chart_data = CategoryChartData()
    chart_data.categories = categories
    for name, values in series_list:
        chart_data.add_series(name or "Series", values)

    left, top, width, height = table_shape.left, table_shape.top, table_shape.width, table_shape.height
    try:
        graphic_frame = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, left, top, width, height, chart_data)
        chart = graphic_frame.chart
        chart.has_legend = len(series_list) > 1
        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
    except Exception:
        return False

    table_shape._element.getparent().remove(table_shape._element)
    return True


def _pptx_fix_reposition_offslide_pictures(prs):
    """Rescales and repositions any picture that extends off the visible
    slide area back to fully within the slide bounds, preserving its
    aspect ratio — deliberately scoped to this one objectively-broken
    state (an image can never be intentionally placed partly off the
    canvas) rather than any subjective "better layout" judgment about
    images that are simply positioned unusually but still fully
    visible, which the source author may well have placed on purpose."""
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                continue
            if not (shape.width and shape.height):
                continue
            off_slide = (
                shape.left < 0 or shape.top < 0
                or shape.left + shape.width > prs.slide_width
                or shape.top + shape.height > prs.slide_height
            )
            if not off_slide:
                continue
            max_w = prs.slide_width - PptxInches(0.5)
            max_h = prs.slide_height - PptxInches(0.5)
            scale = min(max_w / shape.width, max_h / shape.height, 1.0)
            shape.width = int(shape.width * scale)
            shape.height = int(shape.height * scale)
            shape.left = max(0, min(shape.left, prs.slide_width - shape.width))
            shape.top = max(0, min(shape.top, prs.slide_height - shape.height))


def _pptx_fix_force_white_bg_black_text(prs):
    """Forces every slide's background to solid white and every piece of
    text in the deck to solid black — an explicit, absolute request
    that overrides anything else in this pipeline (a custom template's
    brand palette, the smart-emphasis accent color, a slide's own
    inherited layout/master background), since it runs as the very last
    step. Bold from smart emphasis is left in place — it's the one part
    of that feature that isn't a color, so it doesn't conflict with
    "text is always black" the way the accent color would."""
    white, black = PptxRGBColor(255, 255, 255), PptxRGBColor(0, 0, 0)

    def blacken(text_frame):
        for para in text_frame.paragraphs:
            for run in para.runs:
                run.font.color.rgb = black

    for slide in prs.slides:
        try:
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = white
        except Exception:
            pass  # an unusual background fill type shouldn't sink the rest of the fix

        for shape in slide.shapes:
            if shape.has_text_frame:
                blacken(shape.text_frame)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        blacken(cell.text_frame)
            if shape.has_chart:
                chart = shape.chart
                try:
                    if chart.has_legend:
                        chart.legend.font.color.rgb = black
                except Exception:
                    pass
                for axis_attr in ("category_axis", "value_axis"):
                    try:
                        getattr(chart, axis_attr).tick_labels.font.color.rgb = black
                    except (AttributeError, ValueError):
                        pass  # not every chart type has both axes (e.g. a pie chart)


def pptx_to_pptx(src_path, out_dir, style=None, template_path=None,
                  title_size=None, body_size=None, title_font=None, body_font=None,
                  title_color=None, body_color=None, title_align=None):
    """Fixes the two most common problems in an already-existing
    PowerPoint file rather than converting from another format: PowerPoint
    silently auto-shrinking text on any slide with enough content to
    overflow (so a specified font size never actually renders at that
    size), and slides with more bullets than can actually fit at a
    readable size. Both were confirmed directly on a real uploaded deck —
    text auto-shrunk down to near-illegible, ten-item lists crammed onto
    one slide with most of it empty below. Rebuilds every slide with an
    explicit, generous font size and disabled autofit; text-only slides
    that don't fit are split into "(cont.)" continuation slides using
    their own existing content — nothing is invented, and sentence-level
    splitting is only applied to a single bullet that's a run-on paragraph
    long enough to be genuinely hard to read as one bullet. Slides with an
    image, chart, table, or more than one text box are fixed in place
    without being restructured, since splitting them could orphan a visual
    element from the text that refers to it.

    template_path, if given, is an uploaded corporate .potx/.pptx whose
    font scheme, color scheme, and title/body sizes are applied to the
    fixed deck. This deliberately does not migrate every slide onto the
    template's own layouts the way docx_to_pptx does for a fresh
    conversion — moving existing tables, charts, and images between two
    different presentations' layout structures is a much less reliable
    operation than placing fresh content once, and risking a corrupted
    or dropped chart to gain layout positions isn't a good trade for a
    tool whose whole purpose is fixing an existing file, not endangering
    it further. Brand fonts and colors carry through faithfully; the
    original file's own slide structure and placeholder positions do."""
    prs = _safe_load(Presentation, src_path)

    # title_size/body_size, when explicitly passed by a caller (as
    # opposed to arriving via the default/template precedence chain
    # below, which already deals in Pt(...) objects), are plain point
    # values — an int or float, the natural form for a UI's numeric
    # input to send. Converting them here, once, at the point they're
    # received, keeps that contract in exactly one place rather than
    # every call site needing to know to wrap them.
    if title_size is not None and not isinstance(title_size, Length):
        title_size = Pt(title_size)
    if body_size is not None and not isinstance(body_size, Length):
        body_size = Pt(body_size)

    # Precedence, lowest to highest: built-in defaults, then a template's
    # extracted style (if a template was uploaded), then whatever the
    # caller explicitly asked for — an explicit choice in the UI is a
    # more specific, intentional decision than either a hardcoded
    # default or a font size incidentally pulled from a template file.
    _title_size, _body_size, _title_font, _body_font = _PPTX_FIX_TITLE_SIZE, _PPTX_FIX_BODY_SIZE, None, None
    if template_path:
        template_prs = _safe_load(Presentation, template_path)
        template_content_layout = _find_best_pptx_layout(template_prs, "content")
        template_style = _extract_template_style(template_prs, template_content_layout)
        _title_size, _body_size = template_style["title_size"], template_style["body_size"]
        _title_font, _body_font = template_style["title_font"], template_style["body_font"]
        # Swapping the theme part's own XML (color scheme + font scheme)
        # gives the fixed file the template's actual brand palette and
        # font scheme without migrating a single shape between the two
        # presentations — far more reliable than trying to move tables,
        # charts, and images onto a different layout structure.
        try:
            template_theme_part = template_prs.slide_masters[0].part.part_related_by(_PPTX_RT.THEME)
            for master in prs.slide_masters:
                master.part.part_related_by(_PPTX_RT.THEME)._blob = template_theme_part.blob
        except Exception:
            pass  # a template with an unusual/missing theme part shouldn't sink the whole fix

    title_size = title_size if title_size is not None else _title_size
    body_size = body_size if body_size is not None else _body_size
    title_font = title_font if title_font is not None else _title_font
    body_font = body_font if body_font is not None else _body_font

    # ---- Pass 1: fix every title in place, and classify each slide ----
    plan = []  # (slide, is_simple, body_shape_or_None)
    for slide in prs.slides:
        title_shape = slide.shapes.title
        if title_shape is not None and title_shape.has_text_frame:
            _pptx_fix_apply_sizing_in_place(title_shape, title_size, bold_title=True, font_name=title_font, font_color=title_color, text_align=title_align)

        other_shapes = [s for s in slide.shapes if s != title_shape]
        # Upgrade any chart-shaped table (numeric data pasted into a
        # table instead of a real chart) into a native chart before
        # classifying the slide — the classification below needs to see
        # whatever the slide actually ends up with, not the pre-upgrade
        # shape list.
        for s in list(other_shapes):
            if s.has_table:
                _pptx_fix_convert_table_to_chart(slide, s)
        other_shapes = [s for s in slide.shapes if s != title_shape]
        text_shapes = [s for s in other_shapes if _pptx_fix_is_simple_text_shape(s, title_shape)]
        # A slide is eligible for splitting only if it's title + exactly
        # one plain bullet placeholder and nothing else at all — any
        # image, chart, table, or extra text box takes it out of the
        # running, since restructuring around those risks orphaning them.
        has_other_visual = any(
            s.shape_type == MSO_SHAPE_TYPE.PICTURE or s.has_table or s.has_chart
            for s in other_shapes
        )
        is_simple = len(text_shapes) == 1 and len(other_shapes) == 1 and not has_other_visual
        if not is_simple:
            for s in other_shapes:
                if s.has_text_frame:
                    _pptx_fix_apply_sizing_in_place(s, body_size, font_name=body_font, font_color=body_color)
            plan.append((slide, False, None))
        else:
            plan.append((slide, True, text_shapes[0]))

    # ---- Pass 2: for simple slides, split any that don't fit ----
    insertions = []  # (position_in_current_slide_list, new_slide_element_ref) filled in as we go
    slide_list = list(prs.slides)
    for idx, (slide, is_simple, body_shape) in enumerate(plan):
        if not is_simple:
            continue
        title_shape = slide.shapes.title
        base_title = title_shape.text.strip() if title_shape and title_shape.has_text_frame else ""
        # Continuation slides reuse this title exactly as written (nothing
        # added to it, nothing taken from it).
        bullets = _pptx_fix_extract_bullets(body_shape)
        expanded = []
        for level, runs, numbered in bullets:
            expanded.extend(_pptx_fix_split_bullet_if_long(level, runs, numbered))

        # Budget at the size the text will really be (the chosen size, or
        # the 26 pt default) — not always at 26 pt.
        chars_per_line = _pptx_fix_chars_per_line(body_shape.width, body_size or _PPTX_FIX_CAL_SIZE_PT)
        max_lines = _pptx_fix_max_effective_lines(body_shape.height, body_size or _PPTX_FIX_CAL_SIZE_PT)

        def is_section_header(gi):
            """A level-0, fully-bold bullet immediately followed by at
            least one indented (level>0) bullet reads as a genuine
            sub-section header within the slide — structurally distinct
            from an ordinary bullet that just happens to use bold for
            emphasis — mirroring how Heading 1/2 drives splitting for a
            DOCX source. Requires an actual indented child, not just
            "bold and top-level" alone, since that alone is too easy to
            misfire on a bullet a user simply emphasized. Text ending in
            a colon is excluded even when it otherwise matches — a
            colon overwhelmingly signals "here's a list within this same
            topic" ("Common sources:") rather than a new topic of its
            own, confirmed directly: without this, "Common sources:"
            under "Cross-Contamination" was pulled into its own slide,
            leaving both it and the parent slide thin and oddly split."""
            level, runs, numbered = expanded[gi]
            if level != 0 or numbered or not runs or not all(bold for _, bold, _ in runs):
                return False
            text = "".join(t for t, _, _ in runs).strip()
            if text.endswith(":"):
                return False
            if gi + 1 >= len(expanded):
                return False
            next_level, _, _ = expanded[gi + 1]
            return next_level > 0

        groups = [[]]
        group_lines = [0]
        group_titles = [None]  # None = use the slide's own title / "(cont.)" of it
        lines_used = 0
        for gi, (level, runs, numbered) in enumerate(expanded):
            text = "".join(t for t, b, i in runs)
            line_est = _pptx_fix_effective_lines(text, chars_per_line)
            starts_section = bool(groups[-1]) and is_section_header(gi)
            overflows = bool(groups[-1]) and lines_used + line_est > max_lines
            if starts_section or overflows:
                groups.append([])
                group_lines.append(0)
                group_titles.append(text.strip() if starts_section else None)
                lines_used = 0
                if starts_section:
                    continue  # this bullet becomes the new slide's title, not a bullet on it
            groups[-1].append((level, runs, numbered))
            lines_used += line_est
            group_lines[-1] = lines_used
        # A very small trailing group (e.g. one short leftover bullet that
        # only barely tipped past the budget) reads as an awkward,
        # near-empty slide — merging it back into the previous group and
        # accepting a small bounded overage looks better than that, the
        # same tradeoff already made for sentence-splitting. Not applied
        # when the small group is its own genuine section header, though
        # — that split was a deliberate structural choice, not incidental
        # overflow, so it stays even if the section itself is short.
        if len(groups) > 1 and group_lines[-1] < 2.5 and group_titles[-1] is None:
            groups[-2].extend(groups[-1])
            groups.pop()
            group_lines.pop()
            group_titles.pop()

        # rewrite the original slide's body with just the first group
        tf = body_shape.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.clear()
        first_group = groups[0]
        for i, (level, runs, numbered) in enumerate(first_group):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.level = min(level, 4)
            if numbered:
                pPr = p._p.get_or_add_pPr()
                pPr.append(pPr.makeelement(pptx_qn("a:buAutoNum"), {"type": "arabicPeriod"}))
            for text, bold, italic in runs:
                r = p.add_run()
                r.text = text
                r.font.size = body_size
                r.font.bold = bold
                r.font.italic = italic
                if body_font:
                    r.font.name = body_font
                if body_color:
                    r.font.color.rgb = PptxRGBColor.from_string(body_color)

        if len(groups) == 1:
            continue

        # additional groups become new slides — a section-header-triggered
        # group uses that header's own text as its title (the analogue of
        # a DOCX Heading 1/2 becoming a new slide's title); an
        # overflow-triggered group falls back to "(cont.)" of the
        # original slide's title, using the same layout as the slide it
        # split from so images/backgrounds in a custom template carry
        # over correctly
        layout = slide.slide_layout
        insert_position = idx + 1 + sum(insertions_done for pos, insertions_done in insertions if pos <= idx)
        for gi, group in enumerate(groups[1:], 1):
            new_slide = prs.slides.add_slide(layout)
            if new_slide.shapes.title is not None:
                new_slide.shapes.title.text = group_titles[gi] or base_title   # the original title, nothing added
                align_value = _PPTX_FIX_ALIGN_MAP.get(title_align)
                for p in new_slide.shapes.title.text_frame.paragraphs:
                    if align_value is not None:
                        p.alignment = align_value
                    for r in p.runs:
                        r.font.size = title_size
                        r.font.bold = True
                        if title_font:
                            r.font.name = title_font
                        if title_color:
                            r.font.color.rgb = PptxRGBColor.from_string(title_color)
                new_slide.shapes.title.text_frame.auto_size = MSO_AUTO_SIZE.NONE
            new_body = [s for s in new_slide.placeholders if s.placeholder_format.idx == body_shape.placeholder_format.idx]
            new_body = new_body[0] if new_body else [s for s in new_slide.placeholders if s != new_slide.shapes.title][0]
            ntf = new_body.text_frame
            ntf.word_wrap = True
            ntf.auto_size = MSO_AUTO_SIZE.NONE
            ntf.clear()
            for i, (level, runs, numbered) in enumerate(group):
                p = ntf.paragraphs[0] if i == 0 else ntf.add_paragraph()
                p.level = min(level, 4)
                if numbered:
                    pPr = p._p.get_or_add_pPr()
                    pPr.append(pPr.makeelement(pptx_qn("a:buAutoNum"), {"type": "arabicPeriod"}))
                for text, bold, italic in runs:
                    r = p.add_run()
                    r.text = text
                    r.font.size = body_size
                    r.font.bold = bold
                    r.font.italic = italic
                    if body_font:
                        r.font.name = body_font
                    if body_color:
                        r.font.color.rgb = PptxRGBColor.from_string(body_color)
            _pptx_fix_move_slide_to(prs, new_slide, insert_position)
            insert_position += 1
        insertions.append((idx, len(groups) - 1))

    _pptx_fix_add_agenda_slide(prs, title_size, body_size, title_font, body_font,
                                title_color=title_color, body_color=body_color, title_align=title_align)
    _pptx_fix_apply_smart_emphasis(prs)
    _pptx_fix_reposition_offslide_pictures(prs)
    _pptx_fix_apply_transitions(prs)
    # Skipped entirely once the caller has explicitly chosen a text
    # color — this step's whole purpose is forcing solid black text on
    # solid white as an unconditional safety default, deliberately
    # overriding even a template's own brand palette (see its own
    # docstring). That's the right default when nobody's asked for
    # anything else, but it silently overwrote a user's own explicit
    # title/body color choice — confirmed directly: title_color kept
    # coming back 000000 regardless of what was requested, traced all
    # the way to this exact call. Leaving the background alone too,
    # not just the text, avoids a worse failure mode than the original
    # bug: forcing white behind a color someone explicitly picked (say,
    # white title text) would make it invisible instead of merely wrong.
    if not title_color and not body_color:
        _pptx_fix_force_white_bg_black_text(prs)

    out_path = os.path.join(out_dir, "fixed.pptx")
    os.makedirs(out_dir, exist_ok=True)
    prs.save(out_path)
    return out_path


def _pptx_set_footer(slide, text):
    """The slide's own footer (its layout's footer placeholder, filled in) —
    where a Word document's page header and footer lines belong on slides."""
    lp = next((p for p in slide.slide_layout.placeholders if p.placeholder_format.type == PP_PLACEHOLDER.FOOTER), None)
    if lp is None or not text:
        return False
    try:
        slide.shapes.clone_placeholder(lp)
    except Exception:
        return False
    ph = next((p for p in slide.placeholders if p.placeholder_format.type == PP_PLACEHOLDER.FOOTER), None)
    if ph is None:
        return False
    tf = ph.text_frame
    tf.text = ""
    _pptx_add_text(tf.paragraphs[0], text)
    return True


def docx_to_pptx(src_path, out_dir, style="plain", template_path=None):
    """style selects one of this file's own built-in visual themes
    (plain/minimal/academic/bold/classic) and is ignored when template_path is
    given. template_path, if provided, is a path to an uploaded .potx (or
    .pptx used as a template) — a corporate brand template — whose own
    masters, layouts, fonts, and colors are used instead, so the
    converted content maps into the uploaded template's own pre-styled
    layouts rather than one of this file's built-in looks."""
    _run_style_format_cache.clear()
    _paragraph_style_object_cache.clear()
    if template_path:
        prs = _safe_load(Presentation, template_path)
        title_layout = _find_best_pptx_layout(prs, "content")
        blank_layout = _find_best_pptx_layout(prs, "blank")
        template_style = _extract_template_style(prs, title_layout)
        # A .potx/.pptx template's own slide size is part of its brand
        # design (many corporate templates are intentionally 4:3, or a
        # custom size) — overwriting it the way the no-template path does
        # would fight the very template the user uploaded to preserve.
    else:
        template_style = PPTX_TEMPLATES.get(style, PPTX_TEMPLATES["plain"])
        prs = Presentation()
        prs.slide_width = PptxInches(13.333)
        prs.slide_height = PptxInches(7.5)
        _fit_layouts_to_widescreen(prs)
        if template_style.get("plain"):
            _plain_pptx_geometry(prs)
        title_layout = prs.slide_layouts[1]  # title + content
        blank_layout = prs.slide_layouts[6]

    doc = _safe_load(Document, src_path)
    numbering_formats = _get_docx_numbering_formats(doc)

    MAX_LINES_PER_SLIDE = 10  # estimated wrapped-line budget, not a flat bullet count (see
    # _estimate_line_count) — accounts for bullets of very different lengths now that long
    # paragraphs get split into shorter, sentence-level bullets rather than staying as one
    state = {"slide": None, "body_tf": None, "bullet_count": 0, "lines_used": 0, "section_title": None,
             "used": 0.0, "room": 0.0, "last_text": "", "num": {}, "box_seq": {}}
    PLAIN = bool(template_style.get("plain"))
    # The document's own title (its file property), when it has a real one — the topic of
    # anything that comes before the document's first heading.
    try:
        doc_topic = _clean_doc_title(doc.core_properties.title)
    except Exception:
        doc_topic = ""

    def _cont(title):
        """A continuation slide repeats its section's own title exactly as
        written — nothing added, nothing taken away."""
        return (title or "").strip()

    def _continuation_title():
        # Bullets that arrive with no slide open (right after a picture or a
        # table) used to get a generic "Overview" title, which told the
        # audience nothing. They continue the current section instead.
        return _cont(state["section_title"]) if state["section_title"] else ""

    def _slide_title_for(text):
        """A slide's title (R22): its section's own heading; where the document has none yet,
        the document's title, else a short topic in this text's own words. Returns (title,
        derived) — a derived title is marked on the slide, so a handout made from the slides
        leaves it out."""
        if state["section_title"]:
            return _cont(state["section_title"]), False
        t = doc_topic or _topic_from_text(text)
        return t, bool(t)

    def _plain_body_below_title(s, grow):
        body = s.placeholders[1]
        if grow:
            L, T, W, H = body.left, body.top, body.width, body.height
            body.left, body.top, body.width, body.height = L, T + grow, W, max(H - grow, int(PptxInches(1.2)))
        state["room"] = (body.height - _PPTX_INSETS) / _EMU_PT * 0.95

    def new_slide(title_text, topic=False):
        # A heading that produced nothing (common when a document is really a
        # handout of an earlier deck) used to leave an empty slide behind, and
        # a repeated heading produced the same title two or three times in a
        # row with nothing between them. An empty slide is reused instead: a
        # lone section heading still gets its divider slide, but the title
        # never repeats over consecutive empty ones.
        current = state["slide"]
        if current is not None and state["bullet_count"] == 0 and len(current.shapes) <= 2:
            existing = (current.shapes.title.text if current.shapes.title is not None else "").strip().lower()
            derived = _slide_tag_of(current) == "Topic"
            if existing == (title_text or "").strip().lower() or not existing or derived:
                if current.shapes.title is not None and current.shapes.title.text != (title_text or ""):
                    current.shapes.title.text = title_text or ""
                    # its words were replaced, so its look goes back on (and Plain fits it)
                    grow = _pptx_restyle_title(current, template_style)
                    if PLAIN and state["body_tf"] is not None:
                        _plain_body_below_title(current, grow)
                if topic:
                    _slide_tag(current, "Topic")
                elif derived:
                    current._element.cSld.attrib.pop("name", None)
                return current, state["body_tf"]
        s = prs.slides.add_slide(title_layout)
        s.shapes.title.text = title_text or ""
        grow = _style_pptx_slide(s, template_style)
        if topic:
            _slide_tag(s, "Topic")
        tf = s.placeholders[1].text_frame
        tf.clear()
        # The slide master's body placeholder inherits <a:normAutofit/> by
        # default — PowerPoint silently shrinks text to fit whenever it
        # overflows, which would quietly undo the larger, more readable
        # font sizes set below the moment a real slide has enough content
        # to trigger it. Disabling this keeps the specified size honest;
        # MAX_LINES_PER_SLIDE below is tuned to actually fit at that
        # size instead, rather than leaning on autofit to paper over it.
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        if PLAIN:
            _plain_body_below_title(s, grow)
            _pptx_body_list_style(tf, template_style, (template_style["body_size"].pt,) + (_PLAIN_SUB_PT,) * 4)
        state["slide"], state["body_tf"], state["bullet_count"], state["lines_used"] = s, tf, 0, 0
        state["used"], state["box_seq"] = 0.0, {}
        return s, tf

    # ---- Plain: each paragraph measured; what doesn't fit carries on on the next slide ----
    def _plain_size(level):
        return template_style["body_size"].pt if level == 0 else _PLAIN_SUB_PT

    def _plain_width(level, kind):
        body = state["slide"].placeholders[1]
        w = (body.width - 2 * _PPTX_INSETS) / _EMU_PT
        if kind != "plain":
            w -= _PPTX_MARL[min(level, 4)] / _EMU_PT
        return w * 0.95

    def _plain_height(text, level, kind, first, size=None):
        size = size or _plain_size(level)
        n = _calibri_lines(text, size, _plain_width(level, kind))
        return n * _PLAIN_LINE * size + (0 if first else (_PLAIN_GAP_PT if level == 0 else 6))

    def _plain_next_slide():
        cur = state["slide"]
        title = (cur.shapes.title.text if (cur is not None and cur.shapes.title is not None) else "") or ""
        if not title.strip():
            title, derived = _slide_title_for(state["last_text"])
            new_slide(title, topic=derived)
            return
        new_slide(_cont(title))

    def _plain_write(runs_data, level, kind, align, join=False, size=None):
        body_tf = state["body_tf"]
        first = state["bullet_count"] == 0
        if first and body_tf.paragraphs[0].text == "" and len(body_tf.paragraphs) == 1:
            p = body_tf.paragraphs[0]
        else:
            p = body_tf.add_paragraph()
        p.level = min(level, 4)
        if align is not None:
            p.alignment = align
        pPr = p._p.get_or_add_pPr()
        counts, seq = state["num"], state["box_seq"]
        if kind == "number" and not join:
            counts[level] = counts.get(level, 0) + 1
            for deeper in [k for k in counts if k > level]:
                del counts[deeper]
            num = pPr.makeelement(qn("a:buAutoNum"), {"type": "arabicPeriod"})
            if counts[level] > 1 and not seq.get(level):
                num.set("startAt", str(counts[level]))     # a list carried on from the slide before
            pPr.append(num)
            seq[level] = True
        else:
            if kind == "plain" or join:
                pPr.set("marL", "0" if kind == "plain" else str(_PPTX_MARL[min(level, 4)]))
                pPr.set("indent", "0")
                pPr.append(pPr.makeelement(qn("a:buNone"), {}))
            for k in [k for k in seq if k >= level]:
                seq[k] = False
            if not join and level == 0:
                counts.clear()           # the list has ended
        pt = size or _plain_size(level)
        captured = []
        for text, bold, italic, underline, is_link, address, hex_color, size_pt, is_deleted in runs_data:
            for seg_i, seg in enumerate(text.split("\n")):
                if seg_i:
                    p.add_line_break()
                if not seg:
                    continue
                r = p.add_run()
                r.text = seg
                if is_link and address:
                    try:
                        r.hyperlink.address = address
                    except Exception:
                        pass
                if is_deleted:
                    r._r.get_or_add_rPr().set("strike", "sngStrike")
                captured.append((r, bold, italic, underline, is_link, hex_color))
        _style_pptx_body_paragraph(p, template_style, size=Pt(pt))
        # the document's own emphasis, links and (readable) colours — its small sizes are not kept
        for r, bold, italic, underline, is_link, hex_color in captured:
            r.font.bold, r.font.italic, r.font.underline = bold, italic, underline
            if size:
                r.font.size = Pt(size)
            if is_link:
                r.font.color.rgb = PptxRGBColor(0x05, 0x63, 0xC1)
            elif hex_color and _plain_ink_ok(hex_color):
                try:
                    r.font.color.rgb = PptxRGBColor.from_string(hex_color)
                except Exception:
                    pass
        p.line_spacing = 1.0
        p.space_before = Pt(0 if first else (_PLAIN_GAP_PT if level == 0 else 6))
        if join:
            _pptx_mark_runs([c[0] for c in captured], "join")    # invisibly: the rest of the paragraph before
        state["bullet_count"] += 1
        state["lines_used"] += 1

    def _plain_place(runs_data, level, kind, align):
        text = "".join(e[0] for e in runs_data)
        if text.strip():
            state["last_text"] = text
        if state["slide"] is None:
            title, derived = _slide_title_for(text)
            new_slide(title, topic=derived)
        first = state["bullet_count"] == 0
        alone = _plain_height(text, level, kind, True)
        if alone <= state["room"]:
            here = _plain_height(text, level, kind, first)
            if not first and state["used"] + here > state["room"]:
                _plain_next_slide()
                here = alone
            _plain_write(runs_data, level, kind, align)
            state["used"] += here
            return
        # Taller than a whole slide: it starts on a fresh slide and is split at the ends of
        # sentences; each part after the first carries on, under the same title, on the next.
        if not first:
            _plain_next_slide()
        parts = _split_runs_at_sentence_boundaries(runs_data)
        room = state["room"]
        chunks, cur, cur_text = [], [], ""
        for part in parts:
            part_text = "".join(e[0] for e in part)
            trial = (cur_text + " " + part_text) if cur else part_text
            if cur and _plain_height(trial, level, kind, True) > room:
                chunks.append(cur)
                cur, trial = [], part_text
            if cur:          # the space between two sentences, in the earlier one's style (never a link)
                sp = cur[-1]
                cur.append((" ", sp[1], sp[2], sp[3], False, None, sp[6], sp[7], sp[8]))
            cur = cur + list(part)
            cur_text = trial
        if cur:
            chunks.append(cur)
        for k, chunk in enumerate(chunks):
            if k:
                _plain_next_slide()
            chunk_text = "".join(e[0] for e in chunk)
            size = None
            if _plain_height(chunk_text, level, kind, True) > state["room"]:
                # one sentence longer than a whole slide: the only words ever set smaller
                size = _plain_size(level)
                while size > 16 and _plain_height(chunk_text, level, kind, True, size) > state["room"]:
                    size -= 2
            _plain_write(chunk, level, kind, align, join=k > 0, size=size)
            state["used"] += min(_plain_height(chunk_text, level, kind, True, size), state["room"])

    def add_bullet(para, level=0, numbered=False, plain=False):
        if state["slide"] is None and not PLAIN:
            title, derived = _slide_title_for(_full_paragraph_text(para))
            new_slide(title, topic=derived)

        # iter_inner_content() (like .text and .runs) silently drops any
        # text wrapped in a tracked-change insertion or deletion —
        # _iter_all_runs() is the tracked-changes-aware replacement, and
        # since every item it yields is already a real Run, the hyperlink
        # special-casing this used to need is gone too.
        runs_data = []  # (text, bold, italic, underline, is_link, address, hex_color, size_pt, is_deleted)
        for src_run, text, is_link, address, is_deleted in _iter_all_runs(para):
            if not text:
                continue
            style_bold, style_italic, style_underline, style_color, style_size = _effective_run_format(src_run)
            runs_data.append((
                text, bool(style_bold), bool(style_italic),
                True if is_link else bool(style_underline),
                is_link, address, (None if is_link else style_color), style_size, is_deleted,
            ))
        if not runs_data:
            fallback_text = _full_paragraph_text(para).strip()
            if fallback_text:
                runs_data = [(fallback_text, False, False, False, False, None, None, None, False)]

        full_text = "".join(entry[0] for entry in runs_data)
        # Long, multi-sentence paragraphs are split into separate, shorter
        # bullets at real sentence boundaries — see _split_runs_at_sentence_
        # boundaries. Numbered items are left whole since splitting one
        # would break the numbering's meaning (item "3" becoming two
        # separate bullets makes no sense as a numbered step).
        # A paragraph stays one paragraph — splitting it into sentence bullets
        # meant converting back gave a different document.
        groups = [runs_data] if runs_data else []

        pptx_align = _docx_align_to_pptx(_effective_paragraph_alignment(para))
        if PLAIN:
            if runs_data:
                _plain_place(runs_data, level, "number" if numbered else ("plain" if plain else "bullet"), pptx_align)
            return
        if full_text.strip():
            state["last_text"] = full_text
        for group in groups:
            group_text = "".join(entry[0] for entry in group)
            group_lines = _estimate_line_count(group_text)
            # A flat bullet-count cap doesn't account for how much a
            # bullet's text actually wraps (see _estimate_line_count) —
            # overflow onto a "(cont.)" slide is judged by estimated
            # vertical space used instead, so several short bullets and a
            # couple of long ones are both handled correctly rather than
            # just counted the same.
            if state["bullet_count"] > 0 and state["lines_used"] + group_lines > MAX_LINES_PER_SLIDE:
                current_title = state["slide"].shapes.title.text or ""
                new_slide(_cont(current_title))
            body_tf = state["body_tf"]
            if body_tf.paragraphs[0].text == "" and len(body_tf.paragraphs) == 1 and state["bullet_count"] == 0:
                p = body_tf.paragraphs[0]
            else:
                p = body_tf.add_paragraph()
            p.level = min(level, 4)
            if pptx_align is not None:
                p.alignment = pptx_align
            if numbered:
                pPr = p._p.get_or_add_pPr()
                pPr.append(pPr.makeelement(qn("a:buAutoNum"), {"type": "arabicPeriod"}))
            elif plain:
                # An ordinary paragraph shows as an ordinary paragraph — no bullet added.
                pPr = p._p.get_or_add_pPr()
                pPr.set("marL", "0")
                pPr.set("indent", "0")
                pPr.append(pPr.makeelement(qn("a:buNone"), {}))

            captured = []  # (pptx_run, bold, italic, underline, is_link, hex_color, size_pt, is_deleted)
            for text, bold, italic, underline, is_link, address, hex_color, size_pt, is_deleted in group:
                for seg_i, seg in enumerate(text.split("\n")):
                    if seg_i:
                        p.add_line_break()       # a line break in the document stays a line break
                    if not seg:
                        continue
                    r = p.add_run()
                    r.text = seg
                    if is_link and address:
                        try:
                            r.hyperlink.address = address
                        except Exception:
                            pass
                    if is_deleted:
                        r._r.get_or_add_rPr().set("strike", "sngStrike")
                    captured.append((r, bold, italic, underline, is_link, hex_color, size_pt, is_deleted))
            _style_pptx_body_paragraph(p, template_style)
            # Re-apply emphasis after the shared style pass, since that pass
            # sets font attributes on every run and would otherwise stomp the
            # per-run formatting (and hyperlink coloring) just captured above.
            for r, bold, italic, underline, is_link, hex_color, size_pt, is_deleted in captured:
                r.font.bold = bold
                r.font.italic = italic
                r.font.underline = underline
                if size_pt is not None:
                    r.font.size = Pt(size_pt)
                if is_link:
                    try:
                        r.font.color.rgb = PptxRGBColor(0x05, 0x63, 0xC1)
                    except Exception:
                        pass
                elif hex_color:
                    try:
                        r.font.color.rgb = PptxRGBColor.from_string(hex_color)
                    except Exception:
                        pass
            state["bullet_count"] += 1
            state["lines_used"] += group_lines

    def add_subheading(text, level=3):
        """Heading 2/3 in the source document reads as a subsection within
        the current topic, not a brand new topic — giving it a whole new
        slide (the old behavior for every heading level) fragmented what's
        usually meant to be one cohesive slide into several thin ones. This
        keeps it on the current slide as a bolded standalone line instead."""
        if state["slide"] is None:
            new_slide(text)
            return
        if PLAIN:
            size = template_style["subhead_size"].pt
            width = (state["slide"].placeholders[1].width - 2 * _PPTX_INSETS) / _EMU_PT * 0.95
            first = state["bullet_count"] == 0
            need = _calibri_lines(text, size, width, True) * _PLAIN_LINE * size
            follow = _PLAIN_LINE * template_style["body_size"].pt + _PLAIN_GAP_PT   # with a line of what follows
            if not first and state["used"] + need + 18 + follow > state["room"]:
                _plain_next_slide()
                first = state["bullet_count"] == 0
            body_tf = state["body_tf"]
            if first and body_tf.paragraphs[0].text == "" and len(body_tf.paragraphs) == 1:
                p = body_tf.paragraphs[0]
            else:
                p = body_tf.add_paragraph()
            p.level = 0
            _pptx_add_text(p, text)
            _pptx_mark_runs(p.runs, f"h{level}")     # invisibly: this line was a Heading <level>
            pPr = p._p.get_or_add_pPr()
            pPr.set("marL", "0")
            pPr.set("indent", "0")
            pPr.append(pPr.makeelement(qn("a:buNone"), {}))
            _style_pptx_body_paragraph(p, template_style, size=template_style["subhead_size"])
            for run in p.runs:
                run.font.bold = True
            p.line_spacing = 1.0
            p.space_before = Pt(0 if first else 18)
            state["used"] += need + (0 if first else 18)
            state["bullet_count"] += 1
            state["lines_used"] += 1
            state["num"].clear()
            state["box_seq"] = {}
            state["last_text"] = text
            return
        # A subheading's larger font and extra space-before take up more
        # visual room than its line count alone suggests, so it's weighted
        # a bit heavier than a plain estimated-line count would give it.
        subhead_lines = _estimate_line_count(text) + 1
        if state["bullet_count"] > 0 and state["lines_used"] + subhead_lines > MAX_LINES_PER_SLIDE:
            current_title = state["slide"].shapes.title.text or ""
            new_slide(_cont(current_title))
        body_tf = state["body_tf"]
        if body_tf.paragraphs[0].text == "" and len(body_tf.paragraphs) == 1 and state["bullet_count"] == 0:
            p = body_tf.paragraphs[0]
        else:
            p = body_tf.add_paragraph()
        p.level = 0
        _pptx_add_text(p, text)
        _pptx_mark_runs(p.runs, f"h{level}")     # invisibly: this line was a Heading <level>
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", "0")
        pPr.set("indent", "0")
        pPr.append(pPr.makeelement(qn("a:buNone"), {}))
        if state["bullet_count"] > 0:
            p.space_before = Pt(18)
        _style_pptx_body_paragraph(p, template_style, size=template_style["subhead_size"])
        for run in p.runs:
            run.font.bold = True
        state["bullet_count"] += 1
        state["lines_used"] += subhead_lines

    def add_table_slide(table, carried_title=None, own=False):
        src_rows = [list(row.cells) for row in table.rows]
        rows = [[_full_cell_text(cell).strip() for cell in row] for row in src_rows]
        keep = [i for i, r in enumerate(rows) if any(r)]
        if not keep:
            return
        rows = [rows[i] for i in keep]
        src_rows = [src_rows[i] for i in keep]
        n_rows, n_cols = len(rows), max(len(r) for r in rows)
        topic = False
        if not carried_title:
            # R22: a table the document gave no heading is titled by its own caption, the
            # document's title, or the words just before it
            carried_title = _docx_table_caption(table) or doc_topic or _topic_from_text(state["last_text"])
            topic = bool(carried_title)
        if PLAIN:
            _plain_table_slides(rows, src_rows, n_rows, n_cols, carried_title, topic, own)
            return

        # carried_title (the heading immediately preceding this table, if
        # any) comes from the caller's own lookahead, not from deleting an
        # already-created slide here — confirmed directly that python-pptx's
        # documented slide-deletion recipe only detaches the <p:sldId>
        # reference and leaves the underlying part/relationship orphaned,
        # which silently corrupted the saved file (a duplicate slide part
        # name) the moment a second such table appeared later in the same
        # document. Never creating the empty slide in the first place
        # avoids that whole class of problem.
        # A small table right after a short intro line ("A vendor recorded
        # weekly sales at different prices:") goes on that same slide, under
        # the line — rather than leaving the intro alone on one slide and the
        # table, untitled, on the next.
        cur = state["slide"]
        cur_body = cur.placeholders[1] if (cur is not None and len(cur.placeholders) > 1) else None
        placed_here = False
        if cur_body is not None and 0 < state["bullet_count"] and state["lines_used"] <= 3 and not own:
            text_h = int(PptxInches(0.48) * state["lines_used"] + PptxInches(0.35))
            need = int(PptxInches(0.5) * n_rows)
            if need <= cur_body.height - text_h:
                L, T, W = cur_body.left, cur_body.top, cur_body.width
                cur_body.left, cur_body.top, cur_body.width, cur_body.height = L, T, W, text_h
                s = cur
                left, top, width, height = L, T + text_h, W, need
                placed_here = True
        if not placed_here:
            s = prs.slides.add_slide(title_layout if carried_title else blank_layout)
            if carried_title:
                s.shapes.title.text = carried_title
                _style_pptx_slide(s, template_style)
                if topic:
                    _slide_tag(s, "Topic")
                # The title layout's own body placeholder isn't needed here —
                # the table below is the slide's actual content — and leaving
                # an empty placeholder box behind would just be visual clutter.
                body_ph = s.placeholders[1] if len(s.placeholders) > 1 else None
                if body_ph is not None:
                    body_ph._element.getparent().remove(body_ph._element)
                left, top = PptxInches(0.6), PptxInches(1.7)
                width, height = prs.slide_width - PptxInches(1.2), prs.slide_height - PptxInches(2.3)
            else:
                left, top = PptxInches(0.6), PptxInches(0.6)
                width, height = prs.slide_width - PptxInches(1.2), prs.slide_height - PptxInches(1.2)
            # Rows sized to their content instead of stretched to fill the
            # slide: a six-row table used to get ~0.9in-tall rows.
            height = min(height, int(PptxInches(0.55) * n_rows))
        gtable = s.shapes.add_table(n_rows, n_cols, left, top, width, height).table

        merges = _find_docx_merges(src_rows, n_cols)
        # A merged span's duplicated cells would otherwise each get the same
        # text written before merging, concatenating it multiple times into
        # the final merged cell — only the top-left origin of each span
        # should actually receive the text.
        skip_cells = {(r, c) for (min_r, min_c, max_r, max_c) in merges
                      for r in range(min_r, max_r + 1) for c in range(min_c, max_c + 1)
                      if (r, c) != (min_r, min_c)}

        # Cell text had no size of its own and fell back to a small default
        # (about 12pt as rendered) — hard to read across a room. Sized to the
        # table instead, in the deck's own body font.
        cell_pt = Pt(20 if (n_rows <= 6 and n_cols <= 4) else 16 if (n_rows <= 10 and n_cols <= 6) else 13)
        for r_idx, row in enumerate(rows):
            for c_idx in range(n_cols):
                if (r_idx, c_idx) in skip_cells:
                    continue
                cell = gtable.cell(r_idx, c_idx)
                cell.text = row[c_idx] if c_idx < len(row) else ""
                for p in cell.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.size = cell_pt
                        if template_style.get("body_font"):
                            run.font.name = template_style["body_font"]
                if r_idx == 0:
                    for p in cell.text_frame.paragraphs:
                        for run in p.runs:
                            run.font.bold = True
                if c_idx < len(src_rows[r_idx]):
                    shade = _get_docx_cell_shading(src_rows[r_idx][c_idx])
                    if shade:
                        try:
                            cell.fill.solid()
                            cell.fill.fore_color.rgb = PptxRGBColor.from_string(shade)
                        except Exception:
                            pass
        for min_r, min_c, max_r, max_c in merges:
            try:
                gtable.cell(min_r, min_c).merge(gtable.cell(max_r, max_c))
            except Exception:
                pass
        # A loose paragraph appearing right after a table should land on a
        # fresh slide rather than silently reusing the table slide.
        state["slide"], state["body_tf"], state["bullet_count"], state["lines_used"] = None, None, 0, 0

    def _plain_table_slides(rows, src_rows, n_rows, n_cols, title, topic, own=False):
        """Plain: the table on its own slide under its title (or under a short line that
        introduces it, when there's room), white with thin black lines; a table too long for
        one slide carries on over as many as it needs, its first row repeated on each."""
        merges = _find_docx_merges(src_rows, n_cols)
        tall_merge = any(max_r > min_r for (min_r, _a, max_r, _b) in merges)
        # The words' size follows the number of columns: a long table carries on over more
        # slides instead of getting smaller words.
        cell_pt = 20 if n_cols <= 4 else 16 if n_cols <= 6 else 14 if n_cols <= 8 else 12

        def row_heights(pt, width_emu):
            col_pt = max(36.0, (width_emu / n_cols - 2 * _PPTX_INSETS) / _EMU_PT * 0.95)
            out = []
            for r_idx, row in enumerate(rows):
                n = max([_calibri_lines(row[c] if c < len(row) else "", pt, col_pt, r_idx == 0) for c in range(n_cols)] or [1])
                out.append(n * _PLAIN_LINE * pt + 7.2 + 1)
            return out

        cur = state["slide"]
        cur_body = cur.placeholders[1] if (cur is not None and len(cur.placeholders) > 1) else None
        if (cur_body is not None and not own and 0 < state["bullet_count"]
                and state["used"] <= 4 * _PLAIN_LINE * template_style["body_size"].pt):
            heights = row_heights(cell_pt, cur_body.width)
            text_h = int((state["used"] + 14) * _EMU_PT) + _PPTX_INSETS
            need = int(sum(heights) * _EMU_PT)
            if need <= (cur_body.height - text_h) * 0.95:
                L, T, W = cur_body.left, cur_body.top, cur_body.width
                cur_body.left, cur_body.top, cur_body.width, cur_body.height = L, T, W, text_h
                _plain_table_on(cur, rows, src_rows, list(range(n_rows)), merges, n_cols, cell_pt, heights, (L, T + text_h, W), False)
                state["slide"], state["body_tf"], state["bullet_count"], state["lines_used"] = None, None, 0, 0
                return

        def fresh_slide():
            s = prs.slides.add_slide(title_layout if title else blank_layout)
            if title:
                s.shapes.title.text = title
                grow = _style_pptx_slide(s, template_style)
                if topic:
                    _slide_tag(s, "Topic")
                body_ph = s.placeholders[1] if len(s.placeholders) > 1 else None
                if body_ph is not None:
                    area = (body_ph.left, body_ph.top + grow, body_ph.width, body_ph.height - grow)
                    body_ph._element.getparent().remove(body_ph._element)
                else:
                    area = (PptxInches(0.6), PptxInches(1.5) + grow, prs.slide_width - PptxInches(1.2), PptxInches(5.4) - grow)
            else:
                area = (PptxInches(0.6), PptxInches(0.6), prs.slide_width - PptxInches(1.2), prs.slide_height - PptxInches(1.2))
            return s, area

        s, area = fresh_slide()
        room = area[3] / _EMU_PT * 0.95
        heights = row_heights(cell_pt, area[2])
        if sum(heights) > room and (tall_merge or n_rows < 3):
            # can't be split (cells joined across rows): its words made smaller until it fits
            while cell_pt > 10 and sum(heights) > room:
                cell_pt -= 2
                heights = row_heights(cell_pt, area[2])
        pages, page, used = [], [0], heights[0]
        for r in range(1, n_rows):
            if len(page) > 1 and used + heights[r] > room and not tall_merge:
                pages.append(page)
                page, used = [0], heights[0]
            page.append(r)
            used += heights[r]
        pages.append(page)
        for k, page in enumerate(pages):
            if k:
                s, area = fresh_slide()
            _plain_table_on(s, rows, src_rows, page, merges, n_cols, cell_pt, heights, area[:3], k > 0)
        state["slide"], state["body_tf"], state["bullet_count"], state["lines_used"] = None, None, 0, 0

    def _plain_table_on(s, rows, src_rows, page, merges, n_cols, cell_pt, heights, box, continued):
        left, top, width = box
        height = int(sum(heights[r] for r in page) * _EMU_PT)
        shape = s.shapes.add_table(len(page), n_cols, left, top, width, height)
        if continued:
            shape.name = _PPTX_TABLE_CONT          # invisibly: the same table as on the slide before
        gtable = shape.table
        for i, r in enumerate(page):
            gtable.rows[i].height = int(heights[r] * _EMU_PT)
        at = {r: i for i, r in enumerate(page)}
        skip = {(r, c) for (min_r, min_c, max_r, max_c) in merges
                for r in range(min_r, max_r + 1) for c in range(min_c, max_c + 1) if (r, c) != (min_r, min_c)}
        for r in page:
            row = rows[r]
            for c in range(n_cols):
                if (r, c) in skip:
                    continue
                gtable.cell(at[r], c).text = row[c] if c < len(row) else ""
        _plain_pptx_table(gtable, cell_pt, template_style["body_font"])
        for min_r, min_c, max_r, max_c in merges:
            if min_r in at and max_r in at:
                try:
                    gtable.cell(at[min_r], min_c).merge(gtable.cell(at[max_r], max_c))
                except Exception:
                    pass

    def add_chart_slide(title, categories, series_list, chart_type):
        if len(categories) < 1 or len(series_list) < 1:
            return
        s = prs.slides.add_slide(title_layout)
        # R22: the chart's own title, else its section's heading — no made-up "Chart"
        own = title or state["section_title"] or ""
        derived = "" if own else (doc_topic or _topic_from_text(state["last_text"]))
        s.shapes.title.text = own or derived
        grow = _style_pptx_slide(s, template_style)
        if derived:
            _slide_tag(s, "Topic")
        left, top = PptxInches(1.0), PptxInches(1.7)
        width, height = prs.slide_width - PptxInches(2.0), prs.slide_height - PptxInches(2.4)
        if PLAIN:
            body_ph = s.placeholders[1] if len(s.placeholders) > 1 else None
            if body_ph is not None:
                left, top, width, height = body_ph.left, body_ph.top + grow, body_ph.width, body_ph.height - grow
                body_ph._element.getparent().remove(body_ph._element)

        chart_data = CategoryChartData()
        chart_data.categories = categories
        for name, values in series_list:
            numeric_values = []
            for v in values:
                try:
                    numeric_values.append(float(v))
                except (ValueError, TypeError):
                    numeric_values.append(None)  # a blank/non-numeric cell breaks the whole
                    # series if left as a string — python-pptx's chart data requires numbers
                    # or None, not the raw text values straight from the source chart's XML
            chart_data.add_series(name or "Series", numeric_values)

        try:
            graphic_frame = s.shapes.add_chart(chart_type, left, top, width, height, chart_data)
            chart = graphic_frame.chart
            chart.has_legend = len(series_list) > 1
            if chart.has_legend:
                chart.legend.position = XL_LEGEND_POSITION.BOTTOM
                chart.legend.include_in_layout = False
        except Exception:
            # A source chart type or data shape python-pptx's chart API
            # can't build (e.g. malformed values, an unusual combination
            # chart) shouldn't sink the whole conversion — fall back to
            # a plain table of the same data, which was the previous,
            # already-working behavior for every chart before this.
            _add_chart_as_table_fallback(s, categories, series_list, prs.slide_width, prs.slide_height)

        # A loose paragraph right after a chart should land on a fresh
        # slide rather than silently reusing the chart's slide.
        state["slide"], state["body_tf"], state["bullet_count"], state["lines_used"] = None, None, 0, 0

    def _picture_size(image_bytes):
        """Pixel size via python-pptx's own image reader — so anything it can
        embed is accepted, and a malformed picture is rejected before any
        slide is created for it (deleting a slide afterwards isn't safe)."""
        try:
            from pptx.parts.image import Image as _PptxImage
            return _PptxImage.from_blob(image_bytes).size
        except Exception:
            return None

    def _place_image_beside_text(image_bytes, px):
        """With only a few lines on the current slide, the picture sits to
        the right of the text instead of on a slide of its own."""
        s = state["slide"]
        if s is None or state["bullet_count"] == 0 or state["lines_used"] > 5:
            return False
        iw, ih = px
        if not iw or not ih or iw / ih > 2.3:   # very wide pictures would be tiny in half a slide
            return False
        body = s.placeholders[1] if len(s.placeholders) > 1 else None
        if body is None:
            return False
        L, T, W, H = body.left, body.top, body.width, body.height
        text_w, gap = int(W * 0.56), int(W * 0.04)
        img_l, img_w = L + text_w + gap, W - text_w - gap
        scale = min(img_w / iw, H / ih)
        pw, ph_ = int(iw * scale), int(ih * scale)
        s.shapes.add_picture(io.BytesIO(image_bytes), img_l + (img_w - pw) // 2, T + (H - ph_) // 2, pw, ph_)
        body.left, body.top, body.width, body.height = L, T, text_w, H
        # This slide is complete; whatever follows starts a new one.
        state["slide"], state["body_tf"], state["bullet_count"], state["lines_used"] = None, None, 0, 0
        return True

    def _only_title_on(s):
        """True when a slide holds nothing yet but its title."""
        for sh in s.shapes:
            if sh.is_placeholder and sh.placeholder_format.idx == 0:
                continue
            if sh.is_placeholder and sh.has_text_frame and not sh.text_frame.text.strip():
                continue
            return False
        return True

    def _add_scaled_picture(s, image_bytes, px, area):
        """The picture as large as the area allows, proportions kept, centred.
        The area is in EMU and the picture's size in pixels, so the cap on
        enlarging is in EMU per pixel: twice natural size (96 dpi = 9525
        EMU/px), enough to fill a slide without small pictures turning
        blocky. The old cap was 1 EMU per pixel, which shrank a 1200-pixel
        photo to a speck about a thousandth of an inch wide."""
        aL, aT, aW, aH = area
        iw, ih = px
        if iw and ih:
            scale = min(aW / iw, aH / ih, 2 * 9525)
            pw, ph_ = int(iw * scale), int(ih * scale)
        else:
            pw, ph_ = aW, aH
        s.shapes.add_picture(io.BytesIO(image_bytes), aL + (aW - pw) // 2, aT + (aH - ph_) // 2, pw, ph_)

    def _picture_with_caption(s, image_bytes, px, area, caption):
        """The picture as large as the area allows; its caption, if it has one, just under it."""
        if caption is None:
            _add_scaled_picture(s, image_bytes, px, area)
            return
        aL, aT, aW, aH = area
        cap_pt = 20
        lines = _calibri_lines(_full_paragraph_text(caption).strip(), cap_pt, (aW - 2 * _PPTX_INSETS) / _EMU_PT * 0.95)
        cap_h = int((lines * _PLAIN_LINE * cap_pt + 6) * _EMU_PT) + _PPTX_INSETS
        gap = int(PptxInches(0.08))
        _add_scaled_picture(s, image_bytes, px, (aL, aT, aW, max(aH - cap_h - gap, int(PptxInches(1.0)))))
        pic = s.shapes[-1]
        box = s.shapes.add_textbox(aL, pic.top + pic.height + gap, aW, cap_h)
        box.text_frame.word_wrap = True
        box.text_frame.auto_size = MSO_AUTO_SIZE.NONE
        p = box.text_frame.paragraphs[0]
        runs = _write_runs(p, caption, size=Pt(cap_pt))
        if p.alignment is None:
            p.alignment = PP_ALIGN.CENTER
        if (caption.style.name or "Normal") != "Normal":
            _pptx_mark_runs(runs, "style=" + caption.style.name)     # invisibly: its Word style

    def add_image_slide(image_bytes, hint="", caption=None):
        px = _picture_size(image_bytes)
        if px is None:
            return  # a malformed/unsupported embedded image shouldn't sink the whole conversion
        # Plain: no words beside a picture — it has a slide of its own
        if not PLAIN and _place_image_beside_text(image_bytes, px):
            return
        # The section's heading already made a slide and nothing has filled
        # it yet (a handout section that is just a picture): the picture
        # goes on that slide. It used to start another one, leaving the
        # heading's slide behind empty — a blank title-only slide before
        # every picture.
        s = state["slide"]
        if s is not None and state["bullet_count"] == 0 and _only_title_on(s):
            body_ph = s.placeholders[1] if len(s.placeholders) > 1 else None
            if body_ph is not None:
                area = (body_ph.left, body_ph.top, body_ph.width, body_ph.height)
                body_ph._element.getparent().remove(body_ph._element)
            else:
                area = (PptxInches(0.6), PptxInches(1.7), prs.slide_width - PptxInches(1.2), prs.slide_height - PptxInches(2.3))
            _picture_with_caption(s, image_bytes, px, area, caption)
            state["slide"], state["body_tf"], state["bullet_count"], state["lines_used"] = None, None, 0, 0
            return
        # Otherwise a slide of its own — now titled with its section, where
        # it used to be an untitled full-bleed picture cut off from the topic
        # it illustrated.
        # R22: under its section's heading; with none yet, the document's title or a topic in
        # the words around the picture
        pic_title = state["section_title"] or ""
        derived = "" if pic_title else (doc_topic or _topic_from_text(hint or state["last_text"]))
        titled = bool(pic_title or derived)
        s = prs.slides.add_slide(title_layout if titled else blank_layout)
        if titled:
            s.shapes.title.text = pic_title or derived
            grow = _style_pptx_slide(s, template_style)
            if derived:
                _slide_tag(s, "Topic")
            body_ph = s.placeholders[1] if len(s.placeholders) > 1 else None
            if body_ph is not None:
                area = (body_ph.left, body_ph.top + grow, body_ph.width, body_ph.height - grow)
                body_ph._element.getparent().remove(body_ph._element)
            else:
                area = (PptxInches(0.6), PptxInches(1.7), prs.slide_width - PptxInches(1.2), prs.slide_height - PptxInches(2.3))
        else:
            area = (0, 0, prs.slide_width, prs.slide_height)
        _picture_with_caption(s, image_bytes, px, area, caption)
        # A loose paragraph appearing right after an image should land on a
        # fresh slide rather than silently reusing the image slide.
        state["slide"], state["body_tf"], state["bullet_count"], state["lines_used"] = None, None, 0, 0

    def add_shape_slide(mso_shape, fill_hex, shape_text, width_emu, height_emu):
        # R22: under its section's heading (or, with none, a topic in its own words)
        shape_title = state["section_title"] or ""
        derived = "" if shape_title else (doc_topic or _topic_from_text(shape_text or state["last_text"]))
        if shape_title or derived:
            s = prs.slides.add_slide(title_layout)
            s.shapes.title.text = shape_title or derived
            grow = _style_pptx_slide(s, template_style)
            if derived:
                _slide_tag(s, "Topic")
            body_ph = s.placeholders[1] if len(s.placeholders) > 1 else None
            if body_ph is not None:
                aL, aT, aW, aH = body_ph.left, body_ph.top + grow, body_ph.width, body_ph.height - grow
                body_ph._element.getparent().remove(body_ph._element)
            else:
                aL, aT, aW, aH = PptxInches(0.6), PptxInches(1.7), prs.slide_width - PptxInches(1.2), prs.slide_height - PptxInches(2.3)
        else:
            s = prs.slides.add_slide(blank_layout)
            aL, aT, aW, aH = PptxInches(1), PptxInches(1), prs.slide_width - PptxInches(2), prs.slide_height - PptxInches(2)
        # Preserve the shape's own aspect ratio when it had real dimensions,
        # scaled up to a reasonable on-slide size, rather than stretching
        # every recreated shape to one fixed box regardless of its source
        # proportions (a wide arrow shouldn't come out square).
        max_w, max_h = aW, aH
        if width_emu and height_emu:
            scale = min(max_w / width_emu, max_h / height_emu, 4.0)
            w, h = int(width_emu * scale), int(height_emu * scale)
        else:
            w, h = PptxInches(4), PptxInches(2)
        left, top = int(aL + (aW - w) / 2), int(aT + (aH - h) / 2)
        try:
            shape = s.shapes.add_shape(mso_shape, left, top, w, h)
        except Exception:
            return  # an unusual/malformed shape shouldn't sink the whole conversion
        if fill_hex:
            try:
                shape.fill.solid()
                shape.fill.fore_color.rgb = PptxRGBColor.from_string(fill_hex)
            except Exception:
                pass
        if shape_text:
            shape.text_frame.text = shape_text
            for p in shape.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                for run in p.runs:
                    run.font.size = Pt(20)
                    run.font.color.rgb = PptxRGBColor(0xFF, 0xFF, 0xFF) if fill_hex else PptxRGBColor(0x00, 0x00, 0x00)
        # Same reasoning as after an image or table — a loose paragraph
        # right after a shape belongs on its own fresh slide.
        state["slide"], state["body_tf"], state["bullet_count"], state["lines_used"] = None, None, 0, 0

    MAX_IMAGES = 30
    image_count = 0
    MAX_SHAPES = 30
    shape_count = 0

    # A list, not the raw generator, specifically so a heading can look
    # one item ahead to see whether a table follows it immediately — see
    # the comment in add_table_slide for why that lookahead exists.
    blocks = list(_iter_block_items(doc))
    pending_table_title = None
    pending_table_tag = ""
    skip_blocks = set()

    def _make_title_slide(text, block_idx):
        """The document's Title becomes a real title slide: centred, with the
        short line under it (course code, author, date) as its subtitle. It
        used to be laid out like any content slide, with that line shown as
        a bullet point."""
        layout = None
        for lo in prs.slide_layouts:
            kinds = [p.placeholder_format.type for p in lo.placeholders]
            if PP_PLACEHOLDER.SUBTITLE in kinds and any(p.placeholder_format.idx == 0 for p in lo.placeholders):
                layout = lo
                break
        if layout is None:
            return False
        s = prs.slides.add_slide(layout)
        s.shapes.title.text = text
        _style_pptx_slide(s, template_style)
        _slide_tag(s, "Title")
        sub_ph = next((p for p in s.placeholders if p.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE), None)
        nxt = blocks[block_idx + 1] if block_idx + 1 < len(blocks) else None
        used = False
        if sub_ph is not None and nxt is not None and not isinstance(nxt, DocxTable):
            nstyle = (nxt.style.name or "").lower()
            ntext = _full_paragraph_text(nxt).strip()
            if (ntext and not nstyle.startswith("heading") and nstyle != "title" and "list" not in nstyle
                    and len(ntext.split()) <= 25 and not any(True for _ in _iter_inline_images(nxt, doc))):
                sub_ph.text_frame.text = ""
                sub_runs = _pptx_add_text(sub_ph.text_frame.paragraphs[0], ntext)
                for p in sub_ph.text_frame.paragraphs:
                    _style_pptx_body_paragraph(p, template_style, size=Pt(24))
                if (nxt.style.name or "") != "Subtitle":
                    _pptx_mark_runs(sub_runs, "style=" + (nxt.style.name or "Normal"))
                skip_blocks.add(block_idx + 1)
                used = True
        if sub_ph is not None and not used:
            sub_ph._element.getparent().remove(sub_ph._element)
        state["slide"], state["body_tf"], state["bullet_count"], state["lines_used"] = None, None, 0, 0
        return True

    _SLIDE_LABEL_RE = re.compile(r"^slide\s*\d{1,3}$", re.I)

    def _nearby_text(block_idx):
        """The words around a picture: its own paragraph's, else the next one's (a caption),
        else the last words before it."""
        for j in (block_idx, block_idx + 1, block_idx + 2):
            if j < len(blocks) and not isinstance(blocks[j], DocxTable):
                t = _full_paragraph_text(blocks[j]).strip()
                if t and not _SLIDE_LABEL_RE.match(t):
                    return t
        return state["last_text"]

    _CAPTION_RE = re.compile(r"^(?:figure|fig\.?|image|picture|photo(?:graph)?|diagram|chart|graph|plate|illustration|map)\s*\d+", re.I)

    def _caption_for(block_idx):
        """A picture's caption — its own paragraph's words, or the paragraph just after it, when
        that is a caption ("Figure 1. …", or Word's Caption style). Returns (paragraph, whether
        it is the picture's own paragraph); the next paragraph, used, isn't placed again."""
        def is_caption(par, text):
            return bool(text) and len(text.split()) <= 40 and (
                (par.style.name or "").lower() == "caption" or bool(_CAPTION_RE.match(text)))
        own = blocks[block_idx]
        own_text = _full_paragraph_text(own).strip()
        if own_text:
            return (own, True) if is_caption(own, own_text) else (None, False)
        nxt = blocks[block_idx + 1] if block_idx + 1 < len(blocks) else None
        if nxt is not None and not isinstance(nxt, DocxTable):
            text = _full_paragraph_text(nxt).strip()
            if is_caption(nxt, text) and not any(True for _ in _iter_inline_images(nxt, doc)):
                skip_blocks.add(block_idx + 1)
                return nxt, False
        return None, False

    def _is_handout_artefact(text):
        """"Slide 4" on a line of its own: the small labels the app's own
        handouts print under each heading. Converting a handout back into
        slides used to put them on the new slides as the only content."""
        return bool(_SLIDE_LABEL_RE.match((text or "").strip()))

    def _is_lead_in(text):
        t = text.strip()
        return t.endswith(":") and 2 <= len(t.split()) <= 14

    def _next_is_list_item(block_idx):
        nxt = blocks[block_idx + 1] if block_idx + 1 < len(blocks) else None
        if nxt is None or isinstance(nxt, DocxTable):
            return False
        return (_get_paragraph_direct_list_info(nxt, numbering_formats) is not None
                or "list" in (nxt.style.name or "").lower())


    # ---- A handout made by Slides → Handout: its hidden bookmarks mark where
    # each slide began, so the same slides are rebuilt — one per mark, with
    # their own titles, words, tables, pictures and speaker notes.
    def _para_kind(para):
        sname = (para.style.name or "").lower()
        info = _get_paragraph_direct_list_info(para, numbering_formats)
        if info is not None:
            level, numbered = info
            m = re.search(r"list (?:bullet|number|continue)\s*(\d+)", sname)
            if m:
                level = max(level, int(m.group(1)) - 1)   # "List Bullet 2" sits one level in
            return ("number" if numbered else "bullet"), max(0, min(level, 4))
        if "list" in sname:
            m = re.search(r"(\d+)", sname)
            return ("number" if "number" in sname else "bullet"), (max(0, min(int(m.group(1)) - 1, 4)) if m else 0)
        return "plain", 0

    def _write_runs(p, para, size=None, kind="plain", level=0):
        runs_data = []
        for src_run, text, is_link, address, is_deleted in _iter_all_runs(para):
            if not text:
                continue
            b, i, u, color, sz = _effective_run_format(src_run)
            runs_data.append((text, bool(b), bool(i), True if is_link else bool(u), is_link, address, (None if is_link else color), sz, is_deleted))
        p.level = min(level, 4)
        align = _docx_align_to_pptx(_effective_paragraph_alignment(para))
        if align is not None:
            p.alignment = align
        pPr = p._p.get_or_add_pPr()
        if kind == "number":
            pPr.append(pPr.makeelement(qn("a:buAutoNum"), {"type": "arabicPeriod"}))
        elif kind == "plain":
            pPr.set("marL", "0")
            pPr.set("indent", "0")
            pPr.append(pPr.makeelement(qn("a:buNone"), {}))
        captured = []
        for text, bold, italic, underline, is_link, address, color, sz, is_deleted in runs_data:
            for k, seg in enumerate(text.split("\n")):
                if k:
                    p.add_line_break()
                if not seg:
                    continue
                r = p.add_run()
                r.text = seg
                if is_link and address:
                    try:
                        r.hyperlink.address = address
                    except Exception:
                        pass
                if is_deleted:
                    r._r.get_or_add_rPr().set("strike", "sngStrike")
                captured.append((r, bold, italic, underline, is_link, color, sz))
        _style_pptx_body_paragraph(p, template_style, size=size)
        for r, bold, italic, underline, is_link, color, sz in captured:
            r.font.bold, r.font.italic, r.font.underline = bold, italic, underline
            if sz is not None and size is None:
                r.font.size = Pt(sz)
            if is_link:
                r.font.color.rgb = PptxRGBColor(0x05, 0x63, 0xC1)
            elif color:
                try:
                    r.font.color.rgb = PptxRGBColor.from_string(color)
                except Exception:
                    pass
        return [c[0] for c in captured]

    def _fill_table(s, table, left, top, width, height):
        src_rows = [list(row.cells) for row in table.rows]
        if not src_rows:
            return
        n_rows, n_cols = len(src_rows), max(len(r) for r in src_rows)
        gtable = s.shapes.add_table(n_rows, n_cols, left, top, width, max(height, int(PptxInches(0.4) * n_rows))).table
        merges = _find_docx_merges(src_rows, n_cols)
        skip = {(r, c) for (a, b, c2, d) in merges for r in range(a, c2 + 1) for c in range(b, d + 1) if (r, c) != (a, b)}
        cell_pt = Pt(18 if (n_rows <= 6 and n_cols <= 4) else 14 if (n_rows <= 10 and n_cols <= 6) else 12)
        for r_idx, row in enumerate(src_rows):
            for c_idx in range(n_cols):
                if (r_idx, c_idx) in skip or c_idx >= len(row):
                    continue
                cell = gtable.cell(r_idx, c_idx)
                tf = cell.text_frame
                tf.text = ""
                paras = [cp for cp in row[c_idx].paragraphs]
                for k, cp in enumerate(paras):
                    p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                    for src_run, text, is_link, address, is_deleted in _iter_all_runs(cp):
                        b, i, u, color, sz = _effective_run_format(src_run)
                        for rr in _pptx_add_text(p, text):
                            rr.font.size = cell_pt
                            rr.font.bold, rr.font.italic = bool(b), bool(i)
                            if template_style.get("body_font"):
                                rr.font.name = template_style["body_font"]
                shade = _get_docx_cell_shading(row[c_idx])
                if shade:
                    try:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = PptxRGBColor.from_string(shade)
                    except Exception:
                        pass
        for a, b, c2, d in merges:
            try:
                gtable.cell(a, b).merge(gtable.cell(c2, d))
            except Exception:
                pass

    def _marked_build():
        groups, cur = [], None
        for b in blocks:
            mk = None if isinstance(b, DocxTable) else _docx_slide_mark(b)
            if mk is not None or cur is None:
                cur = {"cont": bool(mk and mk[1]), "items": []}
                groups.append(cur)
            cur["items"].append(b)
        prev_title = ""
        title_slide_layout = next((lo for lo in prs.slide_layouts
                                   if PP_PLACEHOLDER.SUBTITLE in [p.placeholder_format.type for p in lo.placeholders]), title_layout)
        # A layout with a title and two text columns side by side (PowerPoint's "Two Content").
        two_layout = next((lo for lo in prs.slide_layouts
                           if [p.placeholder_format.idx for p in lo.placeholders if p.placeholder_format.type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT)][:2] == [1, 2]
                           and len([p for p in lo.placeholders if p.placeholder_format.type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT)]) == 2), None)
        for g in groups:
            title_par, kind, sub_pars, body, tables, images, notes = None, None, [], [], [], [], []
            col2_at = None
            for b in g["items"]:
                if isinstance(b, DocxTable):
                    tables.append(b)
                    continue
                if col2_at is None and any((bm.get(qn("w:name")) or "").startswith("_dct_col") for bm in b._p.iter(qn("w:bookmarkStart"))):
                    col2_at = len(body)
                images += list(_iter_inline_images(b, doc))
                text = _full_paragraph_text(b)
                if not text.strip():
                    continue
                sname = (b.style.name or "").lower()
                if sname == "speaker notes":
                    notes.append(text.strip())
                    continue
                hm = re.match(r"heading (\d+)", sname)
                if title_par is None and not body and not sub_pars and (sname == "title" or hm):
                    title_par, kind = b, ("Title" if sname == "title" else f"Heading {hm.group(1)}")
                    continue
                if kind == "Title" and not body and sname == "subtitle":
                    sub_pars.append(b)
                    continue
                body.append(b)
            title_text = _full_paragraph_text(title_par).strip() if title_par is not None else (prev_title if g["cont"] else "")
            two_col = col2_at is not None and 0 < col2_at < len(body) and kind != "Title" and two_layout is not None
            s = prs.slides.add_slide(title_slide_layout if kind == "Title" else (two_layout if two_col else title_layout))
            if s.shapes.title is not None and title_text:
                s.shapes.title.text_frame.text = ""
                _pptx_add_text(s.shapes.title.text_frame.paragraphs[0], title_text)
            _style_pptx_slide(s, template_style)
            if s.shapes.title is not None and not title_text:
                s.shapes.title._element.getparent().remove(s.shapes.title._element)   # an untitled slide stays untitled
            _slide_tag(s, kind if (kind and not g["cont"]) else "")
            if two_col:
                cols = [body[:col2_at], body[col2_at:]]
                # Both columns at one size: shrunk only if the fuller column wouldn't fit.
                est = max(sum(max(1, -(-len(_full_paragraph_text(b)) // 34)) for b in col) for col in cols)
                size = Pt(max(12, int(20 * MAX_LINES_PER_SLIDE / est))) if est > MAX_LINES_PER_SLIDE else None
                for idx, col in zip((1, 2), cols):
                    ph = next((p for p in s.placeholders if p.placeholder_format.idx == idx), None)
                    if ph is None:
                        continue
                    tf = ph.text_frame
                    tf.text = ""
                    for k, b in enumerate(col):
                        pk, lvl = _para_kind(b)
                        _write_runs(tf.paragraphs[0] if k == 0 else tf.add_paragraph(), b, size=size, kind=pk, level=lvl)
                body = []
            if kind == "Title" and images:
                # Pictures on a title slide sit beside the title, not over it.
                sw, sh = prs.slide_width, prs.slide_height
                for ph in s.placeholders:
                    if ph.placeholder_format.type in (PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.SUBTITLE):
                        top, height = ph.top, ph.height          # inherited from the layout — kept
                        ph.left, ph.top, ph.width, ph.height = PptxInches(1.0), top, int(sw * 0.47), height
                col_l, col_w = int(sw * 0.56), int(sw * 0.39)
                each = int((sh - PptxInches(1.2)) / len(images))
                for k, img in enumerate(images):
                    px = _picture_size(img)
                    if px:
                        _add_scaled_picture(s, img, px, (col_l, PptxInches(0.6) + k * each, col_w, each))
                images = []
            if kind == "Title":
                sub_ph = next((p for p in s.placeholders if p.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE), None)
                lines = sub_pars + body
                if sub_ph is not None:
                    body = []
                    if lines:
                        tf = sub_ph.text_frame
                        tf.text = ""
                        for k, b in enumerate(lines):
                            p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                            runs = _write_runs(p, b, size=Pt(24 if len(lines) <= 2 else 16))
                            if (b.style.name or "") != "Subtitle":
                                _pptx_mark_runs(runs, "style=" + (b.style.name or "Normal"))
                    else:
                        sub_ph._element.getparent().remove(sub_ph._element)
            body_ph = next((ph for ph in s.placeholders if ph.placeholder_format.idx == 1 and ph.placeholder_format.type != PP_PLACEHOLDER.SUBTITLE), None)
            if body_ph is not None:
                L, T, W, H = body_ph.left, body_ph.top, body_ph.width, body_ph.height
            else:
                L, T = PptxInches(0.6), PptxInches(1.7) if title_text else PptxInches(0.6)
                W, H = prs.slide_width - PptxInches(1.2), prs.slide_height - T - PptxInches(0.6)
            text_w = int(W * 0.58) if images and (body or tables) else W
            text_h = 0
            if body:
                if body_ph is None:
                    body_ph = s.shapes.add_textbox(L, T, text_w, H)
                tf = body_ph.text_frame
                tf.word_wrap = True
                tf.text = ""
                est = sum(_estimate_line_count(_full_paragraph_text(b)) for b in body) * (W / max(text_w, 1))
                size = Pt(max(12, int(20 * MAX_LINES_PER_SLIDE / est))) if est > MAX_LINES_PER_SLIDE else None
                for k, b in enumerate(body):
                    p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                    sname = (b.style.name or "").lower()
                    hm = re.match(r"heading (\d+)", sname)
                    if hm:
                        runs = _write_runs(p, b, size=size)
                        for r in runs:
                            r.font.bold = True
                        _pptx_mark_runs(runs, f"h{hm.group(1)}")
                    else:
                        pk, lvl = _para_kind(b)
                        _write_runs(p, b, size=size, kind=pk, level=lvl)
                text_h = H if not tables else min(int(H * 0.55), int(PptxInches(0.42) * est + PptxInches(0.3)))
                body_ph.left, body_ph.top, body_ph.width, body_ph.height = L, T, text_w, text_h
            elif body_ph is not None and not two_col:
                body_ph._element.getparent().remove(body_ph._element)
            if tables:
                top = T + text_h
                each = max(int((T + H - top) / len(tables)), int(PptxInches(0.8)))
                for k, tb in enumerate(tables):
                    _fill_table(s, tb, L, top + k * each, text_w, each)
            if images:
                gap = int(W * 0.04)
                iL, iW = (L + text_w + gap, W - text_w - gap) if (body or tables) else (L, W)
                each = int(H / len(images))
                for k, img in enumerate(images):
                    px = _picture_size(img)
                    if px:
                        _add_scaled_picture(s, img, px, (iL, T + k * each, iW, each))
            if notes:
                s.notes_slide.notes_text_frame.text = "\n".join(notes)
            prev_title = title_text

    marked = any(_docx_slide_mark(b) for b in blocks if not isinstance(b, DocxTable))
    if marked:
        _marked_build()

    for block_idx, block in enumerate([] if marked else blocks):
        if block_idx in skip_blocks:
            continue
        if isinstance(block, DocxTable):
            n_before = len(prs.slides)
            add_table_slide(block, carried_title=pending_table_title or state["section_title"], own=bool(pending_table_title))
            if pending_table_title and len(prs.slides) > n_before:
                _slide_tag(prs.slides[n_before], pending_table_tag)     # its first slide (a long table has more)
            pending_table_title = None
            continue
        pending_table_title = None
        para = block
        own_caption = False
        if image_count < MAX_IMAGES:
            imgs = list(_iter_inline_images(para, doc))
            caption = None
            if PLAIN and imgs:
                caption, own_caption = _caption_for(block_idx)
            for k, img_bytes in enumerate(imgs):
                if image_count >= MAX_IMAGES:
                    break
                add_image_slide(img_bytes, _nearby_text(block_idx), caption if k == len(imgs) - 1 else None)
                image_count += 1
        if shape_count < MAX_SHAPES:
            for mso_shape, fill_hex, shape_text, w_emu, h_emu in _iter_docx_shapes(para):
                if shape_count >= MAX_SHAPES:
                    break
                add_shape_slide(mso_shape, fill_hex, shape_text, w_emu, h_emu)
                shape_count += 1
        for tb_para in _iter_textbox_paragraphs(para):
            tb_text = _full_paragraph_text(tb_para).strip()
            if tb_text:
                add_bullet(tb_para, level=0)
        for chart_title, chart_cats, chart_series, chart_type in _iter_docx_charts(para, doc):
            add_chart_slide(chart_title, chart_cats, chart_series, chart_type)
        if own_caption:
            continue            # its words are the picture's caption, already under it
        text = _full_paragraph_text(para).strip()
        if not text or _is_handout_artefact(text):
            continue
        para_style_name = (para.style.name or "").lower()
        heading_match = re.match(r"heading (\d+)", para_style_name)
        # Heading 1 and Heading 2 both start a new slide — matching how
        # most outline-based DOCX->PPTX tools split a document, and how a
        # reader would expect a document's own structure to map onto
        # slide boundaries. Heading 3+ stays as a bolded subheading within
        # the current slide rather than fragmenting into ever-thinner
        # slides for what's usually meant to be one cohesive topic.
        if para_style_name == "title" or (heading_match and int(heading_match.group(1)) <= 2):
            state["section_title"] = text
            state["num"].clear()
            if para_style_name == "title" and len(prs.slides) == 0 and _make_title_slide(text, block_idx):
                continue
            next_block = blocks[block_idx + 1] if block_idx + 1 < len(blocks) else None
            heading_tag = "Title" if para_style_name == "title" else f"Heading {heading_match.group(1)}"
            if isinstance(next_block, DocxTable):
                pending_table_tag = heading_tag
                # This heading has nothing else to show but a table right
                # after it — deferring the slide entirely (rather than
                # creating one now and a second, unlabeled one for the
                # table) means the table gets exactly one properly
                # titled slide instead of two confusing ones.
                pending_table_title = text
            else:
                new_slide(text)
                _slide_tag(state["slide"], heading_tag)
        elif heading_match:
            add_subheading(text, int(heading_match.group(1)))
        else:
            direct_list_info = _get_paragraph_direct_list_info(para, numbering_formats)
            if direct_list_info is not None:
                level, numbered = direct_list_info
                m_lvl = re.search(r"list (?:bullet|number|continue)\s*(\d+)", para_style_name)
                if m_lvl:
                    level = max(level, int(m_lvl.group(1)) - 1)   # "List Bullet 2" sits one level in
                level = max(0, min(level, 4))
            else:
                level = 0
                numbered = False
                if "list" in para_style_name:
                    m = re.search(r"(\d+)", para_style_name)
                    if m:
                        level = max(0, min(int(m.group(1)) - 1, 4))
                    numbered = "number" in para_style_name
            plain = direct_list_info is None and "list" not in para_style_name
            add_bullet(para, level=level, numbered=numbered, plain=plain)

    # Only needed as a last-resort guarantee that the presentation has at
    # least one slide (e.g. a source document that was nothing but a
    # single image) — if the document already produced real slides and
    # simply ended on a table/chart/image, state["slide"] being None just
    # means nothing followed it, not that a slide is missing. Creating one
    # anyway produced a pointless empty trailing slide a user would have
    # to notice and delete themselves.
    if state["slide"] is None and len(prs.slides) == 0:
        new_slide("")

    footnotes = _get_docx_footnotes(doc)
    if footnotes and PLAIN:
        state["slide"], state["section_title"] = None, None
        new_slide("Footnotes", topic=True)       # R22: says what the slide is (left out of a handout)
        for i, note in enumerate(footnotes, 1):
            _plain_place([(f"{i}. {note}", False, False, False, False, None, None, None, False)], 0, "plain", None)
        footnotes = []
    if footnotes:
        new_slide("Footnotes", topic=True)
        for i, note in enumerate(footnotes, 1):
            add_bullet_text = f"{i}. {note}"
            body_tf = state["body_tf"]
            p = body_tf.paragraphs[0] if state["bullet_count"] == 0 else body_tf.add_paragraph()
            p.text = add_bullet_text
            _style_pptx_body_paragraph(p, template_style)
            state["bullet_count"] += 1

    endnotes = _get_docx_endnotes(doc)
    if endnotes and PLAIN:
        state["slide"], state["section_title"] = None, None
        new_slide("Endnotes", topic=True)
        for i, note in enumerate(endnotes, 1):
            _plain_place([(f"{i}. {note}", False, False, False, False, None, None, None, False)], 0, "plain", None)
        endnotes = []
    if endnotes:
        new_slide("Endnotes", topic=True)
        for i, note in enumerate(endnotes, 1):
            body_tf = state["body_tf"]
            p = body_tf.paragraphs[0] if state["bullet_count"] == 0 else body_tf.add_paragraph()
            p.text = f"{i}. {note}"
            _style_pptx_body_paragraph(p, template_style)
            state["bullet_count"] += 1

    # Review comments are notes about the document, not part of it — they
    # aren't put on a slide. The page header/footer lines become the slides'
    # footer (where they belong on slides), not a slide of their own.
    header_footer = _get_docx_header_footer_text(doc)
    if header_footer:
        footer_text = "\n".join(header_footer)
        for s in prs.slides:
            _pptx_set_footer(s, footer_text)

    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "converted.pptx")
    prs.save(out_path)
    return out_path


# --------------------------------------------------------------- PPTX -> DOCX
def _extract_pdf_page_links(page):
    """Returns a list of (exact_text, url) for each real hyperlink on a PDF
    page. pypdf's link annotations give a URL and an on-page rectangle but
    no text of their own — this uses extract_text()'s visitor_text
    callback to get the actual position of every rendered text fragment,
    then matches each link's rectangle against fragments that fall inside
    it, so the resulting text is the real clickable words rather than a
    guess made after the text has been reflowed into paragraphs."""
    fragments = []

    def visitor(text, cm, tm, font_dict, font_size):
        if text.strip():
            fragments.append((text, tm[4], tm[5]))

    try:
        page.extract_text(visitor_text=visitor)
    except Exception:
        return []

    try:
        annotations = page.get("/Annots") or []
    except Exception:
        return []

    links = []
    for a in annotations:
        try:
            obj = a.get_object()
            if obj.get("/Subtype") != "/Link":
                continue
            action = obj.get("/A")
            url = action.get("/URI") if action else None
            rect = obj.get("/Rect")
            if not url or not rect:
                continue
            x0, y0, x1, y1 = (float(v) for v in rect)
            matched = [t for t, x, y in fragments if x0 <= x <= x1 and y0 - 1 <= y <= y1 + 1]
            text = "".join(matched).strip()
            if text:
                links.append((text, url))
        except Exception:
            continue
    return links


def _add_docx_paragraph_with_links(doc, text, page_links, style=None):
    """Adds a paragraph, converting any substring that exactly matches a
    known link's real text into an actual hyperlink instead of plain text.
    Returns (paragraph, used_links) — used_links is the subset of
    page_links actually placed inline, so the caller can tell exactly
    which links still need to fall back to a plain URL list.

    A link's text that appears more than once in the paragraph is
    deliberately skipped rather than guessed at: with only a single exact
    string match to go on, there's no reliable way to tell which
    occurrence is the real link and which is a coincidentally identical
    word — e.g. a paragraph that says '...here appears first as plain
    text, but the second here is a real link' would otherwise confidently
    hyperlink the *wrong* word. Falling back to listing the link's URL
    separately is honest; guessing the first occurrence is not."""
    matches = []
    for link_text, url in page_links:
        if text.count(link_text) != 1:
            continue  # ambiguous or absent — leave for the caller's fallback list
        idx = text.find(link_text)
        matches.append((idx, idx + len(link_text), url, link_text))
    if not matches:
        p = doc.add_paragraph(text, style=style) if style else doc.add_paragraph(text)
        return p, []
    matches.sort(key=lambda m: m[0])
    accepted = []
    last_end = -1
    for start, end, url, link_text in matches:
        if start >= last_end:  # drop any overlap defensively, keep the leftmost match
            accepted.append((start, end, url, link_text))
            last_end = end
    # used_links reflects what was actually accepted after overlap
    # resolution — a match dropped here for overlapping an earlier one
    # never actually appears in the paragraph, so it must stay eligible
    # for the caller's fallback list rather than being marked as handled.
    used_links = [(link_text, url) for start, end, url, link_text in accepted]
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    pos = 0
    for start, end, url, link_text in accepted:
        if start > pos:
            p.add_run(text[pos:start])
        _add_docx_hyperlink(p, url, text[start:end])
        pos = end
    if pos < len(text):
        p.add_run(text[pos:])
    return p, used_links


def _add_docx_toc_field(doc, heading_texts):
    """Inserts a real Word TOC field (not a manually-typed list) right
    where called, built from headings the document already has — a
    genuinely clickable, navigable table of contents rather than plain
    text. The field's cached display text is pre-populated with the
    actual heading names (one per line) rather than a generic "please
    update this field" placeholder, confirmed directly: without this,
    the document shows that placeholder message until a field refresh
    happens, which many readers never trigger. It becomes a fully
    page-numbered, dot-leadered TOC automatically the first time it's
    opened in real Word, since the document is also marked to refresh
    fields on open — but reads correctly even for a reader who never
    triggers that refresh."""
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = 'TOC \\o "1-3" \\h \\z \\u'
    fld_separate = OxmlElement("w:fldChar")
    fld_separate.set(qn("w:fldCharType"), "separate")
    run._r.append(fld_begin)
    run._r.append(instr_text)
    run._r.append(fld_separate)

    cached_run = OxmlElement("w:r")
    for i, text in enumerate(heading_texts):
        if i > 0:
            cached_run.append(OxmlElement("w:br"))
        t_el = OxmlElement("w:t")
        t_el.set(qn("xml:space"), "preserve")
        t_el.text = text
        cached_run.append(t_el)
    paragraph._p.append(cached_run)

    end_run = paragraph.add_run()
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    end_run._r.append(fld_end)

    update_fields = OxmlElement("w:updateFields")
    update_fields.set(qn("w:val"), "true")
    # <w:updateFields> has a specific required position in CT_Settings'
    # schema sequence — confirmed directly against the OOXML schema
    # after an initial attempt (just appending it) failed validation:
    # it must come immediately before <w:compat>, not at the end of the
    # element, or Word/validators reject the whole settings part as
    # out of sequence.
    compat_el = doc.settings.element.find(qn("w:compat"))
    if compat_el is not None:
        compat_el.addprevious(update_fields)
    else:
        doc.settings.element.append(update_fields)
    return paragraph


# The CT_PPrBase sequence (the subset relevant to elements this file ever
# inserts into a paragraph's pPr) — used to insert a new pPr child at the
# position the OOXML schema actually requires relative to whatever's
# already there, rather than assuming pPr is empty and simply appending.
# Confirmed necessary directly: paragraph_format.space_before (called
# before the callout styling below) already creates a <w:spacing>
# element, and appending pBdr/shd after it put them out of the required
# order, which real validation caught.
_PPR_CHILD_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
    "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
    "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
    "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
    "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
]


def _insert_pPr_child_in_order(pPr, new_el):
    """Inserts new_el into a <w:pPr> at the position CT_PPrBase's schema
    sequence requires relative to whatever children are already
    present, rather than assuming pPr is empty and just appending —
    appending blindly is only safe when nothing else has touched this
    paragraph's formatting yet, which isn't a safe assumption in
    general (space_before, alignment, indentation can all have been set
    first)."""
    new_local = new_el.tag.split("}")[-1]
    new_idx = _PPR_CHILD_ORDER.index(new_local) if new_local in _PPR_CHILD_ORDER else len(_PPR_CHILD_ORDER)
    for child in pPr:
        child_local = child.tag.split("}")[-1]
        child_idx = _PPR_CHILD_ORDER.index(child_local) if child_local in _PPR_CHILD_ORDER else len(_PPR_CHILD_ORDER)
        if child_idx > new_idx:
            child.addprevious(new_el)
            return
    pPr.append(new_el)


def _add_docx_callout_style(paragraph, fill_hex="F2F2F2", border_hex="808080"):
    """Gives a paragraph a shaded background and a colored left border —
    the "callout box" treatment modern documentation tools (Notion,
    GitHub's markdown admonitions) use to visually separate an aside
    from the main flow of text, applied here to speaker notes so they
    read as a distinct annotation rather than blending into the body
    content as more small italic text would."""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "24")
    left.set(qn("w:space"), "4")
    left.set(qn("w:color"), border_hex)
    pBdr.append(left)
    _insert_pPr_child_in_order(pPr, pBdr)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    _insert_pPr_child_in_order(pPr, shd)
    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), "200")
    _insert_pPr_child_in_order(pPr, ind)


def _add_docx_hyperlink(paragraph, url, text, bold=False, italic=False, underline=True):
    """python-docx has no high-level API for creating a hyperlink — this is
    the standard recipe: register the external relationship, then build the
    <w:hyperlink> element by hand with a run inside it styled to look like a
    normal link (blue, underlined) unless told otherwise."""
    part = paragraph.part
    r_id = part.relate_to(
        url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    run_el = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    if bold:
        rpr.append(OxmlElement("w:b"))
    if italic:
        rpr.append(OxmlElement("w:i"))
    if underline:
        u = OxmlElement("w:u")
        u.set(qn("w:val"), "single")
        rpr.append(u)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    rpr.append(color)
    run_el.append(rpr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    run_el.append(t)
    hyperlink.append(run_el)
    paragraph._p.append(hyperlink)


_SMARTART_URI = "http://schemas.microsoft.com/office/drawing/2010/diagram"
_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def _iter_smartart_fallback_text(slide):
    """Yields (found_smartart, text) — found_smartart is True for every
    qualifying SmartArt AlternateContent block encountered, independent of
    whether text came out of it, so a diagram whose fallback shapes happen
    to carry no text still isn't silently indistinguishable from a slide
    with no diagram at all. python-pptx's own shape iterator (slide.shapes,
    and _iter_flat_shapes built on it) silently skips anything wrapped in
    <mc:AlternateContent> — confirmed directly rather than assumed — so a
    SmartArt diagram (which real PowerPoint files wrap this way for
    backward compatibility) is completely invisible to normal shape
    iteration, not just unrecognized. The fallback content inside
    <mc:Fallback> is built from ordinary shapes approximating the
    diagram's own text, and reaching it means walking the slide's raw XML
    directly rather than relying on python-pptx's shape abstraction at
    all for this specific case."""
    sp_tree = slide.shapes._spTree
    for alt in sp_tree.findall(f".//{{{_MC_NS}}}AlternateContent"):
        choice = alt.find(f"{{{_MC_NS}}}Choice")
        if choice is None:
            continue
        graphic_data = choice.find(".//" + pptx_qn("a:graphicData"))
        if graphic_data is None or graphic_data.get("uri") != _SMARTART_URI:
            continue  # an AlternateContent block for something unrelated to SmartArt
        yield (True, None)
        fallback = alt.find(f"{{{_MC_NS}}}Fallback")
        if fallback is None:
            continue
        for sp in fallback.findall(".//" + pptx_qn("p:sp")):
            for para in sp.findall(".//" + pptx_qn("a:p")):
                text = _sanitize_xml_text("".join(t.text or "" for t in para.findall(".//" + pptx_qn("a:t"))).strip())
                if text:
                    yield (False, text)


def _iter_flat_shapes(shapes):
    """Yields every shape in a slide's shape collection, recursing into any
    GROUP shape's own .shapes. A grouped shape (very common — logos paired
    with captions, multi-element diagrams) is otherwise a single opaque
    'GROUP' entry with no text frame, table, or chart of its own, so a plain
    top-level loop silently sees nothing inside it at all."""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_flat_shapes(shape.shapes)
        else:
            yield shape


def _docx_divider(doc):
    """A thin grey line between items — structure without any words."""
    p = doc.add_paragraph()
    p_pr = p._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    for k, v in (("w:val", "single"), ("w:sz", "6"), ("w:space", "1"), ("w:color", "BFBFBF")):
        bottom.set(qn(k), v)
    borders.append(bottom)
    p_pr.append(borders)
    return p


def _handout_title(prs, slides):
    """The handout's title comes from the deck itself: its title slide, with
    that slide's subtitle (course, date, presenter) underneath — else the
    file's own title property. It used to always be the words "Slide
    Handout". Returns (title, subtitle, whether slide 1 is fully used up by
    this, so it can be left out of the body instead of repeated there)."""
    if slides:
        first = slides[0]
        kinds = {ph.placeholder_format.type for ph in first.placeholders}
        t = first.shapes.title
        ttext = t.text_frame.text.strip() if (t is not None and t.has_text_frame) else ""
        if ttext and (PP_PLACEHOLDER.CENTER_TITLE in kinds or PP_PLACEHOLDER.SUBTITLE in kinds):
            sub = next((ph for ph in first.placeholders if ph.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE), None)
            stext = sub.text_frame.text.strip() if (sub is not None and sub.has_text_frame) else ""
            used = {id(t._element)} | ({id(sub._element)} if sub is not None else set())
            other = [sh for sh in first.shapes if id(sh._element) not in used and
                     (getattr(sh, "has_text_frame", False) and sh.text_frame.text.strip()
                      or sh.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.TABLE, MSO_SHAPE_TYPE.CHART))]
            has_notes = first.has_notes_slide and first.notes_slide.notes_text_frame.text.strip()
            return (_sanitize_xml_text(ttext), _sanitize_xml_text(stext) or None,
                    not other and not has_notes)
    core = (prs.core_properties.title or "").strip()
    return (_sanitize_xml_text(core) if core else None), None, False


# ---------------------------------------------------------------- invisible structure markers
# Converting there and back must give back the original, so where a format
# can't express something (where one slide ends inside a Word handout, which
# heading level a slide title came from), it's noted invisibly — a hidden
# bookmark, a slide's internal name, a run's link-target name — never as
# words or labels anyone sees.
_DCT_SLIDE_RE = re.compile(r"^_dct_s(\d+)(c?)$")


def _docx_bookmark(paragraph, name, bid):
    """A zero-length bookmark (Word hides names starting "_")."""
    p = paragraph._p
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bid))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bid))
    pPr = p.find(qn("w:pPr"))
    idx = list(p).index(pPr) + 1 if pPr is not None else 0
    p.insert(idx, start)
    p.insert(idx + 1, end)


def _docx_slide_mark(paragraph):
    """(slide number, continues the previous slide's title) where a handout
    paragraph begins a slide's part of the document, else None."""
    for bm in paragraph._p.iter(qn("w:bookmarkStart")):
        m = _DCT_SLIDE_RE.match(bm.get(qn("w:name")) or "")
        if m:
            return int(m.group(1)), bool(m.group(2))
    return None


def _pptx_add_text(p, text):
    """Text into a slide paragraph, each "\\n" a real line break."""
    runs = []
    for k, seg in enumerate(str(text).split("\n")):
        if k:
            p.add_line_break()
        if seg:
            r = p.add_run()
            r.text = seg
            runs.append(r)
    return runs


def _slide_tag(slide, tag):
    """An invisible note on a slide (its internal name), e.g. "Heading 2"."""
    if tag:
        slide._element.cSld.set("name", tag)


def _slide_tag_of(slide):
    return slide._element.cSld.get("name") or ""


def _pptx_mark_runs(runs, tag):
    """An invisible note on a slide paragraph (its runs' link-target name)."""
    for r in runs[:1]:
        r._r.get_or_add_rPr().set("bmk", "dct:" + tag)


def _pptx_para_mark(para):
    for r in para._p.findall(qn("a:r")):
        rPr = r.find(qn("a:rPr"))
        if rPr is not None and (rPr.get("bmk") or "").startswith("dct:"):
            return rPr.get("bmk")[4:]
    return ""


def _docx_table_caption(table):
    """A Word table's title (its alt text) — where a spreadsheet tab's name rides along."""
    tblPr = table._tbl.tblPr
    cap = tblPr.find(qn("w:tblCaption")) if tblPr is not None else None
    return (cap.get(qn("w:val")) or "").strip() if cap is not None else ""


def _docx_set_table_caption(table, text):
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn("w:tblCaption")):
        tblPr.remove(old)
    cap = OxmlElement("w:tblCaption")
    cap.set(qn("w:val"), str(text)[:200])
    tblPr.append(cap)


def _notes_style(doc):
    """The handout's style for speaker notes — so converting back puts them
    back into the slide's notes instead of onto the slide."""
    try:
        return doc.styles["Speaker Notes"]
    except KeyError:
        st = doc.styles.add_style("Speaker Notes", WD_STYLE_TYPE.PARAGRAPH)
        st.base_style = doc.styles["Normal"]
        st.font.italic = True
        st.font.size = DocxPt(10)
        st.font.color.rgb = DocxRGBColor(0x59, 0x59, 0x59)
        st.paragraph_format.left_indent = DocxInches(0.3)
        st.paragraph_format.space_before = DocxPt(6)
        return st


def _pptx_inherited_bullet(shape, level):
    """Whether a slide paragraph with no bullet setting of its own shows one:
    looked up through its layout placeholder, the master placeholder and the
    master's body text style."""
    try:
        if not shape.is_placeholder:
            return False
        ph_type = shape.placeholder_format.type
        if ph_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.SUBTITLE,
                       PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.DATE):
            return False
        lvl = f"a:lvl{min(level, 8) + 1}pPr"
        chain = []
        try:
            lp = shape.part.slide_layout.placeholders.get(idx=shape.placeholder_format.idx)
            if lp is not None:
                chain.append(lp._element)
        except Exception:
            pass
        master = shape.part.slide_layout.slide_master
        for mp in master.placeholders:
            if mp.placeholder_format.type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT):
                chain.append(mp._element)
                break
        for el in chain:
            for pPr in el.iter(qn(lvl)):
                if pPr.find(qn("a:buNone")) is not None:
                    return False
                if pPr.find(qn("a:buChar")) is not None or pPr.find(qn("a:buAutoNum")) is not None or pPr.find(qn("a:buBlip")) is not None:
                    return True
        body_style = master._element.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}bodyStyle")
        if body_style is not None:
            pPr = body_style.find(qn(lvl))
            if pPr is not None:
                if pPr.find(qn("a:buNone")) is not None:
                    return False
                if pPr.find(qn("a:buChar")) is not None or pPr.find(qn("a:buAutoNum")) is not None:
                    return True
        return ph_type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT)
    except Exception:
        return False


def _pptx_para_kind(shape, para):
    """'number', 'bullet' or 'plain' — exactly how the paragraph shows on the slide."""
    pPr = para._p.find(qn("a:pPr"))
    if pPr is not None:
        if pPr.find(qn("a:buAutoNum")) is not None:
            return "number"
        if pPr.find(qn("a:buNone")) is not None:
            return "plain"
        if pPr.find(qn("a:buChar")) is not None or pPr.find(qn("a:buBlip")) is not None:
            return "bullet"
    return "bullet" if _pptx_inherited_bullet(shape, para.level or 0) else "plain"


def _docx_copy_pptx_runs(doc_p, para):
    """A slide paragraph's words into a Word paragraph: its runs (emphasis,
    colour, size, links), fields and line breaks, in order."""
    from pptx.text.text import _Run as _PptxRun
    for child in para._p:
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "br":
            doc_p.add_run().add_break()
            continue
        if tag not in ("r", "fld"):
            continue
        text = _sanitize_xml_text("".join(t.text or "" for t in child.iter(qn("a:t"))))
        if not text:
            continue
        if tag == "fld":
            doc_p.add_run(text)
            continue
        run = _PptxRun(child, para)
        address = None
        try:
            address = run.hyperlink.address
        except Exception:
            pass
        if address:
            _add_docx_hyperlink(doc_p, address, text, bold=bool(run.font.bold), italic=bool(run.font.italic))
            continue
        r = doc_p.add_run(text)
        if run.font.bold:
            r.bold = True
        if run.font.italic:
            r.italic = True
        if run.font.underline:
            r.underline = True
        hex_color = _safe_hex_color(run.font.color)
        if hex_color:
            try:
                r.font.color.rgb = DocxRGBColor.from_string(hex_color)
            except Exception:
                pass
        if run.font.size is not None:
            r.font.size = DocxPt(run.font.size.pt)


def _docx_text_of_pptx(text_frame):
    """A text frame's words with its line breaks (python-pptx gives them as \\v)."""
    return "\n".join(_sanitize_xml_text(p.text.replace("\x0b", "\n").replace("\v", "\n")).strip("\n") for p in text_frame.paragraphs).strip()


def pptx_to_docx(src_path, out_dir, style="clean"):
    """Slides → Handout. Everything on the slides, in reading order, and
    nothing that isn't: titles become headings, bullets stay bullets and plain
    text stays plain, line breaks and tables keep their shape, speaker notes
    go in their own style, and a deck's footer goes in the page footer. Each
    slide's start is marked with a hidden bookmark, so converting the handout
    back gives the same slides again."""
    prs = _safe_load(Presentation, src_path)
    doc = Document()
    zoom_el = doc.settings.element.find(qn("w:zoom"))
    if zoom_el is not None and zoom_el.get(qn("w:percent")) is None:
        zoom_el.set(qn("w:percent"), "100")
    deck_title = (prs.core_properties.title or "").strip()
    if deck_title:
        doc.core_properties.title = deck_title[:250]      # the file's own title property — never shown as text
    notes_style = _notes_style(doc)
    body = doc.element.body
    BULLETS = ["List Bullet", "List Bullet 2", "List Bullet 3"]
    NUMBERS = ["List Number", "List Number 2", "List Number 3"]
    FURNITURE = (PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.DATE)

    def kids():
        return [c for c in body if c.tag != qn("w:sectPr")]

    merged = {"slide": False}

    def add_text_paragraph(shape, para):
        if not (para.text or "").replace("\x0b", "").replace("\v", "").strip():
            return
        level = min(para.level or 0, 2)
        kind = _pptx_para_kind(shape, para)
        mark = _pptx_para_mark(para)
        if mark == "join":
            # the rest of a paragraph that was split over two slides: one paragraph again
            last = kids()[-1] if kids() else None
            if (last is not None and last.tag == qn("w:p") and last.find(".//" + qn("w:t")) is not None
                    and last.find(".//" + qn("w:drawing")) is None):
                prev = DocxParagraph(last, doc._body)
                if (prev.style.name or "") != notes_style.name:
                    prev.add_run(" ")
                    _docx_copy_pptx_runs(prev, para)
                    merged["slide"] = True
                    return
        style_name = None
        if mark.startswith("h") and mark[1:].isdigit():
            style_name = f"Heading {min(int(mark[1:]), 9)}"
        elif mark.startswith("style="):
            style_name = mark[6:]
        if style_name is None:
            style_name = (NUMBERS if kind == "number" else BULLETS)[level] if kind != "plain" else "Normal"
        try:
            p = doc.add_paragraph(style=style_name)
        except KeyError:
            p = doc.add_paragraph()
        if kind == "plain" and level and style_name == "Normal":
            p.paragraph_format.left_indent = DocxInches(0.3 * level)
        docx_align = _pptx_align_to_docx(para.alignment)
        if docx_align is not None:
            p.alignment = docx_align
        _docx_copy_pptx_runs(p, para)

    def add_table(shape):
        src_rows = [list(row.cells) for row in shape.table.rows]
        if not src_rows:
            return
        if (shape.name or "") in (_PPTX_TABLE_CONT, _PPTX_TABLE_CONT_OLD) and kids() and kids()[-1].tag == qn("w:tbl"):
            # the rest of a table carried over from the slide before (its first row repeated there)
            word_table = DocxTable(kids()[-1], doc)
            n_cols = len(word_table.columns)
            for row in src_rows[1:]:
                cells = word_table.add_row().cells
                for c_idx, src_cell in enumerate(row[:n_cols]):
                    if src_cell.is_spanned and not src_cell.is_merge_origin:
                        continue
                    for k, para in enumerate(src_cell.text_frame.paragraphs):
                        p = cells[c_idx].paragraphs[0] if k == 0 else cells[c_idx].add_paragraph()
                        _docx_copy_pptx_runs(p, para)
            merged["slide"] = True
            return
        n_rows, n_cols = len(src_rows), max(len(r) for r in src_rows)
        word_table = doc.add_table(rows=n_rows, cols=n_cols)
        word_table.style = "Table Grid"
        spans = []
        for r_idx, row in enumerate(src_rows):
            for c_idx in range(n_cols):
                src_cell = row[c_idx] if c_idx < len(row) else None
                if src_cell is None or (src_cell.is_spanned and not src_cell.is_merge_origin):
                    continue
                cell = word_table.cell(r_idx, c_idx)
                first = True
                for para in src_cell.text_frame.paragraphs:
                    p = cell.paragraphs[0] if first else cell.add_paragraph()
                    first = False
                    _docx_copy_pptx_runs(p, para)
                try:
                    if src_cell.fill.type == MSO_FILL_TYPE.SOLID:
                        hex_color = _safe_hex_color(src_cell.fill.fore_color)
                        if hex_color:
                            _set_docx_cell_shading(cell, hex_color)
                except Exception:
                    pass
                if src_cell.is_merge_origin and (src_cell.span_height > 1 or src_cell.span_width > 1):
                    spans.append((r_idx, c_idx, r_idx + src_cell.span_height - 1, c_idx + src_cell.span_width - 1))
        for a, b, c, d in spans:
            try:
                word_table.cell(a, b).merge(word_table.cell(c, d))
            except Exception:
                pass

    def add_chart(shape):
        try:
            chart = shape.chart
            plot = chart.plots[0]
            categories = [_sanitize_xml_text(str(cat)) for cat in plot.categories]
            series_list = list(plot.series)
        except Exception:
            return
        if not series_list:
            return
        try:
            if chart.has_title:
                t = _sanitize_xml_text(chart.chart_title.text_frame.text.strip())
                if t:
                    doc.add_paragraph(t)          # the chart's own title, as shown on the slide
        except Exception:
            pass
        chart_table = doc.add_table(rows=len(categories) + 1, cols=len(series_list) + 1)
        chart_table.style = "Table Grid"
        for s_idx, series in enumerate(series_list, 1):
            chart_table.cell(0, s_idx).text = _sanitize_xml_text(series.name or "")
        for cat_idx, cat_name in enumerate(categories, 1):
            chart_table.cell(cat_idx, 0).text = cat_name
            for s_idx, series in enumerate(series_list, 1):
                values = list(series.values)
                val = values[cat_idx - 1] if cat_idx - 1 < len(values) else None
                chart_table.cell(cat_idx, s_idx).text = _format_cell_value(val)

    slides = list(prs.slides)
    footers = []
    prev_title = None
    for n, slide in enumerate(slides, 1):
        start = len(kids())
        title_shape = slide.shapes.title
        title = _docx_text_of_pptx(title_shape.text_frame) if (title_shape is not None and title_shape.has_text_frame) else ""
        kinds = {ph.placeholder_format.type for ph in slide.placeholders}
        is_title_slide = PP_PLACEHOLDER.CENTER_TITLE in kinds or PP_PLACEHOLDER.SUBTITLE in kinds
        tag = _slide_tag_of(slide)
        cont = bool(title) and not is_title_slide and title == prev_title
        subtitle_ph = next((ph for ph in slide.placeholders if ph.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE), None)

        if title and is_title_slide:
            p = doc.add_paragraph(style="Title")
            _docx_copy_pptx_runs(p, title_shape.text_frame.paragraphs[0])
            for para in title_shape.text_frame.paragraphs[1:]:
                p.add_run().add_break()
                _docx_copy_pptx_runs(p, para)
        elif title and not cont and tag != "Topic":
            m = re.match(r"Heading (\d)$", tag)
            level = int(m.group(1)) if m else 1
            p = doc.add_paragraph(style="Title" if tag == "Title" else f"Heading {level}")
            for k, para in enumerate(title_shape.text_frame.paragraphs):
                if k:
                    p.add_run().add_break()
                _docx_copy_pptx_runs(p, para)
        elif (not title or tag == "Topic") and not cont and n > 1:
            _docx_divider(doc)      # a title the slides added (R22) isn't one of the document's headings
        merged["slide"] = False

        shapes = []
        for shape in _iter_flat_shapes(slide.shapes):
            if title_shape is not None and shape._element is title_shape._element:
                continue
            if shape.is_placeholder and shape.placeholder_format.type in FURNITURE:
                if shape.placeholder_format.type == PP_PLACEHOLDER.FOOTER and shape.has_text_frame and shape.text_frame.text.strip():
                    footers.append(_docx_text_of_pptx(shape.text_frame))
                continue
            shapes.append(shape)
        # Reading order: top to bottom, then left to right (rows within a quarter inch).
        shapes.sort(key=lambda sh: (round((sh.top or 0) / 228600), sh.left or 0))
        if subtitle_ph is not None:
            shapes.sort(key=lambda sh: 0 if sh._element is subtitle_ph._element else 1)
        pictures = []
        body_columns = 0
        for shape in shapes:
            if (shape.is_placeholder and shape.has_text_frame and shape.text_frame.text.strip()
                    and shape.placeholder_format.type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT)):
                body_columns += 1
                if body_columns == 2:
                    # The slide's second column: noted, so converting back gives two columns again.
                    col_start = len(kids())
                    for para in shape.text_frame.paragraphs:
                        add_text_paragraph(shape, para)
                    new_col = kids()[col_start:]
                    if new_col and new_col[0].tag == qn("w:p"):
                        _docx_bookmark(DocxParagraph(new_col[0], doc._body), f"_dct_col{n}", 19000 + n)
                    continue
            if shape.has_table:
                add_table(shape)
            elif getattr(shape, "has_chart", False):
                add_chart(shape)
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                pictures.append(shape)
            elif shape.has_text_frame and shape.text_frame.text.strip():
                if subtitle_ph is not None and shape._element is subtitle_ph._element and is_title_slide:
                    for para in shape.text_frame.paragraphs:
                        if para.text.strip():
                            mark = _pptx_para_mark(para)
                            st = mark[6:] if mark.startswith("style=") else "Subtitle"
                            try:
                                p = doc.add_paragraph(style=st)
                            except KeyError:
                                p = doc.add_paragraph(style="Subtitle")
                            _docx_copy_pptx_runs(p, para)
                    continue
                for para in shape.text_frame.paragraphs:
                    add_text_paragraph(shape, para)
        for _found, text in _iter_smartart_fallback_text(slide):
            if text:
                doc.add_paragraph(_sanitize_xml_text(text), style="List Bullet")
        # Pictures after the words, so converting back sets each beside its text.
        for shape in pictures:
            try:
                doc.add_picture(io.BytesIO(shape.image.blob), width=DocxInches(4))
            except Exception:
                continue
        if slide.has_notes_slide:
            for para in slide.notes_slide.notes_text_frame.paragraphs:
                if para.text.strip():
                    p = doc.add_paragraph(style=notes_style)
                    _docx_copy_pptx_runs(p, para)
                    _add_docx_callout_style(p)

        # The hidden bookmark at this slide's first paragraph.
        new = kids()[start:]
        if not new and merged["slide"]:
            prev_title = title if title else None
            continue                # all of it went back into what came before
        if not new or new[0].tag != qn("w:p"):
            anchor = OxmlElement("w:p")
            if new:
                new[0].addprevious(anchor)
            else:
                body.insert(len(kids()), anchor)
            first_p = DocxParagraph(anchor, doc._body)
        else:
            first_p = DocxParagraph(new[0], doc._body)
        _docx_bookmark(first_p, f"_dct_s{n}{'c' if cont else ''}", 9000 + n)
        prev_title = title if title else None

    if footers:
        common = Counter(footers).most_common(1)[0][0]
        fp = doc.sections[0].footer.paragraphs[0]
        for k, line in enumerate(common.split("\n")):
            if k:
                fp.add_run().add_break()
            fp.add_run(line)
    apply_docx_style(doc, style)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "converted.docx")
    doc.save(out_path)
    return out_path


# ---------------------------------------------------------------- PDF -> DOCX
_TABLE_MARK = "\u2063"   # invisible separator — never appears in real PDF text


def _mark_table_lines(text, tables):
    """pypdf reads a table's rows as ordinary lines of text, so every table
    appeared twice — once as a real table (added at the end of the page) and
    once as a jumbled paragraph of its cell values. This takes those lines
    out of the text and leaves a marker where the table's first row was, so
    the real table goes back in its actual place. Only rows with two or more
    filled cells are matched, so a lone number elsewhere is never removed."""
    if not tables or not text:
        return text
    norm = lambda s: re.sub(r"\s+", " ", s or "").strip().lower()
    row_owner = {}
    for t_idx, rows in enumerate(tables):
        for row in rows:
            cells = [str(c).strip() for c in row if c is not None and str(c).strip()]
            if len(cells) >= 2:
                row_owner.setdefault(norm(" ".join(cells)), t_idx)
    out, placed = [], set()
    for line in text.split("\n"):
        t_idx = row_owner.get(norm(line))
        if t_idx is None:
            out.append(line)
        elif t_idx not in placed:
            placed.add(t_idx)
            out.append(f"{_TABLE_MARK}{t_idx}{_TABLE_MARK}")
    return "\n".join(out)


def _split_table_markers(items):
    """Turns table markers inside reconstructed paragraphs into their own
    ("<index>", "table") items, wherever the paragraph joining put them."""
    pat = re.compile(re.escape(_TABLE_MARK) + r"(\d+)" + re.escape(_TABLE_MARK))
    out = []
    for text, kind in items:
        pos = 0
        for m in pat.finditer(text):
            before = text[pos:m.start()].strip()
            if before:
                out.append((before, kind))
            out.append((m.group(1), "table"))
            pos = m.end()
        rest = text[pos:].strip()
        if rest:
            out.append((rest, kind if pos == 0 else ("para" if kind.startswith("heading") else kind)))
    return out


def _pdf_drawn_images(src_path):
    """{page_index: [(image_name, top, text_just_above)]} for the images each
    page actually draws, top to bottom. pypdf lists every image a page's
    resources *reference* — PDFs that share one resource dictionary across
    pages (LibreOffice's do) made the same chart appear on every page. The
    line of text just above each image is used to put it back in place."""
    result = {}
    try:
        with pdfplumber.open(src_path) as pdf:
            for p_idx, page in enumerate(pdf.pages):
                imgs = sorted(page.images, key=lambda im: im.get("top", 0))
                if not imgs:
                    continue
                try:
                    lines = page.extract_text_lines()
                except Exception:
                    lines = []
                entries = []
                for im in imgs:
                    above = [ln for ln in lines if ln.get("bottom", 0) <= im.get("top", 0) + 1]
                    anchor = above[-1]["text"] if above else ""
                    entries.append((str(im.get("name", "")), im.get("top", 0), anchor))
                result[p_idx] = entries
    except Exception:
        return None
    return result


def _merge_wrapped_headings(items):
    """A heading that wrapped onto two lines in the PDF came through as two
    headings ("Introduction to Economics: Demand" / "and Supply"). Joins a
    heading with the next one at the same level when that one clearly
    continues it: it starts in lowercase, or the first ends on a joining word."""
    out = []
    joiners = ("and", "or", "of", "the", "to", "for", "in", "with", "on", "&", "a", "an")
    for text, kind in items:
        if (out and kind.startswith("heading") and out[-1][1] == kind and text
                and (text[:1].islower() or out[-1][0].rstrip().lower().split()[-1:] and
                     out[-1][0].rstrip().lower().split()[-1] in joiners)
                and not re.search(r"[.!?:]$", out[-1][0].rstrip())):
            out[-1] = (out[-1][0].rstrip() + " " + text.strip(), kind)
        else:
            out.append((text, kind))
    return out


def pdf_to_docx(src_path, out_dir, style="clean"):
    reader = _safe_load(pypdf.PdfReader, src_path)
    if reader.is_encrypted:
        raise ConversionError("This PDF is password-protected and can't be converted until it's unlocked.")
    n_pages = len(reader.pages)
    os.makedirs(out_dir, exist_ok=True)  # needed early: OCR (if any page requires it,
    # below) writes a scratch rasterized PNG into out_dir well before the
    # final document save, which is the only place this used to be called


    # A single pdfplumber.open() for the whole function, not one per
    # page: confirmed directly (profiled against a realistic 10-page
    # document) that _extract_pdf_ruled_tables previously re-opened and
    # re-parsed the entire file from scratch on every single page just
    # to use one page of the result, and _extract_pdf_font_size_headings
    # opened it separately again on top of that — eleven total opens of
    # the same file for one document. try/except here preserves the
    # same graceful fallback _extract_pdf_font_size_headings always had
    # for a PDF structure pdfplumber itself can't open at all: text
    # extraction still proceeds via pypdf below, just without the
    # font-size heading detection or ruled-table extraction this pass
    # adds on top of it.
    tables_by_page = {}
    try:
        with pdfplumber.open(src_path) as pdfplumber_pdf:
            font_headings_by_page, _body_size = _extract_pdf_font_size_headings(pdfplumber_pdf)
            # Ruled tables (real grid lines, not just aligned text — see
            # _extract_pdf_ruled_tables for why the more permissive strategy
            # isn't used) are pulled out per page so they can be rebuilt as real
            # Word tables rather than left to come through the ordinary text
            # flow as a garbled run of individual cell values.
            MAX_TABLE_PAGES = 40  # capped for very long documents
            if n_pages <= MAX_TABLE_PAGES:
                for page_idx in range(n_pages):
                    tables = _extract_pdf_ruled_tables(pdfplumber_pdf, page_idx)
                    if tables:
                        tables_by_page[page_idx] = tables
    except Exception:
        font_headings_by_page, _body_size = {}, None

    page_items = []
    for page_idx, page in enumerate(reader.pages):
        raw_text = page.extract_text() or ""
        page_tables = tables_by_page.get(page_idx, [])
        if page_tables:
            # Remove each table's own cell text from the plain line
            # stream before paragraph reconstruction sees it — otherwise
            # the same content appears twice: once correctly as a real
            # table, once again as a garbled run of individual values
            # from the ordinary text-extraction path, which has no
            # concept of the table's row/column structure. A cell's text
            # is split on embedded newlines before being added to the
            # filter set — pdfplumber sometimes merges two visually-close
            # values into one cell with a line break between them, and
            # pypdf's own line-by-line extraction still produces those as
            # two separate lines, so matching against the whole
            # (possibly multi-line) cell value as a single unit misses
            # them entirely, confirmed directly against a real merged
            # cell leaking into the surrounding paragraph text.
            table_cell_texts = {
                part.strip()
                for table in page_tables
                for row in table
                for cell in row
                if cell
                for part in str(cell).split("\n")
                if part.strip()
            }
            raw_text = "\n".join(
                line for line in raw_text.split("\n") if line.strip() not in table_cell_texts
            )
        text = _sanitize_xml_text(raw_text.strip())
        page_font_headings = font_headings_by_page.get(page_idx, {})
        text = _mark_table_lines(text, tables_by_page.get(page_idx, []))
        page_items.append(_split_table_markers(_reconstruct_paragraphs(text, page_font_headings)) if text else [])

    # A short line that repeats verbatim across most pages is usually a
    # running header/footer (page title, "Confidential", a date stamp,
    # etc.), not real content — repeating it once per page in the
    # reconstructed document just adds clutter. But the same short text
    # can coincidentally also be a genuine section heading on one specific
    # page (e.g. "Summary" as a running header everywhere, but the actual
    # heading for real summary content on the one page that has it) —
    # suppressing every occurrence outright would silently erase that
    # page's real heading along with the boilerplate. Only applies to
    # documents long enough that a real repeat pattern is meaningful, not
    # a 2-page coincidence.
    line_counts = Counter()
    for items in page_items:
        seen_this_page = {t for t, _k in items if len(t) < 100}
        for t in seen_this_page:
            line_counts[t] += 1
    repeat_threshold = max(3, int(n_pages * 0.6))
    noisy_lines = {t for t, count in line_counts.items() if n_pages > 2 and count >= repeat_threshold}

    # For each noisy line, measure how substantial the content immediately
    # following it is on every page where it appears. A specific occurrence
    # is exempted from suppression only when what follows it is both a real
    # paragraph in absolute terms and clearly longer than what typically
    # follows the same line elsewhere — true boilerplate is followed by
    # similarly-sized content everywhere (nothing stands out), while a
    # coincidentally-repeated real heading precedes one page's genuinely
    # substantial section.
    follow_lengths_by_line = {t: [] for t in noisy_lines}
    for items in page_items:
        for item_idx, (t, _k) in enumerate(items):
            if t in noisy_lines:
                flen = len(items[item_idx + 1][0]) if item_idx + 1 < len(items) else 0
                follow_lengths_by_line[t].append(flen)
    medians = {t: statistics.median(lens) if lens else 0 for t, lens in follow_lengths_by_line.items()}
    exempt_occurrences = set()  # (page_idx, item_idx) pairs kept despite matching a noisy line
    for page_idx, items in enumerate(page_items):
        for item_idx, (t, _k) in enumerate(items):
            if t not in noisy_lines:
                continue
            flen = len(items[item_idx + 1][0]) if item_idx + 1 < len(items) else 0
            if flen >= 40 and flen >= medians[t] * 1.8:
                exempt_occurrences.add((page_idx, item_idx))

    doc = Document()
    # python-docx's own default template ships a <w:zoom val="bestFit"/>
    # missing the "percent" attribute the OOXML schema actually requires
    # on it — same pre-existing, library-wide quirk found and fixed in
    # pptx_to_docx, present in every fresh Document() regardless of which
    # function creates it.
    zoom_el = doc.settings.element.find(qn("w:zoom"))
    if zoom_el is not None and zoom_el.get(qn("w:percent")) is None:
        zoom_el.set(qn("w:percent"), "100")
    # Headers/footers repeated on every page are dropped first, so the
    # rearranging below can't disturb which occurrences were exempted.
    filtered_pages = [
        [(t, k) for item_idx, (t, k) in enumerate(items)
         if t not in noisy_lines or (p_idx, item_idx) in exempt_occurrences]
        for p_idx, items in enumerate(page_items)
    ]
    filtered_pages = [_merge_wrapped_headings(items) for items in filtered_pages]

    def _lvl(kind):
        tail = kind[len("heading"):]
        return int(tail) if tail.isdigit() else 3
    all_levels = [_lvl(k) for items in filtered_pages for _t, k in items if k.startswith("heading")]
    doc_title = None
    if filtered_pages and filtered_pages[0] and filtered_pages[0][0][1].startswith("heading") \
            and all_levels and _lvl(filtered_pages[0][0][1]) == min(all_levels):
        # The document's own first, most prominent heading is its title. The
        # Word file used to be titled "Converted from PDF" instead.
        doc_title = filtered_pages[0].pop(0)[0]
    if not doc_title:
        meta_title = ""
        try:
            meta_title = (reader.metadata.title or "").strip() if reader.metadata else ""
        except Exception:
            pass
        doc_title = _sanitize_xml_text(meta_title) if meta_title else None
    if doc_title:
        doc.add_heading(doc_title, 0)
    rest_levels = [_lvl(k) for items in filtered_pages for _t, k in items if k.startswith("heading")]
    base_level = min(rest_levels) if rest_levels else 1

    def _word_level(pdf_level):
        # The biggest remaining headings become Heading 1. They used to be
        # pushed down two levels to sit under "Page N" headings, so a main
        # section like "What is demand?" came out as Heading 4.
        return max(1, min(pdf_level - base_level + 1, 9))

    drawn_images = _pdf_drawn_images(src_path)
    seen_image_hashes = set()
    emitted_tables_by_page = {}
    emitted_images_by_page = {}
    if drawn_images:
        norm = lambda s: re.sub(r"\s+", " ", s or "").strip()
        for p_idx, entries in drawn_images.items():
            if p_idx >= len(filtered_pages):
                continue
            items = filtered_pages[p_idx]
            for name, _top, anchor in entries:
                a = norm(anchor)[-40:]
                pos = None
                if a:
                    for k in range(len(items) - 1, -1, -1):
                        if items[k][1] not in ("table", "image") and a in norm(items[k][0]):
                            pos = k + 1
                            break
                if pos is None and not a:
                    pos = 0        # nothing above it: the picture opens the page
                if pos is not None:
                    items.insert(pos, (name, "image"))

    def _emit_table(table_rows):
        clean_rows = [
            [(_sanitize_xml_text(str(cell).strip()) if cell else "") for cell in row]
            for row in table_rows if any(cell and str(cell).strip() for cell in row)
        ]
        if not clean_rows:
            return
        n_cols = max(len(r) for r in clean_rows)
        word_table = doc.add_table(rows=len(clean_rows), cols=n_cols)
        word_table.style = "Light Grid Accent 1"
        for r_idx, row in enumerate(clean_rows):
            for c_idx in range(n_cols):
                cell = word_table.rows[r_idx].cells[c_idx]
                cell.text = row[c_idx] if c_idx < len(row) else ""
                if r_idx == 0:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            run.bold = True
        doc.add_paragraph()  # breathing room after the table before whatever comes next

    def _emit_image(page, page_idx, name):
        nonlocal image_count
        if image_count >= MAX_IMAGES:
            return
        emitted_images_by_page.setdefault(page_idx, set()).add(name)
        try:
            by_name = {os.path.splitext(img.name)[0]: img for img in page.images}
        except Exception:
            return
        img = by_name.get(name)
        if img is None:
            return
        digest = hashlib.sha1(img.data).hexdigest()
        if digest in seen_image_hashes:
            return   # the same picture (a logo on every page) goes in once
        try:
            doc.add_picture(io.BytesIO(img.data), width=DocxInches(6))
            seen_image_hashes.add(digest)
            image_count += 1
        except Exception:
            pass  # a malformed/unsupported embedded image shouldn't sink the whole conversion

    # A real, clickable table of contents built from the PDF's own
    # most-prominent detected headings (level 1 — the document's own
    # section titles, not the sequential "Page N" markers, which aren't
    # meaningful document structure on their own) — only worth adding
    # once there's enough real structure to summarize.
    level1_headings = []
    for page_idx in sorted(font_headings_by_page.keys()):
        for text, lvl in font_headings_by_page[page_idx].items():
            if lvl == 1:
                level1_headings.append(text)
    # (No contents page: the Word file mirrors the PDF and adds nothing.)

    MAX_IMAGES = 30
    image_count = 0
    MAX_OCR_PAGES = 15  # OCR is comparatively expensive (rasterize + tesseract
    # per page); this covers a few scanned pages mixed into an otherwise
    # text-based PDF, or a fully-scanned document up to a reasonable length,
    # without unbounded processing time on a very long scanned file.
    ocr_page_count = 0

    for i, (page, items) in enumerate(zip(reader.pages, page_items), 1):
        page_links = _extract_pdf_page_links(page)
        unmatched_links = list(page_links)
        page_idx = i - 1
        content_items = filtered_pages[page_idx]
        if content_items:
            for para_text, kind in content_items:
                if not para_text:
                    continue
                if kind == "table":
                    t_list = tables_by_page.get(page_idx, [])
                    if para_text.isdigit() and int(para_text) < len(t_list):
                        emitted_tables_by_page.setdefault(page_idx, set()).add(int(para_text))
                        _emit_table(t_list[int(para_text)])
                    continue
                if kind == "image":
                    _emit_image(page, page_idx, para_text)
                    continue
                candidates = [(lt, url) for lt, url in page_links if lt in para_text]
                if kind == "bullet":
                    _p, used = _add_docx_paragraph_with_links(doc, para_text, candidates, style="List Bullet")
                elif kind == "number":
                    _p, used = _add_docx_paragraph_with_links(doc, para_text, candidates, style="List Number")
                elif kind.startswith("heading"):
                    # "Page N" (when present) already occupies level 2, so
                    # the PDF's own content headings nest under it starting
                    # at level 3 — pdf_level 1 (the document's biggest,
                    # most prominent heading) becomes Word level 3, and so
                    # on, rather than competing with "Page N" for the same
                    # level or going shallower than it.
                    pdf_level = int(kind[len("heading"):]) if kind[len("heading"):].isdigit() else 3
                    doc.add_heading(para_text, level=_word_level(pdf_level))
                    used = []
                else:
                    _p, used = _add_docx_paragraph_with_links(doc, para_text, candidates)
                for m in used:
                    if m in unmatched_links:
                        unmatched_links.remove(m)
        elif not items:
            ocr_text = ""
            if ocr_page_count < MAX_OCR_PAGES:
                ocr_page_count += 1
                ocr_text = _ocr_pdf_page(src_path, i, out_dir)
            if ocr_text:
                for para_text, kind in _reconstruct_paragraphs(_sanitize_xml_text(ocr_text), allow_heading_fallback=False):
                    if not para_text:
                        continue
                    if kind == "bullet":
                        doc.add_paragraph(para_text, style="List Bullet")
                    elif kind == "number":
                        doc.add_paragraph(para_text, style="List Number")
                    elif kind.startswith("heading"):
                        doc.add_heading(para_text, level=2)
                    else:
                        doc.add_paragraph(para_text)
                note_p = doc.add_paragraph()
                note_run = note_p.add_run("(Text on this page was recovered via OCR from a scanned image and may contain recognition errors.)")
                note_run.italic = True
                note_run.font.size = DocxPt(9)
            else:
                doc.add_paragraph("[No extractable text on this page — likely a scanned image, and OCR did not recover any text.]")
        # else: the page had only repeated header/footer noise and nothing
        # else — nothing worth showing, so leave it at just the page heading.

        # Anything that couldn't be put in place (a table whose rows didn't
        # match the text, a picture with no text above it to anchor to)
        # still goes in, at the end of its page as before.
        for t_idx, table_rows in enumerate(tables_by_page.get(page_idx, [])):
            if t_idx not in emitted_tables_by_page.get(page_idx, set()):
                _emit_table(table_rows)
        if drawn_images is not None:
            leftover = [n for n, _t, _a in drawn_images.get(page_idx, [])
                        if n not in emitted_images_by_page.get(page_idx, set())]
        else:
            try:
                leftover = [os.path.splitext(img.name)[0] for img in page.images]
            except Exception:
                leftover = []
        for name in leftover:
            _emit_image(page, page_idx, name)

        if unmatched_links:
            # A link that never matched any reconstructed paragraph text
            # (an image-only link, or a reflow edge case) still shouldn't
            # be silently dropped — list its real URL instead of guessing
            # where it belongs.
            link_p = doc.add_paragraph()
            link_p.paragraph_format.space_before = DocxPt(8)
            label_run = link_p.add_run("Other links on this page: ")
            label_run.italic = True
            label_run.bold = True
            for j, (_text, url) in enumerate(unmatched_links):
                if j > 0:
                    link_p.add_run(", ")
                _add_docx_hyperlink(link_p, url, url)

    apply_docx_style(doc, style)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "converted.docx")
    doc.save(out_path)
    return out_path


# ---------------------------------------------------------------- PDF -> PPTX
def _add_fitted_picture(prs, slide, img_path, top=0, bottom=0):
    """Adds a picture scaled to fit the slide (below `top`, above `bottom`) and
    centred, keeping its proportions."""
    from PIL import Image
    with Image.open(img_path) as im:
        iw, ih = im.size
    W, H = prs.slide_width, prs.slide_height - top - bottom
    scale = min(W / iw, H / ih)
    pw, ph = int(iw * scale), int(ih * scale)
    return slide.shapes.add_picture(img_path, (W - pw) // 2, top + (H - ph) // 2, pw, ph)


def _ocr_page_with_headings(img_path):
    """OCR one page image using tesseract's word boxes, so a scan — which
    carries no font information — still yields headings. Plain OCR text made
    every heading look like body text (one byline ended up titling six
    slides in a row). Two clues, measured on each line:
      - size: its tallest word against a typical body line (catches titles);
      - spacing: extra white space above it and a normal gap below — how
        headings sit in almost every document, and the clue that survives
        scanning (at scan resolution a 14pt heading and 11pt text can measure
        the same height).
    Also turns the bullet marks OCR tends to misread ("¢", "«", "°") back
    into bullets. Returns (text, {line_text: level})."""
    try:
        r = subprocess.run(["tesseract", img_path, "-", "--psm", "3", "tsv"],
                           capture_output=True, text=True, timeout=45)
    except Exception:
        return "", {}
    lines, order = {}, []
    for row in (r.stdout or "").splitlines()[1:]:
        parts = row.split("\t")
        if len(parts) < 12 or parts[0] != "5" or not parts[11].strip():
            continue
        try:
            if float(parts[10]) < 0:
                continue
            key = (int(parts[1]), int(parts[2]), int(parts[3]), int(parts[4]))
            top, h = int(parts[7]), int(parts[9])
        except ValueError:
            continue
        if key not in lines:
            lines[key] = []
            order.append(key)
        lines[key].append((parts[11].strip(), top, h))
    if not order:
        return "", {}
    info = []
    for k in order:
        ws = lines[k]
        info.append({
            "words": [w for w, _t, _h in ws],
            "tall": max(h for _w, _t, h in ws),
            "top": min(t for _w, t, _h in ws),
            "bottom": max(t + h for _w, t, h in ws),
        })
    body_tall = sorted(i["tall"] for i in info if len(i["words"]) >= 5) or sorted(i["tall"] for i in info)
    body = body_tall[len(body_tall) // 2] or 1
    gaps = [None] + [max(0, info[j]["top"] - info[j - 1]["bottom"]) for j in range(1, len(info))]
    normal_gaps = sorted(g for g in gaps[1:] if g is not None) or [0]
    gap_typ = normal_gaps[len(normal_gaps) // 2] or 1
    BULLET_MISREADS = {"¢", "«", "»", "°", "©", "*", "e", "o", "+"}
    text_lines, headings = [], {}
    big_line_before = False
    for j, i in enumerate(info):
        words = list(i["words"])
        if len(words) > 3 and words[0] in BULLET_MISREADS and words[1][:1].isupper():
            words[0] = "\u2022"
        t = " ".join(words)
        text_lines.append(t)
        letters = sum(ch.isalpha() for ch in t)
        looks_like_words = (letters >= 0.6 * len(t.replace(" ", ""))
                            and any(len(w) >= 3 and w[:1].isalpha() for w in words)
                            and words[0] != "\u2022")
        if not looks_like_words or len(words) > 12:
            continue
        if i["tall"] >= 1.2 * body and (t[:1].isupper() or big_line_before):
            # Large type starting in lowercase only counts when it directly
            # continues a large heading line — the second line of a wrapped
            # title ("and Supply"), joined back up later — never on its own,
            # where it's usually stray OCR from inside a picture.
            headings[t] = 1 if i["tall"] >= 1.7 * body else 2
            big_line_before = True
            continue
        big_line_before = False
        if not t[:1].isupper():
            continue
        nxt = info[j + 1]["words"] if j + 1 < len(info) else []
        nxt_text = "".join(nxt)
        if nxt_text and sum(ch.isdigit() for ch in nxt_text) >= 0.5 * len(nxt_text):
            continue   # a line followed by rows of numbers is a table's header row, not a heading
        gap_above = gaps[j] if gaps[j] is not None else 0
        gap_below = gaps[j + 1] if j + 1 < len(gaps) else gap_typ
        if (len(words) <= 10 and gap_above >= 1.8 * gap_typ + 4 and gap_below <= 1.3 * gap_typ + 2
                and not t.rstrip().endswith((".", ",", ";", ":"))):
            headings[t] = 2
    return "\n".join(text_lines), headings

def _pdf_pages_to_editable_deck(src_path, out_dir, page_images, style):
    """Text per page (the PDF's own text, else OCR of the rendered page) → a
    structured Word document → the Notes → Slides engine. Returns the .pptx
    path, or None when there isn't enough readable text for it to be worth it."""
    reader = _safe_load(pypdf.PdfReader, src_path)
    try:   # real font sizes for pages that have a text layer
        with pdfplumber.open(src_path) as ppdf:
            font_headings_by_page, _body = _extract_pdf_font_size_headings(ppdf)
    except Exception:
        font_headings_by_page = {}
    MAX_OCR = 15
    ocr_used, page_items = 0, []
    for idx, img_name in enumerate(page_images):
        text, headings = "", font_headings_by_page.get(idx, {})
        if idx < len(reader.pages):
            try:
                text = reader.pages[idx].extract_text() or ""
            except Exception:
                text = ""
        if len(text.strip()) < 25 and ocr_used < MAX_OCR:
            ocr_used += 1
            text, headings = _ocr_page_with_headings(os.path.join(out_dir, img_name))
        text = _sanitize_xml_text(text)
        page_items.append(_merge_wrapped_headings(_reconstruct_paragraphs(text, headings)) if text.strip() else [])
    if sum(len(t) for items in page_items for t, _k in items) < 80:
        return None
    notes = Document()
    title_done = False
    for items in page_items:
        for text, kind in items:
            if not text:
                continue
            if kind.startswith("heading"):
                if not title_done:
                    notes.add_heading(text, 0)
                    title_done = True
                else:
                    tail = kind[len("heading"):]
                    notes.add_heading(text, min(int(tail), 3) if tail.isdigit() else 2)
            elif kind == "bullet":
                notes.add_paragraph(text, style="List Bullet")
            elif kind == "number":
                notes.add_paragraph(text, style="List Number")
            else:
                notes.add_paragraph(text)
    notes_path = os.path.join(out_dir, "_page_text.docx")
    notes.save(notes_path)
    look = style or "plain"
    deck_path = docx_to_pptx(notes_path, out_dir, style=look)
    prs = Presentation(deck_path)
    layout = prs.slide_layouts[5] if len(prs.slide_layouts) > 5 else prs.slide_layouts[-1]
    look_style = PPTX_TEMPLATES.get(look, PPTX_TEMPLATES["plain"])
    # R22: each original page is titled with its own heading (or the heading it falls under,
    # else a topic in its own words) — never an invented "Original page N"
    page_titles, last = [], ""
    for items in page_items:
        head = next((t for t, k in items if t and k.startswith("heading")), "")
        last = head or last
        page_titles.append(head or last or _topic_from_text(next((t for t, _k in items if t), "")))
    for n, img_name in enumerate(page_images, 1):
        s = prs.slides.add_slide(layout)
        page_title = page_titles[n - 1] if n - 1 < len(page_titles) else ""
        top = PptxInches(0.35)
        if s.shapes.title is not None and page_title:
            s.shapes.title.text = page_title
            _style_pptx_slide(s, look_style)
            _slide_tag(s, "Topic")
            top = s.shapes.title.top + s.shapes.title.height + PptxInches(0.1)
        elif s.shapes.title is not None:
            title_el = s.shapes.title._element
            title_el.getparent().remove(title_el)
        _add_fitted_picture(prs, s, os.path.join(out_dir, img_name), top=top, bottom=PptxInches(0.3))
    prs.save(deck_path)
    return deck_path


def pdf_to_pptx(src_path, out_dir, style=None):
    MAX_PDF_PPTX_PAGES = 50
    try:
        reader = _safe_load(pypdf.PdfReader, src_path)
        page_count = len(reader.pages)
    except ConversionError:
        raise
    if page_count > MAX_PDF_PPTX_PAGES:
        raise ConversionError(
            f"This PDF has {page_count} pages — each one needs to be rendered and OCR'd "
            f"individually for this conversion, which isn't practical past "
            f"{MAX_PDF_PPTX_PAGES} pages. Try splitting it into smaller sections first."
        )
    page_prefix = os.path.join(out_dir, "page")
    try:
        subprocess.run(
            ["pdftoppm", "-png", "-r", "150", src_path, page_prefix],
            capture_output=True, text=True, timeout=60, check=True
        )
    except FileNotFoundError:
        raise ConversionError("PDF rendering tool (poppler-utils) is not installed on the server")
    except subprocess.TimeoutExpired:
        raise ConversionError("PDF rendering timed out — the file may be too large or complex")
    except subprocess.CalledProcessError as e:
        raise ConversionError(f"Could not rasterize PDF: {(e.stderr or e.stdout or '').strip()[:300]}")
    page_images = sorted(
        f for f in os.listdir(out_dir) if f.startswith("page") and f.endswith(".png")
    )
    if not page_images:
        raise ConversionError("Could not rasterize PDF pages")

    # ---- Document pages (portrait: notes, handouts, scans) become real,
    # editable slides. Each page used to become a picture forced to the
    # slide's exact width and height — a portrait page squashed into a wide
    # slide, distorting the text — with nothing editable. The text is read
    # (straight from the PDF, or by OCR for a scan), rebuilt into headings,
    # bullets and numbered steps, and laid out by the same engine as Notes →
    # Slides. The original pages follow at the end, undistorted, so nothing
    # on them (diagrams, handwriting) is lost.
    try:
        first = reader.pages[0]
        portrait = float(first.mediabox.height) >= float(first.mediabox.width)
    except Exception:
        portrait = False
    if portrait:
        try:
            deck_path = _pdf_pages_to_editable_deck(src_path, out_dir, page_images, style)
        except ConversionError:
            raise
        except Exception:
            deck_path = None
        if deck_path:
            return deck_path

    # ---- Slide-deck PDFs (landscape pages), or pages with no readable
    # text: one picture per slide, scaled to fit rather than stretched, with
    # the page's OCR text in the speaker notes.
    prs = Presentation()
    prs.slide_width = PptxInches(13.333)
    prs.slide_height = PptxInches(7.5)
    _fit_layouts_to_widescreen(prs)  # same off-centre, 29%-empty-slide fix as Notes → Slides
    blank_layout = prs.slide_layouts[6]
    from PIL import Image
    for img_name in page_images:
        png_path = os.path.join(out_dir, img_name)
        img_path = png_path
        try:
            jpeg_path = png_path[:-4] + ".jpg"
            with Image.open(png_path) as im:
                im.convert("RGB").save(jpeg_path, "JPEG", quality=85)
            if os.path.getsize(jpeg_path) < os.path.getsize(png_path):
                img_path = jpeg_path
            else:
                os.remove(jpeg_path)
        except Exception:
            pass  # fall back to the original PNG if re-encoding fails for any reason
        slide = prs.slides.add_slide(blank_layout)
        _add_fitted_picture(prs, slide, img_path)
        ocr_text = ""
        try:
            ocr = subprocess.run(
                ["tesseract", img_path, "-", "--psm", "6"],
                capture_output=True, text=True, timeout=30
            )
            ocr_text = ocr.stdout.strip()
        except Exception:
            pass
        notes = slide.notes_slide
        notes.notes_text_frame.text = (
            ocr_text if ocr_text else "(No text detected by OCR on this page.)"
        )
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "converted.pptx")
    prs.save(out_path)
    return out_path


def _count_format_decimals(fmt):
    """Counts decimal placeholder digits (0 or #) after the last '.' in an
    Excel number format string, e.g. '0.00' -> 2, '#,##0' -> 0. Stops at
    the first character that isn't a placeholder digit, so a trailing '%'
    or literal text doesn't get miscounted as more decimals."""
    if "." not in fmt:
        return 0
    after_dot = fmt.rsplit(".", 1)[1]
    count = 0
    for ch in after_dot:
        if ch in "0#":
            count += 1
        else:
            break
    return count


_XL_LOCALE_RE = re.compile(r"\[\$([^\]-]*)(?:-[0-9A-Fa-f]+)?\]")


def _xl_literal_text(part):
    """The literal text of an Excel number-format fragment, as Excel shows
    it: quoted text ("GH¢"), backslash-escaped characters, and currency from
    locale codes ([$€-2] → €) kept; colour/condition codes ([Red]), spacing
    (_)) and fill (*-) codes removed."""
    part = _XL_LOCALE_RE.sub(lambda m: m.group(1), part)
    part = re.sub(r"\[[^\]]*\]", "", part)
    part = re.sub(r"_.", "", part)
    part = re.sub(r"\*.", "", part)
    out, i = [], 0
    while i < len(part):
        ch = part[i]
        if ch == '"':
            j = part.find('"', i + 1)
            j = len(part) if j == -1 else j
            out.append(part[i + 1:j])
            i = j + 1
            continue
        if ch == "\\" and i + 1 < len(part):
            out.append(part[i + 1])
            i += 2
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def _xl_split_number_format(fmt):
    """(prefix, digits, suffix) around a format's digit placeholders,
    ignoring anything inside quotes or [brackets]."""
    in_q = in_b = False
    first = last = None
    i = 0
    while i < len(fmt):
        ch = fmt[i]
        if in_q:
            in_q = ch != '"'
        elif in_b:
            in_b = ch != "]"
        elif ch == '"':
            in_q = True
        elif ch == "[":
            in_b = True
        elif ch == "\\":
            i += 1
        elif ch in "0#?":
            first = i if first is None else first
            last = i
        i += 1
    if first is None:
        return fmt, "", ""
    return fmt[:first], fmt[first:last + 1], fmt[last + 1:]


_XL_DATE_TOKENS = re.compile(r'"[^"]*"|\\.|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am/pm|a/p|.', re.I)


def _excel_date_format_to_strftime(fmt):
    """Translates an Excel date/time display format into a strftime pattern,
    or None when it isn't one. Reports used to print every date as
    2026-09-05 whatever the sheet said; now "dd mmm yyyy" gives 05 Sep 2026
    and "dd/mm/yyyy hh:mm" gives 30/08/2026 14:05. "m"/"mm" means minutes
    right after an hour or right before seconds, months everywhere else —
    the same rule Excel uses."""
    if not fmt or fmt == "General":
        return None
    f = re.sub(r"\[[^\]]*\]", "", fmt.split(";")[0])
    if not re.search(r"[dmyhs]", re.sub(r'"[^"]*"', "", f), re.I):
        return None
    tokens = _XL_DATE_TOKENS.findall(f)
    has_ampm = bool(re.search(r"am/pm|a/p", f, re.I))
    is_code = lambda t: bool(re.fullmatch(r"yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am/pm|a/p", t, re.I))
    table = {"yyyy": "%Y", "yy": "%y", "mmmmm": "%b", "mmmm": "%B", "mmm": "%b", "mm": "%m", "m": "%-m",
             "dddd": "%A", "ddd": "%a", "dd": "%d", "d": "%-d", "ss": "%S", "s": "%-S", "am/pm": "%p", "a/p": "%p",
             "hh": "%I" if has_ampm else "%H", "h": "%-I" if has_ampm else "%-H"}
    out = []
    for idx, tok in enumerate(tokens):
        t = tok.lower()
        if not is_code(tok):
            lit = tok[1:-1] if tok.startswith('"') else (tok[1:] if tok.startswith("\\") else tok)
            out.append(lit.replace("%", "%%"))
            continue
        if t in ("m", "mm"):
            prev = next((x.lower() for x in reversed(tokens[:idx]) if is_code(x)), "")
            nxt = next((x.lower() for x in tokens[idx + 1:] if is_code(x)), "")
            if prev in ("h", "hh") or nxt in ("s", "ss"):
                out.append("%M")
                continue
        out.append(table[t])
    return "".join(out)


def _format_cell_value(val, number_format=None):
    """openpyxl hands back raw Python values — a date becomes a datetime
    object, and floating-point arithmetic in the sheet often leaves noise
    like 3.140000000000001. str()'ing these directly is what the old code
    did, and it looked exactly as raw as that implies. This renders them
    the way a person actually reads a spreadsheet — including honoring
    the cell's actual display format for the common cases (percentage,
    currency, thousands separator), since the raw stored value and what
    Excel actually shows can be completely different: a cell storing 0.15
    with a percentage format displays as '15%', not '0.15'."""
    if val is None:
        return ""
    if isinstance(val, bool):
        return "TRUE" if val else "FALSE"
    if isinstance(val, (datetime.datetime, datetime.date)):
        pattern = _excel_date_format_to_strftime(number_format)
        if pattern:
            try:
                return val.strftime(pattern)
            except (ValueError, TypeError):
                pass   # an unusual pattern — fall back to the unambiguous form below
        if isinstance(val, datetime.datetime) and (val.hour or val.minute or val.second):
            return val.strftime("%Y-%m-%d %H:%M")
        return val.strftime("%Y-%m-%d")
    if isinstance(val, (int, float)) and number_format and number_format != "General":
        # Excel format strings can have up to 4 semicolon-separated
        # sections: positive;negative;zero;text. When an explicit negative
        # section exists (accounting-style '#,##0.00;(#,##0.00)' is the
        # common case), Excel uses THAT section's own formatting for
        # negative numbers — usually wrapping them in parentheses instead
        # of a plain minus sign — rather than just prepending '-' to the
        # positive format.
        sections = number_format.split(";")
        fmt = sections[0]
        work_val = val
        use_parens = False
        if val < 0 and len(sections) > 1:
            fmt = sections[1]
            use_parens = "(" in fmt
            if use_parens:
                # Only strip the sign when parentheses are what convey it —
                # otherwise leave work_val negative so the normal formatting
                # below (Python's own '-', or the currency-sign handling)
                # produces the negative sign itself, instead of a section
                # like '0%;-0%' or '0.00;-0.00' silently losing the sign
                # entirely because nothing was left to represent it.
                work_val = -val
        elif val == 0 and len(sections) > 2:
            fmt = sections[2]

        decimals = _count_format_decimals(fmt)
        result = None
        if fmt.strip().startswith('"') and fmt.strip().endswith('"'):
            # A section that's entirely a quoted literal — e.g. a
            # zero-section of just '"-"', a standard Excel convention for
            # hiding zeros — should show as that literal text. This is
            # deliberately narrower than 'no 0 or # present', which would
            # also (wrongly) match an unquoted, unusual format pattern
            # like a date format string on a raw number openpyxl didn't
            # auto-convert to a real datetime — showing that literal
            # format template ('yyyy-mm-dd') would be worse than falling
            # through to a plain number.
            result = fmt.strip().strip('"')
        elif "%" in fmt:
            result = f"{work_val * 100:.{decimals}f}%"
        else:
            # Whatever the format puts around the digits — a currency in
            # quotes ("GH¢"), a code ([$€-2]), a unit (" kg") — is shown as
            # the sheet shows it. Only $, £, € and ¥ used to be recognised,
            # so a cedi amount lost its "GH¢".
            prefix, core, suffix = _xl_split_number_format(fmt)
            if core:
                decimals = _count_format_decimals(core)
                digits = f"{abs(work_val):,.{decimals}f}" if "," in core else f"{abs(work_val):.{decimals}f}"
                sign = "-" if work_val < 0 else ""
                result = f"{sign}{_xl_literal_text(prefix)}{digits}{_xl_literal_text(suffix)}"
        if result is not None:
            return f"({result})" if use_parens else result
    if isinstance(val, float):
        if val == int(val):
            return str(int(val))
        return f"{val:.6f}".rstrip("0").rstrip(".")
    return str(val)


def _xlsx_with_computed_values(src_path, out_dir):
    """A spreadsheet only carries formula *results* if the program that saved
    it calculated and stored them. Excel does; files made by scripts, some
    web tools and exports often don't — and the report then showed the
    formula itself ("=ROUND(AVERAGE(C2:E2),1)") where the number belonged.
    When any result is missing, LibreOffice opens the file, calculates it,
    and saves a copy with every result stored; otherwise the file is used
    as it is."""
    try:
        wb_f = openpyxl.load_workbook(src_path, data_only=False)
        wb_v = openpyxl.load_workbook(src_path, data_only=True)
    except Exception:
        return src_path
    missing = any(
        c.data_type == "f" and wb_v[ws.title][c.coordinate].value is None
        for ws in wb_f.worksheets for row in ws.iter_rows() for c in row
    )
    if not missing:
        return src_path
    recalc_dir = os.path.join(out_dir, "recalculated")
    os.makedirs(recalc_dir, exist_ok=True)
    try:
        return _soffice_convert(src_path, "xlsx", recalc_dir)
    except Exception:
        return src_path   # still convertible — formulas are shown, as before


_ID_LIKE_HEADER = re.compile(r"\b(id|no\.?|number|num|#|phone|tel|code|year|index)\b", re.I)


def _xlsx_numeric_column_stats(rows):
    """[(header, average, lowest, highest, is_whole_numbers)] for each column
    of numbers under a text header — skipping ID-like columns (student number,
    phone, year), where an average means nothing."""
    if len(rows) < 6:   # a header and at least 5 rows — a 3-line summary needs no summary
        return []
    header = rows[0]
    stats = []
    for c in range(max(len(r) for r in rows)):
        h = header[c].value if c < len(header) and header[c] is not None else None
        if not isinstance(h, str) or not h.strip() or _ID_LIKE_HEADER.search(h):
            continue
        vals, fmt = [], None
        for row in rows[1:]:
            cell = row[c] if c < len(row) else None
            v = cell.value if cell is not None else None
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                vals.append(v)
                fmt = fmt or getattr(cell, "number_format", None)
        if len(vals) >= 3:
            whole = all(float(v).is_integer() for v in vals)
            stats.append((h.strip(), sum(vals) / len(vals), min(vals), max(vals), whole, fmt))
    return stats


def _fmt_stat(v, whole):
    if whole and float(v).is_integer():
        return f"{int(v):,}"
    return f"{v:,.1f}" if abs(v) >= 1 else f"{v:,.2f}"


# ---------------------------------------------------------------- Word ⇄ Excel helpers
_XLSX_LAYOUT_PROP = "docente.layout"


def _docx_comment_map(doc):
    """{comment id: (author, text)} for a Word document's review comments."""
    out = {}
    try:
        for part in doc.part.package.iter_parts():
            if part.partname == "/word/comments.xml":
                root = etree.fromstring(part.blob)
                for c in root.findall(qn("w:comment")):
                    text = _all_text_including_deletions(c).strip()
                    if text:
                        out[c.get(qn("w:id"))] = ((c.get(qn("w:author")) or "").strip(), text)
    except Exception:
        pass
    return out


def _docx_comment_ids(el):
    ids = []
    for tag in ("w:commentRangeStart", "w:commentReference"):
        for c in el.iter(qn(tag)):
            cid = c.get(qn("w:id"))
            if cid not in ids:
                ids.append(cid)
    return ids


def _docx_header_footer_lines(doc):
    heads, foots = [], []
    for section in doc.sections:
        for container, bucket in ((section.header, heads), (section.footer, foots)):
            try:
                for p in container.paragraphs:
                    t = _full_paragraph_text(p).strip()
                    if t and t not in bucket:
                        bucket.append(t)
            except Exception:
                pass
    return heads, foots


def _xlsx_layout(path):
    """The original document's order, if this workbook was made by Word → Excel."""
    try:
        wb = openpyxl.load_workbook(path, read_only=False)
        for prop in wb.custom_doc_props.props:
            if prop.name == _XLSX_LAYOUT_PROP:
                data = _json.loads(prop.value)
                return data if isinstance(data, dict) and data.get("v") == 1 else None
    except Exception:
        return None
    return None


def xlsx_to_docx(src_path, out_dir, style="clean"):
    """Excel → Word, faithfully: each sheet's cells as a table, numbers shown
    exactly as Excel shows them, and nothing invented — no summaries, labels
    or titles. A sheet with a name of its own (not Excel's "Sheet1") keeps it
    as a heading above its table; every table also carries its sheet's name
    invisibly (its alt-text title), so converting back names the tabs the
    same. A workbook made by Word → Excel goes back into the document's own
    order, with the document's own styles."""
    layout = _xlsx_layout(src_path)
    src_path = _xlsx_with_computed_values(src_path, out_dir)
    wb = _safe_load(openpyxl.load_workbook, src_path, data_only=True)
    wb_formulas = _safe_load(openpyxl.load_workbook, src_path, data_only=False)
    sheets, names, max_cols = {}, [], 0
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        ws_formulas = wb_formulas[sheet_name]
        merged_ranges = list(ws.merged_cells.ranges)
        rows = list(ws.iter_rows())
        formula_rows = list(ws_formulas.iter_rows())

        def has_content(row_idx, row):
            for c_idx, c in enumerate(row):
                if c.value is not None and str(c.value).strip():
                    return True
                if c_idx < len(formula_rows[row_idx]) and formula_rows[row_idx][c_idx].data_type == "f":
                    return True
            return False
        if not merged_ranges:
            keep = [i for i, r in enumerate(rows) if has_content(i, r)]
            rows = [rows[i] for i in keep]
            formula_rows = [formula_rows[i] for i in keep]
        # Columns that are empty all the way down (Excel's unused width) aren't part of the data.
        if rows and not merged_ranges:
            last = max((max((i for i, c in enumerate(r) if c.value is not None and str(c.value).strip() != ""), default=-1) for r in rows), default=-1)
            rows = [r[:last + 1] for r in rows]
            formula_rows = [r[:last + 1] for r in formula_rows]
        sheets[sheet_name] = (rows, formula_rows, merged_ranges, ws)
        names.append(sheet_name)
        if rows:
            max_cols = max(max_cols, max(len(r) for r in rows))
    doc = Document()
    if max_cols > 6:
        section = doc.sections[0]
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width
    try:
        wb_title = (wb.properties.title or "").strip()
        if wb_title:
            doc.core_properties.title = wb_title[:250]   # the file's title property, not text on the page
    except Exception:
        pass

    def add_table(sheet_name):
        rows, formula_rows, merged_ranges, ws = sheets[sheet_name]
        if not rows:
            return None
        n_cols = max(len(r) for r in rows)
        table = doc.add_table(rows=len(rows), cols=n_cols)
        table.style = "Table Grid"
        for r_idx, (doc_row, row) in enumerate(zip(table.rows, rows)):
            doc_cells = doc_row.cells
            for c_idx in range(n_cols):
                src_cell = row[c_idx] if c_idx < len(row) else None
                val = src_cell.value if src_cell is not None else None
                if val is None and c_idx < len(formula_rows[r_idx]) and formula_rows[r_idx][c_idx].data_type == "f":
                    val = formula_rows[r_idx][c_idx].value
                cell = doc_cells[c_idx]
                cell.text = _format_cell_value(val, src_cell.number_format if src_cell is not None else None)
                if src_cell is None:
                    continue
                try:
                    if src_cell.font is not None and src_cell.font.b:
                        for p in cell.paragraphs:
                            for run in p.runs:
                                run.bold = True
                except Exception:
                    pass
                fill_hex = _xlsx_conditional_fill_hex(ws, src_cell) or _xlsx_cell_fill_hex(src_cell)
                if fill_hex:
                    _set_docx_cell_shading(cell, fill_hex)
                if src_cell.comment is not None and (src_cell.comment.text or "").strip() and hasattr(doc, "add_comment"):
                    runs = cell.paragraphs[0].runs or [cell.paragraphs[0].add_run("")]
                    try:
                        doc.add_comment(runs, text=src_cell.comment.text.strip(), author=(src_cell.comment.author or "").strip() or "Excel")
                    except Exception:
                        pass
        for mr in merged_ranges:
            try:
                table.cell(mr.min_row - 1, mr.min_col - 1).merge(table.cell(mr.max_row - 1, mr.max_col - 1))
            except IndexError:
                continue
        _docx_set_table_caption(table, sheet_name)
        return table

    def add_text(text, style_name):
        try:
            p = doc.add_paragraph(style=style_name)
        except KeyError:
            p = doc.add_paragraph()
        for k, seg in enumerate(str(text).split("\n")):
            if k:
                p.add_run().add_break()
            if seg:
                p.add_run(_sanitize_xml_text(seg))
        return p

    used = set()
    if layout and layout.get("text") in sheets or (layout and not layout.get("text")):
        text_rows = []
        if layout.get("text"):
            rows = sheets[layout["text"]][0]
            text_rows = [str(r[0].value) for r in rows if r and r[0].value is not None and str(r[0].value).strip()]
        n_p = sum(1 for e in layout.get("b", []) if e and e[0] == "p")
        tables_ok = all(e[1] in sheets for e in layout.get("b", []) if e and e[0] in ("t", "c"))
        if n_p == len(text_rows) and tables_ok:
            k = 0
            for e in layout["b"]:
                if e[0] == "p":
                    add_text(text_rows[k], e[1] if len(e) > 1 else "Normal")
                    k += 1
                elif e[0] in ("t", "c"):
                    if len(e) >= 4 and e[3]:
                        add_text(e[3], e[2])
                    add_table(e[1])
                    used.add(e[1])
            if layout.get("text"):
                used.add(layout["text"])
        else:
            layout = None
    for name in names:
        if name in used or not sheets[name][0]:
            continue
        if not re.fullmatch(r"(?i)sheet\s*\d*", name.strip()):
            doc.add_heading(_sanitize_xml_text(name), level=1)   # the sheet's own name, as on its tab
        add_table(name)

    # Excel's own page header/footer → the document's page header/footer.
    first_ws = wb[wb.sheetnames[0]] if wb.sheetnames else None
    if first_ws is not None:
        for hf, container in ((first_ws.oddHeader, doc.sections[0].header), (first_ws.oddFooter, doc.sections[0].footer)):
            try:
                parts = [x.text for x in (hf.left, hf.center, hf.right) if x is not None and x.text]
            except Exception:
                parts = []
            lines = [ln for part in parts for ln in re.sub(r"&[A-Za-z]", "", part).split("\n") if ln.strip()]
            if lines:
                p = container.paragraphs[0]
                for k, ln in enumerate(lines):
                    if k:
                        p.add_run().add_break()
                    p.add_run(_sanitize_xml_text(ln))
    apply_docx_style(doc, style)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "converted.docx")
    doc.save(out_path)
    return out_path


# --------------------------------------------------------------- DOCX -> XLSX
def _xlsx_sheet_title(wb, text, fallback):
    """A valid, unique sheet-tab name taken from the document's own words
    (Excel's own "SheetN" only when there are none) — never added labels."""
    name = re.sub(r"[:\\/?*\[\]]", " ", str(text or ""))
    name = re.sub(r"\s+", " ", name).strip()
    if len(name) > 31:   # Excel's limit: cut at a whole word, never mid-word
        cut = name[:31]
        name = (cut.rsplit(" ", 1)[0] if " " in cut else cut).strip()
    name = name or fallback
    taken = {w.title.lower() for w in wb.worksheets}
    base, n = name, 2
    while name.lower() in taken:
        suffix = f" {n}"
        name = base[:31 - len(suffix)].rstrip() + suffix
        n += 1
    return name


def docx_to_xlsx(src_path, out_dir):
    doc = _safe_load(Document, src_path)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def autosize_columns(ws, n_cols):
        for col_idx in range(1, n_cols + 1):
            letter = get_column_letter(col_idx)
            longest = max(
                (len(str(cell.value)) for cell in ws[letter] if cell.value is not None),
                default=10,
            )
            ws.column_dimensions[letter].width = min(max(longest + 2, 10), 60)

    def coerce_numeric(text):
        """A table cell that reads '42' or '3.5' should become an actual
        number in the spreadsheet, not a text string that looks like one —
        otherwise SUM/sort/filter in Excel silently don't work on it. Only
        coerces values that are unambiguously numeric; anything else is
        left as text exactly as written."""
        cleaned = text.strip().replace(",", "")
        if re.fullmatch(r"-?\d+", cleaned):
            return int(cleaned)
        if re.fullmatch(r"-?\d+\.\d+", cleaned):
            return float(cleaned)
        return text

    _CUR = r"GH¢|GH₵|GHS|GHC|₵|¢|\$|€|£|¥|₦|NGN|USD|EUR|GBP"
    _NUM_RE = re.compile(
        r"(?P<s1>-)?\s*(?P<pre>" + _CUR + r")?\s*(?P<s2>-)?\s*(?P<int>\d{1,3}(?:,\d{3})+|\d+)"
        r"(?P<frac>\.\d+)?\s*(?P<post>" + _CUR + r")?")

    def coerce_value(text):
        """(value, number_format) for a table cell. Money in any common
        form (GH¢1,450.50, ₵9,800, GHS 200, $12, 1,200 USD), percentages
        (64%, 92.5%) and accounting negatives ((1,200)) become real numbers
        that still display exactly as written — they used to stay text,
        which Excel can't add up or chart and flags with a green warning on
        every cell. Anything that isn't clearly a number stays as written."""
        s = text.strip()
        m = re.fullmatch(r"(-)?\s*(\d+(?:\.\d+)?)\s*%", s)
        if m:
            frac = m.group(2).split(".")[1] if "." in m.group(2) else ""
            val = float(m.group(2)) / 100 * (-1 if m.group(1) else 1)
            return val, "0%" if not frac else "0." + "0" * len(frac) + "%"
        bracketed = bool(re.fullmatch(r"\(\s*[^()]+\s*\)", s))
        inner = s[1:-1].strip() if bracketed else s
        m = _NUM_RE.fullmatch(inner)
        if not m or (m.group("pre") and m.group("post")):
            return coerce_numeric(text), None
        symbol = m.group("pre") or m.group("post")
        if not symbol and not bracketed:
            return coerce_numeric(text), None
        number = float(m.group("int").replace(",", "") + (m.group("frac") or ""))
        if m.group("s1") or m.group("s2") or bracketed:
            number = -number
        decimals = len(m.group("frac")) - 1 if m.group("frac") else 0
        num_fmt = "#,##0" + ("." + "0" * decimals if decimals else "")
        if symbol:
            num_fmt = f'"{symbol}"{num_fmt}' if m.group("pre") else f'{num_fmt} "{symbol}"'
        if bracketed:
            num_fmt = f"{num_fmt};({num_fmt})"
        value = int(number) if not decimals and float(number).is_integer() else number
        return value, num_fmt

    # The document in its own order: each table on a tab of its own, the
    # paragraphs (in order) on one tab, and the order itself kept invisibly in
    # the workbook's properties, so Excel → Word puts it all back. A heading
    # straight above a table names that table's tab instead of being repeated.
    blocks = list(_iter_block_items(doc))
    comment_map = _docx_comment_map(doc)
    layout = []
    text_rows = []          # (text, comment ids)
    chart_data_list = []
    for bi, block in enumerate(blocks):
        if isinstance(block, DocxTable):
            table = block
            heading = None
            prev = blocks[bi - 1] if bi else None
            if (prev is not None and not isinstance(prev, DocxTable) and layout and layout[-1][0] == "p" and text_rows
                    and re.match(r"(heading \d+|title)$", (prev.style.name or "").lower())):
                heading = (prev.style.name, text_rows.pop()[0])
                layout.pop()
            first_cell = next((c.text.strip() for row in table.rows for c in row.cells if c.text.strip()), "")
            name_src = _docx_table_caption(table) or (heading[1] if heading else first_cell)
            ws = wb.create_sheet(title=_xlsx_sheet_title(wb, name_src, f"Sheet{len(wb.worksheets) + 1}"))
            src_rows = [list(row.cells) for row in table.rows]
            n_cols = len(table.columns)
            merges = _find_docx_merges(src_rows, n_cols)
            skip_cells = {(r, c) for (min_r, min_c, max_r, max_c) in merges
                          for r in range(min_r, max_r + 1) for c in range(min_c, max_c + 1)
                          if (r, c) != (min_r, min_c)}
            for r_idx, row in enumerate(table.rows, 1):
                for c_idx, cell in enumerate(row.cells, 1):
                    if (r_idx - 1, c_idx - 1) in skip_cells:
                        continue  # covered by a merge origin elsewhere — left blank, filled via the merge below
                    text = _full_cell_text(cell).strip()
                    value, num_fmt = (text, None) if r_idx == 1 else coerce_value(text)
                    xlsx_cell = ws.cell(row=r_idx, column=c_idx, value=value)
                    if num_fmt:
                        xlsx_cell.number_format = num_fmt
                    for cid in _docx_comment_ids(cell._tc):
                        if cid in comment_map:
                            author, ctext = comment_map[cid]
                            xlsx_cell.comment = XlsxComment(ctext, author or "Word")
                    shade = _get_docx_cell_shading(cell)
                    if shade:
                        xlsx_cell.fill = XlsxPatternFill(start_color=shade, end_color=shade, fill_type="solid")
            for min_r, min_c, max_r, max_c in merges:
                try:
                    ws.merge_cells(start_row=min_r + 1, start_column=min_c + 1, end_row=max_r + 1, end_column=max_c + 1)
                except Exception:
                    pass
            for cell in ws[1]:
                cell.font = XlsxFont(bold=True)
            autosize_columns(ws, len(table.columns))
            entry = ["t", ws.title]
            if heading:
                entry += [heading[0], heading[1]]
            layout.append(entry)
            continue
        p = block
        full_text = _full_paragraph_text(p).strip()
        if full_text:
            text_rows.append((full_text, _docx_comment_ids(p._p)))
            layout.append(["p", p.style.name or "Normal"])
        for tb_para in _iter_textbox_paragraphs(p):
            tb_text = _full_paragraph_text(tb_para).strip()
            if tb_text:
                text_rows.append((tb_text, []))
                layout.append(["p", "Normal"])
        for chart_title, chart_cats, chart_series, _chart_type in _iter_docx_charts(p, doc):
            chart_data_list.append((chart_title, chart_cats, chart_series))
            layout.append(["c", len(chart_data_list) - 1])
    text_sheet = None
    if text_rows:
        ws = wb.create_sheet(title=_xlsx_sheet_title(wb, text_rows[0][0], f"Sheet{len(wb.worksheets) + 1}"))
        ws.column_dimensions["A"].width = 100
        for r, (line, cids) in enumerate(text_rows, 1):
            c = ws.cell(row=r, column=1, value=line)
            if "\n" in line:
                c.alignment = XlsxAlignment(wrap_text=True)
            for cid in cids:
                if cid in comment_map:
                    author, ctext = comment_map[cid]
                    c.comment = XlsxComment(ctext, author or "Word")
        text_sheet = ws.title
    chart_sheet_names = []
    for i, (chart_title, chart_cats, chart_series) in enumerate(chart_data_list, 1):
        ws = wb.create_sheet(title=_xlsx_sheet_title(wb, chart_title or (chart_cats[0] if chart_cats else ""), f"Sheet{len(wb.worksheets) + 1}"))
        chart_sheet_names.append(ws.title)
        ws.cell(row=1, column=1, value="")
        for s_idx, (name, _values) in enumerate(chart_series, 2):
            ws.cell(row=1, column=s_idx, value=name).font = XlsxFont(bold=True)
        for cat_idx, cat_name in enumerate(chart_cats, 2):
            ws.cell(row=cat_idx, column=1, value=cat_name)
            for s_idx, (_name, values) in enumerate(chart_series, 2):
                val = values[cat_idx - 2] if cat_idx - 2 < len(values) else None
                try:
                    val = float(val)
                except (ValueError, TypeError):
                    pass
                ws.cell(row=cat_idx, column=s_idx, value=val)
        autosize_columns(ws, len(chart_series) + 1)

    footnotes = _get_docx_footnotes(doc)
    if footnotes:
        ws = wb.create_sheet(title=_xlsx_sheet_title(wb, footnotes[0], f"Sheet{len(wb.worksheets) + 1}"))
        ws.column_dimensions["A"].width = 100
        for r, note in enumerate(footnotes, 1):
            ws.cell(row=r, column=1, value=f"{r}. {note}")

    endnotes = _get_docx_endnotes(doc)
    if endnotes:
        ws = wb.create_sheet(title=_xlsx_sheet_title(wb, endnotes[0], f"Sheet{len(wb.worksheets) + 1}"))
        ws.column_dimensions["A"].width = 100
        for r, note in enumerate(endnotes, 1):
            ws.cell(row=r, column=1, value=f"{r}. {note}")

    heads, foots = _docx_header_footer_lines(doc)
    for ws in wb.worksheets:
        if heads:
            ws.oddHeader.center.text = "\n".join(heads)[:250]
        if foots:
            ws.oddFooter.center.text = "\n".join(foots)[:250]
    for e in layout:
        if e[0] == "c":
            e[1] = chart_sheet_names[e[1]] if e[1] < len(chart_sheet_names) else ""
    try:
        wb.custom_doc_props.append(XlsxStringProperty(name=_XLSX_LAYOUT_PROP, value=_json.dumps({"v": 1, "text": text_sheet, "b": layout}, ensure_ascii=False)))
    except Exception:
        pass
    if not wb.sheetnames:
        wb.create_sheet(title="Sheet1")   # Excel's own default; nothing written into it

    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "converted.xlsx")
    wb.save(out_path)
    return out_path



# ======================================================================== 3.1: the PDF editor
import logging
_PDFED_MAX_PAGES = 500
_PDFED_MAX_ITEMS = 4000
_PDFED_OCR_PAGES = 12
_PDFED_RASTER_PAGES = 40
_PDFED_FONT_FILES = {
    # regular, bold, italic, bold italic — metric twins of the fonts most documents use
    "calibri": ("Carlito-Regular.ttf", "Carlito-Bold.ttf", "Carlito-Italic.ttf", "Carlito-BoldItalic.ttf"),
    "cambria": ("Caladea-Regular.ttf", "Caladea-Bold.ttf", "Caladea-Italic.ttf", "Caladea-BoldItalic.ttf"),
    "sans": ("LiberationSans-Regular.ttf", "LiberationSans-Bold.ttf", "LiberationSans-Italic.ttf", "LiberationSans-BoldItalic.ttf"),
    "serif": ("LiberationSerif-Regular.ttf", "LiberationSerif-Bold.ttf", "LiberationSerif-Italic.ttf", "LiberationSerif-BoldItalic.ttf"),
    "mono": ("LiberationMono-Regular.ttf", "LiberationMono-Bold.ttf", "LiberationMono-Italic.ttf", "LiberationMono-BoldItalic.ttf"),
    "dejavu": ("DejaVuSans.ttf", "DejaVuSans-Bold.ttf", "DejaVuSans-Oblique.ttf", "DejaVuSans-BoldOblique.ttf"),
}
_PDFED_FONT_DIRS = ("/usr/share/fonts/truetype/crosextra", "/usr/share/fonts/truetype/liberation",
                    "/usr/share/fonts/truetype/liberation2", "/usr/share/fonts/truetype/dejavu")
_PDFED_BUILTIN = {"sans": ("Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique"),
                  "serif": ("Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic"),
                  "mono": ("Courier", "Courier-Bold", "Courier-Oblique", "Courier-BoldOblique")}
_pdfed_fonts = {}
_pdfed_font_paths = {}


def _pdfed_num(v, lo=-1e6, hi=1e6, default=0.0):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return default
    if f != f or f in (float("inf"), float("-inf")):
        return default
    return max(lo, min(hi, f))


def _pdfed_rect(r):
    """[x0, y0, x1, y1] in page points, in order; None unless four numbers."""
    if not isinstance(r, (list, tuple)) or len(r) != 4:
        return None
    a = [_pdfed_num(v) for v in r]
    x0, x1 = sorted((a[0], a[2]))
    y0, y1 = sorted((a[1], a[3]))
    return [x0, y0, x1, y1]


def _pdfed_rgb(h, default=(0, 0, 0)):
    h = str(h or "").strip().lstrip("#")
    if re.fullmatch(r"[0-9A-Fa-f]{6}", h):
        return tuple(int(h[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    return default


def _pdfed_family(name):
    """The kind of font a PDF font name is (or a hint the editor sends): calibri, cambria, serif, mono, sans."""
    n = str(name or "").lower().split("+", 1)[-1]
    if "calibri" in n or "carlito" in n:
        return "calibri"
    if "cambria" in n or "caladea" in n:
        return "cambria"
    if any(k in n for k in ("courier", "mono", "consol", "menlo", "code")):
        return "mono"
    if "sans" not in n and any(k in n for k in ("serif", "times", "roman", "georgia", "garamond", "antiqua", "palatino", "minion", "book")):
        return "serif"
    return "sans"


def _pdfed_font_file(name):
    if name in _pdfed_font_paths:
        return _pdfed_font_paths[name]
    found = None
    for d in _PDFED_FONT_DIRS:
        p = os.path.join(d, name)
        if os.path.exists(p):
            found = p
            break
    if found is None:
        for root in ("/usr/share/fonts", "/usr/local/share/fonts"):
            for dirpath, _dirs, files in os.walk(root):
                if name in files:
                    found = os.path.join(dirpath, name)
                    break
            if found:
                break
    _pdfed_font_paths[name] = found
    return found


def _pdfed_ttf(family, style):
    """(reportlab font name, TTFont) for a family and style (0 regular, 1 bold, 2 italic, 3 both); None if missing."""
    key = (family, style)
    if key in _pdfed_fonts:
        return _pdfed_fonts[key]
    result = None
    files = _PDFED_FONT_FILES.get(family)
    if files:
        path = _pdfed_font_file(files[style]) or _pdfed_font_file(files[0])
        if path:
            try:
                from reportlab.pdfbase import pdfmetrics
                from reportlab.pdfbase.ttfonts import TTFont
                name = "Ded-%s-%d" % (family, style)
                font = TTFont(name, path)
                pdfmetrics.registerFont(font)
                result = (name, font)
            except Exception:
                result = None
    _pdfed_fonts[key] = result
    return result


def _pdfed_font(family, text, bold=False, italic=False):
    """The font to write `text` in: the family asked for when it has every character, else Liberation
    Sans, else DejaVu (which has almost every alphabet), else the PDF built-ins."""
    style = (1 if bold else 0) + (2 if italic else 0)
    chars = {ord(ch) for ch in str(text) if ch not in " \t\r\n"}
    for fam in dict.fromkeys((family, "sans" if family in ("calibri",) else family, "dejavu")):
        f = _pdfed_ttf(fam, style)
        if f and chars <= set(f[1].face.charToGlyph):
            return f[0]
    base = {"cambria": "serif", "serif": "serif", "mono": "mono"}.get(family, "sans")
    return _PDFED_BUILTIN[base][style]


def _pdfed_wrap(text, fontname, size, width):
    from reportlab.pdfbase.pdfmetrics import stringWidth
    out = []
    for para in str(text).replace("\r\n", "\n").split("\n"):
        line = ""
        for word in para.split(" "):
            trial = (line + " " + word) if line else word
            if line and stringWidth(trial, fontname, size) > width:
                out.append(line)
                line = word
            else:
                line = trial
        out.append(line)
    return out


def _pdfed_draw(c, it, assets, cover_rects):
    """Draws one item onto the page's overlay canvas."""
    from reportlab.lib.utils import ImageReader
    t = it.get("type")
    r = _pdfed_rect(it.get("rect"))
    color = _pdfed_rgb(it.get("color"), (0, 0, 0))
    if t in ("whiteout", "edit") and r:
        c.saveState()
        c.setFillColorRGB(*_pdfed_rgb(it.get("bg"), (1, 1, 1)))
        pad = _pdfed_num(it.get("pad"), 0, 6, 0.8)
        c.rect(r[0] - pad, r[1] - pad, (r[2] - r[0]) + 2 * pad, (r[3] - r[1]) + 2 * pad, stroke=0, fill=1)
        c.restoreState()
    if t == "edit" and r:
        text = str(it.get("text") or "")[:4000]
        if text.strip():
            size = _pdfed_num(it.get("size"), 2, 300, 11)
            fam = _pdfed_family(it.get("font"))
            fname = _pdfed_font(fam, text, it.get("bold"), it.get("italic"))
            x = _pdfed_num(it.get("x"), -1e5, 1e5, r[0])
            base = _pdfed_num(it.get("baseline"), -1e5, 1e5, r[1] + (r[3] - r[1]) * 0.22)
            c.setFont(fname, size)
            c.setFillColorRGB(*color)
            c.drawString(x, base, text)
    elif t == "text" and r:
        text = str(it.get("text") or "")[:8000]
        if text.strip():
            size = _pdfed_num(it.get("size"), 2, 300, 14)
            fam = _pdfed_family(it.get("font") or "sans")
            fname = _pdfed_font(fam, text, it.get("bold"), it.get("italic"))
            width = max(r[2] - r[0], size)
            lines = _pdfed_wrap(text, fname, size, width)
            leading = size * 1.2
            y = r[3] - size * 0.95
            align = it.get("align") or "left"
            c.setFont(fname, size)
            c.setFillColorRGB(*color)
            for line in lines:
                if align == "center":
                    c.drawCentredString(r[0] + width / 2, y, line)
                elif align == "right":
                    c.drawRightString(r[0] + width, y, line)
                else:
                    c.drawString(r[0], y, line)
                y -= leading
    elif t in ("image", "replace_image") and r:
        path = assets.get(str(it.get("file") or ""))
        if path and os.path.exists(path):
            try:
                c.drawImage(ImageReader(path), r[0], r[1], r[2] - r[0], r[3] - r[1], mask="auto")
            except Exception:
                pass
    elif t == "highlight":
        c.saveState()
        try:
            c.setBlendMode("Multiply")
        except Exception:
            c.setFillAlpha(0.4)
        c.setFillColorRGB(*_pdfed_rgb(it.get("color"), (1.0, 0.89, 0.30)))
        for q in (it.get("rects") or ([r] if r else []))[:500]:
            q = _pdfed_rect(q)
            if q:
                c.rect(q[0], q[1], q[2] - q[0], q[3] - q[1], stroke=0, fill=1)
        c.restoreState()
    elif t in ("underline", "strike"):
        c.saveState()
        c.setStrokeColorRGB(*color)
        for q in (it.get("rects") or ([r] if r else []))[:500]:
            q = _pdfed_rect(q)
            if not q:
                continue
            h = q[3] - q[1]
            c.setLineWidth(max(0.6, h * 0.07))
            y = q[1] + h * (0.12 if t == "underline" else 0.42)
            c.line(q[0], y, q[2], y)
        c.restoreState()
    elif t == "ink":
        c.saveState()
        c.setStrokeColorRGB(*color)
        c.setLineWidth(_pdfed_num(it.get("width"), 0.3, 40, 2))
        c.setLineCap(1)
        c.setLineJoin(1)
        for path in (it.get("paths") or [])[:300]:
            pts = [(_pdfed_num(p[0]), _pdfed_num(p[1])) for p in (path or [])[:4000] if isinstance(p, (list, tuple)) and len(p) == 2]
            if len(pts) == 1:
                pts.append((pts[0][0] + 0.01, pts[0][1]))
            if len(pts) < 2:
                continue
            p = c.beginPath()
            p.moveTo(*pts[0])
            for q in pts[1:]:
                p.lineTo(*q)
            c.drawPath(p, stroke=1, fill=0)
        c.restoreState()
    elif t in ("check", "cross", "dot") and r:
        w, h = r[2] - r[0], r[3] - r[1]
        c.saveState()
        c.setStrokeColorRGB(*color)
        c.setFillColorRGB(*color)
        c.setLineWidth(max(1.0, min(w, h) * 0.12))
        c.setLineCap(1)
        c.setLineJoin(1)
        if t == "check":
            p = c.beginPath()
            p.moveTo(r[0] + w * 0.12, r[1] + h * 0.52)
            p.lineTo(r[0] + w * 0.40, r[1] + h * 0.20)
            p.lineTo(r[0] + w * 0.88, r[1] + h * 0.84)
            c.drawPath(p, stroke=1, fill=0)
        elif t == "cross":
            c.line(r[0] + w * 0.18, r[1] + h * 0.18, r[2] - w * 0.18, r[3] - h * 0.18)
            c.line(r[0] + w * 0.18, r[3] - h * 0.18, r[2] - w * 0.18, r[1] + h * 0.18)
        else:
            c.circle(r[0] + w / 2, r[1] + h / 2, min(w, h) * 0.22, stroke=0, fill=1)
        c.restoreState()
    elif t == "redact" and r:
        c.saveState()
        c.setFillColorRGB(0, 0, 0)
        c.rect(r[0], r[1], r[2] - r[0], r[3] - r[1], stroke=0, fill=1)
        c.restoreState()
    for q in cover_rects:          # pictures that couldn't be taken out: painted over
        c.saveState()
        c.setFillColorRGB(1, 1, 1)
        c.rect(q[0], q[1], q[2] - q[0], q[3] - q[1], stroke=0, fill=1)
        c.restoreState()


_PDFED_TURNABLE = ("text", "image", "replace_image", "check", "cross", "dot")


def _pdfed_frame(c, rect, rot):
    """Turns the canvas so that (0, 0)–(w, h) is `rect` as it's seen on a page turned by `rot`
    (0, 90, 180 or 270, clockwise, as /Rotate turns it); returns (w, h)."""
    x0, y0, x1, y1 = rect
    c.translate(*{0: (x0, y0), 90: (x1, y0), 180: (x1, y1), 270: (x0, y1)}[rot])
    if rot:
        c.rotate(rot)
    return (x1 - x0, y1 - y0) if rot in (0, 180) else (y1 - y0, x1 - x0)


def _pdfed_overlay(wpage, items, cover, assets, k, total, numbers, watermark, view_rot=0, box=None):
    """One page's drawing, as a one-page PDF to lay over it; None when there's nothing to draw.
    Items added while the page was seen turned carry that turn ("rot") and are drawn upright as seen;
    page numbers and the watermark go where they're seen on the page as it will be shown (view_rot,
    and `box`: the page's new crop, if it has one)."""
    drawn = [it for it in items if it.get("type") not in ("note", "remove_image")]
    if not drawn and not cover and not numbers and not watermark:
        return None
    from reportlab.pdfgen import canvas as rl_canvas
    mb = wpage.mediabox
    cb = wpage.cropbox
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(max(float(mb.right), 1.0), max(float(mb.top), 1.0)))
    if cover:          # pictures that couldn't be taken out are painted over first, under anything new
        _pdfed_draw(c, {"type": "none"}, assets, cover)
    for it in drawn:
        rot = int(_pdfed_num(it.get("rot"), 0, 359, 0)) // 90 * 90
        r = _pdfed_rect(it.get("rect"))
        if rot and r and it.get("type") in _PDFED_TURNABLE:
            c.saveState()
            w, h = _pdfed_frame(c, r, rot)
            _pdfed_draw(c, dict(it, rect=[0, 0, w, h]), assets, [])
            c.restoreState()
        else:
            _pdfed_draw(c, it, assets, [])
    frame = box or [float(cb.left), float(cb.bottom), float(cb.right), float(cb.top)]
    vr = int(view_rot or 0) // 90 * 90 % 360
    if watermark:
        text = str(watermark.get("text") or "")[:120]
        if text.strip():
            size = _pdfed_num(watermark.get("size"), 8, 200, 60)
            fname = _pdfed_font("sans", text, True, False)
            c.saveState()
            W, H = _pdfed_frame(c, frame, vr)
            c.setFillColorRGB(*_pdfed_rgb(watermark.get("color"), (0.5, 0.5, 0.5)))
            c.setFillAlpha(_pdfed_num(watermark.get("opacity"), 0.05, 1, 0.18))
            c.translate(W / 2, H / 2)
            c.rotate(_pdfed_num(watermark.get("angle"), -180, 180, 45))
            from reportlab.pdfbase.pdfmetrics import stringWidth
            long_side = 0.8 * (W * W + H * H) ** 0.5          # a long watermark is made to fit across the page
            wide = stringWidth(text, fname, size)
            if wide > long_side:
                size = max(8, size * long_side / wide)
            c.setFont(fname, size)
            c.drawCentredString(0, -size / 3, text)
            c.restoreState()
    if numbers:
        n = k + int(_pdfed_num(numbers.get("start"), -9999, 99999, 1))
        fmt = str(numbers.get("format") or "{n}")[:60]
        label = fmt.replace("{n}", str(n)).replace("{total}", str(total + int(_pdfed_num(numbers.get("start"), -9999, 99999, 1)) - 1))
        size = _pdfed_num(numbers.get("size"), 6, 40, 11)
        pos = str(numbers.get("position") or "bottom-center")
        fname = _pdfed_font("sans", label)
        c.saveState()
        W, H = _pdfed_frame(c, frame, vr)
        y = 24 if pos.startswith("bottom") else (H - 24 - size * 0.7)
        c.setFillColorRGB(0, 0, 0)
        c.setFont(fname, size)
        if pos.endswith("left"):
            c.drawString(36, y, label)
        elif pos.endswith("right"):
            c.drawRightString(W - 36, y, label)
        else:
            c.drawCentredString(W / 2, y, label)
        c.restoreState()
    c.showPage()
    c.save()
    buf.seek(0)
    return pypdf.PdfReader(buf).pages[0]


def _pdfed_mult(m, n):
    a, b, c_, d, e, f = m
    a2, b2, c2, d2, e2, f2 = n
    return [a * a2 + b * c2, a * b2 + b * d2, c_ * a2 + d * c2, c_ * b2 + d * d2, e * a2 + f * c2 + e2, e * b2 + f * d2 + f2]


def _pdfed_box_of(ctm):
    pts = [(0, 0), (1, 0), (0, 1), (1, 1)]
    xs = [ctm[0] * x + ctm[2] * y + ctm[4] for x, y in pts]
    ys = [ctm[1] * x + ctm[3] * y + ctm[5] for x, y in pts]
    return [min(xs), min(ys), max(xs), max(ys)]


def _pdfed_overlap(a, b):
    ix = max(0.0, min(a[2], b[2]) - max(a[0], b[0]))
    iy = max(0.0, min(a[3], b[3]) - max(a[1], b[1]))
    inter = ix * iy
    union = (a[2] - a[0]) * (a[3] - a[1]) + (b[2] - b[0]) * (b[3] - b[1]) - inter
    return inter / union if union > 0 else 0.0


def _pdfed_clean(wpage, image_rects, text_rects, writer):
    """Takes things out of the page's own content, so they're really gone (not just covered):
      * pictures: the drawing of each picture found in one of image_rects;
      * words: every piece of text that starts inside one of text_rects (the lines that were edited
        — their old words would otherwise still be found by search and copy).
    Text is followed through the page's text positioning (Tm, Td, TD, T*, TL, ' and "), which is all a
    start point needs. Returns the image boxes where no picture was found (to be painted over)."""
    image_rects = [r for r in image_rects if r]
    text_rects = [[r[0] - 1.5, r[1] - 2.0, r[2] + 1.5, r[3] + 2.0] for r in text_rects if r]
    if not image_rects and not text_rects:
        return []
    try:
        from pypdf.generic import ContentStream
        contents = wpage.get_contents()
        if contents is None:
            return image_rects
        cs = ContentStream(contents, writer)
        res = wpage.get("/Resources")
        res = res.get_object() if res is not None else {}
        xobjs = res.get("/XObject") if hasattr(res, "get") else None
        xobjs = xobjs.get_object() if xobjs is not None else {}
        ident = [1, 0, 0, 1, 0, 0]
        ctm, stack, found, keep, changed, dropped = ident[:], [], set(), [], False, set()
        tm, tlm, leading, rise = ident[:], ident[:], 0.0, 0.0

        def text_start():
            m = _pdfed_mult(tm, ctm)
            return m[2] * rise + m[4], m[3] * rise + m[5]

        def in_text_rects(pt):
            return any(r[0] <= pt[0] <= r[2] and r[1] <= pt[1] <= r[3] for r in text_rects)
        for operands, op in cs.operations:
            if op == b"q":
                stack.append(ctm[:])
            elif op == b"Q":
                ctm = stack.pop() if stack else ident[:]
            elif op == b"cm" and len(operands) == 6:
                ctm = _pdfed_mult([float(v) for v in operands], ctm)
            elif op == b"BT":
                tm, tlm = ident[:], ident[:]
            elif op == b"Tm" and len(operands) == 6:
                tm = [float(v) for v in operands]
                tlm = tm[:]
            elif op in (b"Td", b"TD") and len(operands) == 2:
                tx, ty = float(operands[0]), float(operands[1])
                if op == b"TD":
                    leading = -ty
                tlm = _pdfed_mult([1, 0, 0, 1, tx, ty], tlm)
                tm = tlm[:]
            elif op == b"TL" and operands:
                leading = float(operands[0])
            elif op == b"Ts" and operands:
                rise = float(operands[0])
            elif op in (b"T*", b"'", b'"'):
                tlm = _pdfed_mult([1, 0, 0, 1, 0, -leading], tlm)
                tm = tlm[:]
            if text_rects and op in (b"Tj", b"TJ", b"'", b'"') and in_text_rects(text_start()):
                changed = True
                if op in (b"'", b'"'):
                    keep.append(([], b"T*"))      # the move to the next line stays; only the words go
                continue
            if image_rects and op in (b"Do", b"INLINE IMAGE"):
                is_image = op == b"INLINE IMAGE"
                if op == b"Do" and operands:
                    xo = xobjs.get(operands[0]) if hasattr(xobjs, "get") else None
                    xo = xo.get_object() if xo is not None else None
                    is_image = bool(xo is not None and xo.get("/Subtype") == "/Image")
                if is_image:
                    box = _pdfed_box_of(ctm)
                    hit = next((i for i, r in enumerate(image_rects) if i not in found and _pdfed_overlap(box, r) > 0.6), None)
                    if hit is not None:
                        found.add(hit)
                        changed = True
                        if op == b"Do":
                            dropped.add(operands[0])
                        continue
            keep.append((operands, op))
        if changed:
            cs.operations = keep
            wpage.replace_contents(cs)
        # a picture taken out leaves the file too (not only the page's drawing): it's dropped from this
        # page's own copy of its resources — other pages sharing them keep theirs — and, no longer used
        # anywhere, it isn't written (see pdf_apply_plan)
        unused = dropped - {ops[0] for ops, o in keep if o == b"Do" and ops}
        if unused:
            from pypdf.generic import DictionaryObject, NameObject
            new_res = DictionaryObject(dict(res.items()))
            new_res[NameObject("/XObject")] = DictionaryObject({k: v for k, v in xobjs.items() if k not in unused})
            wpage[NameObject("/Resources")] = new_res
            writer._pdfed_dropped = True
        return [r for i, r in enumerate(image_rects) if i not in found]
    except Exception:
        logging.getLogger(__name__).warning("PDF editor: couldn't take something out of a page; painting over it", exc_info=True)
        return image_rects


def _pdfed_note(writer, k, it):
    from pypdf.annotations import Text
    from pypdf.generic import (ArrayObject, DecodedStreamObject, DictionaryObject, FloatObject, NameObject,
                               NumberObject, TextStringObject)
    r = _pdfed_rect(it.get("rect"))
    if not r:
        return
    x, top = r[0], r[3]
    rgb = _pdfed_rgb(it.get("color"), (1.0, 0.84, 0.25))
    annot = Text(rect=(x, top - 22, x + 22, top), text=str(it.get("text") or "")[:8000], open=False)
    annot[NameObject("/C")] = ArrayObject([FloatObject(v) for v in rgb])
    annot[NameObject("/T")] = TextStringObject(str(it.get("author") or "Note")[:80])
    annot[NameObject("/Name")] = NameObject("/Comment")
    annot[NameObject("/F")] = NumberObject(4)
    ap = DecodedStreamObject()
    ap.set_data(("q %.3f %.3f %.3f rg 0.15 0.15 0.15 RG 1 w 1 1 20 20 re B "
                 "0.15 0.15 0.15 RG 1.2 w 5 15 m 17 15 l S 5 11 m 17 11 l S 5 7 m 13 7 l S Q" % rgb).encode("ascii"))
    ap.update({NameObject("/Type"): NameObject("/XObject"), NameObject("/Subtype"): NameObject("/Form"),
               NameObject("/BBox"): ArrayObject([FloatObject(0), FloatObject(0), FloatObject(22), FloatObject(22)])})
    annot[NameObject("/AP")] = DictionaryObject({NameObject("/N"): writer._add_object(ap)})
    writer.add_annotation(k, annot)


def _pdfed_field_name(annot):
    parts, node, guard = [], annot, 0
    while node is not None and guard < 20:
        t = node.get("/T")
        if t is not None:
            parts.append(str(t))
        parent = node.get("/Parent")
        node = parent.get_object() if parent is not None else None
        guard += 1
    return ".".join(reversed(parts))


def _pdfed_field_attr(annot, key):
    node, guard = annot, 0
    while node is not None and guard < 20:
        if key in node:
            return node[key]
        parent = node.get("/Parent")
        node = parent.get_object() if parent is not None else None
        guard += 1
    return None


def _pdfed_fill_form(writer, answers, flatten=False):
    """Writes the answers into the PDF's own form fields (matched by full name or short name): text,
    ticks (their own "on" state) and choices."""
    if not answers:
        return
    for page in writer.pages:
        annots = page.get("/Annots")
        annots = annots.get_object() if annots is not None else []
        values = {}
        for ref in annots or []:
            a = ref.get_object()
            if a.get("/Subtype") != "/Widget":
                continue
            full = _pdfed_field_name(a)
            short = full.rsplit(".", 1)[-1]
            key = full if full in answers else (short if short in answers else None)
            if key is None:
                continue
            v = answers[key]
            ft = _pdfed_field_attr(a, "/FT")
            if ft == "/Btn":
                ap = a.get("/AP")
                ap = ap.get_object() if ap is not None else {}
                states = [str(s) for s in ((ap.get("/N").get_object().keys()) if ap.get("/N") is not None else []) if str(s) != "/Off"]
                if isinstance(v, bool) or v in ("true", "false", "on", "off", "", None):
                    on = v in (True, "true", "on")
                    val = (states[0] if states else "/Yes") if on else "/Off"
                else:
                    want = "/" + str(v).lstrip("/")
                    val = want if want in states else "/Off"
                if val == "/Off" and values.get(short, "/Off") != "/Off":
                    continue        # a group of choices: the chosen button's state wins over the others' "off"
                values[short] = val
            else:
                values[short] = str(v)[:2000]
        if values:
            try:
                writer.update_page_form_field_values(page, values, auto_regenerate=False, flatten=flatten)
            except Exception:
                logging.getLogger(__name__).warning("PDF editor: a form answer couldn't be written", exc_info=True)
    try:
        writer.set_need_appearances_writer(True)
    except Exception:
        pass


def _pdfed_render(pdf_path, k, dpi, unrotated=True):
    """Page k as a PIL picture (RGB) of its visible area; with unrotated=True as stored, not as turned."""
    import pypdfium2 as pdfium
    doc = pdfium.PdfDocument(pdf_path)
    try:
        page = doc[k]
        rot = page.get_rotation()
        if unrotated and rot:
            page.set_rotation(0)
        img = page.render(scale=dpi / 72.0, may_draw_forms=True).to_pil().convert("RGB")
        return img, rot
    finally:
        doc.close()


def _pdfed_ocr_layer(img, dpi, work_dir, tag):
    """An invisible, searchable text layer for a page picture (Tesseract); None when it can't be made."""
    png = os.path.join(work_dir, "ocr_%s.png" % tag)
    img.save(png, dpi=(dpi, dpi))
    base = os.path.join(work_dir, "ocr_%s" % tag)
    try:
        subprocess.run(["tesseract", png, base, "-l", "eng", "--dpi", str(dpi), "-c", "textonly_pdf=1", "pdf"],
                       capture_output=True, timeout=90, check=True)
        return pypdf.PdfReader(base + ".pdf").pages[0]
    except Exception:
        return None


def _pdfed_lay_text(page, layer, box):
    """Lays a text layer (made from a picture of the box) over the page, stretched to the box."""
    from pypdf import Transformation
    lw, lh = float(layer.mediabox.width), float(layer.mediabox.height)
    if lw <= 0 or lh <= 0:
        return
    t = Transformation().scale(box[2] / lw, box[3] / lh).translate(box[0], box[1])
    page.merge_transformed_page(layer, t)


def _pdfed_rasterize(path, pages, work_dir, dpi=200, ocr=True):
    """Turns these pages into pictures of themselves (redaction: nothing under a black box survives),
    with their visible words made searchable again. Returns the new file's path."""
    from reportlab.pdfgen import canvas as rl_canvas
    reader = pypdf.PdfReader(path)
    writer = pypdf.PdfWriter(clone_from=reader)
    for n, k in enumerate(sorted(set(pages))[:_PDFED_RASTER_PAGES]):
        src = writer.pages[k]
        cb = src.cropbox
        x0, y0, w, h = float(cb.left), float(cb.bottom), float(cb.width), float(cb.height)
        img, rot = _pdfed_render(path, k, dpi)
        jpg = os.path.join(work_dir, "raster_%d.jpg" % k)
        img.save(jpg, "JPEG", quality=90)
        buf = io.BytesIO()
        c = rl_canvas.Canvas(buf, pagesize=(w, h))
        c.drawImage(jpg, 0, 0, w, h)
        c.showPage()
        c.save()
        buf.seek(0)
        newp = pypdf.PdfReader(buf).pages[0]
        if ocr and n < _PDFED_OCR_PAGES:
            layer = _pdfed_ocr_layer(img, dpi, work_dir, "r%d" % k)
            if layer is not None:
                _pdfed_lay_text(newp, layer, (0, 0, w, h))
        added = writer.insert_page(newp, k)
        if rot:
            added.rotate(rot)
        writer.remove_page(k + 1)
    out = os.path.join(work_dir, "rasterized.pdf")
    with open(out, "wb") as f:
        writer.write(f)
    return out


def pdf_has_text(path, k):
    try:
        with pdfplumber.open(path) as pdf:
            return len((pdf.pages[k].extract_text() or "").strip()) >= 12
    except Exception:
        return True


def _pdfed_ocr(writer, path, work_dir, dpi=200):
    """Makes scanned pages searchable: pages with no words of their own get Tesseract's invisible
    text layer over them (the page itself isn't changed). Returns how many pages were done."""
    done = 0
    for k, page in enumerate(writer.pages):
        if done >= _PDFED_OCR_PAGES:
            break
        if pdf_has_text(path, k):
            continue
        img, _rot = _pdfed_render(path, k, dpi)
        layer = _pdfed_ocr_layer(img, dpi, work_dir, "o%d" % k)
        if layer is None:
            continue
        cb = page.cropbox
        _pdfed_lay_text(page, layer, (float(cb.left), float(cb.bottom), float(cb.width), float(cb.height)))
        done += 1
    return done


def pdf_plan_needs_pro(plan):
    """Whether this plan uses a Pro tool: redaction, or a password / limits on what readers may do."""
    if not isinstance(plan, dict):
        return False
    if any(isinstance(it, dict) and it.get("type") == "redact" for it in (plan.get("items") or [])):
        return True
    out = plan.get("output") or {}
    perms = out.get("permissions") or {}
    return bool(out.get("password")) or any(v is False for v in perms.values()) if isinstance(perms, dict) else bool(out.get("password"))


def pdf_apply_plan(src_path, plan, out_dir, assets=None, password=None):
    """The PDF editor's Save: the original PDF plus the editor's plan → the edited PDF.
    Returns (path, notes) — notes says what couldn't be done in full (e.g. OCR page limit)."""
    from pypdf.constants import UserAccessPermissions as UAP
    assets = assets or {}
    notes = []
    if not isinstance(plan, dict):
        raise PdfEditError("The changes couldn't be read — please try again.")
    os.makedirs(out_dir, exist_ok=True)

    sources = [src_path] + [assets.get(str(n)) for n in (plan.get("sources") or [])[:20]]
    readers, src_encrypted = [], False
    for k, p in enumerate(sources):
        if not p or not os.path.exists(p):
            raise PdfEditError("A PDF you added is missing — please add it again.")
        r = _safe_load(pypdf.PdfReader, p)
        if r.is_encrypted:
            ok = False
            for pw in ([password] if (k == 0 and password) else []) + [""]:
                try:
                    if r.decrypt(pw):
                        ok = True
                        break
                except Exception:
                    pass
            if not ok:
                raise PdfEditError("This PDF has a password — type it to open the file." if k == 0
                                   else "A PDF you added has a password — open it on its own and save it first.")
            if k == 0:
                src_encrypted = True
        readers.append(r)

    pages = plan.get("pages")
    if not isinstance(pages, list) or not pages:
        raise PdfEditError("There are no pages left to save.")
    if len(pages) > _PDFED_MAX_PAGES:
        raise PdfEditError(f"That's more than {_PDFED_MAX_PAGES} pages — please split the PDF first.")

    # ---- the pages, in their new order (runs from one file go in together: links and forms keep working)
    plain_order = all(isinstance(p, dict) and not p.get("blank") and int(_pdfed_num(p.get("src"), 0, 99, 0)) == 0
                      and int(_pdfed_num(p.get("i"), 0, 1e6, -1)) == n for n, p in enumerate(pages)) \
        and len(pages) == len(readers[0].pages)
    if plain_order:
        writer = pypdf.PdfWriter(clone_from=readers[0])
    else:
        writer = pypdf.PdfWriter()
        seg = {"src": None, "pages": []}

        def flush():
            if seg["pages"]:
                writer.append(readers[seg["src"]], pages=list(seg["pages"]), import_outline=False)
            seg["src"], seg["pages"] = None, []
        for p in pages:
            if not isinstance(p, dict):
                raise PdfEditError("The changes couldn't be read — please try again.")
            if p.get("blank"):
                flush()
                writer.add_blank_page(_pdfed_num(p.get("w"), 72, 14400, 595.3), _pdfed_num(p.get("h"), 72, 14400, 841.9))
                continue
            s = int(_pdfed_num(p.get("src"), 0, len(readers) - 1, 0))
            i = int(_pdfed_num(p.get("i"), 0, 1e6, 0))
            if i >= len(readers[s].pages):
                raise PdfEditError("The changes refer to a page that isn't in the file.")
            if seg["src"] != s:
                flush()
                seg["src"] = s
            seg["pages"].append(i)
        flush()
    if len(writer.pages) != len(pages):
        raise PdfEditError("The pages couldn't be put together — please try again.")

    # ---- what's drawn on each page
    items = [it for it in (plan.get("items") or []) if isinstance(it, dict)]
    if len(items) > _PDFED_MAX_ITEMS:
        raise PdfEditError("That's a lot of changes for one save — please save part of them first.")
    by_page = {}
    for it in items:
        by_page.setdefault(int(_pdfed_num(it.get("page"), -1, 1e6, -1)), []).append(it)
    out_opts = plan.get("output") if isinstance(plan.get("output"), dict) else {}
    numbers = out_opts.get("page_numbers") if isinstance(out_opts.get("page_numbers"), dict) else None
    watermark = out_opts.get("watermark") if isinstance(out_opts.get("watermark"), dict) else None
    redact_pages, note_items = [], []
    total = len(writer.pages)
    for k in range(total):
        wpage = writer.pages[k]
        its = by_page.get(k, [])
        drops = [_pdfed_rect(it.get("rect")) for it in its if it.get("type") in ("remove_image", "replace_image")]
        edited = [_pdfed_rect(it.get("rect")) for it in its if it.get("type") == "edit"]
        cover = _pdfed_clean(wpage, drops, edited, writer) if (drops or edited) else []
        p = pages[k]
        rot = (int(_pdfed_num(p.get("rotate"), -3600, 3600, 0)) // 90 * 90) % 360
        crop = _pdfed_rect(p.get("crop"))
        if not (crop and crop[2] - crop[0] >= 18 and crop[3] - crop[1] >= 18):
            crop = None
        try:
            seen = (int(wpage.rotation or 0) + rot) % 360
        except Exception:
            seen = rot
        overlay = _pdfed_overlay(wpage, its, cover, assets, k, total, numbers, watermark, view_rot=seen, box=crop)
        if overlay is not None:
            wpage.merge_page(overlay)
        for it in its:
            if it.get("type") == "note":
                note_items.append((k, it))
            elif it.get("type") == "redact":
                redact_pages.append(k)
        if rot:
            wpage.rotate(rot)
        if crop:
            from pypdf.generic import RectangleObject
            wpage.cropbox = RectangleObject(crop)
    if getattr(writer, "_pdfed_dropped", False):
        try:
            writer.compress_identical_objects(remove_identicals=False, remove_orphans=True)
        except Exception:
            logging.getLogger(__name__).warning("PDF editor: unused pictures couldn't be left out of the file", exc_info=True)
    for k, it in note_items[:500]:
        try:
            _pdfed_note(writer, k, it)
        except Exception:
            logging.getLogger(__name__).warning("PDF editor: a note couldn't be added", exc_info=True)
    form = plan.get("form")
    if isinstance(form, dict) and form:
        _pdfed_fill_form(writer, dict(list(form.items())[:1000]), flatten=bool(plan.get("flatten_form")))

    stage = os.path.join(out_dir, "stage_a.pdf")
    with open(stage, "wb") as f:
        writer.write(f)

    # ---- redaction: those pages become pictures of themselves (their words searchable again)
    if redact_pages:
        if len(set(redact_pages)) > _PDFED_RASTER_PAGES:
            raise PdfEditError(f"Redaction works on up to {_PDFED_RASTER_PAGES} pages at a time — please save in parts.")
        stage = _pdfed_rasterize(stage, redact_pages, out_dir)

    ocr = bool(out_opts.get("ocr"))
    compress = bool(out_opts.get("compress"))
    pw = out_opts.get("password")
    perms = out_opts.get("permissions") if isinstance(out_opts.get("permissions"), dict) else {}
    keep_old_pw = src_encrypted and "password" not in out_opts and password
    protect = None
    if pw or keep_old_pw or any(v is False for v in perms.values()):
        flags = UAP(0xFFFFFFFC)          # everything allowed (the reserved bits set, as the PDF rules ask)
        if perms.get("print") is False:
            flags &= ~(UAP.PRINT | UAP.PRINT_TO_REPRESENTATION)
        if perms.get("copy") is False:
            flags &= ~(UAP.EXTRACT | UAP.EXTRACT_TEXT_AND_GRAPHICS)
        if perms.get("edit") is False:
            flags &= ~(UAP.MODIFY | UAP.ADD_OR_MODIFY | UAP.ASSEMBLE_DOC | UAP.FILL_FORM_FIELDS)
        protect = (str(pw if pw else (password if keep_old_pw else ""))[:128], flags)
    groups = _pdfed_groups(out_opts.get("split"), total)

    final = os.path.join(out_dir, "edited.pdf")
    if not (ocr or compress or protect):
        os.replace(stage, final)
    else:
        writer = pypdf.PdfWriter(clone_from=pypdf.PdfReader(stage))
        if ocr:
            done = _pdfed_ocr(writer, stage, out_dir)
            if done >= _PDFED_OCR_PAGES:
                notes.append(f"The first {_PDFED_OCR_PAGES} scanned pages were made searchable — save again to do more.")
        if compress:
            for page in writer.pages:
                for img in page.images:
                    try:
                        img.replace(img.image, quality=int(_pdfed_num(out_opts.get("quality"), 20, 95, 65)))
                    except Exception:
                        continue
                try:
                    page.compress_content_streams()
                except Exception:
                    pass
            try:
                writer.compress_identical_objects()
            except Exception:
                pass
        if protect and not groups:
            writer.encrypt(user_password=protect[0], owner_password=os.urandom(16).hex(), permissions_flag=protect[1], algorithm="AES-256")
        with open(final, "wb") as f:
            writer.write(f)
        if compress and os.path.getsize(final) > os.path.getsize(stage) and not ((protect and not groups) or ocr):
            os.replace(stage, final)        # compressing didn't make it smaller: the uncompressed one is kept
    if groups:
        return _pdfed_split(final, groups, out_dir, protect), notes
    return final, notes


def _pdfed_groups(raw, total):
    """The editor's 'split': lists of page numbers (from 0, in the saved order), one list per file."""
    if raw is None or raw == []:
        return None
    if not isinstance(raw, list):
        raise PdfEditError("The split couldn't be read — please try again.")
    groups = []
    for g in raw[:500]:
        if not isinstance(g, list):
            raise PdfEditError("The split couldn't be read — please try again.")
        idx = []
        for v in g[:_PDFED_MAX_PAGES]:
            try:
                n = int(v)
            except (TypeError, ValueError):
                raise PdfEditError("The split couldn't be read — please try again.")
            if not 0 <= n < total:
                raise PdfEditError("The split refers to a page that isn't in the file.")
            idx.append(n)
        if idx:
            groups.append(idx)
    if len(groups) < 2:
        raise PdfEditError("Splitting needs at least two parts.")
    return groups


def _pdfed_split(path, groups, out_dir, protect=None):
    """The edited PDF as several files — one per group of pages, each protected like the whole would be.
    Files are named 'page N.pdf' when every part is a single page, otherwise 'part N.pdf'."""
    reader = pypdf.PdfReader(path)
    singles = all(len(g) == 1 for g in groups)
    used, out = set(), []
    for n, g in enumerate(groups, 1):
        w = pypdf.PdfWriter()
        w.append(reader, pages=list(g), import_outline=False)
        if protect:
            w.encrypt(user_password=protect[0], owner_password=os.urandom(16).hex(), permissions_flag=protect[1], algorithm="AES-256")
        name = ("page %d.pdf" % (g[0] + 1)) if singles else ("part %d.pdf" % n)
        if name in used:
            name = "part %d.pdf" % n
        used.add(name)
        p = os.path.join(out_dir, "split-" + name)
        with open(p, "wb") as f:
            w.write(f)
        out.append(p)
    return out


def pdf_to_images(src_path, out_dir, dpi=150, max_pages=60, password=None):
    """Each page as a JPG picture (as it's shown, turned pages included). Returns the paths."""
    import pypdfium2 as pdfium
    os.makedirs(out_dir, exist_ok=True)
    try:
        doc = pdfium.PdfDocument(src_path, password=password or None)
    except Exception:
        raise PdfEditError("This PDF couldn't be opened — if it has a password, type it first.")
    try:
        n = len(doc)
        if n > max_pages:
            raise PdfEditError(f"That's more than {max_pages} pages — please split the PDF first.")
        out = []
        for k in range(n):
            img = doc[k].render(scale=dpi / 72.0).to_pil().convert("RGB")
            p = os.path.join(out_dir, "page-%03d.jpg" % (k + 1))
            img.save(p, "JPEG", quality=88)
            out.append(p)
        return out
    finally:
        doc.close()


# ======================= 3.1: the Word editor (R17) =======================
# Opening: the .docx as the blocks the editor shows — paragraphs (their words in runs, with the look each
# run really has), tables, pictures — and, kept exactly as they are, the things that can't be typed over
# safely ("atoms": footnote marks, fields, shapes, equations, page breaks). Paragraphs with tracked
# changes or comments, a table of contents and content controls are shown but not editable here.
# Every block has an id from its place in the file, so the same file read again gives the same ids.
# Saving: the same file plus the editor's list of blocks. A block sent back by its id alone is kept
# exactly; a changed paragraph is rebuilt from its own properties and its runs' own properties (fonts
# and sizes kept), with only what was changed set; a new paragraph copies the one it came from. The rest
# of the file — headers, footers, footnotes, comments, styles, settings — is left as it was.
import base64 as _wd_b64
import copy as _wd_copy
import hashlib as _wd_hash
import io as _wd_io


class WordEditError(ConversionError):
    pass


_WD_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_WD_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_WD_V = "urn:schemas-microsoft-com:vml"
_WD_MATH = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_WD_MAX_BLOCKS = 8000
_WD_MAX_PICS = 80
_WD_PIC_BUDGET = 9 * 1024 * 1024
_WD_OFFER = ["Title", "Subtitle", "Heading1", "Heading2", "Heading3"]
_WD_BAD = "The changes don't match this file — please open it again."
_WD_RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow",
                 "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
                 "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs",
                 "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange"]
_WD_PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
                 "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
                 "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
                 "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
                 "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]
_WD_ALIGN = {"center": "center", "right": "right", "end": "right", "both": "justify", "distribute": "justify",
             "lowKashida": "justify", "mediumKashida": "justify", "highKashida": "justify", "thaiDistribute": "justify"}


def _wd_qn(tag):
    from docx.oxml.ns import qn
    return qn(tag)


def _wd_tag(el):
    """(namespace URI, local name) of an element ('' for comments and the like)."""
    t = el.tag
    if not isinstance(t, str) or not t.startswith("{"):
        return ("", "")
    uri, local = t[1:].split("}", 1)
    return (uri, local)


def _wd_is(el, local, uri=_WD_W):
    return _wd_tag(el) == (uri, local)


def _wd_val(el, attr="w:val", default=None):
    if el is None:
        return default
    v = el.get(_wd_qn(attr))
    return default if v is None else v


def _wd_toggle(el):
    """<w:b/> and the like: None when absent, else on — unless its value says off."""
    if el is None:
        return None
    return (el.get(_wd_qn("w:val")) or "true").lower() not in ("0", "false", "off", "none")


def _wd_num(v, default=None):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _wd_text(el):
    out = []
    for e in el.iter():
        uri, t = _wd_tag(e)
        if (uri == _WD_W and t == "t") or (uri == _WD_MATH and t == "t"):
            out.append(e.text or "")
        elif uri == _WD_W and t == "tab":
            out.append(" ")
    return "".join(out).strip()


def _wd_new(tag, **attrs):
    from docx.oxml import OxmlElement
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(_wd_qn("w:" + k), str(v))
    return e


def _wd_drop(parent, local):
    for e in parent.findall(_wd_qn("w:" + local)):
        parent.remove(e)


def _wd_place(parent, el, order):
    """Puts `el` into a properties element where the file format wants it (Word is strict about the order)."""
    uri, local = _wd_tag(el)
    _wd_drop(parent, local)
    rank = order.index(local) if local in order else len(order)
    for i, c in enumerate(parent):
        cu, cl = _wd_tag(c)
        cr = order.index(cl) if (cu == _WD_W and cl in order) else len(order) + 1
        if cr > rank:
            parent.insert(i, el)
            return el
    parent.append(el)
    return el


def _wd_clean(text):
    import re
    return re.sub("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]", "", str(text or ""))[:100000]


def _wd_href(h):
    h = str(h or "").strip()[:2000]
    low = h.lower()
    if low.startswith(("http://", "https://", "mailto:", "tel:")) or (h.startswith("#") and len(h) > 1):
        return h
    return None


def _wd_load(path):
    from docx import Document
    try:
        return Document(path)
    except Exception:
        raise WordEditError("This Word file couldn't be opened — it may be damaged, or not a .docx file.")


def _wd_sig(path, doc):
    import zipfile
    try:
        with zipfile.ZipFile(path) as z:
            return _wd_hash.sha256(z.read(str(doc.part.partname).lstrip("/"))).hexdigest()[:32]
    except Exception:
        return ""


_WD_TEMPLATE = []


def _wd_template():
    if not _WD_TEMPLATE:
        from docx import Document
        _WD_TEMPLATE.append(Document())
    return _WD_TEMPLATE[0]


class _WdStyles:
    def __init__(self, doc):
        self.root = doc.styles.element
        self.el, self.default_p = {}, None
        for s in self.root.findall(_wd_qn("w:style")):
            sid = s.get(_wd_qn("w:styleId"))
            if not sid:
                continue
            self.el[sid] = s
            if s.get(_wd_qn("w:type")) == "paragraph" and (s.get(_wd_qn("w:default")) or "").lower() in ("1", "true", "on"):
                self.default_p = sid
        dd = self.root.find(_wd_qn("w:docDefaults"))
        self.rpr0 = dd.find(_wd_qn("w:rPrDefault") + "/" + _wd_qn("w:rPr")) if dd is not None else None
        self.ppr0 = dd.find(_wd_qn("w:pPrDefault") + "/" + _wd_qn("w:pPr")) if dd is not None else None
        self.theme = _wd_theme_fonts(doc)

    def chain(self, sid):
        out, seen = [], set()
        while sid and sid in self.el and sid not in seen:
            seen.add(sid)
            out.append(self.el[sid])
            sid = _wd_val(self.el[sid].find(_wd_qn("w:basedOn")))
        return out

    def name(self, sid):
        s = self.el.get(sid)
        return _wd_val(s.find(_wd_qn("w:name")) if s is not None else None) or sid

    def font(self, e):
        for a in ("w:ascii", "w:hAnsi"):
            v = e.get(_wd_qn(a))
            if v:
                return v
        for a in ("w:asciiTheme", "w:hAnsiTheme"):
            v = e.get(_wd_qn(a))
            if v:
                return self.theme.get("major" if v.startswith("major") else "minor")
        return None

    def rlayers(self, pstyle, rpr=None, rstyle=None):
        out = [rpr] if rpr is not None else []
        if rstyle is None and rpr is not None:
            rstyle = _wd_val(rpr.find(_wd_qn("w:rStyle")))
        for s in self.chain(rstyle) + self.chain(pstyle or self.default_p):
            x = s.find(_wd_qn("w:rPr"))
            if x is not None:
                out.append(x)
        if self.rpr0 is not None:
            out.append(self.rpr0)
        return out

    def players(self, pstyle, ppr=None):
        out = [ppr] if ppr is not None else []
        for s in self.chain(pstyle or self.default_p):
            x = s.find(_wd_qn("w:pPr"))
            if x is not None:
                out.append(x)
        if self.ppr0 is not None:
            out.append(self.ppr0)
        return out

    def look(self, layers):
        """How words look: bold, italic, underline, strike-out, size (pt), font, colour, highlight, raised/lowered."""
        def first(tag, fn):
            for L in layers:
                e = L.find(_wd_qn("w:" + tag))
                if e is not None:
                    v = fn(e)
                    if v is not None:
                        return v
            return None
        sz = first("sz", lambda e: _wd_num(e.get(_wd_qn("w:val"))))
        col = first("color", lambda e: "" if (e.get(_wd_qn("w:val")) or "auto").lower() == "auto" else e.get(_wd_qn("w:val")).upper())
        hl = first("highlight", lambda e: "" if (e.get(_wd_qn("w:val")) or "none") == "none" else e.get(_wd_qn("w:val")))
        va = first("vertAlign", lambda e: {"superscript": "sup", "subscript": "sub"}.get(e.get(_wd_qn("w:val")) or "", ""))
        return {"b": bool(first("b", _wd_toggle)), "i": bool(first("i", _wd_toggle)),
                "u": bool(first("u", lambda e: (e.get(_wd_qn("w:val")) or "single") != "none")),
                "s": bool(first("strike", _wd_toggle) or first("dstrike", _wd_toggle)),
                "sz": round(sz / 2.0, 2) if sz else 10.0, "f": first("rFonts", self.font),
                "c": col or None, "hl": hl or None, "va": va or None, "caps": bool(first("caps", _wd_toggle))}

    def plook(self, layers):
        def attr(tag, a):
            for L in layers:
                e = L.find(_wd_qn("w:" + tag))
                if e is not None and e.get(_wd_qn("w:" + a)) is not None:
                    return e.get(_wd_qn("w:" + a))
            return None

        def child(tag, sub):
            for L in layers:
                e = L.find(_wd_qn("w:" + tag))
                s = e.find(_wd_qn("w:" + sub)) if e is not None else None
                if s is not None:
                    return _wd_val(s)
            return None
        line, rule = _wd_num(attr("spacing", "line")), attr("spacing", "lineRule") or "auto"
        hang = _wd_num(attr("ind", "hanging"))
        first_line = -hang if hang else _wd_num(attr("ind", "firstLine"), 0.0)
        return {"align": _WD_ALIGN.get(attr("jc", "val") or "", "left"),
                "before": _wd_num(attr("spacing", "before"), 0.0) / 20.0, "after": _wd_num(attr("spacing", "after"), 0.0) / 20.0,
                "line": (round(line / 240.0, 3) if rule == "auto" else -round(line / 20.0, 2)) if line else None,
                "left": _wd_num(attr("ind", "left") or attr("ind", "start"), 0.0) / 20.0, "first": (first_line or 0.0) / 20.0,
                "num": child("numPr", "numId"), "ilvl": int(_wd_num(child("numPr", "ilvl"), 0) or 0)}

    def list_of(self, pstyle):
        """The list a paragraph style makes its paragraphs part of: (numId, level), or None."""
        pl = self.plook(self.players(pstyle))
        return (pl["num"], pl["ilvl"]) if pl["num"] and pl["num"] != "0" else None


def _wd_theme_fonts(doc):
    out = {"minor": None, "major": None}
    try:
        from docx.oxml import parse_xml
        for rel in doc.part.rels.values():
            if rel.reltype.endswith("/theme") and not rel.is_external:
                root = parse_xml(rel.target_part.blob)
                a = "http://schemas.openxmlformats.org/drawingml/2006/main"
                for kind in ("minor", "major"):
                    e = root.find(".//{%s}%sFont/{%s}latin" % (a, kind, a))
                    if e is not None:
                        out[kind] = e.get("typeface")
    except Exception:
        pass
    return out


class _WdNumbering:
    def __init__(self, doc, create=False):
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        self.root = None
        try:
            self.root = doc.part.part_related_by(RT.NUMBERING).element
        except KeyError:
            if create:
                self.root = doc.part.numbering_part.element
        self.made = {}
        self.read()

    def read(self):
        self.abstract, self.nums = {}, {}
        if self.root is None:
            return
        for a in self.root.findall(_wd_qn("w:abstractNum")):
            self.abstract[a.get(_wd_qn("w:abstractNumId"))] = a
        for n in self.root.findall(_wd_qn("w:num")):
            over = {}
            for o in n.findall(_wd_qn("w:lvlOverride")):
                so = o.find(_wd_qn("w:startOverride"))
                if so is not None:
                    over[int(_wd_num(o.get(_wd_qn("w:ilvl")), 0))] = int(_wd_num(_wd_val(so), 1))
            self.nums[n.get(_wd_qn("w:numId"))] = (_wd_val(n.find(_wd_qn("w:abstractNumId"))), over)

    def level(self, num_id, ilvl):
        """(format, text, start) of one level of a list — e.g. ('decimal', '%1.', 1) or ('bullet', '•', 1)."""
        got = self.nums.get(str(num_id))
        a = self.abstract.get(got[0]) if got else None
        if a is None:
            return None
        for lv in a.findall(_wd_qn("w:lvl")):
            if int(_wd_num(lv.get(_wd_qn("w:ilvl")), 0)) == ilvl:
                return (_wd_val(lv.find(_wd_qn("w:numFmt")), default="decimal"), _wd_val(lv.find(_wd_qn("w:lvlText")), default=""),
                        got[1].get(ilvl, int(_wd_num(_wd_val(lv.find(_wd_qn("w:start"))), 1))))
        return None

    def kind(self, num_id):
        lv = self.level(num_id, 0)
        return None if lv is None else ("bullet" if lv[0] in ("bullet", "none") else "number")

    def pick(self, kind, num):
        """The list a paragraph goes in: the one asked for, or a list of the right kind (a numbered list
        asked for as new starts again at 1)."""
        num = str(num) if num is not None else ""
        if num in self.nums and self.kind(num) == kind:
            return num
        if num in self.made:
            return self.made[num]
        abstract = next((self.nums[n][0] for n in sorted(self.nums, key=lambda v: int(_wd_num(v, 0))) if self.kind(n) == kind), None)
        if kind == "bullet" and abstract is not None:
            got = next(n for n in sorted(self.nums, key=lambda v: int(_wd_num(v, 0))) if self.kind(n) == "bullet")
            self.made[num] = got
            return got
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls
        if abstract is None:
            abstract = str(max([int(_wd_num(a, 0)) for a in self.abstract] + [0]) + 1)
            levels = []
            for i in range(9):
                if kind == "bullet":
                    ch = ["•", "◦", "▪"][i % 3]
                    levels.append('<w:lvl w:ilvl="%d"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="%s"/><w:lvlJc w:val="left"/>'
                                  '<w:pPr><w:ind w:left="%d" w:hanging="360"/></w:pPr></w:lvl>' % (i, ch, 720 * (i + 1)))
                else:
                    fmt = ["decimal", "lowerLetter", "lowerRoman"][i % 3]
                    levels.append('<w:lvl w:ilvl="%d"><w:start w:val="1"/><w:numFmt w:val="%s"/><w:lvlText w:val="%%%d."/><w:lvlJc w:val="left"/>'
                                  '<w:pPr><w:ind w:left="%d" w:hanging="360"/></w:pPr></w:lvl>' % (i, fmt, i + 1, 720 * (i + 1)))
            el = parse_xml('<w:abstractNum %s w:abstractNumId="%s"><w:multiLevelType w:val="hybridMultilevel"/>%s</w:abstractNum>'
                           % (nsdecls("w"), abstract, "".join(levels)))
            first_num = self.root.find(_wd_qn("w:num"))
            if first_num is not None:
                first_num.addprevious(el)
            else:
                self.root.append(el)
        nid = str(max([int(_wd_num(n, 0)) for n in self.nums] + [0]) + 1)
        el = parse_xml('<w:num %s w:numId="%s"><w:abstractNumId w:val="%s"/>%s</w:num>' % (
            nsdecls("w"), nid, abstract, '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>' if kind == "number" else ""))
        nums = self.root.findall(_wd_qn("w:num"))
        if nums:
            nums[-1].addnext(el)
        else:
            self.root.append(el)
        self.read()
        self.made[num] = nid
        return nid


class _WdCtx:
    """What a Word file has, by id: blocks, paragraphs, the runs words come from, atoms, bookmarks."""

    def __init__(self, doc, view=True, assets=None):
        self.doc, self.view, self.assets = doc, view, assets or {}
        self.styles = _WdStyles(doc)
        self.blocks, self.paras, self.runs, self.atoms, self.marks, self.locked = {}, {}, {}, {}, {}, set()
        self.field_open, self.notes, self.endnotes, self.pics, self.pic_bytes = 0, 0, 0, 0, 0
        self.used_styles, self.used_lists = set(), set()
        self._num, self.link_rids, self.known_links = None, {}, set()

    def numbering(self, create=False):
        if self._num is None or (create and self._num.root is None):
            self._num = _WdNumbering(self.doc, create=create)
        return self._num

    # ---------------------------------------------------------------- reading
    def link_target(self, h):
        rid = h.get(_wd_qn("r:id"))
        if rid:
            rel = self.doc.part.rels.get(rid)
            if rel is not None and rel.is_external:
                self.known_links.add(rel.target_ref)
                return rel.target_ref
        anchor = h.get(_wd_qn("w:anchor"))
        return "#" + anchor if anchor else None

    def preview(self, rid):
        if not self.view or not rid or self.pics >= _WD_MAX_PICS:
            return None
        part = self.doc.part.related_parts.get(rid)
        blob = getattr(part, "blob", None)
        if not blob:
            return None
        try:
            from PIL import Image
            im = Image.open(_wd_io.BytesIO(blob))
            im.load()
            if max(im.size) > 1200:
                im.thumbnail((1200, 1200))
            buf = _wd_io.BytesIO()
            if im.mode in ("RGBA", "LA", "P"):
                im.convert("RGBA").save(buf, "PNG", optimize=True)
                mime = "image/png"
            else:
                im.convert("RGB").save(buf, "JPEG", quality=80)
                mime = "image/jpeg"
        except Exception:
            return None
        data = buf.getvalue()
        if self.pic_bytes + len(data) > _WD_PIC_BUDGET:
            return None
        self.pic_bytes += len(data)
        self.pics += 1
        return "data:%s;base64,%s" % (mime, _wd_b64.b64encode(data).decode("ascii"))

    def atom_view(self, x):
        uri, t = _wd_tag(x)
        if (uri, t) in ((_WD_W, "drawing"), (_WD_W, "pict"), (_WD_MC, "AlternateContent"), (_WD_W, "object")):
            blip = x.find(".//" + _wd_qn("a:blip"))
            rid = blip.get(_wd_qn("r:embed")) or blip.get(_wd_qn("r:link")) if blip is not None else None
            if rid is None:
                vimg = x.find(".//{%s}imagedata" % _WD_V)
                rid = vimg.get(_wd_qn("r:id")) if vimg is not None else None
            ext = x.find(".//" + _wd_qn("wp:extent"))
            out = {"k": "chip", "label": "Picture"}
            if ext is not None:
                out["w"] = round(_wd_num(ext.get("cx"), 0) / 12700.0, 1)
                out["h"] = round(_wd_num(ext.get("cy"), 0) / 12700.0, 1)
            if x.find(".//" + _wd_qn("wp:anchor")) is not None:
                out["float"] = True
            if rid is None:
                gd = x.find(".//" + _wd_qn("a:graphicData"))
                uri2 = (gd.get("uri") or "") if gd is not None else ""
                txt = _wd_text(x)
                out["label"] = ("Text box: " + txt[:80]) if txt else ("Chart" if "chart" in uri2 else ("Diagram" if "diagram" in uri2 else
                                                                                                        ("Embedded object" if t == "object" else "Shape")))
                return out
            src = self.preview(rid)
            if src:
                out.update({"k": "img", "src": src})
            return out
        if (uri, t) == (_WD_W, "footnoteReference"):
            self.notes += 1
            return {"k": "note", "label": str(self.notes)}
        if (uri, t) == (_WD_W, "endnoteReference"):
            self.endnotes += 1
            return {"k": "note", "label": "i ii iii iv v vi vii viii ix x".split()[(self.endnotes - 1) % 10]}
        if (uri, t) == (_WD_W, "br"):
            return {"k": "break", "label": "Page break" if x.get(_wd_qn("w:type")) == "page" else "Column break"}
        if (uri, t) in ((_WD_MATH, "oMath"), (_WD_MATH, "oMathPara")):
            return {"k": "math", "label": _wd_text(x)[:200] or "Equation"}
        if (uri, t) == (_WD_W, "fldSimple"):
            return {"k": "field", "label": _wd_text(x)[:200] or _wd_field_word(x.get(_wd_qn("w:instr")))}
        if (uri, t) == (_WD_W, "sym"):
            return {"k": "chip", "label": "◆"}
        if (uri, t) == (_WD_W, "ptab"):
            return {"k": "chip", "label": "⇥"}
        return {"k": "chip", "label": _wd_text(x)[:200] or "•"}

    def items(self, parent, ppr, href=None):
        """A paragraph's content as one flat list: (kind, value, run properties, link, element)."""
        out = []
        for c in parent:
            if c is ppr:
                continue
            uri, t = _wd_tag(c)
            if uri != _WD_W:
                out.append(("xel", None, None, None, c))
            elif t == "r":
                rpr = c.find(_wd_qn("w:rPr"))
                for x in c:
                    if x is rpr:
                        continue
                    xu, xt = _wd_tag(x)
                    if xu == _WD_W and xt == "t":
                        out.append(("t", x.text or "", rpr, href, x))
                    elif xu == _WD_W and xt == "tab":
                        out.append(("t", "\t", rpr, href, x))
                    elif xu == _WD_W and (xt == "cr" or (xt == "br" and x.get(_wd_qn("w:type")) in (None, "textWrapping"))):
                        out.append(("t", "\n", rpr, href, x))
                    elif xu == _WD_W and xt == "softHyphen":
                        out.append(("t", "­", rpr, href, x))
                    elif xu == _WD_W and xt == "noBreakHyphen":
                        out.append(("t", "‑", rpr, href, x))
                    elif xu == _WD_W and xt == "lastRenderedPageBreak":
                        continue
                    elif xu == _WD_W and xt == "fldChar":
                        out.append(("fld", x.get(_wd_qn("w:fldCharType")), rpr, href, x))
                    elif xu == _WD_W and xt == "instrText":
                        out.append(("instr", x.text or "", rpr, href, x))
                    elif xu == _WD_W and xt in ("commentReference",):
                        out.append(("lock", "comments", rpr, href, x))
                    elif xu == _WD_W and xt in ("delText", "delInstrText"):
                        out.append(("lock", "changes", rpr, href, x))
                    else:
                        out.append(("x", None, rpr, href, x))
            elif t == "hyperlink" and href is None and self.simple_link(c):
                out.extend(self.items(c, None, href=self.link_target(c) or ""))
            elif t in ("bookmarkStart", "bookmarkEnd"):
                out.append(("mark", None, None, None, c))
            elif t in ("proofErr", "permStart", "permEnd"):
                continue
            elif t in ("ins", "del", "moveFrom", "moveTo", "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd"):
                out.append(("lock", "changes", None, None, c))
            elif t in ("commentRangeStart", "commentRangeEnd"):
                out.append(("lock", "comments", None, None, c))
            else:
                out.append(("xel", None, None, None, c))
        return out

    @staticmethod
    def simple_link(h):
        for r in h:
            uri, t = _wd_tag(r)
            if uri == _WD_W and t == "proofErr":
                continue
            if uri != _WD_W or t != "r":
                return False
            for x in r:
                xu, xt = _wd_tag(x)
                if xu != _WD_W or xt not in ("rPr", "t", "tab", "cr", "br", "softHyphen", "noBreakHyphen", "lastRenderedPageBreak"):
                    return False
                if xt == "br" and x.get(_wd_qn("w:type")) not in (None, "textWrapping"):
                    return False
        return True

    def para(self, p, pid):
        self.blocks[pid] = self.paras[pid] = p
        ppr = p.find(_wd_qn("w:pPr"))
        pstyle = _wd_val(ppr.find(_wd_qn("w:pStyle")) if ppr is not None else None)
        if pstyle not in self.styles.el:
            pstyle = self.styles.default_p
        why = "field" if self.field_open > 0 else None
        if ppr is not None:
            prpr = ppr.find(_wd_qn("w:rPr"))
            if ppr.find(_wd_qn("w:pPrChange")) is not None or (prpr is not None and any(
                    prpr.find(_wd_qn("w:" + t)) is not None for t in ("ins", "del", "moveFrom", "moveTo"))):
                why = why or "changes"
        items = self.items(p, ppr)
        pieces, marks, k, i, n = [], [], 0, 0, len(items)
        base = self.styles.look(self.styles.rlayers(pstyle)) if self.view else None
        while i < n:
            kind, val, rpr, href, x = items[i]
            if kind == "mark":
                marks.append(x)
                i += 1
                continue
            if kind == "lock":
                why = why or val
                i += 1
                continue
            o = "%s:%d" % (pid, k)
            if kind == "fld" and val == "begin":
                depth, j = 0, i
                while j < n:
                    if items[j][0] == "fld" and items[j][1] == "begin":
                        depth += 1
                    elif items[j][0] == "fld" and items[j][1] == "end":
                        depth -= 1
                        if depth == 0:
                            break
                    j += 1
                if depth > 0:                      # it goes on into the next paragraphs (a table of contents)
                    why = why or "field"
                    break
                group = items[i:j + 1]
                self.atoms[o] = ("items", group)
                if self.view:
                    pieces.append(dict({"a": o, "k": "field", "label": _wd_field_label(group)}))
                k += 1
                i = j + 1
                continue
            if kind == "fld" and val == "end" and self.field_open > 0:
                self.field_open -= 1
                why = why or "field"
                i += 1
                continue
            if kind == "t":
                text, j = val, i + 1
                while j < n and items[j][0] == "t" and items[j][2] is rpr and items[j][3] == href:
                    text += items[j][1]
                    j += 1
                self.runs[o] = rpr
                if self.view:
                    look = self.styles.look(self.styles.rlayers(pstyle, rpr))
                    piece = {"x": text, "o": o}
                    if href:
                        piece["h"] = href
                    for key in ("b", "i", "u", "s"):
                        if look[key] != base[key]:
                            piece[key] = look[key]
                    for key in ("sz", "f", "c", "hl", "va"):
                        if look[key] != base[key] and look[key] is not None:
                            piece[key] = look[key]
                    dc = rpr.find(_wd_qn("w:color")) if rpr is not None else None
                    if dc is not None and (dc.get(_wd_qn("w:val")) or "auto").lower() != "auto":
                        piece["dc"] = dc.get(_wd_qn("w:val")).upper()
                    if look["caps"]:
                        piece["caps"] = True
                    pieces.append(piece)
                k += 1
                i = j
                continue
            # anything else is kept exactly as it is
            self.atoms[o] = ("items", [items[i]]) if kind in ("x", "fld", "instr") else ("els", [x])
            if self.view:
                v = self.atom_view(x) if kind in ("x", "xel") else {"k": "field", "label": "Field"}
                v["a"] = o
                pieces.append(v)
            k += 1
            i += 1
        # the rest of a field that started earlier still counts its ends
        while i < n:
            if items[i][0] == "fld" and items[i][1] == "begin":
                self.field_open += 1
            elif items[i][0] == "fld" and items[i][1] == "end" and self.field_open > 0:
                self.field_open -= 1
            i += 1
        self.marks[pid] = marks
        if why:
            self.locked.add(pid)
        if not self.view:
            return None
        self.used_styles.add(pstyle)
        blk = {"id": pid, "t": "p", "style": pstyle, "runs": pieces}
        if why:
            blk["locked"] = why
            blk["text"] = _wd_text(p)[:3000]
        if ppr is not None:
            d = self.styles.plook([ppr])
            if ppr.find(_wd_qn("w:jc")) is not None:
                blk["align"] = d["align"]
            if ppr.find(_wd_qn("w:ind")) is not None:
                blk["ind"] = [d["left"], d["first"]]
            if ppr.find(_wd_qn("w:spacing")) is not None:
                full = self.styles.plook(self.styles.players(pstyle, ppr))
                blk["sp"] = [full["before"], full["after"], full["line"]]
            if ppr.find(_wd_qn("w:sectPr")) is not None:
                blk["sect"] = True
        pl = self.styles.plook(self.styles.players(pstyle, ppr))
        if pl["num"] and pl["num"] != "0":
            lv = self.numbering().level(pl["num"], pl["ilvl"])
            if lv is not None:
                self.used_lists.add(pl["num"])
                blk["list"] = {"kind": "bullet" if lv[0] in ("bullet", "none") else "number", "level": pl["ilvl"], "num": pl["num"]}
        return blk

    def table(self, tbl, tid):
        self.blocks[tid] = tbl
        why = None
        if any(_wd_tag(c)[1] in ("sdt", "customXml") for c in tbl) or tbl.find(_wd_qn("w:tblPr") + "/" + _wd_qn("w:tblPrChange")) is not None:
            why = "controls"
        rows = []
        trs = [c for c in tbl if _wd_is(c, "tr")]
        for r, tr in enumerate(trs):
            if any(_wd_tag(c)[1] in ("sdt", "customXml") for c in tr):
                why = why or "controls"
            trpr = tr.find(_wd_qn("w:trPr"))
            col = int(_wd_num(_wd_val(trpr.find(_wd_qn("w:gridBefore")) if trpr is not None else None), 0) or 0)
            cells = []
            for c, tc in enumerate([x for x in tr if _wd_is(x, "tc")]):
                tcpr = tc.find(_wd_qn("w:tcPr"))
                span = int(_wd_num(_wd_val(tcpr.find(_wd_qn("w:gridSpan")) if tcpr is not None else None), 1) or 1)
                vm = tcpr.find(_wd_qn("w:vMerge")) if tcpr is not None else None
                cell = {"r": r, "c": c, "col": col, "span": span, "blocks": self.container(tc, "%s.%d.%d." % (tid, r, c))}
                if vm is not None:
                    cell["vm"] = "restart" if (vm.get(_wd_qn("w:val")) or "") == "restart" else "continue"
                shd = tcpr.find(_wd_qn("w:shd")) if tcpr is not None else None
                if shd is not None and (shd.get(_wd_qn("w:fill")) or "auto").lower() not in ("auto", "ffffff"):
                    cell["fill"] = shd.get(_wd_qn("w:fill")).upper()
                cells.append(cell)
                col += span
            rows.append(cells)
        if why:
            self.locked.add(tid)
        if not self.view:
            return None
        for r, cells in enumerate(rows):          # merged cells, as rows they cover
            for cell in cells:
                if cell.get("vm") == "restart":
                    n = 1
                    for below in rows[r + 1:]:
                        if any(x.get("vm") == "continue" and x["col"] == cell["col"] for x in below):
                            n += 1
                        else:
                            break
                    cell["rows"] = n
        blk = {"id": tid, "t": "table", "rows": rows}
        if why:
            blk["locked"] = why
        return blk

    def container(self, parent, prefix):
        out = []
        for n, el in enumerate(parent):
            uri, t = _wd_tag(el)
            if uri != _WD_W:
                continue
            bid = prefix + str(n)
            if t == "p":
                b = self.para(el, bid)
            elif t == "tbl":
                b = self.table(el, bid)
            elif t in ("sdt", "customXml", "altChunk"):
                self.blocks[bid] = el
                self.locked.add(bid)
                b = {"id": bid, "t": "locked", "text": _wd_text(el)[:3000], "why": "control"} if self.view else None
            else:
                continue
            if b is not None:
                out.append(b)
            if len(self.blocks) > _WD_MAX_BLOCKS:
                raise WordEditError("This document is too long to edit here — more than %d paragraphs." % _WD_MAX_BLOCKS)
        return out

    # ---------------------------------------------------------------- writing
    def ensure_style(self, sid):
        if sid in self.styles.el:
            return
        tpl = _wd_template().styles.element
        todo, have = [sid], dict((s.get(_wd_qn("w:styleId")), s) for s in tpl.findall(_wd_qn("w:style")))
        if sid not in have:
            raise WordEditError("That style isn't in this document — please choose another.")
        while todo:
            s = todo.pop()
            if s in self.styles.el or s not in have:
                continue
            copy = _wd_copy.deepcopy(have[s])
            self.styles.root.append(copy)
            self.styles.el[s] = copy
            for ref in ("w:basedOn", "w:link", "w:next"):
                v = _wd_val(copy.find(_wd_qn(ref)))
                if v and v not in self.styles.el:
                    todo.append(v)

    def build(self, items, used, within=None):
        out = []
        if not isinstance(items, list):
            raise WordEditError(_WD_BAD)
        for it in items[:_WD_MAX_BLOCKS * 2]:
            if not isinstance(it, dict):
                raise WordEditError(_WD_BAD)
            bid = it.get("id")
            if bid is not None:
                bid = str(bid)
                el = self.blocks.get(bid)
                if el is None or bid in used or (within and not (bid.startswith(within) and "." not in bid[len(within):])) \
                        or (not within and "." in bid):
                    raise WordEditError(_WD_BAD)
                used.add(bid)
                if "p" in it:
                    if not _wd_is(el, "p") or bid in self.locked:
                        raise WordEditError(_WD_BAD)
                    self.fill(el, it["p"], bid)
                elif "cells" in it:
                    if not _wd_is(el, "tbl") or bid in self.locked:
                        raise WordEditError(_WD_BAD)
                    self.cells(el, bid, it["cells"], used)
                out.append((el, bid))
                continue
            p = _wd_new("w:p")
            model = self.paras.get(str(it.get("from"))) if it.get("from") is not None else None
            mp = model.find(_wd_qn("w:pPr")) if model is not None else None
            if mp is not None:
                np_ = _wd_copy.deepcopy(mp)
                for t in ("sectPr", "pPrChange"):
                    _wd_drop(np_, t)
                prpr = np_.find(_wd_qn("w:rPr"))
                if prpr is not None:
                    for t in ("ins", "del", "moveFrom", "moveTo", "rPrChange"):
                        _wd_drop(prpr, t)
                p.append(np_)
            self.fill(p, it.get("p") or {}, None)
            out.append((p, None))
        return out

    def cells(self, tbl, tid, cells, used):
        if not isinstance(cells, dict):
            raise WordEditError(_WD_BAD)
        trs = [c for c in tbl if _wd_is(c, "tr")]
        for key, items in list(cells.items())[:20000]:
            try:
                r, c = [int(v) for v in str(key).split(".")]
                tc = [x for x in trs[r] if _wd_is(x, "tc")][c]
            except (ValueError, IndexError):
                raise WordEditError(_WD_BAD)
            built = self.build(items, used, within="%s.%d.%d." % (tid, r, c))
            tcpr = tc.find(_wd_qn("w:tcPr"))
            extras = [x for x in tc if x is not tcpr and _wd_tag(x)[1] not in ("p", "tbl", "sdt", "customXml", "altChunk")]
            for x in list(tc):
                if x is not tcpr:
                    tc.remove(x)
            for x in extras:
                tc.append(x)
            for el, _ in built:
                tc.append(el)
            if not built or not _wd_is(built[-1][0], "p"):
                tc.append(_wd_new("w:p"))          # a cell always ends with a paragraph

    def fill(self, p, P, pid):
        if not isinstance(P, dict):
            raise WordEditError(_WD_BAD)
        ppr = p.find(_wd_qn("w:pPr"))
        if ppr is None:
            ppr = _wd_new("w:pPr")
            p.insert(0, ppr)
        cur = _wd_val(ppr.find(_wd_qn("w:pStyle"))) or self.styles.default_p
        style = P.get("style")
        if style is not None and str(style)[:120] != cur:
            sid = str(style)[:120]
            if sid == self.styles.default_p:
                _wd_drop(ppr, "pStyle")
            else:
                self.ensure_style(sid)
                _wd_place(ppr, _wd_new("w:pStyle", val=sid), _WD_PPR_ORDER)
        pstyle = _wd_val(ppr.find(_wd_qn("w:pStyle"))) or self.styles.default_p
        if "align" in P:
            _wd_drop(ppr, "jc")
            al = P.get("align")
            if al in ("left", "center", "right", "justify") and al != self.styles.plook(self.styles.players(pstyle))["align"]:
                _wd_place(ppr, _wd_new("w:jc", val={"justify": "both"}.get(al, al)), _WD_PPR_ORDER)
        if "list" in P:
            self.set_list(ppr, pstyle, P.get("list"))
        keep = self.marks.get(pid, []) if pid else []
        for c in list(p):
            if c is not ppr:
                p.remove(c)
        for m in keep:
            if _wd_is(m, "bookmarkStart"):
                p.append(m)
        for r in self.runs_xml(P.get("runs") or [], pstyle):
            p.append(r)
        for m in keep:
            if _wd_is(m, "bookmarkEnd"):
                p.append(m)
        if len(ppr) == 0 and not ppr.attrib:
            p.remove(ppr)

    def set_list(self, ppr, pstyle, lst):
        sl = self.styles.list_of(pstyle)
        _wd_drop(ppr, "numPr")
        if not lst:
            if sl:
                np_ = _wd_new("w:numPr")
                np_.append(_wd_new("w:numId", val="0"))
                _wd_place(ppr, np_, _WD_PPR_ORDER)
            return
        if not isinstance(lst, dict) or lst.get("kind") not in ("bullet", "number"):
            raise WordEditError(_WD_BAD)
        level = max(0, min(8, int(_wd_num(lst.get("level"), 0) or 0)))
        num = self.numbering(create=True).pick(lst["kind"], lst.get("num"))
        if sl and sl[0] == num and sl[1] == level:
            return
        np_ = _wd_new("w:numPr")
        np_.append(_wd_new("w:ilvl", val=level))
        np_.append(_wd_new("w:numId", val=num))
        _wd_place(ppr, np_, _WD_PPR_ORDER)

    def runs_xml(self, runs, pstyle):
        out, link = [], None
        for rs in runs[:50000]:
            if not isinstance(rs, dict):
                continue
            if "a" in rs:
                at = self.atoms.get(str(rs["a"]))
                if at is None:
                    raise WordEditError(_WD_BAD)
                out.extend(_wd_atom_xml(at))
                link = None
                continue
            if "pic" in rs:
                out.append(self.picture(rs))
                link = None
                continue
            text = _wd_clean(rs.get("x"))
            if not text:
                continue
            base = self.runs.get(str(rs.get("o"))) if rs.get("o") is not None else None
            h = str(rs.get("h") or "")
            href = h if h in self.known_links else _wd_href(h)
            r = self.run_xml(text, base, rs, pstyle, href)
            if href:
                if link is not None and link[0] == href:
                    link[1].append(r)
                else:
                    hl = self.hyperlink(href)
                    hl.append(r)
                    out.append(hl)
                    link = (href, hl)
            else:
                out.append(r)
                link = None
        return out

    def run_xml(self, text, base, fmt, pstyle, href):
        import re
        r = _wd_new("w:r")
        rpr = _wd_copy.deepcopy(base) if base is not None else _wd_new("w:rPr")
        _wd_drop(rpr, "rPrChange")
        if href and base is None:
            # a new link looks like one (Word's own link colour, underlined)
            _wd_place(rpr, _wd_new("w:color", val="0563C1"), _WD_RPR_ORDER)
            _wd_place(rpr, _wd_new("w:u", val="single"), _WD_RPR_ORDER)
        rstyle = _wd_val(rpr.find(_wd_qn("w:rStyle")))
        d = self.styles.look(self.styles.rlayers(pstyle, None, rstyle))
        for key, tag in (("b", "b"), ("i", "i"), ("s", "strike")):
            if key not in fmt:
                continue
            want = fmt.get(key)
            _wd_drop(rpr, tag)
            if key == "s":
                _wd_drop(rpr, "dstrike")
            if want is not None and bool(want) != d[key]:
                _wd_place(rpr, _wd_new("w:" + tag) if want else _wd_new("w:" + tag, val="0"), _WD_RPR_ORDER)
        if "u" in fmt and not (href and base is None):
            want, cur = fmt.get("u"), rpr.find(_wd_qn("w:u"))
            had = cur is not None and (cur.get(_wd_qn("w:val")) or "single") != "none"
            if want is None or bool(want) == d["u"]:
                _wd_drop(rpr, "u")
            elif want and not had:
                _wd_place(rpr, _wd_new("w:u", val="single"), _WD_RPR_ORDER)
            elif not want:
                _wd_place(rpr, _wd_new("w:u", val="none"), _WD_RPR_ORDER)
        if "c" in fmt and not (href and base is None):
            c = str(fmt.get("c") or "").lstrip("#").upper()
            cur = rpr.find(_wd_qn("w:color"))
            if not re.fullmatch("[0-9A-F]{6}", c):
                _wd_drop(rpr, "color")
            elif cur is None or (cur.get(_wd_qn("w:val")) or "").upper() != c:
                _wd_place(rpr, _wd_new("w:color", val=c), _WD_RPR_ORDER)
        if "va" in fmt:
            _wd_drop(rpr, "vertAlign")
            if fmt.get("va") in ("sup", "sub"):
                _wd_place(rpr, _wd_new("w:vertAlign", val="superscript" if fmt["va"] == "sup" else "subscript"), _WD_RPR_ORDER)
        if len(rpr):
            r.append(rpr)
        for part in re.split("(\t|\n)", text):
            if part == "\t":
                r.append(_wd_new("w:tab"))
            elif part == "\n":
                r.append(_wd_new("w:br"))
            elif part:
                t = _wd_new("w:t")
                t.text = part
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                r.append(t)
        return r

    def hyperlink(self, href):
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        hl = _wd_new("w:hyperlink")
        if href.startswith("#"):
            hl.set(_wd_qn("w:anchor"), href[1:200])
        else:
            rid = self.link_rids.get(href)
            if rid is None:
                rid = self.link_rids[href] = self.doc.part.relate_to(href, RT.HYPERLINK, is_external=True)
            hl.set(_wd_qn("r:id"), rid)
        hl.set(_wd_qn("w:history"), "1")
        return hl

    def text_width(self):
        sp = self.doc.element.body.find(_wd_qn("w:sectPr"))
        pg = sp.find(_wd_qn("w:pgSz")) if sp is not None else None
        mar = sp.find(_wd_qn("w:pgMar")) if sp is not None else None
        w = _wd_num(pg.get(_wd_qn("w:w")), 12240) if pg is not None else 12240
        m = (_wd_num(mar.get(_wd_qn("w:left")), 1440) + _wd_num(mar.get(_wd_qn("w:right")), 1440)) if mar is not None else 2880
        return max(72.0, (w - m) / 20.0)

    def picture(self, rs):
        from docx.shared import Emu
        name = str(rs.get("pic") or "")
        path = self.assets.get(name)
        if not path or not os.path.exists(path):
            raise WordEditError("A picture you added is missing — please add it again.")
        try:
            from PIL import Image
            with Image.open(path) as im:
                iw, ih = im.size
        except Exception:
            raise WordEditError("A picture you added couldn't be read — please add another.")
        w = _wd_num(rs.get("w")) or 0
        w = min(w if w > 0 else 300.0, self.text_width())
        h = w * ih / max(1, iw)
        inline = self.doc.part.new_pic_inline(path, Emu(int(w * 12700)), Emu(int(h * 12700)))
        r = _wd_new("w:r")
        d = _wd_new("w:drawing")
        d.append(inline)
        r.append(d)
        return r


def _wd_field_word(instr):
    word = (str(instr or "").strip().split() or ["Field"])[0].upper()
    return {"PAGE": "Page number", "NUMPAGES": "Number of pages", "DATE": "Date", "TIME": "Time", "TOC": "Table of contents",
            "REF": "Cross-reference", "PAGEREF": "Page reference", "HYPERLINK": "Link", "SEQ": "Number", "CITATION": "Citation",
            "BIBLIOGRAPHY": "Bibliography"}.get(word, "Field")


def _wd_field_label(group):
    depth, shown, result, instr = 0, False, [], []
    for kind, val, rpr, href, x in group:
        if kind == "fld":
            if val == "begin":
                depth += 1
            elif val == "end":
                depth -= 1
            elif val == "separate" and depth == 1:
                shown = True
        elif kind == "instr" and depth == 1:
            instr.append(val)
        elif kind == "t" and shown and depth == 1:
            result.append(val)
    text = "".join(result).strip()
    return text[:200] if text else _wd_field_word("".join(instr))


def _wd_atom_xml(atom):
    """An atom back as file content: the elements it was, or its pieces each in a run of their own."""
    kind, data = atom
    if kind == "els":
        return [_wd_copy.deepcopy(e) for e in data]
    out = []
    for k, val, rpr, href, x in data:
        if k == "mark":
            out.append(_wd_copy.deepcopy(x))
            continue
        r = _wd_new("w:r")
        if rpr is not None:
            r.append(_wd_copy.deepcopy(rpr))
        r.append(_wd_copy.deepcopy(x))
        out.append(r)
    return out


def _wd_page(body):
    """The first section's page (where a document's words mostly start)."""
    out = {"w": 612.0, "h": 792.0, "ml": 72.0, "mr": 72.0, "mt": 72.0, "mb": 72.0}
    sp = next(iter(body.iter(_wd_qn("w:sectPr"))), None)
    pg = sp.find(_wd_qn("w:pgSz")) if sp is not None else None
    mar = sp.find(_wd_qn("w:pgMar")) if sp is not None else None
    if pg is not None:
        out["w"] = _wd_num(pg.get(_wd_qn("w:w")), 12240) / 20.0
        out["h"] = _wd_num(pg.get(_wd_qn("w:h")), 15840) / 20.0
    if mar is not None:
        for k, a in (("ml", "w:left"), ("mr", "w:right"), ("mt", "w:top"), ("mb", "w:bottom")):
            out[k] = abs(_wd_num(mar.get(_wd_qn(a)), 1440)) / 20.0
    return out


def _wd_style_view(ctx, sid, missing=False):
    st = ctx.styles if not missing else _WdStyles(_wd_template())
    rl = st.look(st.rlayers(sid))
    pl = st.plook(st.players(sid))
    nxt = _wd_val(st.el[sid].find(_wd_qn("w:next"))) if sid in st.el else None
    lst = None
    if not missing and pl["num"] and pl["num"] != "0":
        lv = ctx.numbering().level(pl["num"], pl["ilvl"])
        if lv is not None:
            ctx.used_lists.add(pl["num"])
            lst = {"kind": "bullet" if lv[0] in ("bullet", "none") else "number", "level": pl["ilvl"], "num": pl["num"]}
    v = {"name": st.name(sid), "b": rl["b"], "i": rl["i"], "u": rl["u"], "sz": rl["sz"], "f": rl["f"], "c": rl["c"],
         "caps": rl["caps"], "align": pl["align"], "before": pl["before"], "after": pl["after"], "line": pl["line"],
         "left": pl["left"], "first": pl["first"], "next": nxt if nxt in st.el else None, "list": lst}
    if missing:
        v["missing"] = True
    return v


def word_open(path):
    """A .docx for the Word editor: its blocks, styles, lists and page, as plain data (see _WdCtx)."""
    if path:
        doc = _wd_load(path)
    else:
        from docx import Document
        doc = Document()
    ctx = _WdCtx(doc, view=True)
    blocks = ctx.container(doc.element.body, "b")
    styles = {}
    offer = [ctx.styles.default_p] + [s for s in _WD_OFFER if s != ctx.styles.default_p]
    for sid in [s for s in offer if s] + sorted(s for s in ctx.used_styles if s):
        if sid in styles:
            continue
        if sid in ctx.styles.el:
            styles[sid] = _wd_style_view(ctx, sid)
        elif sid in _WdStyles(_wd_template()).el:
            styles[sid] = _wd_style_view(ctx, sid, missing=True)
    lists = {}
    num = ctx.numbering()
    for nid in sorted(ctx.used_lists):
        lists[nid] = {str(l): list(num.level(nid, l) or ("decimal", "%" + str(l + 1) + ".", 1)) for l in range(9)}
    return {"sig": _wd_sig(path, doc) if path else "", "blocks": blocks, "styles": styles, "default": ctx.styles.default_p,
            "offer": [s for s in offer if s in styles], "lists": lists, "page": _wd_page(doc.element.body)}


def word_save(src_path, plan, out_dir, assets=None):
    """The Word editor's Save: the original .docx (or none, for a new document) plus the plan → the new .docx."""
    from docx import Document
    if not isinstance(plan, dict) or not isinstance(plan.get("body"), list):
        raise WordEditError(_WD_BAD)
    if src_path:
        doc = _wd_load(src_path)
        if plan.get("sig") and plan.get("sig") != _wd_sig(src_path, doc):
            raise WordEditError("This isn't the file that was opened — please open it again.")
    else:
        doc = Document()
    ctx = _WdCtx(doc, view=False, assets=assets)
    body = doc.element.body
    original = [c for c in body]
    sect = original[-1] if original and _wd_is(original[-1], "sectPr") else None
    ctx.container(body, "b")
    follow, pending = {}, []                     # bookmarks and the like between blocks stay before the next block
    for n, c in enumerate(original):
        if c is sect:
            continue
        bid = "b%d" % n
        if bid in ctx.blocks:
            follow[bid], pending = pending, []
        else:
            pending.append(c)
    used = set()
    built = ctx.build(plan["body"], used)
    for c in original:
        body.remove(c)
    kept = set()
    for el, bid in built:
        for x in follow.get(bid, []):
            body.append(x)
        body.append(el)
        kept.add(bid)
    for bid, xs in follow.items():
        if bid not in kept:
            for x in xs:
                body.append(x)
    for x in pending:
        body.append(x)
    if not any(_wd_is(c, "p") or _wd_is(c, "tbl") for c in body):
        body.append(_wd_new("w:p"))
    if sect is not None:
        body.append(sect)
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, "edited.docx")
    doc.save(out)
    return out


CONVERTERS = {
    ("docx", "pdf"): docx_to_pdf,
    ("docx", "pptx"): docx_to_pptx,
    ("pptx", "docx"): pptx_to_docx,
    ("pdf", "docx"): pdf_to_docx,
    ("pdf", "pptx"): pdf_to_pptx,
    ("xlsx", "docx"): xlsx_to_docx,
    ("docx", "xlsx"): docx_to_xlsx,
    ("pptx", "pptx"): pptx_to_pptx,
    ("pptx", "pdf"): pptx_to_pdf,
    ("xlsx", "pdf"): xlsx_to_pdf,
}


def convert(src_path, from_fmt, to_fmt, out_dir, style=None, pptx_options=None):
    # Confirmed directly: a zero-byte upload doesn't fail at all for
    # some conversion pairs (docx->pdf via LibreOffice happily produces
    # a "successful," completely blank PDF from an empty source) —
    # arguably worse than an error, since the user has no idea their
    # own file was the problem rather than the tool. One check here
    # covers every conversion pair, since they all route through this
    # single function.
    if os.path.getsize(src_path) == 0:
        raise ConversionError("This file appears to be empty — there's nothing to convert.")
    key = (from_fmt.lower(), to_fmt.lower())
    if key not in CONVERTERS:
        raise ConversionError(f"Unsupported conversion: {from_fmt} -> {to_fmt}")
    # pptx_options is deliberately only ever forwarded for the pptx->pptx
    # ("Fix Slides") pair specifically, not passed generically to
    # whichever converter `key` happens to resolve to — every other
    # converter function has no matching keyword arguments at all, so
    # passing this through unconditionally would be a hard TypeError on
    # every other conversion pair the moment a caller supplied it.
    if key == ("pptx", "pptx") and pptx_options:
        return CONVERTERS[key](src_path, out_dir, **pptx_options)
    if style:
        return CONVERTERS[key](src_path, out_dir, style=style)
    return CONVERTERS[key](src_path, out_dir)


# ------------------------------------------------------- raw text -> PPTX
_SUMMARY_CHART_TYPES = {
    "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "line": XL_CHART_TYPE.LINE,
    "pie": XL_CHART_TYPE.PIE,
}


_SUMMARY_STYLES = {
    "visual": {
        "show_icon": True,
        "cta_accent": True,
        "bullet_prefix": "",  # relies on spacing/bold alone for hierarchy, not a bullet glyph
    },
    "plain": {
        # "Plain" means undecorated styling, not lower-quality content —
        # the 6x6 condensation, bold key-term spans, real charts, and
        # generous whitespace are genuinely good regardless of visual
        # theme, so none of that changes here. What's actually stripped
        # is decoration: no emoji icon, no color accent distinguishing
        # the CTA slide, and traditional bullet glyphs instead of relying
        # on spacing alone — closer to a plain, formal outline a reader
        # who wants the AI's text without a designed look would expect.
        "show_icon": False,
        "cta_accent": False,
        "bullet_prefix": "\u2022  ",
    },
}


def summary_slides_to_pptx(slides, out_dir, style="plain", filename="Summary.pptx"):
    """Builds the deck from Smart Summarize's slide plan. Plain White by default (white, black
    words, no icons or accents); "visual" keeps the icon-and-colour look. Sizes are fixed and big
    — headings 34 pt, sub-headings 28 pt, points 24 pt — and text is never shrunk to fit: what
    doesn't fit continues on the next slide under the same heading (its sub-heading repeated)."""
    if not slides:
        raise ConversionError("No slides to build a presentation from.")
    style_conf = _SUMMARY_STYLES.get(style, _SUMMARY_STYLES["plain"])
    visual = bool(style_conf["show_icon"])

    prs = Presentation()
    prs.slide_width = PptxInches(13.333)
    prs.slide_height = PptxInches(7.5)
    _fit_layouts_to_widescreen(prs)
    blank_layout = prs.slide_layouts[6]

    margin_x = PptxInches(0.9)
    content_width = prs.slide_width - (2 * margin_x)
    width_in = content_width / 914400.0
    head_top, head_h = PptxInches(0.6), PptxInches(1.25)
    content_top = PptxInches(2.0)
    content_height = prs.slide_height - content_top - PptxInches(0.6)
    budget = content_height / 12700.0          # points of height
    INK = PptxRGBColor(0, 0, 0)
    HEADER_COLOR = PptxRGBColor(0x1A, 0x1A, 0x1A) if visual else INK
    BODY_COLOR = PptxRGBColor(0x26, 0x26, 0x26) if visual else INK
    SUB_COLOR = PptxRGBColor(0x1F, 0x4E, 0x79) if visual else INK
    CTA_ACCENT = PptxRGBColor(0x1F, 0x6F, 0x4B)
    CTA_BG = PptxRGBColor(0xF2, 0xF9, 0xF5)
    SUB_PT, POINT_PT, PLAIN_PT = 28, 24, 24

    def lines(text, size):
        per = max(10, int(width_in * 72 / (size * 0.5)))
        plain = re.sub(r"\*+", "", str(text or ""))
        return max(1, -(-len(plain) // per))

    def item_h(kind, text):
        if kind == "sub":
            return lines(text, SUB_PT) * SUB_PT * 1.15 + 14 + 6
        return lines(text, POINT_PT) * POINT_PT * 1.15 + 8

    def pages_of(items):
        """Items packed into slides by the room they need; a sub-heading never ends a slide, and a
        section continued on the next slide gets its sub-heading again."""
        pages, cur, used, current_sub = [], [], 0.0, None
        for i, (kind, text) in enumerate(items):
            h = item_h(kind, text)
            if kind == "sub":
                current_sub = text
                nxt = items[i + 1] if i + 1 < len(items) and items[i + 1][0] == "pt" else None
                need = h + (item_h(*nxt) if nxt else 0)
                if cur and used + need > budget:
                    pages.append(cur)
                    cur, used = [], 0.0
                cur.append((kind, text))
                used += h
                continue
            if cur and used + h > budget:
                pages.append(cur)
                cur, used = [], 0.0
                if current_sub:
                    cur.append(("sub", current_sub))
                    used += item_h("sub", current_sub)
            cur.append((kind, text))
            used += h
        if cur:
            pages.append(cur)
        return pages or [[]]

    def header(slide, text, is_cta, icon):
        left = margin_x
        if icon:
            ib = slide.shapes.add_textbox(margin_x, head_top, PptxInches(1.1), PptxInches(1.1))
            ib.text_frame.word_wrap = False
            ir = ib.text_frame.paragraphs[0].add_run()
            ir.text = icon
            ir.font.size = Pt(44)
            left = margin_x + PptxInches(1.15)
        hb = slide.shapes.add_textbox(left, head_top, content_width - (left - margin_x), head_h)
        tf = hb.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        _add_markdown_aware_pptx_text(tf.paragraphs[0], text)
        for r in tf.paragraphs[0].runs:
            r.font.size = Pt(34 if len(text) <= 48 else 30)
            r.font.bold = True
            r.font.color.rgb = CTA_ACCENT if is_cta else HEADER_COLOR

    def body(slide, page, plain_bullets):
        bb = slide.shapes.add_textbox(margin_x, content_top, content_width, content_height)
        tf = bb.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE        # never shrunk: what doesn't fit went to the next slide
        first = True
        for kind, text in page:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            if kind == "sub":
                p.space_before = Pt(0 if first else 14)
                p.space_after = Pt(6)
                _add_markdown_aware_pptx_text(p, text)
                for run in p.runs:
                    run.font.size = Pt(SUB_PT)
                    run.font.bold = True
                    if run.font.color.type is None:
                        run.font.color.rgb = SUB_COLOR
            else:
                p.space_after = Pt(8)
                pPr = p._p.get_or_add_pPr()
                indent = not plain_bullets
                pPr.set("marL", str(int(PptxInches(0.55 if indent else 0.4))))
                pPr.set("indent", str(-int(PptxInches(0.32))))
                mark = p.add_run()
                mark.text = "• "
                mark.font.size = Pt(POINT_PT)
                mark.font.color.rgb = SUB_COLOR
                _add_markdown_aware_pptx_text(p, text)
                for run in p.runs:
                    if run.font.size is None:
                        run.font.size = Pt(POINT_PT)
                    if run.font.color.type is None:
                        run.font.color.rgb = BODY_COLOR
            first = False

    for slide_data in slides:
        if slide_data.get("kind") in ("cover", "thanks"):
            slide = prs.slides.add_slide(blank_layout)
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = PptxRGBColor(0x14, 0x2B, 0x4A) if visual else PptxRGBColor(0xFF, 0xFF, 0xFF)
            ink = PptxRGBColor(0xFF, 0xFF, 0xFF) if visual else INK
            head = re.sub(r"\*\*", "", str(slide_data.get("header") or "")).strip()
            sub_lines = [re.sub(r"\*\*", "", str(b)).strip() for b in (slide_data.get("bullets") or []) if str(b).strip()][:4]
            box = slide.shapes.add_textbox(margin_x, PptxInches(2.2 if sub_lines else 2.8), content_width, PptxInches(3.2))
            tf = box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            r = p.add_run()
            r.text = head
            r.font.size = Pt(48 if len(head) <= 34 else 40 if len(head) <= 60 else 32)
            r.font.color.rgb = ink
            for line in sub_lines:
                q = tf.add_paragraph()
                q.alignment = PP_ALIGN.CENTER
                q.space_before = Pt(12)
                rr = q.add_run()
                rr.text = line
                rr.font.size = Pt(24)
                rr.font.color.rgb = PptxRGBColor(0xC9, 0xD6, 0xEA) if visual else PptxRGBColor(0x26, 0x26, 0x26)
            continue

        is_cta = bool(slide_data.get("is_cta")) and style_conf["cta_accent"]
        icon = (slide_data.get("icon") or "").strip() if visual else ""
        chart = slide_data.get("chart")
        if chart:
            slide = prs.slides.add_slide(blank_layout)
            header(slide, slide_data["header"], is_cta, icon)
            chart_data = CategoryChartData()
            chart_data.categories = chart["categories"]
            chart_data.add_series(chart["series_name"], chart["values"])
            try:
                gc = slide.shapes.add_chart(_SUMMARY_CHART_TYPES[chart["type"]], margin_x, content_top, content_width, content_height, chart_data).chart
                gc.has_legend = chart["type"] == "pie"
                if gc.has_legend:
                    gc.legend.position = XL_LEGEND_POSITION.BOTTOM
                    gc.legend.include_in_layout = False
                gc.font.size = Pt(16)
            except Exception:
                _add_chart_as_table_fallback(slide, chart["categories"], [(chart["series_name"], chart["values"])], prs.slide_width, prs.slide_height)
            continue
        if slide_data.get("sections"):
            items = []
            for sec in slide_data["sections"]:
                if sec.get("subheading"):
                    items.append(("sub", sec["subheading"]))
                items.extend(("pt", t) for t in (sec.get("points") or []))
            plain_bullets = False
        else:
            items = [("pt", b) for b in (slide_data.get("bullets") or [])]
            plain_bullets = True
        for page in pages_of(items):
            slide = prs.slides.add_slide(blank_layout)
            if is_cta:
                slide.background.fill.solid()
                slide.background.fill.fore_color.rgb = CTA_BG
            header(slide, slide_data["header"], is_cta, icon)
            if page:
                body(slide, page, plain_bullets)

    result_path = os.path.join(out_dir, filename)
    prs.save(result_path)
    return result_path


def text_to_pptx(raw_text, out_dir):
    """Turn pasted plain text into a slide deck. Blocks separated by a
    blank line become slides; each block's first line is the slide title,
    remaining lines become bullets. No AI involved — pure structural rules,
    same spirit as docx_to_pptx but for text with no formatting to read."""
    blocks = [b.strip() for b in raw_text.strip().split("\n\n") if b.strip()]
    if not blocks:
        raise ConversionError("No text provided")

    prs = Presentation()
    prs.slide_width = PptxInches(13.333)
    prs.slide_height = PptxInches(7.5)
    _fit_layouts_to_widescreen(prs)  # same off-centre, 29%-empty-slide fix as Notes → Slides
    title_layout = prs.slide_layouts[1]
    MAX_BULLETS_PER_SLIDE = 8

    for block in blocks:
        lines = [l.strip() for l in block.split("\n") if l.strip()]
        if not lines:
            continue
        title = _BULLET_RE.sub("", lines[0]).strip()
        body_lines = [_BULLET_RE.sub("", l).strip() for l in lines[1:]]

        def set_slide_title(slide, text):
            tf = slide.shapes.title.text_frame
            tf.clear()
            _add_markdown_aware_pptx_text(tf.paragraphs[0], text[:120])

        slide = None
        tf = None
        count = 0
        for i, line in enumerate(body_lines):
            if slide is None or count >= MAX_BULLETS_PER_SLIDE:
                slide_title = title   # continuation slides keep the original title, nothing added
                slide = prs.slides.add_slide(title_layout)
                set_slide_title(slide, slide_title)
                tf = slide.placeholders[1].text_frame
                tf.clear()
                count = 0
            p = tf.paragraphs[0] if count == 0 else tf.add_paragraph()
            _add_markdown_aware_pptx_text(p, line)
            count += 1
        if slide is None:
            slide = prs.slides.add_slide(title_layout)
            set_slide_title(slide, title)
            slide.placeholders[1].text_frame.clear()

    if not prs.slides:
        raise ConversionError("No usable text found")

    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "Presentation.pptx")
    prs.save(out_path)
    return out_path


def _set_run_font(run, bold=False, italic=False):
    run.font.name = "Times New Roman"
    run.font.size = DocxPt(12)
    run.font.color.rgb = DocxRGBColor(0, 0, 0)
    run.bold = bold
    run.italic = italic
    # Word can silently fall back to a different font for East-Asian text
    # runs unless this is set explicitly alongside the Latin font name above.
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), "Times New Roman")


_MARKDOWN_EMPHASIS_RE = re.compile(
    r"\*\*\*(\S(?:.*?\S)?)\*\*\*"           # ***bold italic*** — must be tried before
    r"|(?<!\w)___(\S(?:.*?\S)?)___(?!\w)"   # ** or *, or a double-marker alternative
    r"|\*\*(\S(?:.*?\S)?)\*\*"              # would consume only 2 of the 3 leading
    r"|(?<!\w)__(\S(?:.*?\S)?)__(?!\w)"     # markers and leave a stray one behind as
    r"|\*(\S(?:.*?\S)?)\*"                  # literal text in the output.
    r"|(?<!\w)_(\S(?:.*?\S)?)_(?!\w)"       # *italic* — content can't start/end with
)                                            # whitespace, so '3 * 4 = 12 and separately
                                              # 5 * 6' (academic multiplication notation)
                                              # isn't mistaken for italic markup.
                                              # _italic_ additionally requires a word
                                              # boundary outside each underscore — the
                                              # same rule CommonMark itself uses — so a
                                              # variable name like 'user_id' is never
                                              # mistaken for emphasis and corrupted.


_XML_ILLEGAL_CHARS_RE = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff]")

# ---------------------------------------------- PPTX sentence-splitting
# Long, multi-sentence source paragraphs read as a dense wall of text once
# dropped onto a slide — the single biggest reason a converted deck still
# needed manual cleanup before it was presentation-ready. These split such
# paragraphs into separate, shorter bullets at real sentence boundaries.
_SENTENCE_SPLIT_ABBREVIATIONS = {
    "mr", "mrs", "ms", "dr", "prof", "rev", "sen", "rep", "gen", "st", "ave",
    "jr", "sr", "vs", "etc", "eg", "ie", "us", "uk", "am", "pm", "no", "vol",
    "fig", "approx", "inc", "corp", "co", "ltd", "dept",
}
_SENTENCE_SPLIT_RE = re.compile(r'([.!?]["\')\]]*)\s+(?=[A-Z"\'(])')
_SENTENCE_TRAILING_END_RE = re.compile(r'[.!?]["\')\]]*\s*$')
_STARTS_NEW_SENTENCE_RE = re.compile(r'^\s*[A-Z"\'(]')
_CHARS_PER_LINE_ESTIMATE = 48  # calibrated against rendered output at 26pt on a 16:9 content placeholder


def _estimate_line_count(text):
    """A flat bullet-count cap doesn't account for how much a given
    bullet's text will actually wrap — six short bullets fit comfortably
    on a slide, but six two-to-three-line bullets (a realistic outcome
    once long paragraphs are being split into sentence-level bullets)
    can overflow it even though the count is the same. This estimates
    wrapped line count from character length as a proportional-font
    approximation; it doesn't need to be exact, just conservative enough
    to trigger a "(cont.)" slide before real overflow happens rather
    than after."""
    if not text.strip():
        return 1
    return max(1, -(-len(text) // _CHARS_PER_LINE_ESTIMATE))  # ceiling division

_SENTENCE_SPLIT_WORD_THRESHOLD = 22  # below this, even a 2-sentence paragraph reads fine as one bullet
_MIN_WORDS_PER_SPLIT_BULLET = 4      # avoids spinning off an awkward, near-empty trailing fragment


def _split_runs_at_sentence_boundaries(runs_data):
    """runs_data is a list of (text, *metadata) tuples — one per source
    run, in order, all belonging to a single source paragraph. Returns a
    list of groups (each itself a list of the same kind of tuples); each
    group becomes one bullet. Splitting operates on run text directly
    rather than the paragraph's plain-text concatenation, so a run's own
    formatting metadata carries over correctly to both sides of a split
    (a run split mid-sentence produces two runs sharing the same
    bold/italic/color, not a loss of formatting). Hyperlink runs are
    never split internally — a link's display text spans one run, and
    breaking it across two bullets would split the link itself.

    Checks both within a single run's text AND across a run boundary —
    a formatting change (bold, italic starting) very often lands exactly
    at a sentence boundary, so "sentence ends in run N, next sentence
    starts in run N+1" is at least as common as an in-run split and is
    not just an edge case to shrug off."""
    groups = []
    current_group = []
    current_word_count = 0
    n = len(runs_data)
    for idx, entry in enumerate(runs_data):
        text = entry[0]
        is_link = entry[4]
        if is_link:
            current_group.append(entry)
            current_word_count += len(text.split())
            continue
        pos = 0
        for m in _SENTENCE_SPLIT_RE.finditer(text):
            split_at = m.end(1)
            preceding = text[pos:split_at - 1]
            words_before = re.split(r"\s+", preceding.strip())
            last_word = words_before[-1].lower().rstrip(".") if preceding.strip() else ""
            if last_word in _SENTENCE_SPLIT_ABBREVIATIONS:
                continue
            chunk = text[pos:split_at]
            if not current_group:
                chunk = chunk.lstrip()
            if chunk:
                current_group.append((chunk,) + entry[1:])
                current_word_count += len(chunk.split())
            pos = split_at
            if current_word_count >= _MIN_WORDS_PER_SPLIT_BULLET and current_group:
                groups.append(current_group)
                current_group = []
                current_word_count = 0
        remainder = text[pos:]
        if not current_group:
            remainder = remainder.lstrip()
        if remainder:
            current_group.append((remainder,) + entry[1:])
            current_word_count += len(remainder.split())
        # Cross-run boundary: this run's remainder ends a sentence, and
        # the NEXT run starts a new one.
        if remainder and _SENTENCE_TRAILING_END_RE.search(remainder):
            trailing_words = re.split(r"\s+", remainder.strip())
            last_word = trailing_words[-1].lower().rstrip(".") if trailing_words else ""
            if last_word not in _SENTENCE_SPLIT_ABBREVIATIONS and idx + 1 < n:
                next_text, next_is_link = runs_data[idx + 1][0], runs_data[idx + 1][4]
                if not next_is_link and _STARTS_NEW_SENTENCE_RE.match(next_text) and \
                        current_word_count >= _MIN_WORDS_PER_SPLIT_BULLET and current_group:
                    # trim the trailing space this run's text carried
                    # before the next sentence, matching the leading-space
                    # trim already applied when a new group starts
                    last_text, last_meta = current_group[-1][0], current_group[-1][1:]
                    current_group[-1] = (last_text.rstrip(),) + last_meta
                    groups.append(current_group)
                    current_group = []
                    current_word_count = 0
    if current_group:
        if groups and current_word_count < _MIN_WORDS_PER_SPLIT_BULLET:
            groups[-1].extend(current_group)
        else:
            groups.append(current_group)
    return groups if len(groups) > 1 else [runs_data]


def _sanitize_xml_text(text):
    """Strips characters that are illegal in XML 1.0 (control characters
    other than tab/newline/carriage return, which are fine) — python-docx
    raises immediately when text containing one of these is assigned, not
    just at save time, and these are a real possibility in AI-generated
    text (an encoding artifact, a copy-paste quirk), not just a
    theoretical edge case. Confirmed directly: a bare null byte or a
    vertical-tab character in otherwise normal text crashes the core
    essay-generation feature outright without this."""
    return _XML_ILLEGAL_CHARS_RE.sub("", text)


def _add_markdown_aware_text(paragraph, text, base_bold=False):
    """Adds text to a paragraph, converting basic markdown emphasis
    (***bold italic***, **bold**, __bold__, *italic*, _italic_) into real
    Word formatting instead of leaving literal asterisks/underscores in
    the output. The text here is AI-generated — LLMs commonly reach for
    markdown emphasis as a natural writing habit even when nothing asked
    for markdown specifically, and literal '**word**' in a finished
    academic Word document reads as broken, not as emphasis. base_bold
    lets a heading's own bold styling combine correctly with an *italic*
    span inside it, rather than the emphasis parsing accidentally
    clearing it."""
    text = _sanitize_xml_text(text)
    pos = 0
    for m in _MARKDOWN_EMPHASIS_RE.finditer(text):
        if m.start() > pos:
            _set_run_font(paragraph.add_run(text[pos:m.start()]), bold=base_bold)
        both_inner = m.group(1) if m.group(1) is not None else m.group(2)
        bold_inner = m.group(3) if m.group(3) is not None else m.group(4)
        if both_inner is not None:
            _set_run_font(paragraph.add_run(both_inner), bold=True, italic=True)
        elif bold_inner is not None:
            _set_run_font(paragraph.add_run(bold_inner), bold=True)
        else:
            italic_inner = m.group(5) if m.group(5) is not None else m.group(6)
            _set_run_font(paragraph.add_run(italic_inner), bold=base_bold, italic=True)
        pos = m.end()
    if pos < len(text):
        _set_run_font(paragraph.add_run(text[pos:]), bold=base_bold)


_PPTX_URL_RE = re.compile(r"https?://[^\s<>\"]+[^\s<>\".,;:)\]]")


def _link_urls_in_runs(paragraph):
    """Web addresses in a paragraph's runs become live links (each in a run of its own)."""
    for r in list(paragraph.runs):
        text = r.text or ""
        if not _PPTX_URL_RE.search(text):
            continue
        pieces, pos = [], 0
        for m in _PPTX_URL_RE.finditer(text):
            if m.start() > pos:
                pieces.append((text[pos:m.start()], None))
            pieces.append((m.group(0), m.group(0)))
            pos = m.end()
        if pos < len(text):
            pieces.append((text[pos:], None))
        anchor = r._r
        r.text = pieces[0][0]
        if pieces[0][1]:
            r.hyperlink.address = pieces[0][1]
        for words, url in pieces[1:]:
            new = copy.deepcopy(anchor)
            for h in new.findall(".//" + pptx_qn("a:hlinkClick")):
                h.getparent().remove(h)
            anchor.addnext(new)
            nr = _Run(new, paragraph)
            nr.text = words
            if url:
                nr.hyperlink.address = url
            anchor = new


def _add_markdown_aware_pptx_text(paragraph, text):
    """The pptx equivalent of _add_markdown_aware_text — same shared
    regex (it's pure text pattern matching, format-agnostic), building
    pptx runs instead of docx ones. Used for pasted text that may itself
    contain markdown emphasis syntax — plausible whenever someone copies
    content from a notes app, a README, or an AI chat response into the
    paste-to-slides feature, not just for AI-authored text."""
    pos = 0
    for m in _MARKDOWN_EMPHASIS_RE.finditer(text):
        if m.start() > pos:
            r = paragraph.add_run()
            r.text = text[pos:m.start()]
        both_inner = m.group(1) if m.group(1) is not None else m.group(2)
        bold_inner = m.group(3) if m.group(3) is not None else m.group(4)
        if both_inner is not None:
            r = paragraph.add_run()
            r.text = both_inner
            r.font.bold = True
            r.font.italic = True
        elif bold_inner is not None:
            r = paragraph.add_run()
            r.text = bold_inner
            r.font.bold = True
        else:
            italic_inner = m.group(5) if m.group(5) is not None else m.group(6)
            r = paragraph.add_run()
            r.text = italic_inner
            r.font.italic = True
        pos = m.end()
    if pos < len(text):
        r = paragraph.add_run()
        r.text = text[pos:]
    _link_urls_in_runs(paragraph)


def _add_page_number_footer(doc):
    """Adds a centered auto-updating page number to the footer — standard
    for academic documents."""
    footer = doc.sections[0].footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run()
    run.font.name = "Times New Roman"
    run.font.size = DocxPt(12)
    run.font.color.rgb = DocxRGBColor(0, 0, 0)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)


# ------------------------------------------------ 3.0: Word downloads like the preview
# A Word file built from Remy's or the writing tool's answer: APA look, real headings, a real
# clickable Table of Contents / List of Tables / List of Figures (with page numbers), real
# tables with captions, in-text citations linked to their references, live reference links.
_DOC_FONT = "Times New Roman"
_LINK_BLUE = "0563C1"
_MD_TABLE_ROW = re.compile(r"^\s*\|.*\|\s*$")
_MD_TABLE_SEP = re.compile(r"^\s*\|?\s*:?-{2,}:?\s*(?:\|\s*:?-{2,}:?\s*)*\|?\s*$")
_MD_CAPTION = re.compile(r"^\s*(?:\*\*|__)?\s*(table|figure)\s*(\d+)?\s*(?:\*\*|__)?\s*[:.\-–—]?\s*(?:\*\*|__)?\s*(.*?)\s*(?:\*\*|__)?\s*$", re.I)
_MD_IMAGE = re.compile(r"^\s*!\[([^\]]*)\]\((data:image/(?:png|jpe?g);base64,[A-Za-z0-9+/=]+)\)\s*$")
_MD_HEADING_LINE = re.compile(r"^(#{1,6})\s+(.+?)\s*#*\s*$")
_MD_BOLD_HEADING = re.compile(r"^\*\*([^*].{0,110}?)\*\*:?\s*$")
_MD_BULLET_LINE = re.compile(r"^\s*[-*•]\s+(.+)$")
_MD_NUMBER_LINE = re.compile(r"^\s*\d{1,3}[.)]\s+(.+)$")


def _emphasis_runs(text):
    """[(text, bold, italic)] for a line of markdown."""
    text = _sanitize_xml_text(str(text or ""))
    out, pos = [], 0
    for m in _MARKDOWN_EMPHASIS_RE.finditer(text):
        if m.start() > pos:
            out.append((text[pos:m.start()], False, False))
        both = m.group(1) if m.group(1) is not None else m.group(2)
        bold = m.group(3) if m.group(3) is not None else m.group(4)
        if both is not None:
            out.append((both, True, True))
        elif bold is not None:
            out.append((bold, True, False))
        else:
            out.append((m.group(5) if m.group(5) is not None else m.group(6), False, True))
        pos = m.end()
    if pos < len(text):
        out.append((text[pos:], False, False))
    return [r for r in out if r[0]]


def _xrun(text, bold=False, italic=False, color=None, underline=None, size=12):
    """One run of text, in the document's font."""
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        fonts.set(qn(a), _DOC_FONT)
    rpr.append(fonts)
    if bold:
        rpr.append(OxmlElement("w:b"))
    if italic:
        rpr.append(OxmlElement("w:i"))
    c = OxmlElement("w:color")
    c.set(qn("w:val"), color or "000000")
    rpr.append(c)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(int(size * 2)))
    rpr.append(sz)
    if underline:
        u = OxmlElement("w:u")
        u.set(qn("w:val"), underline)
        rpr.append(u)
    r.append(rpr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def _xfield(kind):
    r = OxmlElement("w:r")
    f = OxmlElement("w:fldChar")
    f.set(qn("w:fldCharType"), kind)
    r.append(f)
    return r


def _xinstr(text):
    r = OxmlElement("w:r")
    t = OxmlElement("w:instrText")
    t.set(qn("xml:space"), "preserve")
    t.text = f" {text} "
    r.append(t)
    return r


def _xbookmark(p_el, name, bid):
    """A bookmark around a whole paragraph (a heading, a caption, a reference)."""
    s = OxmlElement("w:bookmarkStart")
    s.set(qn("w:id"), str(bid))
    s.set(qn("w:name"), name)
    e = OxmlElement("w:bookmarkEnd")
    e.set(qn("w:id"), str(bid))
    ppr = p_el.find(qn("w:pPr"))
    if ppr is not None:
        ppr.addnext(s)
    else:
        p_el.insert(0, s)
    p_el.append(e)


class _CiteLinks:
    """Where the in-text citations are in a piece of text, and which reference each points at:
    "(Mensah & Boateng, 2023)" (also inside a group, or narrative "Mensah and Boateng (2023)"),
    or Remy's numbers "[1]", "[2, 3]"."""

    def __init__(self, references, numbered):
        self.numbered = numbered
        self.pats = []
        if numbered:
            self.count = len(references)
            return
        for i, r in enumerate(references):
            key = str(r.get("key") or "").strip()
            if not (key.startswith("(") and key.endswith(")")):
                continue
            inner = key[1:-1].strip()
            esc = re.escape(inner).replace("\\&", "(?:&|and)").replace("\\ ", r"\s+")
            self.pats.append((re.compile(r"(?<![\w])" + esc + r"(?![\w])"), f"ref_{i + 1}"))
            name, _, year = inner.rpartition(",") if "," in inner else inner.rpartition(" ")
            name, year = name.strip(), year.strip()
            if name and year:
                n_esc = re.escape(name).replace("\\&", "(?:&|and)").replace("\\ ", r"\s+")
                self.pats.append((re.compile(r"(?<![\w])" + n_esc + r"\s*\(" + re.escape(year) + r"\)"), f"ref_{i + 1}"))

    def spans(self, text):
        found = []
        if self.numbered:
            for m in re.finditer(r"\[(\d{1,3}(?:\s*[,–\-]\s*\d{1,3})*)\]", text):
                for n in re.finditer(r"\d{1,3}", m.group(1)):
                    k = int(n.group(0))
                    if 1 <= k <= self.count:
                        s = m.start(1) + n.start()
                        found.append((s, s + len(n.group(0)), f"ref_{k}"))
        else:
            for pat, anchor in self.pats:
                for m in pat.finditer(text):
                    found.append((m.start(), m.end(), anchor))
        found.sort(key=lambda x: (x[0], -(x[1] - x[0])))
        out, last = [], -1
        for s, e, a in found:
            if s >= last:
                out.append((s, e, a))
                last = e
        return out


def _write_rich(p, text, links=None, bold=False, italic=False):
    """Markdown emphasis as Word formatting; each in-text citation a link to its reference."""
    for seg, b, it in _emphasis_runs(text):
        b, it = b or bold, it or italic
        pos = 0
        for s, e, anchor in (links.spans(seg) if links else []):
            if s > pos:
                p._p.append(_xrun(seg[pos:s], b, it))
            h = OxmlElement("w:hyperlink")
            h.set(qn("w:anchor"), anchor)
            h.set(qn("w:history"), "1")
            h.append(_xrun(seg[s:e], b, it, underline="dotted"))
            p._p.append(h)
            pos = e
        if pos < len(seg):
            p._p.append(_xrun(seg[pos:], b, it))


def _write_url(p, url):
    """A live link (a reference's DOI or PDF), blue and underlined as Word shows links."""
    from docx.opc.constants import RELATIONSHIP_TYPE as _RT
    rid = p.part.relate_to(url, _RT.HYPERLINK, is_external=True)
    h = OxmlElement("w:hyperlink")
    h.set(qn("r:id"), rid)
    h.set(qn("w:history"), "1")
    h.append(_xrun(url, color=_LINK_BLUE, underline="single"))
    p._p.append(h)


def _apa_styles(doc):
    """Normal and Heading 1-3 in APA's look (theme fonts and colours taken off, so they hold)."""
    def plain(st, bold=False, italic=False):
        st.font.name = _DOC_FONT
        st.font.size = DocxPt(12)
        st.font.bold = bold
        st.font.italic = italic
        st.font.color.rgb = DocxRGBColor(0, 0, 0)
        rpr = st.element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.insert(0, rf)
        for a in list(rf.attrib):
            if a.endswith("Theme") or a.endswith("theme"):
                del rf.attrib[a]
        for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rf.set(qn(a), _DOC_FONT)
        col = rpr.find(qn("w:color"))
        if col is not None:
            for a in list(col.attrib):
                if "theme" in a:
                    del col.attrib[a]
        pf = st.paragraph_format
        pf.space_before = DocxPt(0)
        pf.space_after = DocxPt(0)
        pf.line_spacing = 2.0
    plain(doc.styles["Normal"])
    for lvl in (1, 2, 3):
        st = doc.styles[f"Heading {lvl}"]
        plain(st, bold=True, italic=(lvl == 3))
        st.paragraph_format.keep_with_next = True
        st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER if lvl == 1 else WD_ALIGN_PARAGRAPH.LEFT
    try:
        cap = doc.styles["Caption"]
        plain(cap)
        cap.paragraph_format.keep_with_next = True
    except KeyError:
        pass


def _md_blocks(text, skip_heading=None):
    """The answer's markdown as blocks: ("h", level, text) headings, ("p", text) paragraphs,
    ("li", text, numbered) list items, ("table", caption, rows), ("img", caption, bytes)."""
    lines = (text or "").splitlines()
    blocks, buf = [], []
    skip = (skip_heading or "").strip().lower()
    pending_caption = None

    def flush():
        if buf:
            blocks.append(("p", " ".join(buf)))
            buf.clear()

    i = 0
    while i < len(lines):
        raw = lines[i]
        line = raw.strip()
        if not line:
            flush()
            i += 1
            continue
        # a table: | a | b | then |---|---|
        if _MD_TABLE_ROW.match(line) and i + 1 < len(lines) and _MD_TABLE_SEP.match(lines[i + 1].strip()):
            flush()
            rows = []
            j = i
            while j < len(lines) and _MD_TABLE_ROW.match(lines[j].strip()):
                if not _MD_TABLE_SEP.match(lines[j].strip()):
                    cells = [c.strip() for c in lines[j].strip().strip("|").split("|")]
                    rows.append(cells)
                j += 1
            width = max(len(r) for r in rows) if rows else 0
            rows = [r + [""] * (width - len(r)) for r in rows]
            if rows:
                blocks.append(("table", pending_caption, rows))
            pending_caption = None
            i = j
            continue
        m = _MD_IMAGE.match(line)
        if m:
            flush()
            try:
                import base64 as _b64
                data = _b64.b64decode(m.group(2).split(",", 1)[1])
                blocks.append(("img", pending_caption or m.group(1).strip() or None, data))
            except Exception:
                pass
            pending_caption = None
            i += 1
            continue
        m = _MD_CAPTION.match(line)
        nxt = next((lines[k].strip() for k in range(i + 1, min(i + 3, len(lines))) if lines[k].strip()), "")
        if m and len(line) <= 160 and (_MD_TABLE_ROW.match(nxt) or _MD_IMAGE.match(nxt)):
            flush()
            pending_caption = (m.group(3) or "").strip(" *_") or None
            i += 1
            continue
        m = _MD_HEADING_LINE.match(line)
        if m:
            flush()
            h = re.sub(r"[*_`]+", "", m.group(2)).strip()
            if skip and h.lower() == skip:
                skip = ""          # the title repeated as a heading — already at the top
            else:
                n = len(m.group(1))
                blocks.append(("h", 1 if n <= 2 else (2 if n == 3 else 3), h))
            i += 1
            continue
        m = _MD_BOLD_HEADING.match(line)
        if m and len(m.group(1).split()) <= 12:
            flush()
            blocks.append(("h", 1, m.group(1).strip()))
            i += 1
            continue
        m = _MD_BULLET_LINE.match(raw)
        if m:
            flush()
            blocks.append(("li", m.group(1).strip(), False))
            i += 1
            continue
        m = _MD_NUMBER_LINE.match(raw)
        if m:
            flush()
            blocks.append(("li", m.group(1).strip(), True))
            i += 1
            continue
        buf.append(line)
        i += 1
    flush()
    return blocks


def _norm_words(s):
    s = str(s or "").replace("ﬁ", "fi").replace("ﬂ", "fl").replace("­", "")
    s = re.sub(r"[‘’“”'\"`*_]", "", s)
    return re.sub(r"\s+", " ", s).strip().lower()


def _laid_out_pages(docx_path, work_dir, headings, captions):
    """The page each heading and caption lands on, from LibreOffice's layout of the document:
    headings by the PDF's outline, captions by their words (searched after the contents)."""
    pdf = _soffice_convert(docx_path, "pdf", work_dir)
    reader = pypdf.PdfReader(pdf)
    flat = []

    def walk(items):
        for it in items:
            if isinstance(it, list):
                walk(it)
            else:
                try:
                    flat.append((_norm_words(it.title), reader.get_destination_page_number(it)))
                except Exception:
                    pass
    try:
        walk(reader.outline)
    except Exception:
        pass
    pages, k = {}, 0
    for bm, text in headings:
        t = _norm_words(text)[:60]
        for j in range(k, len(flat)):
            if flat[j][0][:60] == t:
                pages[bm] = flat[j][1] + 1
                k = j + 1
                break
    if captions:
        texts = []
        for pg in reader.pages:
            try:
                texts.append(_norm_words(pg.extract_text() or ""))
            except Exception:
                texts.append("")
        first_heading = min(pages.values()) - 1 if pages else None
        if first_heading is None:
            fm = [i for i, t in enumerate(texts) if re.search(r"(?:^| )(?:contents|list of tables|list of figures)(?: |$)", t)]
            first_heading = (max(fm) + 1) if fm else 0
        for bm, label, title in captions:
            want = _norm_words(f"{label} {title}")[:40]
            for i in range(first_heading, len(texts)):
                if want and want in texts[i]:
                    pages[bm] = i + 1
                    break
    return pages


def _build_word(payload, pages):
    """The document itself; pages maps each heading/caption bookmark to its page number."""
    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = DocxInches(1)
    _add_page_number_footer(doc)
    _apa_styles(doc)
    numbered = bool(payload.get("numbered"))
    references = [r for r in (payload.get("references") or []) if isinstance(r, dict) and (r.get("text") or r.get("url"))]
    if not numbered:
        # APA's reference list is alphabetical; the keys stay with their entries
        references = sorted(references, key=lambda r: re.sub(r"[*_]", "", str(r.get("text") or "")).lower())
    links = _CiteLinks(references, numbered) if references else None
    title = re.sub(r"\s+", " ", str(payload.get("title") or "Academic Response")).strip()[:200]

    # ---- the body as blocks
    body = []
    sections = payload.get("sections")
    if sections and isinstance(sections, list) and any(isinstance(s, dict) for s in sections):
        for s in sections:
            if not isinstance(s, dict):
                continue
            head = f"{s.get('number', '')} {s.get('heading', '')}".strip()
            if head:
                try:
                    lvl = max(1, min(int(s.get("level") or 1), 3))
                except (TypeError, ValueError):
                    lvl = 1
                body.append(("h", lvl, head))
            body.extend(_md_blocks(s.get("text") or ""))
    else:
        body = _md_blocks(str(payload.get("text") or "").strip(), skip_heading=title)

    # numbers and bookmarks for headings, tables and figures
    heads, tables, figures = [], [], []
    current_head = None
    for k, b in enumerate(body):
        if b[0] == "h":
            bm = f"_Toc{100000 + len(heads)}"
            heads.append((bm, b[2], b[1]))
            body[k] = ("h", b[1], b[2], bm)
            current_head = b[2]
        elif b[0] == "table":
            n = len(tables) + 1
            cap = b[1] or (current_head or "Summary")
            bm = f"_Tbl{n}"
            tables.append((bm, n, cap))
            body[k] = ("table", cap, b[2], n, bm)
        elif b[0] == "img":
            n = len(figures) + 1
            cap = b[1] or (current_head or "Illustration")
            bm = f"_Fig{n}"
            figures.append((bm, n, cap))
            body[k] = ("img", cap, b[2], n, bm)
    want_toc = len(heads) >= 3
    bid = [1]

    def next_id():
        bid[0] += 1
        return bid[0]

    def para(align=None, indent_first=False, spacing=2.0):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = spacing
        if align is not None:
            p.alignment = align
        if indent_first:
            p.paragraph_format.first_line_indent = DocxInches(0.5)
        return p

    # ---- the title
    tp = para(WD_ALIGN_PARAGRAPH.CENTER)
    _write_rich(tp, title, bold=True)

    # ---- contents, list of tables, list of figures: real Word fields, already filled in
    def contents(heading, entries, instr):
        hp = para(WD_ALIGN_PARAGRAPH.CENTER)
        hp._p.append(_xrun(heading, bold=True))
        first = None
        for n, (bm, text, level) in enumerate(entries):
            p = para(spacing=2.0)
            pf = p.paragraph_format
            pf.left_indent = DocxInches(0.25 * (level - 1)) if level > 1 else DocxInches(0)
            from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
            pf.tab_stops.add_tab_stop(DocxInches(6.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
            if n == 0:
                p._p.append(_xfield("begin"))
                p._p.append(_xinstr(instr))
                p._p.append(_xfield("separate"))
                first = p
            h = OxmlElement("w:hyperlink")
            h.set(qn("w:anchor"), bm)
            h.set(qn("w:history"), "1")
            for seg, b, it in _emphasis_runs(text):
                h.append(_xrun(seg, b, it))
            tab = OxmlElement("w:r")
            tab.append(OxmlElement("w:tab"))
            h.append(tab)
            h.append(_xfield("begin"))
            h.append(_xinstr(f"PAGEREF {bm} \\h"))
            h.append(_xfield("separate"))
            h.append(_xrun(str(pages.get(bm, ""))))
            h.append(_xfield("end"))
            p._p.append(h)
        if first is not None:
            endp = para(spacing=1.0)
            endp._p.append(_xfield("end"))

    front = False
    if want_toc:
        contents("Contents", [(bm, text, lvl) for bm, text, lvl in heads], 'TOC \\o "1-3" \\h \\z \\u')
        front = True
    if tables:
        contents("List of Tables", [(bm, f"Table {n} {cap}", 1) for bm, n, cap in tables], 'TOC \\h \\z \\c "Table"')
        front = True
    if figures:
        contents("List of Figures", [(bm, f"Figure {n} {cap}", 1) for bm, n, cap in figures], 'TOC \\h \\z \\c "Figure"')
        front = True
    if front:
        from docx.enum.text import WD_BREAK
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ---- the body
    def caption(label, n, cap, bm):
        p = doc.add_paragraph(style="Caption")
        p.paragraph_format.line_spacing = 2.0
        p._p.append(_xrun(f"{label} ", bold=True))
        p._p.append(_xfield("begin"))
        p._p.append(_xinstr(f"SEQ {label} \\* ARABIC"))
        p._p.append(_xfield("separate"))
        p._p.append(_xrun(str(n), bold=True))
        p._p.append(_xfield("end"))
        br = OxmlElement("w:r")
        br.append(OxmlElement("w:br"))
        p._p.append(br)
        for seg, b, it in _emphasis_runs(cap):
            p._p.append(_xrun(seg, b, True))
        _xbookmark(p._p, bm, next_id())

    for b in body:
        kind = b[0]
        if kind == "h":
            _, lvl, text, bm = b
            p = doc.add_paragraph(style=f"Heading {lvl}")
            _write_rich(p, text, links, bold=True, italic=(lvl == 3))
            _xbookmark(p._p, bm, next_id())
        elif kind == "p":
            p = para(WD_ALIGN_PARAGRAPH.JUSTIFY, indent_first=True)
            _write_rich(p, b[1], links)
        elif kind == "li":
            p = doc.add_paragraph(style="List Number" if b[2] else "List Bullet")
            p.paragraph_format.line_spacing = 2.0
            _write_rich(p, b[1], links)
        elif kind == "table":
            _, cap, rows, n, bm = b
            caption("Table", n, cap, bm)
            t = doc.add_table(rows=len(rows), cols=len(rows[0]))
            borders = OxmlElement("w:tblBorders")
            for edge, val in (("top", "single"), ("left", "nil"), ("bottom", "single"), ("right", "nil"), ("insideH", "nil"), ("insideV", "nil")):
                e = OxmlElement(f"w:{edge}")
                e.set(qn("w:val"), val)
                if val == "single":
                    e.set(qn("w:sz"), "8")
                    e.set(qn("w:color"), "000000")
                borders.append(e)
            t._tbl.tblPr.append(borders)
            for ri, row in enumerate(rows):
                trpr = t.rows[ri]._tr.get_or_add_trPr()      # a row never breaks across pages
                trpr.append(OxmlElement("w:cantSplit"))
                for ci, cell_text in enumerate(row):
                    cell = t.cell(ri, ci)
                    cp = cell.paragraphs[0]
                    cp.paragraph_format.line_spacing = 1.0
                    cp.paragraph_format.space_before = DocxPt(3)
                    cp.paragraph_format.space_after = DocxPt(3)
                    cp.paragraph_format.keep_with_next = ri < len(rows) - 1   # the table stays together
                    _write_rich(cp, cell_text, links, bold=(ri == 0))
                    if ri == 0:      # a line under the column headings, as APA tables have
                        tcpr = cell._tc.get_or_add_tcPr()
                        tb = OxmlElement("w:tcBorders")
                        bt = OxmlElement("w:bottom")
                        bt.set(qn("w:val"), "single")
                        bt.set(qn("w:sz"), "8")
                        bt.set(qn("w:color"), "000000")
                        tb.append(bt)
                        tcpr.append(tb)
            para(spacing=1.0)      # a little room after the table
        elif kind == "img":
            _, cap, data, n, bm = b
            caption("Figure", n, cap, bm)
            try:
                doc.add_paragraph().add_run().add_picture(io.BytesIO(data), width=DocxInches(6.0))
            except Exception:
                pass

    # ---- the references: each one a place the citations link to, with its live link
    if references:
        doc.add_paragraph()
        rh = para(WD_ALIGN_PARAGRAPH.CENTER)
        rh._p.append(_xrun(str(payload.get("references_title") or "References")[:40], bold=True))
        for i, ref in enumerate(references):
            p = para()
            p.paragraph_format.left_indent = DocxInches(0.5)
            p.paragraph_format.first_line_indent = DocxInches(-0.5)
            text = str(ref.get("text") or "").strip()
            if numbered:
                p._p.append(_xrun(f"[{i + 1}] "))
            if text:
                _write_rich(p, text + " ")
            url = str(ref.get("url") or "").strip()
            if url.startswith("http"):
                _write_url(p, url)
            _xbookmark(p._p, f"ref_{i + 1}", next_id())

    return doc, [(bm, text) for bm, text, _ in heads], [(bm, "Table", n, cap) for bm, n, cap in tables] + [(bm, "Figure", n, cap) for bm, n, cap in figures]


def _ask_word_to_update(doc):
    """Word refreshes the contents' page numbers when the file opens (used only when they
    couldn't be worked out here)."""
    st = doc.settings.element
    if st.find(qn("w:updateFields")) is None:
        u = OxmlElement("w:updateFields")
        u.set(qn("w:val"), "true")
        st.append(u)


def academic_essay_to_docx(payload, out_dir):
    """The Word file for a written answer: APA look (Times New Roman 12, black, double-spaced,
    1-inch margins, justified paragraphs with a first-line indent, page numbers), real
    headings; a real clickable Table of Contents / List of Tables / List of Figures when the
    piece has them (page numbers worked out from the laid-out document); tables with
    captions; in-text citations linked to their references; live reference links."""
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "Academic_Response.docx")
    doc, heads, caps = _build_word(payload, {})
    fronted = len(heads) >= 3 or bool(caps)
    if fronted:
        pages = {}
        try:
            probe = os.path.join(out_dir, "layout")
            os.makedirs(probe, exist_ok=True)
            first = os.path.join(probe, "probe.docx")
            doc.save(first)
            pages = _laid_out_pages(first, probe, heads,
                                    [(bm, label, f"{n} {cap}") for bm, label, n, cap in caps])
        except Exception as e:
            import logging as _logging
            _logging.getLogger("academic-home").warning(f"Contents page numbers not worked out: {e}")
        doc, heads, caps = _build_word(payload, pages)
        wanted = [bm for bm, _ in heads] if len(heads) >= 3 else []
        wanted += [bm for bm, _, _, _ in caps]
        if any(bm not in pages for bm in wanted):
            _ask_word_to_update(doc)
    doc.save(out_path)
    return out_path


def quiz_to_docx(title, questions, out_dir, filename="Practice_Quiz.docx"):
    """Builds a polished, printable practice quiz from AI-generated
    question data — deliberately styled differently from
    academic_essay_to_docx (Calibri, single-spaced, generous blank space
    for handwritten answers) since a worksheet a student writes on and
    a formal essay serve completely different purposes. No AI involved
    here — pure deterministic formatting of content the caller already
    generated and validated.

    questions is a list of dicts, each with:
      "type": "multiple_choice" or "short_answer"
      "question": the question text
      "options": list of answer choices (multiple_choice only)
      "correct_answer": the correct option text, or a model short answer
      "explanation": a brief note on why, shown only in the answer key
    """
    if not questions:
        raise ConversionError("No questions were generated.")

    doc = Document()
    section = doc.sections[0]
    section.left_margin = DocxInches(1)
    section.right_margin = DocxInches(1)
    section.top_margin = DocxInches(1)
    section.bottom_margin = DocxInches(1)
    _add_page_number_footer(doc)

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = DocxPt(11)
    normal.font.color.rgb = DocxRGBColor(0x22, 0x22, 0x22)
    normal.paragraph_format.line_spacing = 1.15

    def set_run(run, size=11, bold=False, italic=False, color=None):
        run.font.name = "Calibri"
        run.font.size = DocxPt(size)
        run.font.color.rgb = color or DocxRGBColor(0x22, 0x22, 0x22)
        run.bold = bold
        run.italic = italic
        return run

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(title_p.add_run(_sanitize_xml_text(str(title)[:200])), size=20, bold=True)

    # A real, printable worksheet has somewhere to write a name and
    # date — a small, deliberate touch that a plain AI-generated
    # question list wouldn't include on its own.
    name_p = doc.add_paragraph()
    name_p.paragraph_format.space_before = DocxPt(18)
    name_p.paragraph_format.space_after = DocxPt(18)
    set_run(name_p.add_run("Name: " + "_" * 28 + "        Date: " + "_" * 14), size=11)

    for i, q in enumerate(questions, 1):
        q_type = q.get("type") or "short_answer"
        q_text = str(q.get("question") or "").strip()
        if not q_text:
            continue

        case = str(q.get("case") or "").strip() if q.get("case_study") else ""
        if case:
            # A case study: its label, the scenario in plain text, then the question.
            lab = doc.add_paragraph()
            lab.paragraph_format.space_before = DocxPt(16)
            lab.paragraph_format.space_after = DocxPt(3)
            set_run(lab.add_run(f"{i}. Case study"), size=10.5, bold=True, color=DocxRGBColor(0x1F, 0x5F, 0x8B))
            case_p = doc.add_paragraph()
            case_p.paragraph_format.left_indent = DocxInches(0.3)
            case_p.paragraph_format.space_after = DocxPt(6)
            set_run(case_p.add_run(_sanitize_xml_text(case)), size=11)
        q_p = doc.add_paragraph()
        q_p.paragraph_format.space_before = DocxPt(4 if case else 14)
        q_p.paragraph_format.space_after = DocxPt(4)
        if case:
            q_p.paragraph_format.left_indent = DocxInches(0.3)
        else:
            set_run(q_p.add_run(f"{i}. "), bold=True)
        set_run(q_p.add_run(_sanitize_xml_text(q_text)), bold=True)

        if q_type == "multiple_choice" and isinstance(q.get("options"), list) and q["options"]:
            letters = "ABCDEFGH"
            for j, opt in enumerate(q["options"][:8]):
                opt_p = doc.add_paragraph()
                opt_p.paragraph_format.left_indent = DocxInches(0.4)
                opt_p.paragraph_format.space_after = DocxPt(2)
                set_run(opt_p.add_run(f"{letters[j]}. "), bold=False)
                set_run(opt_p.add_run(_sanitize_xml_text(str(opt))))
        else:
            # Short answer: three ruled writing lines made of literal
            # underscore text, not paragraph borders — confirmed
            # directly that consecutive paragraph borders here get
            # merged by Word/LibreOffice's rendering into a single
            # line (a real, known OOXML behavior for adjacent
            # identical borders), which silently produced only one
            # visible line no matter how the spacing was adjusted.
            # Underscores as actual run content sidestep that
            # entirely and render exactly as three separate lines.
            for _ in range(3):
                line_p = doc.add_paragraph()
                line_p.paragraph_format.left_indent = DocxInches(0.4)
                line_p.paragraph_format.space_before = DocxPt(16)
                set_run(line_p.add_run("_" * 70), color=DocxRGBColor(0x99, 0x99, 0x99))

    # Answer key on its own page — a tutor handing this to a student
    # hands over the first page only, or the whole thing to grade it
    # themselves; either way the two need to be cleanly separable.
    doc.add_page_break()
    key_title = doc.add_paragraph()
    key_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(key_title.add_run("Answer Key"), size=18, bold=True)

    for i, q in enumerate(questions, 1):
        q_text = str(q.get("question") or "").strip()
        if not q_text:
            continue
        ans = str(q.get("correct_answer") or "").strip()
        expl = str(q.get("explanation") or "").strip()

        ans_p = doc.add_paragraph()
        ans_p.paragraph_format.space_before = DocxPt(10)
        ans_p.paragraph_format.space_after = DocxPt(1)
        set_run(ans_p.add_run(f"{i}. "), bold=True)
        set_run(ans_p.add_run(_sanitize_xml_text(ans)), bold=True, color=DocxRGBColor(0x1a, 0x7a, 0x4a))

        if expl:
            expl_p = doc.add_paragraph()
            expl_p.paragraph_format.left_indent = DocxInches(0.3)
            set_run(expl_p.add_run(_sanitize_xml_text(expl)), size=10, italic=True, color=DocxRGBColor(0x66, 0x66, 0x66))

    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, filename)
    doc.save(out_path)
    return out_path


# --------------------------------------------------------- extract raw text
_PHOTO_EXTS = {"jpg", "jpeg", "png", "webp", "bmp", "tif", "tiff", "heic", "heif"}


def _prepare_photo_for_ocr(src_path, out_path):
    """Phone photos read far better after a little preparation: turned
    upright using the camera's orientation tag (a sideways photo reads as
    gibberish), greyscale, contrast stretched, and scaled down to at most
    2200px — a 12-megapixel photo takes several times longer to read and
    isn't read any better."""
    from PIL import Image, ImageOps
    with Image.open(src_path) as im:
        im = ImageOps.exif_transpose(im).convert("L")
        w, h = im.size
        scale = 2200 / max(w, h)
        if scale < 1:
            im = im.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
        im = ImageOps.autocontrast(im, cutoff=1)
        im.save(out_path, "PNG")
    # The camera's orientation tag is often missing — WhatsApp strips it
    # from forwarded photos, which is how a lot of study material travels —
    # so the text direction is also checked directly, and the page turned
    # if it's sideways or upside down (a sideways page reads as gibberish).
    # If the check can't run, the photo is simply read as it is.
    try:
        r = subprocess.run(["tesseract", out_path, "-", "--psm", "0"], capture_output=True, text=True, timeout=20)
        info = dict(l.split(":", 1) for l in r.stdout.splitlines() if ":" in l)
        turn = int(info.get("Rotate", "0").strip())
        confidence = float(info.get("Orientation confidence", "0").strip())
        if turn in (90, 180, 270) and confidence >= 1.0:
            with Image.open(out_path) as im2:
                im2.rotate(-turn, expand=True).save(out_path, "PNG")
    except Exception:
        pass
    return out_path


def _ocr_photo_text(src_path):
    work = tempfile.mkdtemp(prefix="ocr_")
    try:
        try:
            prepared = _prepare_photo_for_ocr(src_path, os.path.join(work, "page.png"))
        except Exception:
            raise ConversionError(
                "That photo couldn't be opened. JPG or PNG photos work best — on an iPhone, "
                "Settings › Camera › Formats › Most Compatible saves photos that way."
            )
        text, _headings = _ocr_page_with_headings(prepared)
        return text
    finally:
        shutil.rmtree(work, ignore_errors=True)


def _ocr_scanned_pdf_text(src_path, max_pages=10):
    """Text from a scanned PDF (pages that are pictures, with no text layer),
    by rendering each page and reading it — the first `max_pages` pages,
    to keep the wait reasonable."""
    work = tempfile.mkdtemp(prefix="ocrpdf_")
    try:
        subprocess.run(["pdftoppm", "-png", "-r", "150", "-l", str(max_pages), src_path, os.path.join(work, "p")],
                       capture_output=True, timeout=90)
        texts = []
        for name in sorted(os.listdir(work)):
            if name.endswith(".png"):
                t, _h = _ocr_page_with_headings(os.path.join(work, name))
                if t.strip():
                    texts.append(t.strip())
        return "\n\n".join(texts)
    except Exception:
        return ""
    finally:
        shutil.rmtree(work, ignore_errors=True)


def extract_text(src_path, ext):
    """Pull plain text out of an uploaded file for use in the AI tools.
    Supports the same four formats the converter already handles, plus .txt."""
    ext = ext.lower()
    if ext in _PHOTO_EXTS:
        # Photos of textbook pages, handwritten or printed notes, the
        # whiteboard, past questions — the most common "document" a
        # student has, and the AI tools used to refuse them outright.
        text = _ocr_photo_text(src_path)
        if len(text.strip()) < 3:
            raise ConversionError("No readable text was found in that photo. A sharp, well-lit photo taken straight on works best.")
        return text
    if ext == "txt":
        with open(src_path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    if ext == "docx":
        doc = _safe_load(Document, src_path)
        parts = []
        for p in doc.paragraphs:
            full_text = _full_paragraph_text(p).strip()
            if full_text:
                parts.append(full_text)
            for tb_para in _iter_textbox_paragraphs(p):
                tb_text = _full_paragraph_text(tb_para).strip()
                if tb_text:
                    parts.append(tb_text)
            for chart_title, chart_cats, chart_series, _chart_type in _iter_docx_charts(p, doc):
                parts.append(f"Chart: {chart_title}" if chart_title else "Chart data:")
                for name, values in chart_series:
                    row = [name] + list(values)
                    parts.append(" | ".join(row))
        for table in doc.tables:
            for row in table.rows:
                parts.append(" | ".join(_full_cell_text(c) for c in row.cells))
        footnotes = _get_docx_footnotes(doc)
        if footnotes:
            parts.append("Footnotes:")
            parts.extend(f"{i}. {note}" for i, note in enumerate(footnotes, 1))
        endnotes = _get_docx_endnotes(doc)
        if endnotes:
            parts.append("Endnotes:")
            parts.extend(f"{i}. {note}" for i, note in enumerate(endnotes, 1))
        comments = _get_docx_comments(doc)
        if comments:
            parts.append("Comments:")
            parts.extend(f"{author}: {comment_text}" for author, comment_text in comments)
        header_footer = _get_docx_header_footer_text(doc)
        if header_footer:
            parts.append("Header/Footer:")
            parts.extend(header_footer)
        return "\n".join(parts)
    if ext == "pdf":
        reader = _safe_load(pypdf.PdfReader, src_path)
        if reader.is_encrypted:
            raise ConversionError("This PDF is password-protected and can't be read until it's unlocked.")
        text = "\n".join((page.extract_text() or "") for page in reader.pages)
        if len(text.strip()) < 40:
            # A scan: pages that are pictures with no text layer. This used
            # to hand the AI an empty document; now the pages are read.
            text = _ocr_scanned_pdf_text(src_path) or text
        return text
    if ext == "pptx":
        prs = _safe_load(Presentation, src_path)
        parts = []
        for slide in prs.slides:
            for shape in _iter_flat_shapes(slide.shapes):
                if shape.has_text_frame and shape.text_frame.text.strip():
                    parts.append(shape.text_frame.text)
        return "\n".join(parts)
    if ext == "xlsx":
        wb = _safe_load(openpyxl.load_workbook, src_path, data_only=True)
        wb_formulas = _safe_load(openpyxl.load_workbook, src_path, data_only=False)
        parts = []
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            ws_formulas = wb_formulas[sheet]
            for row, formula_row in zip(ws.iter_rows(), ws_formulas.iter_rows()):
                cells = []
                for cell, fcell in zip(row, formula_row):
                    val = cell.value
                    if val is None and fcell.data_type == "f":
                        val = fcell.value  # uncalculated formula — show the formula itself, not blank
                    text = _format_cell_value(val, cell.number_format)
                    if text:
                        cells.append(text)
                if cells:
                    parts.append(" | ".join(cells))
            for row in ws.iter_rows():
                for cell in row:
                    if cell.comment is not None:
                        note_text = (cell.comment.text or "").strip()
                        if note_text:
                            author = (cell.comment.author or "").strip() or "Comment"
                            parts.append(f"[{cell.coordinate} comment by {author}: {note_text}]")
        return "\n".join(parts)
    raise ConversionError(f"Can't extract text from .{ext} files")


# ============================================================
# PDF EDITING ENGINE
# ============================================================
# A real, chainable set of PDF editing operations — merge, split,
# reorder, delete/extract pages, rotate, page numbers, watermark,
# password protect/remove, compress, images-to-PDF, and page
# thumbnails for a page-picker UI. Deliberately does NOT include a
# generic form-filler or true content-removing redaction: form-filling
# (per the pdf skill's own FORMS.md workflow) requires visually
# analyzing each unique form's layout — there's no deterministic,
# one-click version of that for arbitrary uploaded PDFs. Genuine
# redaction requires removing the underlying text object, not just
# drawing a box over it, which is a materially different, higher-risk
# feature than everything else here; it isn't included until it can be
# verified to actually strip the underlying text, not just visually
# cover it.

import reportlab.pdfgen.canvas as _rl_canvas
from reportlab.lib.pagesizes import letter as _RL_LETTER
from reportlab.lib.utils import ImageReader


class PdfEditError(ConversionError):
    """Raised for a PDF-edit-specific problem (an invalid page number,
    an operation applied to the wrong kind of document) — a subclass of
    ConversionError so it's caught the same way everywhere else, with a
    name that's clearer in a traceback when several edit operations are
    chained together in one call."""
    pass


def _parse_pdf_page_spec(spec, total_pages):
    """Parses a page specification into a 0-indexed list of page
    numbers, in the order given (callers that don't care about order —
    delete, most extracts — can sort the result themselves; callers
    that do — reorder, an extract meant to reshuffle pages — get
    exactly the order the caller specified, not an auto-sorted one).
    Accepts: the string "all"; a comma-separated string of 1-indexed
    page numbers and ranges ("1-3,5,7-9"); or a list/tuple of 1-indexed
    ints. Every page number is validated against total_pages — an
    out-of-range or non-numeric entry raises PdfEditError immediately
    with the offending value named, rather than silently clamping or
    skipping it, since a silently-dropped page is a worse surprise in
    an editing tool than a loud, specific error."""
    if spec == "all" or spec is None:
        return list(range(total_pages))
    if isinstance(spec, (list, tuple)):
        raw_numbers = list(spec)
    elif isinstance(spec, str):
        raw_numbers = []
        for part in spec.split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part and not part.startswith("-"):
                bounds = part.split("-")
                if len(bounds) != 2 or not all(b.strip().isdigit() for b in bounds):
                    raise PdfEditError(f"'{part}' isn't a valid page range (expected e.g. '2-5').")
                start, end = int(bounds[0]), int(bounds[1])
                if start > end:
                    raise PdfEditError(f"'{part}' is a backwards range — start page is after end page.")
                raw_numbers.extend(range(start, end + 1))
            else:
                if not part.lstrip("-").isdigit():
                    raise PdfEditError(f"'{part}' isn't a valid page number.")
                raw_numbers.append(int(part))
    else:
        raise PdfEditError(f"Unrecognized page specification: {spec!r}")

    indices = []
    for n in raw_numbers:
        if not isinstance(n, int) or n < 1 or n > total_pages:
            raise PdfEditError(
                f"Page {n} doesn't exist — this PDF has {total_pages} page"
                f"{'s' if total_pages != 1 else ''}."
            )
        indices.append(n - 1)
    return indices


def _pdf_build_from_indices(reader, indices_0based):
    """The core rebuild primitive every structural page operation
    (reorder, delete, extract, split) reduces to: build a fresh
    PdfWriter containing exactly these pages from reader, in this
    order. A fresh writer plus add_page, not an in-place reorder of an
    existing writer's own page list — confirmed directly that
    PdfWriter.pages (a _VirtualList) doesn't support slice
    reassignment, so building a new writer from the source reader is
    the reliable path, not a shortcut taken for convenience."""
    writer = pypdf.PdfWriter()
    for idx in indices_0based:
        writer.add_page(reader.pages[idx])
    return writer


def _pdf_writer_to_file(writer, out_dir, filename="edited.pdf"):
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, filename)
    with open(out_path, "wb") as f:
        writer.write(f)
    return out_path


def _check_pdf_not_encrypted(reader, action="edited"):
    """action is the full past-participle word ("compressed", "split",
    "watermarked"), not a bare verb — a naive action+"d" suffix was
    tried here originally and confirmed broken for most of the actual
    call sites ("compress"+"d" reads as "compressd", "split"+"d" as
    "splitd"), since English past participles don't form from a
    single fixed rule. Passing the whole correct word avoids needing
    one."""
    if reader.is_encrypted:
        raise PdfEditError(
            f"This PDF is password-protected and can't be {action} until it's unlocked "
            f"— use the remove-password tool first, with the correct password."
        )


def pdf_merge(pdf_paths, out_dir, filename="merged.pdf"):
    """Merges multiple PDFs, in the given order, into one document."""
    if len(pdf_paths) < 2:
        raise PdfEditError("Merging needs at least two PDF files.")
    writer = pypdf.PdfWriter()
    for path in pdf_paths:
        reader = _safe_load(pypdf.PdfReader, path)
        _check_pdf_not_encrypted(reader, "merged")
        for page in reader.pages:
            writer.add_page(page)
    return _pdf_writer_to_file(writer, out_dir, filename)


def pdf_reorder_pages(pdf_path, out_dir, order_spec, filename="reordered.pdf"):
    """Reorders every page in the document according to order_spec — a
    page spec (see _parse_pdf_page_spec) that must reference every
    existing page exactly once. Rejected (rather than silently
    dropping or duplicating pages) if it doesn't — a partial reorder
    spec is far more likely to be a mistake (a page the user forgot to
    include) than an intentional page deletion, and delete_pages is
    the explicit, named tool for that."""
    reader = _safe_load(pypdf.PdfReader, pdf_path)
    _check_pdf_not_encrypted(reader, "reordered")
    total = len(reader.pages)
    indices = _parse_pdf_page_spec(order_spec, total)
    if sorted(indices) != list(range(total)):
        missing = sorted(set(range(total)) - set(indices))
        dup_check = len(indices) != len(set(indices))
        if dup_check:
            raise PdfEditError("The new page order lists the same page more than once.")
        raise PdfEditError(
            f"The new page order is missing page(s) "
            f"{', '.join(str(i + 1) for i in missing)} — every page needs a new position, "
            f"not just some of them. Use delete_pages first if you actually want to drop pages."
        )
    writer = _pdf_build_from_indices(reader, indices)
    return _pdf_writer_to_file(writer, out_dir, filename)


def pdf_delete_pages(pdf_path, out_dir, pages_spec, filename="deleted.pdf"):
    """Removes the specified pages, keeping the rest in their original
    order."""
    reader = _safe_load(pypdf.PdfReader, pdf_path)
    _check_pdf_not_encrypted(reader, "edited")
    total = len(reader.pages)
    to_remove = set(_parse_pdf_page_spec(pages_spec, total))
    if len(to_remove) >= total:
        raise PdfEditError("That would delete every page — a PDF can't be left with zero pages.")
    keep = [i for i in range(total) if i not in to_remove]
    writer = _pdf_build_from_indices(reader, keep)
    return _pdf_writer_to_file(writer, out_dir, filename)


def pdf_extract_pages(pdf_path, out_dir, pages_spec, filename="extracted.pdf"):
    """Keeps only the specified pages, in the order given — e.g.
    extracting "5,2" produces a 2-page document with the original page
    5 first, then the original page 2, which is deliberate: this is
    also how a user reorders while dropping pages in one step, rather
    than needing reorder and delete as two separate operations."""
    reader = _safe_load(pypdf.PdfReader, pdf_path)
    _check_pdf_not_encrypted(reader, "edited")
    total = len(reader.pages)
    indices = _parse_pdf_page_spec(pages_spec, total)
    if not indices:
        raise PdfEditError("No pages were specified to extract.")
    writer = _pdf_build_from_indices(reader, indices)
    return _pdf_writer_to_file(writer, out_dir, filename)


def pdf_split(pdf_path, out_dir, ranges=None):
    """Splits a PDF into multiple files. With ranges=None, splits into
    one file per page. With ranges given (a list of page specs, e.g.
    ["1-3", "4-6", "7"]), produces one file per range instead. Returns
    a list of output paths, in order."""
    reader = _safe_load(pypdf.PdfReader, pdf_path)
    _check_pdf_not_encrypted(reader, "split")
    total = len(reader.pages)
    os.makedirs(out_dir, exist_ok=True)
    out_paths = []
    if ranges is None:
        pad = len(str(total))
        for i in range(total):
            writer = _pdf_build_from_indices(reader, [i])
            out_paths.append(_pdf_writer_to_file(writer, out_dir, f"page_{str(i + 1).zfill(pad)}.pdf"))
        return out_paths
    if len(ranges) < 2:
        raise PdfEditError("Splitting by ranges needs at least two ranges — a single range is just an extract.")
    pad = len(str(len(ranges)))
    for i, r in enumerate(ranges, 1):
        indices = _parse_pdf_page_spec(r, total)
        if not indices:
            raise PdfEditError(f"Range '{r}' didn't resolve to any pages.")
        writer = _pdf_build_from_indices(reader, indices)
        out_paths.append(_pdf_writer_to_file(writer, out_dir, f"part_{str(i).zfill(pad)}.pdf"))
    return out_paths


def pdf_rotate_pages(pdf_path, out_dir, pages_spec, degrees, filename="rotated.pdf"):
    """Rotates the specified pages (pages_spec, or "all") clockwise by
    degrees, which must be a multiple of 90 — pypdf's own rotate()
    requires this, and a non-multiple would silently produce a subtly
    skewed page rather than a clean quarter/half/three-quarter turn."""
    if degrees % 90 != 0:
        raise PdfEditError("Rotation must be a multiple of 90 degrees.")
    reader = _safe_load(pypdf.PdfReader, pdf_path)
    _check_pdf_not_encrypted(reader, "rotated")
    total = len(reader.pages)
    targets = set(_parse_pdf_page_spec(pages_spec, total))
    writer = pypdf.PdfWriter()
    for i in range(total):
        page = reader.pages[i]
        added = writer.add_page(page)
        if i in targets:
            added.rotate(degrees % 360)
    return _pdf_writer_to_file(writer, out_dir, filename)


_PDF_NUMBER_POSITIONS = {
    "bottom-center": lambda w, h: (w / 2, 0.4 * 72, "center"),
    "bottom-right": lambda w, h: (w - 0.6 * 72, 0.4 * 72, "right"),
    "bottom-left": lambda w, h: (0.6 * 72, 0.4 * 72, "left"),
    "top-center": lambda w, h: (w / 2, h - 0.6 * 72, "center"),
    "top-right": lambda w, h: (w - 0.6 * 72, h - 0.6 * 72, "right"),
    "top-left": lambda w, h: (0.6 * 72, h - 0.6 * 72, "left"),
}


def pdf_add_page_numbers(pdf_path, out_dir, position="bottom-center", start_at=1,
                          fmt="{n}", pages_spec="all", filename="numbered.pdf"):
    """Overlays a page number on each targeted page. fmt is a template
    with {n} (this page's number, starting from start_at) and {total}
    (the count of numbered pages) — e.g. "Page {n} of {total}". Built
    per-page rather than with one fixed overlay, since pages in a real
    PDF can have different sizes (a merged document combining a
    portrait worksheet with a landscape scan, for instance), and a
    single fixed-size overlay would misplace the number on any page
    that doesn't match it."""
    if position not in _PDF_NUMBER_POSITIONS:
        raise PdfEditError(f"Unknown position '{position}' — choose one of {', '.join(_PDF_NUMBER_POSITIONS)}.")
    reader = _safe_load(pypdf.PdfReader, pdf_path)
    _check_pdf_not_encrypted(reader, "numbered")
    total_pages = len(reader.pages)
    targets = set(_parse_pdf_page_spec(pages_spec, total_pages))
    total_numbered = len(targets)
    writer = pypdf.PdfWriter()
    numbered_so_far = 0
    for i in range(total_pages):
        page = reader.pages[i]
        if i in targets:
            numbered_so_far += 1
            n = start_at + numbered_so_far - 1
            label = fmt.format(n=n, total=total_numbered)
            w, h = float(page.mediabox.width), float(page.mediabox.height)
            buf = io.BytesIO()
            c = _rl_canvas.Canvas(buf, pagesize=(w, h))
            c.setFont("Helvetica", 10)
            c.setFillColorRGB(0.3, 0.3, 0.3)
            x, y, align = _PDF_NUMBER_POSITIONS[position](w, h)
            if align == "center":
                c.drawCentredString(x, y, label)
            elif align == "right":
                c.drawRightString(x, y, label)
            else:
                c.drawString(x, y, label)
            c.save()
            buf.seek(0)
            overlay_page = pypdf.PdfReader(buf).pages[0]
            page.merge_page(overlay_page)
        writer.add_page(page)
    return _pdf_writer_to_file(writer, out_dir, filename)


def pdf_add_watermark(pdf_path, out_dir, text, pages_spec="all", opacity=0.3,
                       angle=45, font_size=60, color="808080", filename="watermarked.pdf"):
    """Overlays diagonal (or any angle) semi-transparent watermark text
    across the targeted pages, sized to each page individually for the
    same reason page numbering is — a mixed-size document shouldn't get
    a watermark positioned for the wrong page size."""
    if not (0 < opacity <= 1):
        raise PdfEditError("Watermark opacity must be greater than 0 and at most 1.")
    try:
        rgb = tuple(int(color[i:i + 2], 16) / 255 for i in (0, 2, 4))
    except (ValueError, IndexError):
        raise PdfEditError(f"'{color}' isn't a valid hex color (expected e.g. '808080').")
    reader = _safe_load(pypdf.PdfReader, pdf_path)
    _check_pdf_not_encrypted(reader, "watermarked")
    total_pages = len(reader.pages)
    targets = set(_parse_pdf_page_spec(pages_spec, total_pages))
    writer = pypdf.PdfWriter()
    for i in range(total_pages):
        page = reader.pages[i]
        if i in targets:
            w, h = float(page.mediabox.width), float(page.mediabox.height)
            buf = io.BytesIO()
            c = _rl_canvas.Canvas(buf, pagesize=(w, h))
            c.saveState()
            c.setFillAlpha(opacity)
            c.setFillColorRGB(*rgb)
            c.setFont("Helvetica-Bold", font_size)
            c.translate(w / 2, h / 2)
            c.rotate(angle)
            c.drawCentredString(0, 0, text)
            c.restoreState()
            c.save()
            buf.seek(0)
            overlay_page = pypdf.PdfReader(buf).pages[0]
            page.merge_page(overlay_page)
        writer.add_page(page)
    return _pdf_writer_to_file(writer, out_dir, filename)


def pdf_set_password(pdf_path, out_dir, user_password, owner_password=None, filename="protected.pdf"):
    """Encrypts the PDF — user_password is required to open it at all;
    owner_password (defaults to the same as user_password if not given)
    governs permissions like printing/editing in readers that enforce
    that distinction."""
    if not user_password:
        raise PdfEditError("A password is required to protect a PDF.")
    reader = _safe_load(pypdf.PdfReader, pdf_path)
    _check_pdf_not_encrypted(reader, "password-protected")
    writer = pypdf.PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.encrypt(user_password, owner_password or user_password)
    return _pdf_writer_to_file(writer, out_dir, filename)


def pdf_remove_password(pdf_path, out_dir, password, filename="unlocked.pdf"):
    """Decrypts a password-protected PDF given the correct password."""
    reader = _safe_load(pypdf.PdfReader, pdf_path)
    if not reader.is_encrypted:
        raise PdfEditError("This PDF isn't password-protected — there's nothing to remove.")
    try:
        result = reader.decrypt(password)
    except Exception:
        result = 0
    if not result:
        raise PdfEditError("That password didn't unlock this PDF — check it and try again.")
    writer = pypdf.PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    return _pdf_writer_to_file(writer, out_dir, filename)


def pdf_compress(pdf_path, out_dir, image_quality=60, filename="compressed.pdf"):
    """Shrinks file size primarily by recompressing embedded images at
    a lower JPEG quality (most of a scanned-worksheet PDF's size is
    almost always its images, not its text/vector content) — pypdf's
    own compress_content_streams also runs, which helps vector-heavy
    pages but does little for an image-heavy scan on its own."""
    if not (1 <= image_quality <= 100):
        raise PdfEditError("Image quality must be between 1 and 100.")
    reader = _safe_load(pypdf.PdfReader, pdf_path)
    _check_pdf_not_encrypted(reader, "compressed")
    writer = pypdf.PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    for page in writer.pages:
        for img in page.images:
            try:
                img.replace(img.image, quality=image_quality)
            except Exception:
                continue  # an image pypdf can't safely re-encode is left as-is, not dropped
    try:
        writer.compress_identical_objects()
    except Exception:
        pass
    for page in writer.pages:
        try:
            page.compress_content_streams()
        except Exception:
            pass
    out_path = _pdf_writer_to_file(writer, out_dir, filename)
    # A PDF whose images were already simple/well-compressed (a
    # screenshot, a flat-color diagram) can come out of a JPEG
    # re-encode *larger* than the source, confirmed directly — a
    # "compress" feature that sometimes makes the file bigger is a
    # real defect, not a rare edge case worth ignoring, so fall back to
    # the original file whenever the result doesn't actually shrink it.
    if os.path.getsize(out_path) >= os.path.getsize(pdf_path):
        os.remove(out_path)
        fallback_path = os.path.join(out_dir, filename)
        with open(pdf_path, "rb") as src, open(fallback_path, "wb") as dst:
            dst.write(src.read())
        return fallback_path
    return out_path
    return _pdf_writer_to_file(writer, out_dir, filename)


# 3.8: Make a PDF smaller — step by step until it fits (see server_shrink38.py)
SHRINK_STEPS = [(2000, 75), (1700, 66), (1450, 58), (1250, 50), (1050, 43), (900, 36), (760, 30)]


def _pdf_shrink_once(pdf_path, out_path, max_px, quality):
    """One try: every big picture at most max_px on its long side, saved as JPEG at `quality`. Returns how many
    pictures were redone."""
    from PIL import Image as _PILImage
    reader = _safe_load(pypdf.PdfReader, pdf_path)
    writer = pypdf.PdfWriter(clone_from=reader)
    done, redone = set(), 0
    for page in writer.pages:
        try:
            images = list(page.images)
        except Exception:
            continue
        for img in images:
            ref = getattr(img, "indirect_reference", None)
            key = getattr(ref, "idnum", None)
            if key is None or key in done:
                continue
            done.add(key)
            try:
                xobj = ref.get_object()
                if "/SMask" in xobj or "/Mask" in xobj or xobj.get("/ImageMask"):
                    continue                     # see-through parts would be lost
                if len(img.data or b"") < 24 * 1024:
                    continue                     # already small: not worth touching
                pic = img.image
                if pic is None or pic.mode in ("1",):
                    continue                     # fax-style black-and-white scans are already tiny
                if pic.mode not in ("RGB", "L"):
                    pic = pic.convert("RGB")
                w, h = pic.size
                k = min(1.0, float(max_px) / max(w, h))
                if k < 1.0:
                    pic = pic.resize((max(1, int(w * k)), max(1, int(h * k))), _PILImage.LANCZOS)
                img.replace(pic, quality=quality)
                redone += 1
            except Exception:
                continue                         # a picture that can't be redone safely stays as it was
    try:
        writer.compress_identical_objects()
    except Exception:
        pass
    for page in writer.pages:
        try:
            page.compress_content_streams()
        except Exception:
            pass
    with open(out_path, "wb") as f:
        writer.write(f)
    return redone


def pdf_shrink(pdf_path, out_dir, target_bytes=None, filename="smaller.pdf", time_budget=55.0):
    """Makes a PDF smaller, aiming at no more than target_bytes (None: as small as the steps go). Returns
    (path, info) with info = {"before", "after", "fits", "pictures"}; the original comes back unchanged when nothing
    makes it smaller (a PDF of text and drawings)."""
    import time as _time
    reader = _safe_load(pypdf.PdfReader, pdf_path)
    _check_pdf_not_encrypted(reader, "made smaller")
    before = os.path.getsize(pdf_path)
    os.makedirs(out_dir, exist_ok=True)
    start = _time.monotonic()
    best_path, best_size, pictures = None, before, 0
    steps = SHRINK_STEPS
    i = 0
    while i < len(steps):
        max_px, q = steps[i]
        trial = os.path.join(out_dir, "try-%d.pdf" % i)
        try:
            n = _pdf_shrink_once(pdf_path, trial, max_px, q)
        except Exception:
            break
        pictures = max(pictures, n)
        size = os.path.getsize(trial)
        if size < best_size:
            if best_path and os.path.exists(best_path):
                os.remove(best_path)
            best_path, best_size = trial, size
        elif os.path.exists(trial):
            os.remove(trial)
        if n == 0:
            break                                # no pictures to work on: more steps change nothing
        if target_bytes and best_size <= target_bytes:
            break
        if _time.monotonic() - start > time_budget:
            break
        # far from the target: skip a step (each step is about a quarter smaller)
        i += 2 if (target_bytes and best_size > target_bytes * 2.2 and i + 2 < len(steps)) else 1
    out_path = os.path.join(out_dir, filename)
    if best_path:
        os.replace(best_path, out_path)
    else:
        shutil.copyfile(pdf_path, out_path)
    after = os.path.getsize(out_path)
    return out_path, {"before": before, "after": after, "fits": bool(target_bytes is None or after <= target_bytes),
                      "pictures": pictures}


def images_to_pdf(image_paths, out_dir, filename="images.pdf", display_names=None):
    """Combines one or more images into a PDF, one image per page,
    each page sized to that image's own aspect ratio (letter-width,
    scaled) rather than forcing every image into one fixed page size
    and distorting or letterboxing it. display_names, if given, is a
    list parallel to image_paths holding the user's original filenames
    — used only for the error message below, since image_paths itself
    is a list of randomized server-side temp paths (see
    safe_upload_filename in app.py) that mean nothing to the user and
    shouldn't appear in anything shown to them. The underlying PIL
    exception is deliberately not included in that message either — it
    embeds the same server-side path in its own text."""
    from PIL import Image
    if not image_paths:
        raise PdfEditError("No images were given to combine into a PDF.")
    display_names = display_names or [os.path.basename(p) for p in image_paths]
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, filename)
    page_w = _RL_LETTER[0]
    c = _rl_canvas.Canvas(out_path)
    for i, img_path in enumerate(image_paths):
        display_name = display_names[i] if i < len(display_names) else os.path.basename(img_path)
        try:
            with Image.open(img_path) as im:
                im = im.convert("RGB")
                iw, ih = im.size
                page_h = page_w * ih / iw
                c.setPageSize((page_w, page_h))
                buf = io.BytesIO()
                im.save(buf, format="JPEG", quality=90)
                buf.seek(0)
                c.drawImage(ImageReader(buf), 0, 0, width=page_w, height=page_h)
                c.showPage()
        except Exception:
            raise PdfEditError(f"'{display_name}' doesn't look like a valid image file — couldn't add it to the PDF.")
    c.save()
    return out_path


def pdf_get_page_thumbnails(pdf_path, out_dir, max_pages=200, dpi=40):
    """Rasterizes each page to a small JPEG thumbnail, for a page-picker
    UI where a user selects/reorders pages visually rather than by
    typing page numbers. Capped at max_pages for a very long document,
    since rendering hundreds of thumbnails at once isn't a reasonable
    single request regardless of how fast any one page is.

    dpi=40, not the 80 this used originally: measured directly against
    the actual on-screen thumbnail size (a 3-column grid on a typical
    phone puts each one around 110 CSS px, ~220px accounting for a 2x
    display) — at that real display size, a side-by-side render of 80
    vs. 40 DPI was visually indistinguishable, while 40 cut the
    rasterization time by roughly a quarter and, more importantly, the
    JPEG payload size by about 70% (7.0MB -> 2.0MB across 50 pages in
    testing) — the more significant win on a mobile connection, where
    transferring that payload back to the client usually costs more
    time than generating it server-side ever did.
    """
    reader = _safe_load(pypdf.PdfReader, pdf_path)
    _check_pdf_not_encrypted(reader, "previewed")
    total = len(reader.pages)
    if total > max_pages:
        raise PdfEditError(
            f"This PDF has {total} pages — thumbnail preview is limited to {max_pages} "
            f"pages at a time. Split it first if you need to work with all of it."
        )
    os.makedirs(out_dir, exist_ok=True)
    prefix = os.path.join(out_dir, "thumb")
    subprocess.run(
        ["pdftoppm", "-jpeg", "-r", str(dpi), pdf_path, prefix],
        capture_output=True, text=True, timeout=60, check=True
    )
    # pdftoppm zero-pads its numeric suffix to the width of the total
    # page count (page 1 of a 12-page PDF becomes "thumb-01.jpg", but
    # "thumb-1.jpg" for a 5-page one) — confirmed directly against both
    # a single-digit and a double-digit page count, rather than assumed.
    pad = len(str(total))
    thumbs = [f"{prefix}-{str(i).zfill(pad)}.jpg" for i in range(1, total + 1)]
    missing = [p for p in thumbs if not os.path.exists(p)]
    if missing:
        raise PdfEditError(f"Thumbnail generation produced {total - len(missing)} of {total} expected pages.")
    return thumbs


def apply_pdf_operations(pdf_path, operations, out_dir, extra_files=None):
    """Applies a list of edit operations in sequence to build one final
    PDF — the real editing-pipeline experience (merge, then number the
    pages, then watermark, all exported together in one pass) rather
    than requiring a separate round-trip download/re-upload between
    each step. Each operation is a dict with a "type" key and whatever
    parameters that type needs:

      {"type": "merge_append", "files": ["key1", "key2"]}
        appends the named files (looked up in extra_files) to the end
        of the current document, in the order given
      {"type": "reorder_pages", "order": "3,1,2,4"}
      {"type": "delete_pages", "pages": "2,4"}
      {"type": "extract_pages", "pages": "1-3,5"}
      {"type": "rotate", "pages": "all", "degrees": 90}
      {"type": "add_page_numbers", "position": "bottom-center",
       "start_at": 1, "fmt": "Page {n} of {total}", "pages": "all"}
      {"type": "add_watermark", "text": "DRAFT", "pages": "all",
       "opacity": 0.3, "angle": 45, "font_size": 60, "color": "808080"}
      {"type": "set_password", "user_password": "...", "owner_password": "..."}
      {"type": "remove_password", "password": "..."}
      {"type": "compress", "image_quality": 60}

    extra_files maps a name referenced by a merge_append operation to
    an actual file path — the additional PDFs being merged in, kept
    separate from pdf_path so the caller doesn't need to pre-merge
    anything before building the operation list. Raises PdfEditError
    (naming which operation, by position, failed) on an unknown
    operation type or invalid parameters — a chain of edits failing
    silently partway through and returning a half-edited file would be
    far worse than stopping with a clear error."""
    if not operations:
        raise PdfEditError("No edit operations were specified.")
    extra_files = extra_files or {}

    # set_password is deferred to the end of the chain regardless of
    # where the caller placed it — confirmed directly that queuing it
    # before any other operation (compress, watermark, anything) fails
    # with an encrypted-file error, since every other operation needs
    # to read the PDF's actual content. The frontend's tool grid lets
    # someone tap "Add password" and "Compress" in either order, and
    # there's exactly one order that ever makes sense regardless of
    # which was tapped first — so the system enforces it rather than
    # relying on the user to have gotten the sequence right themselves.
    # remove_password is deliberately untouched: it has to run first,
    # not last, to unlock the file for whatever follows it.
    password_ops = [op for op in operations if op.get("type") == "set_password"]
    other_ops = [op for op in operations if op.get("type") != "set_password"]
    operations = other_ops + password_ops

    current_path = pdf_path
    os.makedirs(out_dir, exist_ok=True)
    for step_num, op in enumerate(operations, 1):
        op_type = op.get("type")
        try:
            if op_type == "merge_append":
                file_keys = op.get("files") or []
                if not file_keys:
                    raise PdfEditError("merge_append needs at least one file to append.")
                paths = [current_path]
                for key in file_keys:
                    if key not in extra_files:
                        raise PdfEditError(f"merge_append referenced '{key}', which wasn't provided.")
                    paths.append(extra_files[key])
                current_path = pdf_merge(paths, out_dir, filename=f"step{step_num}.pdf")
            elif op_type == "reorder_pages":
                current_path = pdf_reorder_pages(current_path, out_dir, op["order"], filename=f"step{step_num}.pdf")
            elif op_type == "delete_pages":
                current_path = pdf_delete_pages(current_path, out_dir, op["pages"], filename=f"step{step_num}.pdf")
            elif op_type == "extract_pages":
                current_path = pdf_extract_pages(current_path, out_dir, op["pages"], filename=f"step{step_num}.pdf")
            elif op_type == "rotate":
                current_path = pdf_rotate_pages(
                    current_path, out_dir, op.get("pages", "all"), op["degrees"], filename=f"step{step_num}.pdf"
                )
            elif op_type == "add_page_numbers":
                current_path = pdf_add_page_numbers(
                    current_path, out_dir,
                    position=op.get("position", "bottom-center"),
                    start_at=op.get("start_at", 1),
                    fmt=op.get("fmt", "{n}"),
                    pages_spec=op.get("pages", "all"),
                    filename=f"step{step_num}.pdf",
                )
            elif op_type == "add_watermark":
                if "text" not in op:
                    raise PdfEditError("add_watermark needs 'text'.")
                current_path = pdf_add_watermark(
                    current_path, out_dir, text=op["text"],
                    pages_spec=op.get("pages", "all"),
                    opacity=op.get("opacity", 0.3),
                    angle=op.get("angle", 45),
                    font_size=op.get("font_size", 60),
                    color=op.get("color", "808080"),
                    filename=f"step{step_num}.pdf",
                )
            elif op_type == "set_password":
                if "user_password" not in op:
                    raise PdfEditError("set_password needs 'user_password'.")
                current_path = pdf_set_password(
                    current_path, out_dir, op["user_password"], op.get("owner_password"),
                    filename=f"step{step_num}.pdf",
                )
            elif op_type == "remove_password":
                if "password" not in op:
                    raise PdfEditError("remove_password needs 'password'.")
                current_path = pdf_remove_password(current_path, out_dir, op["password"], filename=f"step{step_num}.pdf")
            elif op_type == "compress":
                current_path = pdf_compress(
                    current_path, out_dir, op.get("image_quality", 60), filename=f"step{step_num}.pdf"
                )
            else:
                raise PdfEditError(f"Unknown edit operation '{op_type}'.")
        except KeyError as e:
            raise PdfEditError(f"Step {step_num} ('{op_type}') is missing required field {e}.")
        except PdfEditError as e:
            raise PdfEditError(f"Step {step_num} ('{op_type}'): {e}")
    final_path = os.path.join(out_dir, "edited_final.pdf")
    if current_path != final_path:
        with open(current_path, "rb") as src, open(final_path, "wb") as dst:
            dst.write(src.read())
    return final_path




