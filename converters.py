"""
MasterConvert conversion engine.
Real conversions using LibreOffice (rendering-accurate formats) and
structural rebuilding (python-docx/pptx/openpyxl) for formats that have
no direct renderer path between them.
"""
import datetime
import io
import os
import re
import subprocess
import tempfile
import shutil
from collections import Counter
from docx import Document
from docx.shared import Inches as DocxInches, Pt as DocxPt, RGBColor as DocxRGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt
from pptx.dml.color import RGBColor as PptxRGBColor
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font as XlsxFont
import pypdf


class ConversionError(Exception):
    pass


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
    r"^([\u2022\u25CF\u25AA\u25E6\uf0b7\uf0a7\uf076\uf0d8\-\*]|\d{1,3}[\.\)])\s+"
)
_SENTENCE_END_RE = re.compile(r"[.!?][\'\")\]]*$")
_SHORT_LINE_LEN = 70  # a full wrapped line in a standard document usually runs
                      # much longer than this; a short line ending a sentence
                      # is the most reliable signal pypdf gives us that a
                      # paragraph has actually finished, since it never
                      # exposes real layout/whitespace metadata to check instead.


def _reconstruct_paragraphs(raw_text):
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
    with what follows. Returns a list of (text, kind) tuples, kind in
    {"para", "bullet", "heading"}."""
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
        if _BULLET_RE.match(line):
            flush()
            items.append((_BULLET_RE.sub("", line).strip(), "bullet"))
            continue
        ends_sentence = bool(_SENTENCE_END_RE.search(line))
        if not current and not ends_sentence and len(line) < 80 and line[:1].isupper():
            # A short line, sitting on its own, that doesn't end mid-sentence
            # and starts capitalized reads as a heading far more often than
            # not — e.g. "Section One", "PDF Test Document" — this no longer
            # relies on line.istitle()-style checks, which break on any line
            # containing an all-caps acronym.
            items.append((line, "heading"))
            continue
        current.append(line)
        if ends_sentence and len(line) < _SHORT_LINE_LEN:
            flush()
    flush()
    return items



# ---------------------------------------------------------- PPTX templates
PPTX_TEMPLATES = {
    "minimal": {
        "bg": None, "title_fill": None,
        "title_font": "Calibri Light", "title_size": Pt(36), "title_bold": False,
        "title_color": PptxRGBColor(0x33, 0x33, 0x33),
        "body_font": "Calibri Light", "body_size": Pt(20),
        "body_color": PptxRGBColor(0x55, 0x55, 0x55),
    },
    "academic": {
        "bg": None, "title_fill": None,
        "title_font": "Georgia", "title_size": Pt(32), "title_bold": True,
        "title_color": PptxRGBColor(0x1F, 0x3A, 0x5F),
        "body_font": "Georgia", "body_size": Pt(18),
        "body_color": PptxRGBColor(0x22, 0x22, 0x22),
    },
    "bold": {
        "bg": PptxRGBColor(0x0A, 0x0A, 0x0A), "title_fill": None,
        "title_font": "Arial", "title_size": Pt(44), "title_bold": True,
        "title_color": PptxRGBColor(0xFF, 0xFF, 0xFF),
        "body_font": "Arial", "body_size": Pt(20),
        "body_color": PptxRGBColor(0xE0, 0xE0, 0xE0),
    },
    "classic": {
        "bg": None, "title_fill": PptxRGBColor(0x1F, 0x38, 0x64),
        "title_font": "Calibri", "title_size": Pt(30), "title_bold": True,
        "title_color": PptxRGBColor(0xFF, 0xFF, 0xFF),
        "body_font": "Calibri", "body_size": Pt(18),
        "body_color": PptxRGBColor(0x22, 0x22, 0x22),
    },
}


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
            run.font.color.rgb = style["title_color"]


def _style_pptx_body_paragraph(p, style):
    p.font.name = style["body_font"]
    p.font.size = style["body_size"]
    p.font.color.rgb = style["body_color"]


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


def _style_docx_headings(doc, style):
    """Applies to every heading paragraph already in the document (headings
    are added via doc.add_heading before this runs)."""
    for para in doc.paragraphs:
        if not para.style.name.lower().startswith("heading") and para.style.name != "Title":
            continue
        if style["uppercase_headings"]:
            for run in para.runs:
                run.text = run.text.upper()
        for run in para.runs:
            run.font.name = style["heading_font"]
            run.font.color.rgb = style["heading_color"]


def _style_docx_body(doc, style):
    for para in doc.paragraphs:
        is_heading = para.style.name.lower().startswith("heading") or para.style.name == "Title"
        if is_heading:
            continue
        for run in para.runs:
            run.font.name = style["body_font"]
            run.font.size = style["body_size"]
        if style.get("tight_spacing"):
            para.paragraph_format.space_before = DocxPt(0)
            para.paragraph_format.space_after = DocxPt(2)


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


def apply_docx_style(doc, style_name):
    style = DOCX_STYLES.get(style_name, DOCX_STYLES["clean"])
    _style_docx_headings(doc, style)
    _style_docx_body(doc, style)
    if style["page_numbers"]:
        _add_docx_page_number_footer(doc)


def _soffice_convert(src_path, target_format, out_dir):
    """Use headless LibreOffice for true rendering-based conversions
    (currently: docx->pdf). Raises ConversionError on failure."""
    result = subprocess.run(
        ["soffice", "--headless", "--nologo", "--nofirststartwizard",
         "--convert-to", target_format, "--outdir", out_dir, src_path],
        capture_output=True, text=True, timeout=60
    )
    base = os.path.splitext(os.path.basename(src_path))[0]
    out_path = os.path.join(out_dir, f"{base}.{target_format}")
    if not os.path.exists(out_path):
        raise ConversionError(f"LibreOffice conversion failed: {result.stderr or result.stdout}")
    return out_path


# ---------------------------------------------------------------- DOCX -> PDF
def docx_to_pdf(src_path, out_dir, style=None):
    return _soffice_convert(src_path, "pdf", out_dir)


# --------------------------------------------------------------- DOCX -> PPTX
def _iter_inline_images(paragraph, doc):
    """Yields raw image bytes for each inline picture embedded in this
    paragraph's runs, in order. python-docx's own API doesn't expose
    embedded pictures directly off a paragraph — this reaches into the
    run's XML for the relationship id of each <a:blip> drawing and resolves
    it against the document's image parts."""
    for run in paragraph.runs:
        for blip in run._element.findall(".//" + qn("a:blip")):
            r_id = blip.get(qn("r:embed"))
            if not r_id:
                continue
            try:
                yield doc.part.related_parts[r_id].blob
            except KeyError:
                continue


def docx_to_pptx(src_path, out_dir, style="minimal"):
    template_style = PPTX_TEMPLATES.get(style, PPTX_TEMPLATES["minimal"])
    doc = Document(src_path)
    prs = Presentation()
    prs.slide_width = PptxInches(13.333)
    prs.slide_height = PptxInches(7.5)
    title_layout = prs.slide_layouts[1]  # title + content
    blank_layout = prs.slide_layouts[6]

    MAX_BULLETS_PER_SLIDE = 8  # keep slides readable; overflow spills onto a "(cont.)" slide
    state = {"slide": None, "body_tf": None, "bullet_count": 0}

    def new_slide(title_text):
        s = prs.slides.add_slide(title_layout)
        s.shapes.title.text = title_text or "Untitled"
        _style_pptx_slide(s, template_style)
        tf = s.placeholders[1].text_frame
        tf.clear()
        state["slide"], state["body_tf"], state["bullet_count"] = s, tf, 0
        return s, tf

    def add_bullet(para, level=0):
        if state["slide"] is None:
            new_slide("Overview")
        if state["bullet_count"] >= MAX_BULLETS_PER_SLIDE:
            current_title = state["slide"].shapes.title.text or "Overview"
            new_slide(current_title + " (cont.)")
        body_tf = state["body_tf"]
        if body_tf.paragraphs[0].text == "" and len(body_tf.paragraphs) == 1 and state["bullet_count"] == 0:
            p = body_tf.paragraphs[0]
        else:
            p = body_tf.add_paragraph()
        p.level = min(level, 4)
        runs = [r for r in para.runs if r.text]
        if not runs:
            p.text = para.text.strip()
        for run in runs:
            r = p.add_run()
            r.text = run.text
            r.font.bold = bool(run.bold)
            r.font.italic = bool(run.italic)
            r.font.underline = bool(run.underline)
        _style_pptx_body_paragraph(p, template_style)
        # Re-apply bold/italic after the shared style pass, since that pass
        # sets font attributes on every run and would otherwise stomp the
        # per-run emphasis we just copied over.
        for r, src in zip(p.runs, runs):
            r.font.bold = bool(src.bold)
            r.font.italic = bool(src.italic)
            r.font.underline = bool(src.underline)
        state["bullet_count"] += 1

    def add_table_slide(table):
        rows = [[cell.text.strip() for cell in row.cells] for row in table.rows]
        rows = [r for r in rows if any(r)]
        if not rows:
            return
        n_rows, n_cols = len(rows), max(len(r) for r in rows)
        s = prs.slides.add_slide(blank_layout)
        left, top = PptxInches(0.6), PptxInches(0.6)
        width, height = prs.slide_width - PptxInches(1.2), prs.slide_height - PptxInches(1.2)
        gtable = s.shapes.add_table(n_rows, n_cols, left, top, width, height).table
        for r_idx, row in enumerate(rows):
            for c_idx in range(n_cols):
                cell = gtable.cell(r_idx, c_idx)
                cell.text = row[c_idx] if c_idx < len(row) else ""
                if r_idx == 0:
                    for p in cell.text_frame.paragraphs:
                        for run in p.runs:
                            run.font.bold = True
        # A loose paragraph appearing right after a table should land on a
        # fresh slide rather than silently reusing the table slide.
        state["slide"], state["body_tf"], state["bullet_count"] = None, None, 0

    def add_image_slide(image_bytes):
        s = prs.slides.add_slide(blank_layout)
        try:
            pic = s.shapes.add_picture(io.BytesIO(image_bytes), 0, 0)
        except Exception:
            return  # a malformed/unsupported embedded image shouldn't sink the whole conversion
        scale = min(prs.slide_width / pic.width, prs.slide_height / pic.height, 1) if pic.width and pic.height else 1
        pic.width, pic.height = int(pic.width * scale), int(pic.height * scale)
        pic.left = int((prs.slide_width - pic.width) / 2)
        pic.top = int((prs.slide_height - pic.height) / 2)
        # A loose paragraph appearing right after an image should land on a
        # fresh slide rather than silently reusing the image slide.
        state["slide"], state["body_tf"], state["bullet_count"] = None, None, 0

    MAX_IMAGES = 30
    image_count = 0

    for block in _iter_block_items(doc):
        if isinstance(block, DocxTable):
            add_table_slide(block)
            continue
        para = block
        if image_count < MAX_IMAGES:
            for img_bytes in _iter_inline_images(para, doc):
                if image_count >= MAX_IMAGES:
                    break
                add_image_slide(img_bytes)
                image_count += 1
        text = para.text.strip()
        if not text:
            continue
        para_style_name = (para.style.name or "").lower()
        if "heading" in para_style_name or para_style_name == "title":
            new_slide(text)
        else:
            level = 0
            if "list" in para_style_name:
                m = re.search(r"(\d+)", para_style_name)
                if m:
                    level = max(0, min(int(m.group(1)) - 1, 4))
            add_bullet(para, level=level)

    if state["slide"] is None:
        new_slide("Untitled Document")

    out_path = os.path.join(out_dir, "converted.pptx")
    prs.save(out_path)
    return out_path


# --------------------------------------------------------------- PPTX -> DOCX
def pptx_to_docx(src_path, out_dir, style="clean"):
    prs = Presentation(src_path)
    doc = Document()
    doc.add_heading("Slide Handout", 0)
    BULLET_STYLES = ["List Bullet", "List Bullet 2", "List Bullet 3"]

    for i, slide in enumerate(prs.slides, 1):
        title = None
        text_shapes = []
        table_shapes = []
        for shape in slide.shapes:
            if shape.has_table:
                table_shapes.append(shape)
                continue
            if not shape.has_text_frame or not shape.text_frame.text.strip():
                continue
            if shape == slide.shapes.title:
                title = shape.text_frame.text.strip()
            else:
                text_shapes.append(shape)

        doc.add_heading(title or f"Slide {i}", level=1)

        for shape in text_shapes:
            for para in shape.text_frame.paragraphs:
                line = para.text.strip()
                if not line:
                    continue
                level = min(para.level or 0, 2)
                try:
                    p = doc.add_paragraph(style=BULLET_STYLES[level])
                except KeyError:
                    p = doc.add_paragraph(style="List Bullet")
                    p.paragraph_format.left_indent = DocxInches(0.25 * (level + 1))
                runs = [r for r in para.runs if r.text]
                if not runs:
                    p.add_run(line)
                for run in runs:
                    r = p.add_run(run.text)
                    r.bold = bool(run.font.bold)
                    r.italic = bool(run.font.italic)
                    r.underline = bool(run.font.underline)

        for shape in table_shapes:
            rows = [[cell.text.strip() for cell in row.cells] for row in shape.table.rows]
            rows = [r for r in rows if any(r)]
            if not rows:
                continue
            n_rows, n_cols = len(rows), max(len(r) for r in rows)
            word_table = doc.add_table(rows=n_rows, cols=n_cols)
            word_table.style = "Light Grid Accent 1"
            for r_idx, row in enumerate(rows):
                for c_idx in range(n_cols):
                    cell = word_table.cell(r_idx, c_idx)
                    cell.text = row[c_idx] if c_idx < len(row) else ""
                    if r_idx == 0:
                        for p in cell.paragraphs:
                            for run in p.runs:
                                run.bold = True

        if slide.has_notes_slide:
            notes_text = slide.notes_slide.notes_text_frame.text.strip()
            if notes_text:
                note_p = doc.add_paragraph()
                note_p.paragraph_format.space_before = DocxPt(8)
                run = note_p.add_run(f"Speaker notes: {notes_text}")
                run.italic = True
                run.font.size = DocxPt(9)

        if i < len(prs.slides):
            doc.add_page_break()

    apply_docx_style(doc, style)
    out_path = os.path.join(out_dir, "converted.docx")
    doc.save(out_path)
    return out_path


# ---------------------------------------------------------------- PDF -> DOCX
def pdf_to_docx(src_path, out_dir, style="clean"):
    reader = pypdf.PdfReader(src_path)
    n_pages = len(reader.pages)

    page_items = []
    for page in reader.pages:
        text = (page.extract_text() or "").strip()
        page_items.append(_reconstruct_paragraphs(text) if text else [])

    # A short line that repeats verbatim across most pages is a running
    # header/footer (page title, "Confidential", a date stamp, etc.), not
    # real content — repeating it once per page in the reconstructed
    # document just adds clutter. Only applies to documents long enough
    # that a real repeat pattern is meaningful, not a 2-page coincidence.
    line_counts = Counter()
    for items in page_items:
        seen_this_page = {t for t, _k in items if len(t) < 100}
        for t in seen_this_page:
            line_counts[t] += 1
    repeat_threshold = max(3, int(n_pages * 0.6))
    noisy_lines = {t for t, count in line_counts.items() if n_pages > 2 and count >= repeat_threshold}

    doc = Document()
    doc.add_heading("Converted from PDF", 0)
    MAX_IMAGES = 30
    image_count = 0

    for i, (page, items) in enumerate(zip(reader.pages, page_items), 1):
        if n_pages > 1:
            doc.add_heading(f"Page {i}", level=2)
        content_items = [(t, k) for t, k in items if t not in noisy_lines]
        if content_items:
            for para_text, kind in content_items:
                if not para_text:
                    continue
                if kind == "bullet":
                    doc.add_paragraph(para_text, style="List Bullet")
                elif kind == "heading":
                    doc.add_heading(para_text, level=3)
                else:
                    doc.add_paragraph(para_text)
        elif not items:
            doc.add_paragraph("[No extractable text on this page — likely a scanned image.]")
        # else: the page had only repeated header/footer noise and nothing
        # else — nothing worth showing, so leave it at just the page heading.

        if image_count < MAX_IMAGES:
            try:
                page_images = list(page.images)
            except Exception:
                page_images = []
            for img in page_images:
                if image_count >= MAX_IMAGES:
                    break
                try:
                    doc.add_picture(io.BytesIO(img.data), width=DocxInches(6))
                    image_count += 1
                except Exception:
                    continue  # a malformed/unsupported embedded image shouldn't sink the whole conversion

    apply_docx_style(doc, style)
    out_path = os.path.join(out_dir, "converted.docx")
    doc.save(out_path)
    return out_path


# ---------------------------------------------------------------- PDF -> PPTX
def pdf_to_pptx(src_path, out_dir, style=None):
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

    prs = Presentation()
    prs.slide_width = PptxInches(13.333)
    prs.slide_height = PptxInches(7.5)
    blank_layout = prs.slide_layouts[6]

    from PIL import Image

    for img_name in page_images:
        png_path = os.path.join(out_dir, img_name)
        # Rendered PDF pages are usually flat text/line-art, which PNG often
        # compresses better than JPEG's photo-oriented compression — but a
        # photo-heavy page can go the other way. Keep whichever is smaller
        # rather than assuming either format always wins.
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
        slide.shapes.add_picture(img_path, 0, 0, width=prs.slide_width, height=prs.slide_height)

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

    out_path = os.path.join(out_dir, "converted.pptx")
    prs.save(out_path)
    return out_path


# --------------------------------------------------------------- XLSX -> DOCX
def _format_cell_value(val):
    """openpyxl hands back raw Python values — a date becomes a datetime
    object, and floating-point arithmetic in the sheet often leaves noise
    like 3.140000000000001. str()'ing these directly is what the old code
    did, and it looked exactly as raw as that implies. This renders them the
    way a person actually reads a spreadsheet."""
    if val is None:
        return ""
    if isinstance(val, bool):
        return "TRUE" if val else "FALSE"
    if isinstance(val, datetime.datetime):
        if val.hour or val.minute or val.second:
            return val.strftime("%Y-%m-%d %H:%M")
        return val.strftime("%Y-%m-%d")
    if isinstance(val, datetime.date):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, float):
        if val == int(val):
            return str(int(val))
        return f"{val:.6f}".rstrip("0").rstrip(".")
    return str(val)


def xlsx_to_docx(src_path, out_dir, style="clean"):
    wb = openpyxl.load_workbook(src_path, data_only=True)

    sheet_rows = []
    max_cols = 0
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        merged_ranges = list(ws.merged_cells.ranges)
        rows = list(ws.iter_rows())
        if not merged_ranges:
            rows = [r for r in rows if any(c.value is not None and str(c.value).strip() for c in r)]
        sheet_rows.append((sheet_name, rows, merged_ranges))
        if rows:
            max_cols = max(max_cols, max(len(r) for r in rows))

    doc = Document()
    if max_cols > 6:
        # A wide sheet squeezed into a portrait page becomes unreadable —
        # landscape gives every column real room.
        section = doc.sections[0]
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width

    doc.add_heading("Converted from Excel", 0)
    for sheet_name, rows, merged_ranges in sheet_rows:
        doc.add_heading(sheet_name, level=1)
        if not rows:
            doc.add_paragraph("(Empty sheet)")
            continue
        n_cols = max(len(r) for r in rows)
        table = doc.add_table(rows=len(rows), cols=n_cols)
        table.style = "Light Grid Accent 1"
        for r_idx, row in enumerate(rows):
            for c_idx in range(n_cols):
                val = row[c_idx].value if c_idx < len(row) else None
                cell = table.cell(r_idx, c_idx)
                cell.text = _format_cell_value(val)
                if r_idx == 0:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            run.bold = True
        for mr in merged_ranges:
            try:
                table.cell(mr.min_row - 1, mr.min_col - 1).merge(table.cell(mr.max_row - 1, mr.max_col - 1))
            except IndexError:
                continue  # merge range falls outside the table we built — skip rather than crash

    apply_docx_style(doc, style)
    out_path = os.path.join(out_dir, "converted.docx")
    doc.save(out_path)
    return out_path


# --------------------------------------------------------------- DOCX -> XLSX
def docx_to_xlsx(src_path, out_dir):
    doc = Document(src_path)
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

    for i, table in enumerate(doc.tables, 1):
        ws = wb.create_sheet(title=f"Table {i}"[:31])
        for r_idx, row in enumerate(table.rows, 1):
            for c_idx, cell in enumerate(row.cells, 1):
                text = cell.text.strip()
                value = text if r_idx == 1 else coerce_numeric(text)
                ws.cell(row=r_idx, column=c_idx, value=value)
        for cell in ws[1]:
            cell.font = XlsxFont(bold=True)
        autosize_columns(ws, len(table.columns))

    # Capture the document's own paragraph text too, in its own sheet —
    # tables and surrounding prose commentary often coexist in a document,
    # and the old behavior silently dropped all of it whenever any table
    # was present.
    text_rows = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    if text_rows:
        ws = wb.create_sheet(title="Document Text")
        ws.column_dimensions["A"].width = 100
        for r, line in enumerate(text_rows, 1):
            ws.cell(row=r, column=1, value=line)

    if not wb.sheetnames:
        ws = wb.create_sheet(title="Sheet1")
        ws["A1"] = "(No content found in document)"

    out_path = os.path.join(out_dir, "converted.xlsx")
    wb.save(out_path)
    return out_path


CONVERTERS = {
    ("docx", "pdf"): docx_to_pdf,
    ("docx", "pptx"): docx_to_pptx,
    ("pptx", "docx"): pptx_to_docx,
    ("pdf", "docx"): pdf_to_docx,
    ("pdf", "pptx"): pdf_to_pptx,
    ("xlsx", "docx"): xlsx_to_docx,
    ("docx", "xlsx"): docx_to_xlsx,
}


def convert(src_path, from_fmt, to_fmt, out_dir, style=None):
    key = (from_fmt.lower(), to_fmt.lower())
    if key not in CONVERTERS:
        raise ConversionError(f"Unsupported conversion: {from_fmt} -> {to_fmt}")
    if style:
        return CONVERTERS[key](src_path, out_dir, style=style)
    return CONVERTERS[key](src_path, out_dir)


# ------------------------------------------------------- raw text -> PPTX
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
    title_layout = prs.slide_layouts[1]
    MAX_BULLETS_PER_SLIDE = 8

    for block in blocks:
        lines = [l.strip() for l in block.split("\n") if l.strip()]
        if not lines:
            continue
        title = _BULLET_RE.sub("", lines[0]).strip()
        body_lines = [_BULLET_RE.sub("", l).strip() for l in lines[1:]]

        slide = None
        tf = None
        count = 0
        for i, line in enumerate(body_lines):
            if slide is None or count >= MAX_BULLETS_PER_SLIDE:
                slide_title = title if slide is None else title + " (cont.)"
                slide = prs.slides.add_slide(title_layout)
                slide.shapes.title.text = slide_title[:120]
                tf = slide.placeholders[1].text_frame
                tf.clear()
                count = 0
            p = tf.paragraphs[0] if count == 0 else tf.add_paragraph()
            p.text = line
            count += 1
        if slide is None:
            slide = prs.slides.add_slide(title_layout)
            slide.shapes.title.text = title[:120]
            slide.placeholders[1].text_frame.clear()

    if not prs.slides:
        raise ConversionError("No usable text found")

    out_path = os.path.join(out_dir, "Presentation.pptx")
    prs.save(out_path)
    return out_path


def _set_run_font(run, bold=False):
    run.font.name = "Times New Roman"
    run.font.size = DocxPt(12)
    run.font.color.rgb = DocxRGBColor(0, 0, 0)
    run.bold = bold
    # Word can silently fall back to a different font for East-Asian text
    # runs unless this is set explicitly alongside the Latin font name above.
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), "Times New Roman")


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


def academic_essay_to_docx(payload, out_dir):
    """Builds a strictly-formatted academic Word document from AI-generated
    academic writing: Times New Roman, 12pt, black, double-spaced (2.0),
    1-inch margins, justified body paragraphs with a standard first-line
    indent, a centered page number in the footer, and a hanging-indent APA
    References list at the end. No AI involved — pure deterministic
    formatting of content the caller already generated."""
    doc = Document()

    section = doc.sections[0]
    section.left_margin = DocxInches(1)
    section.right_margin = DocxInches(1)
    section.top_margin = DocxInches(1)
    section.bottom_margin = DocxInches(1)
    _add_page_number_footer(doc)

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = DocxPt(12)
    normal.font.color.rgb = DocxRGBColor(0, 0, 0)
    normal.paragraph_format.line_spacing = 2.0

    title = (payload.get("title") or "Academic Response").strip()[:200]
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.line_spacing = 2.0
    _set_run_font(title_p.add_run(title), bold=True)

    def add_body_paragraph(text):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 2.0
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = DocxInches(0.5)
        _set_run_font(p.add_run(text))
        return p

    def add_section_heading(text, level=1):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 2.0
        if level <= 1:
            # APA Level 1: centered, bold
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _set_run_font(p.add_run(text), bold=True)
        elif level == 2:
            # APA Level 2: left-aligned, bold
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            _set_run_font(p.add_run(text), bold=True)
        else:
            # APA Level 3: left-aligned, bold italic
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(text)
            run.font.name = "Times New Roman"
            run.font.size = DocxPt(12)
            run.font.color.rgb = DocxRGBColor(0, 0, 0)
            run.bold = True
            run.italic = True
        return p

    sections = payload.get("sections")
    if sections:
        for sec in sections:
            add_section_heading(f"{sec['number']} {sec['heading']}", level=sec.get("level", 1))
            add_body_paragraph(sec["text"])
    else:
        text = (payload.get("text") or "").strip()
        paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
        for para in paragraphs or [text]:
            add_body_paragraph(para)

    references = payload.get("references") or []
    if references:
        # APA requires an alphabetical reference list — sort defensively here
        # rather than trusting whatever order the source list arrived in.
        references = sorted(references, key=lambda r: (r.get("text") or "").lower())
        doc.add_paragraph()
        ref_heading = doc.add_paragraph()
        ref_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        ref_heading.paragraph_format.line_spacing = 2.0
        _set_run_font(ref_heading.add_run("References"), bold=True)

        for ref in references:
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 2.0
            p.paragraph_format.left_indent = DocxInches(0.5)
            p.paragraph_format.first_line_indent = DocxInches(-0.5)
            ref_text = (ref.get("text") or "").strip()
            url = (ref.get("url") or "").strip()
            if ref_text:
                _set_run_font(p.add_run(ref_text + (" " if url else "")))
            if url:
                _set_run_font(p.add_run(url))

    out_path = os.path.join(out_dir, "Academic_Response.docx")
    doc.save(out_path)
    return out_path


# --------------------------------------------------------- extract raw text
def extract_text(src_path, ext):
    """Pull plain text out of an uploaded file for use in the AI tools.
    Supports the same four formats the converter already handles, plus .txt."""
    ext = ext.lower()
    if ext == "txt":
        with open(src_path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    if ext == "docx":
        doc = Document(src_path)
        parts = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                parts.append(" | ".join(c.text for c in row.cells))
        return "\n".join(parts)
    if ext == "pdf":
        reader = pypdf.PdfReader(src_path)
        return "\n".join((page.extract_text() or "") for page in reader.pages)
    if ext == "pptx":
        prs = Presentation(src_path)
        parts = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame and shape.text_frame.text.strip():
                    parts.append(shape.text_frame.text)
        return "\n".join(parts)
    if ext == "xlsx":
        wb = openpyxl.load_workbook(src_path, data_only=True)
        parts = []
        for sheet in wb.sheetnames:
            for row in wb[sheet].iter_rows(values_only=True):
                cells = [str(c) for c in row if c is not None]
                if cells:
                    parts.append(" | ".join(cells))
        return "\n".join(parts)
    raise ConversionError(f"Can't extract text from .{ext} files")
