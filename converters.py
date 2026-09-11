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
import copy
import subprocess
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
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
from docx.text.run import Run as DocxRun
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor as PptxRGBColor
from pptx.enum.dml import MSO_FILL_TYPE, MSO_COLOR_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_SHAPE
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
            items.append((_BULLET_RE.sub("", line).strip(), "bullet"))
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




def _extract_pdf_font_size_headings(pdf_path):
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
    page_index -> {line_text: level}."""
    all_sizes = Counter()
    page_lines = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_idx, page in enumerate(pdf.pages):
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
        if prev_size is not None and prev_size - s >= 2:
            level += 1
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


def _extract_pdf_ruled_tables(pdf_path, page_idx):
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
    false positive that corrupts a page of real prose."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if page_idx >= len(pdf.pages):
                return []
            return pdf.pages[page_idx].extract_tables()
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
    "minimal": {
        "bg": None, "title_fill": None,
        "title_font": "Calibri", "title_size": Pt(40), "title_bold": True,
        "title_color": PptxRGBColor(0x1A, 0x1A, 0x1A),
        "body_font": "Calibri", "body_size": Pt(26),
        "body_color": PptxRGBColor(0x33, 0x33, 0x33),
        "subhead_size": Pt(30), "caption_size": Pt(15),
    },
    "academic": {
        # Georgia (serif) replaced with a clean sans-serif — modern
        # presentation-design standards call for sans-serif body/title
        # text specifically because it reads better projected at a
        # distance; the academic template keeps its distinct navy color
        # identity, just not a serif face.
        "bg": None, "title_fill": None,
        "title_font": "Calibri", "title_size": Pt(40), "title_bold": True,
        "title_color": PptxRGBColor(0x1F, 0x3A, 0x5F),
        "body_font": "Calibri", "body_size": Pt(26),
        "body_color": PptxRGBColor(0x22, 0x22, 0x22),
        "subhead_size": Pt(30), "caption_size": Pt(15),
    },
    "bold": {
        "bg": PptxRGBColor(0x0A, 0x0A, 0x0A), "title_fill": None,
        "title_font": "Arial", "title_size": Pt(44), "title_bold": True,
        "title_color": PptxRGBColor(0xFF, 0xFF, 0xFF),
        "body_font": "Arial", "body_size": Pt(26),
        "body_color": PptxRGBColor(0xE8, 0xE8, 0xE8),
        "subhead_size": Pt(30), "caption_size": Pt(15),
    },
    "classic": {
        "bg": None, "title_fill": PptxRGBColor(0x1F, 0x38, 0x64),
        "title_font": "Calibri", "title_size": Pt(38), "title_bold": True,
        "title_color": PptxRGBColor(0xFF, 0xFF, 0xFF),
        "body_font": "Calibri", "body_size": Pt(26),
        "body_color": PptxRGBColor(0x22, 0x22, 0x22),
        "subhead_size": Pt(30), "caption_size": Pt(15),
    },
}


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


def _style_pptx_body_paragraph(p, style, size=None):
    p.font.name = style["body_font"]
    p.font.size = size if size is not None else style["body_size"]
    if style["body_color"] is not None:
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
            if run.font.size is None:
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
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"{base}.{target_format}")
    if not os.path.exists(out_path):
        raise ConversionError(f"LibreOffice conversion failed: {result.stderr or result.stdout}")
    return out_path


# ---------------------------------------------------------------- DOCX -> PDF
def docx_to_pdf(src_path, out_dir, style=None):
    return _soffice_convert(src_path, "pdf", out_dir)


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



def _effective_paragraph_alignment(paragraph):
    """Returns a paragraph's alignment, falling back to its paragraph
    style's own alignment when not set directly — the same style-carries-
    the-real-formatting gap as run-level bold/italic, just for paragraph
    alignment: a custom style (a 'Caption' or 'Quote' style that bakes in
    centering, for instance) leaves paragraph.alignment as None even
    though the paragraph visibly renders centered."""
    if paragraph.alignment is not None:
        return paragraph.alignment
    try:
        return paragraph.style.paragraph_format.alignment
    except Exception:
        return None


def _effective_run_format(run):
    """Returns (bold, italic, underline, hex_color, size_pt) for a run,
    falling back to its referenced character style when the run has no
    direct formatting of its own. Word's built-in 'Strong'/'Emphasis'
    styles — and most custom character styles — carry formatting in the
    STYLE definition rather than as direct run formatting, so reading
    run.bold/font.color/font.size directly returns None for all five of
    these even though the text visibly renders styled, silently losing
    the formatting on conversion. Direct run formatting always wins when
    present; the style is only consulted for whichever is still unset."""
    bold, italic, underline = run.bold, run.italic, run.underline
    hex_color = _safe_hex_color(run.font.color)
    size_pt = run.font.size.pt if run.font.size is not None else None
    if bold is None or italic is None or underline is None or hex_color is None or size_pt is None:
        try:
            style_font = run.style.font
            if bold is None:
                bold = style_font.bold
            if italic is None:
                italic = style_font.italic
            if underline is None:
                underline = style_font.underline
            if hex_color is None:
                hex_color = _safe_hex_color(style_font.color)
            if size_pt is None and style_font.size is not None:
                size_pt = style_font.size.pt
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
        for t in r_el.findall(tag):
            if any(anc.tag in (f"{{{_WPS_NS}}}txbx", qn("w:txbxContent")) for anc in t.iterancestors()):
                continue
            results.append(t.text or "")
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
        return numpr_from(paragraph.style.element.find(qn("w:pPr")))
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


def _pptx_fix_apply_sizing_in_place(shape, size, bold_title=False, font_name=None):
    """Fixes font size and disables autofit on an existing shape without
    restructuring it — the safe path for any slide too complex to split
    (has an image, chart, table, or more than one text box), so a table
    or picture on that slide is never touched. font_name is only passed
    when a custom template's brand font should replace whatever the
    source file's own text used."""
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    for p in tf.paragraphs:
        for r in p.runs:
            r.font.size = size
            if bold_title:
                r.font.bold = True
            if font_name:
                r.font.name = font_name


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


def _pptx_fix_add_agenda_slide(prs, title_size, body_size, title_font, body_font):
    """Inserts an agenda slide right after the title slide, listing each
    distinct section's own title verbatim — built entirely from titles
    the deck already has, nothing invented. Only added once the deck has
    enough distinct sections to actually be worth summarizing, and
    skipped (rather than spilling onto a second agenda slide) once there
    are more sections than a single overview slide can usefully list —
    an agenda that itself needs "(cont.)" defeats its own purpose."""
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
    if agenda_slide.shapes.title is not None:
        agenda_slide.shapes.title.text = "Agenda"
        for p in agenda_slide.shapes.title.text_frame.paragraphs:
            for r in p.runs:
                r.font.size = title_size
                r.font.bold = True
                if title_font:
                    r.font.name = title_font
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


def pptx_to_pptx(src_path, out_dir, style=None, template_path=None):
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

    title_size, body_size, title_font, body_font = _PPTX_FIX_TITLE_SIZE, _PPTX_FIX_BODY_SIZE, None, None
    if template_path:
        template_prs = _safe_load(Presentation, template_path)
        template_content_layout = _find_best_pptx_layout(template_prs, "content")
        template_style = _extract_template_style(template_prs, template_content_layout)
        title_size, body_size = template_style["title_size"], template_style["body_size"]
        title_font, body_font = template_style["title_font"], template_style["body_font"]
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

    # ---- Pass 1: fix every title in place, and classify each slide ----
    plan = []  # (slide, is_simple, body_shape_or_None)
    for slide in prs.slides:
        title_shape = slide.shapes.title
        if title_shape is not None and title_shape.has_text_frame:
            _pptx_fix_apply_sizing_in_place(title_shape, title_size, bold_title=True, font_name=title_font)

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
                    _pptx_fix_apply_sizing_in_place(s, body_size, font_name=body_font)
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
        # Avoid "Title (cont.) (cont.)" when the slide being split was
        # itself already a manually-made continuation slide in the source
        # deck — the new continuation slides this produces use the same
        # base title suffix, not a doubled one.
        if base_title.endswith("(cont.)"):
            base_title = base_title[: -len("(cont.)")].strip()
        bullets = _pptx_fix_extract_bullets(body_shape)
        expanded = []
        for level, runs, numbered in bullets:
            expanded.extend(_pptx_fix_split_bullet_if_long(level, runs, numbered))

        chars_per_line = _pptx_fix_chars_per_line(body_shape.width, _PPTX_FIX_CAL_SIZE_PT)
        max_lines = _pptx_fix_max_effective_lines(body_shape.height, _PPTX_FIX_CAL_SIZE_PT)

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
                new_slide.shapes.title.text = group_titles[gi] or f"{base_title} (cont.)"
                for p in new_slide.shapes.title.text_frame.paragraphs:
                    for r in p.runs:
                        r.font.size = title_size
                        r.font.bold = True
                        if title_font:
                            r.font.name = title_font
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
            _pptx_fix_move_slide_to(prs, new_slide, insert_position)
            insert_position += 1
        insertions.append((idx, len(groups) - 1))

    _pptx_fix_add_agenda_slide(prs, title_size, body_size, title_font, body_font)
    _pptx_fix_apply_smart_emphasis(prs)
    _pptx_fix_reposition_offslide_pictures(prs)
    _pptx_fix_apply_transitions(prs)
    _pptx_fix_force_white_bg_black_text(prs)

    out_path = os.path.join(out_dir, "fixed.pptx")
    os.makedirs(out_dir, exist_ok=True)
    prs.save(out_path)
    return out_path


def docx_to_pptx(src_path, out_dir, style="minimal", template_path=None):
    """style selects one of this file's own built-in visual themes
    (minimal/academic/bold/classic) and is ignored when template_path is
    given. template_path, if provided, is a path to an uploaded .potx (or
    .pptx used as a template) — a corporate brand template — whose own
    masters, layouts, fonts, and colors are used instead, so the
    converted content maps into the uploaded template's own pre-styled
    layouts rather than one of this file's built-in looks."""
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
        template_style = PPTX_TEMPLATES.get(style, PPTX_TEMPLATES["minimal"])
        prs = Presentation()
        prs.slide_width = PptxInches(13.333)
        prs.slide_height = PptxInches(7.5)
        title_layout = prs.slide_layouts[1]  # title + content
        blank_layout = prs.slide_layouts[6]

    doc = _safe_load(Document, src_path)
    numbering_formats = _get_docx_numbering_formats(doc)

    MAX_LINES_PER_SLIDE = 10  # estimated wrapped-line budget, not a flat bullet count (see
    # _estimate_line_count) — accounts for bullets of very different lengths now that long
    # paragraphs get split into shorter, sentence-level bullets rather than staying as one
    state = {"slide": None, "body_tf": None, "bullet_count": 0, "lines_used": 0}

    def new_slide(title_text):
        s = prs.slides.add_slide(title_layout)
        s.shapes.title.text = title_text or "Untitled"
        _style_pptx_slide(s, template_style)
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
        state["slide"], state["body_tf"], state["bullet_count"], state["lines_used"] = s, tf, 0, 0
        return s, tf

    def add_bullet(para, level=0, numbered=False):
        if state["slide"] is None:
            new_slide("Overview")

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
        if not numbered and len(full_text.split()) > _SENTENCE_SPLIT_WORD_THRESHOLD:
            groups = _split_runs_at_sentence_boundaries(runs_data)
        else:
            groups = [runs_data] if runs_data else []

        pptx_align = _docx_align_to_pptx(_effective_paragraph_alignment(para))
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
                current_title = state["slide"].shapes.title.text or "Overview"
                new_slide(current_title + " (cont.)")
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

            captured = []  # (pptx_run, bold, italic, underline, is_link, hex_color, size_pt, is_deleted)
            for text, bold, italic, underline, is_link, address, hex_color, size_pt, is_deleted in group:
                r = p.add_run()
                r.text = text
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

    def add_subheading(text):
        """Heading 2/3 in the source document reads as a subsection within
        the current topic, not a brand new topic — giving it a whole new
        slide (the old behavior for every heading level) fragmented what's
        usually meant to be one cohesive slide into several thin ones. This
        keeps it on the current slide as a bolded standalone line instead."""
        if state["slide"] is None:
            new_slide(text)
            return
        # A subheading's larger font and extra space-before take up more
        # visual room than its line count alone suggests, so it's weighted
        # a bit heavier than a plain estimated-line count would give it.
        subhead_lines = _estimate_line_count(text) + 1
        if state["bullet_count"] > 0 and state["lines_used"] + subhead_lines > MAX_LINES_PER_SLIDE:
            current_title = state["slide"].shapes.title.text or "Overview"
            new_slide(current_title + " (cont.)")
        body_tf = state["body_tf"]
        if body_tf.paragraphs[0].text == "" and len(body_tf.paragraphs) == 1 and state["bullet_count"] == 0:
            p = body_tf.paragraphs[0]
        else:
            p = body_tf.add_paragraph()
        p.level = 0
        p.text = text
        # A subheading is a section divider, not another item in the list
        # — keeping its bullet marker made it look like just another
        # bullet rather than something that visually "breaks up content"
        # the way a real section header should.
        pPr = p._p.get_or_add_pPr()
        pPr.append(pPr.makeelement(qn("a:buNone"), {}))
        if state["bullet_count"] > 0:
            p.space_before = Pt(18)
        _style_pptx_body_paragraph(p, template_style, size=template_style["subhead_size"])
        for run in p.runs:
            run.font.bold = True
        state["bullet_count"] += 1
        state["lines_used"] += subhead_lines

    def add_table_slide(table):
        src_rows = [list(row.cells) for row in table.rows]
        rows = [[_full_cell_text(cell).strip() for cell in row] for row in src_rows]
        keep = [i for i, r in enumerate(rows) if any(r)]
        if not keep:
            return
        rows = [rows[i] for i in keep]
        src_rows = [src_rows[i] for i in keep]
        n_rows, n_cols = len(rows), max(len(r) for r in rows)
        s = prs.slides.add_slide(blank_layout)
        left, top = PptxInches(0.6), PptxInches(0.6)
        width, height = prs.slide_width - PptxInches(1.2), prs.slide_height - PptxInches(1.2)
        gtable = s.shapes.add_table(n_rows, n_cols, left, top, width, height).table

        merges = _find_docx_merges(src_rows, n_cols)
        # A merged span's duplicated cells would otherwise each get the same
        # text written before merging, concatenating it multiple times into
        # the final merged cell — only the top-left origin of each span
        # should actually receive the text.
        skip_cells = {(r, c) for (min_r, min_c, max_r, max_c) in merges
                      for r in range(min_r, max_r + 1) for c in range(min_c, max_c + 1)
                      if (r, c) != (min_r, min_c)}

        for r_idx, row in enumerate(rows):
            for c_idx in range(n_cols):
                if (r_idx, c_idx) in skip_cells:
                    continue
                cell = gtable.cell(r_idx, c_idx)
                cell.text = row[c_idx] if c_idx < len(row) else ""
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

    def add_chart_slide(title, categories, series_list, chart_type):
        if len(categories) < 1 or len(series_list) < 1:
            return
        s = prs.slides.add_slide(title_layout)
        s.shapes.title.text = title or "Chart"
        _style_pptx_slide(s, template_style)
        left, top = PptxInches(1.0), PptxInches(1.7)
        width, height = prs.slide_width - PptxInches(2.0), prs.slide_height - PptxInches(2.4)

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
        state["slide"], state["body_tf"], state["bullet_count"], state["lines_used"] = None, None, 0, 0

    def add_shape_slide(mso_shape, fill_hex, shape_text, width_emu, height_emu):
        s = prs.slides.add_slide(blank_layout)
        # Preserve the shape's own aspect ratio when it had real dimensions,
        # scaled up to a reasonable on-slide size, rather than stretching
        # every recreated shape to one fixed box regardless of its source
        # proportions (a wide arrow shouldn't come out square).
        max_w, max_h = prs.slide_width - PptxInches(2), prs.slide_height - PptxInches(2)
        if width_emu and height_emu:
            scale = min(max_w / width_emu, max_h / height_emu, 4.0)
            w, h = int(width_emu * scale), int(height_emu * scale)
        else:
            w, h = PptxInches(4), PptxInches(2)
        left, top = int((prs.slide_width - w) / 2), int((prs.slide_height - h) / 2)
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
        text = _full_paragraph_text(para).strip()
        if not text:
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
            new_slide(text)
        elif heading_match:
            add_subheading(text)
        else:
            direct_list_info = _get_paragraph_direct_list_info(para, numbering_formats)
            if direct_list_info is not None:
                level, numbered = direct_list_info
                level = max(0, min(level, 4))
            else:
                level = 0
                numbered = False
                if "list" in para_style_name:
                    m = re.search(r"(\d+)", para_style_name)
                    if m:
                        level = max(0, min(int(m.group(1)) - 1, 4))
                    numbered = "number" in para_style_name
            add_bullet(para, level=level, numbered=numbered)

    # Only needed as a last-resort guarantee that the presentation has at
    # least one slide (e.g. a source document that was nothing but a
    # single image) — if the document already produced real slides and
    # simply ended on a table/chart/image, state["slide"] being None just
    # means nothing followed it, not that a slide is missing. Creating one
    # anyway produced a pointless empty trailing slide a user would have
    # to notice and delete themselves.
    if state["slide"] is None and len(prs.slides) == 0:
        new_slide("Untitled Document")

    footnotes = _get_docx_footnotes(doc)
    if footnotes:
        new_slide("Footnotes")
        for i, note in enumerate(footnotes, 1):
            add_bullet_text = f"{i}. {note}"
            body_tf = state["body_tf"]
            p = body_tf.paragraphs[0] if state["bullet_count"] == 0 else body_tf.add_paragraph()
            p.text = add_bullet_text
            _style_pptx_body_paragraph(p, template_style)
            state["bullet_count"] += 1

    endnotes = _get_docx_endnotes(doc)
    if endnotes:
        new_slide("Endnotes")
        for i, note in enumerate(endnotes, 1):
            body_tf = state["body_tf"]
            p = body_tf.paragraphs[0] if state["bullet_count"] == 0 else body_tf.add_paragraph()
            p.text = f"{i}. {note}"
            _style_pptx_body_paragraph(p, template_style)
            state["bullet_count"] += 1

    comments = _get_docx_comments(doc)
    if comments:
        new_slide("Comments")
        for author, comment_text in comments:
            body_tf = state["body_tf"]
            p = body_tf.paragraphs[0] if state["bullet_count"] == 0 else body_tf.add_paragraph()
            p.text = f"{author}: {comment_text}"
            _style_pptx_body_paragraph(p, template_style)
            state["bullet_count"] += 1

    header_footer = _get_docx_header_footer_text(doc)
    if header_footer:
        new_slide("Header/Footer")
        for line in header_footer:
            body_tf = state["body_tf"]
            p = body_tf.paragraphs[0] if state["bullet_count"] == 0 else body_tf.add_paragraph()
            p.text = line
            _style_pptx_body_paragraph(p, template_style)
            state["bullet_count"] += 1

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


def pptx_to_docx(src_path, out_dir, style="clean"):
    prs = _safe_load(Presentation, src_path)
    doc = Document()
    # python-docx's own default template ships a <w:zoom val="bestFit"/>
    # missing the "percent" attribute the OOXML schema actually requires
    # on it — confirmed directly against the schema and present even in
    # a completely unmodified new Document(), unrelated to anything in
    # this function. Fixed here rather than left in, since it's a
    # one-line, safe correction now that it's been found.
    zoom_el = doc.settings.element.find(qn("w:zoom"))
    if zoom_el is not None and zoom_el.get(qn("w:percent")) is None:
        zoom_el.set(qn("w:percent"), "100")
    doc.add_heading("Slide Handout", 0)
    BULLET_STYLES = ["List Bullet", "List Bullet 2", "List Bullet 3"]
    NUMBER_STYLES = ["List Number", "List Number 2", "List Number 3"]

    # A quick pre-pass just for titles, to build a real, clickable table
    # of contents up front before any slide content is added — skipped
    # for a short deck, where flipping through a handful of headings is
    # faster than reading a table of contents for them.
    slide_titles = []
    for i, slide in enumerate(prs.slides, 1):
        title_shape = slide.shapes.title
        if title_shape is not None and title_shape.has_text_frame and title_shape.text_frame.text.strip():
            slide_titles.append(_sanitize_xml_text(title_shape.text_frame.text.strip()))
        else:
            slide_titles.append(None)
    real_titles = [t for t in slide_titles if t]
    if len(real_titles) >= 4:
        toc_heading = doc.add_paragraph()
        toc_heading_run = toc_heading.add_run("Contents")
        toc_heading_run.bold = True
        toc_heading_run.font.size = DocxPt(14)
        _add_docx_toc_field(doc, real_titles)
        doc.add_page_break()

    for i, slide in enumerate(prs.slides, 1):
        title = None
        text_shapes = []
        table_shapes = []
        chart_shapes = []
        image_shapes = []
        smartart_found = False
        smartart_lines = []
        for found, text in _iter_smartart_fallback_text(slide):
            if found:
                smartart_found = True
            if text:
                smartart_lines.append(text)
        for shape in _iter_flat_shapes(slide.shapes):
            if shape.has_table:
                table_shapes.append(shape)
                continue
            if getattr(shape, "has_chart", False):
                chart_shapes.append(shape)
                continue
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                image_shapes.append(shape)
                continue
            if not shape.has_text_frame or not shape.text_frame.text.strip():
                continue
            if shape == slide.shapes.title:
                title = _sanitize_xml_text(shape.text_frame.text.strip())
            else:
                text_shapes.append(shape)

        doc.add_heading(title or f"Slide {i}", level=1)
        if title:
            # A small, secondary "Slide N" label lets a reader cross-
            # reference back to the exact slide in the original
            # presentation — only shown when the heading itself is a
            # real title, since it would just repeat the heading's own
            # text for a slide that has no title of its own.
            slide_num_p = doc.add_paragraph()
            slide_num_run = slide_num_p.add_run(f"Slide {i}")
            slide_num_run.italic = True
            slide_num_run.font.size = DocxPt(9)
            slide_num_run.font.color.rgb = DocxRGBColor(0x80, 0x80, 0x80)
            slide_num_p.paragraph_format.space_after = DocxPt(2)

        for shape in image_shapes:
            try:
                doc.add_picture(io.BytesIO(shape.image.blob), width=DocxInches(4))
            except Exception:
                continue  # a malformed/unsupported embedded image shouldn't sink the whole conversion

        for shape in text_shapes:
            for para in shape.text_frame.paragraphs:
                line = _sanitize_xml_text(para.text.strip())
                if not line:
                    continue
                level = min(para.level or 0, 2)
                pPr = para._p.find(qn("a:pPr"))
                is_numbered = pPr is not None and pPr.find(qn("a:buAutoNum")) is not None
                style_list = NUMBER_STYLES if is_numbered else BULLET_STYLES
                try:
                    p = doc.add_paragraph(style=style_list[level])
                except KeyError:
                    p = doc.add_paragraph(style="List Number" if is_numbered else "List Bullet")
                    p.paragraph_format.left_indent = DocxInches(0.25 * (level + 1))
                docx_align = _pptx_align_to_docx(para.alignment)
                if docx_align is not None:
                    p.alignment = docx_align
                runs = [r for r in para.runs if r.text]
                if not runs:
                    p.add_run(line)
                for run in runs:
                    run_text = _sanitize_xml_text(run.text)
                    address = None
                    try:
                        address = run.hyperlink.address
                    except Exception:
                        pass
                    if address:
                        _add_docx_hyperlink(
                            p, address, run_text,
                            bold=bool(run.font.bold), italic=bool(run.font.italic),
                        )
                    else:
                        r = p.add_run(run_text)
                        r.bold = bool(run.font.bold)
                        r.italic = bool(run.font.italic)
                        r.underline = bool(run.font.underline)
                        hex_color = _safe_hex_color(run.font.color)
                        if hex_color:
                            try:
                                r.font.color.rgb = DocxRGBColor.from_string(hex_color)
                            except Exception:
                                pass
                        if run.font.size is not None:
                            r.font.size = DocxPt(run.font.size.pt)

        for shape in table_shapes:
            src_rows = [list(row.cells) for row in shape.table.rows]
            has_merges = any(cell.is_merge_origin for row in src_rows for cell in row)
            rows = [[_sanitize_xml_text(cell.text.strip()) for cell in row] for row in src_rows]
            if has_merges:
                keep = list(range(len(rows)))
            else:
                keep = [i for i, r in enumerate(rows) if any(r)]
            if not keep:
                continue
            rows = [rows[i] for i in keep]
            src_rows = [src_rows[i] for i in keep]
            n_rows, n_cols = len(rows), max(len(r) for r in rows)
            word_table = doc.add_table(rows=n_rows, cols=n_cols)
            word_table.style = "Light Grid Accent 1"
            merge_spans = []
            for r_idx, (doc_row, row) in enumerate(zip(word_table.rows, rows)):
                doc_cells = doc_row.cells
                for c_idx in range(n_cols):
                    src_cell = src_rows[r_idx][c_idx] if c_idx < len(src_rows[r_idx]) else None
                    if src_cell is not None and src_cell.is_spanned and not src_cell.is_merge_origin:
                        continue  # covered by a merge origin elsewhere — filled in via the merge below
                    cell = doc_cells[c_idx]
                    cell.text = row[c_idx] if c_idx < len(row) else ""
                    if r_idx == 0:
                        for p in cell.paragraphs:
                            for run in p.runs:
                                run.bold = True
                    if src_cell is not None:
                        try:
                            if src_cell.fill.type == MSO_FILL_TYPE.SOLID:
                                hex_color = _safe_hex_color(src_cell.fill.fore_color)
                                if hex_color:
                                    _set_docx_cell_shading(cell, hex_color)
                        except Exception:
                            pass
                        if src_cell.is_merge_origin and (src_cell.span_height > 1 or src_cell.span_width > 1):
                            merge_spans.append((r_idx, c_idx, r_idx + src_cell.span_height - 1, c_idx + src_cell.span_width - 1))
            for min_r, min_c, max_r, max_c in merge_spans:
                try:
                    word_table.cell(min_r, min_c).merge(word_table.cell(max_r, max_c))
                except Exception:
                    pass

        for shape in chart_shapes:
            try:
                chart = shape.chart
                plot = chart.plots[0]
                categories = [_sanitize_xml_text(str(cat)) for cat in plot.categories]
                series_list = list(plot.series)
            except Exception:
                continue
            if not series_list:
                continue
            caption = doc.add_paragraph()
            title_text = None
            try:
                if chart.has_title:
                    title_text = _sanitize_xml_text(chart.chart_title.text_frame.text.strip())
            except Exception:
                pass
            run = caption.add_run(f"Chart: {title_text}" if title_text else "Chart data")
            run.italic = True
            n_rows = len(categories) + 1
            n_cols = len(series_list) + 1
            chart_table = doc.add_table(rows=n_rows, cols=n_cols)
            chart_table.style = "Light Grid Accent 1"
            chart_table.cell(0, 0).text = ""
            for s_idx, series in enumerate(series_list, 1):
                chart_table.cell(0, s_idx).text = _sanitize_xml_text(series.name or f"Series {s_idx}")
            for cat_idx, cat_name in enumerate(categories, 1):
                chart_table.cell(cat_idx, 0).text = cat_name
                for s_idx, series in enumerate(series_list, 1):
                    values = list(series.values)
                    val = values[cat_idx - 1] if cat_idx - 1 < len(values) else None
                    chart_table.cell(cat_idx, s_idx).text = _format_cell_value(val)
            for cell in chart_table.rows[0].cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.bold = True

        if smartart_lines:
            label_p = doc.add_paragraph()
            label_p.paragraph_format.space_before = DocxPt(8)
            label_run = label_p.add_run("Diagram contents:")
            label_run.italic = True
            label_run.bold = True
            for line in smartart_lines:
                doc.add_paragraph(line, style="List Bullet")
        elif smartart_found:
            note_p = doc.add_paragraph()
            note_p.paragraph_format.space_before = DocxPt(8)
            run = note_p.add_run(
                "[This slide contains a diagram (SmartArt) whose content could not be extracted.]"
            )
            run.italic = True
            run.font.size = DocxPt(9)

        if slide.has_notes_slide:
            notes_text = _sanitize_xml_text(slide.notes_slide.notes_text_frame.text.strip())
            if notes_text:
                note_p = doc.add_paragraph()
                note_p.paragraph_format.space_before = DocxPt(8)
                label_run = note_p.add_run("Speaker Notes\n")
                label_run.bold = True
                label_run.font.size = DocxPt(10)
                text_run = note_p.add_run(notes_text)
                text_run.italic = True
                text_run.font.size = DocxPt(10)
                _add_docx_callout_style(note_p)

        if i < len(prs.slides):
            doc.add_page_break()

    apply_docx_style(doc, style)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "converted.docx")
    doc.save(out_path)
    return out_path


# ---------------------------------------------------------------- PDF -> DOCX
def pdf_to_docx(src_path, out_dir, style="clean"):
    reader = _safe_load(pypdf.PdfReader, src_path)
    if reader.is_encrypted:
        raise ConversionError("This PDF is password-protected and can't be converted until it's unlocked.")
    n_pages = len(reader.pages)
    os.makedirs(out_dir, exist_ok=True)  # needed early: OCR (if any page requires it,
    # below) writes a scratch rasterized PNG into out_dir well before the
    # final document save, which is the only place this used to be called

    font_headings_by_page, _body_size = _extract_pdf_font_size_headings(src_path)

    # Ruled tables (real grid lines, not just aligned text — see
    # _extract_pdf_ruled_tables for why the more permissive strategy
    # isn't used) are pulled out per page so they can be rebuilt as real
    # Word tables rather than left to come through the ordinary text
    # flow as a garbled run of individual cell values.
    MAX_TABLE_PAGES = 40  # a table-detection pass opens the PDF again per page; capped for very long documents
    tables_by_page = {}
    if n_pages <= MAX_TABLE_PAGES:
        for page_idx in range(n_pages):
            tables = _extract_pdf_ruled_tables(src_path, page_idx)
            if tables:
                tables_by_page[page_idx] = tables

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
        page_items.append(_reconstruct_paragraphs(text, page_font_headings) if text else [])

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
    doc.add_heading("Converted from PDF", 0)

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
    if len(level1_headings) >= 4:
        toc_heading = doc.add_paragraph()
        toc_heading_run = toc_heading.add_run("Contents")
        toc_heading_run.bold = True
        toc_heading_run.font.size = DocxPt(14)
        _add_docx_toc_field(doc, level1_headings)
        doc.add_page_break()

    MAX_IMAGES = 30
    image_count = 0
    MAX_OCR_PAGES = 15  # OCR is comparatively expensive (rasterize + tesseract
    # per page); this covers a few scanned pages mixed into an otherwise
    # text-based PDF, or a fully-scanned document up to a reasonable length,
    # without unbounded processing time on a very long scanned file.
    ocr_page_count = 0

    for i, (page, items) in enumerate(zip(reader.pages, page_items), 1):
        if n_pages > 1:
            doc.add_heading(f"Page {i}", level=2)
        page_links = _extract_pdf_page_links(page)
        unmatched_links = list(page_links)
        page_idx = i - 1
        content_items = [
            (t, k) for item_idx, (t, k) in enumerate(items)
            if t not in noisy_lines or (page_idx, item_idx) in exempt_occurrences
        ]
        if content_items:
            for para_text, kind in content_items:
                if not para_text:
                    continue
                candidates = [(lt, url) for lt, url in page_links if lt in para_text]
                if kind == "bullet":
                    _p, used = _add_docx_paragraph_with_links(doc, para_text, candidates, style="List Bullet")
                elif kind.startswith("heading"):
                    # "Page N" (when present) already occupies level 2, so
                    # the PDF's own content headings nest under it starting
                    # at level 3 — pdf_level 1 (the document's biggest,
                    # most prominent heading) becomes Word level 3, and so
                    # on, rather than competing with "Page N" for the same
                    # level or going shallower than it.
                    pdf_level = int(kind[len("heading"):]) if kind[len("heading"):].isdigit() else 3
                    doc.add_heading(para_text, level=min(pdf_level + 2, 9))
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
                    elif kind.startswith("heading"):
                        doc.add_heading(para_text, level=3)
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

        for table_rows in tables_by_page.get(page_idx, []):
            # Drop fully-empty rows (a spacer row between ruled sections,
            # common in PDFs that draw extra grid lines for visual
            # padding) rather than rendering a table with blank rows in it.
            clean_rows = [
                [(_sanitize_xml_text(str(cell).strip()) if cell else "") for cell in row]
                for row in table_rows if any(cell and str(cell).strip() for cell in row)
            ]
            if not clean_rows:
                continue
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
def pdf_to_pptx(src_path, out_dir, style=None):
    # Checked upfront, before rasterization — each page needs rendering
    # plus roughly 1.2s of sequential OCR (measured directly), so an
    # unbounded page count could push total processing well past what
    # typical web infrastructure allows before timing out silently with
    # no useful error. Reading the page count via pypdf first is fast
    # (no rendering involved) and avoids wasting the rasterization work
    # entirely on a file that's going to be rejected anyway.
    MAX_PDF_PPTX_PAGES = 50
    try:
        page_count = len(_safe_load(pypdf.PdfReader, src_path).pages)
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

    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "converted.pptx")
    prs.save(out_path)
    return out_path


# --------------------------------------------------------------- XLSX -> DOCX
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
    if isinstance(val, datetime.datetime):
        if val.hour or val.minute or val.second:
            return val.strftime("%Y-%m-%d %H:%M")
        return val.strftime("%Y-%m-%d")
    if isinstance(val, datetime.date):
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
            for symbol in ("$", "£", "€", "¥"):
                if symbol in fmt:
                    if work_val < 0:
                        result = f"-{symbol}{-work_val:,.{decimals}f}"
                    else:
                        result = f"{symbol}{work_val:,.{decimals}f}"
                    break
            if result is None and ("0" in fmt or "#" in fmt):
                result = f"{work_val:,.{decimals}f}" if "," in fmt else f"{work_val:.{decimals}f}"
        if result is not None:
            return f"({result})" if use_parens else result
    if isinstance(val, float):
        if val == int(val):
            return str(int(val))
        return f"{val:.6f}".rstrip("0").rstrip(".")
    return str(val)


def xlsx_to_docx(src_path, out_dir, style="clean"):
    wb = _safe_load(openpyxl.load_workbook, src_path, data_only=True)
    # A formula cell in a workbook that's never been opened in a real
    # spreadsheet app (generated by a script, exported from a database) has
    # no cached result — data_only=True silently returns None for it, which
    # renders as a misleadingly blank cell with no sign a formula was ever
    # there. Loading a second copy without data_only lets an uncalculated
    # formula fall back to showing its actual formula text instead.
    wb_formulas = _safe_load(openpyxl.load_workbook, src_path, data_only=False)

    sheet_rows = []
    max_cols = 0
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
        sheet_rows.append((sheet_name, rows, formula_rows, merged_ranges, ws))
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
    for sheet_name, rows, formula_rows, merged_ranges, ws in sheet_rows:
        doc.add_heading(sheet_name, level=1)
        if not rows:
            doc.add_paragraph("(Empty sheet)")
            continue
        n_cols = max(len(r) for r in rows)
        table = doc.add_table(rows=len(rows), cols=n_cols)
        table.style = "Light Grid Accent 1"
        cell_comments = []
        for r_idx, (doc_row, row) in enumerate(zip(table.rows, rows)):
            doc_cells = doc_row.cells
            for c_idx in range(n_cols):
                src_cell = row[c_idx] if c_idx < len(row) else None
                val = src_cell.value if src_cell is not None else None
                if val is None and c_idx < len(formula_rows[r_idx]) and formula_rows[r_idx][c_idx].data_type == "f":
                    val = formula_rows[r_idx][c_idx].value  # uncalculated formula — show the formula itself, not blank
                cell = doc_cells[c_idx]
                number_format = src_cell.number_format if src_cell is not None else None
                cell.text = _format_cell_value(val, number_format)
                if r_idx == 0:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            run.bold = True
                if src_cell is not None:
                    fill_hex = _xlsx_conditional_fill_hex(ws, src_cell) or _xlsx_cell_fill_hex(src_cell)
                    if fill_hex:
                        _set_docx_cell_shading(cell, fill_hex)
                    if src_cell.comment is not None:
                        note_text = (src_cell.comment.text or "").strip()
                        if note_text:
                            author = (src_cell.comment.author or "").strip() or "Comment"
                            cell_comments.append((src_cell.coordinate, author, note_text))
        for mr in merged_ranges:
            try:
                table.cell(mr.min_row - 1, mr.min_col - 1).merge(table.cell(mr.max_row - 1, mr.max_col - 1))
            except IndexError:
                continue  # merge range falls outside the table we built — skip rather than crash
        if cell_comments:
            note_heading = doc.add_paragraph()
            note_heading.paragraph_format.space_before = DocxPt(8)
            note_run = note_heading.add_run("Cell comments:")
            note_run.bold = True
            note_run.italic = True
            for coord, author, note_text in cell_comments:
                doc.add_paragraph(f"{coord} ({author}): {note_text}", style="List Bullet")

        charts = getattr(ws, "_charts", None) or []
        for chart in charts:
            chart_title = _xlsx_chart_title(chart)
            label = f'Chart: "{chart_title}"' if chart_title else "Chart (based on the data above)"
            chart_p = doc.add_paragraph()
            chart_p.paragraph_format.space_before = DocxPt(8)
            chart_run = chart_p.add_run(label)
            chart_run.italic = True

    apply_docx_style(doc, style)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "converted.docx")
    doc.save(out_path)
    return out_path


# --------------------------------------------------------------- DOCX -> XLSX
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

    for i, table in enumerate(doc.tables, 1):
        ws = wb.create_sheet(title=f"Table {i}"[:31])
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
                value = text if r_idx == 1 else coerce_numeric(text)
                xlsx_cell = ws.cell(row=r_idx, column=c_idx, value=value)
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

    # Capture the document's own paragraph text too, in its own sheet —
    # tables and surrounding prose commentary often coexist in a document,
    # and the old behavior silently dropped all of it whenever any table
    # was present.
    text_rows = []
    chart_data_list = []
    for p in doc.paragraphs:
        full_text = _full_paragraph_text(p).strip()
        if full_text:
            text_rows.append(full_text)
        for tb_para in _iter_textbox_paragraphs(p):
            tb_text = _full_paragraph_text(tb_para).strip()
            if tb_text:
                text_rows.append(tb_text)
        for chart_title, chart_cats, chart_series, _chart_type in _iter_docx_charts(p, doc):
            chart_data_list.append((chart_title, chart_cats, chart_series))
    if text_rows:
        ws = wb.create_sheet(title="Document Text")
        ws.column_dimensions["A"].width = 100
        for r, line in enumerate(text_rows, 1):
            ws.cell(row=r, column=1, value=line)

    for i, (chart_title, chart_cats, chart_series) in enumerate(chart_data_list, 1):
        sheet_title = (chart_title or f"Chart {i}")[:31]
        ws = wb.create_sheet(title=sheet_title)
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
        ws = wb.create_sheet(title="Footnotes")
        ws.column_dimensions["A"].width = 100
        for r, note in enumerate(footnotes, 1):
            ws.cell(row=r, column=1, value=f"{r}. {note}")

    endnotes = _get_docx_endnotes(doc)
    if endnotes:
        ws = wb.create_sheet(title="Endnotes")
        ws.column_dimensions["A"].width = 100
        for r, note in enumerate(endnotes, 1):
            ws.cell(row=r, column=1, value=f"{r}. {note}")

    comments = _get_docx_comments(doc)
    if comments:
        ws = wb.create_sheet(title="Comments")
        ws.column_dimensions["A"].width = 25
        ws.column_dimensions["B"].width = 90
        ws.cell(row=1, column=1, value="Author").font = XlsxFont(bold=True)
        ws.cell(row=1, column=2, value="Comment").font = XlsxFont(bold=True)
        for r, (author, comment_text) in enumerate(comments, 2):
            ws.cell(row=r, column=1, value=author)
            ws.cell(row=r, column=2, value=comment_text)

    header_footer = _get_docx_header_footer_text(doc)
    if header_footer:
        ws = wb.create_sheet(title="Header-Footer")
        ws.column_dimensions["A"].width = 100
        for r, line in enumerate(header_footer, 1):
            ws.cell(row=r, column=1, value=line)

    if not wb.sheetnames:
        ws = wb.create_sheet(title="Sheet1")
        ws["A1"] = "(No content found in document)"

    os.makedirs(out_dir, exist_ok=True)
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
    ("pptx", "pptx"): pptx_to_pptx,
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

        def set_slide_title(slide, text):
            tf = slide.shapes.title.text_frame
            tf.clear()
            _add_markdown_aware_pptx_text(tf.paragraphs[0], text[:120])

        slide = None
        tf = None
        count = 0
        for i, line in enumerate(body_lines):
            if slide is None or count >= MAX_BULLETS_PER_SLIDE:
                slide_title = title if slide is None else title + " (cont.)"
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

    title = str(payload.get("title") or "Academic Response").strip()[:200]
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.line_spacing = 2.0
    _add_markdown_aware_text(title_p, title, base_bold=True)

    def add_body_paragraph(text):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 2.0
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = DocxInches(0.5)
        _add_markdown_aware_text(p, str(text) if text else "")
        return p

    def add_section_heading(text, level=1):
        try:
            level = int(level)
        except (TypeError, ValueError):
            level = 1
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 2.0
        if level <= 1:
            # APA Level 1: centered, bold
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _add_markdown_aware_text(p, text, base_bold=True)
        elif level == 2:
            # APA Level 2: left-aligned, bold
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            _add_markdown_aware_text(p, text, base_bold=True)
        else:
            # APA Level 3: left-aligned, bold italic
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(_sanitize_xml_text(text))
            run.font.name = "Times New Roman"
            run.font.size = DocxPt(12)
            run.font.color.rgb = DocxRGBColor(0, 0, 0)
            run.bold = True
            run.italic = True
        return p

    sections = payload.get("sections")
    if sections and isinstance(sections, list) and any(isinstance(s, dict) for s in sections):
        for sec in sections:
            if not isinstance(sec, dict):
                continue
            heading_text = f"{sec.get('number', '')} {sec.get('heading', '')}".strip()
            if heading_text:
                add_section_heading(heading_text, level=sec.get("level", 1))
            add_body_paragraph(sec.get("text") or "")
    else:
        text = (payload.get("text") or "").strip()
        paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
        for para in paragraphs or [text]:
            add_body_paragraph(para)

    references = payload.get("references") or []
    if references and isinstance(references, list):
        references = [r for r in references if isinstance(r, dict)]
    else:
        references = []
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
            ref_text = str(ref.get("text") or "").strip()
            url = str(ref.get("url") or "").strip()
            if ref_text:
                _add_markdown_aware_text(p, ref_text + (" " if url else ""))
            if url:
                _set_run_font(p.add_run(_sanitize_xml_text(url)))

    os.makedirs(out_dir, exist_ok=True)
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
        return "\n".join((page.extract_text() or "") for page in reader.pages)
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


def _check_pdf_not_encrypted(reader, action="edit"):
    if reader.is_encrypted:
        raise PdfEditError(
            f"This PDF is password-protected and can't be {action}d until it's unlocked "
            f"— use the remove-password tool first, with the correct password."
        )


def pdf_merge(pdf_paths, out_dir, filename="merged.pdf"):
    """Merges multiple PDFs, in the given order, into one document."""
    if len(pdf_paths) < 2:
        raise PdfEditError("Merging needs at least two PDF files.")
    writer = pypdf.PdfWriter()
    for path in pdf_paths:
        reader = _safe_load(pypdf.PdfReader, path)
        _check_pdf_not_encrypted(reader, "merge")
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
    _check_pdf_not_encrypted(reader, "reorder")
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
    _check_pdf_not_encrypted(reader, "edit")
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
    _check_pdf_not_encrypted(reader, "edit")
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
    _check_pdf_not_encrypted(reader, "rotate")
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
    _check_pdf_not_encrypted(reader, "number")
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
    _check_pdf_not_encrypted(reader, "watermark")
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
    _check_pdf_not_encrypted(reader, "password-protect")
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
    _check_pdf_not_encrypted(reader, "compress")
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


def images_to_pdf(image_paths, out_dir, filename="images.pdf"):
    """Combines one or more images into a PDF, one image per page,
    each page sized to that image's own aspect ratio (letter-width,
    scaled) rather than forcing every image into one fixed page size
    and distorting or letterboxing it."""
    from PIL import Image
    if not image_paths:
        raise PdfEditError("No images were given to combine into a PDF.")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, filename)
    page_w = _RL_LETTER[0]
    c = _rl_canvas.Canvas(out_path)
    for img_path in image_paths:
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
        except Exception as e:
            raise PdfEditError(f"Couldn't add image '{os.path.basename(img_path)}': {e}")
    c.save()
    return out_path


def pdf_get_page_thumbnails(pdf_path, out_dir, max_pages=200, dpi=80):
    """Rasterizes each page to a small JPEG thumbnail, for a page-picker
    UI where a user selects/reorders pages visually rather than by
    typing page numbers. Capped at max_pages for a very long document,
    since rendering hundreds of thumbnails at once isn't a reasonable
    single request regardless of how fast any one page is."""
    reader = _safe_load(pypdf.PdfReader, pdf_path)
    _check_pdf_not_encrypted(reader, "preview")
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




