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
from docx.text.run import Run as DocxRun
from pptx import Presentation
from pptx.util import Inches as PptxInches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor as PptxRGBColor
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font as XlsxFont, PatternFill as XlsxPatternFill
from lxml import etree
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
    out_path = os.path.join(out_dir, f"{base}.{target_format}")
    if not os.path.exists(out_path):
        raise ConversionError(f"LibreOffice conversion failed: {result.stderr or result.stdout}")
    return out_path


# ---------------------------------------------------------------- DOCX -> PDF
def docx_to_pdf(src_path, out_dir, style=None):
    return _soffice_convert(src_path, "pdf", out_dir)


# --------------------------------------------------------------- DOCX -> PPTX
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


def _get_docx_footnotes(doc):
    return _get_docx_notes(doc, "/word/footnotes.xml", "w:footnote")


def _get_docx_endnotes(doc):
    return _get_docx_notes(doc, "/word/endnotes.xml", "w:endnote")



def _effective_run_format(run):
    """Returns (bold, italic, underline) for a run, falling back to its
    referenced character style when the run has no direct formatting of
    its own. Word's built-in 'Strong'/'Emphasis' styles — and most custom
    character styles — carry bold/italic in the STYLE definition rather
    than as direct run formatting, so run.bold alone returns None for
    these even though the text visibly renders bold, silently losing the
    formatting on conversion. Direct run formatting always wins when
    present; the style is only consulted for whichever of the three is
    still unset."""
    bold, italic, underline = run.bold, run.italic, run.underline
    if bold is None or italic is None or underline is None:
        try:
            style_font = run.style.font
            if bold is None:
                bold = style_font.bold
            if italic is None:
                italic = style_font.italic
            if underline is None:
                underline = style_font.underline
        except Exception:
            pass
    return bold, italic, underline


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
        return "".join(t.text or "" for t in r_el.findall(tag))

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
    """Returns (level, is_numbered) from the paragraph's own direct numPr
    formatting — the authoritative source when present — or None if the
    paragraph has no direct list formatting (in which case a caller should
    fall back to a style-name-based guess, since a style like 'List Bullet'
    carries its numbering via the *style* definition rather than direct
    per-paragraph numPr)."""
    pPr = paragraph._p.find(qn("w:pPr"))
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


def _iter_textbox_paragraphs(paragraph):
    """Yields each paragraph nested inside a floating text box embedded in
    this paragraph's runs. A text box's own paragraphs live inside a nested
    <w:txbxContent> deep within the run's drawing XML — a completely
    separate tree from the main document body — so normal paragraph
    iteration (doc.paragraphs, iter_block_items) never sees them at all,
    silently dropping the text box's entire content."""
    for run in paragraph.runs:
        for txbx_content in run._element.findall(".//" + qn("w:txbxContent")):
            for p_el in txbx_content.findall(qn("w:p")):
                yield DocxParagraph(p_el, paragraph)


def _iter_inline_images(paragraph, doc):
    """Yields raw image bytes for each inline picture embedded in this
    paragraph, in document order — including images inside a hyperlink.
    python-docx's own paragraph.runs deliberately excludes hyperlink-wrapped
    runs, so walking iter_inner_content() (which covers both) is required
    here, not just runs, or a clickable image would be silently skipped."""
    for item in paragraph.iter_inner_content():
        runs = item.runs if type(item).__name__ == "Hyperlink" else [item]
        for run in runs:
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
    numbering_formats = _get_docx_numbering_formats(doc)
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

    def add_bullet(para, level=0, numbered=False):
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
        pptx_align = _docx_align_to_pptx(para.alignment)
        if pptx_align is not None:
            p.alignment = pptx_align
        if numbered:
            pPr = p._p.get_or_add_pPr()
            pPr.append(pPr.makeelement(qn("a:buAutoNum"), {"type": "arabicPeriod"}))

        # iter_inner_content() (like .text and .runs) silently drops any
        # text wrapped in a tracked-change insertion or deletion —
        # _iter_all_runs() is the tracked-changes-aware replacement, and
        # since every item it yields is already a real Run, the hyperlink
        # special-casing this used to need is gone too.
        captured = []  # (pptx_run, bold, italic, underline, is_link, hex_color, size_pt, is_deleted)
        for src_run, text, is_link, address, is_deleted in _iter_all_runs(para):
            r = p.add_run()
            r.text = text
            style_bold, style_italic, style_underline = _effective_run_format(src_run)
            bold = bool(style_bold)
            italic = bool(style_italic)
            underline = True if is_link else bool(style_underline)
            hex_color = None if is_link else _safe_hex_color(src_run.font.color)
            size_pt = None if src_run.font.size is None else src_run.font.size.pt
            if is_link and address:
                try:
                    r.hyperlink.address = address
                except Exception:
                    pass
            if is_deleted:
                r._r.get_or_add_rPr().set("strike", "sngStrike")
            captured.append((r, bold, italic, underline, is_link, hex_color, size_pt, is_deleted))
        if not captured:
            p.text = _full_paragraph_text(para).strip()
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

    def add_subheading(text):
        """Heading 2/3 in the source document reads as a subsection within
        the current topic, not a brand new topic — giving it a whole new
        slide (the old behavior for every heading level) fragmented what's
        usually meant to be one cohesive slide into several thin ones. This
        keeps it on the current slide as a bolded standalone line instead."""
        if state["slide"] is None:
            new_slide(text)
            return
        if state["bullet_count"] >= MAX_BULLETS_PER_SLIDE:
            current_title = state["slide"].shapes.title.text or "Overview"
            new_slide(current_title + " (cont.)")
        body_tf = state["body_tf"]
        if body_tf.paragraphs[0].text == "" and len(body_tf.paragraphs) == 1 and state["bullet_count"] == 0:
            p = body_tf.paragraphs[0]
        else:
            p = body_tf.add_paragraph()
        p.level = 0
        p.text = text
        _style_pptx_body_paragraph(p, template_style)
        for r in p.runs:
            r.font.bold = True
        state["bullet_count"] += 1

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
        for tb_para in _iter_textbox_paragraphs(para):
            tb_text = _full_paragraph_text(tb_para).strip()
            if tb_text:
                add_bullet(tb_para, level=0)
        text = _full_paragraph_text(para).strip()
        if not text:
            continue
        para_style_name = (para.style.name or "").lower()
        heading_match = re.match(r"heading (\d+)", para_style_name)
        if para_style_name == "title" or (heading_match and int(heading_match.group(1)) <= 1):
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

    if state["slide"] is None:
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

    out_path = os.path.join(out_dir, "converted.pptx")
    prs.save(out_path)
    return out_path


# --------------------------------------------------------------- PPTX -> DOCX
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
    prs = Presentation(src_path)
    doc = Document()
    doc.add_heading("Slide Handout", 0)
    BULLET_STYLES = ["List Bullet", "List Bullet 2", "List Bullet 3"]
    NUMBER_STYLES = ["List Number", "List Number 2", "List Number 3"]

    for i, slide in enumerate(prs.slides, 1):
        title = None
        text_shapes = []
        table_shapes = []
        chart_shapes = []
        image_shapes = []
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
                title = shape.text_frame.text.strip()
            else:
                text_shapes.append(shape)

        doc.add_heading(title or f"Slide {i}", level=1)

        for shape in image_shapes:
            try:
                doc.add_picture(io.BytesIO(shape.image.blob), width=DocxInches(4))
            except Exception:
                continue  # a malformed/unsupported embedded image shouldn't sink the whole conversion

        for shape in text_shapes:
            for para in shape.text_frame.paragraphs:
                line = para.text.strip()
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
                    address = None
                    try:
                        address = run.hyperlink.address
                    except Exception:
                        pass
                    if address:
                        _add_docx_hyperlink(
                            p, address, run.text,
                            bold=bool(run.font.bold), italic=bool(run.font.italic),
                        )
                    else:
                        r = p.add_run(run.text)
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
            rows = [[cell.text.strip() for cell in row] for row in src_rows]
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
            for r_idx, row in enumerate(rows):
                for c_idx in range(n_cols):
                    src_cell = src_rows[r_idx][c_idx] if c_idx < len(src_rows[r_idx]) else None
                    if src_cell is not None and src_cell.is_spanned and not src_cell.is_merge_origin:
                        continue  # covered by a merge origin elsewhere — filled in via the merge below
                    cell = word_table.cell(r_idx, c_idx)
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
                categories = [str(cat) for cat in plot.categories]
                series_list = list(plot.series)
            except Exception:
                continue
            if not series_list:
                continue
            caption = doc.add_paragraph()
            title_text = None
            try:
                if chart.has_title:
                    title_text = chart.chart_title.text_frame.text.strip()
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
                chart_table.cell(0, s_idx).text = series.name or f"Series {s_idx}"
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
    # A formula cell in a workbook that's never been opened in a real
    # spreadsheet app (generated by a script, exported from a database) has
    # no cached result — data_only=True silently returns None for it, which
    # renders as a misleadingly blank cell with no sign a formula was ever
    # there. Loading a second copy without data_only lets an uncalculated
    # formula fall back to showing its actual formula text instead.
    wb_formulas = openpyxl.load_workbook(src_path, data_only=False)

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
        sheet_rows.append((sheet_name, rows, formula_rows, merged_ranges))
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
    for sheet_name, rows, formula_rows, merged_ranges in sheet_rows:
        doc.add_heading(sheet_name, level=1)
        if not rows:
            doc.add_paragraph("(Empty sheet)")
            continue
        n_cols = max(len(r) for r in rows)
        table = doc.add_table(rows=len(rows), cols=n_cols)
        table.style = "Light Grid Accent 1"
        cell_comments = []
        for r_idx, row in enumerate(rows):
            for c_idx in range(n_cols):
                src_cell = row[c_idx] if c_idx < len(row) else None
                val = src_cell.value if src_cell is not None else None
                if val is None and c_idx < len(formula_rows[r_idx]) and formula_rows[r_idx][c_idx].data_type == "f":
                    val = formula_rows[r_idx][c_idx].value  # uncalculated formula — show the formula itself, not blank
                cell = table.cell(r_idx, c_idx)
                cell.text = _format_cell_value(val)
                if r_idx == 0:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            run.bold = True
                if src_cell is not None:
                    fill_hex = _xlsx_cell_fill_hex(src_cell)
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
    for p in doc.paragraphs:
        full_text = _full_paragraph_text(p).strip()
        if full_text:
            text_rows.append(full_text)
        for tb_para in _iter_textbox_paragraphs(p):
            tb_text = _full_paragraph_text(tb_para).strip()
            if tb_text:
                text_rows.append(tb_text)
    if text_rows:
        ws = wb.create_sheet(title="Document Text")
        ws.column_dimensions["A"].width = 100
        for r, line in enumerate(text_rows, 1):
            ws.cell(row=r, column=1, value=line)

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
        parts = []
        for p in doc.paragraphs:
            full_text = _full_paragraph_text(p).strip()
            if full_text:
                parts.append(full_text)
            for tb_para in _iter_textbox_paragraphs(p):
                tb_text = _full_paragraph_text(tb_para).strip()
                if tb_text:
                    parts.append(tb_text)
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
        reader = pypdf.PdfReader(src_path)
        return "\n".join((page.extract_text() or "") for page in reader.pages)
    if ext == "pptx":
        prs = Presentation(src_path)
        parts = []
        for slide in prs.slides:
            for shape in _iter_flat_shapes(slide.shapes):
                if shape.has_text_frame and shape.text_frame.text.strip():
                    parts.append(shape.text_frame.text)
        return "\n".join(parts)
    if ext == "xlsx":
        wb = openpyxl.load_workbook(src_path, data_only=True)
        wb_formulas = openpyxl.load_workbook(src_path, data_only=False)
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
                    if val is not None:
                        cells.append(str(val))
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
