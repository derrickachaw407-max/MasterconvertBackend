"""6.4: PDF to Word, made again from the ground up.

The owner, 19 Nov 2026: "the word pdf feature, it's validity was very low please raise it to the best it can be" and
"Please never add anything I mean anything at all to converted files."

Until 6.4 a PDF's words were read as plain text and poured into a new document: the fonts, sizes, bold and italic, the
colours, centred and justified lines, indents, the places of pictures, the widths and looks of tables, the header and
footer were all lost, and a title, a note and a table style the PDF never had were put in.

Now the page is read as it is drawn (every letter with its font, size, colour and place; every ruled line, filled box
and picture) and the Word file is built to look the same and still be a document a person can change:
  - words keep their font (by the name Word knows it by), size, bold, italic, colour, underline, raised and lowered letters;
  - lines become paragraphs again, with their alignment, indents, the room between lines and between paragraphs;
  - a gap inside a line is a tab at that place (marks at the right of an exam paper, a header's two ends);
  - ruled tables are Word tables with the same columns, borders, shading and merged cells; rows of words standing in
    columns with no rules are a table without borders;
  - pictures, charts and drawings are cut from the page as it is drawn (what is seen is what is kept: a cropped
    picture stays cropped) and stand where they stood;
  - columns of text are read in their order and are columns in Word;
  - what repeats at the top and foot of the pages is Word's header and footer, the page number a real page number;
  - the paper's size, its direction and its margins are the PDF's;
  - a scanned page is read with OCR and set out the same way; a page that cannot be read is kept as its picture.
NOTHING is added: no title, no note, no label, no style the PDF did not have.

It uses what the server already has: pdfplumber (pdfminer) to read, Poppler to draw a page for its pictures,
Tesseract for scanned pages, python-docx to write.
"""
import io
import math
import os
import re
import shutil
import statistics
import subprocess
import tempfile
import time
from collections import Counter, defaultdict

import pdfplumber
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.opc.constants import RELATIONSHIP_TYPE as _RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Pt
from PIL import Image

try:                                    # the server's own error; the module also runs alone (tests, tools)
    if __name__ == "__main__":          # run as a program (the server does: see converters.pdf_to_docx) it stays light
        raise ImportError
    from converters import ConversionError
except Exception:                       # pragma: no cover
    class ConversionError(Exception):
        pass

MAX_PAGES = 400                # a longer PDF is asked to be split first (the server gives a request four minutes)
MAX_OCR_PAGES = 20             # scanned pages read with OCR; the ones after are kept as pictures of the page
OCR_BUDGET = 150.0             # seconds: no new page is given to OCR after this (the request must end in time)
TIME_BUDGET = 170.0            # seconds: the whole conversion is planned to end by then (the server stops it at 200)
RENDER_END = 150.0             # seconds: the pictures are drawn by then, at a lower resolution if need be (_Pace);
RENDER_SLACK = 25.0            # ...or a little after; where that cannot be, the PDF is asked to be split, at once
READ_BUDGET = 120.0            # seconds for reading the pages: a PDF that would need longer is asked to be split, and soon
LOCKED = "This PDF is protected with a password. Take the password off first (PDF tools, Password), then convert it."
NO_COPY = ("The person who made this PDF has locked it against copying its contents, so it cannot be converted. "
           "Ask them for a copy without that lock.")
TOO_LONG = ("This PDF is too long to turn into Word in one go. Split it into smaller parts first (PDF tools, Split a PDF), "
            "then convert each part.")
BACK_DPI = 150                 # a designed page's drawing, behind its words
PIC_DPI = 200                  # pictures are cut from the page drawn at this resolution (or the picture's own, if lower)


# ----------------------------------------------------------------------------------------------------------- fonts
# The name a PDF gives a font ("ABCDEF+TimesNewRomanPS-BoldMT", "Carlito-Italic", "CMBX10") -> the family Word knows,
# and whether it is bold or italic. A font the person's computer does not have is named by its metric twin: Carlito is
# Calibri, Liberation Serif is Times New Roman (so the lines break where they broke).
_FAMILIES = [
    (r"^(timesnewroman|times|nimbusrom|nimbusroman|liberationserif|tinos|texgyretermes|freeserif|thorndale|ptm|utm|stixgeneral|txr)", "Times New Roman"),
    (r"^(arialnarrow|liberationsansnarrow|helveticanarrow|nimbussansnarrow|nimbussannar)", "Arial Narrow"),
    (r"^arialblack", "Arial Black"),
    (r"^(arialunicode)", "Arial"),
    (r"^(arial|helveticaneue|helvetica|nimbussan|liberationsans|arimo|texgyreheros|freesans|albany|phv|uhv)", "Arial"),
    (r"^(couriernew|courier|nimbusmon|liberationmono|cousine|texgyrecursor|freemono|cumberland|pcr|ucr)", "Courier New"),
    (r"^calibrilight", "Calibri Light"),
    (r"^(calibri|carlito)", "Calibri"),
    (r"^cambriamath", "Cambria"),
    (r"^(cambria|caladea)", "Cambria"),
    (r"^(georgia|gelasio)", "Georgia"),
    (r"^(dejavusansmono|bitstreamverasansmono|lucidaconsole|lucidasanstypewriter)", "Lucida Console"),
    (r"^(verdana|dejavusans|bitstreamverasans)", "Verdana"),
    (r"^(dejavuserif|bitstreamveraserif)", "Georgia"),
    (r"^tahoma", "Tahoma"),
    (r"^trebuchet", "Trebuchet MS"),
    (r"^(ebgaramond|garamond|agaramond|adobegaramond)", "Garamond"),
    (r"^(bookantiqua)", "Book Antiqua"),
    (r"^(palatinolinotype|palatino|p052|texgyrepagella|urwpalladio|pplr?|uplr?)", "Palatino Linotype"),
    (r"^(centurygothic|urwgothic|texgyreadventor|avantgarde|pag)", "Century Gothic"),
    (r"^(centuryschoolbook|c059|texgyreschola|newcenturyschlbk|pnc)", "Century Schoolbook"),
    (r"^(bookmanoldstyle|bookman|urwbookman|texgyrebonum|pbk)", "Bookman Old Style"),
    (r"^segoeuisymbol", "Segoe UI Symbol"),
    (r"^segoeui", "Segoe UI"),
    (r"^(comicsans|comicneue)", "Comic Sans MS"),
    (r"^(consolas|inconsolata)", "Consolas"),
    (r"^candara", "Candara"), (r"^corbel", "Corbel"), (r"^constantia", "Constantia"),
    (r"^aptosdisplay", "Aptos Display"), (r"^aptosnarrow", "Aptos Narrow"), (r"^aptosmono", "Aptos Mono"),
    (r"^aptosserif", "Aptos Serif"), (r"^aptos", "Aptos"),
    (r"^gillsans", "Gill Sans MT"), (r"^franklingothic", "Franklin Gothic Book"), (r"^rockwell", "Rockwell"),
    (r"^impact", "Impact"), (r"^lucidasans", "Lucida Sans"), (r"^baskerville", "Baskerville Old Face"),
    (r"^(symbolmt|symbol|standardsymbolsps|standardsymbolsl)", "Symbol"),
    (r"^wingdings3", "Wingdings 3"), (r"^wingdings2", "Wingdings 2"), (r"^wingdings", "Wingdings"), (r"^webdings", "Webdings"),
    (r"^(zapfdingbats|dingbats|d050000l)", "Wingdings"),
    (r"^opensymbol", "Segoe UI Symbol"),
    (r"^(mssansserif|microsoftsansserif)", "Microsoft Sans Serif"),
    (r"^(simsun|nsimsun)", "SimSun"), (r"^(microsoftyahei|msyh)", "Microsoft YaHei"), (r"^(msgothic|msmincho)", "MS Gothic"),
    (r"^(malgungothic)", "Malgun Gothic"),
    (r"^(notosanscjk|sourcehansans)", "Microsoft YaHei"), (r"^(notoserifcjk|sourcehanserif)", "SimSun"),
    (r"^notoserif", "Cambria"), (r"^notosansmono", "Consolas"), (r"^notosans", "Segoe UI"),
    (r"^opensans", "Open Sans"), (r"^roboto", "Roboto"), (r"^lato", "Lato"), (r"^montserrat", "Montserrat"), (r"^poppins", "Poppins"),
    (r"^(sourcesans)", "Source Sans Pro"), (r"^inter", "Inter"),
    # TeX: Computer Modern and its kin have no twin on a phone or a PC; Cambria is the nearest in width and colour
    (r"^(cmss|lmsans|sfss|ecss|cmbr)", "Arial"),
    (r"^(cmtt|cmtex|cmsltt|lmmono|sftt|ectt|cmitt)", "Courier New"),
    (r"^(cmmi|cmsy|cmex|cmmib|cmbsy|msam|msbm|lmmath|eufm|rsfs|esint|stmary|lasy|wasy)", "Cambria"),
    (r"^(cmr|cmbx|cmti|cmsl|cmcsc|cmb|cmu|cmdunh|cmff|cmfi|cmfib|cmvtt|lmroman|lmr|sfrm|sfbx|sfti|sfsl|sfcc|ecrm|ecbx|ecti|ecsl|eccc|sfbi|ecbi)", "Cambria"),
]
_FAMILIES = [(re.compile(p), f) for p, f in _FAMILIES]
_STYLE_WORDS = re.compile(r"(bolditalic|boldoblique|boldit|bolditalicmt|semibolditalic|italic|oblique|slanted|semibold|demibold|extrabold|"
                          r"ultrabold|bold|black|heavy|light|medium|regular|roman|book|normal|condensed|narrow|psmt|ps|mt)$")
_BOLD_TEX = re.compile(r"^(cmbx|cmb\d|cmbsy|cmmib|cmssbx|sfbx|sfbi|ecbx|ecbi|cmbxti|cmbxsl|lmroman\d*bold|lmsans\d*bold)")
_ITAL_TEX = re.compile(r"^(cmti|cmsl|cmmi|cmbxti|cmbxsl|cmssi|cmitt|cmsltt|sfti|sfsl|sfbi|ecti|ecsl|ecbi|cmu.*(it|sl)$)")
# how tall one line of each font is, as a share of its size, when Word sets it "single" (the font's own figures)
_LINE = {"Calibri": 1.2207, "Calibri Light": 1.2207, "Cambria": 1.1724, "Times New Roman": 1.1499, "Arial": 1.1499, "Arial Narrow": 1.147,
         "Courier New": 1.1328, "Georgia": 1.1362, "Verdana": 1.2153, "Tahoma": 1.207, "Trebuchet MS": 1.1611, "Segoe UI": 1.3301,
         "Comic Sans MS": 1.3936, "Consolas": 1.1709, "Candara": 1.2207, "Corbel": 1.2207, "Constantia": 1.2207, "Symbol": 1.2251,
         "Garamond": 1.125, "Book Antiqua": 1.2456, "Palatino Linotype": 1.3486, "Century Gothic": 1.2261, "Century Schoolbook": 1.2002,
         "Bookman Old Style": 1.188, "Lucida Console": 1.0, "Cambria Math": 1.1724}
_FONT_CACHE = {}
# PDF fonts that are the Word font itself or its twin in every letter's width: with these, Word breaks the lines and
# fills the pages exactly as the PDF does (with any other font it only comes near)
_EXACT = re.compile(r"^(timesnewroman|times|nimbusrom|nimbusroman|liberationserif|tinos|texgyretermes|thorndale|arialnarrow|liberationsansnarrow|"
                    r"arialblack|arial|helvetica|nimbussan|liberationsans|arimo|texgyreheros|albany|couriernew|courier|nimbusmon|liberationmono|"
                    r"cousine|texgyrecursor|cumberland|calibri|carlito|cambria|caladea|georgia|gelasio|verdana|tahoma|trebuchet|garamond|"
                    r"bookantiqua|palatinolinotype|palatino|p052|texgyrepagella|centurygothic|urwgothic|texgyreadventor|centuryschoolbook|c059|"
                    r"texgyreschola|bookmanoldstyle|urwbookman|texgyrebonum|segoeui|comicsans|consolas|candara|corbel|constantia|aptos|gillsans|"
                    r"franklingothic|rockwell|impact|lucidasans|lucidaconsole|baskerville|symbol|wingdings|webdings|opensymbol|freeserif|freesans|freemono)")


# fonts every Word has (Windows or Mac): the reader's computer needs no word about them
_WORD_FONTS = set(("Times New Roman", "Arial", "Arial Narrow", "Arial Black", "Courier New", "Calibri", "Calibri Light", "Cambria", "Georgia", "Verdana",
                   "Tahoma", "Trebuchet MS", "Garamond", "Book Antiqua", "Palatino Linotype", "Century Gothic", "Century Schoolbook",
                   "Bookman Old Style", "Segoe UI", "Segoe UI Symbol", "Comic Sans MS", "Consolas", "Candara", "Corbel", "Constantia", "Symbol",
                   "Wingdings", "Wingdings 2", "Wingdings 3", "Webdings", "Lucida Console", "Microsoft Sans Serif", "Impact"))
_SERIF_NAMES = re.compile(r"serif|roman|times|century|garamond|minion|melior|palatin|caslon|baskerv|bembo|sabon|goudy|utopia|charter|bodoni|didot|"
                          r"bookman|schoolbook|georgia|cambria|constantia|antiqua|jenson|plantin|perpetua|rockwell|clarendon|courier|libertine|"
                          r"crimson|merriweather|playfair|lora|pt ?serif|stix|computer ?modern|latin ?modern|cochin|hoefler|elephant|mincho|song|ming")
_SANS_NAMES = re.compile(r"sans|gothic|grotesk|grotesque|helvet|arial|univers|frutiger|futura|avenir|myriad|optima|din\b|segoe|calibri|verdana|tahoma|"
                         r"trebuchet|lato|roboto|inter\b|montserrat|poppins|nunito|ubuntu|raleway|oswald|aptos|candara|corbel|gill|franklin|"
                         r"lucida|akzidenz|proxima|gotham|interstate|eurostile|kozuka ?go|hei\b|yahei|meiryo|malgun")
_MONO_NAMES = re.compile(r"mono|courier|typewriter|consol|\bcode\b|menlo|monaco|terminal|fixed|andale")


# <widths64>  (made by tools/widths64.py: do not change by hand)
_W_CHARS = ' !"#$%&\'()*+,-./0123456789:;<=>?@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`abcdefghijklmnopqrstuvwxyz{|}~\xa1\xa2\xa3\xa4\xa5\xa6\xa7\xa8\xa9\xaa\xab\xac\xad\xae\xaf\xb0\xb1\xb2\xb3\xb4\xb5\xb6\xb7\xb8\xb9\xba\xbb\xbc\xbd\xbe\xbf\xc0\xc1\xc2\xc3\xc4\xc5\xc6\xc7\xc8\xc9\xca\xcb\xcc\xcd\xce\xcf\xd0\xd1\xd2\xd3\xd4\xd5\xd6\xd7\xd8\xd9\xda\xdb\xdc\xdd\xde\xdf\xe0\xe1\xe2\xe3\xe4\xe5\xe6\xe7\xe8\xe9\xea\xeb\xec\xed\xee\xef\xf0\xf1\xf2\xf3\xf4\xf5\xf6\xf7\xf8\xf9\xfa\xfb\xfc\xfd\xfe\xff\u2013\u2014\u2018\u2019\u201c\u201d\u2022\u2026'
_W_TABLE = {
    ("Times New Roman", 0): "250 333 408 500 500 833 778 180 333 333 500 564 250 333 250 278 500 500 500 500 500 500 500 500 500 500 278 278 564 564 564 444 921 722 667 667 722 611 556 722 722 333 389 722 611 889 722 722 556 722 667 556 611 722 722 944 722 722 611 333 278 333 469 500 333 444 500 444 500 444 333 500 500 278 278 500 278 778 500 500 500 500 333 389 278 500 500 722 500 500 444 480 200 480 541 333 500 500 500 500 200 500 333 760 276 500 564 333 760 500 400 549 300 300 333 576 453 333 333 300 310 500 750 750 750 444 722 722 722 722 722 722 889 667 611 611 611 611 333 333 333 333 722 722 722 722 722 722 722 564 722 722 722 722 722 722 556 500 444 444 444 444 444 444 667 444 444 444 444 444 278 278 278 278 500 500 500 500 500 500 500 549 500 500 500 500 500 500 500 500 500 1000 333 333 444 444 350 1000",
    ("Times New Roman", 1): "250 333 555 500 500 1000 833 278 333 333 500 570 250 333 250 278 500 500 500 500 500 500 500 500 500 500 333 333 570 570 570 500 930 722 667 722 722 667 611 778 778 389 500 778 667 944 722 778 611 778 722 556 667 722 722 1000 722 722 667 333 278 333 581 500 333 500 556 444 556 444 333 500 556 278 333 556 278 833 556 500 556 556 444 389 333 556 500 722 500 500 444 394 220 394 520 333 500 500 500 500 220 500 333 747 300 500 570 333 747 500 400 549 300 300 333 576 540 333 333 300 330 500 750 750 750 500 722 722 722 722 722 722 1000 722 667 667 667 667 389 389 389 389 722 722 778 778 778 778 778 570 778 722 722 722 722 722 611 556 500 500 500 500 500 500 722 444 444 444 444 444 278 278 278 278 500 556 500 500 500 500 500 549 500 556 556 556 556 500 556 500 500 1000 333 333 500 500 350 1000",
    ("Times New Roman", 2): "250 333 420 500 500 833 778 214 333 333 500 675 250 333 250 278 500 500 500 500 500 500 500 500 500 500 333 333 675 675 675 500 920 611 611 667 722 611 611 722 722 333 444 667 556 833 667 722 611 722 611 500 556 722 611 833 611 556 556 389 278 389 422 500 333 500 500 444 500 444 278 500 500 278 278 444 278 722 500 500 500 500 389 389 278 500 444 667 444 444 389 400 275 400 541 389 500 500 500 500 275 500 333 760 276 500 675 333 760 500 400 549 300 300 333 576 523 250 333 300 310 500 750 750 750 500 611 611 611 611 611 611 889 667 611 611 611 611 333 333 333 333 722 667 722 722 722 722 722 675 722 722 722 722 722 556 611 500 500 500 500 500 500 500 667 444 444 444 444 444 278 278 278 278 500 500 500 500 500 500 500 549 500 500 500 500 500 444 500 444 500 889 333 333 556 556 350 889",
    ("Times New Roman", 3): "250 389 555 500 500 833 778 278 333 333 500 570 250 333 250 278 500 500 500 500 500 500 500 500 500 500 333 333 570 570 570 500 832 667 667 667 722 667 667 722 778 389 500 667 611 889 722 722 611 722 667 556 611 722 667 889 667 611 611 333 278 333 570 500 333 500 500 444 500 444 333 500 556 278 278 500 278 778 556 500 500 500 389 389 278 556 444 667 500 444 389 348 220 348 570 389 500 500 500 500 220 500 333 747 266 500 606 333 747 500 400 549 300 300 333 576 500 250 333 300 300 500 750 750 750 500 667 667 667 667 667 667 944 667 667 667 667 667 389 389 389 389 722 722 722 722 722 722 722 570 722 722 722 722 722 611 611 500 500 500 500 500 500 500 722 444 444 444 444 444 278 278 278 278 500 556 500 500 500 500 500 549 500 556 556 556 556 444 500 444 500 1000 333 333 500 500 350 1000",
    ("Arial", 0): "278 278 355 556 556 889 667 191 333 333 389 584 278 333 278 278 556 556 556 556 556 556 556 556 556 556 278 278 584 584 584 556 1015 667 667 722 722 667 611 778 722 278 500 667 556 833 722 778 667 778 722 667 611 722 667 944 667 667 611 278 278 278 469 556 333 556 556 500 556 556 278 556 556 222 222 500 222 833 556 556 556 556 333 500 278 556 500 722 500 500 500 334 260 334 584 333 556 556 556 556 260 556 333 737 370 556 584 333 737 552 400 549 333 333 333 576 537 333 333 333 365 556 834 834 834 611 667 667 667 667 667 667 1000 722 667 667 667 667 278 278 278 278 722 722 778 778 778 778 778 584 778 722 722 722 722 667 667 611 556 556 556 556 556 556 889 500 556 556 556 556 278 278 278 278 556 556 556 556 556 556 556 549 611 556 556 556 556 500 556 500 556 1000 222 222 333 333 350 1000",
    ("Arial", 1): "278 333 474 556 556 889 722 238 333 333 389 584 278 333 278 278 556 556 556 556 556 556 556 556 556 556 333 333 584 584 584 611 975 722 722 722 722 667 611 778 722 278 556 722 611 833 722 778 667 778 722 667 611 722 667 944 667 667 611 333 278 333 584 556 333 556 611 556 611 556 333 611 611 278 278 556 278 889 611 611 611 611 389 556 333 611 556 778 556 556 500 389 280 389 584 333 556 556 556 556 280 556 333 737 370 556 584 333 737 552 400 549 333 333 333 576 556 333 333 333 365 556 834 834 834 611 722 722 722 722 722 722 1000 722 667 667 667 667 278 278 278 278 722 722 778 778 778 778 778 584 778 722 722 722 722 667 667 611 556 556 556 556 556 556 889 556 556 556 556 556 278 278 278 278 611 611 611 611 611 611 611 549 611 611 611 611 611 556 611 556 556 1000 278 278 500 500 350 1000",
    ("Arial", 2): "278 278 355 556 556 889 667 191 333 333 389 584 278 333 278 278 556 556 556 556 556 556 556 556 556 556 278 278 584 584 584 556 1015 667 667 722 722 667 611 778 722 278 500 667 556 833 722 778 667 778 722 667 611 722 667 944 667 667 611 278 278 278 469 556 333 556 556 500 556 556 278 556 556 222 222 500 222 833 556 556 556 556 333 500 278 556 500 722 500 500 500 334 260 334 584 333 556 556 556 556 260 556 333 737 370 556 584 333 737 552 400 549 333 333 333 576 537 333 333 333 365 556 834 834 834 611 667 667 667 667 667 667 1000 722 667 667 667 667 278 278 278 278 722 722 778 778 778 778 778 584 778 722 722 722 722 667 667 611 556 556 556 556 556 556 889 500 556 556 556 556 278 278 278 278 556 556 556 556 556 556 556 549 611 556 556 556 556 500 556 500 556 1000 222 222 333 333 350 1000",
    ("Arial", 3): "278 333 474 556 556 889 722 238 333 333 389 584 278 333 278 278 556 556 556 556 556 556 556 556 556 556 333 333 584 584 584 611 975 722 722 722 722 667 611 778 722 278 556 722 611 833 722 778 667 778 722 667 611 722 667 944 667 667 611 333 278 333 584 556 333 556 611 556 611 556 333 611 611 278 278 556 278 889 611 611 611 611 389 556 333 611 556 778 556 556 500 389 280 389 584 333 556 556 556 556 280 556 333 737 370 556 584 333 737 552 400 549 333 333 333 576 556 333 333 333 365 556 834 834 834 611 722 722 722 722 722 722 1000 722 667 667 667 667 278 278 278 278 722 722 778 778 778 778 778 584 778 722 722 722 722 667 667 611 556 556 556 556 556 556 889 556 556 556 556 556 278 278 278 278 611 611 611 611 611 611 611 549 611 611 611 611 611 556 611 556 556 1000 278 278 500 500 350 1000",
    ("Courier New", 0): "600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600",
    ("Courier New", 1): "600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600",
    ("Courier New", 2): "600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600",
    ("Courier New", 3): "600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600 600",
    ("Calibri", 0): "226 326 401 498 507 715 682 221 303 303 498 498 250 306 252 386 507 507 507 507 507 507 507 507 507 507 268 268 498 498 498 463 894 579 544 533 615 488 459 631 623 252 319 520 420 855 646 662 517 673 543 459 487 642 567 890 519 487 468 307 386 307 498 498 291 479 525 423 525 498 305 471 525 229 239 455 229 799 525 527 525 525 349 391 335 525 452 715 433 453 395 314 460 314 498 326 498 507 498 507 498 498 393 834 402 512 498 0 507 394 339 498 336 334 292 550 586 252 307 246 422 512 636 671 675 463 579 579 579 579 579 579 763 533 488 488 488 488 252 252 252 252 625 646 662 662 662 662 662 498 664 642 642 642 642 487 517 527 479 479 479 479 479 479 773 423 498 498 498 498 229 229 229 229 525 525 527 527 527 527 527 498 529 525 525 525 525 453 525 453 498 905 250 250 418 418 498 690",
    ("Calibri", 1): "226 326 438 498 507 729 705 233 312 312 498 498 258 306 267 430 507 507 507 507 507 507 507 507 507 507 276 276 498 498 498 463 898 606 561 529 630 488 459 637 631 267 331 547 423 874 659 676 532 686 563 473 495 653 591 906 551 520 478 325 430 325 498 498 300 494 537 418 537 503 316 474 537 246 255 480 246 813 537 538 537 537 355 399 347 537 473 745 459 474 397 344 475 344 498 326 498 507 498 507 498 498 415 834 416 539 498 0 507 390 342 498 338 336 301 563 598 268 303 252 435 539 658 691 702 463 606 606 606 606 606 606 775 529 488 488 488 488 267 267 267 267 639 659 676 676 676 676 676 498 681 653 653 653 653 520 532 555 494 494 494 494 494 494 775 418 503 503 503 503 246 246 246 246 537 537 538 538 538 538 538 498 544 537 537 537 537 474 537 474 498 905 258 258 435 435 498 711",
    ("Calibri", 2): "226 326 401 498 507 715 682 221 303 303 498 498 250 306 252 388 507 507 507 507 507 507 507 507 507 507 268 268 498 498 498 463 894 579 544 522 615 488 459 631 623 252 319 520 420 855 645 654 517 664 543 452 487 642 567 890 519 487 468 307 384 307 498 498 291 514 514 416 514 478 305 514 514 229 239 455 229 791 514 513 514 514 343 389 335 514 446 715 433 447 395 314 460 314 498 326 498 507 498 507 498 498 393 834 431 512 498 0 507 394 339 498 336 334 292 538 586 252 307 246 422 512 636 671 675 463 579 579 579 579 579 579 763 522 488 488 488 488 252 252 252 252 625 645 654 654 654 654 654 498 658 642 642 642 642 487 517 527 514 514 514 514 514 514 754 416 478 478 478 478 229 229 229 229 525 514 513 513 513 513 513 498 529 514 514 514 514 447 514 447 498 905 250 250 418 418 498 690",
    ("Calibri", 3): "226 326 438 498 507 729 705 233 312 312 498 498 258 306 267 434 507 507 507 507 507 507 507 507 507 507 276 276 498 498 498 463 898 606 561 519 630 488 459 637 631 267 331 547 423 874 656 668 532 677 563 465 495 653 591 907 551 520 478 325 425 325 498 498 300 528 528 412 528 491 316 528 527 246 255 480 246 804 527 527 528 528 352 394 347 527 469 745 459 470 397 344 475 344 498 326 498 507 498 507 498 498 415 834 441 539 498 0 507 390 342 498 338 336 301 554 598 268 303 252 435 539 658 691 702 463 606 606 606 606 606 606 775 519 488 488 488 488 267 267 267 267 639 656 668 668 668 668 668 498 677 653 653 653 653 520 532 555 528 528 528 528 528 528 764 412 491 491 491 491 246 246 246 246 537 527 527 527 527 527 527 498 544 527 527 527 527 470 528 470 498 905 258 258 435 435 498 711",
    ("Cambria", 0): "220 286 393 619 506 890 687 237 382 382 427 554 205 329 206 490 517 362 472 461 492 475 470 470 485 478 264 264 554 554 554 422 885 599 579 519 625 540 509 539 696 343 304 630 501 888 693 599 533 599 613 453 557 687 598 889 571 557 507 350 490 350 554 371 151 470 520 421 520 441 313 459 554 277 255 525 271 813 558 480 520 501 396 392 325 546 470 720 483 482 428 387 316 387 712 286 441 528 544 609 316 500 261 851 417 491 554 332 851 278 375 554 407 407 151 544 588 106 285 407 428 491 829 867 866 422 599 599 599 599 599 599 866 519 540 540 540 540 343 343 343 343 662 693 599 599 599 599 599 554 653 687 687 687 687 557 574 605 470 470 470 470 470 470 708 421 441 441 441 441 277 277 277 277 530 558 480 480 480 480 480 554 490 546 546 546 546 482 547 482 500 1000 221 221 375 375 443 752",
    ("Cambria", 1): "220 335 422 618 543 976 740 251 408 408 453 592 232 311 230 505 533 458 507 508 553 520 538 494 556 534 280 280 592 592 592 452 921 653 634 551 664 559 534 586 719 358 340 680 541 845 686 647 593 647 640 473 587 695 632 959 616 587 554 368 505 368 592 371 184 525 568 450 570 513 335 508 600 308 297 589 311 891 604 530 573 559 454 444 365 609 519 779 525 521 464 393 320 393 592 335 469 556 587 604 320 533 294 851 420 531 592 337 851 311 378 592 437 437 183 605 588 140 285 437 436 531 897 904 970 452 653 653 653 653 653 653 880 551 559 559 559 559 358 358 358 358 709 686 647 647 647 647 647 592 695 695 695 695 695 587 617 677 525 525 525 525 525 525 794 450 513 513 513 513 308 308 308 311 572 604 530 530 530 530 530 592 569 609 609 609 609 521 591 521 500 1000 235 235 398 398 443 772",
    ("Cambria", 2): "220 280 389 619 502 858 659 234 372 372 396 510 206 311 181 466 488 394 459 439 490 455 459 442 479 465 181 206 499 531 499 407 885 578 536 495 598 506 497 509 673 330 302 586 474 835 654 581 503 622 574 414 505 623 553 847 545 523 483 339 466 339 528 463 183 515 496 434 513 436 284 493 528 263 243 489 256 800 532 483 502 512 376 363 340 531 431 694 446 431 433 375 315 375 528 280 437 500 536 624 315 475 261 851 426 454 539 340 851 237 329 509 291 281 152 545 569 4 186 268 392 454 865 884 874 407 578 578 578 578 578 578 834 495 506 506 506 506 330 330 330 330 646 654 581 581 581 581 581 482 622 623 623 623 623 523 560 566 515 515 515 515 515 515 709 434 436 436 436 436 263 263 263 263 509 532 483 483 483 483 483 530 483 531 531 531 531 431 521 431 501 1001 209 209 373 373 443 719",
    ("Cambria", 3): "220 286 421 602 517 972 712 244 396 396 399 509 215 301 201 480 541 438 506 510 558 522 529 513 550 529 202 222 566 570 566 438 921 614 608 542 656 568 538 588 695 345 330 659 516 827 670 647 578 659 631 490 567 662 594 922 586 560 543 355 480 355 573 433 213 564 563 467 575 488 333 545 579 295 281 542 291 870 582 554 574 556 452 427 351 589 491 760 482 491 468 371 309 371 573 328 489 619 582 657 309 506 290 851 415 513 577 311 851 311 350 551 318 333 180 594 569 147 211 316 424 509 949 934 972 438 614 614 614 614 614 614 844 542 568 568 568 568 345 345 345 345 683 670 647 647 647 647 647 530 675 662 662 662 662 560 601 646 564 564 564 564 564 564 751 467 488 488 488 488 295 295 295 295 545 582 554 554 554 554 554 569 558 589 589 589 589 491 567 491 508 996 228 228 405 405 443 742",
}
# </widths64>


_ALT_FONT = {"roman": "Times New Roman", "swiss": "Arial", "modern": "Courier New"}
_KIND_HINT = {}                 # this conversion's fonts: {name as the PDF writes it (no subset tag), lower case: kind}
_W_STAT = {}                    # this conversion's stand-in fonts: {(family, bold, italic): [the letters' widths in the PDF, in the stand-in]}
_W_INDEX = dict((ch, i) for i, ch in enumerate(_W_CHARS))
_W_ROWS = {}


def _w_row(fam, bold, ital):
    key = (fam, (1 if bold else 0) + (2 if ital else 0))
    row = _W_ROWS.get(key)
    if row is None and key in _W_TABLE:
        row = _W_ROWS[key] = [int(v) for v in _W_TABLE[key].split()]
    return row


def _font_kinds(src_path, limit=80):
    """What kind of letter each of the PDF's fonts is: {the font's name in the PDF, lower case: "roman" (with serifs) |
    "swiss" (without) | "modern" (every letter as wide as the next)}. From the PDF's own word about the font where it
    says one, else from the font's name."""
    kinds = {}
    try:
        rd = _open_reader(src_path)
        seen = set()
        for pg in rd.pages[:limit]:
            try:
                fonts = (pg.get("/Resources") or {}).get("/Font") or {}
            except Exception:
                continue
            for _k, v in fonts.items():
                try:
                    v = v.get_object()
                    name = str(v.get("/BaseFont") or "").lstrip("/")
                    if not name or name in seen:
                        continue
                    seen.add(name)
                    d = v.get("/FontDescriptor")
                    if d is None and "/DescendantFonts" in v:
                        d = v["/DescendantFonts"][0].get_object().get("/FontDescriptor")
                    flags = int(d.get_object().get("/Flags") or 0) if d is not None else 0
                except Exception:
                    continue
                if "+" in name[:8]:
                    name = name.split("+", 1)[1]
                low = name.lower()
                words = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", name).lower()
                if flags & 1 or _MONO_NAMES.search(words):
                    kinds[low] = "modern"
                elif _SANS_NAMES.search(words) and not (flags & 2):
                    kinds[low] = "swiss"
                elif flags & 2 or _SERIF_NAMES.search(words):
                    kinds[low] = "roman"
    except Exception:
        pass
    return kinds


def _font(psname):
    """(family or None, bold, italic) for a PDF font name."""
    hit = _FONT_CACHE.get(psname)
    if hit is not None:
        return hit
    name = psname or ""
    if "+" in name[:8]:
        name = name.split("+", 1)[1]
    low = re.sub(r"[^a-z0-9,\-]", "", name.lower())
    bold = bool(re.search(r"bold|black|heavy|semibold|demibold|extrabold|ultrabold|[-,]bd|[-,]demi|[-,]b$", low)) or bool(_BOLD_TEX.match(low))
    ital = bool(re.search(r"italic|oblique|slanted|[-,]it$|[-,]i$|[-,]bi$|bdit$|boldit$", low)) or bool(_ITAL_TEX.match(low))
    if re.search(r"[-,]bi$", low):
        bold = True
    base = re.split(r"[-,]", low)[0] if re.split(r"[-,]", low)[0] else low
    flat = re.sub(r"[^a-z0-9]", "", low)
    fam = None
    for rx, f in _FAMILIES:
        if rx.match(flat):
            fam = f
            break
    if fam is None:
        b = re.sub(r"[^a-z0-9]", "", base)
        for _ in range(3):
            b2 = _STYLE_WORDS.sub("", b)
            if b2 == b or not b2:
                break
            b = b2
        raw = re.split(r"[-,]", name)[0]
        raw = re.sub(r"(?i)(PSMT|MT|PS)$", "", raw)
        raw = re.sub(r"(?i)(BoldItalic|BoldOblique|Bold|Italic|Oblique|Regular|Roman)$", "", raw)
        if len(raw) >= 4 and re.search(r"[a-z]", raw) and not re.match(r"^(T3|TT\d|F\d|C\d|R\d|G\d|Font|Fd|g_d|CIDFont|Type)", raw) and re.match(r"^[A-Za-z][A-Za-z0-9 _]+$", raw):
            fam = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", raw.replace("_", " ")).strip()[:31] or None
        # a font no Word has (Melior, Minion, Frutiger): the reader's computer would put in whatever it likes, a
        # different one in every program. It is named by the nearest font every computer has, and its letters are
        # given the width they have in the PDF (see _W_STAT), so the lines still break where they broke.
        kind = _KIND_HINT.get(name.lower())
        if kind is None and fam:
            words = fam.lower()
            kind = "modern" if _MONO_NAMES.search(words) else "swiss" if _SANS_NAMES.search(words) else "roman" if _SERIF_NAMES.search(words) else None
        if kind:
            fam = _ALT_FONT[kind]
    res = (fam, bold, ital, bool(fam) and bool(_EXACT.match(flat)))
    _FONT_CACHE[psname] = res
    return res


def _hexcol(col, ncs=None):
    """A PDF colour -> 'RRGGBB' (None when it is black or cannot be told)."""
    if col is None:
        return None
    if isinstance(col, (int, float)):
        col = (col,)
    try:
        vals = [float(v) for v in col]
    except (TypeError, ValueError):
        return None
    if any(v > 1.0001 for v in vals):
        vals = [v / 255.0 for v in vals]
    vals = [min(1.0, max(0.0, v)) for v in vals]
    n = len(vals)
    if n == 1:
        if ncs and "Gray" not in str(ncs) and str(ncs) not in ("DeviceGray", "CalGray", "G", "None"):
            return None                   # one ink of a colour space we do not know: black is the honest guess
        r = g = b = vals[0]
    elif n == 3:
        r, g, b = vals
    elif n == 4:
        c, m, y, k = vals
        r, g, b = (1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k)
    else:
        return None
    h = "%02X%02X%02X" % (int(round(r * 255)), int(round(g * 255)), int(round(b * 255)))
    return None if h == "000000" else h


# letters a PDF holds in a symbol font's own numbering, as the sign they draw
_PUA = {"\uf0b7": "\u2022", "\uf0a7": "\u25aa", "\uf0d8": "\u27a2", "\uf0fc": "\u2713", "\uf076": "\u2756", "\uf06e": "\u25a0", "\uf071": "\u2751",
        "\uf0a8": "\u25c6", "\uf0e0": "\u2192", "\uf0df": "\u2190", "\uf0e8": "\u2794", "\uf0b0": "\u00b0", "\uf02d": "\u2212", "\uf0b1": "\u00b1",
        "\uf0b4": "\u00d7", "\uf0b8": "\u00f7", "\uf0a3": "\u2264", "\uf0b3": "\u2265", "\uf0b9": "\u2260", "\uf0bb": "\u2248", "\uf0ae": "\u2192",
        "\uf06f": "\u25cb", "\uf075": "\u25c6", "\uf0a2": "\u2032", "\uf0d7": "\u22c5", "\uf0e5": "\u2211", "\uf0d6": "\u221a", "\uf0a5": "\u221e",
        "\uf061": "\u03b1", "\uf062": "\u03b2", "\uf067": "\u03b3", "\uf064": "\u03b4", "\uf065": "\u03b5", "\uf06c": "\u03bb", "\uf06d": "\u03bc",
        "\uf070": "\u03c0", "\uf073": "\u03c3", "\uf044": "\u0394", "\uf057": "\u03a9", "\uf053": "\u03a3", "\uf0fe": "\u2611",
        "\uf0a1": "\u25cb", "\uf0fb": "\u2717", "\uf0f0": "\u21e8", "\uf046": "\u261e", "\uf0d9": "\u21e7"}
_CLOSERS = ".,;:!?)]}%\u2019\u201d\u00bb"
_OPENERS = "([{\u2018\u201c\u00ab"
_LEADERS = (".", "\u00b7", "\u2026")
_LIG = {"\ufb00": "ff", "\ufb01": "fi", "\ufb02": "fl", "\ufb03": "ffi", "\ufb04": "ffl", "\ufb05": "st", "\ufb06": "st"}
_BULLETS = set("\u2022\u25e6\u25aa\u25a0\u25a1\u25cf\u25cb\u25c6\u25c7\u27a2\u25ba\u25b6\u2713\u2714\u2756\u2751\u2023\u2043\u2219\u00b7\u25b8\u25b9\u25e7\u2794\u2192\u261e\u2666\u2605\u2606\u2611\u2717")
_DASH_BULLETS = set("-\u2013\u2014*")
_NUM_MARK = re.compile(r"^(\(?\d{1,3}[.)]|\(?[A-Za-z][.)]|\(?(?:[ivxlc]{1,6}|[IVXLC]{1,6})[.)]|\d{1,2}(?:\.\d{1,2}){1,3}\.?|\[\d{1,3}\])$")


def _clean(t):
    """One letter as Word should hold it."""
    if not t:
        return ""
    if t in _PUA:
        return _PUA[t]
    if t in _LIG:
        return _LIG[t]
    if t.startswith("(cid:"):
        return "\ufffd"
    out = []
    for ch in t:
        o = ord(ch)
        if ch in _LIG:
            out.append(_LIG[ch])
        elif o == 0xAD:
            out.append("-")                # a soft hyphen that is drawn is a hyphen at a line's end
        elif o < 32 or o in (0x7F, 0xFFFE, 0xFFFF) or 0xD800 <= o <= 0xDFFF:
            continue
        elif 0xF000 <= o <= 0xF0FF:
            out.append(_PUA.get(ch, chr(o - 0xF000) if 0x20 < o - 0xF000 < 0x7F else "\u2022"))
        else:
            out.append(ch)
    return "".join(out)


# ----------------------------------------------------------------------------------------------- what a page holds
class Wd(object):
    """A word: letters that touch. parts = [[text, fmt], ...]; fmt = (font, size, bold, italic, colour, raised)."""
    __slots__ = ("x0", "x1", "top", "bottom", "base", "size", "parts", "space", "blk", "art", "ul", "st", "link", "gapspace", "field", "spw", "lead", "xs")

    def text(self):
        return "".join(p[0] for p in self.parts)

    def width(self):
        return self.x1 - self.x0


class Ln(object):
    """A line of one column: words left to right; segs = the parts a tab keeps apart [(first word index, x0, x1)];
    dots = the indexes of the parts that a row of dots leads to (a contents page)."""
    __slots__ = ("words", "x0", "x1", "top", "bottom", "base", "size", "segs", "blk", "allbold", "startbold", "font", "marker", "cen", "rgt", "dots")


class Para(object):
    __slots__ = ("lines", "align", "left", "first", "right", "before", "after", "pitch", "tabs", "level", "L", "R", "top", "bottom", "size",
                 "font", "brk", "runs", "rule", "shade", "keep", "open_end", "open_start", "last_base", "first_base", "cont", "colbrk", "sect", "spec",
                 "hard")


class Pic(object):
    __slots__ = ("x0", "x1", "top", "bottom", "data", "ext", "align", "L", "R", "brk", "before", "after", "left", "page", "colbrk", "sect",
                 "anchor", "behind", "size", "kids", "flow")


class Rule(object):
    __slots__ = ("x0", "x1", "top", "bottom", "width", "color", "L", "R", "brk", "before", "after", "colbrk", "sect")


class Tbl(object):
    __slots__ = ("x0", "x1", "top", "bottom", "xs", "ys", "cells", "L", "R", "brk", "before", "after", "borderless", "pad", "colbrk", "sect",
                 "hdr", "brk_rows")


class Note(object):
    """A footnote: its mark in the PDF ("1", "*"), its paragraphs, and how Word should number it."""
    __slots__ = ("mark", "items", "auto", "id", "page", "found", "sup", "order")


class Cell(object):
    __slots__ = ("r0", "r1", "c0", "c1", "items", "shade", "borders", "valign", "x0", "x1", "top", "bottom", "padl", "padr", "pics")


def _median(v, d=0.0):
    v = sorted(v)
    return v[len(v) // 2] if v else d


def _fmt_of(c):
    return c[5]


# ------------------------------------------------------------------------------------------------ reading one page
class _Row(object):
    __slots__ = ("base", "top", "bottom", "size", "words", "x0", "x1", "sx")


class _Page(object):
    __slots__ = ("n", "w", "h", "rows", "hs", "vs", "fills", "curves", "images", "links", "rot", "blk", "nchars", "bad", "spaces", "items",
                 "head", "foot", "scan", "zones", "body_top", "body_bottom", "ocr", "figs", "tables", "rules", "taken", "hp", "fp", "sig", "hfset", "exact", "ow", "oh",
                 "hfigs", "ffigs", "hrules", "frules", "hpics", "fpics", "design", "back", "skew")


def _read_chars(page):
    """The page's letters: [x0, x1, top, bottom, base, fmt, text, mcid, artifact, size] and the boxes of letters that
    are not upright (turned text is kept as a picture of itself)."""
    out, rot = [], []
    W, H = float(page.width), float(page.height)
    bad = spaces = 0
    exact = [0]
    for c in page.chars:
        t = c.get("text") or ""
        if not t:
            continue
        x0, x1, top, bottom = c["x0"], c["x1"], c["top"], c["bottom"]
        if x1 < -1 or x0 > W + 1 or bottom < -1 or top > H + 1:
            continue
        size = float(c.get("size") or (bottom - top) or 0)
        if size < 1.5 or size > 400:
            continue
        if not c.get("upright", True):
            if not t.isspace():
                rot.append((x0, top, x1, bottom))
            continue
        m = c.get("matrix")
        base = None
        skew = False
        if m and len(m) == 6:
            try:
                base = top + (c["y1"] - m[5])
                skew = abs(m[2]) > 0.12 * abs(m[3]) and abs(m[3]) > 0
            except Exception:
                base = None
        if base is None or not (top - 0.2 * size <= base <= bottom + 0.2 * size):
            base = bottom - 0.22 * size
        if t.isspace() or t == "\xa0":
            spaces += 1
            out.append([x0, x1, top, bottom, base, None, " ", c.get("mcid"), c.get("tag") == "Artifact", size])
            continue
        fam, bold, ital, same = _font(c.get("fontname") or "")
        tx = _clean(t)
        if not tx:
            continue
        if same:
            exact[0] += 1
        elif fam and len(tx) == 1 and tx in _W_INDEX:
            row = _w_row(fam, bold, ital)
            if row is not None:
                sub = row[_W_INDEX[tx]]
                wd = x1 - x0
                if sub > 0 and 0 < wd < 2.5 * size:
                    st = _W_STAT.get((fam, bold, ital))
                    if st is None:
                        st = _W_STAT[(fam, bold, ital)] = [0.0, 0.0, 0]
                    st[0] += wd / size * 1000.0
                    st[1] += sub
                    st[2] += 1
        if "\ufffd" in tx:
            bad += 1
        fmt = (fam, round(size * 2) / 2.0, bold, ital or skew, _hexcol(c.get("non_stroking_color"), c.get("ncs")), "")
        out.append([x0, x1, top, bottom, base, fmt, tx, c.get("mcid"), c.get("tag") == "Artifact", size])
    return out, rot, bad, spaces, exact[0]


def _group_rows(chs):
    """Letters -> rows (one baseline each), with raised and lowered letters brought into the row they belong to.
    Letters join a row when they stand on its baseline AND are of its size: a small raised figure is never taken by a
    line of another column that happens to run at its height."""
    import bisect
    chs.sort(key=lambda c: (c[4], c[0]))
    rows = []                                    # [sum of baselines, count, letters, size of its first letter]
    for c in chs:
        base, size = c[4], c[9]
        hit = None
        for r in reversed(rows[-6:]):
            if abs(base - r[0] / r[1]) <= max(0.8, 0.07 * max(size, r[3])) and 0.78 <= size / r[3] <= 1.28:
                hit = r
                break
        if hit is None:
            rows.append([base, 1, [c], size])
        else:
            hit[0] += base
            hit[1] += 1
            hit[2].append(c)

    def describe(r):
        cs = [c for c in r[2] if c[5] is not None]
        if not cs:
            return None
        sz = Counter()
        for c in cs:
            sz[round(c[9] * 2) / 2.0] += (c[1] - c[0])
        return [r[0] / r[1], sz.most_common(1)[0][0], min(c[0] for c in cs), max(c[1] for c in cs)]
    info = [describe(r) for r in rows]
    # raised / lowered: the letters of a small row that lie inside a bigger row's line, next to its letters
    order = [i for i in range(len(rows)) if info[i]]
    for qi in order:
        q = info[qi]
        if q is None:
            continue
        for ri in order:
            r = info[ri]
            if ri == qi or r is None or r[1] * 0.86 < q[1] or abs(q[0] - r[0]) > 0.62 * r[1]:
                continue
            big = sorted((c for c in rows[ri][2] if c[5] is not None and c[5][5] == ""), key=lambda c: c[0])
            if not big:
                continue
            xs = [c[0] for c in big]
            keep, moved = [], 0
            s = r[1]
            lo, hi = r[0] - 0.80 * s, r[0] + 0.30 * s
            for c in rows[qi][2]:
                if c[5] is None:
                    keep.append(c)
                    continue
                k = bisect.bisect_left(xs, c[0])
                near = False
                for j in (k - 1, k):
                    if 0 <= j < len(big):
                        b = big[j]
                        if max(b[0] - c[1], c[0] - b[1]) <= 0.55 * s:
                            near = True
                ov = min(c[3], hi) - max(c[2], lo)
                if near and ov >= 0.5 * (c[3] - c[2]):
                    f = c[5]
                    if abs(c[4] - r[0]) >= 0.05 * s:
                        c[5] = (f[0], s, f[2], f[3], f[4], "sup" if c[4] < r[0] else "sub")
                    c[4] = r[0]
                    rows[ri][2].append(c)
                    big.insert(bisect.bisect_left(xs, c[0]), c)
                    xs.insert(bisect.bisect_left(xs, c[0]), c[0])
                    moved += 1
                else:
                    keep.append(c)
            if moved:
                rows[qi][2] = keep
                if not any(c[5] is not None for c in keep):
                    info[qi] = None
                    break
                info[qi] = describe(rows[qi])
                q = info[qi]
    # what is left on one baseline in two sizes, side by side (a big figure before a heading's words), is one row
    live = [i for i in range(len(rows)) if info[i] is not None]
    live.sort(key=lambda i: (info[i][0], info[i][2]))
    for a_i in range(len(live)):
        i = live[a_i]
        if info[i] is None:
            continue
        for j in live[a_i + 1:a_i + 5]:
            if info[j] is None or abs(info[j][0] - info[i][0]) > 0.8:
                continue
            gap = max(info[j][2] - info[i][3], info[i][2] - info[j][3])
            if gap <= 1.2 * max(info[i][1], info[j][1]):
                rows[i][2].extend(rows[j][2])
                info[j] = None
                info[i] = describe(rows[i])
    out = []
    for i, r in enumerate(rows):
        if info[i] is None or not any(c[5] is not None for c in r[2]):
            continue
        out.append((info[i][0], info[i][1], sorted(r[2], key=lambda c: (c[0], c[1]))))
    out.sort(key=lambda r: r[0])
    return out


_KEEP_LEADERS = False          # PowerPoint has no dotted tabs: for slides a row of dots stays the dots it is


def _make_words(base, dom, cs, spaced=True):
    """One row's letters -> (words, where a space after the last word begins or None). spaced: does this page write
    its spaces as letters? Then a small gap with no space in it is still a break (a tab that had no room)."""
    words = []
    lim = 0.13 if spaced else 0.17
    # a row of five or more dots (a contents page's leader) is not letters: it is a gap the next word is led across
    if not _KEEP_LEADERS and sum(1 for c in cs if c[6] in _LEADERS) >= 5:
        out, i, n = [], 0, len(cs)
        while i < n:
            if cs[i][5] is not None and cs[i][6] in _LEADERS:
                j = i
                while j < n and ((cs[j][5] is not None and cs[j][6] in _LEADERS) or (cs[j][5] is None and j + 1 < n and cs[j + 1][5] is not None and cs[j + 1][6] in _LEADERS)):
                    j += 1
                if sum(1 for c in cs[i:j] if c[5] is not None) >= 5:
                    c0 = cs[i]
                    out.append([c0[0], cs[j - 1][1], c0[2], c0[3], c0[4], None, "\x00", c0[7], c0[8], c0[9]])
                    i = j
                    continue
            out.append(cs[i])
            i += 1
        cs = out
    cur = prev = None
    pend = False
    led = False
    spw = 0.0
    for c in cs:
        if c[5] is None:                          # a space that is in the PDF (or a row of leader dots)
            pend = True
            if c[6] == "\x00":
                led = True
            else:
                spw = c[1] - c[0]
            continue
        s = max(c[9], prev[9]) if prev else c[9]
        gap = (c[0] - prev[1]) if prev else 0.0
        if prev is not None and not pend and c[6] == prev[6] and abs(c[0] - prev[0]) < 0.12 * s and c[5] == prev[5]:
            continue                              # the same letter printed twice (a poor man's bold)
        if cur is not None and not pend and lim * s < gap <= 0.22 * s and (c[6][:1] in _CLOSERS or prev[6][-1:] in _OPENERS):
            gap = 0.0                             # a full stop (or a bracket) a hair away from its word still belongs to it
        if cur is None or pend or gap > lim * s or gap < -0.7 * s:
            cur = Wd()
            cur.x0, cur.x1, cur.top, cur.bottom = c[0], c[1], c[2], c[3]
            cur.base, cur.size, cur.parts = base, dom, []
            cur.space = prev is not None and (pend or gap > lim * s)
            cur.gapspace = not pend
            cur.blk, cur.art, cur.ul, cur.st, cur.link, cur.field = c[7], True, False, False, None, None
            cur.xs = []
            cur.spw = spw if pend else 0.0
            cur.lead = led and prev is not None
            led = False
            spw = 0.0
            words.append(cur)
        else:
            if c[1] > cur.x1:
                cur.x1 = c[1]
        if c[5][5] == "":
            if c[2] < cur.top:
                cur.top = c[2]
            if c[3] > cur.bottom:
                cur.bottom = c[3]
        if not c[8]:
            cur.art = False
        if cur.blk is None:
            cur.blk = c[7]
        if cur.parts and cur.parts[-1][1] == c[5]:
            cur.parts[-1][0] += c[6]
        else:
            cur.parts.append([c[6], c[5]])
        nl = len(c[6])
        for k in range(nl):                      # where each letter ends (a ligature's letters share its width)
            cur.xs.append(c[0] + (c[1] - c[0]) * (k + 1) / float(nl))
        prev, pend = c, False
    for w in words:
        sz = Counter()
        for t, f in (q[:2] for q in w.parts):
            if f[5] == "":
                sz[f[1]] += len(t)
        if sz:
            w.size = sz.most_common(1)[0][0]
    tail = None
    if pend and prev is not None:
        for c in reversed(cs):
            if c[5] is None:
                tail = c[0]
                break
    return words, tail


def _blockmap(page):
    """For a tagged PDF: marked-content id -> (the paragraph it belongs to, its kind). {} when the PDF has no tags."""
    out = {}
    try:
        tree = page.structure_tree
    except Exception:
        return out
    if not tree:
        return out
    BLOCK = {"P", "H", "H1", "H2", "H3", "H4", "H5", "H6", "Caption", "TOCI", "BlockQuote", "Code", "Formula", "Note", "FENote", "Title", "BibEntry", "Index"}
    counter = [0]

    def walk(node, cur, depth, li):
        if depth > 60 or not isinstance(node, dict):
            return
        typ = node.get("type") or ""
        me = cur
        if typ in BLOCK:
            counter[0] += 1
            me = (counter[0], typ)
            if li is not None and li[0] is not None and li[1] is None:
                li[1] = me                           # the item's first paragraph: its label joins it
                for m in li[2]:
                    out[m] = me
        elif typ == "LI":
            li = [True, None, []]
        elif typ in ("TD", "TH") and cur is None:
            counter[0] += 1
            me = (counter[0], "P")
        for m in node.get("mcids") or []:
            if typ == "Lbl" and li is not None:
                if li[1] is not None:
                    out[m] = li[1]
                else:
                    li[2].append(m)
            elif me is not None:
                out[m] = me
            elif typ in ("Span", "Link", "Reference", "Quote", "Em", "Strong", "LBody", "Lbl", "Div", "Sect", "Part", "Art", "NonStruct", "Document", "L", "LI"):
                counter[0] += 1
                out[m] = (counter[0], "P")
        for ch in node.get("children") or []:
            walk(ch, me, depth + 1, li)

    try:
        for n in tree:
            walk(n, None, 0, None)
    except Exception:
        return {}
    return out


def _read_vectors(page):
    """Ruled lines (horizontal, vertical), filled boxes, other drawing, pictures: what is drawn that is not a letter."""
    hs, vs, fills, curves, images = [], [], [], [], []
    W, H = float(page.width), float(page.height)

    def seg(x0, top, x1, bottom, lw, col, real=True):
        if x1 < x0:
            x0, x1 = x1, x0
        if bottom < top:
            top, bottom = bottom, top
        w, h = x1 - x0, bottom - top
        if w < 0.2 and h < 0.2:
            return
        if h <= 1.2 and w > h:
            hs.append([x0, x1, (top + bottom) / 2.0, max(lw, h, 0.25), col, real, False])
        elif w <= 1.2 and h > w:
            vs.append([top, bottom, (x0 + x1) / 2.0, max(lw, w, 0.25), col, real, False])
        else:
            curves.append((x0, top, x1, bottom))

    try:
        lines = page.lines
    except Exception:
        lines = []
    for l in lines:
        if not l.get("stroke", True) and not l.get("fill"):
            continue
        seg(l["x0"], l["top"], l["x1"], l["bottom"], float(l.get("linewidth") or 0.5), _hexcol(l.get("stroking_color")))
    try:
        rects = page.rects
    except Exception:
        rects = []
    for r in rects:
        x0, x1, top, bottom = r["x0"], r["x1"], r["top"], r["bottom"]
        w, h = x1 - x0, bottom - top
        if w <= 0 and h <= 0:
            continue
        if x1 < 0 or x0 > W or bottom < 0 or top > H:
            continue
        fill, stroke = r.get("fill"), r.get("stroke")
        lw = float(r.get("linewidth") or 0)
        if fill and min(w, h) <= 2.2 and max(w, h) > min(w, h):
            seg(x0, top, x1, bottom, min(w, h), _hexcol(r.get("non_stroking_color")))
            continue
        if fill:
            col = _hexcol(r.get("non_stroking_color"))
            raw = r.get("non_stroking_color")
            black = col is None and raw is not None
            if col != "FFFFFF" and (col is not None or black):
                fills.append([x0, top, x1, bottom, col or "000000", False])
        if stroke and lw > 0:
            col = _hexcol(r.get("stroking_color"))
            seg(x0, top, x1, top, lw, col)
            seg(x0, bottom, x1, bottom, lw, col)
            seg(x0, top, x0, bottom, lw, col)
            seg(x1, top, x1, bottom, lw, col)
    try:
        cvs = page.curves
    except Exception:
        cvs = []
    for c in cvs:
        x0, x1, top, bottom = c["x0"], c["x1"], c["top"], c["bottom"]
        if x1 < 0 or x0 > W or bottom < 0 or top > H:
            continue
        if not c.get("stroke") and not c.get("fill"):
            continue
        if c.get("fill") and not c.get("stroke") and _hexcol(c.get("non_stroking_color")) == "FFFFFF":
            continue
        curves.append((max(0, x0), max(0, top), min(W, x1), min(H, bottom)))
    try:
        ims = page.images
    except Exception:
        ims = []
    for im in ims:
        x0, x1, top, bottom = max(0.0, im["x0"]), min(W, im["x1"]), max(0.0, im["top"]), min(H, im["bottom"])
        if x1 - x0 < 2 or bottom - top < 2:
            continue
        src = im.get("srcsize") or (0, 0)
        images.append([x0, top, x1, bottom, src, bool(im.get("imagemask")), im.get("tag") == "Artifact"])
    return hs, vs, fills, curves, images


def _read_links(page):
    out = []
    try:
        for h in page.hyperlinks:
            uri = h.get("uri")
            if isinstance(uri, bytes):
                uri = uri.decode("utf-8", "replace")
            if uri and re.match(r"^(https?://|mailto:|ftp://|tel:)", uri, re.I):
                out.append((h["x0"], h["top"], h["x1"], h["bottom"], uri))
    except Exception:
        pass
    return out


def _read_page(page, n):
    pg = _Page()
    pg.n, pg.w, pg.h = n, float(page.width), float(page.height)
    pg.ow, pg.oh = pg.w, pg.h
    chs, pg.rot, bad, spaces, pg.exact = _read_chars(page)
    pg.nchars = sum(1 for c in chs if c[5] is not None)
    pg.bad, pg.spaces = bad, spaces
    pg.blk = _blockmap(page) if any(c[7] is not None for c in chs[:400]) else {}
    pg.hs, pg.vs, pg.fills, pg.curves, pg.images = _read_vectors(page)
    pg.links = _read_links(page)
    rows = []
    spaced = spaces > 0.02 * max(1, pg.nchars)
    for base, dom, cs in _group_rows(chs):
        ws, tail = _make_words(base, dom, cs, spaced)
        if not ws:
            continue
        r = _Row()
        r.base, r.size, r.words, r.sx = base, dom, ws, tail
        r.top, r.bottom = min(w.top for w in ws), max(w.bottom for w in ws)
        r.x0, r.x1 = ws[0].x0, max(w.x1 for w in ws)
        rows.append(r)
    pg.rows = rows
    if pg.blk:
        for r in rows:
            for w in r.words:
                w.blk = pg.blk.get(w.blk)
    else:
        for r in rows:
            for w in r.words:
                w.blk = None
    pg.items, pg.head, pg.foot, pg.scan, pg.zones, pg.ocr = [], [], [], False, [], False
    pg.figs, pg.tables, pg.rules, pg.taken = [], [], [], None
    pg.hfigs, pg.ffigs, pg.hrules, pg.frules, pg.hpics, pg.fpics, pg.design, pg.back = [], [], [], [], [], [], False, None
    pg.body_top = pg.body_bottom = None
    return pg


# ----------------------------------------------------------------------------------------------------------- tables
def _cluster(vals, tol=2.0):
    """Sorted numbers, near ones made one (their mean)."""
    out = []
    for v in sorted(vals):
        if out and v - out[-1][-1] <= tol:
            out[-1].append(v)
        else:
            out.append([v])
    return [sum(g) / len(g) for g in out]


def _near(vals, v):
    best, bi = None, -1
    for i, x in enumerate(vals):
        d = abs(x - v)
        if best is None or d < best:
            best, bi = d, i
    return bi


def _find_tables(page, pg):
    """Ruled tables: the grid the drawn lines (and the edges of filled cells) make. Returns [Tbl] with empty cells'
    words not yet put in."""
    if len(pg.hs) + len(pg.fills) < 2:
        return []
    hl, vl = [], []
    for s in pg.hs:
        if s[1] - s[0] >= 6:
            hl.append({"object_type": "line", "x0": s[0], "x1": s[1], "top": s[2], "bottom": s[2], "width": s[1] - s[0], "height": 0})
    for s in pg.vs:
        if s[1] - s[0] >= 6:
            vl.append({"object_type": "line", "x0": s[2], "x1": s[2], "top": s[0], "bottom": s[1], "width": 0, "height": s[1] - s[0]})
    for f in pg.fills:
        x0, top, x1, bottom = f[:4]
        if x1 - x0 < 8 or bottom - top < 6 or (x1 - x0) * (bottom - top) > 0.6 * pg.w * pg.h:
            continue
        hl.append({"object_type": "line", "x0": x0, "x1": x1, "top": top, "bottom": top, "width": x1 - x0, "height": 0})
        hl.append({"object_type": "line", "x0": x0, "x1": x1, "top": bottom, "bottom": bottom, "width": x1 - x0, "height": 0})
        vl.append({"object_type": "line", "x0": x0, "x1": x0, "top": top, "bottom": bottom, "width": 0, "height": bottom - top})
        vl.append({"object_type": "line", "x0": x1, "x1": x1, "top": top, "bottom": bottom, "width": 0, "height": bottom - top})
    if len(hl) < 2 or len(vl) < 2:
        return []
    try:
        from pdfplumber.table import TableFinder
        tf = TableFinder(page, {"vertical_strategy": "explicit", "horizontal_strategy": "explicit", "explicit_vertical_lines": vl,
                                "explicit_horizontal_lines": hl, "snap_tolerance": 2.5, "join_tolerance": 2.5, "intersection_tolerance": 2.5,
                                "edge_min_length": 4})
        found = tf.tables
    except Exception:
        return []
    out = []
    for t in found:
        cells = [c for c in (t.cells or []) if c and c[2] - c[0] > 3 and c[3] - c[1] > 3]
        if not cells:
            continue
        x0, top, x1, bottom = min(c[0] for c in cells), min(c[1] for c in cells), max(c[2] for c in cells), max(c[3] for c in cells)
        if (x1 - x0) * (bottom - top) > 0.82 * pg.w * pg.h and len(cells) <= 2:
            continue                                  # a frame round the whole page is not a table
        xs = _cluster([c[0] for c in cells] + [c[2] for c in cells])
        ys = _cluster([c[1] for c in cells] + [c[3] for c in cells])
        if len(xs) < 2 or len(ys) < 2:
            continue
        tb = Tbl()
        tb.x0, tb.x1, tb.top, tb.bottom, tb.xs, tb.ys, tb.cells, tb.borderless, tb.brk, tb.before, tb.pad = x0, x1, top, bottom, xs, ys, [], False, False, 0.0, None
        tb.after, tb.colbrk, tb.sect, tb.L, tb.R = 0.0, False, None, x0, x1
        tb.hdr, tb.brk_rows = 0, set()
        seen = set()
        for c in sorted(cells, key=lambda c: (c[1], c[0])):
            c0, c1, r0, r1 = _near(xs, c[0]), _near(xs, c[2]), _near(ys, c[1]), _near(ys, c[3])
            if c1 <= c0 or r1 <= r0 or (r0, c0) in seen:
                continue
            seen.add((r0, c0))
            ce = Cell()
            ce.r0, ce.r1, ce.c0, ce.c1, ce.items, ce.shade, ce.valign = r0, r1, c0, c1, [], None, "top"
            ce.padl = ce.padr = None
            ce.pics = []
            ce.x0, ce.x1, ce.top, ce.bottom = xs[c0], xs[c1], ys[r0], ys[r1]
            ce.borders = {}
            tb.cells.append(ce)
        if tb.cells:
            out.append(tb)
    # one table must not sit inside another: keep the bigger
    out.sort(key=lambda t: -(t.x1 - t.x0) * (t.bottom - t.top))
    keep = []
    for t in out:
        if any(k.x0 - 2 <= t.x0 and t.x1 <= k.x1 + 2 and k.top - 2 <= t.top and t.bottom <= k.bottom + 2 for k in keep):
            continue
        keep.append(t)
    keep.sort(key=lambda t: (t.top, t.x0))
    return keep


def _table_looks(tb, pg):
    """Each cell's borders (from the lines that are really drawn along its edges) and its shading."""
    for ce in tb.cells:
        w, h = ce.x1 - ce.x0, ce.bottom - ce.top
        for side, y in (("top", ce.top), ("bottom", ce.bottom)):
            best = None
            for s in pg.hs:
                if abs(s[2] - y) <= 2.2 and s[5]:
                    ov = min(s[1], ce.x1) - max(s[0], ce.x0)
                    if ov >= 0.6 * w and (best is None or s[3] > best[3]):
                        best = s
            if best is not None:
                ce.borders[side] = (best[3], best[4])
                best[6] = True
        for side, x in (("left", ce.x0), ("right", ce.x1)):
            best = None
            for s in pg.vs:
                if abs(s[2] - x) <= 2.2 and s[5]:
                    ov = min(s[1], ce.bottom) - max(s[0], ce.top)
                    if ov >= 0.6 * h and (best is None or s[3] > best[3]):
                        best = s
            if best is not None:
                ce.borders[side] = (best[3], best[4])
                best[6] = True
        area = max(1.0, w * h)
        for f in pg.fills:
            ov = max(0.0, min(f[2], ce.x1) - max(f[0], ce.x0)) * max(0.0, min(f[3], ce.bottom) - max(f[1], ce.top))
            if ov >= 0.7 * area:
                ce.shade = f[4]
                f[5] = True
    # lines that are part of the grid are used up even when no cell claimed them (double rules, outer frames)
    for s in pg.hs:
        if not s[6] and tb.top - 2.5 <= s[2] <= tb.bottom + 2.5 and s[0] >= tb.x0 - 3 and s[1] <= tb.x1 + 3:
            s[6] = True
    for s in pg.vs:
        if not s[6] and tb.x0 - 2.5 <= s[2] <= tb.x1 + 2.5 and s[0] >= tb.top - 3 and s[1] <= tb.bottom + 3:
            s[6] = True
    for f in pg.fills:
        if not f[5] and f[0] >= tb.x0 - 3 and f[2] <= tb.x1 + 3 and f[1] >= tb.top - 3 and f[3] <= tb.bottom + 3:
            f[5] = True


# ---------------------------------------------------------------------------------- underlines, rules, drawn figures
_TAIL_MARKS = ".,;:!?)]}\u201d\u2019\"'\u00bb\u2026"
_HEAD_MARKS = "([{\u201c\u2018\"'\u00ab"


def _split_word(r, w, k):
    """Cut the word w of row r after its k-th letter: the rest becomes a word of its own that follows with no space.
    Returns the new word, or None when the word cannot be cut there."""
    xs = w.xs
    if not xs or w.field or k <= 0 or k >= len(xs) or len(xs) != len(w.text()) or any(len(q) > 2 for q in w.parts):
        return None
    a_parts, b_parts, n = [], [], 0
    for t, f in w.parts:
        if n + len(t) <= k:
            a_parts.append([t, f])
        elif n >= k:
            b_parts.append([t, f])
        else:
            a_parts.append([t[:k - n], f])
            b_parts.append([t[k - n:], f])
        n += len(t)
    b = Wd()
    b.x0, b.x1, b.top, b.bottom, b.base, b.size = xs[k - 1], w.x1, w.top, w.bottom, w.base, w.size
    b.parts, b.space, b.gapspace, b.blk, b.art = b_parts, False, False, w.blk, w.art
    b.ul, b.st, b.link, b.field, b.spw, b.lead, b.xs = w.ul, w.st, w.link, None, 0.0, False, xs[k:]
    w.parts, w.x1, w.xs = a_parts, xs[k - 1], xs[:k]
    r.words.insert(r.words.index(w) + 1, b)
    return b


def _trim_to(r, w, x0, x1):
    """A line or a link's box runs from x0 to x1, and the word w lies on it but for its closing marks (the full stop
    after a link) or its opening ones (the bracket before underlined words). Those are cut off: returns the word
    that is truly on it."""
    xs = w.xs
    if not xs or len(xs) != len(w.text()):
        return w
    t = w.text()
    n = len(t)
    k = n
    while k > 1 and t[k - 1] in _TAIL_MARKS and (xs[k - 2] + xs[k - 1]) / 2.0 > x1 + 0.3:
        k -= 1
    if k < n:
        _split_word(r, w, k)
        t = w.text()
        xs = w.xs
        n = len(t)
    k = 0
    while k < n - 1 and t[k] in _HEAD_MARKS and ((xs[k - 1] if k else w.x0) + xs[k]) / 2.0 < x0 - 0.3:
        k += 1
    if k:
        b = _split_word(r, w, k)
        if b is not None:
            return b
    return w


def _underlines(pg):
    """A thin line just under words is their underline; one through their middle strikes them out."""
    for s in pg.hs:
        if s[6] or s[1] - s[0] < 2:
            continue
        for r in pg.rows:
            d = s[2] - r.base
            if not (-0.5 * r.size <= d <= 0.32 * r.size + 1.0):
                continue
            hit = [w for w in r.words if min(w.x1, s[1]) - max(w.x0, s[0]) > 0.5 * (w.x1 - w.x0)]
            if not hit:
                continue
            cov = sum(min(w.x1, s[1]) - max(w.x0, s[0]) for w in hit)
            span = max(w.x1 for w in hit) - min(w.x0 for w in hit)
            if span < 0.75 * (s[1] - s[0]) or s[3] > 0.2 * r.size:
                continue
            if d >= 0.02 * r.size:
                for w in hit:
                    _trim_to(r, w, s[0], s[1]).ul = True
                s[6] = True
            elif -0.42 * r.size <= d <= -0.15 * r.size and cov > 0:
                for w in hit:
                    _trim_to(r, w, s[0], s[1]).st = True
                s[6] = True
            break


def _boxes_touch(a, b, pad):
    return a[0] - pad <= b[2] and b[0] - pad <= a[2] and a[1] - pad <= b[3] and b[1] - pad <= a[3]


def _find_figures(pg):
    """Pictures, and drawing that is not a table, an underline or a plain rule: each cluster is one figure, kept as a
    picture cut from the page. Returns boxes (x0, top, x1, bottom). Plain long rules are set aside in pg.rules."""
    atoms = []                                               # (x0, top, x1, bottom, kind)
    page_area = pg.w * pg.h
    full = None
    for im in pg.images:
        a = (im[2] - im[0]) * (im[3] - im[1])
        if a >= 0.6 * page_area:
            full = im                                        # a picture as big as the page: a scan, or a background
            continue
        atoms.append((im[0], im[1], im[2], im[3], "img"))
    for c in pg.curves:
        if (c[2] - c[0]) * (c[3] - c[1]) >= 0.85 * page_area:
            continue
        atoms.append((c[0], c[1], c[2], c[3], "curve"))
    for r in pg.rot:
        atoms.append((r[0], r[1], r[2], r[3], "rot"))
    loose_h = [s for s in pg.hs if not s[6]]
    loose_v = [s for s in pg.vs if not s[6]]
    for s in loose_v:
        atoms.append((s[2] - 0.5, s[0], s[2] + 0.5, s[1], "v"))
    for f in pg.fills:
        if f[5]:
            continue
        a = (f[2] - f[0]) * (f[3] - f[1])
        if a >= 0.5 * page_area:
            continue                                         # a wash over the page
        atoms.append((f[0], f[1], f[2], f[3], "fill"))
    # a horizontal line is a plain rule unless other drawing touches it
    rules = []
    others = list(atoms)
    for s in loose_h:
        box = (s[0], s[2] - 0.5, s[1], s[2] + 0.5)
        if any(_boxes_touch(box, o, 3.0) for o in others):
            atoms.append((box[0], box[1], box[2], box[3], "h"))
        elif s[1] - s[0] >= 18:
            rules.append(s)
        else:
            atoms.append((box[0], box[1], box[2], box[3], "h"))
    pg.rules = rules
    if not atoms:
        return [], full
    # clusters of atoms that touch (6 pt)
    par = list(range(len(atoms)))

    def find(i):
        while par[i] != i:
            par[i] = par[par[i]]
            i = par[i]
        return i
    order = sorted(range(len(atoms)), key=lambda i: atoms[i][0])
    for ii, i in enumerate(order):
        a = atoms[i]
        for j in order[ii + 1:]:
            b = atoms[j]
            if b[0] > a[2] + 6:
                break
            if _boxes_touch(a, b, 6.0):
                ra, rb = find(i), find(j)
                if ra != rb:
                    par[ra] = rb
    groups = defaultdict(list)
    for i in range(len(atoms)):
        groups[find(i)].append(atoms[i])
    figs = []
    for g in groups.values():
        x0, top, x1, bottom = min(a[0] for a in g), min(a[1] for a in g), max(a[2] for a in g), max(a[3] for a in g)
        kinds = Counter(a[4] for a in g)
        w, h = x1 - x0, bottom - top
        has_img = kinds.get("img", 0) > 0
        if not has_img and not kinds.get("rot"):
            if w < 14 and h < 14:
                continue                                     # a dot, a tick box, a drawn bullet: too small to be a figure
            if kinds.get("curve", 0) == 0 and len(g) < 3 and kinds.get("fill", 0) == len(g):
                # one or two plain filled boxes: a band of colour behind words, not a figure
                continue
            if kinds.get("curve", 0) == 0 and len(g) <= 2 and (w < 20 or h < 20):
                continue
        figs.append([max(0.0, x0), max(0.0, top), min(pg.w, x1), min(pg.h, bottom), has_img, len(g), kinds.get("curve", 0)])
    return figs, full


def _settle_figures(pg, figs):
    """Words that lie on a figure are part of its picture (axis labels, a legend); a 'figure' that is really a boxed
    or tinted block of running text is left as text. Sets pg.figs to the final boxes and takes their words out of
    the rows."""
    final = []
    for f in figs:
        box = f[:4]
        for _ in range(2):
            inside = []
            for r in pg.rows:
                for w in r.words:
                    cx, cy = (w.x0 + w.x1) / 2.0, (w.top + w.bottom) / 2.0
                    if box[0] - 1 <= cx <= box[2] + 1 and box[1] - 1 <= cy <= box[3] + 1:
                        inside.append((r, w))
            nb = list(box)
            for r, w in inside:
                nb[0], nb[1], nb[2], nb[3] = min(nb[0], w.x0), min(nb[1], w.top), max(nb[2], w.x1), max(nb[3], w.bottom)
            if nb == list(box):
                break
            box = nb
        rows_in = defaultdict(int)
        for r, w in inside:
            rows_in[id(r)] += 1
        long_rows = sum(1 for v in rows_in.values() if v >= 6)
        area = (box[2] - box[0]) * (box[3] - box[1])
        if not f[4] and long_rows >= 3 and area > 0.12 * pg.w * pg.h:
            continue                                         # running text with decoration round it: keep the text
        if not f[4] and long_rows >= 2 and len(inside) > 40 and f[5] <= 6:
            continue
        final.append((box[0], box[1], box[2], box[3], inside))
    # two figures that now overlap are one
    merged = []
    for b in sorted(final, key=lambda b: (b[1], b[0])):
        for m in merged:
            if _boxes_touch(m, b, 2.0):
                m[0], m[1], m[2], m[3] = min(m[0], b[0]), min(m[1], b[1]), max(m[2], b[2]), max(m[3], b[3])
                m[4].extend(b[4])
                break
        else:
            merged.append([b[0], b[1], b[2], b[3], list(b[4])])
    gone = set()
    for m in merged:
        for r, w in m[4]:
            gone.add(id(w))
    if gone:
        for r in pg.rows:
            r.words = [w for w in r.words if id(w) not in gone]
        pg.rows = [r for r in pg.rows if r.words]
        for r in pg.rows:
            r.x0, r.x1 = r.words[0].x0, max(w.x1 for w in r.words)
    pg.figs = [(m[0], m[1], m[2], m[3]) for m in merged]


# ------------------------------------------------------------------------------------------ columns and the order
class _Band(object):
    """One thing in the page's top-to-bottom order: a row of words, or a block (table, figure, rule)."""
    __slots__ = ("top", "bottom", "x0", "x1", "row", "block", "spans")


def _row_spans(words):
    """The stretches of a row that words cover (gaps a word space would not make stay open)."""
    out = []
    for w in words:
        if out and w.x0 - out[-1][1] <= max(0.95 * w.size, 6.0):
            out[-1][1] = max(out[-1][1], w.x1)
        else:
            out.append([w.x0, w.x1])
    return out


def _bands_of(rows, blocks):
    bands = []
    for r in rows:
        b = _Band()
        b.top, b.bottom, b.x0, b.x1, b.row, b.block = r.top, r.bottom, r.x0, r.x1, r, None
        b.spans = _row_spans(r.words)
        bands.append(b)
    for k in blocks:
        b = _Band()
        b.top, b.bottom, b.x0, b.x1, b.row, b.block = k.top, k.bottom, k.x0, k.x1, None, k
        b.spans = [[k.x0, k.x1]]
        bands.append(b)
    bands.sort(key=lambda b: ((b.row.base if b.row is not None else b.top + 0.01), b.x0))
    return bands


def _split_row(r, g0, g1):
    """A row cut at a gutter -> (left row or None, right row or None)."""
    mid = (g0 + g1) / 2.0
    lw = [w for w in r.words if (w.x0 + w.x1) / 2.0 < mid]
    rw = [w for w in r.words if (w.x0 + w.x1) / 2.0 >= mid]
    res = []
    for ws in (lw, rw):
        if not ws:
            res.append(None)
            continue
        n = _Row()
        n.base, n.size, n.words, n.sx = r.base, r.size, ws, (r.sx if (ws is rw and r.sx is not None) else None)
        sz = Counter()
        for w in ws:
            sz[w.size] += w.x1 - w.x0
        n.size = sz.most_common(1)[0][0]
        n.top, n.bottom, n.x0, n.x1 = min(w.top for w in ws), max(w.bottom for w in ws), ws[0].x0, max(w.x1 for w in ws)
        ws[0].space = False
        res.append(n)
    return res


def _regroup_rows(rows):
    """The rows of one column, after the page's rows were cut at a gutter: words on one baseline are one row again, and
    a small raised or lowered word (a note's number, a power) joins the line it belongs to."""
    rows = sorted(rows, key=lambda r: (r.base, r.x0))
    out = []
    for r in rows:
        hit = None
        for q in reversed(out[-3:]):
            if abs(q.base - r.base) <= max(0.8, 0.07 * max(q.size, r.size)) and 0.78 <= r.size / q.size <= 1.28:
                hit = q
                break
        if hit is None:
            out.append(r)
            continue
        ws = sorted(hit.words + r.words, key=lambda w: w.x0)
        for i, w in enumerate(ws):
            if i and w is r.words[0] or (i and w is hit.words[0]):
                w.space = (w.x0 - ws[i - 1].x1) > 0.13 * w.size
        ws[0].space = False
        hit.words = ws
        hit.top, hit.bottom = min(hit.top, r.top), max(hit.bottom, r.bottom)
        hit.x0, hit.x1 = ws[0].x0, max(w.x1 for w in ws)
        if r.sx is not None and (hit.sx is None or r.sx > hit.sx):
            hit.sx = r.sx
    for q in list(out):
        if not q.words:
            continue
        for r in out:
            if r is q or not r.words or r.size * 0.86 < q.size or abs(q.base - r.base) > 0.62 * r.size:
                continue
            s = r.size
            lo, hi = r.base - 0.80 * s, r.base + 0.30 * s
            keep = []
            for w in q.words:
                near = any(max(v.x0 - w.x1, w.x0 - v.x1) <= 0.55 * s for v in r.words)
                ov = min(w.bottom, hi) - max(w.top, lo)
                if near and ov >= 0.5 * (w.bottom - w.top):
                    va = "sup" if w.base < r.base else "sub"
                    for part in w.parts:
                        f = part[1]
                        part[1] = (f[0], s, f[2], f[3], f[4], va if abs(w.base - r.base) >= 0.05 * s else f[5])
                    w.base, w.size = r.base, s
                    left = [v for v in r.words if v.x0 <= w.x0]
                    w.space = bool(left) and (w.x0 - max(v.x1 for v in left)) > 0.13 * s
                    r.words.append(w)
                    r.words.sort(key=lambda v: v.x0)
                    k = r.words.index(w)
                    if k + 1 < len(r.words):
                        nx = r.words[k + 1]
                        nx.space = (nx.x0 - w.x1) > 0.13 * s
                    r.x0, r.x1 = r.words[0].x0, max(v.x1 for v in r.words)
                else:
                    keep.append(w)
            q.words = keep
            if not keep:
                break
    out = [r for r in out if r.words]
    for r in out:
        r.words[0].space = False
        r.x0, r.x1 = r.words[0].x0, max(w.x1 for w in r.words)
    return out


def _zones(bands, XL, XR, depth=0):
    """Bands in page order -> zones: ("one", [bands]) or ("cols", [[zones of column 1], [zones of column 2], ...],
    [(x0, x1) of each column]). A gutter is an upright strip of the page that nothing touches, over a height of many
    lines, with real text on both of its sides. The page is looked at as it is drawn (every word's own place), not row
    by row: lines of two columns need not stand at the same height. Words that only stand apart on their lines (marks
    at the right, a small grid) are not columns."""
    n = len(bands)
    if n < 6 or depth > 3 or XR - XL < 120:
        return [("one", bands)] if bands else []
    lo, hi = int(math.floor(XL)), int(math.ceil(XR))
    width = hi - lo
    if width <= 0:
        return [("one", bands)]
    top = min(b.top for b in bands)
    bottom = max(b.bottom for b in bands)
    HY = 2.0
    ny = int((bottom - top) / HY) + 2
    occ = [None] * ny

    def mark(x0, x1, y0, y1):
        a, z = max(0, int(math.floor(x0)) - lo), min(width, int(math.ceil(x1)) - lo)
        if z < a:
            return
        for y in range(max(0, int((y0 - top) / HY)), min(ny - 1, int((y1 - top) / HY)) + 1):
            row = occ[y]
            if row is None:
                row = occ[y] = bytearray(width + 1)
            for x in range(a, z + 1):
                row[x] = 1
    for b in bands:
        if b.row is None:
            mark(b.x0, b.x1, b.top, b.bottom)
            continue
        ws = b.row.words
        i = 0
        while i < len(ws):
            j = i
            x1, t0, t1 = ws[i].x1, ws[i].top, ws[i].bottom
            while j + 1 < len(ws) and ws[j + 1].x0 - x1 <= max(0.95 * ws[j + 1].size, 6.0):
                j += 1
                x1 = max(x1, ws[j].x1)
                t0, t1 = min(t0, ws[j].top), max(t1, ws[j].bottom)
            # the middle of the letters' height marks the strip (so that neighbouring lines do not run into each other)
            h = t1 - t0
            mark(ws[i].x0, x1, t0 + 0.15 * h, t1 - 0.15 * h)
            i = j + 1
    active = [y for y in range(ny) if occ[y] is not None]
    na = len(active)
    if na < 12:
        return [("one", bands)]
    empty = [0] * (width + 1)
    for y in active:
        row = occ[y]
        for x in range(width + 1):
            if not row[x]:
                empty[x] += 1
    need = max(10, int(0.2 * na))
    cands = []
    x = int(0.1 * width)
    end = int(0.9 * width)
    while x <= end:
        if empty[x] >= need:
            a = x
            while x <= end and empty[x] >= need:
                x += 1
            if x - a >= 6:
                cands.append((a, x - 1, max(empty[a:x])))
        else:
            x += 1
    if not cands:
        return [("one", bands)]
    cands.sort(key=lambda c: (-c[2], -(c[1] - c[0])))
    for g0, g1, _cnt in cands[:5]:
        best = max(range(g0, g1 + 1), key=lambda i: empty[i])
        a = z = best
        top_cnt = empty[best]
        slack = max(1, int(0.06 * na))
        while a - 1 >= g0 and empty[a - 1] >= top_cnt - slack:
            a -= 1
        while z + 1 <= g1 and empty[z + 1] >= top_cnt - slack:
            z += 1
        if z - a < 5:
            continue
        c0, c1 = a + lo, z + lo
        clear = {}
        for y in active:
            row = occ[y]
            clear[y] = not any(row[x] for x in range(a + 1, z))
        # runs of strips in which the gutter is clear (blank strips between lines do not end a run)
        runs, cur = [], None
        for y in active:
            if clear[y]:
                if cur is None:
                    cur = [y, y]
                else:
                    cur[1] = y
            else:
                if cur is not None:
                    runs.append(cur)
                    cur = None
        if cur is not None:
            runs.append(cur)
        runs = [r for r in runs if (r[1] - r[0] + 1) * HY >= 30.0]
        if not runs:
            continue
        zones, used_any = [], False
        pending = []
        ri = 0

        def mid_y(b):
            return (b.top + b.bottom) / 2.0
        order = sorted(bands, key=lambda b: ((b.row.base if b.row is not None else b.top + 0.01), b.x0))
        taken = [None] * len(order)
        for k, b in enumerate(order):
            y = int((mid_y(b) - top) / HY)
            for r_i, r in enumerate(runs):
                if r[0] <= y <= r[1]:
                    crosses = (b.x0 < c0 - 0.5 and b.x1 > c1 + 0.5) and (b.row is None or any(w.x0 < c1 - 0.5 and w.x1 > c0 + 0.5 for w in b.row.words))
                    if not crosses:
                        taken[k] = r_i
                    break
        k = 0
        while k < len(order):
            if taken[k] is None:
                pending.append(order[k])
                k += 1
                continue
            j = k
            while j < len(order) and taken[j] == taken[k]:
                j += 1
            run = order[k:j]
            cols = _as_columns(run, c0, c1, XL, XR)
            if cols is None:
                pending.extend(run)
            else:
                if pending:
                    zones.append(("one", pending))
                    pending = []
                left, right, lbox, rbox = cols
                zones.append(("cols", [_zones(left, lbox[0], lbox[1], depth + 1), _zones(right, rbox[0], rbox[1], depth + 1)], [lbox, rbox]))
                used_any = True
            k = j
        if pending:
            zones.append(("one", pending))
        if used_any:
            return zones
    return [("one", bands)]


def _mode_near(vals, pick_max):
    """Among numbers that lie within 3 pt of the outermost one that at least two of them share, the place most of
    them stand at (to a tenth of a point); None when no place is shared."""
    if not vals:
        return None
    vs = sorted(vals, reverse=pick_max)
    n = len(vs)
    need = 1 if n < 4 else max(2, min(8, int(0.02 * n)))
    edge = vs[0]
    for i, v in enumerate(vs):
        if i + need - 1 < n and abs(vs[i + need - 1] - v) <= 1.0:
            edge = v
            break
    band = [v for v in vs if abs(v - edge) <= 3.0 and ((v <= edge + 0.6) if pick_max else (v >= edge - 0.6))]
    best, bc = None, 0
    for v in band:
        c = sum(1 for u in band if abs(u - v) <= 0.12)
        if c > bc or (c == bc and best is not None and ((v > best) if pick_max else (v < best))):
            best, bc = v, c
    if best is not None and bc >= max(2, int(0.2 * len(band))):
        grp = [u for u in band if abs(u - best) <= 0.12]
        return sum(grp) / len(grp)
    return edge


def _edges(rows):
    """Where a column's lines begin and end. Justified lines all end at the edge (a space written after a line's last
    word begins there too, even where a program stopped the letters a hair short of it); so the edge is the place
    most lines end at, not the farthest any one of them reaches. One line set out into the margin, or one thing wider
    than the column, does not move it."""
    if not rows:
        return None
    L = _mode_near([r.x0 for r in rows], False)
    ends = [(r.sx if (r.sx is not None and 0 <= r.sx - r.x1 <= 2.0) else r.x1) for r in rows]
    R = _mode_near(ends, True)
    if L is None or R is None or R - L < 20:
        L, R = min(r.x0 for r in rows), max(r.x1 for r in rows)
    return L, R


def _rebands(bands):
    """One column's bands with its rows grouped again by baseline (see _regroup_rows)."""
    rows = _regroup_rows([b.row for b in bands if b.row is not None])
    out = [b for b in bands if b.row is None]
    for r in rows:
        r.top, r.bottom = min(w.top for w in r.words), max(w.bottom for w in r.words)
        nb = _Band()
        nb.top, nb.bottom, nb.x0, nb.x1, nb.row, nb.block, nb.spans = r.top, r.bottom, r.x0, r.x1, r, None, _row_spans(r.words)
        out.append(nb)
    out.sort(key=lambda b: ((b.row.base if b.row is not None else b.top + 0.01), b.x0))
    return out


def _as_columns(run, c0, c1, XL, XR):
    """The bands of a run on each side of a gutter, if they are two columns of text; else None."""
    left, right = [], []
    nl = nr = both = 0
    chl = chr_ = 0
    for b in run:
        if b.row is None:
            if b.x1 <= (c0 + c1) / 2.0 + 2:
                left.append(b)
                nl += 1
                chl += 40
            elif b.x0 >= (c0 + c1) / 2.0 - 2:
                right.append(b)
                nr += 1
                chr_ += 40
            else:
                return None
            continue
        l, r = _split_row(b.row, c0, c1)
        if l is not None:
            nb = _Band()
            nb.top, nb.bottom, nb.x0, nb.x1, nb.row, nb.block, nb.spans = l.top, l.bottom, l.x0, l.x1, l, None, _row_spans(l.words)
            left.append(nb)
            nl += 1
            chl += sum(len(w.text()) for w in l.words) + len(l.words) - 1
        if r is not None:
            nb = _Band()
            nb.top, nb.bottom, nb.x0, nb.x1, nb.row, nb.block, nb.spans = r.top, r.bottom, r.x0, r.x1, r, None, _row_spans(r.words)
            right.append(nb)
            nr += 1
            chr_ += sum(len(w.text()) for w in r.words) + len(r.words) - 1
        if l is not None and r is not None:
            both += 1
    if nl < 3 or nr < 3:
        return None
    left, right = _rebands(left), _rebands(right)
    nl, nr = len(left), len(right)
    le = _edges([b.row for b in left if b.row is not None])
    re_ = _edges([b.row for b in right if b.row is not None])
    lx0, lx1 = le if le else (min(b.x0 for b in left), max(b.x1 for b in left))
    rx0, rx1 = re_ if re_ else (min(b.x0 for b in right), max(b.x1 for b in right))
    total = max(1.0, XR - XL)
    if lx1 - lx0 < 0.15 * total or rx1 - rx0 < 0.15 * total:
        return None
    al, ar = chl / float(nl), chr_ / float(nr)
    ok = (al >= 18 and ar >= 18 and min(nl, nr) >= 4) or (both <= 0.5 * min(nl, nr) and max(al, ar) >= 22 and min(nl, nr) >= 4 and min(al, ar) >= 6)
    if not ok:
        return None
    return left, right, (lx0, lx1), (rx0, rx1)


# -------------------------------------------------------------------------------------------- lines and paragraphs
def _is_marker(w, nxt):
    """Is this first word a list's bullet or number, standing clear of the words after it?"""
    t = w.text()
    if nxt is None:
        return False
    gap = nxt.x0 - w.x1
    if len(t) == 1 and t in _BULLETS:
        return gap > 0.2 * w.size
    if len(t) == 1 and t in _DASH_BULLETS:
        return gap > 0.75 * w.size
    if t == "o" and gap > 0.9 * w.size:
        return True
    if _NUM_MARK.match(t):
        # a number is a list's number when the words stand clear of it: by a gap no space would make
        return gap > 0.62 * w.size and (nxt.gapspace or gap > 1.1 * w.size or (nxt.spw > 0.5 and gap > 2.4 * nxt.spw))
    return False


def _make_line(row, doc):
    """A row of one column -> a line, cut into the parts a tab keeps apart."""
    ws = row.words
    ln = Ln()
    ln.words, ln.base, ln.size = ws, row.base, row.size
    ln.x0, ln.x1, ln.top, ln.bottom = ws[0].x0, max(w.x1 for w in ws), row.top, row.bottom
    ln.marker = len(ws) > 1 and _is_marker(ws[0], ws[1])
    segs = [[0, ws[0].x0, ws[0].x1]]
    for i in range(1, len(ws)):
        a, b = ws[i - 1], ws[i]
        gap = b.x0 - a.x1
        s = max(a.size, b.size)
        wide = gap > 2.0 * s or (i == 1 and ln.marker) or b.lead
        if not wide and b.gapspace:
            wide = gap > (0.9 * s if doc["spaces"] else 1.6 * s)
        if wide:
            segs.append([i, b.x0, b.x1])
        else:
            segs[-1][2] = max(segs[-1][2], b.x1)
    ln.segs = segs
    ln.dots = set(k for k, sg in enumerate(segs) if k and ws[sg[0]].lead)
    blk = Counter(w.blk for w in ws if w.blk is not None)
    ln.blk = blk.most_common(1)[0][0] if blk else None
    bolds = [q[1][2] for w in ws for q in w.parts]
    ln.allbold = bool(bolds) and all(bolds)
    first = ws[1] if ln.marker and len(ws) > 1 else ws[0]
    ln.startbold = first.parts[0][1][2]
    fam = Counter()
    for w in ws:
        for t, f in (q[:2] for q in w.parts):
            fam[f[0]] += len(t)
    ln.font = fam.most_common(1)[0][0] if fam else None
    ln.cen = ln.rgt = False
    return ln


def _mark_alignment(lines, L, R):
    width = max(1.0, R - L)
    for ln in lines:
        lo, ro = ln.x0 - L, R - ln.x1
        ln.cen = abs(lo - ro) <= max(2.5, 0.025 * width) and lo > max(8.0, 0.05 * width) and len(ln.segs) == 1
        ln.rgt = (not ln.cen) and ro <= 1.8 and lo > 0.3 * width and len(ln.segs) == 1


def _breaks(cur, ln, L, R, doc):
    """Does this line start a new paragraph (cur = the lines of the paragraph so far)?"""
    prev = cur[-1]
    gap = ln.base - prev.base
    if gap <= 0.3 * min(prev.size, ln.size):
        return True
    if prev.blk is not None and ln.blk is not None:
        if prev.blk[1] != "P?":
            return prev.blk != ln.blk
        # a paragraph proposed by OCR: taken, unless the page's own look says otherwise
        if prev.blk != ln.blk:
            return True
        if abs(prev.size - ln.size) > max(0.7, 0.1 * max(prev.size, ln.size)) or ln.marker or len(ln.segs) > 1 or len(prev.segs) > 1:
            return True
        if len(cur) >= 2 and gap > (cur[-1].base - cur[-2].base) * 1.3 + 0.6:
            return True
        # lines set in the middle (a title's three lines), a bold line before plain words, a line ended on purpose
        # (the next line's first word would have fitted on it with room to spare): each its own paragraph, as on
        # any other page. Tesseract takes such a block for one paragraph, and it came out as one, re-wrapped.
        if prev.cen or ln.cen or prev.rgt or ln.rgt:
            return True
        if prev.allbold and not ln.startbold:
            return True
        first = ln.words[0]
        if R - prev.x1 > (first.x1 - first.x0) + 0.6 * max(prev.size, ln.size) + 2.0:
            return True
        return False
    if (len(prev.segs) > 1 and not (prev.marker and len(prev.segs) == 2 and len(cur) == 1)) or len(ln.segs) > 1:
        return True
    if ln.marker:
        return True
    if abs(prev.size - ln.size) > 0.7:
        return True
    if prev.cen or ln.cen or prev.rgt or ln.rgt:
        return True
    if prev.allbold and not ln.startbold:
        return True
    size = max(prev.size, ln.size)
    if len(cur) >= 2:
        pitch = cur[-1].base - cur[-2].base
        if gap > pitch * 1.22 + 0.6 or gap < pitch * 0.8 - 0.6:
            return True
    else:
        usual = doc["pitch"].get((prev.font, prev.size))
        if usual:
            if gap > usual * 1.22 + 0.6:
                return True
        elif gap > 1.75 * size:
            return True
    # would the next line's first word have fitted on the line before? then that line was ended on purpose
    first = ln.words[0]
    free = R - prev.x1
    if free > (first.x1 - first.x0) + 0.33 * size + 1.0:
        return True
    body_x0 = prev.segs[1][1] if (prev.marker and len(prev.segs) > 1) else prev.x0
    tol = max(2.0, 0.25 * size)
    if len(cur) >= 2:
        rest = cur[1].x0
        if abs(ln.x0 - rest) > tol:
            return True
    else:
        if ln.x0 > body_x0 + 7.0 * size or ln.x0 < L - tol:
            return True
    return False


def _finish_para(lines, L, R, doc):
    p = Para()
    p.lines, p.L, p.R = lines, L, R
    p.top, p.bottom = lines[0].top, lines[-1].bottom
    p.first_base, p.last_base = lines[0].base, lines[-1].base
    p.level = p.rule = p.shade = None
    p.brk = p.keep = p.cont = p.colbrk = p.hard = False
    p.sect = None
    p.before = p.after = 0.0
    p.tabs = []
    sz, fam = Counter(), Counter()
    for ln in lines:
        for w in ln.words:
            for t, f in (q[:2] for q in w.parts):
                if f[5] == "":
                    sz[f[1]] += len(t)
                fam[f[0]] += len(t)
    p.size = sz.most_common(1)[0][0] if sz else lines[0].size
    p.font = fam.most_common(1)[0][0] if fam else None
    n = len(lines)
    width = max(1.0, R - L)
    p.pitch = _median([lines[i].base - lines[i - 1].base for i in range(1, n)]) if n >= 2 else None
    first = lines[0]
    body_x0 = first.segs[1][1] if (first.marker and len(first.segs) > 1) else None
    tol = max(1.6, 0.012 * width)
    p.right = 0.0
    if first.cen and all(l.cen or abs((l.x0 - L) - (R - l.x1)) <= max(3.0, 0.03 * width) for l in lines):
        p.align, p.left, p.first = "center", 0.0, 0.0
    elif n == 1 and first.rgt:
        p.align, p.left, p.first = "right", 0.0, 0.0
    elif n >= 2 and all(l.rgt or (R - l.x1) <= tol for l in lines) and max(l.x0 for l in lines) - min(l.x0 for l in lines) > 6 and not all(abs(l.x0 - lines[1].x0) <= tol for l in lines[1:]):
        p.align, p.left, p.first = "right", 0.0, 0.0
    else:
        rest = min(l.x0 for l in lines[1:]) if n >= 2 else (body_x0 if body_x0 is not None else first.x0)
        if body_x0 is not None and n >= 2 and abs(rest - body_x0) > max(3.0, 0.4 * p.size):
            rest = min(rest, body_x0) if rest < body_x0 else rest
        p.left = rest - L
        p.first = first.x0 - rest
        if abs(p.first) < 0.9:
            p.first = 0.0
        if abs(p.left) < 0.9:
            p.left = 0.0
        full = [l for l in lines[:-1] if (R - l.x1) <= tol + 0.6]
        if n >= 2 and len(full) == n - 1:
            p.align = "both"
        else:
            p.align = "left"
            mx = max(l.x1 for l in lines)
            if n >= 3 and R - mx > 0.18 * width:
                p.right = max(0.0, R - mx - 0.04 * width)
        over = max(l.x1 for l in lines) - R
        if over > 0.9:
            p.right = -over                                    # the lines run past the column's edge: so they do in Word
    # tabs: where each later part of a one-line paragraph starts (or ends, when it ends at the right edge)
    for li, ln in enumerate(lines):
        if len(ln.segs) > 1:
            for k in range(1, len(ln.segs)):
                i, x0, x1 = ln.segs[k]
                if k == 1 and ln.marker and li == 0:
                    if n >= 2 and abs(x0 - (L + p.left)) > 1.5:
                        p.tabs.append((x0 - L, "left", None))  # the words after the bullet do not start at the indent
                    continue                                   # (else they start at the hanging indent: no stop needed)
                last = k == len(ln.segs) - 1
                dot = "dot" if k in ln.dots else None
                if last and (R - x1 <= 2.0 or dot) and x0 - L > 0.5 * width:
                    t = ((R - L) if R - x1 <= 2.0 else (x1 - L), "right", dot)
                elif abs((x0 + x1) / 2.0 - (L + R) / 2.0) <= 2.5 and len(ln.segs) == 3 and k == 1:
                    t = ((L + R) / 2.0 - L, "center", dot)
                else:
                    t = (x0 - L, "left", dot)
                if not any(abs(t[0] - q[0]) < 1.0 for q in p.tabs):
                    p.tabs.append(t)
    p.spec = _line_spec(p, doc)
    p.open_end = (R - lines[-1].x1) <= max(tol + 1.0, 0.02 * width) and len(lines[-1].segs) == 1 and p.align in ("both", "left") and not _ends_sentence(lines[-1])
    p.open_start = False
    p.runs = None
    return p


_END = re.compile(r"[.!?:;\u2026][\"'\u201d\u2019)\]]*$")


def _ends_sentence(ln):
    return bool(_END.search(ln.words[-1].text()))


def _paragraphs(lines, L, R, doc):
    """Lines of one column, top to bottom -> paragraphs."""
    if not lines:
        return []
    _mark_alignment(lines, L, R)
    groups, cur = [], None
    for ln in lines:
        if cur is None or _breaks(cur, ln, L, R, doc):
            cur = [ln]
            groups.append(cur)
        else:
            cur.append(ln)
    return [_finish_para(g, L, R, doc) for g in groups]


# how far below the baseline each font's line reaches, as a share of its size (the rest of the line is above it):
# Word (and LibreOffice reading a Word file) put the baseline of a single-spaced line this far above the line's foot
_DESC = {"Calibri": 0.2686, "Calibri Light": 0.2686, "Cambria": 0.2222, "Times New Roman": 0.2163, "Arial": 0.2121, "Arial Narrow": 0.2121,
         "Courier New": 0.3003, "Georgia": 0.2192, "Verdana": 0.2100, "Tahoma": 0.2070, "Trebuchet MS": 0.2222, "Segoe UI": 0.2510,
         "Comic Sans MS": 0.2910, "Consolas": 0.2578, "Garamond": 0.2637, "Book Antiqua": 0.2832, "Palatino Linotype": 0.2832,
         "Century Gothic": 0.2197, "Century Schoolbook": 0.2202, "Bookman Old Style": 0.2500, "Candara": 0.2686, "Corbel": 0.2686,
         "Constantia": 0.2686, "Cambria Math": 0.2222}


def _desc(font):
    return _DESC.get(font, 0.22)


def _natural(font, size):
    """The height of one line of this font set single."""
    return _LINE.get(font, 1.17) * size


def _eff_pitch(p, doc):
    """How tall a line of this paragraph is: measured when it has two lines, else what is usual in this document for
    its font and size, else the font's own."""
    if p.pitch:
        return p.pitch
    u = doc["pitch"].get((p.font, p.size))
    if u and u < 3.2 * p.size:
        return max(u, 0.8 * _natural(p.font, p.size))
    return _natural(p.font, p.size)


# ----------------------------------------------------------------------------------------------- a page -> its flow
def _cell_take(tb, pg):
    """The page's words that lie in the table's cells leave the page's rows and wait in the cells."""
    taken = set()
    pads = []
    for ce in tb.cells:
        rows = []
        for r in pg.rows:
            ws = [w for w in r.words if ce.x0 - 0.5 <= (w.x0 + w.x1) / 2.0 <= ce.x1 + 0.5 and ce.top - 0.5 <= (w.top + w.bottom) / 2.0 <= ce.bottom + 0.5 and id(w) not in taken]
            if not ws:
                continue
            for w in ws:
                taken.add(id(w))
            n = _Row()
            n.base, n.words, n.sx = r.base, ws, None
            sz = Counter()
            for w in ws:
                sz[w.size] += w.x1 - w.x0
            n.size = sz.most_common(1)[0][0]
            n.top, n.bottom, n.x0, n.x1 = min(w.top for w in ws), max(w.bottom for w in ws), ws[0].x0, max(w.x1 for w in ws)
            rows.append(n)
        ce.items = rows                                # rows for now; paragraphs once the document's spacing is known
        if rows:
            pads.append(min(r.x0 for r in rows) - ce.x0)
    tb.pad = max(0.0, min(8.0, min(pads))) if pads else 5.4
    if taken:
        for r in pg.rows:
            r.words = [w for w in r.words if id(w) not in taken]
        pg.rows = [r for r in pg.rows if r.words]
        for r in pg.rows:
            r.x0, r.x1 = r.words[0].x0, max(w.x1 for w in r.words)
            r.words[0].space = False
        for ce in tb.cells:
            for r in ce.items:
                r.words[0].space = False


def _cell_pic(tb, f, pg):
    """A picture that lies in one cell of a table belongs to that cell (words drawn on it are part of the picture)."""
    cx, cy = (f[0] + f[2]) / 2.0, (f[1] + f[3]) / 2.0
    for ce in tb.cells:
        if ce.x0 - 1 <= cx <= ce.x1 + 1 and ce.top - 1 <= cy <= ce.bottom + 1:
            if f[0] < ce.x0 - 3 or f[2] > ce.x1 + 3 or f[1] < ce.top - 3 or f[3] > ce.bottom + 3:
                return                                   # it reaches over the cell's lines: the table's own drawing
            pc = Pic()
            pc.x0, pc.top, pc.x1, pc.bottom = max(f[0], ce.x0), max(f[1], ce.top), min(f[2], ce.x1), min(f[3], ce.bottom)
            pc.data = pc.ext = pc.anchor = pc.size = None
            pc.behind = False
            pc.kids = pc.flow = None
            pc.brk = pc.colbrk = False
            pc.sect = None
            pc.before = pc.after = pc.left = 0.0
            pc.page, pc.align = pg.n, "left"
            ce.pics.append(pc)
            keep = []
            for r in ce.items:
                r.words = [w for w in r.words if not (pc.x0 - 1 <= (w.x0 + w.x1) / 2.0 <= pc.x1 + 1 and pc.top - 1 <= (w.top + w.bottom) / 2.0 <= pc.bottom + 1)]
                if r.words:
                    r.x0, r.x1 = r.words[0].x0, max(w.x1 for w in r.words)
                    r.words[0].space = False
                    keep.append(r)
            ce.items = keep
            return


def _cell_build(tb, doc):
    """Each cell's rows -> its paragraphs, with the room above each."""
    pad = tb.pad
    for ce in tb.cells:
        rows = ce.items
        L = ce.x0 + (ce.padl if ce.padl is not None else pad)
        R = ce.x1 - (ce.padr if ce.padr is not None else pad)
        if R - L < 6:
            L, R = ce.x0, ce.x1
        lines = [_make_line(r, doc) for r in rows]
        for ln in lines:
            if ln.x1 > R:
                R = min(ce.x1, ln.x1)
        paras = _paragraphs(lines, L, R, doc)
        for pc in ce.pics:
            pc.L, pc.R = L, R
            mid = (pc.x0 + pc.x1) / 2.0
            if abs(mid - (L + R) / 2.0) <= max(2.0, 0.03 * (R - L)) and pc.x0 - L > 3:
                pc.align = "center"
            elif R - pc.x1 < 2 and pc.x0 - L > 6:
                pc.align = "right"
            else:
                pc.left = max(0.0, pc.x0 - L)
        if ce.pics:
            paras = sorted(paras + ce.pics, key=lambda it: it.top)
        if paras:
            t_free = paras[0].top - ce.top
            b_free = ce.bottom - paras[-1].bottom
            if abs(t_free - b_free) <= max(2.5, 0.12 * (ce.bottom - ce.top)) and t_free > 4.5:
                ce.valign = "center"
            elif b_free < t_free - 3 and b_free < 6:
                ce.valign = "bottom"
        _space(paras, ce.top if ce.valign == "top" else None, doc, limit=ce.bottom)
        ce.items = paras


def _link_words(pg):
    """Words that lie on a link's box carry the link."""
    for (x0, top, x1, bottom, uri) in pg.links:
        for r in pg.rows:
            if r.bottom < top - 1 or r.top > bottom + 1:
                continue
            for w in list(r.words):
                cx, cy = (w.x0 + w.x1) / 2.0, (w.top + w.bottom) / 2.0
                if x0 - 0.5 <= cx <= x1 + 0.5 and top - 1 <= cy <= bottom + 1:
                    _trim_to(r, w, x0, x1).link = uri


def _open_reader(src_path):
    """pypdf's reader of the PDF, opened (a PDF that needs no password to be read may still be encrypted)."""
    from pypdf import PdfReader
    rd = PdfReader(src_path)
    if rd.is_encrypted:
        try:
            rd.decrypt("")
        except Exception:
            pass
    return rd


def pdf_lock(src_path):
    """None when the PDF may be read and its contents taken out; else the words of the refusal.
    A PDF can be encrypted and still open without a password: its maker set limits on it (no changes, no printing,
    no copying). Such a file is converted like any other where its maker allows its contents to be copied out, and
    only there. (Until 6.4 every such file was called "password-protected" and sent to a tool that asks for a
    password the person was never given: there was no way on.)"""
    try:
        from pypdf import PdfReader
        from pypdf.constants import UserAccessPermissions as _P
        rd = PdfReader(src_path)
        if not rd.is_encrypted:
            return None
        try:
            ok = rd.decrypt("")
        except Exception:
            ok = 0
        if not ok:
            return LOCKED
        try:
            perm = rd.user_access_permissions
        except Exception:
            perm = None
        if perm is not None and not (perm & _P.EXTRACT):
            return NO_COPY
        return None
    except Exception:
        return None


def _lum(col):
    """How light a colour 'RRGGBB' is: 0 black ... 1 white (None is black)."""
    if not col:
        return 0.0
    try:
        r, g, b = int(col[0:2], 16), int(col[2:4], 16), int(col[4:6], 16)
    except (ValueError, TypeError):
        return 0.0
    return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0


_SLIDE_SHAPES = (4 / 3.0, 16 / 10.0, 16 / 9.0)


def _looks_designed(pg, cell_rows):
    """A page that is a picture with words on it, not paper with words: a slide, a poster, a certificate.
    Its drawing is kept whole, behind the words (see Pic.behind): a colour over the whole page; light words, which on
    white paper would not be seen; the shape of a slide."""
    area = pg.w * pg.h
    for f in pg.fills:
        if (f[2] - f[0]) * (f[3] - f[1]) >= 0.5 * area and _lum(f[4]) < 0.93:
            return True
    for im in pg.images:
        if (im[2] - im[0]) * (im[3] - im[1]) >= 0.6 * area:
            return True
    light = total = 0
    for r in pg.rows:
        for w in r.words:
            for part in w.parts:
                k = len(part[0])
                total += k
                if _lum(part[1][4]) >= 0.8:
                    light += k
    for rows, shade in cell_rows:
        dark = shade is not None and _lum(shade) < 0.62
        for r in rows:
            for w in r.words:
                for part in w.parts:
                    k = len(part[0])
                    total += k
                    if not dark and _lum(part[1][4]) >= 0.8:
                        light += k
    if light >= 3 and light >= 0.02 * total:
        return True
    a = pg.w / max(1.0, pg.h)
    if pg.w > pg.h and any(abs(a - k) <= 0.02 * k for k in _SLIDE_SHAPES):
        return True
    return False


def _analyse_page(page, n, doc):
    """Everything that needs the PDF's own page object: letters, drawing, tables, figures."""
    pg = _read_page(page, n)
    page_area = pg.w * pg.h
    big = [im for im in pg.images if (im[2] - im[0]) * (im[3] - im[1]) >= 0.6 * page_area]
    letters = sum(len(w.text()) for r in pg.rows for w in r.words)
    if pg.nchars and pg.bad > 0.25 * pg.nchars:
        pg.rows, pg.scan = [], True                     # letters with no meaning the PDF can tell us: read the page by OCR
        return pg
    if letters < 8 and (big or (not pg.rows and (pg.images or len(pg.curves) > 40))):
        pg.scan = True                                  # a picture of a page
        pg.rows = []
        return pg
    _link_words(pg)
    tables = _find_tables(page, pg)
    for tb in tables:
        _table_looks(tb, pg)
    pg.tables = tables
    _underlines(pg)
    for tb in tables:
        _cell_take(tb, pg)
    if doc.get("all_design") or _looks_designed(pg, [(ce.items, ce.shade) for tb in tables for ce in tb.cells]):
        # everything that is not words is drawn once, behind the words, as it stands in the PDF: no picture is cut out,
        # no line is made a border, and a table keeps its words only (its lines and tints are in that drawing)
        pg.design = True
        pg.figs, pg.rules = [], []
        for tb in tables:
            for ce in tb.cells:
                ce.borders, ce.shade = {}, None
        return pg
    # a picture that lies in one cell of a table is that cell's (two cells' pictures must not grow into one figure)
    if tables and pg.images:
        rest = []
        for im in pg.images:
            ce_tb = None
            for tb in tables:
                for ce in tb.cells:
                    if ce.x0 - 1.5 <= im[0] and im[2] <= ce.x1 + 1.5 and ce.top - 1.5 <= im[1] and im[3] <= ce.bottom + 1.5:
                        ce_tb = tb
                        break
                if ce_tb is not None:
                    break
            if ce_tb is not None and im[2] - im[0] >= 4 and im[3] - im[1] >= 4:
                _cell_pic(ce_tb, [im[0], im[1], im[2], im[3], True, 1, 0], pg)
            else:
                rest.append(im)
        pg.images = rest
    figs, full = _find_figures(pg)
    outside = []
    for f in figs:
        tb = next((t for t in tables if t.x0 - 2 <= f[0] and f[2] <= t.x1 + 2 and t.top - 2 <= f[1] and f[3] <= t.bottom + 2), None)
        if tb is None:
            outside.append(f)
        elif f[4] or (f[6] >= 2 and f[2] - f[0] >= 20 and f[3] - f[1] >= 20):
            _cell_pic(tb, f, pg)                        # a picture, or a drawing, in a cell of the table
    _settle_figures(pg, outside)
    return pg


def _norm_hf(text):
    return re.sub(r"\s+", " ", re.sub(r"\d+", "#", text)).strip().lower()


def _alike(a, b, at=0.85):
    """Two short texts read from pictures that are the same text, but for a letter or two read wrong."""
    import difflib
    if not a or not b or min(len(a), len(b)) < 0.7 * max(len(a), len(b)):
        return False
    return difflib.SequenceMatcher(None, a, b, autojunk=False).ratio() >= at


def _row_text(r):
    return " ".join(w.text() for w in r.words)


def _one_line(rows):
    """In a page's head or foot, words that stand on one baseline are one line, however far apart and whatever their
    size (a running title in the middle and a bigger page number at the edge). Left as two rows they would be two
    lines in Word, and the taller head would push the page's text down."""
    out = []
    for r in rows:
        if out and abs(out[-1].base - r.base) <= 0.6:
            a = out[-1]
            ws = sorted(a.words + r.words, key=lambda w: w.x0)
            m = _Row()
            m.words, m.sx = ws, None
            m.base = a.base if (a.x1 - a.x0) >= (r.x1 - r.x0) else r.base
            m.size = a.size if (a.x1 - a.x0) >= (r.x1 - r.x0) else r.size
            m.top, m.bottom, m.x0, m.x1 = min(a.top, r.top), max(a.bottom, r.bottom), ws[0].x0, max(w.x1 for w in ws)
            for i, w in enumerate(ws):
                if i and not w.space and w.x0 - ws[i - 1].x1 > 0.2 * max(w.size, ws[i - 1].size):
                    w.space = True
            ws[0].space = False
            out[-1] = m
        else:
            out.append(r)
    return out


def _split_head_foot(pages):
    """What repeats at the top and the foot of the pages (or is marked as such in a tagged PDF) leaves the body."""
    n = len(pages)
    counts = Counter()
    cand = {}

    def hkey(zone, pg, r):
        # where on the page: to 5 points on a page made on a computer; on a page read from a picture (a photograph
        # stands a little higher or lower on each page) the words alone, within the top or the foot of the page
        return (zone, _norm_hf(_row_text(r)), None if getattr(pg, "ocr", False) else int(round(r.base / 5.0)))
    for pg in pages:
        if pg.scan or pg.design:
            continue
        tops = [r for r in pg.rows if r.bottom <= 0.13 * pg.h][:4]
        bots = [r for r in pg.rows if r.top >= 0.87 * pg.h][-4:]
        cand[pg.n] = (tops, bots)
        seen = set()
        for zone, rs in (("t", tops), ("b", bots)):
            for r in rs:
                key = hkey(zone, pg, r)
                if key not in seen:
                    seen.add(key)
                    counts[key] += 1
    live = sum(1 for pg in pages if not (pg.scan or pg.design))
    need = max(2, int(math.ceil(0.4 * live)))
    # on pages read from pictures a letter can be read wrong on one page ("Rivers" as "Rlvers"): a row that is nearly
    # the same as a row at the top (the foot) of the other pages counts as the same
    ocr_rows = {"t": [], "b": []}
    for pg in pages:
        if getattr(pg, "ocr", False) and pg.n in cand:
            for zone, rs in (("t", cand[pg.n][0]), ("b", cand[pg.n][1])):
                for r in rs:
                    ocr_rows[zone].append((pg.n, _norm_hf(_row_text(r))))

    def seen_on(zone, pg, r):
        if not getattr(pg, "ocr", False):
            return counts[hkey(zone, pg, r)]
        t = _norm_hf(_row_text(r))
        if len(t) < 3:
            return counts[hkey(zone, pg, r)]
        on = set([pg.n])
        for (m, u) in ocr_rows[zone]:
            if m not in on and (u == t or (len(u) >= 3 and _alike(u, t))):
                on.add(m)
        return len(on)
    for pg in pages:
        if pg.scan or pg.design:
            continue
        tops, bots = cand.get(pg.n, ([], []))
        head, foot = [], []
        for r in pg.rows:
            art = all(w.art for w in r.words)
            if art and r.bottom <= 0.16 * pg.h:
                head.append(r)
            elif art and r.top >= 0.84 * pg.h:
                foot.append(r)
        if live >= 2:
            for r in tops:
                if r not in head and seen_on("t", pg, r) >= need:
                    head.append(r)
            for r in bots:
                if r not in foot and seen_on("b", pg, r) >= need:
                    foot.append(r)
        # the header is everything above its lowest repeating row; the footer everything below its highest
        if head:
            lim = max(r.base for r in head)
            head = [r for r in pg.rows if r.base <= lim + 0.5]
        if foot:
            lim = min(r.base for r in foot)
            foot = [r for r in pg.rows if r.base >= lim - 0.5]
        gone = set(id(r) for r in head) | set(id(r) for r in foot)
        pg.head, pg.foot = _one_line(sorted(head, key=lambda r: r.base)), _one_line(sorted(foot, key=lambda r: r.base))
        pg.rows = [r for r in pg.rows if id(r) not in gone]
    # the pictures and lines of the head and the foot: what stands beside or above the head's words, or comes again at
    # the same place on the pages (a logo, the line under a running title)

    def fkey(zone, f):
        return (zone, int(round(f[0] / 3.0)), int(round(f[1] / 3.0)), int(round((f[2] - f[0]) / 3.0)), int(round((f[3] - f[1]) / 3.0)))

    def rkey(zone, q):
        return (zone, int(round(q[0] / 3.0)), int(round(q[1] / 3.0)), int(round(q[2] / 2.0)))
    fc, rc = Counter(), Counter()
    for pg in pages:
        if pg.scan or pg.design:
            continue
        for f in pg.figs:
            if f[3] <= 0.2 * pg.h:
                fc[fkey("t", f)] += 1
            elif f[1] >= 0.8 * pg.h:
                fc[fkey("b", f)] += 1
        for q in pg.rules:
            if q[2] <= 0.2 * pg.h:
                rc[rkey("t", q)] += 1
            elif q[2] >= 0.8 * pg.h:
                rc[rkey("b", q)] += 1
    for pg in pages:
        if pg.scan or pg.design:
            continue
        hlim = max(r.bottom for r in pg.head) if pg.head else None
        flim = min(r.top for r in pg.foot) if pg.foot else None
        body_top = min([r.top for r in pg.rows] + [t.top for t in pg.tables] + [pg.h])
        body_bottom = max([r.bottom for r in pg.rows] + [t.bottom for t in pg.tables] + [0.0])
        keep = []
        for f in pg.figs:
            again_t = live >= 2 and fc[fkey("t", f)] >= need
            again_b = live >= 2 and fc[fkey("b", f)] >= need
            if f[3] <= 0.2 * pg.h and f[3] <= body_top + 2 and (again_t or (hlim is not None and f[1] < hlim)):
                pg.hfigs.append(f)
            elif f[1] >= 0.8 * pg.h and f[1] >= body_bottom - 2 and (again_b or (flim is not None and f[3] > flim)):
                pg.ffigs.append(f)
            else:
                keep.append(f)
        pg.figs = keep
        keep = []
        for q in pg.rules:
            if q[2] <= 0.2 * pg.h and q[2] < body_top and ((live >= 2 and rc[rkey("t", q)] >= need) or (hlim is not None and q[2] <= hlim + 14)):
                pg.hrules.append(q)
            elif q[2] >= 0.8 * pg.h and q[2] > body_bottom and ((live >= 2 and rc[rkey("b", q)] >= need) or (flim is not None and q[2] >= flim - 14)):
                pg.frules.append(q)
            else:
                keep.append(q)
        pg.rules = keep


def _flow_of(pg, doc):
    """The page's body in reading order: [("one", items, (L, R)) | ("cols", [items of each column], [(L, R), ...])]."""
    blocks = list(pg.tables)
    for f in pg.figs:
        pc = Pic()
        pc.x0, pc.top, pc.x1, pc.bottom = f
        pc.data = pc.ext = pc.anchor = pc.size = None
        pc.behind = False
        pc.kids = pc.flow = None
        pc.brk = pc.colbrk = False
        pc.sect = None
        pc.before = pc.after = pc.left = 0.0
        pc.page = pg.n
        pc.align = "left"
        blocks.append(pc)
    for s in pg.rules:
        ru = Rule()
        ru.x0, ru.x1, ru.top, ru.bottom, ru.width, ru.color = s[0], s[1], s[2], s[2], s[3], s[4]
        ru.brk = ru.colbrk = False
        ru.sect = None
        ru.before = ru.after = 0.0
        blocks.append(ru)
    bands = _bands_of(pg.rows, blocks)
    if not bands:
        return []
    XL, XR = min(b.x0 for b in bands), max(b.x1 for b in bands)
    return _zones(bands, XL, XR)


def _rivers(lines):
    """The strips no word of these lines touches, wide enough not to be the space between two words: the gaps
    between a table's columns."""
    x0 = min(l.x0 for l in lines)
    x1 = max(l.x1 for l in lines)
    if x1 - x0 < 20:
        return []
    n = int((x1 - x0) * 2) + 3
    cov = bytearray(n)
    sizes = []
    for l in lines:
        sizes.append(l.size)
        for w in l.words:
            a = max(0, int((w.x0 - x0) * 2))
            z = min(n - 1, int(math.ceil((w.x1 - x0) * 2)))
            for k in range(a, z + 1):
                cov[k] = 1
    need = int(max(4.0, 0.45 * _median(sizes, 10.0)) * 2)
    out = []
    k = 0
    while k < n:
        if not cov[k]:
            a = k
            while k < n and not cov[k]:
                k += 1
            if k - a >= need and a > 0 and k < n:
                out.append((x0 + a / 2.0, x0 + k / 2.0))
        else:
            k += 1
    return out


def _multi(ln):
    return len(ln.segs) >= 2 and not ln.marker


def _grow_table(ents, i0, L, R, doc):
    """From a line whose words stand in separate stretches: the run of lines (and rules between them) that share its
    column gaps. Returns (index after the run, Tbl or None). A plain label-and-value list stays what it is: lines
    with a tab."""
    n = len(ents)
    lines, rules = [], []
    j = i0
    seeds = conts = 0
    end = i0
    while j < n:
        kind, o = ents[j]
        if kind == "bk":
            if isinstance(o, Rule) and lines:
                rules.append((j, o))
                j += 1
                continue
            break
        ln = o
        if ln.marker:
            break
        if lines:
            prev = lines[-1][1]
            gap = ln.base - prev.base
            if gap > 3.3 * max(ln.size, prev.size) or gap <= 0:
                break
        if _multi(ln):
            if lines and not _rivers([l for _, l in lines] + [ln]):
                break
            lines.append((j, ln))
            seeds += 1
        else:
            if not lines:
                break
            cur = _rivers([l for _, l in lines])
            if not cur:
                break
            if any(ln.x0 < r1 - 0.5 and ln.x1 > r0 + 0.5 for r0, r1 in cur):
                break                                          # it crosses a gap between columns: ordinary text again
            if gap > 1.75 * max(ln.size, prev.size):
                break
            if ln.x0 < cur[0][0] and abs(ln.x0 - min(l.x0 for _, l in lines)) > 3.0:
                break
            lines.append((j, ln))
            conts += 1
        j += 1
        end = j
    while lines and not _multi(lines[-1][1]) and False:
        lines.pop()
    if seeds < 2:
        return i0, None
    last_line_idx = lines[-1][0]
    end = last_line_idx + 1
    rules = [(k, r) for k, r in rules if k < end or k == end]
    if rules and rules[-1][0] == end:
        end += 1                                               # the rule that closes the table
    rv = _rivers([l for _, l in lines])
    if not rv:
        return i0, None
    ncols = len(rv) + 1
    if ncols < 3 and conts == 0 and not rules:
        return i0, None
    tx0, tx1 = min(l.x0 for _, l in lines), max(l.x1 for _, l in lines)
    # the last column may use the room up to the column's edge (or a little past its widest words)
    far = max(tx1 + 2.0, min(R, tx1 + 12.0))
    xs = [tx0] + [(a + b) / 2.0 for a, b in rv] + [far]
    rules = [(k, r) for k, r in rules if min(r.x1, tx1) - max(r.x0, tx0) >= 0.5 * (tx1 - tx0) or (r.x1 - r.x0) >= 0.3 * (tx1 - tx0)]
    if rules:
        far = max(far, max(r.x1 for _, r in rules))            # ... or as far as the table's own rules reach
        xs[-1] = far
        if min(r.x0 for _, r in rules) < tx0 - 0.5:
            xs[0] = min(r.x0 for _, r in rules)
    tx0, tx1 = xs[0], xs[-1]
    # rows: a line with words in the first column, or with words in several, starts a row; a line that only goes on
    # under one of the other columns is the rest of a cell
    rows = []
    for k, ln in lines:
        first_col = any((w.x0 + w.x1) / 2.0 < xs[1] for w in ln.words)
        rule_between = rows and any(rows[-1][-1][0] < rk < k for rk, _ in rules)
        if not rows or _multi(ln) or first_col or rule_between:
            rows.append([(k, ln)])
        else:
            rows[-1].append((k, ln))
    # where each row begins and ends
    ys = []
    for ri, row in enumerate(rows):
        top = min(l.top for _, l in row)
        if ri == 0:
            ys.append(top - 1.0)
        else:
            prev_bottom = max(l.bottom for _, l in rows[ri - 1])
            between = [r for rk, r in rules if rows[ri - 1][-1][0] < rk < row[0][0]]
            ys.append(between[0].top if between else (prev_bottom + top) / 2.0)
    bottom_rule = [r for rk, r in rules if rk > lines[-1][0]]
    ys.append(bottom_rule[0].top if bottom_rule else max(l.bottom for _, l in rows[-1]) + 1.0)
    tb = Tbl()
    tb.x0, tb.x1, tb.top, tb.bottom, tb.xs, tb.ys, tb.cells = tx0, tx1, ys[0], ys[-1], xs, ys, []
    tb.borderless, tb.brk, tb.before, tb.after, tb.pad, tb.colbrk, tb.sect, tb.L, tb.R = True, False, 0.0, 0.0, 0.0, False, None, L, R
    tb.hdr, tb.brk_rows = 0, set()
    grid = {}
    for ri, row in enumerate(rows):
        for ci in range(ncols):
            ce = Cell()
            ce.r0, ce.r1, ce.c0, ce.c1, ce.items, ce.shade, ce.valign, ce.borders = ri, ri + 1, ci, ci + 1, [], None, "top", {}
            ce.x0, ce.x1, ce.top, ce.bottom = xs[ci], xs[ci + 1], ys[ri], ys[ri + 1]
            ce.padl = ce.padr = 0.0
            ce.pics = []
            grid[(ri, ci)] = ce
            tb.cells.append(ce)
        for k, ln in row:
            per = defaultdict(list)
            for w in ln.words:
                mid = (w.x0 + w.x1) / 2.0
                ci = 0
                while ci + 1 < ncols and mid >= xs[ci + 1]:
                    ci += 1
                per[ci].append(w)
            for ci, ws in per.items():
                r = _Row()
                r.base, r.words, r.sx = ln.base, ws, None
                sz = Counter()
                for w in ws:
                    sz[w.size] += w.x1 - w.x0
                r.size = sz.most_common(1)[0][0]
                r.top, r.bottom, r.x0, r.x1 = min(w.top for w in ws), max(w.bottom for w in ws), ws[0].x0, max(w.x1 for w in ws)
                ws[0].space = False
                grid[(ri, ci)].items.append(r)
    # each column's words keep their distance from the column's edges
    for ci in range(ncols):
        col = [grid[(ri, ci)] for ri in range(len(rows))]
        x0s = [r.x0 for ce in col for r in ce.items]
        x1s = [r.x1 for ce in col for r in ce.items]
        if not x0s:
            continue
        lo, hi = min(x0s), max(x1s)
        slack = 1.0
        ragged_left = max(x0s) - lo > 3.0
        flush_right = hi - min(x1s) <= 1.5
        mids = [(a + b) / 2.0 for a, b in zip(x0s, x1s)]
        if len(x0s) >= 2 and ragged_left and flush_right:
            padl, padr = slack, xs[ci + 1] - hi                # a column set flush right (figures): its right edge is kept
        elif len(x0s) >= 2 and ragged_left and max(mids) - min(mids) <= 2.0:
            side = min(lo - xs[ci], xs[ci + 1] - hi)          # a centred column: the same room on both sides
            padl, padr = max(slack, side - 1.0), max(slack, side - 1.0)
        else:
            padl, padr = lo - xs[ci], slack                    # flush left: the words may use the room up to the next column
        for ce in col:
            ce.padl, ce.padr = max(0.0, padl), max(0.0, padr)
    # rules are the borders of the rows they lie between
    for rk, r in rules:
        if rk < lines[0][0]:
            continue
        above = [ri for ri, row in enumerate(rows) if row[-1][0] < rk]
        if not above:
            continue
        ri = above[-1]
        for ci in range(ncols):
            ce = grid[(ri, ci)]
            if min(r.x1, ce.x1) - max(r.x0, ce.x0) >= 0.5 * (ce.x1 - ce.x0):
                ce.borders["bottom"] = (r.width, r.color)
    return end, tb


def _stream_tables(ents, L, R, doc):
    """Lines whose words stand in columns, with no ruled grid round them, become a table without borders (a rule
    drawn above, between or under its rows is a border of those rows)."""
    out = []
    i, n = 0, len(ents)
    while i < n:
        kind, o = ents[i]
        if kind == "ln" and _multi(o):
            j, tb = _grow_table(ents, i, L, R, doc)
            if tb is not None:
                # a rule that lies just above the first row is the table's top border
                if out and out[-1][0] == "bk" and isinstance(out[-1][1], Rule):
                    r = out[-1][1]
                    if min(r.x1, tb.x1) - max(r.x0, tb.x0) >= 0.5 * (tb.x1 - tb.x0) and 0 <= tb.ys[0] + 1.0 - r.top <= 2.6 * o.size:
                        out.pop()
                        tb.ys[0] = tb.top = r.top
                        for ce in tb.cells:
                            if ce.r0 == 0:
                                ce.top = r.top
                                if min(r.x1, ce.x1) - max(r.x0, ce.x0) >= 0.5 * (ce.x1 - ce.x0):
                                    ce.borders["top"] = (r.width, r.color)
                # the table reaches as far as its rules do
                _cell_build(tb, doc)
                out.append(["bk", tb])
                i = j
                continue
        out.append(ents[i])
        i += 1
    return out


_NOTE_MARK = re.compile(r"^(\d{1,3}|[*\u2217\u2020\u2021\u00a7\u00b6#]{1,3}|[a-z])[.)]?$")


def _take_notes(bands, L, R, doc, pg):
    """Footnotes: a short rule that starts at the column's left edge, low on the page, with nothing under it but small
    text. Returns (the bands above the rule, [Note])."""
    cut = None
    for i in range(len(bands) - 1, -1, -1):
        b = bands[i]
        if b.row is not None:
            continue
        k = b.block
        if not isinstance(k, Rule):
            break
        wdt = k.x1 - k.x0
        if abs(k.x0 - L) <= 4.5 and 28 <= wdt <= 0.55 * (R - L) + 2 and k.top > 0.3 * pg.h and i < len(bands) - 1:
            cut = i
        break
    if cut is None:
        return bands, []
    rows = [b.row for b in bands[cut + 1:]]
    if any(r is None for r in rows) or not rows:
        return bands, []
    body = doc["size"]
    small = max(r.size for r in rows) <= body - 0.4
    w0 = rows[0].words[0]
    first = w0.parts[0]
    marked = first[1][5] == "sup" or bool(_NOTE_MARK.match(first[0])) or bool(_NOTE_MARK.match(w0.text()))
    if not (small or marked):
        return bands, []
    lines = [_make_line(r, doc) for r in rows]
    paras = _paragraphs(lines, L, max(R, max(l.x1 for l in lines)), doc)
    notes = []
    for p in paras:
        w = p.lines[0].words[0]
        part = w.parts[0]
        mark = None
        if part[1][5] == "sup" and _NOTE_MARK.match(part[0]):
            mark, sup = part[0], True
        elif len(w.parts) == 1 and _NOTE_MARK.match(part[0]) and len(p.lines[0].words) > 1 and (part[0].isdigit() or not part[0].isalpha()):
            mark, sup = part[0], False
        if mark is None:
            if notes:
                notes[-1].items.append(p)                    # a note's second paragraph
                continue
            mark, sup = "", False                            # the end of a note begun on the page before
        n = Note()
        n.mark, n.sup, n.items, n.auto, n.id, n.page, n.found, n.order = mark.rstrip(".)"), sup, [p], False, 0, pg.n, False, 0
        notes.append(n)
    for n in notes:
        _space(n.items, None, doc)
        for it in n.items:
            it.before = it.before if it is not n.items[0] else 0.0
    return bands[:cut], notes


def _column_items(bands, L, R, doc, pg=None):
    items, buf = [], []
    if pg is not None:
        bands, notes = _take_notes(bands, L, R, doc, pg)
        if notes:
            doc.setdefault("notes", {}).setdefault(pg.n, []).extend(notes)

    ents = []
    for b in bands:
        if b.row is not None:
            ents.append(["ln", _make_line(b.row, doc)])
        else:
            b.block.L, b.block.R = L, R
            ents.append(["bk", b.block])
    ents = _stream_tables(ents, L, R, doc)

    def flush():
        if buf:
            items.extend(_paragraphs(list(buf), L, R, doc))
            del buf[:]
    for kind, k in ents:
        if kind == "ln":
            buf.append(k)
        else:
            flush()
            k.L, k.R = L, R
            if isinstance(k, Pic):
                mid, w = (k.x0 + k.x1) / 2.0, max(1.0, R - L)
                if abs(mid - (L + R) / 2.0) <= max(4.0, 0.03 * w) and k.x0 - L > 3:
                    k.align = "center"
                elif R - k.x1 <= 3 and k.x0 - L > 0.25 * w:
                    k.align = "right"
                else:
                    k.align, k.left = "left", k.x0 - L
            items.append(k)
    flush()
    return items


def _flatten_zones(zones, L, R, doc, pg=None):
    """zones -> [("one", items, (L, R)) or ("cols", [items of each column], [(L, R), ...])]. A column that is itself
    cut into columns from top to bottom makes three (or more) columns side by side; a column that is only partly cut
    is read in its order as one."""
    out = []
    for z in zones:
        if z[0] == "one":
            out.append(("one", _column_items(z[1], L, R, doc, pg), (L, R)))
            continue
        cols, boxes = [], []
        for sub, box in zip(z[1], z[2]):
            inner = _flatten_zones(sub, box[0], box[1], doc, None)
            if len(inner) == 1 and inner[0][0] == "cols":
                cols.extend(inner[0][1])
                boxes.extend(inner[0][2])
                continue
            its = []
            for f in inner:
                if f[0] == "one":
                    its.extend(f[1])
                else:
                    for c in f[1]:
                        its.extend(c)
            cols.append(its)
            boxes.append(box)
        out.append(("cols", cols, boxes))
    return out


# ------------------------------------------------------------------------------------------ the whole document
def _row_font(r):
    fam = Counter()
    for w in r.words:
        for t, f in (q[:2] for q in w.parts):
            fam[f[0]] += len(t)
    return fam.most_common(1)[0][0] if fam else None


def _doc_stats(pages, doc):
    """What is usual in this PDF: the body's font and size, the distance between lines, whether lines end in hyphens."""
    fonts, sizes = Counter(), Counter()
    gaps = defaultdict(Counter)
    ends = hyph = 0
    for pg in pages:
        prev = None
        right = max((r.x1 for r in pg.rows), default=0.0)
        left = min((r.x0 for r in pg.rows), default=0.0)
        for r in pg.rows:
            f = _row_font(r)
            n = sum(len(w.text()) for w in r.words)
            fonts[f] += n
            sizes[(f, r.size)] += n
            if prev is not None and prev[1] == f and prev[2] == r.size:
                g = r.base - prev[0]
                ov = min(r.x1, prev[4]) - max(r.x0, prev[3])
                ran_on = right - prev[4] <= 0.1 * max(1.0, right - left)       # the line above reached the right edge
                if ran_on and 0.6 * r.size < g < 3.4 * r.size and ov > 0.3 * min(r.x1 - r.x0, prev[4] - prev[3]):
                    gaps[(f, r.size)][round(g * 4) / 4.0] += 1
            prev = (r.base, f, r.size, r.x0, r.x1)
            last = r.words[-1].text()
            ends += 1
            if len(last) > 2 and last.endswith("-") and last[-2].isalpha():
                hyph += 1
    doc["font"] = next((f for f, _ in fonts.most_common() if f), None)
    body = [(k, v) for k, v in sizes.most_common() if k[0] == doc["font"]] or sizes.most_common()
    doc["size"] = body[0][0][1] if body else 11.0
    if body and body[0][0][0] and not doc["font"]:
        doc["font"] = body[0][0][0]
    pitch = {}
    for key, c in gaps.items():
        if sum(c.values()) < 2:
            continue
        # the usual distance between lines: the smallest one that comes more than once (bigger ones are the gaps
        # between paragraphs), a little smoothed
        best, bv = None, -1
        for g in sorted(c):
            v = c[g] + c.get(g - 0.25, 0) + c.get(g + 0.25, 0)
            if v > bv * 1.5:                       # the smallest of the common ones (bigger ones end a paragraph)
                best, bv = g, v
        near = [(g, k) for g, k in c.items() if abs(g - best) <= 0.3]
        pitch[key] = sum(g * k for g, k in near) / float(sum(k for g, k in near))
    doc["pitch"] = pitch
    doc["hyph"] = ends > 30 and hyph > 0.035 * ends
    full = squeezed = 0
    for pg in pages:
        if not pg.rows:
            continue
        right = max(r.x1 for r in pg.rows)
        for r in pg.rows:
            if right - r.x1 > 2.0 or len(r.words) < 6:
                continue
            ratios = [(w.x0 - r.words[i - 1].x1) / w.spw for i, w in enumerate(r.words) if i and w.spw > 0.5]
            if len(ratios) < 4:
                continue
            full += 1
            if _median(ratios) < 0.955:
                squeezed += 1
    # Word 2013 and later squeeze the spaces of some justified lines to take in one more word; older layouts never do.
    # Where the PDF's justified lines are never squeezed, only the older rule breaks the lines where the PDF does.
    doc["compat"] = 14 if (full >= 1 and squeezed == 0) else 15
    doc["just"] = (full, squeezed)


def _body_box(pages, doc):
    """The text area of each paper size: (left, right, top, bottom) in points from the page's corner."""
    groups = defaultdict(list)
    for pg in pages:
        groups[(int(round(pg.w)), int(round(pg.h)))].append(pg)
    boxes = {}
    for key, pgs in groups.items():
        x0s, x1s, tops, bots = [], [], [], []
        tails, edges, allrows = [], [], []
        for pg in pgs:
            if pg.scan:
                continue
            ext = [(r.x0, r.x1, r.top, r.bottom) for r in pg.rows]
            ext += [(t.x0 + (t.pad or 0), t.x1 - (t.pad or 0), t.top, t.bottom) for t in pg.tables]
            ext += [(f[0], f[2], f[1], f[3]) for f in pg.figs]
            if not ext:
                continue
            x0s.append(min(e[0] for e in ext))
            x1s.append(max(e[1] for e in ext))
            allrows.extend(pg.rows)
            for t in pg.tables:
                edges.append(t.x1)
            first = min(pg.rows, key=lambda r: r.base) if pg.rows else None
            top = min(e[2] for e in ext)
            if first is not None and first.top <= top + 1:
                f = _row_font(first) or doc["font"]
                top = first.base - (_natural(f, first.size) - _desc(f) * first.size)
            tops.append(top)
            pg.body_top = top
            last = max(pg.rows, key=lambda r: r.base) if pg.rows else None
            bot = max(e[3] for e in ext)
            if last is not None and last.bottom >= bot - 1:
                f = _row_font(last) or doc["font"]
                usual = doc["pitch"].get((_row_font(last), last.size))
                nat = _natural(f, last.size)
                bot = last.base + _desc(f) * last.size + (max(0.0, usual - nat) if usual and usual < 3.2 * last.size else 0.0)
            bots.append((bot, pg.n))
        w, h = float(key[0]), float(key[1])
        if not x0s:
            boxes[key] = (72.0 if w > 300 else 0.1 * w, w - (72.0 if w > 300 else 0.1 * w), 72.0 if h > 300 else 0.1 * h, h - (72.0 if h > 300 else 0.1 * h))
            continue
        ed = _edges(allrows)
        if ed:
            L, R = ed
        else:
            L, R = min(x0s), max(x1s)
        L = max(0.0, min(L, w * 0.45))
        R = min(w, max(R, w * 0.55))
        # the right edge: justified lines end at it. A space written after a line's last word starts at it too (some
        # programs stop the letters a hair short); a table as wide as the text shows it; and where no line is
        # justified, lines end short of the true edge, which is most often as far from the paper's side as the left one
        sym = w - L
        if 0 < sym - R < 0.16 * (R - L):
            R = sym
        live = [pg for pg in pgs if not pg.scan]
        if live and all(pg.design for pg in live) and sym > R:
            R = sym                                      # slides: the words' own boxes are not known; the lines are kept as they are
        T = max(0.0, min(tops))
        # how far down the text area reaches: the fullest page says (a last page is seldom full; a lone page says
        # nothing, so its foot margin is taken to be like its top one)
        lastn = max(pg.n for pg in pgs)
        inner = [b for b, n in bots if n != lastn] if len(bots) > 1 else []
        if inner:
            B = min(h, max(max(inner), max(b for b, n in bots) if max(b for b, n in bots) <= max(inner) + 40 else max(inner)))
            B = max(B, max(b for b, n in bots))
        else:
            B = max(max(b for b, n in bots), h - max(T, 36.0))
        B = min(h, B)
        boxes[key] = (L, R, T, B)
    doc["boxes"] = boxes


def _line_box(p, base, doc):
    """(top, bottom) of the line box Word gives a line of this paragraph whose baseline is at `base`:
    single or a multiple: the baseline sits the font's own height below the top, the extra room is under the line;
    exactly: the baseline is four fifths down; at least: the extra room is above."""
    font = p.font or doc["font"]
    nat = _natural(font, p.size)
    d = _desc(font) * p.size
    spec = p.spec
    if spec is None:
        return base - (nat - d), base + d
    rule, v = spec
    if rule == "auto":
        return base - (nat - d), base + d + (v / 240.0 - 1.0) * nat
    h = v / 20.0
    if rule == "exact":
        return base - 0.8 * h, base + 0.2 * h
    return base - (max(h, nat) - d), base + d


def _item_edges(it, doc):
    """(where the item's first line box begins, where its last ends) for spacing."""
    if isinstance(it, Para):
        if not it.lines:
            return it.top, it.bottom
        return _line_box(it, it.first_base, doc)[0], _line_box(it, it.last_base, doc)[1]
    if isinstance(it, Pic):
        if it.flow is not None:
            return it.flow.top, it.flow.bottom           # a picture that stands at its own place: a point of the flow holds it
        return it.top, it.bottom
    if isinstance(it, Rule):
        return it.top - it.width / 2.0, it.top + it.width / 2.0 + 1.0
    return it.top, it.bottom


def _space(items, top_ref, doc, limit=None):
    """The room above each item of one column (and under an item that a table follows)."""
    # a one-line paragraph was given the line spacing that is usual in this document; where what follows it stands
    # too close for that (a table's cell, a heading set solid), it is single-spaced after all
    for i, it in enumerate(items):
        if isinstance(it, Para) and it.lines and it.pitch is None and it.spec is not None:
            lim = limit if i == len(items) - 1 else _item_edges(items[i + 1], doc)[0] if not (isinstance(items[i + 1], Para) and not items[i + 1].lines) else None
            if lim is not None and _line_box(it, it.last_base, doc)[1] > lim + 0.4:
                it.spec = None
    prev_end = top_ref
    prev = None
    for it in items:
        a, z = _item_edges(it, doc)
        gap = max(0.0, a - prev_end) if prev_end is not None else 0.0
        gap = round(gap * 2) / 2.0
        if gap < 0.75:
            gap = 0.0
        if isinstance(it, Tbl):
            if prev is not None and not isinstance(prev, Tbl):
                prev.after = gap
            it.before = gap
        else:
            it.before = gap
        prev_end, prev = z, it


def _runs(p, doc):
    """A paragraph's words as runs: [text, fmt, underline, strike, link, field]."""
    out = []

    def add(t, f, ul, st, link, field=None):
        if field is None and out and out[-1][5] is None and out[-1][1] == f and out[-1][2] == ul and out[-1][3] == st and out[-1][4] == link:
            out[-1][0] += t
        else:
            out.append([t, f, ul, st, link, field])
    prev = None
    for li, ln in enumerate(p.lines):
        starts = set(s[0] for s in ln.segs[1:])
        for wi, w in enumerate(ln.words):
            sep = ""
            if wi in starts:
                sep = "\t"
            elif wi > 0:
                sep = " " if w.space else ""
            elif li > 0 and prev is not None and p.hard:
                pf = prev.parts[-1][1]
                out.append(["", (pf[0], pf[1], False, False, None, ""), False, False, None, "BRK:line"])
            elif li > 0 and prev is not None:
                last = prev.text()
                first = w.text()
                if len(last) > 2 and last.endswith("-") and last[-2].isalpha() and first[:1].islower():
                    if doc.get("hyph") and out and out[-1][5] is None and out[-1][0].endswith("-"):
                        out[-1][0] = out[-1][0][:-1]              # a word the typesetter broke: joined again
                    sep = ""
                else:
                    sep = " "
            if sep:
                pf = prev.parts[-1][1] if prev is not None else w.parts[0][1]
                if pf[5]:
                    pf = (pf[0], pf[1], pf[2], pf[3], pf[4], "")
                if sep == "\t":
                    add(sep, (pf[0], pf[1], False, False, None, ""), False, False, None)
                else:
                    add(sep, pf, prev is not None and prev.ul and w.ul, prev is not None and prev.st and w.st, prev.link if (prev is not None and prev.link == w.link) else None)
            if w.field:
                add(w.text(), w.parts[0][1], w.ul, w.st, None, w.field)
            else:
                for part in w.parts:
                    if len(part) > 2:
                        add(part[0], part[1], False, False, None, ("FN", part[2]))
                    else:
                        add(part[0], part[1], w.ul, w.st, w.link)
            prev = w
    return out


def _cell_text(ce):
    return " ".join(" ".join(w.text() for ln in it.lines for w in ln.words) for it in ce.items if isinstance(it, Para))


def _row_texts(tb, r):
    return [(_cell_text(ce), ce.c0, ce.c1) for ce in sorted(tb.cells, key=lambda c: c.c0) if ce.r0 == r]


def _join_tables(a, b, pin):
    """Is table b (first on its page) the rest of table a (last on the page before)? Then b's rows go into a: one
    table again, as it was written. Rows the PDF repeats at the top of the new page are the table's heading rows."""
    if len(a.xs) != len(b.xs) or a.borderless != b.borderless:
        return False
    if any(abs((x - a.xs[0]) - (y - b.xs[0])) > 2.0 for x, y in zip(a.xs, b.xs)) or abs(a.xs[0] - b.xs[0]) > 3.0:
        return False
    na, nb = len(a.ys) - 1, len(b.ys) - 1
    rep_rows = 0
    for k in range(min(3, na - 1, nb - 1)):
        ta, tb_ = _row_texts(a, k), _row_texts(b, k)
        if ta == tb_ and any(t for t, _, _ in ta):
            rep_rows = k + 1
        else:
            break
    if rep_rows:
        a.hdr = max(a.hdr, rep_rows)
        b.cells = [ce for ce in b.cells if ce.r0 >= rep_rows]
        if any(ce.r0 < rep_rows < ce.r1 for ce in b.cells):
            return False
        for ce in b.cells:
            ce.r0 -= rep_rows
            ce.r1 -= rep_rows
        b.ys = b.ys[rep_rows:]
        nb -= rep_rows
        if nb <= 0:
            return True
    off = a.ys[-1] - b.ys[0]
    a.ys = a.ys + [y + off for y in b.ys[1:]]
    for ce in b.cells:
        ce.r0 += na
        ce.r1 += na
        ce.top += off
        ce.bottom += off
        a.cells.append(ce)
    if pin:
        a.brk_rows.add(na)
    a.bottom = a.ys[-1]
    return True


def _can_join(p, q, doc):
    """Is q (first on the next page or column) the rest of paragraph p?"""
    if not (isinstance(p, Para) and isinstance(q, Para)):
        return False
    if not p.open_end or p.align not in ("both", "left") or q.align not in ("both", "left"):
        return False
    a, b = p.lines[-1], q.lines[0]
    if b.marker or len(b.segs) > 1 or b.cen or b.rgt:
        return False
    if a.font != b.font or abs(a.size - b.size) > 0.3 or a.allbold != b.allbold:
        return False
    if b.blk is not None and b.blk[1] not in ("P", "P?"):
        return False
    if len(q.lines) >= 2 and abs(q.first) > max(2.0, 0.25 * q.size):
        return False                                    # q opens with an indent of its own: a new paragraph
    if len(q.lines) == 1 and (q.left - p.left) > max(2.0, 0.25 * q.size) and len(p.lines) >= 2:
        return False
    return True


def _page_paras(pg):
    for f in pg.zones or []:
        cols = [f[1]] if f[0] == "one" else f[1]
        for items in cols:
            for it in items:
                if isinstance(it, Para):
                    yield it
                elif isinstance(it, Tbl):
                    for ce in it.cells:
                        for q in ce.items:
                            if isinstance(q, Para):
                                yield q


def _link_notes(pages, doc):
    """Each footnote is tied to the raised mark that calls it in the page's text. Where the marks run 1, 2, 3 through
    the document (or start again on every page) and every one is found, they are Word's own numbered footnotes;
    otherwise each note keeps the mark the PDF gave it, as it is written there."""
    allnotes = doc.get("notes") or {}
    if not allnotes:
        doc["notes_list"] = []
        return
    order = []
    for pg in pages:
        notes = allnotes.get(pg.n) or []
        if not notes:
            continue
        paras = list(_page_paras(pg))
        used = set()
        for n in notes:
            if not n.mark:
                continue
            for p in paras:
                hit = None
                for ln in p.lines:
                    for w in ln.words:
                        for i, part in enumerate(w.parts):
                            if len(part) == 2 and part[1][5] == "sup" and part[0].strip(".),") == n.mark and id(part) not in used:
                                hit = (w, i, part)
                                break
                        if hit:
                            break
                    if hit:
                        break
                if hit:
                    w, i, part = hit
                    used.add(id(part))
                    w.parts[i] = [part[0], part[1], n]
                    n.found = True
                    break
        for n in notes:
            if not n.found:
                # no mark in the text calls it: it hangs, unseen, on the end of the page's last paragraph
                last = None
                for p in paras:
                    if p.lines:
                        last = p
                if last is not None:
                    w = last.lines[-1].words[-1]
                    w.parts.append(["", w.parts[-1][1], n])
                else:
                    continue
            order.append(n)
    nums = [n.mark for n in order]
    digits = bool(order) and all(m.isdigit() for m in nums) and all(n.found for n in order)
    running = digits and [int(m) for m in nums] == list(range(1, len(nums) + 1))
    by_page = False
    if digits and not running:
        by_page = True
        seen = defaultdict(list)
        for n in order:
            seen[n.page].append(int(n.mark))
        by_page = all(v == list(range(1, len(v) + 1)) for v in seen.values())
    for i, n in enumerate(order):
        n.id = i + 1
        n.auto = running or by_page
    doc["notes_list"] = order
    doc["notes_each_page"] = by_page and not running


def _assemble(pages, doc):
    """All pages -> one list of items in reading order, with page breaks, column breaks and section starts marked.
    A paragraph that runs from one page (or column) into the next is one paragraph again."""
    body, sections = [], []
    cur = None                                          # the Word section being filled
    tail, tail_full = None, False                       # the last item put down; was its page (column) filled to the foot?
    forced = doc.get("sect_pages", set())
    pin = bool(doc.get("pin"))
    # where the text begins is a matter of each run of pages with the same page head: a first page that begins higher
    # must not set the top margin of the pages after it (their head would then push their text down the page)
    set_top = {}
    for pg in pages:
        bt = pg.body_top
        if bt is not None and not pg.scan:
            k2 = ((int(round(pg.w)), int(round(pg.h))), pg.hfset)
            set_top[k2] = min(set_top.get(k2, bt), bt)
    for pg in pages:
        key = (int(round(pg.w)), int(round(pg.h)))
        L, R, T, B = doc["boxes"][key]
        T = max(0.0, set_top.get((key, pg.hfset), T))
        if pg.hp:
            T = max(T, min(_item_edges(pg.hp[-1], doc)[1], T + 40.0))       # a head that reaches below it pushes the text down
        flows = [f for f in (pg.zones or []) if any(f[1] if f[0] == "cols" else [f[1]])]
        if not flows:
            flows = [("one", [_empty_para(L, R)], (L, R))]          # an empty page is still a page
        first_items = flows[0][1] if flows[0][0] == "one" else (flows[0][1][0] if flows[0][1] else [])
        if first_items and isinstance(first_items[0], Pic) and first_items[0].flow is not None:
            first_items[0].flow.top, first_items[0].flow.bottom = T, T + 1.0        # the holder of standing pictures: at this page's top
        page_start = True
        prev_end = T
        for fi, flow in enumerate(flows):
            if flow[0] == "one":
                cols, want_cols = [flow[1]], None
            else:
                cols = flow[1]
                want_cols = tuple((round(b[0], 1), round(b[1], 1)) for b in flow[2])
            if want_cols and cur is not None and cur["cols"] and len(cur["cols"]) == len(want_cols) and cur["paper"] == key \
                    and all(abs(a[0] - b[0]) <= 3.0 and abs(a[1] - b[1]) <= 3.0 for a, b in zip(cur["cols"], want_cols)):
                want_cols = cur["cols"]
            last_flow = fi == len(flows) - 1
            tops = [_item_edges(c[0], doc)[0] for c in cols if c]
            zone_top = min(tops) if tops else prev_end
            ends = []
            for ci, items in enumerate(cols):
                if not items:
                    continue
                items = list(items)
                _space(items, prev_end if ci == 0 else zone_top, doc)
                head = items[0]
                starts = ci == 0 and (cur is None or cur["paper"] != key or cur["cols"] != want_cols or cur["hf"] != pg.hfset or (page_start and pg.n in forced)
                                      or (page_start and cur.get("design") != pg.design))
                joined = False
                if starts:
                    cur = {"paper": key, "cols": want_cols, "page": pg.n, "hf": pg.hfset, "new_page": page_start and bool(body), "design": pg.design,
                           "top": T}
                    head.sect = cur
                    sections.append(cur)
                elif (page_start and ci == 0) or ci > 0:
                    if tail_full and _can_join(tail, head, doc):
                        joined = True
                    elif tail_full and page_start and isinstance(tail, Tbl) and isinstance(head, Tbl) and _join_tables(tail, head, pin):
                        items = items[1:]
                    elif not tail_full or pin:
                        if ci > 0:
                            head.colbrk = True
                        else:
                            head.brk = True
                if page_start and ci == 0 and not starts and (pg.design or (pg.n > 0 and pages[pg.n - 1].design)):
                    joined = False
                    head.brk = True
                col_end = _item_edges(items[-1], doc)[1] if not (isinstance(items[-1], Para) and not items[-1].lines) else prev_end
                if joined:
                    p = tail
                    if p.runs is None:
                        p.runs = _runs(p, doc)
                    qr = _runs(head, doc)
                    last = p.runs[-1][0] if p.runs else ""
                    firstw = qr[0][0] if qr else ""
                    broken = len(last) > 2 and last.endswith("-") and last[-2:-1].isalpha() and firstw[:1].islower()
                    if pin:
                        # the paragraph goes on, and the new page (column) starts at the very word it starts with in the PDF
                        if p.runs and qr and not broken:
                            p.runs.append([" ", p.runs[-1][1], False, False, None, None])
                        p.runs.append(["", (None, doc["size"], False, False, None, ""), False, False, None, "BRK:column" if ci > 0 else "BRK:page"])
                    elif broken:
                        if doc.get("hyph") and p.runs[-1][5] is None:
                            p.runs[-1][0] = last[:-1]
                    elif p.runs and qr:
                        p.runs.append([" ", p.runs[-1][1], False, False, None, None])
                    p.runs.extend(qr)
                    p.open_end = head.open_end
                    p.cont = True
                    items = items[1:]
                body.extend(items)
                if items:
                    tail = items[-1]
                ends.append(col_end)
                line_h = _eff_pitch(tail, doc) if isinstance(tail, Para) and tail.lines else doc["size"] * 1.2
                tail_full = last_flow and (B - col_end) <= max(2.4 * line_h, 0.055 * (B - T))
                page_start = False
            if ends:
                prev_end = max(ends)
        extra = (doc.get("page_notes") or {}).get(pg.n)
        if extra:
            # what the person wrote about this page but not on it (a slide's speaker notes): it follows on a page of
            # its own, in plain words
            first = True
            for line in extra:
                q = _empty_para(L, R)
                q.runs = [[line, (None, 12.0, False, False, None, ""), False, False, None, None]]
                q.brk = first
                q.after = 6.0
                first = False
                body.append(q)
                tail = q
            tail_full = False
    doc["sections"] = sections
    return body


def _empty_para(L, R):
    p = Para()
    p.lines, p.L, p.R = [], L, R
    p.align, p.left, p.first, p.right, p.before, p.after, p.pitch, p.tabs = "left", 0.0, 0.0, 0.0, 0.0, 0.0, None, []
    p.level = p.rule = p.shade = p.sect = p.spec = None
    p.brk = p.keep = p.cont = p.colbrk = p.open_end = p.open_start = p.hard = False
    p.runs = []
    p.size, p.font = 11.0, None
    p.top = p.bottom = p.first_base = p.last_base = 0.0
    return p


# ------------------------------------------------------------------------------------------------- writing Word
def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn("w:" + k), str(v))
    return e


def _tw(pt):
    return int(round(pt * 20))


_SNAP = ((1.0, 240), (1.08, 259), (1.15, 276), (1.5, 360), (2.0, 480))


def _line_spec(p, doc):
    """How Word should space this paragraph's lines: None (single), or (rule, value)."""
    font = p.font or doc["font"]
    pe = p.pitch or doc["pitch"].get((p.font, p.size))
    if not pe:
        return None
    nat = _natural(font, p.size)
    if font in _LINE and not (doc.get("pin") and p.pitch):
        ratio = pe / nat
        if ratio > 3.2:
            return None
        for r, tw in _SNAP:
            if abs(ratio - r) <= 0.012:
                return None if tw == 240 else ("auto", tw)
        if ratio >= 1.0:
            return ("auto", int(round(240 * ratio)))
        return ("exact", _tw(pe))
    return ("exact", _tw(pe))


def _r_el(text, fmt, ul, st, dflt, field=None):
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    fam, size, bold, ital, col, va = fmt
    if fam:
        dflt.setdefault("used", set()).add(fam)
    if fam and fam != dflt["font"]:
        rpr.append(_el("w:rFonts", ascii=fam, hAnsi=fam, eastAsia=fam, cs=fam))
    if bold:
        rpr.append(_el("w:b"))
        rpr.append(_el("w:bCs"))
    if ital:
        rpr.append(_el("w:i"))
        rpr.append(_el("w:iCs"))
    if st:
        rpr.append(_el("w:strike"))
    if col:
        rpr.append(_el("w:color", val=col))
    pct = dflt["wscale"].get((fam or dflt["font"], bool(bold), bool(ital))) if dflt.get("wscale") else None
    if pct:
        rpr.append(_el("w:w", val=pct))                  # a stand-in font's letters, as wide as the PDF's own
    if size and abs(size - dflt["size"]) > 0.01:
        rpr.append(_el("w:sz", val=int(round(size * 2))))
        rpr.append(_el("w:szCs", val=int(round(size * 2))))
    if ul:
        rpr.append(_el("w:u", val="single"))
    if va:
        rpr.append(_el("w:vertAlign", val="superscript" if va == "sup" else "subscript"))
    if len(rpr):
        r.append(rpr)
    if field in ("BRK:page", "BRK:column"):
        r.append(_el("w:br", type=field[4:]))
        return [r]
    if field == "BRK:line":
        r.append(OxmlElement("w:br"))                    # the line ends here, as in the PDF (a designed page)
        return [r]
    if isinstance(field, tuple) and field[0] == "FN":
        note = field[1]
        ref = _el("w:footnoteReference", id=note.id)
        if not note.auto:
            ref.set(qn("w:customMarkFollows"), "1")
        r.append(ref)
        if not note.auto and text:
            t = OxmlElement("w:t")
            t.set(qn("xml:space"), "preserve")
            t.text = text
            r.append(t)
        return [r]
    if field == "FNREF":
        r.append(OxmlElement("w:footnoteRef"))
        return [r]
    if field:
        import copy
        out = []
        for kind in ("begin", None, "separate", "text", "end"):
            rr = OxmlElement("w:r")
            if len(rpr):
                rr.append(copy.deepcopy(rpr))
            if kind in ("begin", "separate", "end"):
                rr.append(_el("w:fldChar", fldCharType=kind))
            elif kind is None:
                it = OxmlElement("w:instrText")
                it.set(qn("xml:space"), "preserve")
                it.text = " %s " % field
                rr.append(it)
            else:
                t = OxmlElement("w:t")
                t.text = text
                rr.append(t)
            out.append(rr)
        return out
    for piece in re.split("(\t)", text):
        if piece == "\t":
            r.append(OxmlElement("w:tab"))
        elif piece:
            t = OxmlElement("w:t")
            t.set(qn("xml:space"), "preserve")
            t.text = piece
            r.append(t)
    return [r]


def _p_el(p, doc, part, extra_before=0.0):
    """A paragraph as Word's XML."""
    el = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    if p.brk:
        ppr.append(_el("w:pageBreakBefore"))
    if p.tabs:
        tabs = OxmlElement("w:tabs")
        seen = set()
        for pos, kind, leader in sorted(p.tabs, key=lambda t: t[0]):
            tw = _tw(pos)
            if tw in seen or tw <= 0:
                continue
            seen.add(tw)
            te = _el("w:tab", val=kind, pos=tw)
            if leader:
                te.set(qn("w:leader"), leader)
            tabs.append(te)
        if len(tabs):
            ppr.append(tabs)
    sp = {}
    before = (p.before or 0.0) + extra_before
    if before >= 0.75:
        sp["before"] = _tw(before)
    if p.after and p.after >= 0.75:
        sp["after"] = _tw(p.after)
    spec = p.spec if p.lines else None
    if spec:
        sp["line"], sp["lineRule"] = spec[1], spec[0]
    if sp:
        ppr.append(_el("w:spacing", **sp))
    ind = {}
    if abs(p.left) >= 0.9:
        ind["left"] = _tw(p.left)
    if abs(p.right) >= 0.9:
        ind["right"] = _tw(p.right)
    if p.first >= 0.9:
        ind["firstLine"] = _tw(p.first)
    elif p.first <= -0.9:
        ind["hanging"] = _tw(-p.first)
        if "left" not in ind and -p.first > 0:
            ind["left"] = _tw(-p.first) if p.left >= 0 else 0
    if ind:
        ppr.append(_el("w:ind", **ind))
    if p.align in ("center", "right", "both"):
        ppr.append(_el("w:jc", val=p.align))
    if p.level is not None:
        ppr.append(_el("w:outlineLvl", val=int(p.level)))
    if len(ppr):
        el.append(ppr)
    if p.colbrk:
        r = OxmlElement("w:r")
        r.append(_el("w:br", type="column"))
        el.append(r)
    runs = p.runs if p.runs is not None else _runs(p, doc)
    dflt = doc["dflt"]
    i = 0
    while i < len(runs):
        text, fmt, ul, st, link, field = runs[i]
        if link:
            j = i
            hl = OxmlElement("w:hyperlink")
            try:
                rid = part.relate_to(link, _RT.HYPERLINK, is_external=True)
                hl.set(qn("r:id"), rid)
                hl.set(qn("w:history"), "1")
            except Exception:
                hl = None
            while j < len(runs) and runs[j][4] == link:
                for r in _r_el(runs[j][0], runs[j][1], runs[j][2], runs[j][3], dflt, runs[j][5]):
                    (hl if hl is not None else el).append(r)
                j += 1
            if hl is not None:
                el.append(hl)
            i = j
            continue
        for r in _r_el(text, fmt, ul, st, dflt, field):
            el.append(r)
        i += 1
    return el


def _tiny_p(before=0.0, brk=False, colbrk=False):
    """The smallest paragraph Word allows (it needs one after a table, and to carry a section's end)."""
    el = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    if brk:
        ppr.append(_el("w:pageBreakBefore"))
    sp = {"before": _tw(max(0.0, before - 1.0)) if before > 1.5 else 0, "after": 0, "line": 20, "lineRule": "exact"}
    ppr.append(_el("w:spacing", **sp))
    rpr = OxmlElement("w:rPr")
    rpr.append(_el("w:sz", val=2))
    rpr.append(_el("w:szCs", val=2))
    ppr.append(rpr)
    el.append(ppr)
    if colbrk:
        r = OxmlElement("w:r")
        r.append(_el("w:br", type="column"))
        el.append(r)
    return el


def _rule_el(ru, doc):
    el = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    if ru.brk:
        ppr.append(_el("w:pageBreakBefore"))
    bd = OxmlElement("w:pBdr")
    bd.append(_el("w:top", val="single", sz=max(2, min(96, int(round(ru.width * 8)))), space=0, color=ru.color or "000000"))
    ppr.append(bd)
    ppr.append(_el("w:spacing", before=_tw(ru.before) if ru.before >= 0.75 else 0, after=0, line=20, lineRule="exact"))
    ind = {}
    if ru.x0 - ru.L >= 1.5:
        ind["left"] = _tw(ru.x0 - ru.L)
    if ru.R - ru.x1 >= 1.5:
        ind["right"] = _tw(ru.R - ru.x1)
    if ind:
        ppr.append(_el("w:ind", **ind))
    rpr = OxmlElement("w:rPr")
    rpr.append(_el("w:sz", val=2))
    rpr.append(_el("w:szCs", val=2))
    ppr.append(rpr)
    el.append(ppr)
    if ru.colbrk:
        r = OxmlElement("w:r")
        r.append(_el("w:br", type="column"))
        el.append(r)
    return el


def _pic_el(pc, doc, holder):
    """A picture in a paragraph of its own. holder: a python-docx container whose part keeps the picture."""
    from docx.text.paragraph import Paragraph
    el = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    if pc.brk:
        ppr.append(_el("w:pageBreakBefore"))
    sp = {}
    if pc.before >= 0.75:
        sp["before"] = _tw(pc.before)
    if pc.after >= 0.75:
        sp["after"] = _tw(pc.after)
    if sp:
        ppr.append(_el("w:spacing", **sp))
    if pc.align == "left" and abs(pc.left) >= 1.5:
        ppr.append(_el("w:ind", left=_tw(pc.left)))
    if pc.align in ("center", "right"):
        ppr.append(_el("w:jc", val=pc.align))
    # the line holds the picture and nothing else: its own (unseen) end mark is set as small as can be, so that the
    # line is exactly as tall as the picture, in Word and wherever else the file is opened
    rpr = OxmlElement("w:rPr")
    rpr.append(_el("w:sz", val=2))
    rpr.append(_el("w:szCs", val=2))
    ppr.append(rpr)
    el.append(ppr)
    if pc.colbrk:
        r = OxmlElement("w:r")
        r.append(_el("w:br", type="column"))
        el.append(r)
    if pc.data:
        par = Paragraph(el, holder)
        run = par.add_run()
        try:
            run.add_picture(io.BytesIO(pc.data), width=Emu(int(round((pc.x1 - pc.x0) * 12700))), height=Emu(int(round((pc.bottom - pc.top) * 12700))))
            inl = run._r.find(".//" + qn("wp:inline"))
            if inl is not None:                         # no room round the picture: said outright, or LibreOffice gives it 9 points
                for k in ("distT", "distB", "distL", "distR"):
                    inl.set(k, "0")
            rp = OxmlElement("w:rPr")
            rp.append(_el("w:sz", val=2))
            rp.append(_el("w:szCs", val=2))
            run._r.insert(0, rp)
        except Exception:
            el.remove(run._r)
    return el


def _anchor_run(pc, holder, behind=False, order=1):
    # (a page's drawing has the page's size; a logo its own)
    """A run with a picture that stands at its own place on the page, whatever the words do: a logo in a page's head,
    the drawing behind the words of a designed page. pc.anchor = (x, y) from the paper's corner, in points."""
    from docx.text.paragraph import Paragraph
    tmp = OxmlElement("w:p")
    run = Paragraph(tmp, holder).add_run()
    wd, ht = pc.size if pc.size else (pc.x1 - pc.x0, pc.bottom - pc.top)
    run.add_picture(io.BytesIO(pc.data), width=Emu(int(round(wd * 12700))), height=Emu(int(round(ht * 12700))))
    inline = run._r.find(".//" + qn("wp:inline"))
    anchor = OxmlElement("wp:anchor")
    for k, v in (("distT", "0"), ("distB", "0"), ("distL", "0"), ("distR", "0"), ("simplePos", "0"), ("relativeHeight", str(int(order))),
                 ("behindDoc", "1" if behind else "0"), ("locked", "0"), ("layoutInCell", "1"), ("allowOverlap", "1")):
        anchor.set(k, v)
    sp = OxmlElement("wp:simplePos")
    sp.set("x", "0")
    sp.set("y", "0")
    anchor.append(sp)
    for tag, v in (("wp:positionH", pc.anchor[0]), ("wp:positionV", pc.anchor[1])):
        e = OxmlElement(tag)
        e.set("relativeFrom", "page")
        o = OxmlElement("wp:posOffset")
        o.text = str(int(round(v * 12700)))
        e.append(o)
        anchor.append(e)
    kids = dict((ch.tag, ch) for ch in list(inline))
    anchor.append(kids[qn("wp:extent")])
    ee = kids.get(qn("wp:effectExtent"))
    if ee is None:
        ee = OxmlElement("wp:effectExtent")
        for k in ("l", "t", "r", "b"):
            ee.set(k, "0")
    anchor.append(ee)
    anchor.append(OxmlElement("wp:wrapNone"))
    for t in ("wp:docPr", "wp:cNvGraphicFramePr", "a:graphic"):
        ch = kids.get(qn(t))
        if ch is not None:
            anchor.append(ch)
    parent = inline.getparent()
    parent.replace(inline, anchor)
    tmp.remove(run._r)
    return run._r


def _border(tag, b):
    if not b:
        return _el(tag, val="nil")
    return _el(tag, val="single", sz=max(2, min(96, int(round(b[0] * 8)))), space=0, color=b[1] or "000000")


def _tbl_el(tb, doc, part, holder):
    nrows, ncols = len(tb.ys) - 1, len(tb.xs) - 1
    owner = {}
    for ce in tb.cells:
        for r in range(ce.r0, ce.r1):
            for c in range(ce.c0, ce.c1):
                owner[(r, c)] = ce
    tbl = OxmlElement("w:tbl")
    pr = OxmlElement("w:tblPr")
    width = tb.xs[-1] - tb.xs[0]
    pr.append(_el("w:tblW", w=_tw(width), type="dxa"))
    span = max(1.0, tb.R - tb.L)
    mid = (tb.x0 + tb.x1) / 2.0
    if abs(mid - (tb.L + tb.R) / 2.0) <= max(3.0, 0.02 * span) and tb.x0 - tb.L > 4:
        pr.append(_el("w:jc", val="center"))
    else:
        # before Word 2013 a table stood left of where it was put by its cells' inner margin
        ind = tb.x0 - tb.L + ((tb.pad if tb.pad is not None else 5.4) if doc.get("compat", 15) < 15 else 0.0)
        if abs(ind) > 0.4:
            pr.append(_el("w:tblInd", w=_tw(ind), type="dxa"))
    pr.append(_el("w:tblLayout", type="fixed"))
    mar = OxmlElement("w:tblCellMar")
    pad = tb.pad if tb.pad is not None else 5.4
    for side, v in (("top", 0), ("left", pad), ("bottom", 0), ("right", pad)):
        mar.append(_el("w:" + side, w=_tw(v), type="dxa"))
    pr.append(mar)
    tbl.append(pr)
    grid = OxmlElement("w:tblGrid")
    for c in range(ncols):
        grid.append(_el("w:gridCol", w=_tw(tb.xs[c + 1] - tb.xs[c])))
    tbl.append(grid)
    first_p = None
    for r in range(nrows):
        tr = OxmlElement("w:tr")
        trpr = OxmlElement("w:trPr")
        trpr.append(_el("w:trHeight", val=_tw(tb.ys[r + 1] - tb.ys[r]), hRule="atLeast"))
        if r < tb.hdr:
            trpr.append(_el("w:tblHeader"))
        tr.append(trpr)
        row_first = None
        c = 0
        while c < ncols:
            ce = owner.get((r, c))
            tc = OxmlElement("w:tc")
            tcpr = OxmlElement("w:tcPr")
            if ce is None:
                tcpr.append(_el("w:tcW", w=_tw(tb.xs[c + 1] - tb.xs[c]), type="dxa"))
                tc.append(tcpr)
                tc.append(_tiny_p())
                tr.append(tc)
                c += 1
                continue
            span_c = ce.c1 - ce.c0
            tcpr.append(_el("w:tcW", w=_tw(tb.xs[ce.c1] - tb.xs[ce.c0]), type="dxa"))
            if span_c > 1:
                tcpr.append(_el("w:gridSpan", val=span_c))
            if ce.r1 - ce.r0 > 1:
                tcpr.append(_el("w:vMerge", val="restart") if r == ce.r0 else _el("w:vMerge"))
            bd = OxmlElement("w:tcBorders")
            for side in ("top", "left", "bottom", "right"):
                bd.append(_border("w:" + side, ce.borders.get(side)))
            tcpr.append(bd)
            if ce.shade:
                tcpr.append(_el("w:shd", val="clear", color="auto", fill=ce.shade))
            if ce.padl is not None or ce.padr is not None:
                tm = OxmlElement("w:tcMar")
                tm.append(_el("w:left", w=_tw(max(0.0, ce.padl or 0.0)), type="dxa"))
                tm.append(_el("w:right", w=_tw(max(0.0, ce.padr or 0.0)), type="dxa"))
                tcpr.append(tm)
            if ce.valign in ("center", "bottom"):
                tcpr.append(_el("w:vAlign", val=ce.valign))
            tc.append(tcpr)
            wrote = False
            if r == ce.r0:
                for it in ce.items:
                    if isinstance(it, Para):
                        pe = _p_el(it, doc, part)
                    elif isinstance(it, Pic):
                        pe = _pic_el(it, doc, holder)
                    else:
                        continue
                    if first_p is None:
                        first_p = pe
                    if row_first is None:
                        row_first = pe
                    tc.append(pe)
                    wrote = True
            if not wrote:
                e = _tiny_p()
                if first_p is None:
                    first_p = e
                if row_first is None:
                    row_first = e
                tc.append(e)
            tr.append(tc)
            c += span_c
        if r in tb.brk_rows and row_first is not None:
            ppr = row_first.find(qn("w:pPr"))
            if ppr is None:
                ppr = OxmlElement("w:pPr")
                row_first.insert(0, ppr)
            ppr.insert(0, _el("w:pageBreakBefore"))
        tbl.append(tr)
    if tb.brk and first_p is not None:
        ppr = first_p.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            first_p.insert(0, ppr)
        ppr.insert(0, _el("w:pageBreakBefore"))
    return tbl


# ------------------------------------------------------------------------------------------------- scanned pages
_PAPERS = ((595.28, 841.89), (612.0, 792.0), (612.0, 1008.0), (419.53, 595.28), (841.89, 1190.55))
_ASCENDERS = set("bdfhklt")


def _paper_for(w, h):
    """A scanned page's size in the PDF is whatever the scanner or phone wrote (often far bigger than paper). The Word
    page is real paper: the page scaled so its width is A4's or Letter's, whichever its shape is nearer to.
    Returns (scale, page width, page height)."""
    for a, b in _PAPERS:
        for pw, ph in ((a, b), (b, a)):
            if abs(w - pw) <= 4 and abs(h - ph) <= 4:
                return 1.0, w, h
    land = w > h
    lo, hi = (h, w) if land else (w, h)
    ratio = hi / max(1.0, lo)
    paper = _PAPERS[1] if abs(ratio - 1.2941) < abs(ratio - 1.4142) else _PAPERS[0]
    sc = paper[0] / lo
    long_side = hi * sc
    if abs(long_side - paper[1]) <= 0.03 * paper[1]:
        long_side = max(long_side, paper[1])
    long_side = min(long_side, 1580.0)
    return sc, (long_side if land else paper[0]), (paper[0] if land else long_side)


def _scan_figures(im, boxes, sure=None):
    """Pictures on a scanned page: the places still dark when every read word is painted over, big enough and not
    a table's ruled lines. im: the page, grey. boxes: the boxes (pixels) of the words read with some trust: "words"
    read with none are mostly a picture's own bars and lines, and painting them over painted the picture away.
    sure: the boxes of the words that are surely words (trusted, three letters or more): a place with many of them
    in it is writing on a tint, not a picture. Returns boxes in pixels."""
    from PIL import ImageDraw
    w, h = im.size
    work = im.copy()
    d = ImageDraw.Draw(work)
    cell = max(6, int(round(w / 170.0)))
    for (x0, y0, x1, y1) in boxes:
        d.rectangle((x0 - 3, y0 - 3, x1 + 3, y1 + 3), fill=255)
    gw, gh = max(1, w // cell), max(1, h // cell)
    small = work.resize((gw, gh), Image.BOX)
    px = list(small.getdata())
    srt = sorted(px)
    paper = srt[int(0.6 * len(srt))]
    for (x0, y0, x1, y1) in boxes:                            # painted-over words must read as paper, not as white
        pass
    dark = bytearray(gw * gh)
    for i, v in enumerate(px):
        if v < paper - 26 and v < 215:
            dark[i] = 1
    seen = bytearray(gw * gh)
    out = []
    for start in range(gw * gh):
        if not dark[start] or seen[start]:
            continue
        stack = [start]
        seen[start] = 1
        cnt = 0
        x0 = x1 = start % gw
        y0 = y1 = start // gw
        while stack:
            i = stack.pop()
            cnt += 1
            x, y = i % gw, i // gw
            if x < x0:
                x0 = x
            if x > x1:
                x1 = x
            if y < y0:
                y0 = y
            if y > y1:
                y1 = y
            for dx in (-2, -1, 0, 1, 2):                      # (two cells: the gap between a chart's bars is
                for dy in (-2, -1, 0, 1, 2):                  # one picture, not two)
                    nx, ny = x + dx, y + dy
                    if 0 <= nx < gw and 0 <= ny < gh:
                        j = ny * gw + nx
                        if dark[j] and not seen[j]:
                            seen[j] = 1
                            stack.append(j)
        bw, bh = x1 - x0 + 1, y1 - y0 + 1
        if bw < 12 or bh < 9:
            continue
        if x0 == 0 or y0 == 0 or x1 == gw - 1 or y1 == gh - 1:
            continue                                         # the table or the dark round a photographed page
        dens = cnt / float(bw * bh)
        inside = sum(1 for (a, b, c, e) in (boxes if sure is None else sure) if x0 * cell <= (a + c) / 2.0 <= (x1 + 1) * cell and y0 * cell <= (b + e) / 2.0 <= (y1 + 1) * cell)
        if dens >= 0.3 or (cnt >= 30 and inside <= 6 and dens >= 0.04):
            # marks just outside it that were not read as words (the letters under a chart's bars) are part of it
            for _ in range(10):
                grew = False
                if y1 + 1 < gh - 1 and any(dark[(y1 + 1) * gw + x] for x in range(x0, x1 + 1)):
                    y1 += 1
                    grew = True
                if y0 - 1 > 0 and any(dark[(y0 - 1) * gw + x] for x in range(x0, x1 + 1)):
                    y0 -= 1
                    grew = True
                if x0 - 1 > 0 and any(dark[y * gw + x0 - 1] for y in range(y0, y1 + 1)):
                    x0 -= 1
                    grew = True
                if x1 + 1 < gw - 1 and any(dark[y * gw + x1 + 1] for y in range(y0, y1 + 1)):
                    x1 += 1
                    grew = True
                if not grew:
                    if y1 + 2 < gh - 1 and any(dark[(y1 + 2) * gw + x] for x in range(x0, x1 + 1)):
                        y1 += 2                          # one clear row, then marks again (a row of labels under the axis)
                        continue
                    break
            out.append((x0 * cell, y0 * cell, (x1 + 1) * cell, (y1 + 1) * cell))
    return out


def _paper_of(im):
    """The paper's own brightness on a grey picture of a page (most of a page is paper)."""
    small = im.resize((max(1, im.width // 16), max(1, im.height // 16)), Image.BOX)
    px = sorted(small.tobytes())
    return px[int(0.6 * len(px))]


def _skew_of(im, paper):
    """How far a scanned page is turned: the angle (degrees, as Image.rotate takes it) that sets its lines level;
    0.0 when there is nothing sure to go by. Lines of writing, and ruled lines, make the sharpest rows when they lie
    level: a small copy of the page is turned by each angle and its rows compared (Pillow only, about 0.2 s)."""
    w = 800
    h = max(8, int(round(im.height * w / float(im.width))))
    thr = max(60, min(200, paper - 50))
    ink = im.resize((w, h), Image.BILINEAR).point(lambda v: 255 if v < thr else 0)
    mx, my = int(w * 0.1), int(h * 0.1)                      # (the dark round a photographed page says nothing of its lines)

    def score(a):
        r = ink.rotate(a, resample=Image.BILINEAR, fillcolor=0) if a else ink
        r = r.crop((mx, my, w - mx, h - my))
        prof = r.resize((1, r.height), Image.BOX).tobytes()
        return sum((prof[i + 1] - prof[i]) ** 2 for i in range(len(prof) - 1))
    base = score(0.0)
    if base <= 0:
        return 0.0
    best, ba = base, 0.0
    for i in range(-12, 13):
        a = i * 0.5
        if a:
            sc = score(a)
            if sc > best:
                best, ba = sc, a
    fine = ba
    for d in (-0.375, -0.25, -0.125, 0.125, 0.25, 0.375):
        sc = score(ba + d)
        if sc > best:
            best, fine = sc, ba + d
    if abs(fine) >= 5.9 or best < 1.06 * base:
        return 0.0                                           # at the end of what was tried, or no clearer than level
    return fine


def _scan_bands(ink):
    """Dark filled boxes on a scanned page (a table's dark heading row, with its light words in it): [(x0, y0, x1, y1)]
    in pixels, their edges set on the last rows and columns that are mostly ink. ink: the page, ink 255, paper 0."""
    iw, ih = ink.size
    cw, chh = max(14, iw // 100), 8
    gw, gh = iw // cw, ih // chh
    if gw < 4 or gh < 4:
        return []
    cells = ink.crop((0, 0, gw * cw, gh * chh)).resize((gw, gh), Image.BOX).tobytes()
    dark = bytearray(1 if b >= 140 else 0 for b in cells)
    seen = bytearray(len(dark))
    out = []
    start = dark.find(1)
    while start != -1:
        if not seen[start]:
            stack = [start]
            seen[start] = 1
            x0 = x1 = start % gw
            y0 = y1 = start // gw
            cnt = 0
            while stack:
                i = stack.pop()
                cnt += 1
                x, y = i % gw, i // gw
                if x < x0:
                    x0 = x
                if x > x1:
                    x1 = x
                if y < y0:
                    y0 = y
                if y > y1:
                    y1 = y
                for j in (i - 1 if x > 0 else -1, i + 1 if x < gw - 1 else -1, i - gw, i + gw):
                    if 0 <= j < len(dark) and dark[j] and not seen[j]:
                        seen[j] = 1
                        stack.append(j)
            bw, bh = x1 - x0 + 1, y1 - y0 + 1
            if bw >= 4 and bh >= 2 and cnt >= 0.8 * bw * bh and x0 > 0 and x1 < gw - 1 and y0 > 0 and y1 < gh - 1:
                a0, b0, a1, b1 = x0 * cw, y0 * chh, (x1 + 1) * cw, (y1 + 1) * chh
                # the edges, to the pixel
                rows = ink.crop((a0, max(0, b0 - chh), a1, min(ih, b1 + chh))).resize((1, min(ih, b1 + chh) - max(0, b0 - chh)), Image.BOX).tobytes()
                top = max(0, b0 - chh)
                ys = [top + i for i, v in enumerate(rows) if v >= 150]
                cols = ink.crop((max(0, a0 - cw), b0, min(iw, a1 + cw), b1)).resize((min(iw, a1 + cw) - max(0, a0 - cw), 1), Image.BOX).tobytes()
                left = max(0, a0 - cw)
                xs_ = [left + i for i, v in enumerate(cols) if v >= 150]
                if ys and xs_:
                    out.append((xs_[0], ys[0], xs_[-1] + 1, ys[-1] + 1))
        start = dark.find(1, start + 1)
    return out


def _scan_rules(im, paper):
    """Ruled lines on a scanned page, found in its picture (grey, its lines level): a table's grid.
    Returns (hs, vs, bands) in pixels: hs = [(x0, x1, y, thick)], vs = [(y0, y1, x, thick)], bands = dark filled
    boxes [(x0, y0, x1, y1)] (a table's dark heading row).
    A line is a run of ink far longer than any letter's stroke: the picture is squeezed along the line's direction, so
    that a point of the squeezed picture is dark only where a whole stretch of the line is ink (Pillow only)."""
    iw, ih = im.size
    thr = max(50, min(185, paper - 75))
    ink = im.point(lambda v: 255 if v < thr else 0)
    bands = _scan_bands(ink)
    L = max(14, iw // 100)
    tmax = max(8, iw // 190)
    edge = max(8, iw // 70)                                  # (the dark round a photographed page is no line)
    found = []
    for horizontal in (True, False):
        src = ink if horizontal else ink.transpose(Image.TRANSPOSE)
        W, H = src.size
        gw = W // L
        if gw < 4 or H < 8:
            found.append([])
            continue
        cells = src.crop((0, 0, gw * L, H)).resize((gw, H), Image.BOX).tobytes()
        full = bytearray(1 if b >= 215 else 0 for b in cells)
        # nothing in a dark box is a line of its own; a line that runs up to one ends at it
        for (bx0, by0, bx1, by1) in bands:
            if not horizontal:
                bx0, by0, bx1, by1 = by0, bx0, by1, bx1
            c0, c1 = max(0, bx0 // L), min(gw, -(-bx1 // L))
            blank = bytes(c1 - c0)
            for yy in range(max(0, by0 - 2), min(H, by1 + 2)):
                full[yy * gw + c0:yy * gw + c1] = blank
        # what is taller than a line can be is no line
        for x in range(gw):
            col = full[x::gw]
            i = col.find(1)
            while i != -1:
                j = i
                while j < H and col[j]:
                    j += 1
                if j - i > tmax + 2:
                    for yy in range(i, j):
                        full[yy * gw + x] = 0
                i = col.find(1, j)
        seen = bytearray(len(full))
        lines = []
        start = full.find(1)
        while start != -1:
            if not seen[start]:
                stack = [start]
                seen[start] = 1
                x0 = x1 = start % gw
                y0 = y1 = start // gw
                cnt = 0
                while stack:
                    i = stack.pop()
                    cnt += 1
                    x, y = i % gw, i // gw
                    if x < x0:
                        x0 = x
                    if x > x1:
                        x1 = x
                    if y < y0:
                        y0 = y
                    if y > y1:
                        y1 = y
                    for dy in (-1, 0, 1):
                        ny = y + dy
                        if ny < 0 or ny >= H:
                            continue
                        for dx in (-1, 0, 1):
                            nx = x + dx
                            if nx < 0 or nx >= gw:
                                continue
                            j = ny * gw + nx
                            if full[j] and not seen[j]:
                                seen[j] = 1
                                stack.append(j)
                ncell, thick = x1 - x0 + 1, y1 - y0 + 1
                avg = cnt / float(ncell)                         # rows of ink at a place along it
                mid = (y0 + y1 + 1) / 2.0
                if ncell >= 4 and thick <= tmax + 2 + 0.012 * ncell * L and edge <= mid <= H - edge:
                    # its ends, to the pixel: on from the last whole stretch while there is ink in its rows
                    a, b = x0 * L, (x1 + 1) * L
                    ya, yb = max(0, y0 - 1), min(H, y1 + 2)
                    lo, hi = max(0, a - L), min(W, b + L)
                    prof = src.crop((lo, ya, hi, yb)).resize((hi - lo, 1), Image.BOX).tobytes()
                    need = 255.0 * 0.6 / (yb - ya)
                    while a > lo and prof[a - 1 - lo] >= need:
                        a -= 1
                    while b < hi and prof[b - lo] >= need:
                        b += 1
                    if b - a >= max(4 * L, 14 * avg):
                        lines.append((a, b, mid, max(1.0, avg)))
            start = full.find(1, start + 1)
        found.append(lines)
    return found[0], found[1], bands


def _scan_leaders(im, paper):
    """Rows of dots that lead across a line to a number (a contents page), found in a scanned page's picture:
    [(x0, x1, y, h)] in pixels. Tesseract reads such a row as letters ("cececesee"), and the words before it wrong
    with it: the rows are painted over before the page is read, and the number they lead to is reached by a dotted
    tab in Word. A row of dots is ink a few points tall, half of it gaps, with nothing above it, on and on."""
    iw, ih = im.size
    thr = max(60, min(210, paper - 38))                 # (dots are small: on a photographed page they are pale)
    ink = im.point(lambda v: 255 if v < thr else 0)
    L = max(14, iw // 100)
    gw = iw // L
    cells = ink.crop((0, 0, gw * L, ih)).resize((gw, ih), Image.BOX).tobytes()
    hmax = max(5, iw // 210)
    clear = max(8, iw // 150)
    marks = {}                                      # cell x -> [y of a thin band's middle]
    for x in range(gw):
        col = cells[x::gw]
        y = 0
        ys = []
        while y < ih:
            if col[y] < 16:
                y += 1
                continue
            y0 = y
            mx = 0
            while y < ih and col[y] >= 16:
                if col[y] > mx:
                    mx = col[y]
                y += 1
            if 2 <= y - y0 <= hmax and 36 <= mx <= 205:
                above = col[max(0, y0 - clear):y0]
                below = col[y:min(ih, y + 4)]
                if (not above or max(above) < 14) and (not below or max(below) < 14):
                    ys.append((y0 + y) / 2.0)
        marks[x] = ys
    # chains of cells side by side at one height (a cell or two may be missing: a pale dot, a speck above)
    raw = []
    used = set()
    for x in range(gw):
        for y in marks[x]:
            if (x, y) in used:
                continue
            chain = [(x, y)]
            used.add((x, y))
            cx, cy = x, y
            while True:
                hit = None
                for step in (1, 2, 3):
                    if cx + step >= gw:
                        break
                    nxt = [yy for yy in marks[cx + step] if abs(yy - cy) <= 2.5 + step and (cx + step, yy) not in used]
                    if nxt:
                        hit = (cx + step, min(nxt, key=lambda yy: abs(yy - cy)))
                        break
                if hit is None:
                    break
                cx, cy = hit
                chain.append(hit)
                used.add(hit)
            if len(chain) >= 4:
                raw.append(chain)
    out = []
    for chain in raw:
        ym = sorted(c[1] for c in chain)[len(chain) // 2]
        x0, x1 = chain[0][0] * L, (chain[-1][0] + 1) * L
        # the ends, to the pixel: on while the dots go on and nothing stands above them
        ya, yb = max(0, int(ym - hmax)), min(ih, int(ym + hmax))
        lo, hi = max(0, x0 - 3 * L), min(iw, x1 + 3 * L)
        strip = ink.crop((lo, ya, hi, yb)).resize((hi - lo, 1), Image.BOX).tobytes()
        over = ink.crop((lo, max(0, ya - clear - hmax), hi, ya)).resize((hi - lo, 1), Image.BOX).tobytes()
        # (from its middle outwards: a chain's last piece may be a figure's foot, which stands at the dots' height)
        a = b = (x0 + x1) // 2
        gap = 0
        while a > lo and over[a - 1 - lo] < 8 and gap <= 3 * L:
            gap = 0 if strip[a - 1 - lo] >= 10 else gap + 1
            a -= 1
        a += gap
        gap = 0
        while b < hi and over[b - lo] < 8 and gap <= 3 * L:
            gap = 0 if strip[b - lo] >= 10 else gap + 1
            b += 1
        b -= gap
        out.append([a, b, ym, hmax])
    # pieces of one row of dots are one
    out.sort(key=lambda l: (round(l[2] / 6.0), l[0]))
    merged = []
    for l in out:
        if merged and abs(merged[-1][2] - l[2]) <= 5 and l[0] - merged[-1][1] <= 3 * L:
            merged[-1][1] = max(merged[-1][1], l[1])
        else:
            merged.append(l)
    return [tuple(l) for l in merged if l[1] - l[0] >= 5 * L]


class _Sheet(object):
    """What pdfplumber's table finder asks of a page: its box."""
    def __init__(self, w, h):
        self.bbox = (0, 0, w, h)


def _scan_tables(im, paper, k, pw, ph):
    """A scanned page's ruled tables, from the lines found in its picture (grey, level; k = page points a pixel).
    Returns (tables with empty cells, hs, vs, fills in page points as a read page holds them, the boxes in pixels to
    paint over before the page is read). Only what makes a table's grid is used: any other line stays in the picture
    and is read as before."""
    hs, vs, bands = _scan_rules(im, paper)
    if len(vs) < 2 or len(hs) + 2 * len(bands) < 2:
        return [], [], [], [], [], []
    iw, ih = im.size
    L = max(14, iw // 100)
    tol = max(6, iw // 260)
    vs = [list(v) for v in vs]
    fills_px = []
    for (bx0, by0, bx1, by1) in bands:
        bh = by1 - by0
        if bh > 0.12 * ih or bx1 - bx0 < 0.08 * iw:
            continue
        # a dark row of a table: a ruled line of its own width lies just under or over it
        mate = None
        for (x0, x1, y, th) in hs:
            if abs(x0 - bx0) <= 1.5 * L and abs(x1 - bx1) <= 1.5 * L and (-tol <= y - by1 <= 6 * bh or -tol <= by0 - y <= 6 * bh):
                mate = (x0, x1)
                break
        if mate is None:
            continue
        box = im.crop((bx0 + L, by0 + 2, bx1 - L, by1 - 2)) if bx1 - bx0 > 3 * L and bh > 6 else im.crop((bx0, by0, bx1, by1))
        px = sorted(box.resize((max(1, box.width // 4), max(1, box.height // 2)), Image.BOX).tobytes())
        level = px[len(px) // 3]                                 # the fill itself (the row's words are lighter or darker)
        fills_px.append((mate[0], by0, mate[1], by1, level))
        # the columns' lines go on through the row where the picture shows them in it
        for v in vs:
            x = v[2]
            if not (mate[0] + tol < x < mate[1] - tol):
                continue
            below, above = abs(v[0] - by1) <= L + tol, abs(v[1] - by0) <= L + tol
            if not (below or above) or bh < 8:
                continue
            xa, xb = max(0, int(x) - 3), min(iw, int(x) + 4)
            cols = im.crop((xa, by0 + 3, xb, by1 - 3)).resize((xb - xa, 1), Image.BOX).tobytes()
            if max(abs(c - level) for c in cols) >= 12:
                if below:
                    v[0] = by0
                else:
                    v[1] = by1
    t = _Page()
    t.w, t.h = pw, ph
    t.hs = [[x0 * k, x1 * k, y * k, max(0.5, min(3.0, th * k)), "000000", True, False] for (x0, x1, y, th) in hs]
    t.vs = [[y0 * k, y1 * k, x * k, max(0.5, min(3.0, th * k)), "000000", True, False] for (y0, y1, x, th) in vs]
    t.fills = [[x0 * k, y0 * k, x1 * k, y1 * k, "%02X%02X%02X" % (lv, lv, lv), False] for (x0, y0, x1, y1, lv) in fills_px]
    try:
        tables = _find_tables(_Sheet(pw, ph), t)
    except Exception:
        tables = []
    keep = []
    for tb in tables:
        area = (tb.x1 - tb.x0) * (tb.bottom - tb.top)
        if len(tb.cells) >= 2 or (area <= 0.5 * pw * ph and tb.x1 - tb.x0 >= 30 and tb.bottom - tb.top >= 14):
            keep.append(tb)
    if not keep:
        return [], [], [], [], [], []

    def inside(x0, y0, x1, y1):
        return any(tb.x0 - 3 <= x0 and x1 <= tb.x1 + 3 and tb.top - 3 <= y0 and y1 <= tb.bottom + 3 for tb in keep)
    paint, khs, kvs = [], [], []
    for line, (x0, x1, y, th) in zip(t.hs, hs):
        if inside(line[0], line[2], line[1], line[2]):
            khs.append(line)
            paint.append((x0 - 2, int(y - th / 2.0 - 2), x1 + 2, int(math.ceil(y + th / 2.0 + 2))))
    for line, (y0, y1, x, th) in zip(t.vs, vs):
        if inside(line[2], line[0], line[2], line[1]):
            kvs.append(line)
            paint.append((int(x - th / 2.0 - 2), y0 - 2, int(math.ceil(x + th / 2.0 + 2)), y1 + 2))
    kfills = [f for f in t.fills if inside(f[0], f[1], f[2], f[3])]
    used = [fp for f, fp in zip(t.fills, fills_px) if inside(f[0], f[1], f[2], f[3])]
    dark = [(x0, y0, x1, y1) for (x0, y0, x1, y1, lv) in fills_px]
    # a line is not painted where it runs in a dark row (the paper's colour there would be a white stroke through it)
    cut = []
    for (x0, y0, x1, y1) in paint:
        parts = [(x0, y0, x1, y1)]
        for (dx0, dy0, dx1, dy1) in dark:
            nxt = []
            for (a0, b0, a1, b1) in parts:
                if a1 <= dx0 or a0 >= dx1 or b1 <= dy0 or b0 >= dy1:
                    nxt.append((a0, b0, a1, b1))
                    continue
                if b1 - b0 > a1 - a0:                            # a line down the page: what is above and below the row
                    if b0 < dy0:
                        nxt.append((a0, b0, a1, dy0))
                    if b1 > dy1:
                        nxt.append((a0, dy1, a1, b1))
                elif not (dy0 - 4 <= b0 and b1 <= dy1 + 4):      # a line across, not the row's own edge
                    nxt.append((a0, b0, a1, b1))
            parts = nxt
        cut += parts
    return keep, khs, kvs, kfills, cut, used


def _marks_again(im, paper, lines, figs, root, leaders=()):
    """What stands before a line and was not read: a list's number, a bullet.
    The numbers of a list stand in a column of their own, and Tesseract often takes such a column for one tall line
    of writing turned on its side ("PoP", "er", trusted at 0 of 100) or passes it by: the list came out without its
    numbers. Here every line is looked at again: is there ink just left of its first word that no read word covers?
    A small round blot at the letters' middle height is a bullet. Anything else is read once more WITH its line (the
    strips of all such lines, one under the other, in one small picture), and taken only if it reads as a list's mark.
    lines: [(par, box, slope, off, x_size, words)] as _ocr_page holds them; a mark found is put before the words."""
    from lxml import etree
    from PIL import ImageDraw
    iw, ih = im.size
    thr = max(50, min(185, paper - 60))
    unread = im.point(lambda v: 255 if v < thr else 0)
    d = ImageDraw.Draw(unread)
    lead = {}
    for ln in lines:
        order = sorted(ln[5], key=lambda w_: w_[1][0])
        # what was read of a mark: ".", "_", ":" alone before the words; or a short word read with little trust that
        # stands clear of the words ("(SS", "ES", "go" for "2.", "3.", "4." on grey paper). It is looked at again, and
        # gives way only to something that reads as a list's mark.
        junk = []
        xs_ = ln[4]
        while len(junk) < 2 and len(order) - len(junk) >= 2:
            w_, nx_ = order[len(junk)], order[len(junk) + 1]
            alone = len(w_[0]) == 1 and w_[0] in "._:,;~-'`|"
            doubt = not junk and len(w_[0]) <= 3 and w_[2] < 75 and nx_[1][0] - w_[1][2] > 0.45 * xs_ and not _NUM_MARK.match(w_[0]) and w_[0] not in _BULLETS
            if not (alone or doubt):
                break
            junk.append(w_)
        lead[id(ln)] = (order[len(junk)][1], junk)
        for w_ in order[len(junk):]:
            wb = w_[1]
            d.rectangle((wb[0] - 2, wb[1] - 2, wb[2] + 2, wb[3] + 2), fill=0)
    for f in figs:
        d.rectangle(f, fill=0)
    strips = []
    for ln in lines:
        lb, off, xs, words = ln[1], ln[3], ln[4], ln[5]
        if xs < 8:
            continue
        fw, junk = lead[id(ln)]
        x1 = fw[0] - max(2, int(0.1 * xs))
        x0 = max(0, int(fw[0] - 3.4 * xs))
        ya, yb = max(0, lb[1] - 1), min(ih, lb[3] + 1)
        if x1 - x0 < 0.3 * xs or yb - ya < 4:
            continue
        cols = unread.crop((x0, ya, x1, yb)).resize((x1 - x0, 1), Image.BOX).tobytes()
        need = 255.0 * 2.5 / (yb - ya)                           # two or three points of ink in a column
        i, gap, lo, hi = len(cols) - 1, 0, None, None
        while i >= 0:                                            # the ink nearest the words, leftwards from them
            if cols[i] >= need:
                if hi is None:
                    hi = i
                lo, gap = i, 0
            else:
                gap += 1
                if (hi is not None and gap > 0.4 * xs) or (hi is None and gap > 2.6 * xs):
                    break
            i -= 1
        if hi is None:
            continue
        bw = hi - lo + 1
        if bw < 3 or bw > 2.4 * xs:
            continue
        rws = unread.crop((x0 + lo, ya, x0 + hi + 1, yb)).resize((1, yb - ya), Image.BOX).tobytes()
        ys = [j for j, v in enumerate(rws) if v >= max(8.0, 255.0 * 1.5 / bw)]
        if not ys:
            continue
        box = [x0 + lo, ya + ys[0], x0 + hi + 1, ya + ys[-1] + 1]
        bh = box[3] - box[1]
        if bh < 0.12 * xs:
            continue                                             # a speck, a thin stroke
        base = lb[3] + off                                       # (the page is level: the line's baseline)
        cy = (box[1] + box[3]) / 2.0
        if not junk and 0.12 * xs <= bw <= 0.55 * xs and bh <= 0.55 * xs and 0.6 <= bw / float(bh) <= 1.7 and base - 0.8 * xs <= cy <= base - 0.05 * xs:
            words.insert(0, ("•", box, 60))                 # a round blot before the words, at the letters' middle
            lb[0] = min(lb[0], box[0])
            continue
        strips.append((ln, box, (max(0, box[0] - int(0.6 * xs)), max(0, lb[1] - 8), min(iw, min(lb[2], fw[0] + int(14 * xs)) + 6), min(ih, lb[3] + 8)), fw[0], junk))
    # The number a row of dots leads to (a contents page's page numbers stand in a column of their own at the right:
    # the same trouble as a list's numbers at the left). Where nothing was read after the dots and ink stands there,
    # the line is read once more with it.
    ends = []
    for (lx0, lx1, ly, lh) in leaders:
        ln = None
        for cand in lines:
            lb, off, xs = cand[1], cand[3], cand[4]
            if abs(lb[3] + off - ly) <= 0.5 * xs and lb[0] < lx0 + xs and lb[2] <= lx1 + 0.6 * xs:
                if ln is None or cand[1][2] > ln[1][2]:
                    ln = cand
        if ln is None:
            continue
        xs = ln[4]
        if any(abs(c[1][3] + c[3] - ly) <= 0.5 * c[4] and any(w_[1][0] >= lx1 - 0.4 * xs and w_[1][0] <= lx1 + 6 * xs for w_ in c[5]) for c in lines):
            continue                                             # something was read after the dots
        x0, x1 = min(iw, int(lx1 + 1)), min(iw, int(lx1 + 5 * xs))
        ya, yb = max(0, int(ly - 1.1 * xs)), min(ih, int(ly + 0.35 * xs))
        if x1 - x0 < 4 or yb - ya < 4:
            continue
        cols = unread.crop((x0, ya, x1, yb)).resize((x1 - x0, 1), Image.BOX).tobytes()
        need = 255.0 * 2.5 / (yb - ya)
        inked = [i for i, v in enumerate(cols) if v >= need]
        if not inked or inked[0] > 2.5 * xs:
            continue
        hi = inked[0]
        for i in inked:                                          # the ink nearest the dots, over gaps of half a letter
            if i - hi > 0.5 * xs:
                break
            hi = i
        box = [x0 + inked[0], ya, x0 + hi + 1, yb]
        rws = unread.crop((box[0], ya, box[2], yb)).resize((1, yb - ya), Image.BOX).tobytes()
        ys = [j for j, v in enumerate(rws) if v >= 8]
        if not ys or ys[-1] - ys[0] < 0.25 * xs:
            continue
        box[1], box[3] = ya + ys[0], ya + ys[-1] + 1
        lb = ln[1]
        ends.append((ln, box, (max(0, lb[0] - 6), max(0, min(lb[1], box[1]) - 8), min(iw, box[2] + 10), min(ih, max(lb[3], box[3]) + 8)), lx1))
    strips = strips[:40]
    ends = ends[:60]
    if not strips and not ends:
        return
    pad = 36
    both = [st[2] for st in strips] + [e[2] for e in ends]
    cw = max(b[2] - b[0] for b in both) + 2 * pad
    ch = sum(b[3] - b[1] for b in both) + pad * (len(both) + 1)
    comp = Image.new("L", (cw, ch), paper)
    y = pad
    places = []
    for b in both:
        comp.paste(im.crop(b), (pad, y))
        places.append((y, y + b[3] - b[1], b))
        y += b[3] - b[1] + pad
    png = root + "_m.png"
    comp.save(png)
    try:
        subprocess.run(["tesseract", png, root + "_m", "--psm", "6", "-l", "eng", "hocr"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=40, check=False)
        tree = etree.parse(root + "_m.hocr", etree.HTMLParser(encoding="utf-8"))
    except Exception:
        return
    got = defaultdict(list)
    for wd in tree.iter("span"):
        if (wd.get("class") or "") != "ocrx_word":
            continue
        t = _clean("".join(wd.itertext()).strip())
        mw = re.search(r"bbox (\d+) (\d+) (\d+) (\d+)", wd.get("title") or "")
        mc = re.search(r"x_wconf (\d+)", wd.get("title") or "")
        if not t or not mw:
            continue
        wb = [int(v) for v in mw.groups()]
        cy = (wb[1] + wb[3]) / 2.0
        for i, (y0, y1, b) in enumerate(places):
            if y0 - pad / 2.0 <= cy <= y1 + pad / 2.0:
                got[i].append((t, wb[0] - pad + b[0], wb[2] - pad + b[0], int(mc.group(1)) if mc else 50))
                break
    for i, (ln, box, b, first, junk) in enumerate(strips):
        left = sorted([g for g in got.get(i, []) if g[2] <= first + 2], key=lambda g: g[1])
        if not left or len(left) > 2:
            continue
        mark = "".join(g[0] for g in left)
        if not (_NUM_MARK.match(mark) or re.match(r"^\d{1,3}$", mark) or (len(mark) == 1 and mark in _BULLETS)):
            continue
        for w_ in junk:                                          # what was first read of the mark gives way to the mark
            if w_ in ln[5]:
                ln[5].remove(w_)
            box = [min(box[0], w_[1][0]), min(box[1], w_[1][1]), max(box[2], w_[1][2]), max(box[3], w_[1][3])]
        ln[5].insert(0, (mark, box, min(g[3] for g in left)))
        ln[1][0] = min(ln[1][0], box[0])
    for j, (ln, box, b, lx1) in enumerate(ends):
        right = sorted([g for g in got.get(len(strips) + j, []) if g[1] >= lx1 - 3], key=lambda g: g[1])
        if not right or len(right) > 2:
            continue
        num = "".join(g[0] for g in right).strip(".")
        if not re.match(r"^(\d{1,4}|[ivxlc]{1,6}|[IVXLC]{1,6})$", num):
            continue                                             # only what a row of dots leads to: a page's number
        ln[5].append((num, box, min(g[3] for g in right)))
        ln[1][2] = max(ln[1][2], box[2])


def _lines_again(im, paper, lines, figs, tables, k, root):
    """A line that Tesseract's look at the whole page passed by: a page number alone at the foot of a photographed
    page, a short caption, a last line by itself. (Measured 19 Nov 2026: on a made-up phone photograph of a handout,
    "Page 1" at the foot of the first page was in no line at all, and the Word file lost it.)
    Ink that no read word covers, in a band about a line high, clear of the pictures, the tables and the edge of the
    picture, is read once more: all such strips one under the other in one small picture (as _marks_again does),
    taken only where they read as words with trust. Nothing is added that is not on the page: specks, a fold, the
    table's edge do not read as words with trust, and a strip that reads as signs is left.
    lines as _ocr_page holds them; what is read is added to them as lines of their own. Returns how many."""
    from lxml import etree
    from PIL import ImageDraw
    iw, ih = im.size
    sizes = sorted(l[4] for l in lines if l[4] >= 8)
    if len(sizes) < 3:
        return 0
    xs = sizes[len(sizes) // 2]
    thr = max(50, min(185, paper - 60))
    unread = im.point(lambda v: 255 if v < thr else 0)
    d = ImageDraw.Draw(unread)
    for ln in lines:
        lb = ln[1]
        d.rectangle((lb[0] - 4, lb[1] - 4, lb[2] + 4, lb[3] + 4), fill=0)
        for w in ln[5]:
            wb = w[1]
            d.rectangle((wb[0] - 4, wb[1] - 4, wb[2] + 4, wb[3] + 4), fill=0)
    for f in figs:
        d.rectangle((f[0] - 4, f[1] - 4, f[2] + 4, f[3] + 4), fill=0)
    for tb in tables:
        d.rectangle((tb.x0 / k - 6, tb.top / k - 6, tb.x1 / k + 6, tb.bottom / k + 6), fill=0)
    # the unread ink, a cell to every 3 by 3 pixels; the cells' runs along each row, joined where they are closer than a
    # word's gap; runs that touch from one row to the next are one blot (a word, a line of words, or a mark)
    sc = 3
    sw, sh = max(1, iw // sc), max(1, ih // sc)
    cells = unread.resize((sw, sh), Image.BOX).tobytes()
    g = max(1, int(1.0 * xs / sc))
    runs, parent = [], []

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i
    prev = []
    for y in range(sh):
        row, cur, x = cells[y * sw:(y + 1) * sw], [], 0
        while x < sw:
            if row[x] < 40:
                x += 1
                continue
            x0, last = x, x
            while x < sw and (row[x] >= 40 or x - last <= g):
                if row[x] >= 40:
                    last = x
                x += 1
            i = len(runs)
            runs.append((y, x0, last + 1))
            parent.append(i)
            for j in prev:
                if runs[j][1] <= last + 1 and x0 <= runs[j][2]:
                    ri, rj = find(i), find(j)
                    if ri != rj:
                        parent[ri] = rj
            cur.append(i)
        prev = cur
        if len(runs) > 60000:
            return 0                                 # a picture full of marks (a photo, a texture): not looked at
    blots = {}
    for i, (y, x0, x1) in enumerate(runs):
        r = find(i)
        b_ = blots.get(r)
        blots[r] = [min(b_[0], x0), min(b_[1], y), max(b_[2], x1), max(b_[3], y + 1)] if b_ else [x0, y, x1, y + 1]
    cands = []
    for b_ in blots.values():
        xa, ya, xb, yb = b_[0] * sc, b_[1] * sc, b_[2] * sc, b_[3] * sc
        # (xs is the running words' size: a line of smaller print, a running head, a footnote, stands lower; its
        # capitals alone are a third of it)
        if not (0.35 * xs <= yb - ya <= 3.5 * xs) or xb - xa < 0.35 * xs:
            continue
        if xa <= 0.01 * iw or xb >= 0.99 * iw or ya <= 0.01 * ih or yb >= 0.99 * ih:
            continue                                 # at the picture's edge: the table, the hand, the paper's edge
        dens = unread.crop((xa, ya, xb, yb)).resize((1, 1), Image.BOX).tobytes()[0] / 255.0
        if 0.05 <= dens <= 0.6:
            cands.append((xa, ya, xb, yb))
    cands.sort(key=lambda c: (c[1], c[0]))
    if not cands or len(cands) > 12:
        return 0                                     # nothing missed; or the picture is full of marks (a photo, a texture)
    # each strip a little wider than its ink (a letter's faint edge is not cut), its paper made white: a strip's own
    # paper lighter or darker than the rest drew an edge round it, and the edge was read as "|"
    m = max(2, int(0.2 * xs))
    cands = [(max(0, c[0] - m), max(0, c[1] - m), min(iw, c[2] + m), min(ih, c[3] + m)) for c in cands]
    pad, gap = int(0.6 * xs), int(1.6 * xs)
    W = max(c[2] - c[0] for c in cands) + 2 * pad
    H = sum(c[3] - c[1] + 2 * pad for c in cands) + gap * (len(cands) - 1)
    strip = Image.new("L", (W, H), 255)
    slots, at = [], 0
    for c in cands:
        piece = im.crop(c)
        hist = piece.histogram()
        tot, acc, pl = sum(hist), 0, 255
        for v in range(255, -1, -1):                 # the paper: the lightest tenth of the strip
            acc += hist[v]
            if acc >= 0.1 * tot:
                pl = max(1, v)
                break
        strip.paste(piece.point(lambda v, pl=pl: min(255, int(v * 255.0 / pl))), (pad, at + pad))
        slots.append((at, at + (c[3] - c[1]) + 2 * pad, c))
        at += (c[3] - c[1]) + 2 * pad + gap
    png = root + "_l.png"
    strip.save(png)
    try:
        subprocess.run(["tesseract", png, root + "_l", "--psm", "6", "-l", "eng", "hocr"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=40, check=False)
        tree = etree.parse(root + "_l.hocr", etree.HTMLParser(encoding="utf-8"))
    except Exception:
        return 0
    got = {}
    for ln in tree.iter("span"):
        if (ln.get("class") or "") not in ("ocr_line", "ocr_header", "ocr_caption", "ocr_textfloat"):
            continue
        mx = re.search(r"x_size ([\d.]+)", ln.get("title") or "")
        for wd in ln.iter("span"):
            if (wd.get("class") or "") != "ocrx_word":
                continue
            t = _clean("".join(wd.itertext()).strip())
            mw = re.search(r"bbox (\d+) (\d+) (\d+) (\d+)", wd.get("title") or "")
            mc = re.search(r"x_wconf (\d+)", wd.get("title") or "")
            if not t or not mw:
                continue
            wb = [int(v) for v in mw.groups()]
            cy = (wb[1] + wb[3]) / 2.0
            for si, (s0, s1, c) in enumerate(slots):
                if s0 <= cy < s1:
                    ox, oy = c[0] - pad, c[1] - pad - s0
                    got.setdefault(si, []).append((t, [wb[0] + ox, wb[1] + oy, wb[2] + ox, wb[3] + oy], int(mc.group(1)) if mc else 0,
                                                   float(mx.group(1)) if mx else None))
                    break
    added = 0
    for si, ws in sorted(got.items()):
        ws = [w for w in ws if w[2] >= 60]
        if not ws:
            continue
        text = "".join(w[0] for w in ws)
        alnum = sum(1 for ch in text if ch.isalnum())
        conf = sum(w[2] for w in ws) / float(len(ws))
        if alnum < 0.6 * len(text):
            continue                                 # signs, not words
        if not (alnum >= 2 and conf >= 75) and not (text.isdigit() and conf >= 85):
            continue
        lb = [min(w[1][0] for w in ws), min(w[1][1] for w in ws), max(w[1][2] for w in ws), max(w[1][3] for w in ws)]
        lxs = next((w[3] for w in ws if w[3]), None) or xs
        lines.append((-(1000 + si), lb, 0.0, 0.0, lxs, [(w[0], w[1], w[2]) for w in ws]))
        added += 1
    return added


def _ocr_bold(doc):
    """Which read words are bold, from how thick their strokes are (bold letters' stems are half as thick again).
    The measure is the document's own: the middle stroke of all its read words (of a book face's plain letters, about
    a ninth of the letters' height, where the document is too short to say). A line is bold when its words' strokes
    are 1.4 times that (1.22 for a line set larger than its page's common size, as headings are); a word in a plain
    line, when it stands out by 1.45 and its neighbour does too, or it is long, or a figure follows it ("Table 1.").
    Never on little evidence: bold that is not on the paper would be an addition."""
    pages = doc.get("ocr_bold") or []
    vals = [v for (_rows, thick, _b) in pages for v in thick.values()]
    if len(vals) < 4:
        return
    usual_t = _median(vals) if len(vals) >= 60 else min(_median(vals), 0.118)
    if usual_t <= 0:
        return
    for rows, thick, body_sz in pages:
        for words, size in rows:
            got = [thick[id(w)] for w in words if id(w) in thick]
            if not got:
                continue
            need = 1.22 if (body_sz and size >= 1.06 * body_sz) else 1.4
            # (a heading's number, "1." or "5.", is a thin mark: it does not say the words are not bold)
            worded = [thick[id(w)] for w in words if id(w) in thick and len(re.sub(r"[\W_]", "", w.text())) >= 3] or got
            if _median(got) >= need * usual_t and min(worded) >= 1.1 * usual_t:
                hits = list(words)
            else:
                strong = [w for w in words if thick.get(id(w), 0) >= 1.45 * usual_t]
                hits = []
                for w in strong:
                    i = words.index(w)
                    beside = any(0 <= j < len(words) and words[j] in strong for j in (i - 1, i + 1))
                    if beside or len(w.text()) >= 7:
                        hits.append(w)
                    elif i + 1 < len(words) and id(words[i + 1]) not in thick and re.match(r"^\d{1,3}[.:)]?$", words[i + 1].text()):
                        hits += [w, words[i + 1]]                # "Table 1.", "Figure 2:": the number is the word's
            for w in hits:
                f = w.parts[0][1]
                w.parts[0][1] = (f[0], f[1], True, f[3], f[4], f[5])


def _ocr_page(src_path, pg, doc):
    """A page that is a picture of a page: read with Tesseract, its words set out like any other page's.
    True when words were read."""
    from lxml import etree
    sc, pw, ph = _paper_for(pg.ow, pg.oh)
    tmp = tempfile.mkdtemp(prefix="p2w_ocr_")
    try:
        root = os.path.join(tmp, "p")
        target = 2100 if pg.ow <= pg.oh else 2800
        try:
            subprocess.run(["pdftoppm", "-gray", "-scale-to-x", str(target), "-scale-to-y", "-1", "-f", str(pg.n + 1), "-l", str(pg.n + 1), "-png", "-singlefile", src_path, root],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=60, check=False)
        except Exception:
            return False
        png = root + ".png"
        if not os.path.exists(png):
            return False
        with Image.open(png) as im0:
            im = im0.convert("L")
        iw, ih = im.size
        k = (pg.ow * sc) / float(iw)                          # page points per pixel
        # The page is set level first (a scan a little turned, a page photographed by hand): the reading is surer,
        # and what stands in a row or a column on the paper stands so in the picture. Then the tables' ruled lines
        # are found in the picture and painted over in the copy that is read: Tesseract takes them for ";", ":", "I".
        skew, paper = 0.0, 255
        tables, t_hs, t_vs, t_fills = [], [], [], []
        leaders = []
        clean = im
        try:
            paper = _paper_of(im)
            skew = _skew_of(im, paper)
            if abs(skew) >= 0.2:
                im = im.rotate(skew, resample=Image.BICUBIC, fillcolor=paper)
            else:
                skew = 0.0
            clean = im
            tables, t_hs, t_vs, t_fills, paint, dark_rows = _scan_tables(im, paper, k, pw, ph)
            if paint or dark_rows:
                from PIL import ImageDraw
                clean = im.copy()
                # light words on a table's dark row are read as dark words on paper: the row is turned about (its
                # own dark becomes the paper's light, its light words dark); Tesseract passes light-on-dark by
                for (x0, y0, x1, y1, lv) in dark_rows:
                    if lv < paper - 60:
                        box = (max(0, int(x0)), max(0, int(y0)), min(iw, int(x1)), min(ih, int(y1)))
                        span = float(max(1, 255 - lv))
                        clean.paste(im.crop(box).point(lambda v, lv=lv, span=span: max(0, min(paper, int(paper * (255 - v) / span)))), box)
                d = ImageDraw.Draw(clean)
                for box in paint:
                    d.rectangle(box, fill=paper)
            # rows of dots that lead to a number (a contents page) are read as letters, and spoil the words before
            # them: painted over too (never in a table: its dotted rules are its own)
            try:
                leaders = [l for l in _scan_leaders(im, paper)
                           if not any(tb.x0 - 2 <= l[0] * k and l[1] * k <= tb.x1 + 2 and tb.top - 2 <= l[2] * k <= tb.bottom + 2 for tb in tables)]
            except Exception:
                leaders = []
            if leaders:
                from PIL import ImageDraw
                if clean is im:
                    clean = im.copy()
                d = ImageDraw.Draw(clean)
                for (lx0, lx1, ly, lh) in leaders:
                    d.rectangle((lx0, ly - lh, lx1, ly + lh), fill=paper)
            if clean is not im or skew:
                clean.save(png)
        except Exception:
            # the page as it was drawn, read as before
            skew, tables, t_hs, t_vs, t_fills, leaders = 0.0, [], [], [], [], []
            try:
                with Image.open(png) as im0:
                    im = im0.convert("L")
            except Exception:
                return False
            clean = im
        try:
            subprocess.run(["tesseract", png, root, "--psm", "3", "-l", "eng", "hocr"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=120, check=False)
        except Exception:
            return False
        hocr = root + ".hocr"
        if not os.path.exists(hocr):
            return False
        try:
            tree = etree.parse(hocr, etree.HTMLParser(encoding="utf-8"))
        except Exception:
            return False
        lines = []
        slopes = []
        dropped = []
        no_xs = set()                                        # lines Tesseract gave no letter size for
        try:
            thr_s = max(50, min(185, paper - 60))
            solid = clean.point(lambda v: 255 if v < thr_s else 0)
        except Exception:
            solid = None
        for par in tree.iter("p"):
            if "ocr_par" not in (par.get("class") or ""):
                continue
            for ln in par.iter("span"):
                cls = ln.get("class") or ""
                if cls not in ("ocr_line", "ocr_header", "ocr_caption", "ocr_textfloat"):
                    continue
                title = ln.get("title") or ""
                m = re.search(r"bbox (\d+) (\d+) (\d+) (\d+)", title)
                if not m:
                    continue
                lb = [int(v) for v in m.groups()]
                mb = re.search(r"baseline (-?[\d.]+) (-?[\d.]+)", title)
                slope, off = (float(mb.group(1)), float(mb.group(2))) if mb else (0.0, 0.0)
                mx = re.search(r"x_size ([\d.]+)", title)
                xs = float(mx.group(1)) if mx else (lb[3] - lb[1])
                if not mx:
                    no_xs.add(id(lb))
                words = []
                for wd in ln.iter("span"):
                    if (wd.get("class") or "") != "ocrx_word":
                        continue
                    t = _clean("".join(wd.itertext()).strip())
                    mw = re.search(r"bbox (\d+) (\d+) (\d+) (\d+)", wd.get("title") or "")
                    mc = re.search(r"x_wconf (\d+)", wd.get("title") or "")
                    conf = int(mc.group(1)) if mc else 50
                    if not t or not mw or (conf < 15 and not any(ch.isalnum() for ch in t)):
                        continue
                    if t in ("|", "||", "¦") and conf < 70:
                        continue                             # an upright stroke read with little trust: a speck, a fold
                    wb = [int(v) for v in mw.groups()]
                    if wb[2] - wb[0] < 2 or wb[3] - wb[1] < 2:
                        continue
                    if solid is not None and wb[2] - wb[0] >= 6 and wb[3] - wb[1] >= 6:
                        # letters leave most of their box as paper; a chart's bar, a block of colour does not. Such a
                        # "word" is part of a picture, and painting it over painted the picture away.
                        try:
                            cut = solid.crop((wb[0], wb[1], wb[2], wb[3]))
                            fill = cut.resize((1, 1), Image.BOX).tobytes()[0] / 255.0
                        except Exception:
                            fill = 0.0
                        # (measured: letters, bold capitals in small print too, fill at most six tenths of their box;
                        # a single upright stroke, "I", "l", "1", may fill all of it: it is narrow)
                        if fill >= 0.75 and (wb[3] - wb[1]) * (wb[2] - wb[0]) >= 0.25 * xs * xs and wb[2] - wb[0] >= 0.3 * xs:
                            continue
                    words.append((t, wb, conf))
                if not words:
                    continue
                if all(w[2] < 30 for w in words) and (len(words) == 1 or lb[3] - lb[1] > 2.2 * xs):
                    dropped.append(lb)
                    continue                                 # nothing in it was read with any trust: specks, or marks read as letters
                if len(words) >= 2 and lb[3] - lb[1] > 2.6 * xs:
                    cys = [(w[1][1] + w[1][3]) / 2.0 for w in words]
                    if max(cys) - min(cys) > 1.5 * xs:
                        # a column of figures or marks taken for one line of writing ("NY NY FF KF" for the page
                        # numbers of a contents page): no line at all. Each is read again with the line it stands in.
                        dropped.append(lb)
                        continue
                if len(words) >= 3 and abs(slope) < 0.01:
                    # a word that stands at another line's height (a list's "4." read at the end of line 3) is no word
                    # of this line: it becomes a line of its own, and joins the row it stands in
                    mid_y = _median([(w[1][1] + w[1][3]) / 2.0 for w in words])
                    stray = [w for w in words if abs((w[1][1] + w[1][3]) / 2.0 - mid_y) > 0.85 * xs]
                    if stray and len(stray) <= 2:
                        base0 = lb[3] + off
                        words = [w for w in words if w not in stray]
                        lb = [min(w[1][0] for w in words), min(w[1][1] for w in words), max(w[1][2] for w in words), max(w[1][3] for w in words)]
                        off = min(0.0, base0 - lb[3])
                        for w in stray:
                            sb = list(w[1])
                            lines.append((id(par), sb, 0.0, 0.0, xs, [w]))
                if len(words) >= 3:
                    slopes.append(slope)
                lines.append((id(par), lb, slope, off, xs, words))
        if not lines:
            return False
        nwords = sum(len(l[5]) for l in lines)
        if nwords < 3:
            return False
        figs_px = _scan_figures(clean, [w[1] for l in lines for w in l[5] if w[2] >= 50],
                                [w[1] for l in lines for w in l[5] if w[2] >= 70 and sum(1 for ch in w[0] if ch.isalnum()) >= 3])
        # one drawing found in two parts (a chart's bars to the left and to the right of a low one) is one picture
        merged_f = []
        for f in sorted(figs_px, key=lambda f: (f[1], f[0])):
            for i, g in enumerate(merged_f):
                ov = min(f[3], g[3]) - max(f[1], g[1])
                gap = max(f[0], g[0]) - min(f[2], g[2])
                if ov >= 0.6 * min(f[3] - f[1], g[3] - g[1]) and gap <= 0.03 * iw:
                    merged_f[i] = (min(f[0], g[0]), min(f[1], g[1]), max(f[2], g[2]), max(f[3], g[3]))
                    break
            else:
                merged_f.append(f)
        figs_px = merged_f
        try:
            _marks_again(clean, paper, lines, figs_px, root, leaders)
        except Exception:
            pass
        try:
            _lines_again(clean, paper, lines, figs_px, tables, k, root)
        except Exception:
            pass
        tilt = _median(slopes) if len(slopes) >= 3 else 0.0
        if abs(tilt) < 0.004:
            tilt = 0.0
        cx, cy = iw / 2.0, ih / 2.0
        ct, st = 1.0 / math.sqrt(1 + tilt * tilt), tilt / math.sqrt(1 + tilt * tilt)

        def turn(x, y):
            dx, dy = x - cx, y - cy
            return cx + dx * ct + dy * st, cy - dx * st + dy * ct
        # short labels standing beside or under a picture (an axis's numbers, a legend) belong to it
        margin = 0.42 * iw / 8.27
        for fi, f in enumerate(figs_px):
            f = list(f)
            for pid, lb, slope, off, xs, words in lines:
                if len(words) > 3:
                    continue
                mx, my = (lb[0] + lb[2]) / 2.0, (lb[1] + lb[3]) / 2.0
                beside = f[1] - 4 <= my <= f[3] + 4 and (0 <= f[0] - lb[2] <= margin or 0 <= lb[0] - f[2] <= margin)
                under = f[0] - 4 <= mx <= f[2] + 4 and 0 <= lb[1] - f[3] <= 0.35 * margin and (lb[2] - lb[0]) < 0.3 * (f[2] - f[0])
                if beside or under:
                    f[0], f[1], f[2], f[3] = min(f[0], lb[0] - 2), min(f[1], lb[1] - 2), max(f[2], lb[2] + 2), max(f[3], lb[3] + 2)
            figs_px[fi] = tuple(f)
        # How thick the letters' strokes are, word by word (bold letters' stems are half as thick again): the runs of
        # ink along three rows through the middle of the small letters, as a share of the line's letter height.
        stroke = {}
        try:
            thr_i = max(50, min(185, paper - 60))
            ink = clean.point(lambda v: 255 if v < thr_i else 0)
            for pid, lb, slope, off, xs, words in lines:
                if xs < 10:
                    continue
                for (t, wb, conf) in words:
                    if sum(1 for ch in t if ch.isalpha()) < 3:
                        continue
                    base_y = lb[3] + off + slope * (wb[0] - lb[0])
                    runs = []
                    for f_ in (0.14, 0.22, 0.30):
                        yy = int(round(base_y - f_ * xs))
                        if 0 <= yy < ih:
                            seg = ink.crop((wb[0], yy, wb[2], yy + 1)).tobytes()
                            runs += [len(r_) for r_ in seg.split(b"\x00") if r_]
                    if len(runs) >= 4:
                        stroke[id(wb)] = _median(runs) / float(xs)
        except Exception:
            stroke = {}
        thick = {}                                           # the same, by the word as the page holds it
        rows = []
        others = {}                                          # a row's second measure of its letters (Tesseract's own)
        for pid, lb, slope, off, xs, words in lines:
            if any(f[0] <= (lb[0] + lb[2]) / 2.0 <= f[2] and f[1] <= (lb[1] + lb[3]) / 2.0 <= f[3] for f in figs_px):
                continue                                     # "words" read in a picture are the picture
            text = "".join(w[0] for w in words)
            # From the letters' tops to the baseline. A line that RISES to the right (a page photographed a little
            # turned) has a box as much taller as it rises: taken as letter height, a long line came out twice the
            # size of a short one and the lines stood on each other.
            asc_h = (lb[3] - lb[1]) + off + min(0.0, slope) * (lb[2] - lb[0])
            # Tesseract's own measure of the line's letters (from the bottom of a "p" to the top of a "d": nine tenths
            # of the size for a book face). The line's box is the finer measure where it is the letters' box, and it
            # is not always: a rule's edge left under a table's row, a speck over a word, and the box is too tall.
            by_x = xs * k * 1.09
            other = None
            if any(ch.isupper() or ch.isdigit() or ch in _ASCENDERS for ch in text) and asc_h > 4:
                size = asc_h * k / 0.70
                if id(lb) not in no_xs:
                    if not (0.72 * by_x <= size <= 1.1 * by_x):
                        # the box is taller than the letters (a raised figure, a root sign, brackets, a speck):
                        # Tesseract's measure of the row stands
                        size = by_x
                    else:
                        other = max(4.0, min(120.0, by_x))   # near: the page's common sizes decide between the two
            else:
                size = by_x
            size = max(4.0, min(120.0, size))
            mid = (lb[0] + lb[2]) / 2.0
            by = lb[3] + off + slope * (mid - lb[0])
            _bx, base = turn(mid, by)
            r = _Row()
            r.words, r.sx = [], None
            xh = max(3.0, (lb[3] - lb[1]) * 0.5)
            for wi, (t, wb, conf) in enumerate(words):
                x0, y0 = turn(wb[0], (wb[1] + wb[3]) / 2.0)
                x1, _y1 = turn(wb[2], (wb[1] + wb[3]) / 2.0)
                hh = (wb[3] - wb[1]) / 2.0
                if wi == 0 and len(words) > 1 and len(t) == 1 and (wb[3] - wb[1]) <= 0.9 * xh and (wb[2] - wb[0]) <= 0.9 * xh \
                        and words[1][1][0] - wb[2] > 1.2 * xh and abs((wb[1] + wb[3]) / 2.0 - (lb[1] + lb[3] + off) / 2.0) <= 0.5 * xh \
                        and (conf < 80 or not t.isalnum()):
                    t = "\u2022"                             # a small mark standing clear before the words: a bullet, whatever it was read as
                w = Wd()
                w.x0, w.x1, w.top, w.bottom = x0 * k, x1 * k, (y0 - hh) * k, (y0 + hh) * k
                w.base, w.size = base * k, size
                w.parts = [[t, (None, size, False, False, None, "")]]
                w.space, w.gapspace, w.blk, w.art, w.ul, w.st, w.link, w.field, w.spw, w.lead = bool(r.words), False, (pid, "P?"), False, False, False, None, None, 0.0, False
                w.xs = None
                if id(wb) in stroke:
                    thick[id(w)] = stroke[id(wb)]
                r.words.append(w)
            r.words.sort(key=lambda w: w.x0)
            r.words[0].space = False
            r.base, r.size = base * k, size
            r.top, r.bottom = min(w.top for w in r.words), max(w.bottom for w in r.words)
            r.x0, r.x1 = r.words[0].x0, max(w.x1 for w in r.words)
            if other is not None:
                others[id(r)] = other
            rows.append(r)
        # Tesseract reads a table cell by cell and columns one after the other: lines that share a baseline are one
        # row of the page again
        rows.sort(key=lambda r: (r.base, r.x0))
        merged = []
        for r in rows:
            hit = None
            for q in reversed(merged[-6:]):
                if abs(q.base - r.base) <= 0.3 * min(q.size, r.size) and abs(q.size - r.size) <= 0.25 * max(q.size, r.size) and (r.x0 >= q.x1 - 1 or r.x1 <= q.x0 + 1):
                    hit = q
                    break
            if hit is None:
                merged.append(r)
            else:
                hit.words = sorted(hit.words + r.words, key=lambda w: w.x0)
                for i, w in enumerate(hit.words):
                    w.space = i > 0
                hit.top, hit.bottom = min(hit.top, r.top), max(hit.bottom, r.bottom)
                hit.x0, hit.x1 = hit.words[0].x0, max(w.x1 for w in hit.words)
        rows = merged
        if not rows:
            return False
        # letters of one size were read a little differently from line to line (by a twentieth either way, and more
        # on a photographed page): the common sizes win, and sizes within a ninth of each other are one size
        sizes = Counter()
        for r in rows:
            sizes[round(r.size * 2) / 2.0] += sum(len(w.text()) for w in r.words)
        common = [sz for sz, c in sizes.most_common()]
        chosen = []
        for sz in common:
            if not any(abs(sz - c) <= 0.11 * c for c in chosen):
                chosen.append(sz)
        def set_size(r, sz):
            r.size = sz
            for w in r.words:
                w.size = sz
                f = w.parts[0][1]
                w.parts[0][1] = (f[0], sz, f[2], f[3], f[4], f[5])
        for r in rows:
            best = min(chosen, key=lambda c: abs(c - r.size))
            if abs(best - r.size) <= 0.11 * best:
                sz = best
            else:
                # the line's box gave a size the page does not know: does Tesseract's own measure give one it knows?
                alt = others.get(id(r))
                b2 = min(chosen, key=lambda c: abs(c - alt)) if alt else None
                sz = b2 if b2 is not None and abs(b2 - alt) <= 0.11 * b2 else round(r.size * 2) / 2.0
            set_size(r, round(sz * 2) / 2.0)
        # A line among evenly pitched lines of one size has their size (one paragraph is set in one size; the measure
        # of a single line is off by a tenth now and then, and Word then breaks that line elsewhere).
        rows.sort(key=lambda r: r.base)
        for i, r in enumerate(rows):
            up = dn = None
            for q in reversed(rows[max(0, i - 8):i]):
                if min(q.x1, r.x1) - max(q.x0, r.x0) >= 0.4 * min(q.x1 - q.x0, r.x1 - r.x0) and r.base - q.base > 0.5 * r.size:
                    up = q
                    break
            for q in rows[i + 1:i + 9]:
                if min(q.x1, r.x1) - max(q.x0, r.x0) >= 0.4 * min(q.x1 - q.x0, r.x1 - r.x0) and q.base - r.base > 0.5 * r.size:
                    dn = q
                    break
            if up is None or dn is None or up.size != dn.size or r.size == up.size:
                continue
            d1, d2 = r.base - up.base, dn.base - r.base
            if abs(d1 - d2) <= 0.1 * max(d1, d2) and max(d1, d2) <= 1.7 * up.size and abs(r.size - up.size) <= 0.22 * up.size:
                set_size(r, up.size)
        rows.sort(key=lambda r: r.base)
        # A run of three lines or more, one under the other at one pitch, is one paragraph or one list: a line of it
        # read smaller (by up to a fifth: a line in italics, a line without tall letters) has the run's size. (Measured
        # on a made-up phone photograph, 20 Nov 2026: a paragraph's last line and two points of a list of four came out
        # smaller, each a footnote's size; the rule above needs a line of the size above AND below, at one pitch.)
        # Only a smaller line is raised: a larger one may be a heading set close above its words.
        def over_(a_, b_):
            return min(a_.x1, b_.x1) - max(a_.x0, b_.x0) >= 0.4 * min(a_.x1 - a_.x0, b_.x1 - b_.x0)
        i = 0
        while i < len(rows):
            run, j, p0 = [rows[i]], i + 1, None
            while j < len(rows):
                q, last = rows[j], run[-1]
                d = q.base - last.base
                if not over_(q, last) or d <= 0.5 * max(q.size, last.size) or d > 1.7 * max(q.size, last.size):
                    break
                if p0 is not None and abs(d - p0) > 0.1 * p0:
                    break
                p0 = d if p0 is None else p0
                run.append(q)
                j += 1
            if len(run) >= 3:
                weight = Counter()
                for r in run:
                    weight[r.size] += sum(len(w.text()) for w in r.words)
                dom = weight.most_common(1)[0][0]
                for r in run:
                    if r.size < dom and dom - r.size <= 0.22 * dom:
                        set_size(r, dom)
                i = j
            else:
                i += 1
        # Rows of a table without ruled lines (words under words in three columns or more) are of one size: a row
        # of single figures is measured poorly by itself.
        def cols_of(r):
            out = [r.words[0].x0]
            for a_, b_ in zip(r.words, r.words[1:]):
                if b_.x0 - a_.x1 > 1.8 * r.size:
                    out.append(b_.x0)
            return out
        for a_, b_ in zip(rows, rows[1:]):
            if a_.size == b_.size or b_.base - a_.base > 2.6 * max(a_.size, b_.size) or abs(a_.size - b_.size) > 0.3 * max(a_.size, b_.size):
                continue
            ca, cb = cols_of(a_), cols_of(b_)
            if len(ca) >= 3 and len(cb) >= 3:
                same = sum(1 for x in ca if any(abs(x - y) <= 0.6 * a_.size for y in cb))
                if same >= 0.6 * min(len(ca), len(cb)):
                    na, nb = sum(len(w.text()) for w in a_.words), sum(len(w.text()) for w in b_.words)
                    set_size(a_ if na < nb else b_, (b_ if na < nb else a_).size)
        # Bold is settled when every page has been read (_ocr_bold): a cover page is all bold, and has no plain words
        # of its own to go by.
        if thick:
            doc.setdefault("ocr_bold", []).append(([(list(r.words), r.size) for r in rows], thick, chosen[0] if chosen else None))
        # A list's number stands clear of its words by more than a space. A read page has no letters' widths to tell
        # a space by: the usual gap between this page's words is the measure.
        gaps = sorted(b_.x0 - a_.x1 for r in rows for a_, b_ in zip(r.words, r.words[1:]) if 0 < b_.x0 - a_.x1 < 1.2 * r.size)
        usual = gaps[len(gaps) // 2] if len(gaps) >= 12 else None
        if usual:
            for r in rows:
                if len(r.words) >= 2 and _NUM_MARK.match(r.words[0].text()) and r.words[1].x0 - r.words[0].x1 >= 1.7 * usual:
                    r.words[1].gapspace = True
        # The word a row of dots leads to (a contents page's page number) is reached by a dotted tab, as on a page
        # that was never a picture. A few dots left at either end of the row belong to the row, not to the words.
        for (lx0, lx1, ly, lh) in leaders:
            a0, a1, ay = lx0 * k, lx1 * k, ly * k
            row = None
            for r in rows:
                if abs(r.base - ay) <= 0.5 * r.size and r.x0 < a0 + r.size and (row is None or len(r.words) > len(row.words)):
                    row = r
            if row is None:
                continue
            after = [w for w in row.words if w.x0 >= a1 - 0.5 * row.size]
            if not after:
                # read as a line of its own (a lone figure's size is measured poorly): it is this row's last word
                for q in rows:
                    if q is not row and abs(q.base - ay) <= 0.6 * row.size and q.x0 >= a1 - 0.5 * row.size and q.x0 <= a1 + 6 * row.size:
                        for w in q.words:
                            w.size, w.base = row.size, row.base
                            f = w.parts[0][1]
                            w.parts[0][1] = (f[0], row.size, f[2], f[3], f[4], f[5])
                        row.words = sorted(row.words + q.words, key=lambda w: w.x0)
                        q.words = []
                        after = [w for w in row.words if w.x0 >= a1 - 0.5 * row.size]
                        break
                rows = [q for q in rows if q.words]
            before = [w for w in row.words if w.x1 <= a0 + 0.5 * row.size]
            if before and before[-1].text().endswith(".."):
                before[-1].parts[0][0] = before[-1].text().rstrip(".")
            if after:
                t0 = after[0].text().lstrip(".")
                if t0:
                    after[0].parts[0][0] = t0
                after[0].lead = True
                after[0].space = True
            row.words = [w for w in row.words if w.text()]
            if row.words:
                row.x0, row.x1 = row.words[0].x0, max(w.x1 for w in row.words)
        rows = [r for r in rows if r.words]
        # a bullet read as a letter ("e", "o") between lines that begin with a bullet at the same place is a bullet
        for i, r in enumerate(rows):
            w0 = r.words[0]
            if len(r.words) < 2 or len(w0.text()) != 1 or w0.text() in _BULLETS or r.words[1].x0 - w0.x1 < 0.25 * r.size:
                continue
            near = [q for q in rows[max(0, i - 2):i] + rows[i + 1:i + 3]
                    if len(q.words) >= 2 and q.words[0].text() == "\u2022" and abs(q.words[0].x0 - w0.x0) <= 0.35 * r.size
                    and abs(q.words[1].x0 - r.words[1].x0) <= 0.35 * r.size and abs(q.size - r.size) <= 0.15 * r.size]
            if near and w0.x1 - w0.x0 <= 0.8 * r.size:
                w0.parts[0][0] = "\u2022"
        pg.rows = rows
        pg.w, pg.h = pw, ph
        pg.skew = skew
        figs = [(f[0] * k, f[1] * k, f[2] * k, f[3] * k) for f in figs_px]
        if tables:
            # the page's own drawn lines (a scan has none that mean anything) give way to the ones found in its picture
            pg.hs, pg.vs, pg.fills = t_hs, t_vs, t_fills
            for tb in tables:
                _table_looks(tb, pg)
            for tb in tables:
                _cell_take(tb, pg)
                sizes_t = Counter()
                for ce in tb.cells:
                    # words on a dark row are light: they keep to be read on it (their strokes could not be measured
                    # there, and a mark with no letter in it is the picture's grain)
                    if ce.shade and _lum(ce.shade) < 0.4:
                        keep_r = []
                        for r in ce.items:
                            r.words = [w for w in r.words if any(ch.isalnum() for ch in w.text())]
                            if not r.words:
                                continue
                            r.words[0].space = False
                            r.x0, r.x1 = r.words[0].x0, max(w.x1 for w in r.words)
                            for w in r.words:
                                f = w.parts[0][1]
                                w.parts[0][1] = (f[0], f[1], f[2], f[3], "FFFFFF", f[5])
                            keep_r.append(r)
                        ce.items = keep_r
                    for r in ce.items:
                        sizes_t[r.size] += sum(len(w.text()) for w in r.words)
                # one table is set in one size (a short line's measure is off now and then)
                if sizes_t:
                    main_t = sizes_t.most_common(1)[0][0]
                    for ce in tb.cells:
                        for r in ce.items:
                            if r.size != main_t and 0.75 * main_t <= r.size <= 1.3 * main_t:
                                set_size(r, main_t)
            rest = []
            for f in figs:
                tb = next((t_ for t_ in tables if t_.x0 - 2 <= f[0] and f[2] <= t_.x1 + 2 and t_.top - 2 <= f[1] and f[3] <= t_.bottom + 2), None)
                if tb is None:
                    rest.append(f)
                elif f[2] - f[0] >= 14 and f[3] - f[1] >= 14:
                    _cell_pic(tb, f, pg)                     # a picture in a cell of the table
            figs = rest
            pg.tables = tables
        pg.figs = figs
        pg.nchars = sum(len(w.text()) for r in pg.rows for w in r.words) + sum(len(w.text()) for tb in tables for ce in tb.cells for r in ce.items for w in r.words)
        return True
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ---------------------------------------------------------------------------------------------- headers and footers
def _plan_head_foot(pages, doc):
    """Each page's header and footer as paragraphs; which pages share them; which numbers are the page's number."""
    any_hf = False
    for pg in pages:
        key = (int(round(pg.w)), int(round(pg.h)))
        L, R, T, B = doc["boxes"][key]
        pg.hp = _paragraphs([_make_line(r, doc) for r in pg.head], L, R, doc) if pg.head else []
        pg.fp = _paragraphs([_make_line(r, doc) for r in pg.foot], L, R, doc) if pg.foot else []
        for rules, items in ((pg.hrules, pg.hp), (pg.frules, pg.fp)):
            for q in rules:                              # the line under a running title, over a page number
                ru = Rule()
                ru.x0, ru.x1, ru.top, ru.bottom, ru.width, ru.color = q[0], q[1], q[2], q[2], q[3], q[4]
                ru.brk = ru.colbrk = False
                ru.sect = None
                ru.before = ru.after = 0.0
                ru.L, ru.R = min(L, q[0]), max(R, q[1])
                items.append(ru)
            if rules:
                items.sort(key=lambda it: it.top)
        for figs, pics in ((pg.hfigs, pg.hpics), (pg.ffigs, pg.fpics)):
            del pics[:]
            for f in figs:                               # a logo: it stands where it stands, whatever the words do
                pc = Pic()
                pc.x0, pc.top, pc.x1, pc.bottom = f[0], f[1], f[2], f[3]
                pc.data = pc.ext = pc.size = None
                pc.behind = False
                pc.kids = pc.flow = None
                pc.anchor = (f[0], f[1])
                pc.brk = pc.colbrk = False
                pc.sect = None
                pc.before = pc.after = pc.left = 0.0
                pc.page, pc.align, pc.L, pc.R = pg.n, "left", L, R
                pics.append(pc)
        if pg.hp:
            _space(pg.hp, _item_edges(pg.hp[0], doc)[0], doc)
        if pg.fp:
            _space(pg.fp, _item_edges(pg.fp[0], doc)[0], doc)
        if getattr(pg, "ocr", False):
            # a page read from a picture: its head's and foot's words (one row or two, the reader may part them), and
            # how many pictures and lines; where they stand differs from photograph to photograph
            pg.sig = ("ocr", " ".join(_norm_hf(_row_text(r)) for r in pg.head), " ".join(_norm_hf(_row_text(r)) for r in pg.foot),
                      len(pg.hfigs) + len(pg.ffigs), len(pg.hrules) + len(pg.frules))
        else:
            pg.sig = (tuple(_norm_hf(_row_text(r)) for r in pg.head), tuple(_norm_hf(_row_text(r)) for r in pg.foot),
                      tuple((int(round(f[0] / 3.0)), int(round(f[1] / 3.0)), int(round(f[2] / 3.0)), int(round(f[3] / 3.0))) for f in pg.hfigs + pg.ffigs),
                      tuple((int(round(q[0] / 3.0)), int(round(q[1] / 3.0)), int(round(q[2] / 2.0))) for q in pg.hrules + pg.frules))
        pg.hfset = 0
        if pg.head or pg.foot or pg.hfigs or pg.ffigs or pg.hrules or pg.frules:
            any_hf = True
    doc["hfsets"], doc["sect_pages"], doc["evenodd"], doc["hf_reps"] = [], set(), False, set()
    if not any_hf:
        return
    n = len(pages)
    sigs = [pg.sig for pg in pages]

    def same(a, b):
        if a == b:
            return True
        if a[0] != "ocr" or b[0] != "ocr" or a[3:] != b[3:]:
            return False
        return all(x == y or _alike(x, y) for x, y in ((a[1], b[1]), (a[2], b[2])))
    def plan(same):
        doc["sect_pages"], doc["evenodd"] = set(), False
        for pg in pages:
            pg.hfset = 0
            for zone in ("head", "foot"):
                for r in getattr(pg, zone):
                    for w in r.words:
                        if getattr(w, "field", None) in ("PAGE", "NUMPAGES"):
                            w.field = None
        sets = []
        if all(same(s, sigs[0]) for s in sigs):
            sets.append({"default": list(range(n))})
        elif n >= 2 and all(same(s, sigs[1]) for s in sigs[1:]):
            sets.append({"default": list(range(1, n)), "first": [0]})
        elif n >= 4 and all(same(s, sigs[0]) for s in sigs[0::2]) and all(same(s, sigs[1]) for s in sigs[1::2]):
            sets.append({"default": list(range(0, n, 2)), "even": list(range(1, n, 2))})
            doc["evenodd"] = True
        elif n >= 5 and all(same(s, sigs[1]) for s in sigs[1::2]) and all(same(s, sigs[2]) for s in sigs[2::2]):
            # a first page of its own, then left and right pages (a journal, a book's chapter)
            sets.append({"first": [0], "even": list(range(1, n, 2)), "default": list(range(2, n, 2))})
            doc["evenodd"] = True
        else:
            i = 0
            while i < n:
                j = i
                while j + 1 < n and same(sigs[j + 1], sigs[i]):
                    j += 1
                sets.append({"default": list(range(i, j + 1))})
                for k in range(i, j + 1):
                    pages[k].hfset = len(sets) - 1
                if i > 0:
                    doc["sect_pages"].add(i)
                i = j + 1
        for st in sets:
            for kind in ("default", "first", "even"):
                idx = st.get(kind)
                if not idx:
                    continue
                rep = pages[idx[0]]
                for zone in ("head", "foot"):
                    rows = getattr(rep, zone)
                    for ri, r in enumerate(rows):
                        for wi, w in enumerate(r.words):
                            t = w.text()
                            if not t.isdigit() or len(idx) < 2:
                                continue
                            vals = []
                            for k in idx:
                                rr = getattr(pages[k], zone)
                                try:
                                    vals.append(int(rr[ri].words[wi].text()))
                                except Exception:
                                    vals = None
                                    break
                            if not vals:
                                continue
                            if all(v - (k + 1) == vals[0] - (idx[0] + 1) for v, k in zip(vals, idx)):
                                w.field = "PAGE"
                                if kind == "default" or "start" not in st:
                                    st["start"] = vals[0] - idx[0]
                            elif all(v == vals[0] for v in vals) and vals[0] == n and wi > 0 and r.words[wi - 1].text().lower() in ("of", "/", "sur", "de", "von"):
                                w.field = "NUMPAGES"
        return sets

    sets = plan(same)

    # A number in the head or the foot that changes from page to page and is not the page's number (a chapter's, a
    # week's, a date; or a page number read wrong from a picture) would stand on every page as the first page of the
    # set has it. Then the pages share a head only where it is the same word for word, and each keeps its own numbers.
    def varies(st):                           # the zones (head, foot) where such a number is
        out = set()
        for kind in ("default", "first", "even"):
            idx = st.get(kind)
            if not idx or len(idx) < 2:
                continue
            rep = pages[idx[0]]
            for zone in ("head", "foot"):
                for ri, r in enumerate(getattr(rep, zone)):
                    for wi, w in enumerate(r.words):
                        t = w.text()
                        if zone in out or not re.search(r"\d", t) or getattr(w, "field", None):
                            continue
                        for k in idx[1:]:
                            try:
                                if getattr(pages[k], zone)[ri].words[wi].text() != t:
                                    out.add(zone)
                                    break
                            except Exception:
                                out.add(zone)
                                break
        return out
    vz = set()
    for st in sets:
        vz |= varies(st)
    if vz:
        # where the number is, the words must be the same, number and all; the other zone is compared as before (its
        # page number then still becomes Word's page number within each run of pages)
        def zone_sig(pg, zone):
            rows = getattr(pg, zone)
            if zone in vz:
                return ("=",) + tuple(_row_text(r) for r in rows)
            return tuple(_norm_hf(_row_text(r)) for r in rows)
        sigs = [(zone_sig(pg, "head"), zone_sig(pg, "foot"), pg.sig[3:] if pg.sig[0] == "ocr" else pg.sig[2:]) for pg in pages]
        sets = plan(lambda a_, b_: a_ == b_)
    # Word gives the "even" head to pages whose NUMBER is even: where the pages are numbered from an even number on, the
    # two heads change places
    for st in sets:
        if st.get("even") and st.get("default") and st.get("start") is not None and (st["start"] + st["default"][0]) % 2 == 0:
            st["default"], st["even"] = st["even"], st["default"]
    doc["hfsets"] = sets
    doc["hf_reps"] = set(st[kind][0] for st in sets for kind in ("default", "first", "even") if st.get(kind))


def _hf_dist(pg, doc):
    """(distance of the header's first line from the paper's top, of the footer's last line from its foot)."""
    hd = fd = None
    if pg.hp:
        hd = max(0.0, _item_edges(pg.hp[0], doc)[0])
    if pg.fp:
        fd = max(0.0, pg.h - _item_edges(pg.fp[-1], doc)[1])
    if hd is None and pg.hpics:
        hd = 0.0                                         # only a picture up there: its place is its own, the head is empty
    if fd is None and pg.fpics:
        fd = 0.0
    return hd, fd


# ------------------------------------------------------------------------------------------------------ pictures
def _raster(src_path, pno, dpi, w_pt, h_pt):
    """One page drawn as a picture (Poppler; PDFium when Poppler is not there)."""
    longest = max(w_pt, h_pt)
    dpi = int(max(50, min(dpi, 3400 * 72.0 / max(1.0, longest))))
    tmp = tempfile.mkdtemp(prefix="p2w_")
    try:
        root = os.path.join(tmp, "p")
        try:
            subprocess.run(["pdftoppm", "-r", str(dpi), "-f", str(pno + 1), "-l", str(pno + 1), "-png", "-singlefile", src_path, root],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=90, check=False)
        except Exception:
            pass
        f = root + ".png"
        if os.path.exists(f) and os.path.getsize(f) > 0:
            with Image.open(f) as im:
                return im.convert("RGB")
        try:
            import pypdfium2 as pdfium
            with _PDFIUM_LOCK:
                pdf = pdfium.PdfDocument(src_path)
                try:
                    return pdf[pno].render(scale=dpi / 72.0).to_pil().convert("RGB")
                finally:
                    pdf.close()
        except Exception:
            return None
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


import threading as _threading
_PDFIUM_LOCK = _threading.Lock()


def _encode(im):
    """A cut-out as the smaller of what suits it: PNG for flat drawing, JPEG for a photograph."""
    w, h = im.size
    small = im.resize((max(1, min(w, 96)), max(1, min(h, 96))))
    colours = len(small.getcolors(small.size[0] * small.size[1] + 1) or ())
    buf = io.BytesIO()
    if colours > 0.45 * small.size[0] * small.size[1] and colours > 1500:
        im.save(buf, "JPEG", quality=88, optimize=True)
        return buf.getvalue(), "jpg"
    im.save(buf, "PNG", optimize=True)
    data = buf.getvalue()
    if len(data) > 900000:
        b2 = io.BytesIO()
        im.save(b2, "JPEG", quality=90, optimize=True)
        if b2.tell() < 0.5 * len(data):
            return b2.getvalue(), "jpg"
    return data, "png"


def _all_pics(body, pages):
    out = []

    def walk(items):
        for it in items:
            if isinstance(it, Pic):
                out.append(it)
                out.extend(it.kids or [])
            elif isinstance(it, Tbl):
                for ce in it.cells:
                    walk(ce.items)
    walk(body)
    for pg in pages:
        walk(getattr(pg, "hp", None) or [])
        walk(getattr(pg, "fp", None) or [])
    return out


def _hf_pics(pages, doc):
    """The pictures of the page heads and feet that the Word file will hold (one of each kind, not every page's)."""
    out = []
    for k in sorted(doc.get("hf_reps") or ()):
        out.extend(pages[k].hpics)
        out.extend(pages[k].fpics)
    return out


_RENDER_STEPS = (1.0, 0.8, 0.64, 0.5)


def _cost(q):
    """What drawing a page costs at q times its best resolution, beside its cost at the best (measured with Poppler:
    a part of the work does not depend on the picture's size)."""
    return 0.1 + 0.9 * q * q


class _Pace(object):
    """Keeps the drawing of pictures inside the time the request has. Every picture is drawn at its best resolution
    while that lets the rest be drawn in time; when it would not, the rest are drawn a step smaller: a long PDF then
    ends as a whole file with slightly softer pictures, not as "too long" after minutes of waiting. And where even
    the smallest step cannot make it, that is said at once."""
    __slots__ = ("t0", "left", "norm", "n")

    def __init__(self, t0, total):
        self.t0, self.left, self.norm, self.n = t0, max(0, int(total)), 0.0, 0

    def q(self):
        if self.n == 0:
            return 1.0
        per = self.norm / self.n                         # the seconds one drawing takes at its best resolution
        room = self.t0 + RENDER_END - time.time()
        for q in _RENDER_STEPS:
            if per * _cost(q) * self.left <= room:
                return q
        if per * _cost(_RENDER_STEPS[-1]) * self.left > room + RENDER_SLACK:
            raise ConversionError(TOO_LONG)
        return _RENDER_STEPS[-1]

    def took(self, q, seconds):
        self.norm += max(0.0, seconds) / _cost(q)
        self.n += 1
        self.left = max(0, self.left - 1)


_PAINT_OPS = frozenset((b"f", b"F", b"f*", b"B", b"B*", b"b", b"b*", b"S", b"s", b"sh", b"Do", b"BI", b"EI", b"INLINE IMAGE"))


def _paints(page, reader):
    """Does this page (its words already taken out) draw anything at all? True when unsure."""
    try:
        from pypdf.generic import ContentStream
        for a in (page.get("/Annots") or []):
            a = a.get_object()
            if a.get("/Subtype") != "/Link" and a.get("/AP") is not None:
                return True                                  # a stamp, a filled-in field: it is drawn with the page
        cont = page.get_contents()
        if cont is None:
            return False
        for _operands, op in ContentStream(cont, reader).operations:
            if op in _PAINT_OPS:
                return True
        return False
    except Exception:
        return True


def _strip_words(src_path, page_numbers, blank=None):
    """A copy of some of the PDF's pages with their words taken out: what is drawn behind the words.
    Returns (the copy's path, {page number: its place in the copy}); (None, {}) if it cannot be made.
    blank, a set, is given the numbers of the pages that draw nothing but their words (no need to draw those)."""
    try:
        from pypdf import PdfWriter
        rd = _open_reader(src_path)
        wr = PdfWriter()
        idx = {}
        for k, n in enumerate(page_numbers):
            wr.add_page(rd.pages[n])
            idx[n] = k
        wr.remove_text()
        if blank is not None:
            for n, k in idx.items():
                if not _paints(wr.pages[k], wr):
                    blank.add(n)
        out = os.path.join(tempfile.mkdtemp(prefix="p2w_bg_"), "bg.pdf")
        with open(out, "wb") as fh:
            wr.write(fh)
        return out, idx
    except Exception:
        return None, {}


def _is_blank(im):
    """Nothing but white paper?"""
    small = im.convert("L").resize((max(1, min(im.width, 160)), max(1, min(im.height, 160))))
    return small.getextrema()[0] >= 250


def _back_image(src_path, bg, idx, pg, dpi):
    """The drawing behind a designed page's words: the page drawn without its words. Words that are turned on their
    side cannot be words in the new file, so they stay in the drawing, as they look."""
    im = _raster(bg, idx[pg.n], dpi, pg.w, pg.h)
    if im is None or not pg.rot:
        return im
    full = _raster(src_path, pg.n, dpi, pg.w, pg.h)
    if full is None or full.size != im.size:
        return im
    sx, sy = im.width / pg.w, im.height / pg.h
    for (x0, top, x1, bottom) in pg.rot:
        box = (max(0, int(math.floor((x0 - 0.5) * sx))), max(0, int(math.floor((top - 0.5) * sy))),
               min(im.width, int(math.ceil((x1 + 0.5) * sx))), min(im.height, int(math.ceil((bottom + 0.5) * sy))))
        if box[2] > box[0] and box[3] > box[1]:
            im.paste(full.crop(box), box[:2])
    full.close()
    return im


def _render_backs(src_path, backs, pages, pace):
    """The drawing behind the words of each designed page (pc.behind): the page drawn without its words."""
    if not backs:
        return
    bg, idx = _strip_words(src_path, sorted(set(pc.page for pc in backs)))
    try:
        for pc in backs:
            pg = pages[pc.page]
            if bg is None or pc.page not in idx:
                pace.took(1.0, 0.0)
                continue
            q, t = pace.q(), time.time()
            im = _back_image(src_path, bg, idx, pg, int(BACK_DPI * q))
            if im is not None:
                try:
                    if not _is_blank(im):
                        pc.data, pc.ext = _encode(im)
                except Exception:
                    pc.data = None
                im.close()
            pace.took(q, time.time() - t)
    finally:
        if bg:
            shutil.rmtree(os.path.dirname(bg), ignore_errors=True)


def _render_pictures(src_path, body, pages, doc, t0):
    by_page = defaultdict(list)
    backs = []
    for pc in _all_pics(body, pages) + _hf_pics(pages, doc):
        if pc.behind:
            backs.append(pc)
        elif pc.data is None:
            by_page[pc.page].append(pc)
    pace = _Pace(t0, len(backs) + len(by_page))
    _render_backs(src_path, backs, pages, pace)
    for pno in sorted(by_page):
        pg = pages[pno]
        q, t = pace.q(), time.time()
        dpi = int(PIC_DPI * q)
        im = _raster(src_path, pno, dpi if not pg.ocr else max(100, int(dpi * pg.w / max(1.0, pg.ow))), pg.ow, pg.oh)
        if im is None:
            pace.took(q, time.time() - t)
            continue
        turn = getattr(pg, "skew", 0.0) if pg.ocr else 0.0
        if turn:
            try:
                edge = im.resize((1, 1), Image.BOX).getpixel((0, 0))
                im = im.rotate(turn, resample=Image.BICUBIC, fillcolor=edge)     # as the page was set level to be read
            except Exception:
                pass
        sx = im.width / pg.w
        sy = im.height / (pg.oh * (pg.w / pg.ow)) if pg.ocr else im.height / pg.h
        for pc in by_page[pno]:
            box = (max(0, int(math.floor((pc.x0 - 0.25) * sx))), max(0, int(math.floor((pc.top - 0.25) * sy))),
                   min(im.width, int(math.ceil((pc.x1 + 0.25) * sx))), min(im.height, int(math.ceil((pc.bottom + 0.25) * sy))))
            if box[2] - box[0] < 2 or box[3] - box[1] < 2:
                continue
            try:
                pc.data, pc.ext = _encode(im.crop(box))
            except Exception:
                pc.data = None
        im.close()
        pace.took(q, time.time() - t)


# ------------------------------------------------------------------------------------------- the document itself
def _styles(d, doc):
    """The new document's own defaults are the PDF's body text: its font and size, single lines, no room after."""
    st = d.styles.element
    dd = st.find(qn("w:docDefaults"))
    if dd is None:
        dd = OxmlElement("w:docDefaults")
        st.insert(0, dd)
    for ch in list(dd):
        dd.remove(ch)
    rd = OxmlElement("w:rPrDefault")
    rpr = OxmlElement("w:rPr")
    f = doc["dflt"]["font"]
    rpr.append(_el("w:rFonts", ascii=f, hAnsi=f, eastAsia=f, cs=f))
    rpr.append(_el("w:sz", val=int(round(doc["dflt"]["size"] * 2))))
    rpr.append(_el("w:szCs", val=int(round(doc["dflt"]["size"] * 2))))
    rpr.append(_el("w:lang", val="en-US", eastAsia="en-US", bidi="ar-SA"))
    rd.append(rpr)
    dd.append(rd)
    pd = OxmlElement("w:pPrDefault")
    ppr = OxmlElement("w:pPr")
    ppr.append(_el("w:spacing", after=0, line=240, lineRule="auto"))
    pd.append(ppr)
    dd.append(pd)
    # the Normal style must not bring a font, a size or spacing of its own
    for s in st.findall(qn("w:style")):
        if s.get(qn("w:styleId")) == "Normal":
            for tag in ("w:rPr", "w:pPr"):
                e = s.find(qn(tag))
                if e is not None:
                    s.remove(e)
    se = d.settings.element
    zoom = se.find(qn("w:zoom"))
    if zoom is not None and zoom.get(qn("w:percent")) is None:
        zoom.set(qn("w:percent"), "100")          # python-docx's blank document leaves out what Word's own rules require
    compat = se.find(qn("w:compat"))
    if compat is None:
        compat = OxmlElement("w:compat")
        se.append(compat)
    for cs in compat.findall(qn("w:compatSetting")):
        if cs.get(qn("w:name")) == "compatibilityMode":
            compat.remove(cs)
    compat.append(_el("w:compatSetting", name="compatibilityMode", uri="http://schemas.microsoft.com/office/word", val=str(doc.get("compat", 15))))


def _sect_pr(sect, doc, pages, first):
    """Word's description of one section: paper, margins, columns."""
    from docx.oxml.section import CT_SectPr          # noqa: F401 (python-docx registers the class for w:sectPr)
    key = sect["paper"]
    w, h = float(key[0]), float(key[1])
    L, R, T, B = doc["boxes"][key]
    if sect.get("cols"):
        L, R = sect["cols"][0][0], sect["cols"][-1][1]
    el = OxmlElement("w:sectPr")
    if not first:
        el.append(_el("w:type", val="nextPage" if sect.get("new_page") else "continuous"))
    sz = {"w": _tw(w), "h": _tw(h)}
    if w > h:
        sz["orient"] = "landscape"
    el.append(_el("w:pgSz", **sz))
    pg = pages[sect["page"]]
    hd, fd = sect.get("hd"), sect.get("fd")
    bottom = max(0.0, h - B)
    if doc.get("pin") or sect.get("design"):
        bottom = max(min(bottom, 14.0), bottom * 0.45)
    if sect.get("top") is not None:
        T = sect["top"]
    mar = {"top": _tw(max(0.0, T)), "right": _tw(max(0.0, w - R)), "bottom": _tw(max(0.0, bottom - min(3.0, 0.25 * doc["size"]))), "left": _tw(max(0.0, L)),
           "header": _tw(hd if hd is not None else min(max(0.0, T) / 2.0, 36.0)), "footer": _tw(fd if fd is not None else min(bottom / 2.0, 36.0)), "gutter": 0}
    el.append(_el("w:pgMar", **mar))
    if sect.get("start") is not None and sect["start"] != 1:
        el.append(_el("w:pgNumType", start=int(sect["start"])))
    cols = sect.get("cols")
    if cols and len(cols) >= 2:
        ce = _el("w:cols", num=len(cols), equalWidth=0)
        for i, (a, b) in enumerate(cols):
            at = {"w": _tw(b - a)}
            if i + 1 < len(cols):
                at["space"] = _tw(max(0.0, cols[i + 1][0] - b))
            ce.append(_el("w:col", **at))
        el.append(ce)
    return el


def _fill_hf(holder, items, doc, pics=(), shift=None):
    """Put paragraphs, lines and pictures into a Word header or footer (python-docx gives it one empty paragraph: it
    goes). shift = (dl, dr, width): this section's margins are not the ones the paragraphs were measured from (a
    section of columns): how far its left margin lies left of theirs, its right one right of theirs, and its width."""
    import copy as _copy
    root = holder._element
    for ch in list(root):
        root.remove(ch)
    wrote = False
    for it in items:
        if isinstance(it, Para):
            if shift and it.lines:
                dl, dr, width = shift
                it = _copy.copy(it)
                if it.align in ("left", "both") or it.tabs:
                    it.left = it.left + dl
                    it.right = it.right + dr
                tabs = []
                for pos, kind, leader in it.tabs or []:
                    pos = pos + dl
                    if kind == "right" and pos > width - 0.1:
                        pos = width - 0.1                # a stop past the margin would throw the words onto a new line
                    tabs.append((pos, kind, leader))
                it.tabs = tabs
            root.append(_p_el(it, doc, holder.part))
            wrote = True
        elif isinstance(it, Rule):
            root.append(_rule_el(it, doc))
            wrote = True
        elif isinstance(it, Pic) and it.data:
            root.append(_pic_el(it, doc, holder))
            wrote = True
    pics = [pc for pc in pics if pc.data and pc.anchor]
    if not wrote:
        root.append(_tiny_p() if pics else OxmlElement("w:p"))
    if pics:
        first = root.find(qn("w:p"))
        for k, pc in enumerate(pics):
            try:
                first.append(_anchor_run(pc, holder, order=k + 1))
            except Exception:
                pass


_FT_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"


def _width_scales():
    """{(family, bold, italic): percent}: how much Word must widen or narrow a stand-in font's letters for them to take
    the room the PDF's letters take (a hair less: a line that is full must still fit)."""
    out = {}
    for key, (a, b, n) in _W_STAT.items():
        if n < 30 or b <= 0:
            continue
        r = a / b * 0.994
        if 0.6 <= r <= 1.6 and abs(r - 1.0) >= 0.012:
            out[key] = int(round(r * 100))
    return out


def _font_table(d, doc):
    """The Word file's list of its fonts. For a font not every computer has, it says what kind of font it is and which
    common font is nearest: Word (and LibreOffice) then puts in that one where the font is missing, instead of
    whatever it likes (a text set in a serif font came out in a wide sans-serif one, and filled more pages)."""
    used = set(doc["dflt"].get("used") or ())
    if doc["dflt"].get("font"):
        used.add(doc["dflt"]["font"])
    rows = []
    for fam in sorted(used):
        if fam in _WORD_FONTS:
            continue
        low = fam.lower()
        kind = "modern" if _MONO_NAMES.search(low) else "swiss" if _SANS_NAMES.search(low) else "roman" if _SERIF_NAMES.search(low) else None
        name = fam.replace("&", "&amp;").replace("<", "&lt;").replace('"', "&quot;")
        if kind:
            rows.append('<w:font w:name="%s"><w:altName w:val="%s"/><w:charset w:val="00"/><w:family w:val="%s"/><w:pitch w:val="%s"/></w:font>'
                        % (name, _ALT_FONT[kind], kind, "fixed" if kind == "modern" else "variable"))
        else:
            rows.append('<w:font w:name="%s"><w:charset w:val="00"/><w:family w:val="auto"/><w:pitch w:val="variable"/></w:font>' % name)
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           + "".join(rows) + "</w:fonts>").encode("utf-8")
    try:
        part = d.part.part_related_by(_FT_RT)
        part._blob = xml
    except Exception:
        pass


_BLANK_APP = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"/>')


def _drop_thumbnail(package):
    """What the blank file a Word or PowerPoint file grows from says about itself must not go out as the person's
    file's own: its little picture of an empty page (the picture a computer shows for the file), and "Microsoft
    Macintosh Word", "1 page, 0 words" (docProps/app.xml: the part stays, programs expect it, and says nothing).
    python-docx keeps the package's relationships in .rels, python-pptx in ._rels: for PowerPoint files this did
    nothing before 6.4's last round, and every PowerPoint made from a PDF carried the empty slide's picture."""
    rels = getattr(package, "rels", None)
    if rels is None:
        rels = getattr(package, "_rels", None)
    if rels is None:
        return
    try:
        for rel in list(rels.values()):
            if rel.reltype.endswith("/thumbnail"):
                rels.pop(rel.rId)
            elif rel.reltype.endswith("/extended-properties"):
                rel.target_part._blob = _BLANK_APP
    except Exception:
        pass


_FN_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
_FN_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"


def _write_notes(d, doc):
    """Word's footnotes part, with each note's paragraphs. Must run after the body is written (a note's place in the
    text is its reference there)."""
    notes = doc.get("notes_list") or []
    if not notes:
        return
    from docx.opc.part import Part
    from docx.opc.packuri import PackURI
    from lxml import etree
    part = Part(PackURI("/word/footnotes.xml"), _FN_CT, b"", d.part.package)
    d.part.relate_to(part, _FN_RT)
    root = OxmlElement("w:footnotes")
    for nid, kind in ((-1, "separator"), (0, "continuationSeparator")):
        fn = _el("w:footnote", type=kind, id=nid)
        p = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        ppr.append(_el("w:spacing", after=0, line=240, lineRule="auto"))
        p.append(ppr)
        r = OxmlElement("w:r")
        r.append(OxmlElement("w:" + kind))
        p.append(r)
        fn.append(p)
        root.append(fn)
    for n in notes:
        fn = _el("w:footnote", id=n.id)
        for k, it in enumerate(n.items):
            if not isinstance(it, Para):
                continue
            if k == 0 and n.auto and it.lines:
                # Word writes the note's own number: the mark the PDF printed there is not written twice
                it.runs = _runs(it, doc)
                if it.runs and it.runs[0][0].strip(".) \t") == n.mark:
                    f0 = it.runs[0][1]
                    it.runs[0] = ["", f0 if f0[5] else (f0[0], f0[1], f0[2], f0[3], f0[4], "sup" if n.sup else ""), False, False, None, "FNREF"]
            fn.append(_p_el(it, doc, part))
        if not len(fn):
            fn.append(OxmlElement("w:p"))
        root.append(fn)
    part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    se = d.settings.element
    pr = OxmlElement("w:footnotePr")
    if doc.get("notes_each_page"):
        pr.append(_el("w:numRestart", val="eachPage"))
    pr.append(_el("w:footnote", id=-1))
    pr.append(_el("w:footnote", id=0))
    compat = se.find(qn("w:compat"))
    if compat is not None:
        compat.addprevious(pr)
    else:
        se.append(pr)


def _write(body, pages, doc, out_path):
    from docx.section import Section
    d = Document()
    _styles(d, doc)
    bodyel = d.element.body
    final = bodyel.find(qn("w:sectPr"))
    for ch in list(bodyel):
        bodyel.remove(ch)
    sections = doc["sections"]
    sets = doc.get("hfsets") or []
    sect_els = []
    part = d.part
    holder = d._body
    last_p, last_tbl = None, False
    cur = -1

    def put(el):
        bodyel.append(el)

    def close_section():
        nonlocal last_p, last_tbl
        if last_p is None or last_tbl:
            last_p = _tiny_p()
            put(last_p)
            last_tbl = False
        ppr = last_p.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            last_p.insert(0, ppr)
        se = _sect_pr(sections[cur], doc, pages, cur == 0)
        ppr.append(se)
        sect_els.append(se)

    for it in body:
        if getattr(it, "sect", None) is not None:
            if cur >= 0:
                close_section()
            cur += 1
            sec = sections[cur]
            st = sets[sec["hf"]] if sets and 0 <= sec["hf"] < len(sets) else None
            if st:
                rep = pages[st["default"][0]]
                sec["hd"], sec["fd"] = _hf_dist(rep, doc)
                if st.get("start") is not None and (cur == 0 or sec["page"] == st["default"][0]):
                    sec["start"] = st["start"] + sec["page"] if cur else st["start"]
            last_p, last_tbl = None, False
        if isinstance(it, Para):
            el = _p_el(it, doc, part)
            put(el)
            last_p, last_tbl = el, False
        elif isinstance(it, Pic) and it.anchor is not None:
            el = _tiny_p(brk=it.brk)                    # one point of the flow holds the page's drawing, behind the words
            for k, one in enumerate([it] + list(it.kids or [])):
                if one.data:
                    try:
                        el.append(_anchor_run(one, holder, behind=one.behind, order=k + 1))
                    except Exception:
                        pass
            put(el)
            last_p, last_tbl = el, False
        elif isinstance(it, Pic):
            el = _pic_el(it, doc, holder)
            put(el)
            last_p, last_tbl = el, False
        elif isinstance(it, Rule):
            el = _rule_el(it, doc)
            put(el)
            last_p, last_tbl = el, False
        elif isinstance(it, Tbl):
            if last_tbl or it.colbrk:
                put(_tiny_p(before=it.before if last_tbl else 0.0, colbrk=it.colbrk))
            put(_tbl_el(it, doc, part, holder))
            last_tbl = True
    if last_tbl or last_p is None:
        put(_tiny_p())
    if cur < 0:
        sections.append({"paper": (int(round(pages[0].w)), int(round(pages[0].h))), "cols": None, "page": 0, "hf": 0, "new_page": False})
        cur = 0
    fe = _sect_pr(sections[cur], doc, pages, cur == 0)
    bodyel.append(fe)
    sect_els.append(fe)
    # headers and footers
    if sets:
        if doc.get("evenodd"):
            d.settings.odd_and_even_pages_header_footer = True
        for se, sec in zip(sect_els, sections):
            st = sets[sec["hf"]] if 0 <= sec["hf"] < len(sets) else None
            if st is None:
                continue
            s = Section(se, part)
            rep = pages[st["default"][0]]
            bL, bR = doc["boxes"][sec["paper"]][0], doc["boxes"][sec["paper"]][1]
            sL, sR = (sec["cols"][0][0], sec["cols"][-1][1]) if sec.get("cols") else (bL, bR)
            sh = (bL - sL, sR - bR, sR - sL)
            s.header.is_linked_to_previous = False
            _fill_hf(s.header, rep.hp, doc, rep.hpics, sh)
            s.footer.is_linked_to_previous = False
            _fill_hf(s.footer, rep.fp, doc, rep.fpics, sh)
            if st.get("first"):
                fp = pages[st["first"][0]]
                s.different_first_page_header_footer = True
                s.first_page_header.is_linked_to_previous = False
                _fill_hf(s.first_page_header, fp.hp, doc, fp.hpics, sh)
                s.first_page_footer.is_linked_to_previous = False
                _fill_hf(s.first_page_footer, fp.fp, doc, fp.fpics, sh)
            if st.get("even"):
                ep = pages[st["even"][0]]
                s.even_page_header.is_linked_to_previous = False
                _fill_hf(s.even_page_header, ep.hp, doc, ep.hpics, sh)
                s.even_page_footer.is_linked_to_previous = False
                _fill_hf(s.even_page_footer, ep.fp, doc, ep.fpics, sh)
    _write_notes(d, doc)
    _font_table(d, doc)
    _drop_thumbnail(d.part.package)
    cp = d.core_properties
    cp.author = ""
    cp.title = doc.get("title") or ""
    cp.comments = ""
    cp.keywords = ""
    cp.subject = ""
    try:
        import datetime as _dt
        now = _dt.datetime.now(_dt.timezone.utc).replace(microsecond=0, tzinfo=None)
        cp.created = now
        cp.modified = now
        cp.last_modified_by = ""
        cp.revision = 1
    except Exception:
        pass
    d.save(out_path)


# ------------------------------------------------------------------------------------------------ PowerPoint
# 6.4: PDF to PowerPoint. Every page is a slide that looks like the page: what is drawn on it (colours, shapes,
# pictures, a table's lines) is one picture at the back, and every paragraph is a text box at its own place, in its own
# font, size and colour, which can be changed. Nothing is added and nothing is left out.
E_PT = 12700                    # English Metric Units in a point


def _ppt_paras(pg):
    """The page's paragraphs, wherever they stand (in the columns, in a table's cells)."""
    out = []

    def walk(items):
        for it in items:
            if isinstance(it, Para) and it.lines:
                out.append(it)
            elif isinstance(it, Tbl):
                for ce in it.cells:
                    walk(ce.items)
    for f in pg.zones or []:
        if f[0] == "one":
            walk(f[1])
        else:
            for c in f[1]:
                walk(c)
    return out


class _Mini(object):
    """Some of a paragraph's lines, or parts of them, as _runs() wants them."""
    __slots__ = ("lines", "hard")


def _line_part(ln, i0, i1, marker):
    """Words i0..i1 of a line as a line of their own (marker: the first word is a bullet or number, a tab follows it)."""
    q = Ln()
    q.words = ln.words[i0:i1]
    q.base, q.size, q.top, q.bottom = ln.base, ln.size, ln.top, ln.bottom
    q.x0, q.x1 = q.words[0].x0, max(w.x1 for w in q.words)
    q.marker = bool(marker and len(q.words) > 1)
    q.segs = [[0, q.words[0].x0, q.words[0].x1]]
    if q.marker:
        q.segs.append([1, q.words[1].x0, q.x1])
    q.dots, q.blk, q.font, q.allbold, q.startbold, q.cen, q.rgt = set(), ln.blk, ln.font, ln.allbold, ln.startbold, ln.cen, ln.rgt
    return q


def _ppt_text(slide, lines, doc, k, ox, oy, align, pitch, R=None, just=False):
    """A text box with these lines: where they stand, in their fonts, sizes and colours."""
    from pptx.util import Emu, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
    from pptx.oxml.xmlchemy import OxmlElement as PEl
    n = len(lines)
    first = lines[0]
    s1 = max([q[1][1] for w in first.words for q in w.parts if not q[1][5]] or [first.size or 12.0])
    body_x0 = first.segs[1][1] if (first.marker and len(first.segs) > 1) else None
    mar = ind = 0.0
    if align == "center":
        bx0, bx1, algn = min(l.x0 for l in lines), max(l.x1 for l in lines), PP_ALIGN.CENTER
    elif align == "right":
        bx0, bx1, algn = min(l.x0 for l in lines), max(l.x1 for l in lines), PP_ALIGN.RIGHT
    else:
        x_first = first.x0
        x_rest = min(l.x0 for l in lines[1:]) if n >= 2 else (body_x0 if body_x0 is not None else x_first)
        bx0 = min(x_first, x_rest)
        mar, ind = x_rest - bx0, x_first - x_rest
        bx1 = max(l.x1 for l in lines)
        algn = PP_ALIGN.LEFT
        # (a paragraph set flush on both sides: each line still ends where it ends in the PDF, so no word can run into
        # the paragraph below. The lines are not spread to the right edge: tried, PowerPoint's and LibreOffice's ways of
        # spreading differ from Word's, a squeezed line then wraps and words land on the next paragraph)
    pitch = pitch if n >= 2 else None
    h1 = pitch if pitch else 1.2 * s1
    top = first.base - (h1 - 0.2 * s1)                   # PowerPoint's line: its baseline is a fifth of the letter size above its foot
    bottom = lines[-1].base + 0.25 * s1
    width = max(4.0, bx1 - bx0)
    tb = slide.shapes.add_textbox(Emu(int(round((ox + bx0 * k) * E_PT))), Emu(int(round((oy + top * k) * E_PT))),
                                  Emu(int(round(width * k * E_PT)) + 2 * E_PT), Emu(int(round(max(4.0, bottom - top) * k * E_PT))))
    tf = tb.text_frame
    tf.word_wrap = False
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    par = tf.paragraphs[0]
    par.alignment = algn
    if pitch:
        par.line_spacing = Pt(round(pitch * k, 2))
    ppr = par._p.get_or_add_pPr()
    if abs(mar) >= 0.5 or abs(ind) >= 0.5:
        ppr.set("marL", str(int(round(max(0.0, mar) * k * E_PT))))
        ppr.set("indent", str(int(round(ind * k * E_PT))))
    mini = _Mini()
    mini.lines, mini.hard = lines, n >= 2
    for text, fmt, ul, st, link, field in _runs(mini, doc):
        if field == "BRK:line":
            par.add_line_break()
            continue
        if field is not None and not isinstance(field, str):
            continue
        if not text:
            continue
        fam, size, bold, ital, col, va = fmt
        r = par.add_run()
        r.text = text
        f = r.font
        name = fam or doc.get("font")
        f.size = Pt(round((size or first.size or 12.0) * k * 2) / 2.0)
        f.bold = bool(bold)
        f.italic = bool(ital)
        if ul:
            f.underline = True
        rpr = r._r.get_or_add_rPr()
        if st:
            rpr.set("strike", "sngStrike")
        if va:
            rpr.set("baseline", "30000" if va == "sup" else "-25000")
        f.color.rgb = RGBColor.from_string(col or "000000")
        if name:
            f.name = name
            for tag in ("a:ea", "a:cs"):
                e = PEl(tag)
                e.set("typeface", name)
                rpr.append(e)
        if link:
            try:
                r.hyperlink.address = link
            except Exception:
                pass
    return tb


def _ppt_box(slide, p, doc, k, ox, oy):
    """One paragraph as text boxes: the paragraph itself, and one for each part of a line that stands apart from it
    (the marks at the right of an exam question, the columns of a list without lines).
    k: how much the page is scaled to the slide; (ox, oy): where the page's corner is on the slide."""
    main, apart = [], []
    for ln in p.lines:
        keep = 2 if (ln.marker and len(ln.segs) > 1) else 1
        if len(ln.segs) <= keep:
            main.append(ln)
            continue
        cut = ln.segs[keep][0]
        main.append(_line_part(ln, 0, cut, ln.marker))
        for j in range(keep, len(ln.segs)):
            i0 = ln.segs[j][0]
            i1 = ln.segs[j + 1][0] if j + 1 < len(ln.segs) else len(ln.words)
            part = _line_part(ln, i0, i1, False)
            last = j == len(ln.segs) - 1
            apart.append((part, "right" if (last and p.R - part.x1 <= 2.0 and part.x0 - p.L > 0.5 * (p.R - p.L)) else "left"))
    n = len(main)
    just = p.align == "both" and n >= 2 and not apart
    _ppt_text(slide, main, doc, k, ox, oy, p.align if p.align in ("center", "right") else "left", p.pitch if n >= 2 else None,
              R=(p.R - max(0.0, p.right)) if just else None, just=just)
    for part, al in apart:
        _ppt_text(slide, [part], doc, k, ox, oy, al, None)


def pdf_to_pptx(src_path, out_dir, info=None):
    """A PDF -> a PowerPoint file: one slide a page, each looking like its page, with its words in text boxes that can
    be changed. Nothing is added."""
    from pptx import Presentation
    from pptx.util import Emu
    t0 = time.time()
    os.makedirs(out_dir, exist_ok=True)
    why = pdf_lock(src_path)
    if why:
        raise ConversionError(why)
    doc = {"spaces": True, "pitch": {}, "font": None, "size": 11.0, "hyph": False, "title": None, "all_design": True}
    pages = []
    try:
        pdf = pdfplumber.open(src_path)
    except Exception:
        raise ConversionError("This PDF could not be opened. It may be damaged: try opening it and saving it again as a PDF, then convert that.")
    global _KEEP_LEADERS
    _KEEP_LEADERS = True
    _fonts_begin(src_path)
    try:
        return _pdf_to_pptx(pdf, src_path, out_dir, doc, pages, t0, info)
    finally:
        _KEEP_LEADERS = False


def _pdf_to_pptx(pdf, src_path, out_dir, doc, pages, t0, info):
    from pptx import Presentation
    from pptx.util import Emu
    with pdf:
        n = len(pdf.pages)
        if n == 0:
            raise ConversionError("This PDF has no pages.")
        if n > MAX_PAGES:
            raise ConversionError("This PDF has %d pages. Up to %d can be turned into slides in one go: split it first (PDF tools, Split a PDF), then convert each part." % (n, MAX_PAGES))
        for i, page in enumerate(pdf.pages):
            try:
                pg = _analyse_page(page, i, doc)
            except Exception:
                pg = _Page()
                pg.n, pg.w, pg.h = i, float(page.width), float(page.height)
                pg.ow, pg.oh = pg.w, pg.h
                pg.rows, pg.hs, pg.vs, pg.fills, pg.curves, pg.images, pg.links, pg.rot, pg.blk = [], [], [], [], [], [], [], [], {}
                pg.nchars = pg.bad = pg.spaces = pg.exact = 0
                pg.items, pg.head, pg.foot, pg.scan, pg.zones, pg.ocr = [], [], [], True, [], False
                pg.figs, pg.tables, pg.rules, pg.taken = [], [], [], None
                pg.hfigs, pg.ffigs, pg.hrules, pg.frules, pg.hpics, pg.fpics, pg.design, pg.back = [], [], [], [], [], [], False, None
                pg.body_top = pg.body_bottom = None
            pages.append(pg)
            try:
                page.flush_cache()
            except Exception:
                pass
            el = time.time() - t0
            if el > READ_BUDGET or (i >= 5 and el > 12.0 and el * n / float(i + 1) > READ_BUDGET):
                raise ConversionError(TOO_LONG.replace("into Word", "into slides"))
    doc["pin"] = True
    _doc_stats(pages, doc)
    doc["dflt"] = {"font": doc["font"] or "Calibri", "size": doc["size"]}
    for pg in pages:
        for tb in pg.tables:
            _cell_build(tb, doc)
    _body_box(pages, doc)
    for pg in pages:
        if pg.scan:
            continue
        key = (int(round(pg.w)), int(round(pg.h)))
        L, R, T, B = doc["boxes"][key]
        try:
            pg.zones = _flatten_zones(_flow_of(pg, doc), L, R, doc, pg)
        except Exception:
            pg.zones, pg.scan = [], True                 # a page that cannot be read as words is kept whole, as its picture
    # the slide's size: the size most pages have
    sizes = Counter((round(pg.w, 1), round(pg.h, 1)) for pg in pages)
    SW, SH = sizes.most_common(1)[0][0]
    SW, SH = max(72.0, min(4032.0, SW)), max(72.0, min(4032.0, SH))
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(int(round(SW * E_PT))), Emu(int(round(SH * E_PT)))
    blank = None
    for lay in prs.slide_layouts:
        if not list(lay.placeholders):
            blank = lay
            break
    if blank is None:
        blank = prs.slide_layouts[len(prs.slide_layouts) - 1]
    live = [pg.n for pg in pages if not pg.scan]
    empty = set()                    # the pages that draw nothing but their words: no picture goes behind those
    bg, idx = _strip_words(src_path, live, empty) if live else (None, {})
    kept = 0
    pace = _Pace(t0, sum(1 for pg in pages if pg.n not in empty or pg.rot))
    try:
        for pg in pages:
            slide = prs.slides.add_slide(blank)
            for ph in list(slide.placeholders):
                ph._element.getparent().remove(ph._element)
            k = min(SW / pg.w, SH / pg.h)
            ox, oy = (SW - pg.w * k) / 2.0, (SH - pg.h * k) / 2.0
            im = None
            if pg.scan or bg is None or pg.n not in idx:
                q, t = pace.q(), time.time()
                im = _raster(src_path, pg.n, int(BACK_DPI * q), pg.w, pg.h)        # the page as it is, words and all
                if pg.scan:
                    kept += 1
            elif pg.n not in empty or pg.rot:
                q, t = pace.q(), time.time()
                im = _back_image(src_path, bg, idx, pg, int(BACK_DPI * q))
            if im is not None:
                try:
                    if not _is_blank(im):
                        data, ext = _encode(im)
                        slide.shapes.add_picture(io.BytesIO(data), Emu(int(round(ox * E_PT))), Emu(int(round(oy * E_PT))),
                                                 width=Emu(int(round(pg.w * k * E_PT))), height=Emu(int(round(pg.h * k * E_PT))))
                except Exception:
                    pass
                im.close()
                pace.took(q, time.time() - t)
            if pg.scan or (bg is None or pg.n not in idx):
                continue
            for p in sorted(_ppt_paras(pg), key=lambda q: (q.lines[0].base, q.lines[0].x0)):
                try:
                    _ppt_box(slide, p, doc, k, ox, oy)
                except Exception:
                    continue
    finally:
        if bg:
            shutil.rmtree(os.path.dirname(bg), ignore_errors=True)
    _drop_thumbnail(prs.part.package)
    cp = prs.core_properties
    cp.author = cp.title = cp.comments = cp.keywords = cp.subject = cp.category = cp.last_modified_by = ""
    try:
        import datetime as _dt
        now = _dt.datetime.now(_dt.timezone.utc).replace(microsecond=0, tzinfo=None)
        cp.created = cp.modified = now
        cp.revision = 1
    except Exception:
        pass
    base = os.path.splitext(os.path.basename(src_path))[0] or "converted"
    out = os.path.join(out_dir, base + ".pptx")
    prs.save(out)
    if info is not None:
        info["pages"] = len(pages)
        info["ocr"] = 0
        info["pictures"] = kept
    return out


class _Holder(object):
    """Where a page's standing pictures are kept in the text's flow: a line one point high at the page's top."""
    __slots__ = ("top", "bottom")


def _with_first(flows, pc, L, R, T):
    """Put a picture that stands at its own place (pc.anchor) first on its page. In the flow it is a line one point
    high at the top of the text area; _item_edges() reads that from pc.flow."""
    pc.flow = _Holder()
    pc.flow.top, pc.flow.bottom = T, T + 1.0
    if flows and flows[0][0] == "one":
        flows[0] = ("one", [pc] + list(flows[0][1]), flows[0][2])
    elif flows and flows[0][0] == "cols" and flows[0][1] and flows[0][1][0]:
        cols = list(flows[0][1])
        cols[0] = [pc] + list(cols[0])
        flows[0] = ("cols", cols, flows[0][2])
    else:
        flows.insert(0, ("one", [pc], (L, R)))
    return flows


def _fonts_begin(src_path):
    """A new PDF: what is known of the last one's fonts is forgotten, and this one's fonts are looked up."""
    _FONT_CACHE.clear()
    _W_STAT.clear()
    _KIND_HINT.clear()
    _KIND_HINT.update(_font_kinds(src_path))


def _harden(pg):
    """On a designed page the words stand on a drawing: every line must end where it ends in the PDF and stand at the
    PDF's distance from the line before, whatever font the reader's computer puts in."""
    def walk(items):
        for it in items:
            if isinstance(it, Para) and it.lines:
                if len(it.lines) >= 2:
                    it.hard = True
                    it.runs = None
                    if it.pitch:
                        it.spec = ("exact", _tw(it.pitch))
                    if it.align == "both":
                        it.right = max(0.0, it.R - max(l.x1 for l in it.lines))
                    elif it.align == "left" and it.right > 0:
                        it.right = 0.0
            elif isinstance(it, Tbl):
                for ce in it.cells:
                    walk(ce.items)
    for f in pg.zones or []:
        if f[0] == "one":
            walk(f[1])
        else:
            for c in f[1]:
                walk(c)


def pdf_to_docx(src_path, out_dir, style=None, info=None, notes=None):
    """A PDF -> a Word file that looks like it and can be changed. Nothing is added.
    info, a dict, is told what was done: pages, ocr (scanned pages whose words were read), pictures (scanned pages kept
    as pictures of the page). notes: {page number from 0: [lines]}, words that belong to a page without being on it (a
    slide's speaker notes, when the PDF was made from slides): each page's follow it on a page of their own."""
    t0 = time.time()
    os.makedirs(out_dir, exist_ok=True)
    why = pdf_lock(src_path)
    if why:
        raise ConversionError(why)
    doc = {"spaces": True, "pitch": {}, "font": None, "size": 11.0, "hyph": False, "page_notes": notes or {}}
    try:
        pdf = pdfplumber.open(src_path)
    except Exception as e:
        raise ConversionError("This PDF could not be opened. It may be damaged: try opening it and saving it again as a PDF, then convert that.")
    _fonts_begin(src_path)
    pages = []
    with pdf:
        n = len(pdf.pages)
        if n == 0:
            raise ConversionError("This PDF has no pages.")
        if n > MAX_PAGES:
            raise ConversionError("This PDF has %d pages. PDF to Word takes up to %d pages at a time: split it first (PDF tools, Split a PDF), then convert each part." % (n, MAX_PAGES))
        try:
            doc["title"] = ""
        except Exception:
            pass
        nch = nsp = nex = 0
        for i, page in enumerate(pdf.pages):
            try:
                pg = _analyse_page(page, i, doc)
            except Exception:
                pg = _Page()
                pg.n, pg.w, pg.h = i, float(page.width), float(page.height)
                pg.ow, pg.oh = pg.w, pg.h
                pg.rows, pg.hs, pg.vs, pg.fills, pg.curves, pg.images, pg.links, pg.rot, pg.blk = [], [], [], [], [], [], [], [], {}
                pg.nchars = pg.bad = pg.spaces = pg.exact = 0
                pg.items, pg.head, pg.foot, pg.scan, pg.zones, pg.ocr = [], [], [], True, [], False
                pg.figs, pg.tables, pg.rules, pg.taken = [], [], [], None
                pg.hfigs, pg.ffigs, pg.hrules, pg.frules, pg.hpics, pg.fpics, pg.design, pg.back = [], [], [], [], [], [], False, None
                pg.body_top = pg.body_bottom = None
            nch += pg.nchars
            nsp += pg.spaces
            nex += pg.exact
            pages.append(pg)
            try:
                page.flush_cache()
            except Exception:
                pass
            el = time.time() - t0
            if el > READ_BUDGET or (i >= 5 and el > 12.0 and el * n / float(i + 1) > READ_BUDGET):
                raise ConversionError(TOO_LONG)          # said after seconds, not after minutes of waiting
    done = 0
    for pg in pages:
        if not pg.scan or done >= MAX_OCR_PAGES or time.time() - t0 > OCR_BUDGET:
            continue
        done += 1
        try:
            if _ocr_page(src_path, pg, doc):
                pg.scan, pg.ocr = False, True
        except Exception:
            pass
    try:
        _ocr_bold(doc)
    except Exception:
        pass
    doc["spaces"] = nsp > 0.02 * max(1, nch)
    # with the PDF's own fonts Word fills the pages as the PDF does, and the text is left to flow; with stand-in fonts
    # every page and column is made to start where it starts in the PDF
    doc["pin"] = nex < 0.9 * max(1, nch)
    _split_head_foot(pages)
    _doc_stats(pages, doc)
    if not doc["font"]:
        # no font is named anywhere (a scan read by OCR): a book face for printed pages
        doc["font"] = "Times New Roman" if any(pg.ocr for pg in pages) else "Calibri"
    doc["dflt"] = {"font": doc["font"], "size": doc["size"] or 11.0, "wscale": _width_scales()}
    for pg in pages:
        for tb in pg.tables:
            _cell_build(tb, doc)
    _body_box(pages, doc)
    _plan_head_foot(pages, doc)
    for pg in pages:
        key = (int(round(pg.w)), int(round(pg.h)))
        L, R, T, B = doc["boxes"][key]
        if pg.scan:
            pc = Pic()
            pc.x0, pc.top, pc.x1, pc.bottom = L, T, R, min(B, T + (R - L) * pg.h / pg.w)
            pc.data = pc.ext = pc.anchor = pc.size = None
            pc.behind = False
            pc.kids = pc.flow = None
            pc.brk = pc.colbrk = False
            pc.sect = None
            pc.before = pc.after = pc.left = 0.0
            pc.page, pc.align, pc.L, pc.R = pg.n, "left", L, R
            pg.zones = [("one", [pc], (L, R))]
            continue
        side = [f for f in pg.figs if f[2] <= L - 1.0 or f[0] >= R + 1.0] if not pg.design else []
        if side:
            pg.figs = [f for f in pg.figs if f not in side]
        pg.zones = _flatten_zones(_flow_of(pg, doc), L, R, doc, pg)
        if side:
            # a picture (or words turned on their side) in the paper's margin: it stands where it stands, whatever the
            # text does. In the text's flow it would push a column's words down by its own height.
            hold = None
            for f in side:
                pc = Pic()
                pc.x0, pc.top, pc.x1, pc.bottom = f[0], f[1], f[2], f[3]
                pc.data = pc.ext = pc.size = pc.kids = pc.flow = None
                pc.anchor, pc.behind = (f[0], f[1]), False
                pc.brk = pc.colbrk = False
                pc.sect = None
                pc.before = pc.after = pc.left = 0.0
                pc.page, pc.align, pc.L, pc.R = pg.n, "left", L, R
                if hold is None:
                    hold = pc
                    hold.kids = []
                else:
                    hold.kids.append(pc)
            pg.zones = _with_first([f for f in (pg.zones or []) if any(f[1] if f[0] == "cols" else [f[1]])], hold, L, R, T)
        if pg.design:
            _harden(pg)
            pc = Pic()                                  # the page's drawing, whole, behind its words
            pc.x0, pc.x1, pc.top, pc.bottom = 0.0, pg.w, T, T + 1.0     # (in the flow it is a line one point high)
            pc.data = pc.ext = None
            pc.anchor, pc.behind, pc.size, pc.kids, pc.flow = (0.0, 0.0), True, (pg.w, pg.h), None, None
            pc.brk = pc.colbrk = False
            pc.sect = None
            pc.before = pc.after = pc.left = 0.0
            pc.page, pc.align, pc.L, pc.R = pg.n, "left", L, R
            pg.back = pc
            pg.zones = _with_first([f for f in (pg.zones or []) if any(f[1] if f[0] == "cols" else [f[1]])], pc, L, R, T)
    _link_notes(pages, doc)
    body = _assemble(pages, doc)
    _render_pictures(src_path, body, pages, doc, t0)
    base = os.path.splitext(os.path.basename(src_path))[0] or "converted"
    out = os.path.join(out_dir, base + ".docx")
    _write(body, pages, doc, out)
    if info is not None:
        info["pages"] = len(pages)
        info["ocr"] = sum(1 for pg in pages if pg.ocr)
        info["pictures"] = sum(1 for pg in pages if pg.scan)
    return out


if __name__ == "__main__":
    # python3 pdf2word.py [--pptx] <the PDF> <the folder for the Word (PowerPoint) file>: one line of JSON says how it went.
    # The server runs every PDF to Word this way, in a process of its own: its memory goes back when it ends, a PDF
    # that is too much for the server stops this process and not the server, and other people's requests do not wait.
    import json as _json
    import sys as _sys
    _info = {}
    try:
        # the server says how many seconds this conversion has (it may have waited for its turn first)
        _has = float(os.environ.get("P2W_SECONDS") or 0)
        if 40.0 <= _has < 200.0:
            _k = _has / 200.0
            TIME_BUDGET, RENDER_END, OCR_BUDGET, READ_BUDGET = TIME_BUDGET * _k, RENDER_END * _k, OCR_BUDGET * _k, READ_BUDGET * _k
    except Exception:
        pass
    try:
        if _sys.argv[1] == "--pptx":
            _out = pdf_to_pptx(_sys.argv[2], _sys.argv[3], info=_info)
        elif _sys.argv[1] == "--notes":
            with open(_sys.argv[2], encoding="utf-8") as _fh:
                _notes = dict((int(k), v) for k, v in _json.load(_fh).items())
            _out = pdf_to_docx(_sys.argv[3], _sys.argv[4], info=_info, notes=_notes)
        else:
            _out = pdf_to_docx(_sys.argv[1], _sys.argv[2], info=_info)
        print(_json.dumps({"ok": _out, "info": _info}))
    except ConversionError as _e:
        print(_json.dumps({"error": str(_e)}))
        _sys.exit(3)
    except MemoryError:
        print(_json.dumps({"error": TOO_LONG}))
        _sys.exit(4)
