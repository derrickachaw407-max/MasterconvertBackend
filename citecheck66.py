"""6.6, round 5 (the owner, 21 Nov 2026 00:47: "Raise the app's validity. Do as much as a research tool can be."): Check my
citations. The person pastes their own writing with its reference list; the app finds each reference's paper in its
registry (Crossref), says what differs from the publisher's record (year, journal, volume, pages, the first author), what
couldn't be found, what was retracted; matches every in-text citation to its reference (and says which have none, and which
references are never cited); and checks each cited sentence against its paper's abstract (server/app.py cite_check_endpoint,
with the AI and the rules of server/valid65.py).

Here only what needs neither the network nor the AI: the writing cut into its body and its reference list, the reference
entries read (number, first author, year, DOI), the in-text citations found (author-date, narrative, MLA's author and page,
numbers), each citation matched to its reference, and a reference compared with its registry record. Tests:
tests/test_citecheck66.py."""
import re
import unicodedata

TEXT_MAX = 30000                 # characters of writing, references included
REFS_MAX = 60                    # references looked up
SENTS_MAX = 40                   # cited sentences checked against their papers

_REF_HEAD = re.compile(r"^\s*(?:#{1,6}\s*)?(?:\*\*)?\s*(?:\d+(?:\.\d+)*\.?\s+)?(?:references?(?:\s+list)?|reference\s+list|works\s+cited|"
                       r"bibliography|list\s+of\s+references|literature\s+cited|sources(?:\s+cited)?)\s*(?:\*\*)?\s*:?\s*$", re.I)
_REF_START = re.compile(r"^\s*(?:\[\d{1,3}\]|\d{1,3}[.)]\s|[A-ZÀ-ɏ]|(?:(?:van|von|de|del|der|den|du|da|di|le|la|al|el|ten|ter)\s+)+[A-ZÀ-ɏ]|"
                        r"(?:-{2,3}|—{1,3}|–{2,3}|_{3,})\s*[.,])")                 # (6.6, round 28: MLA's "---." for the same author)
_DOI = re.compile(r"\b(10\.\d{4,9}/[^\s\"<>]+)", re.I)
_YEAR = re.compile(r"\((?P<p>(?:19|20)\d{2})[a-z]?(?:,[^)]*)?\)|\b(?P<b>(?:19|20)\d{2})[a-z]?\b")


def norm_name(s):
    """A surname for comparing: accents off, lower case, letters only ("van den Berg" -> "vandenberg")."""
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z]+", "", s)


def split_writing(text):
    """(body, [reference entries]): the reference list is what follows the last heading that names one ("References",
    "Reference list", "Works Cited", "Bibliography"); an entry that runs over several lines is joined back into one."""
    lines = str(text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    at = None
    for i, l in enumerate(lines):
        if _REF_HEAD.match(l):
            at = i
    if at is None:
        return str(text or "").strip(), []
    body = "\n".join(lines[:at]).strip()
    entries = []
    for l in lines[at + 1:]:
        t = re.sub(r"\s+", " ", l).strip()
        if not t:
            continue
        starts = bool(_REF_START.match(t)) and not t.lower().startswith(("http", "doi:", "available at", "retrieved"))
        numbered = bool(re.match(r"^\s*(?:\[\d{1,3}\]|\d{1,3}[.)]\s)", t))
        open_end = bool(entries) and not re.search(r"(?:[.)\]”\"]|https?://\S+|doi:\s*\S+)\s*$", entries[-1])
        if entries and not numbered and (not starts or open_end or len(t) < 25):
            entries[-1] = entries[-1] + " " + t           # the rest of the entry before it
        else:
            entries.append(t)
    return body, entries[:REFS_MAX]


def _first_author(entry):
    """The first author's surname (or a group's name) as the entry writes it."""
    t = re.sub(r"^\s*(?:\[\d{1,3}\]|\d{1,3}[.)])\s*", "", entry).strip()
    ieee = re.match(r"^(?:[A-Z]\.\s?(?:-?[A-Z]\.\s?)*)\s*((?:(?:van|von|de|der|den|du|da|di|le|la)\s+)*[A-Z][\w'’\-]+)", t)
    if ieee:                                                                    # IEEE: "K. A. Mensah and A. Boateng, …"
        return ieee.group(1)
    head = re.split(r",|\.\s|\s\(|\s(?:&|and)\s", t, maxsplit=1)[0].strip()
    m = re.fullmatch(r"(?:[A-Z]\.\s?(?:-?[A-Z]\.\s?)*)\s*(.+)", head)            # IEEE: "K. A. Mensah"
    if m:
        return m.group(1).strip()
    m = re.fullmatch(r"(.+?)\s+[A-Z]{1,3}", head)                               # Vancouver: "Mensah KA"
    if m and not re.search(r"\s", m.group(1).split()[-1]):
        return m.group(1).strip()
    return head


# 6.6, round 24: a work with no date ("(n.d.)") or not yet out ("(in press)", "forthcoming"): no year, though the entry
# may name one elsewhere ("Retrieved March 3, 2024, from …")
_UNDATED = re.compile(r"\(\s*(n\.\s?d\.|no\s+date|in\s+press|forthcoming)|(?:^|[.,]\s)(n\.\s?d\.|(?:in\s+press|forthcoming)(?=[.,;]))", re.I)


def _undated_kind(t):
    t = re.sub(r"\s+", " ", str(t or "")).lower()
    return "nd" if t.startswith("n") else "press" if t else ""


def read_reference(entry, i):
    """One entry: {"n": its place in the list (1…), "num": the number it is written with, if any, "text", "surname",
    "year", "doi", "undated": "nd" (no date) or "press" (in press, forthcoming) or ""}."""
    num = re.match(r"^\s*(?:\[(\d{1,3})\]|(\d{1,3})[.)]\s)", entry)
    years = [m.group("p") or m.group("b") for m in _YEAR.finditer(entry)]
    paren = [m.group("p") for m in _YEAR.finditer(entry) if m.group("p")]
    doi = _DOI.search(entry)
    und = _UNDATED.search(entry)
    first_paren = next((m for m in _YEAR.finditer(entry) if m.group("p")), None)
    undated = _undated_kind(und.group(1) or und.group(2)) if und and not (first_paren and first_paren.start() < und.start()) else ""
    return {"n": i + 1, "num": int(num.group(1) or num.group(2)) if num else None, "text": entry[:700], "surname": _first_author(entry)[:120],
            "year": None if undated else (int((paren or years or [0])[0]) or None), "doi": doi.group(1).rstrip(".,;)]").lower() if doi else None,
            "undated": undated}


# ---------------------------------------------------------------- in-text citations
_PAREN = re.compile(r"\(([^()]{2,300})\)")
_ONE = re.compile(r"^\s*(?:see\s+also\s+|see\s+|e\.g\.,?\s+|cf\.\s+|also\s+|as\s+cited\s+in\s+)?(?P<names>[^\d()]{2,160}?),?\s+"
                  r"(?P<years>(?:(?:19|20)\d{2}[a-z]?|n\.\s?d\.(?:-[a-z])?|no\s+date|in\s+press(?:-[a-z])?|forthcoming)(?:\s*,\s*(?:(?:19|20)\d{2}[a-z]?|"
                  r"n\.\s?d\.(?:-[a-z])?|no\s+date|in\s+press(?:-[a-z])?|forthcoming))*)(?:\s*,\s*(?:(?:p|pp|para|chap)\.?\s*)?[\w–\-]+)?\s*$", re.I)
_MLA_ONE = re.compile(r"^\s*(?P<names>[A-ZÀ-ɏ][^\d()]{1,120}?)\s+(?P<page>\d{1,5}(?:\s*[–-]\s*\d{1,5})?)\s*$")
_NARR = re.compile(r"(?P<names>(?:(?:van|von|de|der|den|du|da|di|le|la|al|el)\s+)*[A-Z][\w'’\-]+(?:\s+[A-Z][\w'’\-]+){0,3}(?:\s+et\s+al\.?|\s+(?:and|&)\s+(?:(?:van|von|de|der|den|du|da|di)\s+)*[A-Z][\w'’\-]+)?)"
                   r"(?:['’]s)?\s+\((?P<years>(?:(?:19|20)\d{2}[a-z]?|n\.\s?d\.|no\s+date|in\s+press|forthcoming)(?:\s*,\s*(?:(?:19|20)\d{2}[a-z]?|"
                   r"n\.\s?d\.|no\s+date|in\s+press|forthcoming))*)(?:,\s*(?:p|pp)\.?\s*[\w–\-]+)?\)")
_NUM = re.compile(r"\[(\d{1,3}(?:\s*[,–\-]\s*\d{1,3})*)\]")
# 6.6, round 14: Vancouver's other ways: "(1)", "(1, 3–5)", and numbers raised as Word writes them ("sleep¹˒²", "use³⁻⁵")
_NUM_PAREN = re.compile(r"\((\d{1,3}(?:\s*[,–\-]\s*\d{1,3})*)\)")
_SUP_DIGITS = {"⁰": "0", "¹": "1", "²": "2", "³": "3", "⁴": "4", "⁵": "5", "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9", "⁻": "-", "˒": ",", "‚": ","}
_NUM_SUP = re.compile(r"[⁰¹²³⁴⁵⁶⁷⁸⁹]+(?:\s*[,˒‚⁻–\-]\s*[⁰¹²³⁴⁵⁶⁷⁸⁹]+)*")
# "(Mensah, 2021, as cited in Owusu, 2023)", "Mensah (2021, as cited in Owusu, 2023)": the source read is the second
_AS_CITED = re.compile(r",?\s*as\s+cited\s+in\s+", re.I)
_NOT_NAMES = re.compile(r"^(?:table|figure|fig|appendix|chapter|section|see|page|pp?\.?\s*\d|n\s*=|n\.d|years?|aged?|age|i\.e|e\.g)\b|"
                        r"^(?:pp?|n)\s*$", re.I)                 # (6.6, round 46: "P. Owusu" is a name, "p. 4" is not)


def _surname_of(names):
    """The first author of a citation's names: "Mensah et al." -> "Mensah", "Osei & Lamptey" -> "Osei"; APA's first citation
    of a group with its short name, "World Health Organization [WHO]" -> "World Health Organization" (6.6, round 41)."""
    names = re.sub(r"\s*\[[^\]]{1,15}\]", "", names)
    first = re.split(r"\s+et\s+al\.?|\s*&\s*|\s+and\s+|,", names.strip(), maxsplit=1)[0].strip()
    first = re.sub(r"^(?:the|a|an|in|as|by|from|according\s+to)\s+", "", first, flags=re.I)
    first = _INITIALS_LEAD.sub("", first)                   # (6.6, round 46: "Y. Mensah", two first authors of one surname)
    return re.sub(r"['’]s$", "", first)


_INITIALS_LEAD = re.compile(r"^(?:[A-ZÀ-ɏ]\.\s*(?:-\s*)?)+\s*(?=[A-ZÀ-ɏa-z])")


def initials_of_cite(names):
    """The initials a citation puts before its first author ("(Y. Mensah et al., 2021)": "Y"), or ""."""
    names = re.sub(r"^\s*\(|\)\s*$", "", str(names or "")).strip()
    m = _INITIALS_LEAD.match(names)
    return "".join(re.findall(r"[A-ZÀ-ɏ]", m.group(0))) if m else ""


def initials_of_entry(entry, surname):
    """The first author's initials as an entry gives them: "Mensah, Y. A.," / "Mensah, Yaw," -> "YA" / "Y"; IEEE's
    "Y. A. Mensah" -> "YA"; "" when it can't tell."""
    t = re.sub(r"^\s*(?:\[\d{1,3}\]|\d{1,3}[.)])\s*", "", str(entry or ""))
    sn = str(surname or "").strip()
    if not sn:
        return ""
    m = re.match(re.escape(sn) + r",\s*((?:[A-ZÀ-ɏ][\w'’\-]*\.?\s*)+?)(?:,|\.\s|\s\(|\s&|\sand\s|\(|$)", t)
    if m:
        return "".join(re.findall(r"(?:^|[\s.\-])([A-ZÀ-ɏ])", " " + m.group(1)))
    m = re.match(r"((?:[A-ZÀ-ɏ]\.\s?(?:-?[A-ZÀ-ɏ]\.\s?)*))\s*" + re.escape(sn) + r"\b", t)      # IEEE
    return "".join(re.findall(r"[A-ZÀ-ɏ]", m.group(1))) if m else ""


def _expand(nums):
    out = []
    for part in re.split(r"\s*,\s*", nums):
        a, _, b = re.sub(r"\s+", "", part).replace("–", "-").partition("-")
        if a.isdigit() and b.isdigit() and int(a) <= int(b) <= int(a) + 50:
            out += list(range(int(a), int(b) + 1))
        elif a.isdigit():
            out.append(int(a))
    return out


# 6.6, round 28: MLA's citation of one of two works by the same author, by its author and a short title
_MLA_TITLED = re.compile(r"^\s*(?P<names>(?:[A-Z]\.\s*)*(?:(?:van|von|de|der|den|du|da|di|le|la)\s+)*[A-ZÀ-ɏ][\w'’\-]+(?:\s+[A-Z][\w'’\-]+){0,5}?(?:\s+et\s+al\.?|"
                         r"\s+and\s+(?:(?:van|von|de|der|den|du|da|di)\s+)*[A-Z][\w'’\-]+)?)\s*,\s*[“\"‘](?P<title>[^”\"’]{2,120})[”\"’]"
                         r"(?:\s+(?P<page>\d{1,5}(?:\s*[–\-]\s*\d{1,5})?))?\s*$")
_MLA_NAME = re.compile(r"^\s*(?P<names>(?:[A-Z]\.\s*)*(?:(?:van|von|de|der|den|du|da|di|le|la)\s+)*[A-ZÀ-ɏ][\w'’\-]+(?:\s+et\s+al\.?|\s+and\s+(?:(?:van|von|de|der|den|du|da|di)\s+)*[A-Z][\w'’\-]+)?)\s*$")


def find_citations(body, numbered=False, names=None):
    """The in-text citations of the writing: [{"text": as written, "surname", "year" (int or None), "suffix" ("a"…),
    "num" (for a numbered style), "start", "end"}]. A parenthesis that isn't a citation ("(n = 40)", "(Table 2)") is
    left alone. names (6.6, round 10: an MLA list's first authors, as norm_name gives them): MLA's citation of a work
    with no page, "(Mensah et al.)", "(Darko and Adjei)", is one when its first name is one of the list's."""
    out = []
    taken = []
    for m in _PAREN.finditer(body or ""):
        inside = m.group(1)
        parts = [p for p in re.split(r"\s*;\s*", inside) if p.strip()]
        got = []
        for p in parts:
            shown = "(" + p.strip() + ")"
            sec = _AS_CITED.split(p, maxsplit=1)
            if len(sec) == 2 and not numbered:
                p = sec[1]                              # 6.6, round 14: the work read (the list's entry) is the one cited in
            t_ = _MLA_TITLED.match(p) if not numbered else None
            if t_:
                got.append({"text": shown, "surname": _surname_of(t_.group("names")), "year": None, "suffix": "", "num": None, "mla": True,
                            "short_title": t_.group("title").strip()})
                continue
            c = _MLA_NAME.match(p) if names and not numbered else None
            if c and norm_name(_surname_of(c.group("names"))) in names:
                got.append({"text": shown, "surname": _surname_of(c.group("names")), "year": None, "suffix": "", "num": None, "mla": True})
                continue
            if names and not numbered and re.fullmatch(r"\s*[A-Z][\w'’\-]*(?:\s+(?:of|for|and|the|on|in|&)?\s*[A-Z][\w'’\-]*){1,7}\s*", p) \
                    and norm_name(p) in names:              # 6.6, round 24: MLA's group author, "(Ghana Statistical Service)"
                got.append({"text": shown, "surname": p.strip(), "year": None, "suffix": "", "num": None, "mla": True})
                continue
            a = _ONE.match(p)
            if a and not _NOT_NAMES.match(a.group("names").strip()) and re.search(r"[A-ZÀ-ɏ]", a.group("names")):
                for y in re.split(r"\s*,\s*", a.group("years")):
                    ym = re.match(r"((?:19|20)\d{2})([a-z]?)", y)
                    got.append({"text": shown, "surname": _surname_of(a.group("names")), "year": int(ym.group(1)) if ym else None,
                                "suffix": ym.group(2) if ym else "", "num": None, "undated": "" if ym else _undated_kind(y)})
                continue
            b = _MLA_ONE.match(p)
            if b and not numbered and not _NOT_NAMES.match(b.group("names").strip()) and len(b.group("names").split()) <= 6:
                got.append({"text": "(" + p.strip() + ")", "surname": _surname_of(b.group("names")), "year": None, "suffix": "", "num": None, "mla": True})
        if got and len(got) >= max(1, len(parts) // 2):
            for g in got:
                g.update(start=m.start(), end=m.end())
            out += got
            taken.append((m.start(), m.end()))
    for m in re.finditer(r"\((?:19|20)\d{2}[a-z]?\s*,?\s*as\s+cited\s+in\s+([^()]{3,160})\)", body or "", re.I):
        if any(a <= m.start() < b for a, b in taken):
            continue                                    # 6.6, round 14: "Mensah (2021, as cited in Owusu, 2023)"
        a = _ONE.match(m.group(1))
        if a and re.search(r"[A-ZÀ-ɏ]", a.group("names")):
            for y in re.split(r"\s*,\s*", a.group("years")):
                ym = re.match(r"((?:19|20)\d{2})([a-z]?)", y)
                out.append({"text": m.group(0), "surname": _surname_of(a.group("names")), "year": int(ym.group(1)) if ym else None,
                            "suffix": ym.group(2) if ym else "", "num": None, "undated": "" if ym else _undated_kind(y),
                            "start": m.start(), "end": m.end()})
            taken.append((m.start(), m.end()))
    for m in _NARR.finditer(body or ""):
        if any(a <= m.start() < b for a, b in taken) or _NOT_NAMES.match(m.group("names")):
            continue
        for y in re.split(r"\s*,\s*", m.group("years")):
            ym = re.match(r"((?:19|20)\d{2})([a-z]?)", y)
            out.append({"text": m.group(0), "surname": _surname_of(m.group("names")), "year": int(ym.group(1)) if ym else None,
                        "suffix": ym.group(2) if ym else "", "num": None, "undated": "" if ym else _undated_kind(y),
                        "start": m.start(), "end": m.end()})
    if numbered:
        for m in _NUM.finditer(body or ""):
            for n in _expand(m.group(1)):
                out.append({"text": m.group(0), "surname": None, "year": None, "suffix": "", "num": n, "start": m.start(), "end": m.end()})
        # 6.6, round 14: "(1)", "(1, 3–5)" (a number alone in brackets, not a year) and raised numbers
        for m in _NUM_PAREN.finditer(body or ""):
            if any(a <= m.start() < b for a, b in taken):
                continue
            for n in _expand(m.group(1)):
                out.append({"text": m.group(0), "surname": None, "year": None, "suffix": "", "num": n, "start": m.start(), "end": m.end()})
        for m in _NUM_SUP.finditer(body or ""):
            before = (body or "")[max(0, m.start() - 6):m.start()]
            if re.search(r"(?:^|[^A-Za-z])(?:[A-Za-zχηωρσμ]|cm|km|mm|kg/m|ft)$|\d$|[⁻⁺⁼^]$", before):
                continue                                # (m², R², χ², 10⁻⁵: a power, not a citation)
            flat = "".join(_SUP_DIGITS.get(ch, ch) for ch in m.group(0))
            for n in _expand(flat):
                out.append({"text": m.group(0), "surname": None, "year": None, "suffix": "", "num": n, "start": m.start(), "end": m.end()})
    out.sort(key=lambda c: c["start"])
    return out


_SAME_AUTHOR = re.compile(r"^\s*(?:-{2,3}|—{1,3}|–{2,3}|_{3,})\s*[.,]")


def fill_repeated_authors(refs):
    """6.6, round 28: "---." (MLA 9, Chicago) stands for the same author(s) as the entry before it: that entry's first
    author, so the entry is matched to its citations."""
    for i, r in enumerate(refs):
        if i and _SAME_AUTHOR.match(r.get("text") or ""):
            r["surname"] = refs[i - 1].get("surname") or r.get("surname")
            r["same_author"] = True
    return refs


def is_numbered(refs):
    """A reference list in a numbered style: most of its entries begin with a number."""
    return bool(refs) and sum(1 for r in refs if r.get("num")) >= max(1, (len(refs) * 2) // 3)


def match(cites, refs, style=None):
    """Each citation's reference (its "n"), or None when the list has none for it. By number in a numbered style; else
    by the first author's surname and the year (a letter after the year tells two works of one year apart)."""
    by_num = {r["num"]: r["n"] for r in refs if r.get("num")}
    for c in cites:
        c["ref"] = None
        if c.get("num"):
            c["ref"] = by_num.get(c["num"]) or (c["num"] if not by_num and 1 <= c["num"] <= len(refs) else None)
            continue
        sn = norm_name(c.get("surname"))
        if not sn:
            continue
        # "Ama Mensah (2021)", "World Health Organization (2023)": the whole name, or one of its words, is the entry's
        alts = {sn} | {norm_name(w) for w in str(c.get("surname") or "").split() if len(norm_name(w)) > 2}
        cands = [r for r in refs if norm_name(r.get("surname")) in alts or (len(sn) > 3 and norm_name(r.get("surname")).startswith(sn))]
        if not cands and re.fullmatch(r"[A-Z][A-Z&]{1,9}", str(c.get("surname") or "").strip()):
            cands = [r for r in refs if _is_acronym_of(c["surname"].strip(), r.get("surname"))]     # 6.6, round 41: "(WHO, 2023)"
        ini = initials_of_cite(c.get("text") or "")
        if ini and len(cands) > 1:                      # 6.6, round 46: "(Y. Mensah et al., 2021)" is Mensah, Y., not Mensah, A.
            fit_i = [r for r in cands if initials_of_entry(r["text"], r.get("surname")).startswith(ini)]
            cands = fit_i or cands
        if c.get("short_title"):                          # 6.6, round 28: MLA's "(Darko, “Same Year A”)"
            flat = lambda t: " ".join(re.findall(r"[a-z0-9]+", unicodedata.normalize("NFKD", str(t or "")).encode("ascii", "ignore").decode().lower()))
            st = flat(c["short_title"])
            tw = {w for w in _words(c["short_title"]) if w not in ("the", "and", "for", "with")}

            def score(r):
                qt = re.search(r"[“\"]([^”\"]{2,300})[”\"]", r["text"])
                title = flat(qt.group(1) if qt else r["text"])
                return 3 if title == st else 2 if (title + " ").startswith(st + " ") else 1 if tw and tw <= _words(title) else 0
            best = max((score(r) for r in cands), default=0)
            cands = [r for r in cands if best and score(r) == best]
        if c.get("year"):
            cands = [r for r in cands if r.get("year") == c["year"]]
        elif c.get("undated"):                          # 6.6, round 24: "(Haddad & Lee, in press)", "(… n.d.)"
            cands = [r for r in cands if not r.get("year") and r.get("undated") in (c["undated"], "")]
        if len(cands) > 1 and c.get("suffix"):
            letters = "abcdefghijklmnopqrstuvwxyz"
            k = letters.find(c["suffix"])
            same = [r for r in cands if re.search(r"\(" + str(c["year"]) + c["suffix"] + r"\)|\b" + str(c["year"]) + c["suffix"] + r"\b", r["text"])]
            cands = same or (cands[k:k + 1] if 0 <= k < len(cands) else cands)
        if len(cands) > 1:
            # ("(Osei, 2020)" and "(Osei & Mensah, 2020)", MLA's "(Osei)" and "(Osei et al.)": works of one first author,
            # told apart by how many authors the citation names, and by its second author)
            t = re.sub(r"\(|\)|,?\s*(?:(?:19|20)\d{2}[a-z]?|n\.\s?d\.).*$", "", str(c.get("text") or "")).strip()
            t = re.sub(r"\s+\d{1,5}(?:\s*[–\-]\s*\d{1,5})?\s*$", "", t)              # (MLA's page)
            n_names = len([x for x in re.split(r"\s*,\s*|\s*&\s*|\s+and\s+", re.sub(r",?\s*\bet\s+al\b\.?", "", t)) if x.strip()])
            has_etal = bool(re.search(r"\bet\s+al\b", t))
            k = ETAL_FROM.get(style or "", 3)

            def fits(r):
                n = entry_authors(r["text"])
                if not n:
                    want = "many" if has_etal or n_names >= 3 else "two" if n_names == 2 else "one"
                    return _shape(r["text"]) == want
                if n >= 100:
                    return has_etal or n_names >= 3
                return n >= max(k, n_names + 1) if has_etal else n == n_names
            fit = [r for r in cands if fits(r)]
            # every name the citation gives after the first is in the entry's authors: "(Osei & Mensah, 2020)", and APA's
            # "(Osei, Darko, et al., 2023)" for one of two works that would both be "Osei et al., 2023" (6.6, round 45)
            # (Harvard's "Mensah-Bonsu, Lamptey et al." has its "et al." on the last name, not after a comma: 6.6, round 46)
            listed = [y for y in (re.sub(r"\s*\bet\s+al\b\.?", "", x).strip() for x in re.split(r"\s*,\s*(?:and\s+|&\s*)?|\s*&\s*|\s+and\s+", t)) if y]
            if len(listed) >= 2 and len(fit) > 1:
                def names_fit(r):
                    sur = [norm_name(x) for x in entry_surnames(r["text"])]
                    if len(sur) >= len(listed):                 # (in their order: "(Agyeman, Quaye, et al.)" has Quaye second)
                        return all(norm_name(n) == sur[i] for i, n in enumerate(listed) if i)
                    return all(norm_name(n) in norm_name(re.split(r"\(|[“\"‘]", r["text"], maxsplit=1)[0]) for n in listed[1:])
                fit = [r for r in fit if names_fit(r)] or fit
            cands = fit or cands
        if cands:
            c["ref"] = cands[0]["n"]
    return cites


_ACRONYMS = {"unicef": "unitednationschildrensfund", "unesco": "unitednationseducationalscientificandculturalorganization",
             "unaids": "jointunitednationsprogrammeonhivaids", "usaid": "unitedstatesagencyforinternationaldevelopment",
             "unhcr": "unitednationshighcommissionerforrefugees", "undp": "unitednationsdevelopmentprogramme"}


def _is_acronym_of(abbr, name):
    """6.6, round 41: "(WHO, 2023)" is the list's "World Health Organization. (2023)": a group cited by its initials, with
    or without its small words ("MOH": Ministry of Health; "CDC": Centers for Disease Control and Prevention), or by a
    name everyone shortens otherwise (UNICEF)."""
    words = re.findall(r"[A-Za-z][\w'’\-]*", str(name or ""))
    if len(words) < 2:
        return False
    ab = abbr.lower().replace("&", "")
    if _ACRONYMS.get(ab) == norm_name(name):
        return True
    small = ("of", "for", "and", "the", "on", "in")
    full = "".join(w[0].lower() for w in words)
    core = "".join(w[0].lower() for w in words if w.lower() not in small)
    return ab in (full, core) or (len(ab) >= 3 and (core.startswith(ab) or full.startswith(ab)))


def _shape(entry):
    """How many authors an entry names, roughly: "one", "two" or "many" (MLA's "et al." for three or more)."""
    head = re.split(r"[“\"‘]|\((?:19|20)\d{2}|\.\s(?:19|20)\d{2}\.", re.sub(r"^\s*(?:\[\d{1,3}\]|\d{1,3}[.)])\s*", "", str(entry or "")), maxsplit=1)[0]
    if re.search(r"\bet\s+al\b", head) or len(re.findall(r",\s", head)) >= 4:
        return "many"
    return "two" if re.search(r"\s(?:and|&)\s", head) else "one"


# ---------------------------------------------------------------- a reference against its registry record
def _words(t):
    return set(re.findall(r"[a-z0-9]{3,}", unicodedata.normalize("NFKD", str(t or "")).encode("ascii", "ignore").decode().lower()))


def title_in(entry, title):
    """The record's title is the entry's: most of its words (70%) are in the entry."""
    tw = _words(title)
    if not tw:
        return False
    return len(tw & _words(entry)) >= max(2, round(0.7 * len(tw)))


def _abbrev_of(t, w):
    """t is w, or w shortened as journal names are ("int" international, "natl" national, "nurs" nursing, "res" research):
    its first letter, then its letters in order."""
    if t == w:
        return True
    if len(t) < 3 or t[0] != w[0] or len(t) > len(w):
        return False
    it = iter(w)
    return all(ch in it for ch in t)


def _same_journal(journal, entry, title=""):
    """The entry names the record's journal: most of the journal's words (60%, both of two) are in the entry, written out
    or shortened as Vancouver and IEEE shorten them ("Int J Nurs Stud"). (6.6, round 27) Words of the paper's own title
    don't count: a "Journal of Rest Studies" entry for a paper called "Sleep among nurses" is not the "Journal of Sleep
    Research"; where the entry has the title, only what follows it is read."""
    import math
    jw = _words(journal) - {"the", "and", "journal", "for", "of"}
    if not jw:
        return True
    flat = lambda t: " ".join(re.findall(r"[a-z0-9]+", unicodedata.normalize("NFKD", str(t or "")).encode("ascii", "ignore").decode().lower()))
    ft, fe = flat(title), flat(entry)
    if len(ft) >= 12 and ft in fe:
        ew, want = _words(fe[fe.index(ft) + len(ft):]), jw           # (what follows the title: the journal, its volume, pages)
    else:
        tw = _words(title)
        own = jw - tw
        ew, want = (_words(entry) - tw, own) if own else (_words(entry), jw)
    hit = sum(1 for w in want if any(_abbrev_of(t, w) for t in ew))
    return hit >= max(1, math.ceil(0.6 * len(want)))


def _after_title(entry, title):
    """What an entry gives after the paper's title (its journal, volume, issue, pages), so that a title's own numbers
    ("aged 15–19", "COVID-19 (SARS-CoV-2)") are never read as its pages or volume; the whole entry when the title isn't
    found in it."""
    words = re.findall(r"\w+", re.sub(r"<[^>]+>", " ", str(title or "")))
    if len(words) < 2:
        return str(entry or "")
    rx = r"\W+".join(re.escape(w) for w in words[-4:])
    m = None
    for m in re.finditer(r"\b" + rx + r"\b", str(entry or ""), re.I):
        pass
    return str(entry or "")[m.end():] if m else str(entry or "")


def _not_year(x, year):
    return not (len(x) == 4 and x.isdigit() and (x.startswith(("19", "20"))) and (not year or int(x) == int(year)))


def _entry_volume(ev, year=None):
    """The volume an entry gives (6.6, round 27): "vol. 8" (MLA, IEEE), "8(2)" (APA, Harvard, Vancouver, Chicago with an
    issue), ";8:" (Vancouver), ", 8, 100–120" (APA, Harvard without an issue), "Journal 8: 100" (Chicago); "" when none."""
    for rx in (r"\bvol\.?\s*(\d{1,4})\b", r"(?<![\d.\-])(\d{1,4})\*?\s*\(\s*(?:\d{1,3}(?:\s*[–\-/]\s*\d{1,3})?|Suppl\.?\s*\d*|S\d+)\s*\)",
               r";\s*(\d{1,4})\s*:",
               r",\s*\*?(\d{1,4})\*?\s*,\s*(?:pp?\.\s*)?(?:Article\s+)?e?\d", r"[A-Za-z]\s+(\d{1,4})\s*:\s*e?\d"):
        for m in re.finditer(rx, ev, re.I):
            if _not_year(m.group(1), year):
                return m.group(1)
    return ""


def _entry_issue(ev):
    """The issue an entry gives (6.6, round 42): "8(2)" (APA, Harvard, Vancouver), "8 (2)" (Chicago), "no. 2" (MLA,
    IEEE); "" when none."""
    m = re.search(r"\bno\.\s*(\d{1,4}|[A-Za-z]?\d{1,4}[A-Za-z]?)\b", ev) or \
        re.search(r"(?<![\d.\-])\d{1,4}\*?\s*\(\s*(\d{1,3}(?:\s*[–\-/]\s*\d{1,3})?|Suppl\.?\s*\d*|S\d+)\s*\)", ev)
    return re.sub(r"\s+", "", m.group(1)) if m else ""


def _entry_first_page(ev, year=None):
    """The first page (or article number) an entry gives (6.6, round 27): a range "100–120", "pp. 100", ":100" (Vancouver),
    ", 1234." after the volume (APA), "Article e0231234", "Art. no. 1234"; "" when none."""
    t = re.sub(r"\b(?:19|20)\d{2}[a-z]?\b", " ", ev)
    m = re.search(r"\b(?:pp?\.\s*)?(\d{1,6})\s*[–\-]\s*\d{1,6}\b", t)
    if m:
        return m.group(1)
    for rx in (r"\b(?:Article|Art\.\s*no\.)\s*(e?\d{1,9})\b", r"\bpp?\.\s*(e?\d{1,9})\b", r"\d\s*(?:\(\s*[\w–\-/]+\s*\))?\s*:\s*(e?\d{1,9})\b",
               r"(?<![\d.])\d{1,4}\*?\s*(?:\(\s*[\w–\-/]+\s*\))?\s*,\s*(e?\d{1,9})\b(?=\s*(?:\.|$|https?:))"):
        m = re.search(rx, t, re.I)
        if m:
            return m.group(1)
    return ""


def compare(ref, rec):
    """What the entry says differently from the publisher's record, in plain words: [sentence, …]."""
    out = []
    e = ref["text"]
    ry = rec.get("year")
    if ry and ref.get("year") and ref["year"] != ry:
        oy = rec.get("online_year")
        if oy and oy != ry and ref["year"] == oy:      # (6.6, round 44) the year it came out online, not a wrong year
            out.append(f"Your list gives {oy}, the year it came out online; the issue it is in is from {ry}. Give the issue’s year "
                       f"if your style asks for it (APA does).")
        else:
            out.append(f"The year is {ry} in the publisher’s record ({ref['year']} in your list).")
    jn = str(rec.get("journal") or "").strip()
    if jn and rec.get("type") == "journal-article" and not _same_journal(jn, e, rec.get("title")):
        out.append(f"The journal is {jn} in the publisher’s record.")
    # (6.6, round 27) the entry's own volume and pages, never a DOI's or a link's numbers ("s12889-021-10212-9",
    # "S1389-9457(21)00123-4"); a volume without an issue (", 81, 100–120") and a single page or article number are read too
    ev = re.sub(r"https?://\S+|\bdoi:\s*\S+|\b10\.\d{4,9}/\S+", " ", _after_title(e, rec.get("title")), flags=re.I)
    vol = str(rec.get("volume") or "").strip()
    said_v = _entry_volume(ev, ref.get("year"))
    if vol and said_v and said_v != vol:
        out.append(f"The volume is {vol} in the publisher’s record.")
    # (6.6, round 42) the issue: the one the entry gives, or none where the record has one (every style gives it)
    iss = str(rec.get("issue") or "").strip()
    said_i = _entry_issue(ev)
    if iss and said_i and said_i.lower() != iss.lower() and said_v == vol:
        out.append(f"The issue is {iss} in the publisher’s record.")
    elif iss and not said_i and said_v and said_v == vol and rec.get("type") == "journal-article" and re.fullmatch(r"\d{1,4}", iss):
        out.append(f"Your entry gives no issue: the publisher’s record has volume {vol}, issue {iss}.")
    first = str(rec.get("first_page") or "").strip()
    said_p = _entry_first_page(ev, ref.get("year"))
    if first and said_p and said_p.lower().lstrip("e") != first.lower().lstrip("e"):
        last = str(rec.get("last_page") or "").strip()
        out.append(f"The pages are {first}–{last} in the publisher’s record." if last else f"The first page is {first} in the publisher’s record.")
    fa = norm_name(rec.get("first_surname"))
    sn = norm_name(ref.get("surname"))
    if fa and sn and fa not in sn and sn not in fa:
        out.append(f"The first author is {rec.get('first_surname')} in the publisher’s record.")
    elif fa and sn and fa != sn and min(len(fa), len(sn)) >= 4 and _edits(fa, sn) <= 2:
        # (6.6, round 40) "Mensa" for "Mensah": a letter or two off, which the test above let through (a part of a
        # double name, "Mensah-Bonsu" for "Mensah", stays allowed)
        out.append(f"The first author is {rec.get('first_surname')} in the publisher’s record: check the spelling.")
    # 6.6, round 16: what the entry leaves out that the record gives (a journal article's volume and pages)
    if rec.get("type") == "journal-article":
        bare = re.sub(r"https?://\S+|\bdoi:\s*\S+|\b10\.\d{4,9}/\S+|\b(?:19|20)\d{2}[a-z]?\b", " ", e)
        nums = set(re.findall(r"\b\d{1,6}\b", bare))
        said_vol = any(x.startswith("The volume is") for x in out) or bool(said_v)
        if vol and vol.isdigit() and vol not in nums and not said_vol:
            out.append(f"Your entry gives no volume: the publisher’s record has volume {vol}.")
        last = str(rec.get("last_page") or "").strip()
        said_pg = any(x.startswith(("The pages are", "The first page is")) for x in out) or bool(said_p)
        if first.isdigit() and first not in nums and not said_pg:
            out.append(f"Your entry gives no pages: the publisher’s record has {first}–{last}." if last.isdigit()
                       else f"Your entry gives no page: the publisher’s record has {first}.")
    return out


# 6.6, round 31: a source most universities don't accept, or a link to a copy rather than to the paper
_SRC_NOTES = (
    (re.compile(r"\bwikipedia\.org\b|\bIn\s+Wikipedia\b|\bWikipedia,?\s+the\s+free\s+encyclopedia\b|[.,]\s*Wikipedia\.?\s*(?:https?://|Retrieved|$)", re.I),
     "Wikipedia is an encyclopedia anyone can edit: most universities ask you to cite the source it draws on instead."),
    (re.compile(r"\b(?:researchgate\.net|academia\.edu)\b", re.I),
     "This links to a copy shared on a researchers’ site: cite the published paper (its journal, and its DOI if it has one)."),
    (re.compile(r"\bscholar\.google\.[a-z.]+\b", re.I),
     "This links to a Google Scholar page, not to the work: give the publisher’s page or the DOI."),
    (re.compile(r"\bsci-?hub\.[a-z]{2,}|\blibgen\.[a-z]{2,}|\blibrary\.lol\b", re.I),
     "This links to an unauthorised copy: give the publisher’s page or the DOI instead."))


# 6.6, round 80 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): an entry dated in a year still to come, which no
# registry could confirm. A work accepted but not yet out is written "in press" (APA 7, Cite Them Right; "forthcoming" in
# MLA and Chicago); a journal that has already placed a paper in next year's issue is confirmed by its record, so nothing is
# said then.
_AHEAD = re.compile(r"\b(?:in\s+press|forthcoming|advance\s+(?:online\s+)?(?:publication|access)|(?:online\s+)?ahead\s+of\s+print|"
                    r"early\s+(?:view|access|online)|epub\b)", re.I)


def future_note(ref, this_year):
    """"" or the line for an entry whose year is still to come."""
    y = ref.get("year")
    if not isinstance(y, int) or y <= this_year or _AHEAD.search(str(ref.get("text") or "")):
        return ""
    return f"{y} is a year still to come: check the year (a work accepted but not yet out is written “in press”)."


def source_note(entry):
    """What a reader of the list would question about the source itself ("" when nothing)."""
    for rx, note in _SRC_NOTES:
        if rx.search(str(entry or "")):
            return note
    return ""


def offer_doi(ref, rec):
    """The record's DOI when the entry gives none (an addition, not a difference), or ""."""
    if (rec or {}).get("doi") and not ref.get("doi"):
        return f"Its DOI is https://doi.org/{rec['doi']} (your list doesn’t give it)."
    return ""


# ---------------------------------------------------------------- 6.6, round 7: the citations themselves
_QUOTE = re.compile(r"[“\"]([^”\"]{12,400})[”\"]")
_LOCATOR = re.compile(r"\b(?:pp?|para|paras|chap|sec)\.\s*[\w–\-]+|"                     # APA, Harvard: p. 4, pp. 4–5, para. 2
                      r"(?<![\d,])\s(?!(?:19|20)\d{2}[a-z]?\b)\d{1,5}(?:\s*[–\-]\s*\d{1,5})?\s*\)\s*$|"   # MLA: (Mensah 45)
                      r",\s*(?!(?:19|20)\d{2}[a-z]?\b)\d{1,5}(?:\s*[–\-]\s*\d{1,5})?\s*\)\s*$", re.I)          # Chicago: (Mensah 2021, 45)


# 6.6, round 89 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): APA 7 (8.27) sets a quotation of 40 words or more
# as a block quotation: on its own lines, indented, without quotation marks, its citation after the final full stop.
_LONG_Q = re.compile(r"[“\"]([^”\"]{150,4000})[”\"]")


def long_quotes(body, limit=40):
    """The first words of each quotation of `limit` words or more written inside quotation marks."""
    out = []
    for m in _LONG_Q.finditer(str(body or "")):
        words = re.findall(r"[\w’'\-]+", m.group(1))
        if len(words) >= limit:
            out.append(" ".join(words[:8]))
    return out


# 6.6, round 90: APA 7 (8.12) lists several works in one bracket as the reference list does: alphabetically by the first
# author's surname, one author's works by year, earliest first ("(Mensah, 2021; Osei, 2019)"). A bracket with "see also",
# "e.g." or a secondary citation is left as written (APA lets those go their own way).
def order_notes(body, cites):
    """[{"cite": the bracket as written, "note": it in APA 7's order}] for brackets of several works out of order."""
    groups = {}
    for c in cites:
        if c.get("num") or c.get("mla") or not str(c.get("text") or "").startswith("(") or "start" not in c:
            continue
        groups.setdefault((c["start"], c["end"]), []).append(c)
    out = []
    for (a, b), cs in sorted(groups.items()):
        inside = str(body or "")[a + 1:b - 1]
        parts = [x.strip() for x in re.split(r"\s*;\s*", inside) if x.strip()]
        if len(parts) < 2 or re.search(r"\b(?:see|e\.g\.|cf\.|as\s+cited\s+in|for\s+example|also)\b", inside, re.I):
            continue
        keyed = []
        for prt in parts:
            m = re.match(r"^(?P<names>[^\d()]+?),?\s+(?P<y>(?:19|20)\d{2}[a-z]?|n\.\s?d\.|in\s+press)", prt, re.I)
            if not m:
                keyed = []
                break
            y = m.group("y")
            keyed.append(((norm_name(m.group("names")), 0 if y[:1].isdigit() is False and y.lower().startswith("n") else 1,
                           y if y[:1].isdigit() else ("0000" if y.lower().startswith("n") else "9999")), prt))
        if len(keyed) != len(parts):
            continue
        ordered = [prt for _k, prt in sorted(keyed, key=lambda kp: kp[0])]
        if ordered != parts:
            merged = []                                  # one author's works: the names once, the years after ("Mensah, 2019, 2021")
            for prt in ordered:
                m2 = re.match(r"^(?P<names>[^\d()]+?),?\s+(?P<rest>(?:19|20)\d{2}.*|n\.\s?d\..*|in\s+press.*)$", prt, re.I)
                if merged and m2 and m2.group("names").strip() == merged[-1][0]:
                    merged[-1] = (merged[-1][0], merged[-1][1] + ", " + m2.group("rest"))
                else:
                    merged.append((m2.group("names").strip(), m2.group("rest")) if m2 else (prt, ""))
            ordered = [f"{n}, {r}" if r else n for n, r in merged]
            out.append({"cite": "(" + "; ".join(parts) + ")",
                        "note": "APA 7 lists the works in one bracket as the reference list does (by the first author’s surname, then year): ("
                                + "; ".join(ordered) + ")."})
    return out


# 6.6, round 91 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): APA 7 (8.21) writes a group's full name at its
# first citation, with the abbreviation in square brackets, "(World Health Organization [WHO], 2023)", and the abbreviation
# after that, "(WHO, 2023)". A first citation by the abbreviation alone leaves the reader to guess the group (the list gives
# the full name).
_ABBR_CITE = re.compile(r"^\(?(?P<abbr>[A-Z][A-Z&]{1,9})(?:,\s*|\s+)(?:(?:19|20)\d{2}|n\.\s?d\.)")
_GROUP_HEAD = re.compile(r"^\s*(?P<g>[A-Z][^.(]{6,140}?)\.\s*\(")


def group_abbr_notes(cites, refs):
    """[{"cite", "note"}] for a group cited first by its abbreviation alone, where the list gives its full name."""
    by_n = {r["n"]: r for r in refs}
    seen, out = set(), []
    for c in sorted((c for c in cites if c.get("ref")), key=lambda c: c.get("start", 0)):
        if c["ref"] in seen:
            continue
        seen.add(c["ref"])
        t = str(c.get("text") or "").strip()
        m = _ABBR_CITE.match(t)
        g = _GROUP_HEAD.match(str(by_n.get(c["ref"], {}).get("text") or ""))
        if not m or not g or "[" in t or not _is_acronym_of(m.group("abbr"), g.group("g")):
            continue
        name, abbr = g.group("g").strip(), m.group("abbr")
        full = t.replace(abbr, f"{name} [{abbr}]", 1)
        out.append({"cite": t, "note": f"APA 7 gives a group’s full name at its first citation, with its abbreviation: {full}; after that, {t}."})
    return out


# 6.6, round 92: a secondary citation lists only the work read (APA 7 8.6): "(Mensah, 2019, as cited in Osei, 2021)" has Osei
# (2021) in the list and not Mensah (2019), unless the writer read Mensah too (then it is cited directly).
_PRIMARY = re.compile(r"^\(?\s*(?P<names>[^\d()]+?),?\s+(?P<y>(?:19|20)\d{2})[a-z]?\s*,?\s*as\s+cited\s+in\b", re.I)


def secondary_notes(cites, refs):
    """[{"cite", "note"}] for a secondary citation whose original work is also in the list."""
    out, seen = [], set()
    for c in cites:
        t = str(c.get("text") or "")
        if t in seen or not re.search(r"\bas\s+cited\s+in\b", t, re.I):
            continue
        seen.add(t)
        m = _PRIMARY.match(t)
        if not m:
            continue
        sur, y = _surname_of(m.group("names")), int(m.group("y"))
        if any(norm_name(r.get("surname")) == norm_name(sur) and r.get("year") == y for r in refs):
            out.append({"cite": t[:120], "note": f"You read {sur} ({y}) in {c.get('surname')} ({c.get('year')}): APA 7 lists only the work you read, "
                                                 f"so {sur} ({y}) has no entry of its own (if you read it yourself, cite it directly instead)."})
    return out


def quote_needs_page(sentence, cites):
    """A direct quotation (four words or more in quotation marks) whose citation gives no page: every style asks for it."""
    m = _QUOTE.search(str(sentence or ""))
    if not m or len(m.group(1).split()) < 4 or not cites:
        return False
    return not any(_LOCATOR.search(str(c.get("text") or "")) for c in cites)


ETAL_FROM = {"apa": 3, "mla": 3, "harvard": 4, "chicago": 3}       # from how many authors the text gives "et al."
# (Chicago 18, 2024, the current edition: from three; the 17th named all three, and a list written to it is not told wrong)


def etal_note(cite, n_authors, style):
    """What the style asks of a citation's names, given how many authors the paper has ("" when it is right)."""
    t = str(cite.get("text") or "")
    k = ETAL_FROM.get(style)
    if not k or not n_authors or cite.get("num"):
        return ""
    names = re.sub(r"\(|\)|,?\s*(?:(?:19|20)\d{2}[a-z]?|n\.\s?d\.|no\s+date|in\s+press|forthcoming).*$|\s+(?:19|20)\d{2}.*$", "", t, flags=re.I).strip()
    has_etal = bool(re.search(r"\bet\s+al\b", t))
    listed = len([x for x in re.split(r"\s*,\s*|\s*&\s*|\s+and\s+", names) if x.strip() and not re.search(r"\bet\s+al\b", x)])
    if has_etal and n_authors < k:
        return (f"The paper has {n_authors} author{'' if n_authors == 1 else 's'}: name {'it' if n_authors == 1 else 'both' if n_authors == 2 else 'all three'} "
                f"instead of “et al.”.")
    if not has_etal and n_authors >= (4 if style == "chicago" else k) and listed >= 2:
        return f"The paper has {n_authors} authors: the first author’s name and “et al.” are enough."
    return ""


def amp_note(cite, n_authors, style):
    """6.6, round 30: how two or more authors' names are joined. APA 7: "&" inside brackets, "and" in a sentence; MLA 9
    and Chicago: "and" always. Only for two people's names (n_authors == 2): a group's own name may hold "and" ("Ministry of
    Health and Social Welfare"). "" when it is right."""
    t = str(cite.get("text") or "")
    if n_authors != 2 or cite.get("num") or style not in ("apa", "mla", "chicago", "harvard"):
        return ""                               # (three or more: "et al.", which etal_note asks for)
    names = re.sub(r"\(|\)|,?\s*(?:(?:19|20)\d{2}[a-z]?|n\.\s?d\.|no\s+date|in\s+press|forthcoming).*$|\s+(?:19|20)\d{2}.*$|"
                   r"\s+\d{1,5}(?:\s*[–\-]\s*\d{1,5})?\s*\)?$", "", t, flags=re.I).strip()
    if re.search(r"\bet\s+al\b", names):
        return ""
    bracketed = t.lstrip().startswith("(")
    amp, word = bool(re.search(r"\s&\s", names)), bool(re.search(r"\s(?:and)\s", names))
    if style == "apa" and bracketed and word and not amp:
        return "In APA 7, “&” joins the authors’ names inside brackets: " + re.sub(r"\s+and\s+", " & ", t, count=1) + "."
    if style == "apa" and not bracketed and amp:
        return "In APA 7, “and” joins the authors’ names in a sentence: " + re.sub(r"\s+&\s+", " and ", t, count=1) + "."
    if style in ("mla", "chicago", "harvard") and amp:          # (6.6, round 81: Cite Them Right too, in brackets and in a sentence)
        return {"mla": "MLA 9", "chicago": "Chicago", "harvard": "Cite Them Right"}[style] + " joins the authors’ names with “and”: " \
            + re.sub(r"\s+&\s+", " and ", t, count=1) + "."
    return ""


# 6.6, round 68 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): how a citation in brackets separates the name
# from the year. APA 7 (8.11) and Cite Them Right put a comma, "(Mensah, 2021)"; Chicago's author-date (CMOS 18) none,
# "(Mensah 2021)". A citation in a sentence, "Mensah (2021)", has none in any of them; MLA cites a page, not a year.
_NAME_YEAR = re.compile(r"^\(\s*(?P<names>[^()]*?[A-Za-z\u00c0-\u024f.’'\]])(?P<sep>\s*,\s*|\s+)(?P<year>(?:19|20)\d{2}[a-z]?|"
                        r"n\.\s?d\.(?:-[a-z])?|no\s+date|in\s+press(?:-[a-z])?|forthcoming)(?P<rest>[^()]*)\)$", re.I)
COMMA_RULE = {"apa": "APA 7 puts a comma between the name and the year", "harvard": "Cite Them Right puts a comma between the name and the year",
              "chicago": "Chicago puts no comma between the name and the year"}


# 6.6, round 72: "et al." as every style writes it: one full stop, after "al" ("et. al.", "et.al", "et al" and "Et al." are not)
_ETAL_BAD = re.compile(r"\b[Ee]t\.\s?al\b\.?|\b[Ee]t\s?al\b(?!\.)|\bEt\s?al\b\.?")
ETAL_RULE = "“et al.” is written with one full stop, after “al”"


def etal_form(t):
    """The text with its "et al." as the styles write it."""
    return re.sub(r"\b[Ee]t\.?\s?al\b\.?", "et al.", str(t or ""))


def etal_punct_fix(cite):
    """The citation with "et al." as every style writes it, when it is written otherwise ("" when right)."""
    t = str(cite.get("text") or "")
    if cite.get("num") or not _ETAL_BAD.search(t):
        return ""
    return etal_form(t)


def comma_fix(cite, style):
    """The citation as the list's style writes it, when it separates the name from the year the other way ("" when right)."""
    t = str(cite.get("text") or "").strip()
    if cite.get("num") or cite.get("mla") or style not in COMMA_RULE or ";" in t or _AS_CITED.search(t):
        return ""
    m = _NAME_YEAR.match(t)
    if not m or not re.search(r"[A-ZÀ-Þ]", m.group("names")):
        return ""
    names, comma = etal_form(m.group("names").strip()), "," in m.group("sep")         # (round 72: its "et al." as written by the styles)
    if style == "chicago":
        return f"({names} {m.group('year')}{m.group('rest')})" if comma else ""
    return "" if comma else f"({names}, {m.group('year')}{m.group('rest')})"


def comma_note(cite, style):
    """The note for one citation ("" when it is right)."""
    f = comma_fix(cite, style)
    return f"{COMMA_RULE[style]}: {f}." if f else ""


def comma_many(texts, style, example):
    """One line for three or more citations that separate the name from the year the other way, so the same note isn't
    said again and again: {"cite": the first three and how many more, "note": the rule, with the first one as the style
    writes it}."""
    more = len(texts) - 3
    shown = ", ".join(t if len(t) <= 60 else t[:57].rstrip() + "…)" for t in texts[:3])
    return {"cite": (shown + (f" and {more} more" if more > 0 else ""))[:200], "note": f"{COMMA_RULE[style]} in brackets, as in {example}."}


def ambiguity_note(cite, refs, style):
    """6.6, round 45: APA 7 (8.18): two works that would both be "(Osei et al., 2023)" are told apart by naming as many
    authors as it takes: "(Osei, Darko, et al., 2023)". Round 47: Chicago the same way (CMOS 15.28: "(Osei, Darko, et al.
    2023)"); Cite Them Right, which names the first author alone from four authors, by a letter after the year in the text
    and in the list ("(Osei et al., 2023a)"). "" when the citation is not one of two such."""
    if style not in ("apa", "chicago", "harvard") or cite.get("num") or not cite.get("year"):
        return ""
    t = re.sub(r"\(|\)|,?\s*(?:(?:19|20)\d{2}[a-z]?).*$", "", str(cite.get("text") or "")).strip()
    if not re.search(r"\bet\s+al\b", t) or len([x for x in re.split(r"\s*,\s*|\s*&\s*|\s+and\s+", re.sub(r",?\s*\bet\s+al\b\.?", "", t))
                                                 if x.strip()]) != 1:
        return ""
    if re.search(r"\b(?:19|20)\d{2}[a-z]\b", str(cite.get("text") or "")):
        return ""                                        # ("2023a": told apart by its letter)
    sn = norm_name(cite.get("surname"))
    k = ETAL_FROM.get(style, 3)

    def many(r):
        n = entry_authors(r["text"])
        return n >= 100 or n >= k if n else _shape(r["text"]) == "many"
    same = [r for r in refs if norm_name(r.get("surname")) == sn and r.get("year") == cite["year"] and many(r)]
    ini = initials_of_cite(cite.get("text") or "")
    if ini:                                             # ("(Y. Mensah et al., 2021)": its initials tell it apart, round 46)
        same = [r for r in same if initials_of_entry(r["text"], r.get("surname")).startswith(ini)]
    if len(same) < 2:
        return ""
    n = {2: "Two", 3: "Three", 4: "Four", 5: "Five"}.get(len(same), str(len(same)))
    sep = " " if style == "chicago" else ", "
    if style == "harvard":
        return (f"{n} works in your list would be “{cite.get('surname')} et al., {cite['year']}”: Cite Them Right tells them apart "
                f"with a letter after the year, in the text and in the list ({cite['year']}a, {cite['year']}b).")
    sur = entry_surnames(same[0]["text"])
    eg = f", as in ({cite.get('surname')}, {sur[1]}, et al.{sep}{cite['year']})" if len(sur) >= 2 else ""
    return (f"{n} works in your list would be “{cite.get('surname')} et al.{sep}{cite['year']}”: "
            + ("APA 7 names" if style == "apa" else "Chicago names") + f" as many authors as tell them apart{eg}.")


def _edits(a, b):
    """How many letters differ between two names (Levenshtein's distance)."""
    if a == b:
        return 0
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def near_miss(cite, refs):
    """A citation with no entry whose author IS in the list with another year: what the list has. (6.6, round 13) Or
    whose name is one letter or two from an entry's of the same year: that entry, to check the spelling."""
    sn = norm_name(cite.get("surname"))
    if not sn:
        return ""
    same = [r for r in refs if norm_name(r.get("surname")) == sn and (r.get("year") or r.get("undated"))]
    said = {"nd": "“n.d.”", "press": "“in press”"}
    if same and (cite.get("year") or cite.get("undated")):
        years = sorted({str(r["year"]) if r.get("year") else said[r["undated"]] for r in same})
        return f"Your list has {cite.get('surname')} with {', '.join(years)}, not {cite['year'] if cite.get('year') else said[cite['undated']]}."
    if not same and cite.get("year") and len(sn) >= 4:
        close = [r for r in refs if r.get("year") == cite["year"] and norm_name(r.get("surname"))
                 and _edits(sn, norm_name(r.get("surname"))) <= (2 if len(sn) >= 7 else 1)]
        if len(close) == 1:
            return f"Your list has {close[0].get('surname')} ({cite['year']}): check the spelling."
    return ""


_PAGE_OF = re.compile(r"\bpp?\.\s*(\d{1,5})(?:\s*[–\-]\s*(\d{1,5}))?|(?<![\d,])\s(\d{1,5})(?:\s*[–\-]\s*(\d{1,5}))?\s*\)\s*$|"
                      r",\s*(?!(?:19|20)\d{2}\b)(\d{1,5})(?:\s*[–\-]\s*(\d{1,5}))?\s*\)\s*$")


def page_note(cite, rec):
    """6.6, round 13: a page the citation gives that the paper doesn't have (its pages in the publisher's record)."""
    m = _PAGE_OF.search(str(cite.get("text") or ""))
    first, last = str((rec or {}).get("first_page") or ""), str((rec or {}).get("last_page") or "")
    if not m or not first.isdigit() or not last.isdigit() or int(last) < int(first):
        return ""
    nums = [int(x) for x in m.groups() if x]
    if not nums or (cite.get("year") and nums[0] == cite["year"]):
        return ""
    if all(int(first) <= n <= int(last) for n in nums):
        return ""
    if len(nums) > 1:
        return f"Pages {nums[0]}–{nums[1]} go outside the paper’s pages in the publisher’s record ({first}–{last})."
    return f"Page {nums[0]} is outside the paper’s pages in the publisher’s record ({first}–{last})."


def entry_surnames(entry):
    """The authors' surnames in the entry's order (6.6, round 46): APA, Harvard ("Osei, A., Quaye, P., & Badu, P."), MLA and
    Chicago ("Osei, Ama, Kofi Quaye, and Kojo Badu"), IEEE ("A. Osei, P. Quaye, and P. Badu"); [] when it can't tell."""
    t = re.sub(r"^\s*(?:\[\d{1,3}\]|\d{1,3}[.)])\s*", "", str(entry or ""))
    head = re.split(r"\((?:19|20)\d{2}|\(n\.\s?d\.|\(in\s+press|\.\s(?:19|20)\d{2}\.|[“\"‘]", t, maxsplit=1)[0]
    h = re.sub(r",?\s*\bet\s+al\b\.?", "", head).strip().rstrip(".,").strip()
    part = r"(?:(?:van|von|de|del|der|den|du|da|di|le|la)\s+)*[A-ZÀ-ɏ][\w'’\-]+"
    if re.match(r"\s*(?:[A-ZÀ-ɏ]\.\s?(?:-?[A-ZÀ-ɏ]\.\s?)*)\s?" + part, h):                          # IEEE
        return [m for m in re.findall(r"(?:^|,\s*(?:and\s+)?|\s+and\s+)(?:[A-ZÀ-ɏ]\.\s?(?:-?[A-ZÀ-ɏ]\.\s?)*)\s?(" + part + ")", h)]
    apa = re.findall(r"(" + part + r"(?:\s[A-ZÀ-ɏ][\w'’\-]+)?),\s(?:[A-Z]\.\s?-?)+", h + ".")
    if apa:
        return apa
    bits = [x.strip() for x in re.split(r",\s*(?:and\s+)?|\s+and\s+|\s*&\s*", h) if x.strip()]
    if len(bits) >= 2:                                                                                    # MLA, Chicago
        return [bits[0]] + [b.split()[-1] for b in bits[2:]]
    return bits[:1]


def entry_authors(entry):
    """How many authors an entry names (APA, Harvard: "Surname, A. B."; Vancouver: "Surname AB"); 0 when it can't tell."""
    t = re.sub(r"^\s*(?:\[\d{1,3}\]|\d{1,3}[.)])\s*", "", str(entry or ""))
    head = re.split(r"\((?:19|20)\d{2}|\.\s(?:19|20)\d{2}\.|[“\"‘]", t, maxsplit=1)[0]
    if re.match(r"\s*(?:[A-ZÀ-ɏ]\.\s?(?:-?[A-ZÀ-ɏ]\.\s?)*)\s?(?:(?:van|von|de|der|den|du|da|di|le|la)\s+)*[A-ZÀ-ɏ][a-zà-ɏ]", head):
        hh = re.sub(r",?\s*\bet\s+al\b\.?", "", head).strip().rstrip(".,")      # (IEEE: "K. A. Mensah, Y. Owusu, and A. Boateng")
        n = len(re.findall(r"(?:^|,\s*(?:and\s+)?|\s+and\s+)(?:[A-ZÀ-ɏ]\.\s?(?:-?[A-ZÀ-ɏ]\.\s?)*)\s?(?:(?:van|von|de|der|den|du|da|di|le|la)\s+)*"
                           r"[A-ZÀ-ɏ][\w'’\-]+", hh))
        if n:
            return n + (100 if re.search(r"\bet\s+al\b", head) else 0)
    apa = re.findall(r"[A-ZÀ-ɏ][\w'’\-]+(?:\s[A-ZÀ-ɏ][\w'’\-]+)?,\s(?:[A-Z]\.\s?-?)+", head)
    if apa:
        # ("Osei, A. et al. (2023)": a list that shortens its authors doesn't say how many there are: many, not one)
        return len(apa) + (1 if re.search(r"\.\s\.\s\.", head) else 0) + (100 if re.search(r"\bet\s+al\b", head) else 0)
    van = re.findall(r"\b[A-Z][\w'’\-]+\s[A-Z]{1,3}(?=,|\.)", head)
    if van:
        return len(van) + (100 if re.search(r"\bet\s+al\b", head) else 0)
    # (6.6, round 46) IEEE: "K. A. Mensah, Y. Owusu, and A. Boateng,"; MLA, Chicago: "Mensah, Kofi, Esi Frimpong, and Akua Appiah"
    etal = 100 if re.search(r"\bet\s+al\b", head) else 0
    h = re.sub(r",?\s*\bet\s+al\b\.?", "", head).strip().rstrip(".,").strip()
    ieee = re.findall(r"(?:^|,\s*(?:and\s+)?|\s+and\s+)(?:[A-ZÀ-ɏ]\.\s?(?:-?[A-ZÀ-ɏ]\.\s?)*)\s?(?:(?:van|von|de|der|den|du|da|di|le|la)\s+)*[A-ZÀ-ɏ][\w'’\-]+", h)
    if ieee and re.match(r"(?:[A-ZÀ-ɏ]\.\s?)+", h):
        return len(ieee) + etal
    part = r"(?:(?:van|von|de|del|der|den|du|da|di|le|la)\s+)*[A-ZÀ-ɏ][\w'’\-]*\.?(?:\s+(?:(?:van|von|de|del|der|den|du|da|di|le|la)\s+)*[A-ZÀ-ɏ][\w'’\-]*\.?)*"
    parts = [x.strip() for x in re.split(r",\s*(?:and\s+)?|\s+and\s+|\s*&\s*", h) if x.strip()]
    if parts and all(re.fullmatch(part, x) for x in parts):
        return (1 if len(parts) == 1 else 1 + (len(parts) - 2)) + etal
    return 0


# ---------------------------------------------------------------- 6.6, round 8: the paper's words, and the list itself
_SMALL = set(("a an the of in on at to for from by with and or but as is are was were be been being this that these those "
              "it its their there which who whom whose than then into over under among between within without about after "
              "before during while when where how what not no nor so such can could may might must shall should will would "
              "do does did has have had also more most less least very we our they them he she his her i").split())


def copied_run(sentence, source, run=8, content=5):
    """The longest stretch of `run` or more words in a row that a sentence shares with its source (outside the sentence's
    own quotation marks) with at least `content` of them more than small words: those words as the sentence writes them,
    or "" when there is none. (Copied words need quotation marks and a page, even with a citation.)"""
    s = str(sentence or "")
    masked = re.sub(r"“[^”]*”|\"[^\"]*\"", lambda m: " " * len(m.group(0)), s)
    toks = [(m.group(0).lower(), m.start(), m.end()) for m in re.finditer(r"[A-Za-z0-9]+", masked)]
    src = re.findall(r"[a-z0-9]+", unicodedata.normalize("NFKC", str(source or "")).lower())
    if len(toks) < run or len(src) < run:
        return ""
    grams = {tuple(src[i:i + run]) for i in range(len(src) - run + 1)}
    hits = [i for i in range(len(toks) - run + 1) if tuple(t for t, _a, _b in toks[i:i + run]) in grams]
    if not hits:
        return ""
    stretches, a, b = [], hits[0], hits[0] + run
    for i in hits[1:]:
        if i <= b - run + 1:
            b = i + run
        else:
            stretches.append((a, b))
            a, b = i, i + run
    stretches.append((a, b))
    for a, b in sorted(stretches, key=lambda x: x[1] - x[0], reverse=True):
        if sum(1 for t, _x, _y in toks[a:b] if t not in _SMALL) >= content:
            return s[toks[a][1]:toks[b - 1][2]]
    return ""


def copied_note(run):
    words = run.split()
    shown = run if len(words) <= 24 else " ".join(words[:12]) + " … " + " ".join(words[-6:])
    return (f"These {len(words)} words are the abstract’s own: “{shown}”. Put them in quotation marks with the page number, "
            "or say it in your own words.")


def _sort_key(r):
    t = re.sub(r"^\s*(?:\[\d{1,3}\]|\d{1,3}[.)])\s*", "", str(r.get("text") or ""))
    who = str(r.get("surname") or "") or t
    who = re.sub(r"^(?:the|a|an)\s+", "", who.strip(), flags=re.I)
    return norm_name(who)


def _same_entry(a, b):
    if a.get("doi") and b.get("doi"):
        return a["doi"] == b["doi"]
    if norm_name(a.get("surname")) != norm_name(b.get("surname")) or a.get("year") != b.get("year") or not a.get("year"):
        return False
    ya = re.search(r"\b" + str(a["year"]) + r"([a-z])\b", a["text"])
    yb = re.search(r"\b" + str(b["year"]) + r"([a-z])\b", b["text"])
    if ya and yb and ya.group(1) != yb.group(1):
        return False                          # (2021a and 2021b: two works of one year)
    wa, wb = _words(a["text"]), _words(b["text"])
    return bool(wa and wb) and len(wa & wb) / len(wa | wb) >= 0.75


def list_notes(refs, cites, numbered, style_name):
    """What the reference list itself gets wrong: the same paper listed twice; an author-date list out of alphabetical
    order; a numbered list not numbered in the order its references are first cited. [sentence, …]"""
    out = []
    lab = (lambda r: f"[{r.get('num') or r['n']}]") if numbered else (lambda r: str(r["n"]))
    for i, r in enumerate(refs):
        for q in refs[:i]:
            if _same_entry(q, r):
                out.append(f"Entries {lab(q)} and {lab(r)} are the same paper: keep one." if numbered
                           else f"Entries {lab(q)} and {lab(r)} of your list are the same paper: keep one.")
                break
    if numbered:
        num_of = {r["n"]: (r.get("num") or r["n"]) for r in refs}
        first = []
        for c in cites:
            if c.get("ref") and c["ref"] not in first:
                first.append(c["ref"])
        seq = [num_of[n] for n in first]
        for x, y in zip(seq, seq[1:]):
            if y < x:
                out.append(f"In {style_name}, references are numbered in the order they are first cited: [{x}] is cited before [{y}].")
                break
    else:
        keys = [(_sort_key(r), r) for r in refs]
        # (6.6, round 43) "van den Berg", "de Souza": listed under the particle or under the main name, as people and
        # styles differ; either way is right, so a pair is out of order only when it is so both ways
        def ka(r, k):
            main = norm_name(re.sub(r"^\s*(?:(?:van|von|de|del|della|der|den|du|da|di|le|la|ten|ter|dos|das)\s+)+", "", str(r.get("surname") or "")))
            return {k, main or k}
        bad = [(x, y) for x, y in zip(keys, keys[1:]) if x[0] and y[0] and all(b < a for a in ka(x[1], x[0]) for b in ka(y[1], y[0]))]
        if bad:
            (_k1, r1), (_k2, r2) = bad[0]
            n1 = str(r1.get("surname") or "").strip() or r1["text"][:40]
            n2 = str(r2.get("surname") or "").strip() or r2["text"][:40]
            more = f" (in {len(bad)} places)" if len(bad) > 1 else ""
            out.append(f"{style_name} lists references in alphabetical order of the first author’s surname{more}: "
                       f"{n2} comes before {n1}.")
    return out


# 6.6, round 88 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): two works by the same authors in the same year are
# told apart by a letter after the year, in the list and in the text (APA 7 8.19, Cite Them Right, CMOS 18 13.106): without
# it "(Mensah, 2021)" points to either, and a reader (or this check) can't tell which one a sentence rests on. MLA tells
# them apart by a short title instead (round 28), so it isn't asked there.
def _author_head(entry):
    """The authors as the entry writes them (before the year or the title), in a form two entries can be compared in."""
    t = re.sub(r"^\s*(?:\[\d{1,3}\]|\d{1,3}[.)])\s*", "", str(entry or ""))
    head = re.split(r'\((?:19|20)\d{2}|\(n\.\s?d\.|\(in\s+press|\.\s(?:19|20)\d{2}[a-z]?\.|[“"‘]', t, maxsplit=1)[0]
    return re.sub(r"[^a-z]+", " ", head.lower()).strip()


def same_year_note(refs, style):
    """"" or the line for the first pair of entries with the same authors and year and no letters to tell them apart."""
    if style not in ("apa", "harvard", "chicago"):
        return ""
    for i, r in enumerate(refs):
        for q in refs[:i]:
            if not r.get("year") or r.get("year") != q.get("year") or norm_name(r.get("surname")) != norm_name(q.get("surname")):
                continue
            if _same_entry(q, r):
                continue                                    # (the same paper twice: list_notes says so)
            a, b = _author_head(q["text"]), _author_head(r["text"])
            if not a or a != b:
                continue                                    # (other co-authors, or other initials, tell them apart in the text)
            y = str(r["year"])
            if re.search(r"\b" + y + r"[a-z]\b", q["text"]) or re.search(r"\b" + y + r"[a-z]\b", r["text"]):
                continue
            who = str(q.get("surname") or "").strip() or q["text"][:40]
            return (f"Entries {q['n']} and {r['n']} are both {who}, {y}: {_STYLE_NAME[style]} tells them apart with a letter after the year, "
                    f"in the list and in the text ({y}a, {y}b), in the order of their titles.")
    return ""


# ---------------------------------------------------------------- 6.6, round 9: an entry as APA 7 writes it
_APA_HEAD = re.compile(r"^(?P<head>.*?)\(\s*(?:(?:19|20)\d{2}[a-z]?|n\.\s?d\.|in\s+press)", re.S | re.I)


# ---------------------------------------------------------------- 6.6, round 51: how a title is capitalised
_TITLE_SMALL = set("a an the and but or nor for so yet as at by in of off on per to up via vs with from into onto over upon than versus "
                   "among amongst between through during without within against across".split())


def entry_title(entry, style):
    """The title of a work as the entry writes it ("" when it can't tell): in quotation marks (Cite Them Right's ‘…’, MLA,
    Chicago, IEEE), after the year (APA, Harvard), or after the authors (Vancouver)."""
    t = re.sub(r"^\s*(?:\[\d{1,3}\]|\d{1,3}[.)])\s*", "", str(entry or "")).replace("*", "")
    m = re.search(r"[‘“\"]([^’”\"]{8,300})[’”\"]", t)
    if m and style in ("harvard", "mla", "chicago", "ieee"):
        return m.group(1).strip().rstrip(".,")
    if style in ("apa", "harvard"):
        m = re.search(r"\((?:19|20)\d{2}[a-z]?\)\.?\s+(.{8,300}?)(?:[.?!](?:\s|$)|,\s+\d)", t)
        return m.group(1).strip() if m else ""
    if style == "vancouver":
        m = re.match(r"^.+?[A-Z]{1,3}(?:,? et al)?\.\s+(.{8,300}?)[.?!]\s", t)
        return m.group(1).strip() if m else ""
    return ""


def _forms(text):
    out = {}
    for sent in re.split(r"(?<=[.!?:;])\s+|\n+", str(text or "")):
        for k, m in enumerate(re.finditer(r"[A-Za-z\u00C0-\u024F][a-z\u00DF-\u024F'’]+", sent)):
            if k:
                out.setdefault(m.group(0).lower(), set()).add(m.group(0))
    return out


def title_note(entry, style, evidence=""):
    """6.6, round 51: a title capitalised the other way from the style: APA 7, Cite Them Right, Vancouver and IEEE write an
    article's or a book's title in sentence case; MLA 9 and Chicago capitalise its principal words. In sentence case a word
    counts only where the person's own writing (or the paper's abstract) writes it in lower case inside a sentence (so a
    name is never called a fault). "" when it is right, or can't be told."""
    title = entry_title(entry, style)
    words = re.findall(r"[A-Za-z\u00C0-\u024F][A-Za-z\u00C0-\u024F'’\-]*", title)
    if len(words) < 4:
        return ""
    if style in ("mla", "chicago"):
        low = [w for i, w in enumerate(words) if i and len(w) > 3 and w.islower() and w.lower() not in _TITLE_SMALL]
        if len(low) >= 2:
            return (("MLA 9" if style == "mla" else "Chicago") + " capitalises a title’s principal words (“" + " ".join(
                (w.lower() if w.lower() in _TITLE_SMALL and i else w[:1].upper() + w[1:]) for i, w in enumerate(words[:6])) + "…”).")
        return ""
    if style not in ("apa", "harvard", "vancouver", "ieee"):
        return ""
    forms = _forms(evidence)
    after_colon = {i + 1 for i, w in enumerate(words[:-1]) if re.search(re.escape(w) + r"\s*[:?!]", title)}
    capped = [w for i, w in enumerate(words) if i and i not in after_colon and len(w) > 3 and w[:1].isupper() and not w.isupper()
              and not re.search(r"[a-z][A-Z]", w) and "-" not in w
              and w.lower() in forms and all(x.islower() for x in forms[w.lower()])]
    if len(capped) < 2:
        return ""
    name = {"apa": "APA 7", "harvard": "Cite Them Right", "vancouver": "Vancouver", "ieee": "IEEE"}[style]
    rest = "its first word" + (", the first word after a colon" if style == "apa" else "") + " and names"
    return f"{name} writes a title in sentence case: a capital only for {rest} (“{capped[0]}” and “{capped[1]}” here would be lower case)."


# 6.6, round 76 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): an entry of an MLA 9, Chicago or Cite Them Right
# list as its own style writes it, as round 9 does for APA 7. Only what each manual states plainly: the authors joined with
# "and" (MLA 9 5.6, CMOS 18 13.76, Cite Them Right), MLA's "vol. 12, no. 3" and "pp." before a range of pages (MLA 9 5.82,
# 5.88), and Cite Them Right's "Available at:" before a web address or a DOI.
_OTHER_NAME = {"mla": "MLA 9", "chicago": "Chicago", "harvard": "Cite Them Right"}
_HEAD_MLA = re.compile(r"^(?P<head>[^“\"]{2,400}?)[“\"]")
_HEAD_CHI = re.compile(r"^(?P<head>.{2,400}?)\.\s+(?:(?:19|20)\d{2}[a-z]?|n\.\s?d\.|forthcoming)\.\s")
_HEAD_CTR = re.compile(r"^(?P<head>.{2,400}?)\(\s*(?:(?:19|20)\d{2}[a-z]?|no\s+date|n\.\s?d\.)")
_VOL_ISSUE = re.compile(r"(?<![\d(])(?P<v>\d{1,4})\s*\(\s*(?P<i>\d{1,4}(?:\s*[–\-/]\s*\d{1,4})?)\s*\)")
_RANGE = re.compile(r"(?P<pp>\bpp?\.\s*)?(?<![\d.])(?P<a>\d{1,5})\s*[–\-]\s*(?P<b>\d{1,5})(?![\d])")


# 6.6, round 82: a journal's name keeps the capitals of its principal words in every author-date style ("Journal of Sleep
# Research", "BMC Public Health"; only an article's or a book's title goes in sentence case where the style asks it); and
# APA 7 writes a volume and issue as 12(3), with no "Vol." or "No.".
_JOURNAL_AT = {"apa": re.compile(r"[.?!]\s+(?P<j>[A-Z][^.,;:()\[\]]{3,140}?),\s*\*?\d{1,4}\*?\s*(?:\(\s*[\w–\-]+\s*\))?\s*,"),
               "harvard": re.compile(r"[’'],\s+(?P<j>[A-Z][^.,;:()\[\]]{3,140}?),\s*\*?\d{1,4}"),
               "mla": re.compile(r"[”\"]\s+(?P<j>[A-Z][^.,;:()\[\]]{3,140}?),\s*vol\.", re.I),
               "chicago": re.compile(r"[”\"]\s+(?P<j>[A-Z][^.,;:()\[\]]{3,140}?)\s+\d{1,4}\s*(?:\(|,|:)")}
_STYLE_NAME = {"apa": "APA 7", "harvard": "Cite Them Right", "mla": "MLA 9", "chicago": "Chicago"}


def journal_note(entry, style):
    """"" or the line for a journal's name written with its principal words in lower case."""
    rx = _JOURNAL_AT.get(style)
    if not rx:
        return ""
    t = re.sub(r"^\s*(?:\[\d{1,3}\]|\d{1,3}[.)])\s*", "", str(entry or "")).replace("*", "")
    m = rx.search(t)
    if not m:
        return ""
    j = m.group("j").strip()
    words = j.split()
    low = [w for i, w in enumerate(words) if len(w) >= 4 and w.isalpha() and w.islower() and (i == 0 or w.lower() not in _TITLE_SMALL)]
    if not low:
        return ""
    fixed = " ".join(w if (i and w.lower() in _TITLE_SMALL) or not w.isalpha() or not w.islower() else w[:1].upper() + w[1:]
                     for i, w in enumerate(words))
    return f"{_STYLE_NAME[style]} writes a journal’s name with capitals for its principal words: {fixed}."


# 6.6, round 86: APA 7 and Cite Them Right give the authors' first names as initials ("Mensah, K."), MLA and Chicago in full.
_FULL_FIRST = re.compile(r"^(?P<sur>(?:(?:van|von|de|del|der|den|du|da|di|le|la)\s+)*[A-Z][\w'’\-]+),\s+(?P<first>[A-Z][a-z]{2,}(?:[\s-][A-Z][a-z]{2,})?)"
                         r"(?P<rest>\s+[A-Z]\.)?\s*(?:[.,(&]|and\b)")


def first_names_note(entry, style):
    """"" or the line for an APA 7 or Cite Them Right entry that gives its first author's first name in full."""
    if style not in ("apa", "harvard"):
        return ""
    t = re.sub(r"^\s*(?:\[\d{1,3}\]|\d{1,3}[.)])\s*", "", str(entry or "")).replace("*", "")
    m = _FULL_FIRST.match(t)
    if not m or m.group("first").lower() in ("and", "the", "for", "with", "ministry", "department", "university", "school", "office", "service"):
        return ""
    ini = " ".join(w[0] + "." for w in re.split(r"[\s-]+", m.group("first"))) + ((" " + m.group("rest").strip()) if m.group("rest") else "")
    return f"{_STYLE_NAME[style]} gives the authors’ first names as initials: {m.group('sur')}, {ini}"


def _other_notes(entry, style):
    t = re.sub(r"^\s*(?:\[\d{1,3}\]|\d{1,3}[.)])\s*", "", str(entry or "")).replace("*", "")
    name, out = _OTHER_NAME[style], []
    hm = {"mla": _HEAD_MLA, "chicago": _HEAD_CHI, "harvard": _HEAD_CTR}[style].match(t)
    head = hm.group("head") if hm else ""
    if re.search(r"\s&\s", head):
        out.append(f"{name} joins the authors’ names with “and”, not “&”.")
    if style == "mla":
        rest = t[len(head):] if head else t
        vi = _VOL_ISSUE.search(rest)
        if vi and not re.search(r"\bvol\.\s*\d", t, re.I):
            out.append(f"MLA 9 writes the volume and issue as “vol. {vi.group('v')}, no. {vi.group('i')}”.")
        if re.search(r"\bvol\.\s*\d", t, re.I):
            for m in _RANGE.finditer(rest):
                a, b = m.group("a"), m.group("b")
                if m.group("pp") or (len(a) == 4 and len(b) == 4 and a[:2] in ("19", "20") and b[:2] in ("19", "20")):
                    continue
                if t[:m.start()].rstrip().endswith(("no.", "vol.")):
                    continue
                out.append(f"MLA 9 writes a range of pages after “pp.”: pp. {a}–{b}.")
                break
    if style in ("mla", "chicago") and re.search(r"(?<![/\w])doi\s*:\s*10\.|https?://dx\.doi\.org/", t, re.I):     # 6.6, round 84
        out.append(f"{name} writes a DOI as a link: https://doi.org/10.…")
    if style == "harvard":
        fi = first_names_note(t, style)                                             # 6.6, round 86
        if fi:
            out.append(fi)
    if style == "harvard" and re.search(r"https?://|\bwww\.|\bdoi\.org/|\bdoi:\s*10\.", t, re.I) and not re.search(r"\bavailable\s+at\s*:", t, re.I):
        out.append("Cite Them Right writes “Available at:” before the address (a DOI as https://doi.org/…).")
    return out


def style_notes(entry, style, evidence=""):
    """What an entry of an APA 7 list writes differently from APA 7 itself (the publication manual's 7th edition, chapter
    9-10): "et al." in the list, a DOI not written as a link, "Retrieved from", "pp." for a journal article's pages, "and"
    before the last author, no date. [sentence, …]. 6.6, round 51: in every style, a title capitalised the other way
    (title_note)."""
    tn = title_note(entry, style, evidence)
    en = (ETAL_RULE + ".") if _ETAL_BAD.search(str(entry or "")) else ""          # 6.6, round 72: in every style
    jn = journal_note(entry, style)                                                # 6.6, round 82
    if style in ("mla", "chicago", "harvard"):
        return [x for x in [en] + _other_notes(entry, style) + [jn, tn] if x]
    if style != "apa":
        return [x for x in (en, tn) if x]
    t = re.sub(r"^\s*(?:\[\d{1,3}\]|\d{1,3}[.)])\s*", "", str(entry or ""))
    out = []
    hm = _APA_HEAD.match(t)
    head = hm.group("head") if hm else ""
    if re.search(r"\bet\.?\s*al\b", head, re.I):
        out.append("APA 7 names every author, up to 20, in the reference list: “et al.” belongs in the text.")
    elif en:
        out.append(en)
    if re.search(r"(?<!/)\bdoi\s*:\s*10\.|https?://dx\.doi\.org/|http://doi\.org/", t, re.I):
        out.append("APA 7 writes the DOI as a link: https://doi.org/10.…")
    if re.search(r"\bretrieved\s+from\b", t, re.I):
        out.append("APA 7 no longer writes “Retrieved from” before a link (only “Retrieved [date], from” for a page that changes).")
    if re.search(r"\d\s*(?:\(\s*\d[\w–\-]*\s*\))?\s*,\s*pp?\.\s*\d", t) and not re.search(r"\(\s*pp?\.\s*\d", t):
        out.append("APA 7 gives a journal article’s pages without “pp.”.")
    if head and "&" not in head and re.search(r"[A-Z]\.\s*,?\s+and\s+(?:(?:van|von|de|der|den|du|da|di|le|la)\s+)*[A-Z][\w'’\-]+,\s+[A-Z]\.", head):
        out.append("APA 7 puts “&” before the last author’s name.")
    if not re.search(r"\b(?:19|20)\d{2}[a-z]?\b|\bn\.\s?d\.|\(\s*in\s+press", t, re.I):
        out.append("APA 7 writes “(n.d.)” where a work has no date.")
    links = re.findall(r"https?://\S+", t)                                       # 6.6, round 87
    if any(re.search(r"doi\.org/10\.", u, re.I) for u in links) and any(not re.search(r"doi\.org/", u, re.I) for u in links):
        out.append("APA 7 gives a work’s DOI alone when it has one: the other link goes.")
    if re.search(r"\bavailable\s+(?:at|from)\s*:?\s*(?:https?://|www\.)|\(?\baccessed\s*:?\s+\d", t, re.I):
        out.append("APA 7 writes a link on its own, without “Available at” or “Accessed”.")
    if re.search(r"\bpersonal\s+communication", t, re.I):                       # 6.6, round 92
        out.append("APA 7 cites a personal communication in the text only: it has no entry in the reference list.")
    fi = first_names_note(t, "apa")                                               # 6.6, round 86
    if fi:
        out.append(fi)
    city = re.search(r"(?:^|[.)]\s)(?P<city>[A-Z][a-z]+(?:\s[A-Z][a-z]+)?(?:,\s*[A-Z]{2})?)\s*:\s*(?P<pub>[A-Z][\w&.'’-]*(?:\s+(?:[A-Z][\w&.'’-]*|of|and|the|for|&)){0,7})\.?\s*$", t)
    if city and not re.search(r"\d|\.\s", city.group("pub")) and city.group("city").lower() not in ("retrieved", "available", "accessed", "note", "in"):
        out.append(f"APA 7 gives the publisher without its city: {city.group('pub').strip().rstrip('.')}.")
    vn = re.search(r"\bvol(?:ume)?\.?\s*(?P<v>\d{1,4})(?:\s*,?\s*(?:no|issue)\.?\s*(?P<i>\d{1,4}))?", t, re.I)
    if vn:                                                                        # 6.6, round 82: "Vol. 12, No. 3"
        out.append(f"APA 7 writes the volume and issue as {vn.group('v')}" + (f"({vn.group('i')})" if vn.group("i") else "") + ", without “Vol.” or “No.”.")
    if jn:
        out.append(jn)
    if tn:
        out.append(tn)
    return out


# ---------------------------------------------------------------- 6.6, round 17: a book
_JOURNALISH = re.compile(r"\b\d{1,4}\s*\(\s*\d+[\w–\-]*\s*\)\s*[,:]|\bvol\.\s*\d|\bJournal\b|;\s*\d{1,4}\s*[(:]|\bArticle\s+e?\d", re.I)
_BOOKISH = re.compile(r"\b\d+(?:st|nd|rd|th)\s+(?:ed|edn)\b|\bed(?:ition|n)?\.|\(Eds?\.\)|\bPress\b|\bPublish(?:ers?|ing)\b|\bBooks\b|"
                      r"\b(?:Wiley|Elsevier|Sage|SAGE|Springer|Pearson|McGraw|Lippincott|Wolters\s+Kluwer|Routledge|Saunders|Mosby|"
                      r"Jones\s*(?:&|and)\s*Bartlett|Churchill\s+Livingstone|Oxford\s+University|Cambridge\s+University|Macmillan|Cengage)\b")


def looks_book(entry):
    """An entry that reads as a book: an edition, a publisher's name, editors; never a journal's volume and issue."""
    t = str(entry or "")
    return not _JOURNALISH.search(t) and bool(_BOOKISH.search(t))


def book_title(entry):
    """A book's title as the entry writes it (APA and Harvard: after the year; MLA: after the authors), without its edition."""
    t = re.sub(r"^\s*(?:\[\d{1,3}\]|\d{1,3}[.)])\s*", "", str(entry or ""))
    m = re.search(r"\((?:19|20)\d{2}[a-z]?\)\.?\s+(?P<t>.+)", t)
    if m:
        t = m.group("t")
    else:
        m = re.search(r"^(?P<a>.+?[a-z\u00e0-\u024f]{2,})\.\s+(?P<t>[A-Z\u201c\"].+)", t)    # MLA: "Surname, Given F., and Name. Title. …"
        if not m:
            return ""
        t = m.group("t")
    t = re.split(r"\s*\((?:\d+(?:st|nd|rd|th)\s+ed|Ed|Vol)|\.\s+(?:\d+(?:st|nd|rd|th)\s+edn?\b|[A-Z][a-z]+:\s)|\.\s*$|\.\s+[A-Z][\w&’'\- ]{2,40}(?:,|\.)", t, maxsplit=1)[0]
    return re.sub(r"[“”\"*_]", "", t).strip(" .,")


def book_match(entry, surname, docs):
    """The Open Library work that is the entry's book: its title's words in the entry, its author the entry's first author.
    {"title", "authors", "years", "first_year", "publisher", "url"} or None."""
    sn = norm_name(surname)
    for d in docs or []:
        if not isinstance(d, dict) or not d.get("title"):
            continue
        if not title_in(entry, d["title"]):
            continue
        names = [norm_name(str(a).split()[-1]) for a in (d.get("author_name") or []) if str(a).strip()]
        if sn and names and not any(sn == n or (len(sn) > 3 and n.endswith(sn)) for n in names):
            continue
        years = sorted({int(y) for y in (d.get("publish_year") or []) if str(y).isdigit() and 1450 < int(y) < 2100})
        return {"title": str(d["title"])[:300], "authors": [str(a)[:80] for a in (d.get("author_name") or [])[:6]], "years": years,
                "first_year": d.get("first_publish_year") if isinstance(d.get("first_publish_year"), int) else (years[0] if years else None),
                "publisher": str((d.get("publisher") or [""])[0])[:120], "url": "https://openlibrary.org" + str(d.get("key") or "")}
    return None


def book_years_note(year, years):
    """What to say when the entry's year is not one of the book's editions."""
    if not year or not years or year in years:
        return ""
    near = [y for y in years if abs(y - year) <= 3]
    shown = ", ".join(str(y) for y in (near or years[-6:]))
    return (f"Open Library has no edition of this book from {year} (it lists {shown}"
            + ("" if near or len(years) <= 6 else " and earlier") + "): check the year of the edition you read.")


# 6.6, round 140 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."; his rule of 15 Nov: "every statement should be
# cited"): a sentence of the person's writing that gives a figure about the world, or says what studies or the evidence show,
# with no source in it nor anywhere after it in its paragraph (a paragraph often cites once, at its end: that stands). Their
# own findings are left alone (a sentence about "this study", its respondents or participants, a table or figure, and every
# sentence of a results chapter or section, until its discussion). At most four, the first ones.
_UC_MARK = re.compile(r"\b(?:globally|worldwide|across\s+the\s+(?:world|globe)|statistics\s+(?:show|indicate|suggest)|"
                      r"(?:several|many|previous|recent|other|some|most)?\s*stud(?:y|ies)\s+(?:have\s+|has\s+)?(?:shown|show|found|reported|revealed|"
                      r"indicated|suggested|demonstrated)|research(?:ers)?\s+(?:has\s+|have\s+)?(?:shown|found|revealed|indicated|reported|suggested|"
                      r"demonstrated)|(?:the\s+)?evidence\s+(?:shows|suggests|indicates)|it\s+(?:is|has\s+been)\s+(?:estimated|reported|documented|"
                      r"established)\s+that|according\s+to)\b", re.I)
_UC_FIG = re.compile(r"\d+(?:\.\d+)?\s*(?:%|per\s*cent\b|percent\b)|\b\d[\d,.]*\s*(?:million|billion)\b|\b\d{1,3}(?:,\d{3})+\s+(?:people|deaths|cases|"
                     r"patients|children|women|men|students|adults|lives)\b", re.I)
_UC_OWN = re.compile(r"\b(?:this|the\s+present|the\s+current|our)\s+(?:study|research|work|survey|project|thesis|proposal)\b|\brespondents?\b|"
                     r"\bparticipants?\b|\binterviewees?\b|\b(?:table|figure|fig\.|chart)\s*\d|\bthe\s+researchers?\b|\bwe\s+(?:found|observed|recorded)\b|"
                     r"\bthe\s+(?:sample|data)\s+(?:was|were|showed)\b|\b(?:level\s+of\s+significance|significance\s+level|confidence\s+(?:level|interval)|"
                     r"margin\s+of\s+error|response\s+rate|cronbach|spss|stata|sample\s+size|was\s+set\s+at|were\s+(?:analy[sz]ed|computed|"
                     r"calculated|coded|entered|sampled|selected|recruited))\b", re.I)          # (its own methods: no source asked)
_UC_ANYCITE = re.compile(r"\[\s*\d{1,3}(?:\s*[,–-]\s*\d{1,3})*\s*\]|[¹²³⁴⁵⁶⁷⁸⁹⁰]+(?:[˒,⁻–][¹²³⁴⁵⁶⁷⁸⁹⁰]+)*|\((?=[^()]*\b(?:1[89]|20)\d{2}[a-z]?\b)[^()]{2,160}\)|"
                         r"\b[A-Z][A-Za-z'’-]+(?:\s+(?:et\s+al\.?|and\s+[A-Z][A-Za-z'’-]+|&\s+[A-Z][A-Za-z'’-]+))?\s*\((?:19|20)\d{2}[a-z]?\)")
_UC_ABBR = re.compile(r"(?:\bet\s+al|\be\.g|\bi\.e|\bcf|\bvs|\bDr|\bMr|\bMrs|\bMs|\bProf|\bFig|\bNo|\bVol|\bpp|\bSt|\bapprox|\bca)\.$")
_UC_RESULTS = re.compile(r"^\s*(?:chapter\s+four\b|4\.0\b|(?:\d+\.?\d*\s+)?(?:results|findings|presentation\s+of\s+(?:results|findings|data)|"
                         r"data\s+(?:analysis\s+and\s+)?presentation|analysis\s+and\s+(?:results|presentation))\b)", re.I)
_UC_AFTER = re.compile(r"^\s*(?:chapter\s+five\b|5\.0\b|(?:\d+\.?\d*\s+)?(?:discussion|conclusions?|summary|recommendations)\b)", re.I)
_UC_SPLIT = re.compile(r"(?<=[.!?])[\"”’)]?\s+(?=[A-Z“\"(])")


def uncited_claims(body, cites):
    """[{"cite": the sentence's beginning, "note"}]: claims about the world or the research with no source in their paragraph
    from them on."""
    t = str(body or "")
    spans = [(c["start"], c.get("end", c["start"] + 1)) for c in (cites or []) if isinstance(c.get("start"), int)]
    spans += [(m.start(), m.end()) for m in _UC_ANYCITE.finditer(t)]      # (any citation's shape, in any style)
    out, pos, in_results = [], 0, False
    for para in t.split("\n"):
        p0 = pos
        pos += len(para) + 1
        line = para.strip()
        if not line:
            continue
        if len(line) <= 90 and _UC_RESULTS.match(line):
            in_results = True
            continue
        if len(line) <= 90 and _UC_AFTER.match(line):
            in_results = False
            continue
        if in_results or "\t" in para:
            continue
        end = p0 + len(para)
        cuts, last = [], 0
        for m in _UC_SPLIT.finditer(para):
            if _UC_ABBR.search(para[last:m.start()].rstrip("\"”’)")):
                continue                          # ("et al. (2020)", "e.g. Ghana": not a sentence's end)
            cuts.append((last, m.start()))
            last = m.end()
        cuts.append((last, len(para)))
        for i, j in cuts:
            sent = para[i:j]
            s0 = p0 + i
            if len(sent) < 25 or sent.rstrip().endswith("?") or _UC_OWN.search(sent):
                continue
            fig, mark = _UC_FIG.search(sent), _UC_MARK.search(sent)
            if not (fig or mark):
                continue
            if any(b > s0 and a < end for a, b in spans):
                continue                          # (a source in it, or later in its paragraph)
            head = sent.strip()
            shown = "“" + (head[:110].rstrip() + " …" if len(head) > 110 else head) + "”"
            note = ("It gives a figure but cites no source: add the study or report it comes from." if fig else
                    "It says what research or the evidence shows but cites none: add the studies.")
            out.append({"cite": shown, "note": note})
            if len(out) >= 4:
                return out
    return out
