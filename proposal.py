"""Academic-Home 5.4: a research proposal, written the way a university asks for one.

The owner, 12 Nov 2026, with a nursing school's proposal as the model: "Apply the same approach to research as in the
attached document when writing a proposal. Make everything beautifully and high quality like or than it. If there's any
information the user didn't provide please pause and ask. and any question you ask the user should really count in the
research. Ask about their Research topic (can suggest some for them immediately if they don't have; about five to chose
from), main objectives, specific objectives (how many objectives, they may state them or ai choose for them using
their information provided."

The approach of that document, which this module follows:
    a cover; a contents list; a structured abstract;
    CHAPTER ONE: 1.0 Background of the Study (from the world to the continent, the region, the country and the place,
        every statement cited), 1.1 Problem Statement, 1.2 Research Objectives (general, specific, research questions,
        one question for each objective), 1.3 Significance, 1.4 Scope, 1.5 Limitations, 1.6 Organisation of the Study;
    CHAPTER TWO: LITERATURE REVIEW (2.0 Introduction, themes that follow the objectives, a summary);
    CHAPTER THREE: METHODOLOGY (3.0 Introduction, 3.1 Study Design, 3.2 Study Setting, 3.3 Study Population, 3.4 Sample
        Size (the formula worked out), 3.5 Sampling Technique, 3.6 Inclusion and Exclusion Criteria, 3.7 Research
        Instrument, 3.8 Data Collection Procedure, 3.9 Data Analysis Method, 3.10 Ethical Considerations);
    References; Appendices (the instrument, informed consent, the budget, the work plan as a Gantt chart).

Two kinds of words, kept apart (rule 9: everything presented as sourced must be academically valid):
    - what is said ABOUT THE WORLD (the background, the problem, the literature) comes only from papers that were found,
      every sentence cited and checked against its paper (server/app.py: the same checks as every piece of writing);
    - THE PERSON'S OWN PLAN (objectives, questions, significance, scope, limits, method, ethics, the appendices) comes
      only from what the person gave or agreed to. It carries no citation and no figure that isn't theirs or worked out
      from theirs (guard_plan_text): nothing is invented about their school, their people or their numbers.

This file holds what can be worked out without the AI or the network: the plan made safe, what is still missing, the
parts a proposal is written in, the sample size, the work plan, the consent form, the budget, and the words given to
the AI for each part. The requests themselves are in app.py; the Word file is converters.proposal_to_docx.
"""
import datetime
import math
import re

# ---------------------------------------------------------------- what can be chosen (the app shows the same lists)
DESIGNS = {
    "quant_cross": ("Quantitative, descriptive cross-sectional survey",
                    "a quantitative, descriptive, cross-sectional design"),
    "quant_corr": ("Quantitative, correlational", "a quantitative, correlational design"),
    "quant_exp": ("Quantitative, experimental or quasi-experimental", "a quantitative, quasi-experimental design"),
    "qual_phenom": ("Qualitative, in-depth interviews", "a qualitative, descriptive phenomenological design"),
    "qual_case": ("Qualitative, case study", "a qualitative case study design"),
    "mixed": ("Mixed methods (numbers and interviews)", "a convergent mixed-methods design"),
}
QUALITATIVE = ("qual_phenom", "qual_case")
SAMPLING = {
    "stratified": ("Stratified random sampling", "a stratified random sampling technique"),
    "simple_random": ("Simple random sampling", "a simple random sampling technique"),
    "systematic": ("Systematic sampling", "a systematic sampling technique"),
    "cluster": ("Cluster sampling", "a cluster sampling technique"),
    "convenience": ("Convenience sampling", "a convenience sampling technique"),
    "purposive": ("Purposive sampling", "a purposive sampling technique"),
    "snowball": ("Snowball sampling", "a snowball sampling technique"),
    "census": ("Everyone in the population (a census)", "a census of the whole study population"),
}
INSTRUMENTS = {
    "questionnaire": ("A questionnaire", "a structured questionnaire"),
    "interview": ("An interview guide", "a semi-structured interview guide"),
    "focus_group": ("A focus group guide", "a focus group discussion guide"),
    "observation": ("An observation checklist", "a structured observation checklist"),
    "records": ("A records review form", "a data extraction form for existing records"),
    "questionnaire_interview": ("A questionnaire and an interview guide", "a structured questionnaire and a semi-structured interview guide"),
}
ANALYSES = {
    "statistics": ("Statistics (frequencies, tests of difference and relationship)", "descriptive and inferential statistics"),
    "thematic": ("Thematic analysis", "thematic analysis"),
    "both": ("Statistics and thematic analysis", "descriptive and inferential statistics together with thematic analysis"),
}
# how full the proposal is: about how many words, how the literature review is shared out, and its parts
DEPTHS = {
    "short": {"name": "Short", "words": 3500, "bg": 4, "problem": 1, "themes": 4, "theme_paras": 1, "ch2_parts": 1},
    "standard": {"name": "Standard", "words": 6000, "bg": 5, "problem": 2, "themes": 7, "theme_paras": 2, "ch2_parts": 2},
    "full": {"name": "Full", "words": 10000, "bg": 7, "problem": 2, "themes": 11, "theme_paras": 2, "ch2_parts": 3},
}
MAX_OBJECTIVES, MIN_OBJECTIVES = 6, 2
MARGIN = 0.05                      # the margin of error used for a sample size
_WS = re.compile(r"\s+")


def _line(v, cap):
    return _WS.sub(" ", str(v if v is not None else "")).strip()[:cap]


def _lines(v, n, cap):
    out = []
    for x in (v if isinstance(v, list) else str(v or "").split("\n")):
        t = _line(x, cap)
        t = re.sub(r"^(?:\d{1,2}[.)]|[-•*])\s+", "", t)             # a typed "1." or "-" in front
        if t and t.lower() not in [o.lower() for o in out]:
            out.append(t)
        if len(out) >= n:
            break
    return out


def _sentence(t):
    """A statement as a sentence: a capital first, one full stop last (a question keeps its mark)."""
    t = _line(t, 400).strip(" .;")
    if not t:
        return ""
    t = t[0].upper() + t[1:]
    return t if t.endswith("?") else t + "."


def _int(v, lo, hi):
    try:
        n = int(float(str(v).replace(",", "").strip()))
    except (TypeError, ValueError):
        return 0
    return n if lo <= n <= hi else 0


def _month(v):
    """"2026-09" → (2026, 9), or None."""
    m = re.match(r"^\s*((?:19|20)\d{2})-(\d{1,2})\s*$", str(v or ""))
    if not m or not 1 <= int(m.group(2)) <= 12:
        return None
    return int(m.group(1)), int(m.group(2))


def clean_plan(raw):
    """The plan the app sent, made safe: known keys, short lines, sane numbers. Nothing is added to it here."""
    raw = raw if isinstance(raw, dict) else {}
    p = {
        "title": _line(raw.get("title"), 300).strip(" ."),
        "general": _sentence(raw.get("general")),
        "specific": [_sentence(x) for x in _lines(raw.get("specific"), MAX_OBJECTIVES, 300)],
        "questions": [_line(x, 300) for x in _lines(raw.get("questions"), MAX_OBJECTIVES, 300)],
        "gap": _sentence(raw.get("gap")),
        "design": raw.get("design") if raw.get("design") in DESIGNS else "",
        "design_other": _line(raw.get("design_other"), 120),
        "setting": _line(raw.get("setting"), 200).strip(" ."),
        "country": _line(raw.get("country"), 60).strip(" ."),
        "region": _line(raw.get("region"), 60).strip(" ."),
        "continent": _line(raw.get("continent"), 40).strip(" ."),
        "population": _line(raw.get("population"), 200).strip(" ."),
        "pop_size": _int(raw.get("pop_size"), 2, 100000000),
        "sample_given": _int(raw.get("sample_given"), 1, 1000000),
        "sampling": raw.get("sampling") if raw.get("sampling") in SAMPLING else "",
        "instrument": raw.get("instrument") if raw.get("instrument") in INSTRUMENTS else "",
        "instrument_from": _line(raw.get("instrument_from"), 200).strip(" ."),
        "analysis": raw.get("analysis") if raw.get("analysis") in ANALYSES else "",
        "software": _line(raw.get("software"), 60).strip(" ."),
        "ethics_body": _line(raw.get("ethics_body"), 160).strip(" ."),
        "keywords": _lines(raw.get("keywords"), 8, 60),
        "depth": raw.get("depth") if raw.get("depth") in DEPTHS else "standard",
        # the cover
        "institution": _lines(raw.get("institution"), 4, 120),
        "authors": [],
        "supervisor": _line(raw.get("supervisor"), 120),
        "programme": _line(raw.get("programme"), 120),
        "date": _line(raw.get("date"), 40),
        # the work plan and the budget
        "start": raw.get("start") if _month(raw.get("start")) else "",
        "months": _int(raw.get("months"), 2, 36),
        "currency": _line(raw.get("currency"), 12),
        "budget": [],
        "budget_mode": raw.get("budget_mode") if raw.get("budget_mode") in ("blank", "mine", "none") else "blank",
        "contact": _line(raw.get("contact"), 160),
        # how many researchers there are (the writing says "the researcher" or "the researchers"; their names are for the
        # cover and the consent form, which are made when the Word file is, and never go to the AI)
        "researchers": _int(raw.get("researchers"), 1, 12),
        # who it is written for (their level, their field), and anything the person said it must follow
        "about": _line(raw.get("about"), 200),
        "notes": _line(raw.get("notes"), 600),
    }
    for a in (raw.get("authors") if isinstance(raw.get("authors"), list) else [])[:12]:
        if isinstance(a, dict):
            name, ident = _line(a.get("name"), 80), _line(a.get("id"), 40)
        else:
            m = re.match(r"^(.*?)[\s,;:–—-]+([A-Za-z]{0,6}[/\-]?\d[\w/\-]{2,30})\s*$", _line(a, 140))   # a name, then its index number: "Ama Mensah SN/0412/22" (a made-up one)
            name, ident = (m.group(1).strip(" ,;"), m.group(2)) if m else (_line(a, 80), "")
        if name:
            p["authors"].append({"name": name, "id": ident})
    for b in (raw.get("budget") if isinstance(raw.get("budget"), list) else [])[:20]:
        if not isinstance(b, dict) or not _line(b.get("item"), 80):
            continue
        try:
            total = round(float(str(b.get("total")).replace(",", "")), 2) if str(b.get("total") or "").strip() else None
        except (TypeError, ValueError):
            total = None
        if total is not None and not 0 <= total <= 1e9:
            total = None
        p["budget"].append({"item": _line(b.get("item"), 80), "qty": _line(b.get("qty"), 12), "unit": _line(b.get("unit"), 16), "total": total})
    if len(p["questions"]) != len(p["specific"]):
        p["questions"] = [question_for(o) for o in p["specific"]]
    return p


LOGO_MAX_BYTES = 600 * 1024          # as sent (the app sends far less: it makes the picture small first)
LOGO_MAX_SIDE = 900


def clean_logo(raw):
    """The institution's logo for the cover, sent with the Word file's request as a data URL (PNG or JPEG). It is read
    as a picture and written again from its pixels alone (so that nothing else a file can carry comes along), on white,
    no larger than a cover needs. Returns PNG bytes; None when it is not such a picture, is too large or can't be read.
    The logo is for the Word file only: it is never kept and never among the words given to the AI."""
    if not isinstance(raw, str) or len(raw) > LOGO_MAX_BYTES * 4 // 3 + 200:
        return None
    m = re.match(r"^data:image/(png|jpeg);base64,([A-Za-z0-9+/=\s]+)$", raw)
    if not m:
        return None
    try:
        import base64
        import io
        from PIL import Image
        data = base64.b64decode(m.group(2), validate=False)
        if not 67 <= len(data) <= LOGO_MAX_BYTES:
            return None
        if not (data[:8] == b"\x89PNG\r\n\x1a\n" or data[:3] == b"\xff\xd8\xff"):
            return None
        im = Image.open(io.BytesIO(data))
        if im.format not in ("PNG", "JPEG"):
            return None
        w, h = im.size
        if w < 16 or h < 16 or w > 2400 or h > 2400:          # (looked at before the picture itself is read)
            return None
        im.load()
        im = im.convert("RGBA")
        if max(w, h) > LOGO_MAX_SIDE:
            k = LOGO_MAX_SIDE / float(max(w, h))
            im = im.resize((max(1, int(round(w * k))), max(1, int(round(h * k)))), Image.LANCZOS)
        page = Image.new("RGB", im.size, (255, 255, 255))      # a see-through picture is drawn on black by some programs
        page.paste(im, mask=im.split()[3])
        out = io.BytesIO()
        page.save(out, format="PNG", optimize=True)
        return out.getvalue()
    except Exception:
        return None


def how_many(plan):
    """How many researchers: the names on the cover when they were given, else the number the app sent, else one."""
    return len(plan.get("authors") or []) or plan.get("researchers") or 1


def missing(plan):
    """What must be known before a proposal can be written well (the app pauses and asks for each): [(key, why)]."""
    need = []
    if not plan.get("title"):
        need.append(("title", "The proposal needs its topic."))
    if not plan.get("general"):
        need.append(("general", "The main objective says what the whole study is for."))
    if len(plan.get("specific") or []) < MIN_OBJECTIVES:
        need.append(("specific", "A proposal needs at least two specific objectives."))
    if not plan.get("design") and not plan.get("design_other"):
        need.append(("design", "The study design decides the whole of Chapter Three."))
    if not plan.get("setting"):
        need.append(("setting", "Chapter Three describes where the study will take place."))
    if not plan.get("population"):
        need.append(("population", "Chapter Three describes who will be studied."))
    return need


# ---------------------------------------------------------------- objectives and questions
_Q_VERBS = r"(?:identify|determine|describe|assess|examine|explore|establish|investigate|document|find out|evaluate|measure|ascertain|analyse|analyze|understand|test|estimate)"
_Q_PREP = re.compile(r"\s(?:among|amongst|of|in|at|on|toward|towards|for|between|by|with|to|that|which|who|within|across|during|about|regarding|concerning|from)\s", re.I)
_Q_ONE = {"diabetes", "series", "species", "status", "analysis", "basis", "emphasis", "news", "measles", "aids", "access", "progress", "success", "stress", "bias",
          "consensus", "census", "process", "illness", "business", "focus", "virus", "campus", "syllabus", "crisis", "thesis", "diagnosis", "hypothesis"}


def _q_plural(phrase):
    """Is the thing an objective names more than one ("the attitudes of…", "the factors that…")? Its head noun is the
    last word before the first preposition."""
    words = [w.lower().strip(",;:") for w in _Q_PREP.split(" " + phrase + " ", 1)[0].strip().split(" ")]
    while len(words) > 1 and words[-1].endswith(("ed", "ing")):       # "the coping strategies used by…": the noun is before "used"
        words.pop()
    head = words[-1] if words else ""
    if not head.endswith("s") or head in _Q_ONE or head.endswith(("ness", "ics", "ss", "us", "is")):
        return False
    return True


def question_for(objective):
    """A research question for one specific objective, when the person gave none ("To determine the level of X" →
    "What is the level of X?"). A plain rewording: nothing is added to what the objective says. (RamzAI's own
    suggestions come with their questions; this is for objectives a person typed without any.)"""
    o = _line(objective, 300).strip(" .")
    if not o:
        return ""
    low = o[0].lower() + o[1:]
    low = low if low.startswith("to ") else "to " + low
    m = re.match(r"^to " + _Q_VERBS + r" (?:whether|if) (.+)$", low, re.I)
    if m:
        return "Is it the case that %s?" % m.group(1)
    m = re.match(r"^to " + _Q_VERBS + r" (?:the )?(relationship|association|correlation|link) between (.+)$", low, re.I)
    if m:
        return "Is there a significant %s between %s?" % (m.group(1), m.group(2))
    m = re.match(r"^to (?:" + _Q_VERBS[3:-1] + r"|compare) (?:the )?differences? in (.+)$", low, re.I)
    if m:
        return "Is there a significant difference in %s?" % m.group(1)
    m = re.match(r"^to compare (.+)$", low, re.I)
    if m:
        return "How do %s compare?" % m.group(1)
    m = re.match(r"^to " + _Q_VERBS + r" (how|why|when|where) (.+)$", low, re.I)
    if m:
        return "What does the evidence show about %s %s?" % (m.group(1).lower(), m.group(2))
    m = re.match(r"^to " + _Q_VERBS + r" (?:the )?(.+)$", low, re.I)
    if m:
        return "What %s the %s?" % ("are" if _q_plural(m.group(1)) else "is", m.group(1))
    return "What can be said about %s?" % re.sub(r"^to \w+ ", "", low)


# ---------------------------------------------------------------- the parts a proposal is written in
ORDINALS = ("ONE", "TWO", "THREE", "FOUR", "FIVE")


def themes_wanted(plan):
    return DEPTHS[plan.get("depth") or "standard"]["themes"]


def parts(plan):
    """The requests a proposal is written in, in order: [{"key", "label"}]. One chapter to a part; the literature
    review takes one, two or three parts as it grows; the last part is the method, the abstract and the appendices."""
    d = DEPTHS[plan.get("depth") or "standard"]
    out = [{"key": "ch1", "label": "Chapter One: Introduction"}]
    n = d["ch2_parts"]
    share = {1: [""], 2: [", first half", ", second half"], 3: [", first part", ", middle part", ", last part"]}[n]
    for i in range(n):
        out.append({"key": "ch2-%d" % (i + 1), "label": "Chapter Two: Literature review" + share[i]})
    out.append({"key": "ch3", "label": "Chapter Three: Methodology, the abstract and the appendices"})
    return out


def share_themes(outline, plan, part_key):
    """The themes of the literature review that one part writes (the outline is shared out evenly, in order; the last
    part also writes the summary)."""
    n = DEPTHS[plan.get("depth") or "standard"]["ch2_parts"]
    try:
        k = int(part_key.split("-")[1]) - 1
    except (IndexError, ValueError):
        return []
    outline = list(outline or [])
    per = int(math.ceil(len(outline) / float(n))) if outline else 0
    return outline[k * per:(k + 1) * per]


def clean_outline(raw, plan):
    """The literature review's themes as the AI proposed them, made safe: [{"n": "2.1", "h": heading, "q": search}]."""
    out, seen = [], set()
    for x in (raw if isinstance(raw, list) else [])[:16]:
        if not isinstance(x, dict):
            continue
        h = _line(x.get("h") or x.get("heading"), 140).strip(" .:")
        h = re.sub(r"^\d+(?:\.\d+)*\s+", "", h)
        q = _line(x.get("q") or x.get("query") or h, 160)
        if not h or h.lower() in seen or re.match(r"^(?:introduction|summary|conclusion|summary and conclusion)$", h, re.I):
            continue
        seen.add(h.lower())
        out.append({"h": h, "q": q})
    out = out[:themes_wanted(plan)]
    for i, t in enumerate(out):
        t["n"] = "2.%d" % (i + 1)
    return out


def fallback_outline(plan):
    """Themes straight from the objectives, when the AI's outline can't be used: one theme for each objective."""
    core = plan.get("title") or ""
    raw = [{"h": "The Concept of " + _topic_noun(plan), "q": core}]
    for o in plan.get("specific") or []:
        h = re.sub(r"^to \w+ (?:the )?", "", _line(o, 200).strip(" ."), flags=re.I)
        raw.append({"h": h[:1].upper() + h[1:], "q": h})
    return clean_outline(raw, plan)


def _topic_noun(plan):
    kw = plan.get("keywords") or []
    t = kw[0] if kw else re.split(r"\b(?:among|in|of|at|on|and)\b", plan.get("title") or "the Topic", 1)[0]
    return _line(t, 60).strip(" ,.:").title() or "the Topic"


# ---------------------------------------------------------------- the sample size, worked out
def _fmt(n):
    return "{:,}".format(int(n))


def sample_size(plan):
    """How many people the study needs, worked out from what the person gave. Returns a dict:
      {"method": "yamane", "N", "e", "n", "steps": [...], "text": [paragraphs]}   a known population (Yamane)
      {"method": "cochran", "z", "p", "e", "n", ...}                              a population of unknown size
      {"method": "census", "N", "n"}                                              everyone is studied
      {"method": "saturation", "n" (the person's number, or 0)}                    interviews: until nothing new comes
      {"method": "given", "n"}                                                    the person's own number
    A sample size is always rounded UP to the next whole person: rounding down would give less precision than asked."""
    design, N, e = plan.get("design") or "", plan.get("pop_size") or 0, MARGIN
    who = plan.get("population") or "participants"
    if plan.get("sample_given") and not N:
        n = plan["sample_given"]
        return {"method": "given", "n": n, "text": [
            "The study will recruit %s %s. This number was set by the researcher%s for the time and resources available to the study."
            % (_fmt(n), who, "s" if how_many(plan) > 1 else "")]}
    if design in QUALITATIVE:
        n = plan.get("sample_given") or 0
        lead = ("About %s %s will be invited to take part. " % (_fmt(n), who)) if n else ""
        return {"method": "saturation", "n": n, "text": [
            lead + "Because the study is qualitative, the final number of participants will not be fixed in advance. Recruitment and "
            "interviewing will continue until data saturation is reached, that is, until further interviews add no new themes to the analysis."]}
    if plan.get("sampling") == "census" and N:
        return {"method": "census", "N": N, "n": N, "text": [
            "All %s %s in the study population will be invited to take part, so no sample will be drawn. A census is appropriate "
            "because the population is small enough to be reached in full." % (_fmt(N), who)]}
    if N:
        exact = N / (1.0 + N * e * e)
        n = int(math.ceil(exact - 1e-9))
        ne2 = N * e * e
        whole = abs(exact - round(exact)) < 1e-9
        steps = ["n = %s / 1 + %s(%s)²" % (_fmt(N), _fmt(N), _num(e)), "n = %s / 1 + %s(%s)" % (_fmt(N), _fmt(N), _num(e * e)),
                 "n = %s / 1 + %s" % (_fmt(N), _num(ne2)), "n = %s / %s" % (_fmt(N), _num(1 + ne2))]
        steps += ["n = %s" % _fmt(n)] if whole else ["n = %s" % _num(round(exact, 2)), "n ≈ %s" % _fmt(n)]
        return {"method": "yamane", "N": N, "e": e, "n": n, "exact": round(exact, 2), "formula": "n = N / 1 + N(e)²", "steps": steps, "text": [
            "The sample size for the study will be determined using Yamane’s formula, which is appropriate for calculating a sample from a finite population. The formula is expressed as:",
            "Where n is the sample size, N is the total population and e is the margin of error.",
            "For this study, the total population (N) is estimated to be %s %s. A margin of error (e) of %s will be used to ensure adequate precision." % (_fmt(N), who, _num(e)),
            ("" if whole else "The result is rounded up to the next whole number, so that the precision asked for is fully met. ")
            + "A minimum sample of %s %s will therefore be required for the study." % (_fmt(n), who)]}
    z, p = 1.96, 0.5
    exact = z * z * p * (1 - p) / (e * e)
    n = int(math.ceil(exact - 1e-9))
    steps = ["n = (%s)² × %s × (1 − %s) / (%s)²" % (_num(z), _num(p), _num(p), _num(e)), "n = %s × %s / %s" % (_num(round(z * z, 4)), _num(p * (1 - p)), _num(e * e)),
             "n = %s" % _num(round(exact, 2)), "n ≈ %s" % _fmt(n)]            # (a line "A / B" is the fraction A over B: everything after the stroke is under it)
    return {"method": "cochran", "z": z, "p": p, "e": e, "n": n, "exact": round(exact, 2), "formula": "n = Z² p (1 − p) / e²", "steps": steps, "text": [
        "The exact size of the study population is not known, so the sample size will be determined using Cochran’s formula for a large or unknown population. The formula is expressed as:",
        "Where n is the sample size, Z is the standard normal value for the chosen level of confidence, p is the estimated proportion of the population with the characteristic of interest, and e is the margin of error.",
        "For this study, a confidence level of 95%% (Z = %s), an estimated proportion (p) of %s, which gives the largest sample, and a margin of error (e) of %s will be used." % (_num(z), _num(p), _num(e)),
        "The result is rounded up to the next whole number. A minimum sample of %s %s will therefore be required for the study." % (_fmt(n), who)]}


def sample_text(sample):
    """The Sample Size section as plain words (the chat shows this; the Word file sets the formula out as a fraction)."""
    t = sample.get("text") or []
    if not sample.get("steps"):
        return "\n\n".join(t)
    flat = lambda line: re.sub(r"^(.+?[=≈]\s*)(.+?)\s+/\s+(.+)$", lambda m: "%s%s / (%s)" % (m.group(1), m.group(2), m.group(3)) if re.search(r"[+\-×]", m.group(3)) else m.group(0), line)
    return "\n\n".join([t[0], flat(sample["formula"]), t[1], t[2], "\n".join(flat(x) for x in sample["steps"]), t[3]])


def _num(x):
    """A number the way a worked example shows it: 0.05, 0.0025, 6, 333.33."""
    if abs(x - round(x)) < 1e-9:
        return _fmt(round(x))
    s = ("%.4f" % x).rstrip("0").rstrip(".")
    return s


# ---------------------------------------------------------------- the work plan (a Gantt chart)
ACTIVITIES = (("Proposal writing", 0.00, 0.14), ("Proposal defence", 0.14, 0.20), ("Ethical review and permissions", 0.20, 0.34),
              ("Pre-testing the instrument", 0.30, 0.40), ("Data collection", 0.40, 0.62), ("Data analysis", 0.60, 0.76),
              ("Writing the results and discussion", 0.72, 0.90), ("Submission and final defence", 0.90, 1.00))
_MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


def gantt(plan):
    """The work plan: {"months": [labels], "rows": [(activity, first month index, last month index)], "dated": bool}.
    With a start month the columns are real months; without one they are "Month 1", "Month 2"… (no date is invented).
    The usual order of a study's work is shared out over the months the person gave (nine when they gave none)."""
    n = plan.get("months") or 9
    start = _month(plan.get("start"))
    if start:
        y, m = start
        labels = []
        for i in range(n):
            mm, yy = (m - 1 + i) % 12, y + (m - 1 + i) // 12
            labels.append("%s %d" % (_MONTHS[mm], yy))
    else:
        labels = ["Month %d" % (i + 1) for i in range(n)]
    rows = []
    for name, a, b in ACTIVITIES:
        first = min(n - 1, int(math.floor(a * n + 1e-9)))
        last = min(n - 1, max(first, int(math.ceil(b * n - 1e-9)) - 1))
        rows.append((name, first, last))
    return {"months": labels, "rows": rows, "dated": bool(start)}


# ---------------------------------------------------------------- the budget
BUDGET_ITEMS = (("Stationery and printing", ("Printing of questionnaires and consent forms", "Photocopying", "Bond paper and stationery")),
                ("Communication", ("Internet data and airtime",)),
                ("Field work", ("Transport during data collection",)),
                ("Analysis and production", ("Statistical or analysis services", "Plagiarism check", "Printing and binding of the report")))


def budget(plan):
    """The budget's rows: {"currency", "rows": [("head", name) | ("item", name, qty, unit, total or None)], "total"}.
    The person's own costs when they typed them; else the usual items with empty amounts for them to fill in (no cost
    is invented). None when they asked for no budget."""
    if plan.get("budget_mode") == "none":
        return None
    cur = plan.get("currency") or ""
    mine = [b for b in (plan.get("budget") or []) if b.get("item")]
    if plan.get("budget_mode") == "mine" and mine:
        rows = [("item", b["item"], b.get("qty") or "", b.get("unit") or "", b.get("total")) for b in mine]
        known = [b["total"] for b in mine if b.get("total") is not None]
        return {"currency": cur, "rows": rows, "total": round(sum(known), 2) if known else None, "filled": True}
    rows = []
    for head, items in BUDGET_ITEMS:
        rows.append(("head", head))
        rows.extend(("item", it, "", "", None) for it in items)
    return {"currency": cur, "rows": rows, "total": None, "filled": False}


# ---------------------------------------------------------------- informed consent
def consent(plan):
    """The informed consent form: [(heading or "", paragraph)]. Standard wording, filled with what the person gave; a
    line they didn't give (a contact) is left as a line to write on, never made up."""
    who = "we" if how_many(plan) > 1 else "I"
    Who = who.capitalize()
    are = "are" if who == "we" else "am"
    names = [a["name"] for a in plan.get("authors") or [] if a.get("name")]
    inst = (plan.get("institution") or [""])[0]
    lead = ("%s %s %s" % (Who, are, _join(names))) if names else ("%s %s the researcher%s for this study" % (Who, are, "s" if who == "we" else ""))
    if inst:
        lead += ", of " + inst
    instrument = INSTRUMENTS.get(plan.get("instrument") or "questionnaire", INSTRUMENTS["questionnaire"])[1]
    minutes = "30 to 45 minutes" if plan.get("design") in QUALITATIVE else "10 to 15 minutes"
    contact = plan.get("contact") or "______________________________"
    sup = (" You may also contact the study’s supervisor, %s." % plan["supervisor"]) if plan.get("supervisor") else ""
    task = ("take part in an interview" if plan.get("instrument") in ("interview", "focus_group") else
            "complete a questionnaire and, if you agree, take part in a short interview" if plan.get("instrument") == "questionnaire_interview" else
            "complete %s" % instrument)
    return [
        ("", "%s. %s %s inviting you to take part in a research study titled “%s”." % (lead, Who, are, plan.get("title") or "")),
        ("Purpose of the study", (plan.get("general") or "").replace("To ", "The study aims to ", 1) if plan.get("general") else "The purpose of the study will be explained to you by the researcher."),
        ("Voluntary participation", "Taking part is entirely voluntary. You may refuse to take part, decline to answer any question, or withdraw from the study at any time without giving a reason and without any penalty."),
        ("What you will be asked to do", "You will be asked to %s. This will take about %s of your time." % (task, minutes)),
        ("Confidentiality", "Your answers will be kept confidential. No name, index number or other detail that could identify you will appear in any report of the study. "
                            "The information will be stored securely, will be seen only by the research team, and will be used for academic purposes only."),
        ("Risks and benefits", "There are no known risks in taking part beyond those of everyday life. You will receive no payment or direct benefit, "
                               "but your answers will help to improve understanding of the subject of the study."),
        ("Questions", "If you have any question about the study, please contact the researcher%s at %s.%s" % ("s" if who == "we" else "", contact, sup)),
        ("Consent", "I have read and understood the information above. My questions have been answered, and I agree to take part in this study."),
    ]


def _join(names):
    names = [n for n in names if n]
    return names[0] if len(names) == 1 else ", ".join(names[:-1]) + " and " + names[-1]


# ---------------------------------------------------------------- the topic without its place
_LEAD_IN = re.compile(r"\s+(?:in|at|within|from|across|of)\s+(?:the\s+)?", re.I)


def wide_topic(plan):
    """The study's subject without where it is done ("… Among Student Nurses in Northern Ghana" gives "… Among Student
    Nurses"): what the background's widest search looks for, so that it can open with the world before it narrows to the
    region and the country. The title itself when no place can be told apart in it."""
    title = re.sub(r"\s+", " ", str(plan.get("title") or "")).strip()
    low = title.lower()
    places = [plan.get("country"), plan.get("region"), plan.get("continent")] + re.split(r"\s*,\s*", plan.get("setting") or "")
    cut = len(title)
    for place in places:
        place = re.sub(r"\s+", " ", str(place or "")).strip().lower()
        m = re.search(r"(?<![a-z])" + re.escape(place) + r"(?![a-z])", low) if len(place) >= 4 else None
        if m:
            cut = min(cut, m.start())
    if cut == len(title):
        return title
    head = title[:cut]
    lead = [m for m in _LEAD_IN.finditer(head)]
    if lead and len(head) - lead[-1].start() <= 40:
        head = head[:lead[-1].start()]              # "… in Northern ", "… at the ", "… in the Tamale Metropolis of "
    head = head.strip(" ,:;-–—(")
    return head if len(head.split()) >= 3 else title


# ---------------------------------------------------------------- what the plan allows the AI to say
_STD_NUMBERS = {0.05, 5.0, 95.0, 1.96, 0.5, 0.7, 0.70, 1.0, 2.0, 3.0, 4.0, 5.0, 10.0, 15.0, 30.0, 45.0, 100.0}


def allowed_numbers(plan, sample=None, extra_text=""):
    """The figures a plan's words may hold: the person's own, those worked out from them, and the method's own constants
    (the 0.05 level of significance, the 95% level of confidence, a scale's points, 0.70 for Cronbach's alpha)."""
    ok = set(_STD_NUMBERS)
    for k in ("pop_size", "sample_given", "months"):
        if plan.get(k):
            ok.add(float(plan[k]))
    ok.add(float(len(plan.get("specific") or [])))
    if sample:
        for k in ("n", "N", "exact"):
            if sample.get(k):
                ok.add(float(sample[k]))
    words = " ".join([plan.get(k) or "" for k in ("title", "general", "gap", "setting", "population", "software", "instrument_from", "ethics_body", "design_other",
                                                    "notes", "about")]
                     + (plan.get("specific") or []) + [extra_text or ""])
    for m in re.finditer(r"\d+(?:[.,]\d+)*", words):
        try:
            ok.add(float(m.group(0).replace(",", "")))
        except ValueError:
            pass
    return ok


_CITE = re.compile(r"\s*\((?:[^()]{0,160}?\b(?:1[5-9]|20)\d{2}[a-z]?|[^()]{0,120}?\bn\.d\.)\)")
_NARR = re.compile(r"\b[A-Z][A-Za-z’'\-]+(?:\s+(?:and|&)\s+[A-Z][A-Za-z’'\-]+|\s+et\s+al\.?)?,?\s+\((?:1[5-9]|20)\d{2}[a-z]?\)")
_NUM = re.compile(r"(?<![\w.])(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)(?![\w])")
_NUM_LABEL = re.compile(r"(?:(?:chapter|section|appendix|table|figure|objective|question|part|step|phase|sdg|goal|item|sections?)\s*|covid[\s\-–]*)$", re.I)
# a number that only names something ("Chapter 3", "COVID-19"). A level ("Level 300") or a version ("version 26") is a
# fact about the place or the software: it stays only when it is in the person's own words (allowed_numbers).
_SENT_END = re.compile(r"(?<=[.!?])\s+(?=[A-Z“\"(])")


def guard_plan_text(text, allowed):
    """A plan's words as the AI wrote them, with anything that isn't the person's taken out: a sentence that cites a
    paper (the plan cites none: citations belong to the cited chapters, where each is checked), and a sentence that
    states a figure which is neither the person's nor the method's. Headings and list lines are kept as they are.
    Returns (text, how many sentences went)."""
    out, gone = [], 0
    for line in str(text or "").split("\n"):
        st = line.strip()
        if not st or st.startswith("#") or st.startswith("⟦"):
            out.append(line)
            continue
        lead = re.match(r"^\s*(?:[-•*]|\d{1,2}[.)])\s+", line)
        body = line[lead.end():] if lead else line
        kept = []
        for s in _SENT_END.split(body):
            if not s.strip():
                continue
            if _CITE.search(s) or _NARR.search(s):
                gone += 1
                continue
            bad = False
            for m in _NUM.finditer(s):
                raw = m.group(1)
                try:
                    v = float(raw.replace(",", ""))
                except ValueError:
                    continue
                if len(raw) == 4 and raw.isdigit() and 1500 <= v <= 2099:
                    continue                                   # a year
                if _NUM_LABEL.search(s[max(0, m.start() - 14):m.start()]):
                    continue                                   # "Chapter 3", "Section 2"
                if re.match(r"\s*(?:=|[-–]\s*\d)", s[m.end():m.end() + 4]) and v <= 5:
                    continue                                   # "1 = Strongly disagree", "1–5"
                if not any(abs(v - a) < 1e-6 for a in allowed):
                    bad = True
                    break
            if bad:
                gone += 1
                continue
            kept.append(s.strip())
        if kept:
            out.append((lead.group(0) if lead else "") + " ".join(kept))
    return re.sub(r"\n{3,}", "\n\n", "\n".join(out)).strip(), gone


# ---------------------------------------------------------------- the words given to the AI
def plan_block(plan, sample=None):
    """The person's plan, as lines the AI writes from (and nothing else)."""
    d = plan.get("design")
    rows = [("TITLE", plan.get("title")), ("GENERAL OBJECTIVE", plan.get("general")),
            ("SPECIFIC OBJECTIVES", "\n".join("  %d. %s" % (i + 1, o) for i, o in enumerate(plan.get("specific") or []))),
            ("RESEARCH QUESTIONS", "\n".join("  %d. %s" % (i + 1, o) for i, o in enumerate(plan.get("questions") or []))),
            ("THE GAP THE STUDY FILLS (the researcher's own statement)", plan.get("gap")),
            ("STUDY DESIGN", DESIGNS[d][1] if d in DESIGNS else plan.get("design_other")),
            ("STUDY SETTING", plan.get("setting")), ("COUNTRY", plan.get("country")),
            ("STUDY POPULATION", plan.get("population")),
            ("SIZE OF THE POPULATION", (_fmt(plan["pop_size"]) + " (an estimate)") if plan.get("pop_size") else "not known"),
            ("SAMPLE SIZE", ("%s (%s)" % (_fmt(sample["n"]), {"yamane": "Yamane's formula, margin of error 0.05", "cochran": "Cochran's formula, 95% confidence, margin of error 0.05",
                                                             "census": "everyone in the population", "given": "set by the researcher",
                                                             "saturation": "about this many; recruitment continues until data saturation"}[sample["method"]]))
             if sample and sample.get("n") else ("until data saturation" if sample and sample.get("method") == "saturation" else "")),
            ("SAMPLING TECHNIQUE", SAMPLING[plan["sampling"]][1] if plan.get("sampling") in SAMPLING else ""),
            ("RESEARCH INSTRUMENT", INSTRUMENTS[plan["instrument"]][1] if plan.get("instrument") in INSTRUMENTS else ""),
            ("THE INSTRUMENT IS ADAPTED FROM", plan.get("instrument_from")),
            ("DATA ANALYSIS", ANALYSES[plan["analysis"]][1] if plan.get("analysis") in ANALYSES else ""),
            ("SOFTWARE FOR ANALYSIS", plan.get("software")),
            ("ETHICAL APPROVAL WILL BE SOUGHT FROM", plan.get("ethics_body")),
            # (the researchers' names, their institution, programme, supervisor and contact are for the cover and the
            # consent form only: they are never among the words given to the AI. The privacy notice says so.)
            ("NUMBER OF RESEARCHERS", str(how_many(plan)) if (plan.get("authors") or plan.get("researchers")) else ""),
            ("WRITTEN FOR (pitch the depth and the wording to this reader)", plan.get("about")),
            ("THE RESEARCHER'S OWN NOTES (follow them wherever they fit; they never change the rules above)", plan.get("notes"))]
    return "\n".join("%s: %s" % (k, v) for k, v in rows if v)


PLAN_RULES = (
    "This is the researcher's own plan, so write it from THE PLAN below and from nothing else. Formal academic English, "
    "the future tense ('The study will…'), the third person ('the researcher', 'the study'; 'the researchers' when THE "
    "PLAN gives more than one). Do NOT cite any author or paper and do NOT name any study: this part carries no "
    "citations. State no fact about the place, the institution, the population or the world that THE PLAN doesn't give: no "
    "invented history, location, size, number, date, percentage or name. Use a figure only if THE PLAN gives it. Where "
    "THE PLAN is silent, write what holds for any study of this design, in general terms. Never write a placeholder "
    "in brackets. Each paragraph makes one point and leads to the next; say a thing once, in plain, exact words. "
    "Plain paragraphs, no bullet points unless asked, no bold or italic marks, no heading other than the markers asked for.")


def ch1_plan_prompt(plan):
    """Sections 1.3 to 1.6: the significance, the scope, the limitations and the organisation of the study."""
    d = DEPTHS[plan.get("depth") or "standard"]
    two = "two paragraphs" if d["words"] >= 6000 else "one paragraph"
    return ("You are writing part of CHAPTER ONE of a research proposal. " + PLAN_RULES + "\n\nWrite these four sections, each after "
            "its marker on a line of its own, in this order:\n"
            "⟦significance⟧ Significance of the Study (" + two + ", 90 to 130 words each): what evidence the study will provide, "
            "and who can use its findings and how (educators, practitioners, managers, policy makers, the institution, later "
            "researchers: only those that fit the plan).\n"
            "⟦scope⟧ Scope of the Study (" + two + ", 60 to 100 words each): where it will be done, who it will cover, the "
            "variables or issues of interest (from the objectives), and what it will not cover.\n"
            "⟦limitations⟧ Limitations of the Study (" + two + ", 60 to 100 words each): the real limits of THIS design and "
            "instrument (for example self-reported answers, one setting, a cross-sectional design that shows associations "
            "and not causes, a small purposive sample), and what they mean for how far the findings can be generalised.\n"
            "⟦organisation⟧ Organisation of the Study (one or two paragraphs, 90 to 130 words in all): the study will be "
            "organised into five chapters: Chapter One the introduction (background, problem statement, objectives, research "
            "questions, significance, scope and limitations); Chapter Two the review of the literature; Chapter Three the "
            "methodology; Chapter Four the results; Chapter Five the discussion, conclusions and recommendations.\n\n"
            "THE PLAN\n" + plan_block(plan))


CH3_HEADS = (("intro", "3.0", "Introduction"), ("design", "3.1", "Study Design"), ("setting", "3.2", "Study Setting"),
             ("population", "3.3", "Study Population"), ("sample", "3.4", "Sample Size"), ("sampling", "3.5", "Sampling Technique"),
             ("criteria", "3.6", "Inclusion and Exclusion Criteria"), ("instrument", "3.7", "Research Instrument"),
             ("procedure", "3.8", "Data Collection Procedure"), ("analysis", "3.9", "Data Analysis Method"),
             ("ethics", "3.10", "Ethical Considerations"))


def ch3_prompt(plan, sample):
    qual = plan.get("design") in QUALITATIVE
    return ("You are writing CHAPTER THREE (METHODOLOGY) of a research proposal. " + PLAN_RULES + "\n\nWrite these sections, each after "
            "its marker on a line of its own, in this order (the Sample Size section is written separately: do not write it):\n"
            "⟦intro⟧ Introduction (one paragraph, 70 to 100 words): what the chapter presents, in the order of its sections.\n"
            "⟦design⟧ Study Design (one paragraph, 90 to 130 words): the design in THE PLAN and why it suits these objectives.\n"
            "⟦setting⟧ Study Setting (one paragraph, 60 to 110 words): where the study will be done, only as THE PLAN names it; "
            "say why it is a suitable place for this study. Invent nothing about the place.\n"
            "⟦population⟧ Study Population (one paragraph, 70 to 110 words): who will be studied, as THE PLAN gives them, "
            "and why they are the right people for these objectives.\n"
            "⟦sampling⟧ Sampling Technique (one or two paragraphs, 100 to 150 words): the technique in THE PLAN, how it will be "
            "carried out step by step for this population, and why it fits. Use THE PLAN's SAMPLE SIZE exactly as it is: never "
            "recalculate it, and add no allowance (for non-response or anything else) that THE PLAN doesn't give.\n"
            "⟦criteria⟧ Inclusion and Exclusion Criteria: the line 'Inclusion criteria' then three or four bullet points "
            "(each starting '- '), then the line 'Exclusion criteria' then three or four bullet points; each follows from "
            "THE PLAN's population and setting (who belongs to it, is present during data collection, gives informed consent).\n"
            "⟦instrument⟧ Research Instrument (two paragraphs, 150 to 200 words): the instrument in THE PLAN; its sections "
            "(the participants' background, then one section for each specific objective) and how each will be answered"
            + ("" if qual else " (for example a four- or five-point scale of agreement or frequency)") + "; "
            + ("how the guide will be reviewed and tried out before use. " if qual else
               "how it will be pre-tested with a small group like the participants, and that its internal consistency will be "
               "checked with Cronbach's alpha, a value of 0.70 or more being taken as acceptable. ")
            + "If THE PLAN says what the instrument is adapted from, say so in those words; otherwise say it will be developed "
            "by the researcher from the study's objectives and the reviewed literature.\n"
            "⟦procedure⟧ Data Collection Procedure (two paragraphs, 130 to 180 words): approvals and permission first, then "
            "explaining the study and obtaining informed consent, then how and where the data will be collected and by whom, "
            "and how the completed tools will be checked and kept.\n"
            "⟦analysis⟧ Data Analysis Method (two paragraphs, 120 to 170 words): how the data will be prepared and analysed, "
            "naming for EACH specific objective the analysis that answers it"
            + (" (themes developed from the transcripts)" if qual else " (frequencies, percentages and means for a description; a test of "
               "relationship or of difference where an objective asks for one, at the 0.05 level of significance)")
            + ". Name the software only if THE PLAN names it.\n"
            "⟦ethics⟧ Ethical Considerations (two paragraphs, 120 to 170 words): ethical approval (from the body THE PLAN names, "
            "or 'the relevant ethics review committee'), permission from the setting, voluntary participation, informed consent, "
            "the right to withdraw, confidentiality and anonymity, and safe keeping of the data.\n\n"
            "THE PLAN\n" + plan_block(plan, sample))


def outline_prompt(plan):
    n = themes_wanted(plan)
    return ("You plan the themes of CHAPTER TWO (LITERATURE REVIEW) of a research proposal. From THE PLAN, give exactly %d "
            "themes, in the order they should be read: first the main concept(s) of the study (what it is, how common it is, "
            "what is known about its causes or effects), then one or more themes for EACH specific objective, in the "
            "objectives' order. Each theme: a heading in Title Case (at most 14 words, no number) and a search query (4 to 9 "
            "plain words, as typed into a database of research papers, with no quotation marks). Themes must not overlap. Do "
            "not include an introduction or a summary: those are added separately. Reply with JSON only: "
            "{\"themes\": [{\"h\": \"heading\", \"q\": \"search query\"}]}.\n\nTHE PLAN\n" % n) + plan_block(plan)


def abstract_prompt(plan, sample):
    return ("You are writing the ABSTRACT of a research proposal from what has already been written and agreed. Write four "
            "labelled parts, each after its marker on a line of its own:\n"
            "⟦introduction⟧ two or three sentences on the problem, drawn ONLY from CHAPTER ONE below (reword it; add nothing; "
            "no citations and no figures);\n"
            "⟦aim⟧ one or two sentences giving the aim and what the objectives cover, from THE PLAN;\n"
            "⟦methodology⟧ three or four sentences on the design, the setting, the population, the sample and the "
            "instrument, from THE PLAN, in the future tense;\n"
            "⟦analysis⟧ two sentences on how the data will be analysed, from THE PLAN, in the future tense;\n"
            "⟦keywords⟧ four to six keywords, separated by commas, taken from the title and objectives.\n"
            "150 to 250 words in all. Formal academic English, the third person. No citations, no heading, no bold or italic "
            "marks, nothing in brackets, and no fact that the two sources below don't give.\n\nTHE PLAN\n" + plan_block(plan, sample))


def instrument_prompt(plan):
    qual = plan.get("design") in QUALITATIVE or plan.get("instrument") in ("interview", "focus_group")
    if qual:
        return ("You draft the INTERVIEW GUIDE for a research proposal, from THE PLAN. Reply with JSON only:\n"
                "{\"kind\": \"guide\", \"intro\": \"two sentences that thank the participant and say what the interview is about\", "
                "\"sections\": [{\"title\": \"Section A: Background\", \"questions\": [\"...\"]}, {\"title\": \"Section B: <objective 1 in a few words>\", "
                "\"questions\": [\"...\"]}]}\n"
                "Section A: three or four easy background questions that fit this population. Then ONE section for EACH specific "
                "objective, in order, each with three to five open questions (with a short follow-up probe in brackets where it "
                "helps). Plain words, one thing in each question, nothing leading, nothing that asks for a name or identifies a "
                "person.\n\nTHE PLAN\n" + plan_block(plan))
    return ("You draft the QUESTIONNAIRE for a research proposal, from THE PLAN. Reply with JSON only:\n"
            "{\"kind\": \"questionnaire\", \"intro\": \"two sentences that say who is asking, what the study is about, and that answers are confidential\", "
            "\"background\": [{\"q\": \"Age (years)\", \"options\": [\"18-24\", \"25-29\", \"30-34\", \"35 and above\"]}], "
            "\"sections\": [{\"title\": \"Section B: <objective 1 in a few words>\", \"scale\": \"agreement\" or \"frequency\", \"items\": [\"a statement\"]}]}\n"
            "background: four or five questions that fit THIS population (for example age, sex, programme, level or years of "
            "work), each with three to five answer options that cover everyone and don't overlap. Then ONE section for each "
            "specific objective that a respondent can answer (an objective about a relationship or a difference is answered "
            "from the other sections: give it no section), each with six to eight short statements a respondent can rate: "
            "scale \"agreement\" (Strongly disagree to Strongly agree) for views, beliefs and attitudes; \"frequency\" (Never to "
            "Always) for what people do. Plain words, one idea in each statement, nothing leading, nothing that asks for a "
            "name or identifies a person. Do not number the statements.\n\nTHE PLAN\n" + plan_block(plan))


def clean_instrument(raw, plan):
    """The instrument as the AI drafted it, made safe; None when it can't be used."""
    if not isinstance(raw, dict):
        return None
    intro = _line(raw.get("intro"), 500)
    if raw.get("kind") == "guide" or (raw.get("sections") and isinstance(raw["sections"][0], dict) and raw["sections"][0].get("questions")):
        secs = []
        for s in (raw.get("sections") if isinstance(raw.get("sections"), list) else [])[:8]:
            if not isinstance(s, dict):
                continue
            qs = [_line(q, 260) for q in (s.get("questions") if isinstance(s.get("questions"), list) else [])[:7] if _line(q, 260)]
            if qs:
                secs.append({"title": _line(s.get("title"), 120) or "Section %s" % chr(65 + len(secs)), "questions": qs})
        return {"kind": "guide", "intro": intro, "sections": secs} if len(secs) >= 2 else None
    bg = []
    for b in (raw.get("background") if isinstance(raw.get("background"), list) else [])[:6]:
        if isinstance(b, dict) and _line(b.get("q"), 120):
            opts = [_line(o, 50) for o in (b.get("options") if isinstance(b.get("options"), list) else [])[:6] if _line(o, 50)]
            bg.append({"q": _line(b.get("q"), 120).rstrip(":"), "options": opts})
    secs = []
    for s in (raw.get("sections") if isinstance(raw.get("sections"), list) else [])[:7]:
        if not isinstance(s, dict):
            continue
        items = [re.sub(r"^\d{1,2}[.)]\s*", "", _line(i, 220)) for i in (s.get("items") if isinstance(s.get("items"), list) else [])[:9] if _line(i, 220)]
        if len(items) >= 3:
            secs.append({"title": _line(s.get("title"), 120) or "Section %s" % chr(66 + len(secs)),
                         "scale": "frequency" if s.get("scale") == "frequency" else "agreement", "items": items})
    return {"kind": "questionnaire", "intro": intro, "background": bg, "sections": secs} if secs else None


SCALES = {"agreement": (("SD", "Strongly disagree"), ("D", "Disagree"), ("A", "Agree"), ("SA", "Strongly agree")),
          "frequency": (("N", "Never"), ("R", "Rarely"), ("S", "Sometimes"), ("O", "Often"), ("A", "Always"))}


# ---------------------------------------------------------------- the fixed words of Chapter One's objectives
def objectives_sections(plan):
    """Section 1.2, in the person's own words: the general objective, the specific objectives, the research questions."""
    spec = "\n".join("%d. %s" % (i + 1, o) for i, o in enumerate(plan.get("specific") or []))
    ques = "\n".join("%d. %s" % (i + 1, q) for i, q in enumerate(plan.get("questions") or []))
    return [{"number": "1.2", "heading": "Research Objectives", "level": 2, "text": "", "kind": "plan"},
            {"number": "1.2.1", "heading": "General Objective", "level": 3, "text": plan.get("general") or "", "kind": "plan"},
            {"number": "1.2.2", "heading": "Specific Objectives", "level": 3, "text": "The study will seek:\n\n" + spec, "kind": "plan"},
            {"number": "1.2.3", "heading": "Research Questions", "level": 3, "text": "The study will answer the following questions:\n\n" + ques, "kind": "plan"}]


def split_marked(text, keys):
    """{key: text} from what the AI wrote after each ⟦key⟧ marker."""
    out = {}
    parts_ = re.split(r"⟦\s*([a-z0-9_]+)\s*⟧", str(text or ""))
    for i in range(1, len(parts_) - 1, 2):
        k = parts_[i].strip().lower()
        if k in keys and k not in out:
            out[k] = re.sub(r"\n{3,}", "\n\n", parts_[i + 1]).strip()
    return out


def today_words():
    t = datetime.date.today()
    return "%s %d" % (("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")[t.month - 1], t.year)
