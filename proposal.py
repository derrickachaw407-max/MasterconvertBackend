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
        # one question for each objective: the person's own, in their order, as far as they go; then one in the words of
        # each objective left (6.6, round 33: before, a person's questions were all set aside when there were fewer)
        given = p["questions"][:len(p["specific"])]
        p["questions"] = given + [question_for(o) for o in p["specific"][len(given):]]
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
    plural = lambda w: w.endswith("s") and w not in _Q_ONE and not w.endswith(("ness", "ics", "ss", "us", "is")) and "'" not in w and "’" not in w
    # 6.6, round 105: "the coping strategies nurses use…", "the challenges nurses face…": the noun comes first
    first = [w.lower().strip(",;:") for w in phrase.split()[:2]]
    if first and first[0] in ("the", "a", "an"):
        first = [w.lower().strip(",;:") for w in phrase.split()[1:3]]
    if any(plural(w) for w in first) and not any(w in ("of", "for", "in", "on", "to", "towards", "among", "about") for w in first[:1]):
        return True
    return plural(head)


# 6.6, round 83 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): an objective RamzAI suggests starts with "To" and
# one measurable verb. "To know", "to understand", "to find out", "to look at" can't be measured or answered by a design: a
# supervisor's first correction. The person's own objectives are never touched (only RamzAI's suggestions pass here).
MEASURABLE = ("determine", "assess", "examine", "identify", "describe", "compare", "explore", "establish", "evaluate", "investigate",
              "measure", "ascertain", "estimate", "analyse", "analyze", "document")
_VAGUE = (("get to know", "determine"), ("find out", "determine"), ("look into", "examine"), ("look at", "examine"),
          ("learn about", "describe"), ("know about", "determine"), ("know", "determine"), ("see", "determine"), ("study", "examine"),
          ("check", "assess"), ("discover", "identify"), ("understand", None), ("appreciate", "describe"), ("be aware of", "assess"))


def measurable(objective, qualitative=False):
    """The objective with a measurable verb ("To know the level of…" -> "To determine the level of…"; "to understand" becomes
    "to explore" in a qualitative study, "to assess" otherwise); "Assess the…" gets its "To". Anything else stays as it is."""
    t = str(objective or "").strip()
    m = re.match(r"^(to\s+)?(.*)$", t, re.I | re.S)
    has_to, rest = bool(m.group(1)), m.group(2)
    low = rest.lower()
    if any(low.startswith(v + " ") for v in MEASURABLE):
        return t if has_to else "To " + rest[:1].lower() + rest[1:]
    if not has_to:
        return t
    for vague, good in _VAGUE:
        if low.startswith(vague + " "):
            good = good or ("explore" if qualitative else "assess")
            return "To " + good + rest[len(vague):]
    return t


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



# 6.6, round 105 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): each research question asks what its own
# objective sets out to find. RamzAI's suggested questions were taken whenever they came one for each objective, so a question
# about another objective (or none) could stand under it, and a supervisor checks that line first. A question that shares
# too little of what tells its objective apart from the others is written again from the objective (question_for); the
# person's own questions stay as they wrote them.
def question_fits(objective, question, objectives, plan=None):
    """True when the question asks about what tells its objective apart from the other objectives (half of those words, at
    least one), the study's shared topic, people and place aside."""
    objs = [str(o) for o in (objectives or []) if str(o).strip()]
    shared = set.intersection(*[_obj_words(o) for o in objs]) if len(objs) > 1 else set()
    plan = plan or {}
    shared |= _obj_words(plan.get("population")) | _obj_words(plan.get("setting"))
    own = _obj_words(objective) - shared
    if not own:
        return True
    hit = own & _obj_words(question)
    return bool(hit) and len(hit) * 2 >= min(len(own), 4)


def fit_questions(objectives, questions, plan=None):
    """The questions, one for each objective, each asking what its objective sets out to find: one that doesn't is written
    again from its objective. (questions as RamzAI suggested them; the person's own are never passed here.)"""
    objs = [str(o) for o in (objectives or [])]
    out = []
    for i, o in enumerate(objs):
        q = str(questions[i]) if i < len(questions or []) else ""
        out.append(q if q and question_fits(o, q, objs, plan) else question_for(o))
    return out

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
# 6.6, round 64 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a study whose participants are children (under 18)
# asks a parent or guardian for consent and the child for assent. The consent form spoke only to the participant ("I agree
# to take part"), whoever they were, and chapter three's ethics section was not held to it: a proposal for pupils went out
# without parental consent, which every ethics committee asks for first.
_MINOR = re.compile(r"\b(?:pupils?|school\s*children|children|child|kids|adolescents?|teen(?:ager)?s|minors?|juveniles?|infants?|toddlers?|"
                    r"under[\s-]*(?:fives?|5s?|18s?|eighteens?)|kindergarten\w*|"
                    r"(?:junior|senior)\s+high\s+(?:school\s+)?(?:students?|pupils?|learners?)|"
                    r"(?:secondary|high|primary|basic|elementary|middle|junior|senior)\s+school\s+(?:students?|pupils?|learners?)|"
                    r"(?:jhs|shs)\s*\d?\s+(?:students?|pupils?|learners?)|aged\s+(?:[5-9]|1[0-7])\s*(?:-|\u2013|to)\s*(?:[5-9]|1[0-9])\b)", re.I)
_ADULT_OF = re.compile(r"\b(?:mothers?|fathers?|parents?|caregivers?|care-givers?|guardians?|teachers?|nurses?|midwives|health\s+workers?|"
                       r"staff|head\s*teachers?|principals?|carers?|families|households?)\b\s+(?:of|for|with|caring\s+for|who\s+(?:have|care\s+for)|attending)\b",
                       re.I)


def minors(plan):
    """Do the participants include children (under 18)? From who will be studied (or, when that isn't given, the title):
    pupils, children, adolescents, secondary-school students; not the adults around them ("mothers of children under
    five", "teachers of pupils")."""
    pop = str(plan.get("population") or plan.get("title") or "")
    return bool(_MINOR.search(pop)) and not _ADULT_OF.search(pop)


def ethics_gaps(plan, ethics):
    """What chapter three's ethics section leaves out for children: ["parent", "assent"] (empty when nothing)."""
    if not minors(plan):
        return []
    t = str(ethics or "")
    out = []
    if not re.search(r"\b(?:parents?|guardians?)\b", t, re.I):
        out.append("parent")
    if not re.search(r"\bassent\b", t, re.I):
        out.append("assent")
    return out


ETHICS_MINORS_ASK = ("Write ONLY the ⟦ethics⟧ section again, after its marker. The participants include children (under 18): say "
                     "that written informed consent will be obtained from each child's parent or guardian, that each child's assent "
                     "will be sought in words they understand, that a child who does not want to take part will not be included even "
                     "when a parent has agreed, and that permission will be obtained from the setting; keep the rest of the section.")


def consent(plan):
    """The informed consent form: [(heading or "", paragraph)]. Standard wording, filled with what the person gave; a
    line they didn't give (a contact) is left as a line to write on, never made up. For children (6.6, round 64): the
    form speaks to a parent or guardian, and the child gives assent."""
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
    if minors(plan):
        return [
            ("", "Dear parent or guardian. %s. %s %s inviting your child to take part in a research study titled “%s”." % (lead, Who, are, plan.get("title") or "")),
            ("Purpose of the study", (plan.get("general") or "").replace("To ", "The study aims to ", 1) if plan.get("general") else "The purpose of the study will be explained to you by the researcher."),
            ("Voluntary participation", "Your child’s taking part is entirely voluntary. You or your child may refuse, your child may decline to answer any question, and you or your "
                                        "child may withdraw from the study at any time without giving a reason and without any penalty."),
            ("What your child will be asked to do", "Your child will be asked to %s. This will take about %s, at a time that does not disturb lessons or care." % (task, minutes)),
            ("Confidentiality", "Your child’s answers will be kept confidential. No name, index number or other detail that could identify your child will appear in any "
                                "report of the study. The information will be stored securely, will be seen only by the research team, and will be used for academic purposes only."),
            ("Risks and benefits", "There are no known risks in taking part beyond those of everyday life. Your child will receive no payment or direct benefit, "
                                   "but the answers will help to improve understanding of the subject of the study."),
            ("Your child’s agreement", "The study will also be explained to your child in words they understand, and your child will be asked to agree. A child who does not "
                                       "want to take part will not be included, even when a parent or guardian has agreed. A participant aged 18 or over signs for themselves."),
            ("Questions", "If you have any question about the study, please contact the researcher%s at %s.%s" % ("s" if who == "we" else "", contact, sup)),
            ("Consent of the parent or guardian", "I am the parent or guardian of the child named below. I have read and understood the information above, my questions "
                                                  "have been answered, and I agree that my child may take part in this study."),
            ("Assent of the child", "The study has been explained to me. I know that I can stop at any time, and I agree to take part."),
        ]
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
def fit_methods(plan):
    """6.6, round 15: a design has its own kinds of sampling, instrument and analysis, whatever was sent (the app keeps the
    same rule, src/app-54.js _p54Fitting; so does /api/brief/plan): a qualitative study has no random sample, no
    questionnaire and no statistics; a mixed-methods study has both kinds of instrument and analysis; a quantitative study
    has statistics. The plan is changed in place and returned."""
    d = plan.get("design")
    if d in QUALITATIVE:
        if plan.get("sampling") not in ("purposive", "snowball", "convenience"):
            plan["sampling"] = "purposive"
        if plan.get("instrument") not in ("interview", "focus_group", "observation"):
            plan["instrument"] = "interview"
        plan["analysis"] = "thematic"
    elif d == "mixed":
        plan["sampling"] = plan.get("sampling") or "stratified"
        if not plan.get("instrument") or plan.get("instrument") in ("interview", "focus_group"):
            plan["instrument"] = "questionnaire_interview"
        plan["analysis"] = "both"
    elif d:
        if not plan.get("sampling") or plan.get("sampling") == "snowball":
            plan["sampling"] = "simple_random"
        if not plan.get("instrument") or plan.get("instrument") in ("interview", "focus_group", "questionnaire_interview"):
            plan["instrument"] = "questionnaire"
        plan["analysis"] = "statistics"
    return plan


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
            "THE PLAN's population and setting (who belongs to it, is present during data collection, gives informed consent). "
            "An exclusion criterion is a reason to leave out someone who meets every inclusion criterion (for example too ill to "
            "take part, away during data collection, took part in the pre-test), never an inclusion criterion turned round "
            "('students who do not consent' only repeats 'students who consent').\n"
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
            fair = [q for q in qs if not _leading(q)]          # 6.6, round 65: no question that leads to its answer
            if len(fair) >= 2:
                qs = fair
            title = _line(s.get("title"), 120)
            if not re.search(r"\bbackground\b|\bdemograph|\bpersonal\s+(?:information|details|data)\b|\bbio-?data\b", title or "", re.I):
                qs = [x for q in qs for x in split_question(q)][:8]       # 6.6, round 99: one thing in each question
                qs = [q if _leading(q) else open_question(q) for q in qs]  # 6.6, round 98: open, not yes or no (one kept as drafted stays)
            if qs:
                secs.append({"title": _line(s.get("title"), 120) or "Section %s" % chr(65 + len(secs)), "questions": qs})
        return {"kind": "guide", "intro": intro, "sections": secs} if len(secs) >= 2 else None
    bg = []
    for b in (raw.get("background") if isinstance(raw.get("background"), list) else [])[:6]:
        if isinstance(b, dict) and _line(b.get("q"), 120):
            opts = [_line(o, 50) for o in (b.get("options") if isinstance(b.get("options"), list) else [])[:6] if _line(o, 50)]
            bg.append({"q": _line(b.get("q"), 120).rstrip(":"), "options": opts})
    bg = [dict(b, options=fix_ranges(b["options"])) for b in bg]          # 6.6, round 9: answers that don't overlap or leave a gap
    secs, seen = [], set()
    for s in (raw.get("sections") if isinstance(raw.get("sections"), list) else [])[:7]:
        if not isinstance(s, dict):
            continue
        items = [re.sub(r"^\d{1,2}[.)]\s*", "", _line(i, 220)) for i in (s.get("items") if isinstance(s.get("items"), list) else [])[:9] if _line(i, 220)]
        # 6.6, round 9: a statement once in the whole questionnaire; a question ("Do you…?") can't be agreed with or rated
        # for how often, so it goes when the section keeps three statements without it
        items = [i for k, i in enumerate(items) if _item_key(i) not in {_item_key(x) for x in items[:k]} and _item_key(i) not in seen]
        rated = [i for i in items if not i.rstrip().endswith("?")]
        if len(rated) >= 3:
            items = rated
        fair = [i for i in items if not _leading(i)]            # 6.6, round 65: no statement that leads to its answer
        if len(fair) >= 3:
            items = fair
        one = [x for i in items for x in split_statement(i)]     # 6.6, round 99: one thing in each statement
        if len([x for x in one if x]) >= 3:
            items = [x for k, x in enumerate(one) if _item_key(x) not in {_item_key(y) for y in one[:k]} and _item_key(x) not in seen][:9]
        if s.get("scale") == "frequency":
            items = _fit_frequency(items)
        seen |= {_item_key(i) for i in items}
        if len(items) >= 3:
            secs.append({"title": _line(s.get("title"), 120) or "Section %s" % chr(66 + len(secs)),
                         "scale": "frequency" if s.get("scale") == "frequency" else "agreement", "items": items})
    return {"kind": "questionnaire", "intro": intro, "background": bg, "sections": secs} if secs else None


# 6.6, round 70 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): the instrument answers every specific objective it
# should (its content validity): a questionnaire one section for each objective a respondent can answer (an objective about a
# relationship or a difference is answered from the others, the participants' background from its own questions), an
# interview guide one for each objective. The draft was used as it came: a section left out meant an objective no data would
# answer, and a draft cut off at its room meant no instrument at all.
_OBJ_BG = re.compile(r"\b(?:socio[\s-]*demographic|demographic|background\s+(?:characteristics|information|of)|personal\s+characteristics|"
                     r"characteristics\s+of\s+(?:the\s+)?(?:respondents|participants))", re.I)
_OBJ_STOP = set("assess determine identify describe examine explore establish investigate document find evaluate measure ascertain analyse "
                "analyze understand test estimate among amongst towards toward between within level levels extent their there this these those "
                "about into with from that which what section respondents participants study".split())


def _obj_words(t):
    return {w[:6] for w in re.findall(r"[a-z][a-z’'-]{3,}", str(t or "").lower()) if w not in _OBJ_STOP}


def instrument_gaps(plan, tool):
    """The specific objectives the instrument should answer and has no section for: [(n, objective)] (empty when each has
    one). Sections are matched to objectives by the words that tell the objectives apart (not the topic, the people or the
    place, which they share), then in order; only objectives beyond the sections there are can be left without one."""
    objs = [str(o) for o in (plan.get("specific") or []) if str(o).strip()]
    if not tool or not objs:
        return []
    guide = tool.get("kind") == "guide"
    need = [(i + 1, o) for i, o in enumerate(objs) if not _OBJ_BG.search(o) and (guide or not (_OBJ_REL.search(o) or _OBJ_DIFF.search(o)))]
    secs = [str(x.get("title") or "") for x in (tool.get("sections") or []) if isinstance(x, dict)]
    if guide:
        secs = [t for t in secs if not re.search(r"\bbackground\b|\bdemograph", t, re.I)]
    if len(secs) >= len(need):
        return []
    shared = set.intersection(*[_obj_words(o) for o in objs]) if len(objs) > 1 else set()
    shared |= _obj_words(plan.get("population")) | _obj_words(plan.get("setting"))
    ow = {n: _obj_words(o) - shared for n, o in need}
    sw = [_obj_words(t) - shared for t in secs]
    pairs = sorted(((len(ow[n] & sw[k]), n, k) for n, _ in need for k in range(len(secs)) if ow[n] & sw[k]), reverse=True)
    done_o, done_s = set(), set()
    for score, n, k in pairs:
        if n not in done_o and k not in done_s:
            done_o.add(n)
            done_s.add(k)
    left_o = [(n, o) for n, o in need if n not in done_o]
    left_s = [k for k in range(len(secs)) if k not in done_s]
    return left_o[len(left_s):]


def instrument_ask(gaps):
    """What the instrument is asked for again, naming the objectives it left without a section."""
    if not gaps:
        return "Draft it."
    return ("Draft it. Your last draft had no section for " + ("this specific objective" if len(gaps) == 1 else "these specific objectives")
            + ", which the instrument must answer: " + "; ".join("%d. %s" % (n, o.rstrip(".")) for n, o in gaps)
            + ". Give ONE section for EACH specific objective it should answer, in order, and keep the reply short enough to finish.")


# 6.6, round 70: the ethics section says what every ethics committee looks for first: approval, voluntary participation,
# informed consent, the right to withdraw, confidentiality or anonymity, and how the data will be kept. Chapter three is told
# to (ch3_prompt); a section that leaves one out is asked for once more, as round 64 does for children.
ETHICS_BASICS = (
    ("approval", re.compile(r"\b(?:approv\w*|clearance|ethic\w*\s+(?:review\s+)?(?:committee|board)|review\s+(?:committee|board)|IRB|ERC)\b", re.I),
     "that ethical approval will be obtained from the relevant ethics review committee (or the body THE PLAN names)"),
    ("voluntary", re.compile(r"\bvoluntar\w*|\bfree(?:ly)?\s+to\s+(?:take\s+part|participate|decline|refuse)|\bnot\s+(?:be\s+)?(?:forced|coerced|compelled)|"
                             r"\bwithout\s+(?:any\s+)?(?:coercion|pressure)", re.I), "that taking part is voluntary"),
    ("consent", re.compile(r"\bconsent\w*", re.I), "that informed consent will be obtained before anyone takes part"),
    ("withdraw", re.compile(r"\bwithdr[ae]w\w*|\bpull\s+out\b|\bleave\s+the\s+study|\bstop\s+(?:taking\s+part|participating)|\bdiscontinue", re.I),
     "that participants may withdraw at any time without any penalty"),
    ("confidentiality", re.compile(r"\bconfidential\w*|\banonym\w*|\bprivacy\b|\bno\s+names?\b|\bcodes?\b", re.I),
     "how confidentiality and anonymity will be kept (no names on the instrument, codes instead)"),
    ("keeping", re.compile(r"(?:^|(?<=[.!?;]))(?=[^.!?;]*\b(?:data|questionnaires?|records?|information|recordings?|transcripts?|forms?|responses|"
                           r"files?|notes)\b)(?=[^.!?;]*\b(?:kept|keep\w*|stor\w*|lock\w*|password\w*|secur\w*|safe\w*|destroy\w*|delet\w*|dispos\w*|"
                           r"access\w*|encrypt\w*)\b)", re.I),
     "how the data will be kept safe and who may see it"),       # (in one sentence with what is kept: "Confidentiality will be kept" isn't it)
)


def ethics_basics(ethics):
    """What the ethics section leaves out of the basics: ["approval", "voluntary", …] (empty when nothing)."""
    t = str(ethics or "")
    return [k for k, rx, _ in ETHICS_BASICS if not rx.search(t)]


def ethics_ask(minor_gaps, basic_gaps):
    """The ethics section asked for again: round 64's words for children, and the basics it left out."""
    words = {k: w for k, _, w in ETHICS_BASICS}
    more = [words[k] for k in basic_gaps if k in words]
    if minor_gaps and not more:
        return ETHICS_MINORS_ASK
    tail = ("It must also say " + "; ".join(more) + "; keep the rest of the section.") if more else ""
    if minor_gaps:
        return ETHICS_MINORS_ASK[:-len("; keep the rest of the section.")] + ". " + tail
    return "Write ONLY the ⟦ethics⟧ section again, after its marker. " + tail


# 6.6, round 63 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a quantitative proposal's analysis names, for an
# objective that asks about a relationship or a difference, a test that can answer it. Chapter three is told to (ch3_prompt),
# and nothing checked it: "frequencies and percentages" alone can't answer "to determine the relationship between sleep and
# grades", and a supervisor marks the method as unable to meet its objective.
_OBJ_DIFF = re.compile(r"\b(?:differen\w*|differ|compar\w*|vary|varies|between\s+(?:male|female|men|women|boys|girls|rural|urban|"
                       r"public|private)\b)", re.I)
_OBJ_REL = re.compile(r"\b(?:relationships?|relat(?:e|es|ed|ing)\s+to|associat\w*|correlat\w*|influenc\w*|effects?\s+of|impacts?\s+of|"
                      r"predict\w*|determinants?|factors?\s+(?:that\s+)?(?:associated|influencing|affecting|contributing|related|"
                      r"influence|affect|contribute)|link(?:ed|s)?\s+(?:to|with|between))\b", re.I)
_TEST_REL = re.compile(r"chi[\s-]*squ?are|\u03c7\s*2|\u03c7\u00b2|correlation|pearson|spearman|kendall|regression|odds\s+ratio|fisher\W?s\s+exact|"
                       r"test\s+of\s+(?:association|independence|relationship)", re.I)
_TEST_DIFF = re.compile(r"\bt[\s-]*tests?\b|anova|analysis\s+of\s+variance|mann[\s-]*whitney|kruskal|wilcoxon|chi[\s-]*squ?are|"
                        r"\u03c7\s*2|\u03c7\u00b2|test\s+of\s+difference|independent[\s-]+samples", re.I)


def analysis_gaps(plan, analysis):
    """The specific objectives of a quantitative (or mixed) plan whose kind of question the analysis names no test for:
    [(their number, "relationship" or "difference")]. A qualitative plan answers with themes: none."""
    if plan.get("design") in QUALITATIVE or plan.get("analysis") == "thematic" or re.search(
            r"qualitative|phenomenolog|ethnograph|grounded\s+theory|narrative\s+inquiry", str(plan.get("design_other") or ""), re.I):
        return []
    t = str(analysis or "")
    out = []
    for i, ob in enumerate(plan.get("specific") or [], 1):
        if _OBJ_DIFF.search(ob) and not _TEST_DIFF.search(t):
            out.append((i, "difference"))
        elif _OBJ_REL.search(ob) and not _TEST_REL.search(t):
            out.append((i, "relationship"))
    return out


def analysis_ask(gaps):
    """What chapter three's analysis is asked to add, in words for the AI."""
    return ("Write ONLY the ⟦analysis⟧ section again, after its marker. It must name, for EACH specific objective, the analysis "
            "that answers it. " + " ".join(
                f"Specific objective {n} asks about a {kind}: name the test of {kind} that answers it "
                + ("(for example the chi-square test of independence for categories, or a correlation for scores)" if kind == "relationship"
                   else "(for example an independent-samples t-test or ANOVA for scores, or the chi-square test for categories)")
                + ", at the 0.05 level of significance." for n, kind in gaps))


def _item_key(t):
    return re.sub(r"[^a-z0-9]+", " ", str(t or "").lower()).strip()


# 6.6, round 65: a question or statement that leads to its answer is no fair measure: "Don't you think…?", "Wouldn't you
# agree…?", "Isn't it true that…?", "Do you agree that…?" (in an interview), "How bad is…?" (the answer assumed), "Obviously",
# "Surely", "Everyone knows that…". It goes while its section keeps enough (two questions, three statements); otherwise it
# stays as drafted for the person to read.
_LEADING = re.compile(r"^\s*(?:don['\u2019]?t|doesn['\u2019]?t|didn['\u2019]?t|wouldn['\u2019]?t|shouldn['\u2019]?t|isn['\u2019]?t|aren['\u2019]?t|"
                      r"wasn['\u2019]?t|weren['\u2019]?t|won['\u2019]?t|can['\u2019]?t|couldn['\u2019]?t)\s+(?:you|it|they|this|that|there|we|people|"
                      r"students|nurses|patients)\b"
                      r"|^\s*(?:do|would|wouldn['\u2019]?t)\s+you\s+(?:not\s+)?agree\b|\bisn['\u2019]?t\s+it\s+(?:true|obvious|clear)\b"
                      r"|\b(?:obviously|surely|undoubtedly|of\s+course)\b|\beveryone\s+knows\b|\bit\s+is\s+(?:obvious|clear)\s+that\b"
                      r"|^\s*how\s+(?:bad|terrible|harmful|dangerous|awful|good|great|helpful|wonderful|useful)\s+(?:is|are|was|were)\b", re.I)


def _leading(t):
    return bool(_LEADING.search(str(t or "")))


# 6.6, round 116 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): an exclusion criterion is a reason to leave out
# someone who meets every inclusion criterion (too ill to take part, away during data collection, took part in the pre-test);
# "students who do not give consent" only repeats "students who give informed consent" the other way round, a point
# supervisors mark. Such a mirror goes while another exclusion criterion stays (chapter three is told so too).
_CRIT_NEG = re.compile(r"\b(?:not|no|non|never|unwilling(?:ness)?|refus\w*|declin\w*|fail\w*\s+to|without|absent|unable|n['’]t)\b|n['’]t\b", re.I)
_CRIT_STOP = set("a an the who whom whose which that are is was were be been being will would to of in on at for and or with by as do does did "
                 "have has had their they them those these students participants respondents nurses persons people individuals "
                 "willing unwilling refuse refused refusing decline declined declining consent consented consenting give gives given "
                 "provide provides provided informed written verbal agree agreed agreeing".split())


def _crit_words(t):
    return {w[:6] for w in re.findall(r"[a-z][a-z’'-]{2,}", str(t or "").lower()) if w not in _CRIT_STOP and not _CRIT_NEG.fullmatch(w)}


def _crit_consent(t):
    return bool(re.search(r"\b(?:consent\w*|agree\w*|willing\w*|unwilling\w*|refus\w*|declin\w*)\b", str(t or ""), re.I))


def mirror_exclusions(text):
    """Chapter three's 3.6 without the exclusion criteria that only turn an inclusion criterion round, as long as another
    exclusion criterion stays."""
    t = str(text or "")
    m = re.search(r"(?im)^\s*exclusion\s+criteria\s*:?\s*$", t)
    if not m:
        return t
    inc_part, exc_part = t[:m.start()], t[m.end():]
    incs = [x.strip() for x in re.findall(r"(?m)^\s*[-•*]\s+(.+)$", inc_part)]
    lines = exc_part.split("\n")
    bullets = [(k, re.sub(r"^\s*[-•*]\s+", "", ln).strip()) for k, ln in enumerate(lines) if re.match(r"^\s*[-•*]\s+", ln)]
    if not incs or len(bullets) < 2:
        return t
    mirrors = set()
    for k, e in bullets:
        if not _CRIT_NEG.search(e):
            continue
        ew = _crit_words(e)
        for i in incs:
            iw = _crit_words(i)
            if _CRIT_NEG.search(i):
                continue
            both_consent = _crit_consent(e) and _crit_consent(i)
            if both_consent or (iw and ew and (ew <= iw or len(iw & ew) * 2 >= len(iw | ew))):
                mirrors.add(k)
                break
    if not mirrors or len(mirrors) >= len(bullets):
        return t
    kept = [ln for k, ln in enumerate(lines) if k not in mirrors]
    return t[:m.end()] + "\n".join(kept)


# 6.6, round 98: an interview guide's questions are open. A question that can be answered "yes" or "no" ("Do you use your phone
# at night?") gives the interviewer a word, not an account. One about an objective with no open probe after it is given one,
# "(Please tell me more.)", which keeps the question and asks for the account. "Can you describe…", "Could you tell me…",
# "Is there anything else…?" and the background section are left as they are.
_CLOSED_Q = re.compile(r"^\s*(?:do|does|did|is|are|was|were|have|has|had|will|would|should|could|can|may|must|shall)\b"
                       r"(?!\s+(?:you|u)\s+(?:please\s+)?(?:describe|tell|explain|share|talk|walk|give|say\s+(?:more|something)|elaborate|narrate|"
                       r"expand|recall|list|mention|outline|discuss|comment|think\s+(?:of|about)\s+(?:a|an|any|some)\b))", re.I)
_ANYTHING_ELSE = re.compile(r"\banything\s+else\b|\bany\s+other\s+(?:thing|comment|issue|point|idea)s?\b|\bwould\s+you\s+like\s+to\s+add\b", re.I)
_OPEN_PROBE = re.compile(r"\((?=[^)]*\b(?:how|what|why|when|where|who|which|explain|describe|tell|examples?|elaborate|more|probe)\b)[^)]*\)\s*\.?\s*$", re.I)
OPEN_PROBE = "(Please tell me more.)"


def open_question(q):
    """A question of an interview guide's objective section, open: a yes-or-no question with no open probe gets one."""
    q = str(q or "").strip()
    if not q or not _CLOSED_Q.match(q) or _ANYTHING_ELSE.search(q) or _OPEN_PROBE.search(q):
        return q
    return q + " " + OPEN_PROBE


# 6.6, round 99: one thing at a time. A statement that rates two things ("I use my phone in bed and I sleep late") can't be
# answered for either: it becomes two statements when it is two simple ones (each "I …", three words or more), else it goes
# while its section keeps three. An interview question that asks two things ("What do you use your phone for at night, and
# how does it affect your sleep?", or two questions in one line) becomes two questions; a probe in brackets stays whole.
_TWO_I = re.compile(r"^(?P<a>I\s+[^,;()]+?)(?:,\s*|\s+)(?:and|or|but)\s+(?P<b>I\s+[^;()]+?)[.!]?$")
_TWO_Q = re.compile(r"^(?P<a>[^()?]+?),?\s+and\s+(?P<b>(?:how|what|why|when|where|who|which|in\s+what\s+ways?)\b[^()?]+\?)(?P<probe>\s*\([^)]*\))?\s*$", re.I)


def _words(t):
    return len(re.findall(r"[A-Za-z’']+", str(t or "")))


def split_statement(i):
    """[i] as it is, [two statements] for one that is two simple ones, or [] for one that rates two things otherwise."""
    t = str(i or "").strip()
    two = [m for m in re.finditer(r"(?:,\s*|\s)(?:and|or)\s+I\s+\w", t)
           if not re.search(r"\b(?:my|our|the|his|her|their|your)\s+[\w’'-]+(?:\s+[\w’'-]+)?$", t[:m.start()], re.I)]   # ("my friends and I")
    if not two:
        return [t]
    m = _TWO_I.match(t)
    if m and len(two) == 1 and _words(m.group("a")) >= 4 and _words(m.group("b")) >= 4:
        return [m.group("a").strip() + ".", m.group("b").strip().rstrip(".") + "."]
    return []


def split_question(q):
    """[q] as it is, or the two questions it asks."""
    t = str(q or "").strip()
    probe = ""
    pm = re.search(r"\s*\([^)]*\)\s*\.?\s*$", t)
    core = t[:pm.start()].strip() if pm else t
    probe = t[pm.start():].strip() if pm else ""
    parts = [x.strip() for x in re.split(r"(?<=\?)\s+(?=[A-Z])", core) if x.strip()]
    if len(parts) == 2 and all(_words(x) >= 3 and x.endswith("?") for x in parts):
        return [parts[0], (parts[1] + (" " + probe if probe else "")).strip()]
    m = _TWO_Q.match(core + (" " + probe if probe else ""))
    if m and _words(m.group("a")) >= 6 and _words(m.group("b")) >= 4:
        b = m.group("b").strip()
        return [m.group("a").strip().rstrip(",") + "?", (b[:1].upper() + b[1:] + (m.group("probe") or "")).strip()]
    return [t]


# 6.6, round 62 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a statement rated for how often (Never to Always)
# that says its own frequency ("I always check my phone after midnight") can't be answered: "Sometimes" to "always" means
# nothing. A word that only says how often goes ("I check my phone after midnight"); a statement that says "never",
# "rarely" or "seldom" turns round without it ("I never skip breakfast" is not "I skip breakfast"), so it goes, as long as
# its section keeps three statements (then it stays as it was: the person reads the draft).
_FREQ_POS = re.compile(r"\s*\b(?:always|often|usually|sometimes|frequently|regularly|occasionally|constantly|generally|normally|"
                       r"typically|habitually|at\s+times|all\s+the\s+time|most\s+of\s+the\s+time|every\s+(?:single\s+)?(?:day|night|"
                       r"week|time|morning|evening)|most\s+(?:days|nights|weeks|mornings|evenings))\b", re.I)
_FREQ_NEG = re.compile(r"\b(?:never|rarely|seldom|hardly\s+ever|almost\s+never|not\s+often)\b", re.I)


def _fit_frequency(items):
    out = []
    for it in items:
        if _FREQ_NEG.search(it):
            continue
        t = re.sub(r"\s{2,}", " ", _FREQ_POS.sub("", it)).strip()
        t = re.sub(r"\s+([,.;:])", r"\1", t).strip(" ,")
        out.append(t[:1].upper() + t[1:] if t else it)
    if len(out) >= 3:
        return [i for k, i in enumerate(out) if _item_key(i) not in {_item_key(x) for x in out[:k]}]
    return [re.sub(r"\s{2,}", " ", _FREQ_POS.sub("", it)).strip() or it for it in items]


# ---------------------------------------------------------------- 6.6, round 13: a proposal tells what WILL be done
# A proposal's method is a plan: "Data were collected", "Participants were recruited" or "The results showed" claim work
# that hasn't been done. In chapter three (its setting apart: a place's history is past) and the abstract's methodology and
# analysis, a sentence about the study's own steps is put in the future ("will be collected"); a sentence that reports
# results the study can't have yet is left out. A sentence with a year in it, or about earlier work, is left as it is.
_PROC_NOUN = re.compile(r"\b(?:data|participants?|respondents?|questionnaires?|interviews?|sample|samples|consent|approval|permission|"
                        r"instruments?|pre-?test(?:ing)?|pilot|analysis|responses|information|students|nurses|patients|the\s+study|"
                        r"the\s+researchers?|focus\s+groups?|records?|scores?|items?|variables?|confidentiality|anonymity)\b", re.I)
_PAST_PASSIVE = re.compile(r"\b(?:was|were)\s+((?:\w+ly\s+)?(?:collected|conducted|analy[sz]ed|recruited|selected|sampled|interviewed|"
                           r"administered|obtained|sought|distributed|coded|entered|cleaned|checked|pre-?tested|piloted|used|employed|"
                           r"adopted|applied|computed|calculated|determined|estimated|summari[sz]ed|presented|described|compared|tested|"
                           r"assessed|measured|explained|informed|assured|given|asked|invited|approached|included|excluded|stored|kept|"
                           r"transcribed|translated|recorded|gathered|retrieved|reviewed|approved|granted|performed|carried\s+out|done|"
                           r"guaranteed|maintained|protected|ensured|trained|visited|contacted|allocated|assigned|randomi[sz]ed|weighted))\b", re.I)
_PAST_ACTIVE = {"collected": "collect", "administered": "administer", "obtained": "obtain", "recruited": "recruit", "selected": "select",
                "used": "use", "analysed": "analyse", "analyzed": "analyze", "employed": "employ", "adopted": "adopt", "conducted": "conduct",
                "sought": "seek", "interviewed": "interview", "distributed": "distribute", "visited": "visit", "explained": "explain",
                "ensured": "ensure", "assured": "assure", "recorded": "record", "transcribed": "transcribe", "coded": "code", "entered": "enter"}
_ACTIVE_RX = re.compile(r"\b((?:the|a|each)\s+(?:researchers?|research\s+assistants?|study\s+team|investigators?|study))\s+(" + "|".join(_PAST_ACTIVE) + r")\b", re.I)
_RESULTS_SAID = re.compile(r"\b(?:the\s+)?(?:results?|findings?)\s+(?:of\s+(?:the|this)\s+study\s+)?(?:showed|revealed|indicated|demonstrated|"
                           r"found|suggested|confirmed)\b|\bit\s+was\s+found\s+that\b|\bthe\s+study\s+found\b", re.I)
_EARLIER = re.compile(r"\b(?:19|20)\d{2}\b|\b(?:previous|previously|earlier|prior|past|other)\s+(?:stud|research|work|surveys?|researchers?)|\bwas\s+established\b", re.I)


def future_tense(text):
    """A proposal's method in the future ("Data were collected" -> "Data will be collected"; "The researcher administered"
    -> "The researcher will administer"); a sentence reporting results the study can't have yet is taken out. Only sentences
    about the study's own steps change; one with a year or about earlier work stays as it is."""
    out = []
    for para in str(text or "").split("\n"):
        sents = re.split(r"(?<=[.!?])\s+(?=[A-Z(“\"])", para) if para.strip() else [para]
        kept = []
        for snt in sents:
            if _EARLIER.search(snt):
                kept.append(snt)
                continue
            if _RESULTS_SAID.search(snt):
                continue                                   # (a plan has no results yet)
            if _PROC_NOUN.search(snt):
                snt = _PAST_PASSIVE.sub(lambda m: "will be " + m.group(1), snt)
                snt = _ACTIVE_RX.sub(lambda m: m.group(1) + " will " + _PAST_ACTIVE[m.group(2).lower()], snt)
            kept.append(snt)
        out.append(" ".join(k for k in kept if k is not None))
    return "\n".join(out)


# ---------------------------------------------------------------- 6.6, round 9: number ranges a respondent can answer
_N = r"\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?"            # 500, 1,000, 2.5
_OPT_RANGE = re.compile(r"^(?P<pre>\D*?)(?P<a>" + _N + r")(?P<sep>\s*(?:-|–|—|to)\s*)(?P<cur>[^\d\s]{0,4})(?P<b>" + _N + r")(?P<post>\D*)$", re.I)
_OPT_BELOW = re.compile(r"^\s*(?:below|under|less\s+than|fewer\s+than|younger\s+than|<)\s*(?P<cur>[^\d\s]{0,4})(?P<a>" + _N + r")(?P<post>\D*)$", re.I)
_OPT_ATLEAST = re.compile(r"^\s*(?P<cur>[^\d\s]{0,4})(?P<a>" + _N + r")(?P<mid>\s*\+|\D*?\b(?:and|or)\s+(?:above|over|more|older|longer)\b)(?P<post>.*)$", re.I)
_OPT_ABOVE = re.compile(r"^\s*(?:above|over|more\s+than|greater\s+than|older\s+than|>)\s*(?P<cur>[^\d\s]{0,4})(?P<a>" + _N + r")(?P<post>\D*)$", re.I)


def _num_fmt(x, dec, commas=False):
    return (f"{x:,.{dec}f}" if commas else f"{x:.{dec}f}") if dec else (f"{int(round(x)):,}" if commas else str(int(round(x))))


def fix_ranges(options):
    """The answer options of a background question whose answers are numbers ("18-24", "25–29 years", "Below 18",
    "30 and above", "Above 30", "GH₵500–1,000"), made to neither overlap nor leave a gap ("20-25", "25-30" -> "20-24",
    "25-30"; "18-20", "22-25" -> "18-21", "22-25"; "25-29", "Above 30" -> "25-30", "Above 30"; "Below 18", "19-24" ->
    "Below 18", "18-24"). Options that aren't all numbers, or aren't in rising order, are left as they are."""
    parsed = []
    for o in options:
        t = str(o or "")
        for kind, rx in (("range", _OPT_RANGE), ("below", _OPT_BELOW), ("atleast", _OPT_ATLEAST), ("above", _OPT_ABOVE)):
            m = rx.match(t)
            if m:
                parsed.append((kind, m))
                break
        else:
            return list(options)
    if len(parsed) < 2:
        return list(options)
    nums = [m.group(g) for _k, m in parsed for g in ("a", "b") if g in m.groupdict() and m.group(g)]
    dec = max((len(x.split(".")[1]) for x in nums if "." in x), default=0)
    commas = any("," in x for x in nums)
    step = 10 ** -dec
    lo_hi = []
    for kind, m in parsed:
        a = float(m.group("a").replace(",", ""))
        if kind == "range":
            b = float(m.group("b").replace(",", ""))
            if b < a:
                return list(options)
            lo_hi.append((a, b))
        elif kind == "below":
            lo_hi.append((float("-inf"), a - step))
        elif kind == "atleast":
            lo_hi.append((a, float("inf")))
        else:
            lo_hi.append((a + step, float("inf")))
    if any(lo_hi[k + 1][0] < lo_hi[k][0] for k in range(len(lo_hi) - 1)) or any(h == float("inf") for _l, h in lo_hi[:-1]):
        return list(options)                  # (not in rising order, or an open top before the end)
    out = [str(o) for o in options]
    eps = step / 10
    for k in range(len(parsed) - 1):
        (kind, m), (nkind, nm) = parsed[k], parsed[k + 1]
        hi, nlo = lo_hi[k][1], lo_hi[k + 1][0]
        if abs(nlo - (hi + step)) < eps:
            continue                          # (they meet exactly)
        if kind == "range" and nlo - step > lo_hi[k][0] + eps:
            nb = nlo - step                   # the earlier range ends where the next begins
            out[k] = out[k][:m.start("b")] + _num_fmt(nb, dec, commas) + out[k][m.end("b"):]
            lo_hi[k] = (lo_hi[k][0], nb)
        elif kind == "range" and nkind == "range" and hi + step < lo_hi[k + 1][1] - eps:
            na = hi + step                    # ("1-2", "2-4": "1-1" is no range) the next one begins after it: "1-2", "3-4"
            out[k + 1] = out[k + 1][:nm.start("a")] + _num_fmt(na, dec, commas) + out[k + 1][nm.end("a"):]
            lo_hi[k + 1] = (na, lo_hi[k + 1][1])
        elif nkind == "range" and kind == "below":
            na = hi + step                    # "Below 18", "19-24" -> "18-24"
            out[k + 1] = out[k + 1][:nm.start("a")] + _num_fmt(na, dec, commas) + out[k + 1][nm.end("a"):]
            lo_hi[k + 1] = (na, lo_hi[k + 1][1])
        elif nkind == "above" and nlo > hi + step + eps:
            # "25-29", "Above 30": 30 is nobody's answer -> "30 and above"
            out[k + 1] = (nm.group("cur") or "") + _num_fmt(hi + step, dec, commas) + (nm.group("post") or "") + " and above"
            lo_hi[k + 1] = (hi + step, float("inf"))
    return [re.sub(r"\s{2,}", " ", o).strip() for o in out]


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
