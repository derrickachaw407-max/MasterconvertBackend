"""5.7: slides made from research work are a research DEFENCE.

The owner, 13 Nov 2026 11:41: "A slides made out of research work should be a research defense in other for the user to
use for their research defense." and 11:56: "The defense should follow the other of research preferred choices in the
work and why they chosed them(eg; research topic, problem statement, objectives, the significance of the study, study,
design , setting, population, sample sample size and so on). And/or every thing that the user needs to defend themselves
and their work and be able to convince people that it is good research work and they know what they are doing."

So a defence deck goes choice by choice, in the order of the work; each slide says WHAT was chosen and WHY; and the
person also gets what to say on each slide, the questions a panel is likely to ask with answers from their own work,
and a list of what the work does not say yet (a panel will ask for it).

What keeps it honest (rule 9: everything presented as sourced must be valid; nothing invented):
- WHAT was chosen is taken only from the work. RamzAI copies the work's own words beside each choice; a choice whose
  words cannot be found in the work is left out and listed under "your work does not say".
- A REASON the work itself gives carries the work's own words too. A reason the work does not give is RamzAI's
  suggestion (the usual methodological reason for that choice). It is MARKED as a suggestion in the app and in the
  presenter's notes, for the person to check before presenting.
- A number that is not in the work never reaches a slide, a note or an answer.

Nothing here calls the AI or the network: server/app.py hands in `ask`. No name, place or figure in this file comes
from anyone's real document.
"""
import json
import re

TEXT_MAX = 150000          # characters of the work that are read (about 60 pages)
ROOM_A, ROOM_B = 5200, 6000     # room for the words of each of the two answers

# ---------------------------------------------------------------- is it research work?
_MARKS = [r"problem\s+statement|statement\s+of\s+the\s+problem", r"research\s+objectives?|general\s+objective|specific\s+objectives?|aims?\s+and\s+objectives",
          r"research\s+questions?|hypothes[ie]s", r"significance\s+of\s+the\s+study|justification\s+of\s+the\s+study|rationale",
          r"literature\s+review|review\s+of\s+(?:related\s+)?literature", r"methodology|research\s+methods?|materials\s+and\s+methods",
          r"(?:study|research)\s+design", r"(?:study|target)\s+population", r"sample\s+size|sampling\s+(?:technique|method|procedure)",
          r"data\s+collection", r"data\s+analysis|method\s+of\s+analysis", r"ethical\s+(?:considerations?|clearance|approval)",
          r"scope\s+of\s+the\s+study|delimitations?", r"background\s+(?:of|to)\s+the\s+study", r"chapter\s+(?:one|two|three|four|five|[1-5i]+)\b",
          r"recommendations?", r"findings|results\s+and\s+discussion"]
_MARKS = [re.compile(m, re.I) for m in _MARKS]


def marks_found(text):
    t = str(text or "")
    return sum(1 for m in _MARKS if m.search(t))


def looks_like_research(text):
    """Research work names its own parts (a problem statement, objectives, a methodology…). Five of them is not chance."""
    return marks_found(text) >= 5


# ---------------------------------------------------------------- the choices of a study, in the order of the work
# key, the slide's heading, the sub-heading for a long statement of the choice, the sub-heading for its reasons,
# what a panel calls it when it is missing, and whether a proposal / a finished study is expected to state it
ITEMS_A = [
    ("topic", "Background: Why This Topic", "The topic", "Why this topic", "the background to the topic", True, True),
    ("problem", "Problem Statement", "The problem", "Why it must be studied", "the problem statement", True, True),
    ("general", "Research Objectives", "General objective", "Why these objectives", "the general objective", True, True),
    ("specific", "Research Objectives", "Specific objectives", "", "the specific objectives", True, True),
    ("questions", "Research Questions", "", "", "the research questions or hypotheses", False, False),
    ("significance", "Significance of the Study", "Who the study will help", "", "the significance of the study", True, True),
    ("scope", "Scope of the Study", "What the study covers", "Why these limits", "the scope of the study", False, False),
    ("limitations", "Limitations", "The limits of the study", "How they are handled", "the limitations", False, False),
]
ITEMS_B = [
    ("design", "Study Design", "The design", "Why this design", "the study design", True, True),
    ("setting", "Study Setting", "The setting", "Why this setting", "the study setting", True, True),
    ("population", "Study Population", "The population", "Why this population", "the study population", True, True),
    ("criteria", "Inclusion and Exclusion Criteria", "Who is included and who is not", "Why these criteria", "the inclusion and exclusion criteria", False, False),
    ("size", "Sample Size", "The sample size", "How it was worked out", "the sample size and how it was worked out", True, True),
    ("sampling", "Sampling Technique", "The technique", "Why this technique", "the sampling technique", True, True),
    ("instrument", "Research Instrument", "The instrument", "Why this instrument", "the research instrument", True, True),
    ("quality", "Validity and Reliability", "How the instrument is tested", "Why this matters", "how the instrument's validity and reliability are tested", False, False),
    ("collection", "Data Collection Procedure", "The procedure", "Why this way", "the data collection procedure", True, True),
    ("analysis", "Data Analysis", "The analysis", "Why these methods", "the data analysis", True, True),
    ("ethics", "Ethical Considerations", "What will be done", "Why it matters", "the ethical considerations", True, True),
    ("results", "Results", "What the study found", "", "the results", False, True),
    ("discussion", "Discussion", "What the findings mean", "", "the discussion of the findings", False, False),
    ("conclusion", "Conclusion", "What the study concludes", "", "the conclusion", False, True),
    ("recommendations", "Recommendations", "What the study recommends", "", "the recommendations", False, True),
    ("plan", "Work Plan", "When each step is done", "", "the work plan", False, False),
    ("budget", "Budget", "What the study will cost", "", "the budget", False, False),
]
ITEMS = {k: (h, lab, why, missing, prop, study) for k, h, lab, why, missing, prop, study in ITEMS_A + ITEMS_B}
ORDER = [k for k, *_ in ITEMS_A] + ["literature", "gap"] + [k for k, *_ in ITEMS_B]
LISTS = ("specific", "questions", "significance", "criteria", "results", "recommendations", "limitations", "ethics", "plan", "budget", "conclusion", "discussion")
GROUPS = [("Background and the problem", ("topic", "problem")), ("Objectives and research questions", ("general", "specific", "questions")),
          ("Significance and scope", ("significance", "scope", "limitations")), ("Literature review", ("literature", "gap")),
          ("Methodology", ("design", "setting", "population", "criteria", "size", "sampling", "instrument", "quality", "collection", "analysis", "ethics")),
          ("Results and discussion", ("results", "discussion")), ("Conclusion and recommendations", ("conclusion", "recommendations")),
          ("Work plan and budget", ("plan", "budget"))]

_COMMON = (
    "You help a researcher prepare to DEFEND their own research work before a panel. You are given the work itself. You read "
    "it and return ONE JSON object and nothing else: no code fences, no words before or after, starting with { and ending with }.\n\n"
    "THE RULES, each one checkable:\n"
    "1. ONLY THE WORK. What was chosen, every fact, figure, name, place, date and citation comes from the work, with its exact "
    "meaning. Never add a fact about the study that the work does not state. Keep in-text citations exactly as the work writes "
    "them, at the end of the point that uses them. Never invent a citation.\n"
    "2. THE WORK'S OWN WORDS BESIDE EACH CHOICE. \"src\" is 6 to 14 consecutive words copied EXACTLY, letter for letter, from the "
    "place in the work that states that choice. It is used to check your answer against the work.\n"
    "3. REASONS. \"why\" holds 2 to 4 reasons for the choice, each at most 20 words. First the reasons the WORK itself gives: "
    "for each of those, \"src\" is 6 to 14 consecutive words copied exactly from where the work gives it. When the work gives "
    "fewer than two reasons, add the usual methodological reason a good researcher would give for THAT choice in THIS study "
    "(its objectives, its population), with \"src\" as an empty string: that marks it as your suggestion. A suggested reason is "
    "true of the method in general; it never states a fact about the study, a number, or something the researcher did.\n"
    "4. SHORT AND PLAIN. A point is at most 20 words, a full thought, in plain words. No emoji, no bold or italic marks, no "
    "symbols used as decoration.\n"
    "5. WHAT TO SAY. \"say\" is what the presenter can say aloud for that slide: 1 or 2 sentences, at most 45 words, in the first "
    "person ({voice}), only from the work.\n"
    "6. IF THEY ASK. \"ask\" holds 1 or 2 questions a panel is likely to ask about that choice, each with a short, confident answer "
    "(at most 40 words) drawn from the work; \"src\" is 6 to 14 words copied exactly from the work that support the answer, or an "
    "empty string when the answer is your suggestion of what a good researcher would say.\n"
    "7. WHAT THE WORK DOES NOT SAY. When the work does not state something, its key is null. Do not fill a gap.\n"
    "8. Write in the language of the work.\n\n"
)
_ITEM = ('{"what": ["<1 to 3 points: what was chosen or stated>"], "src": "<the work\'s own words>", '
         '"why": [{"t": "<a reason>", "src": "<the work\'s own words, or empty for your suggestion>"}], '
         '"say": "<what to say>", "ask": [{"q": "<a likely question>", "a": "<the answer>", "src": "<the work\'s own words, or empty>"}]}')
PROMPT_A = _COMMON + (
    "THIS TIME: the study and why it is needed (its introduction and its literature review). Return these keys:\n"
    '  "title": the title of the work, copied from it (at most 30 words)\n'
    '  "kind": "proposal" when the study is still to be done (it is written as a plan and reports no results), "study" when it '
    "reports results or findings\n"
    '  "people": "one" or "group", as the work shows who wrote it ("" when it does not show)\n'
    "  then one object for each of these, shaped as " + _ITEM + ", or null when the work does not state it:\n"
    '  "topic": what the study is about, in one point; its reasons are the facts in the background that show why the topic '
    "matters (from the wider world down to the place of the study), each with its citation\n"
    '  "problem": the problem the study addresses; its reasons say why it must be studied now (what is not known, what goes '
    "wrong if nothing is done)\n"
    '  "general": the general objective or aim, in the work\'s words; its reasons say why this aim follows from the problem\n'
    '  "specific": the specific objectives, one point each, in the work\'s words (up to 6); "why" may be empty\n'
    '  "questions": the research questions or hypotheses, one point each (up to 6); "why" may be empty\n'
    '  "significance": who will use the findings and how, one point for each group (up to 5); "why" may be empty\n'
    '  "scope": what the study covers (people, place, time, subject); its reasons say why it is limited so\n'
    '  "limitations": the limits the work itself names (up to 4); its reasons say how each is handled, when the work says\n'
    '  "literature": {"themes": [{"h": "<a theme of the review, at most 7 words>", "points": ["<what the studies found, with '
    'their citations>", "..."], "src": "<the work\'s own words>"}], "say": "...", "ask": [...]} with up to 3 themes of 2 or 3 '
    "points each, or null\n"
    '  "gap": what earlier studies have not answered that this study will; its reasons may be empty\n'
    '  "references": the entries of the work\'s reference list for the citations you kept on the points, copied exactly as '
    "the work writes them (up to 12); [] when the work has no reference list\n"
)
PROMPT_B = _COMMON + (
    "THIS TIME: how the study is done and, when it is finished, what it found (its methodology and what follows it). Return one "
    "object for each of these keys, shaped as " + _ITEM + ", or null when the work does not state it:\n"
    '  "design": the study design, named in one short point (at most 9 words)\n'
    '  "setting": where the study is done, in one short point, with what the work says about the place\n'
    '  "population": who is studied, with their number when the work gives it\n'
    '  "criteria": who is included and who is excluded, one point each, beginning "Included:" or "Excluded:"; "why" may be empty\n'
    '  "size": the sample size, in one short point; "why" gives how it was worked out, step by step with the work\'s own numbers '
    "(the formula, the population, the margin of error, any allowance), and why it is enough\n"
    '  "sampling": the sampling technique, named in one short point (at most 9 words)\n'
    '  "instrument": the tool for collecting data, in one short point, with its parts when the work describes them\n'
    '  "quality": how validity and reliability are ensured (a pre-test, an expert review, a reliability figure)\n'
    '  "collection": how the data are collected: who, how, over what time\n'
    '  "analysis": how the data are analysed: the software and the methods, as the work names them\n'
    '  "ethics": the ethical steps (approval, permission, consent, confidentiality), one point each; "why" may be empty\n'
    '  "results": for a finished study, its main findings, one point each with the work\'s own figures (up to 6); "why" may be empty\n'
    '  "discussion": what the findings mean and how they compare with earlier studies (up to 4 points); "why" may be empty\n'
    '  "conclusion": the conclusion (up to 4 points); "why" may be empty\n'
    '  "recommendations": the recommendations, one point each (up to 6); "why" may be empty\n'
    '  "plan": for a proposal, the work plan: each stage with its time, as the work gives it (up to 6); "why" may be empty\n'
    '  "budget": for a proposal, the budget: its main items and its total, as the work gives them (up to 6); "why" may be empty\n'
    '  "general_asks": 3 or 4 more questions a panel often asks about the whole study (what is new in it, what could go wrong, '
    'how the findings will be used), as [{"q": "...", "a": "...", "src": "..."}]\n'
)


def prompts(people=""):
    voice = "\"we\", because several people present" if people == "group" else "\"I\" for one presenter, \"we\" when the work was written by a group"
    return PROMPT_A.replace("{voice}", voice), PROMPT_B.replace("{voice}", voice)


# ---------------------------------------------------------------- reading the answers against the work
_WORD = re.compile(r"[a-z0-9À-ɏ]+")
_NUM = re.compile(r"(?<![A-Za-z])\d[\d,]*(?:\.\d+)?%?")


def _norm(text):
    return " ".join(_WORD.findall(str(text or "").lower()))


def _numbers(text):
    """The figures in a text, as people write them: 1,250 = 1250; 62% = 62."""
    out = set()
    for m in _NUM.findall(str(text or "")):
        n = m.replace(",", "").rstrip("%").rstrip(".")
        if n:
            out.add(n)
            if "." in n:
                out.add(n.rstrip("0").rstrip("."))
    return out


class Work:
    """The work, ready to be checked against."""
    def __init__(self, text):
        self.text = str(text or "")
        self.norm = " " + _norm(self.text) + " "
        self.numbers = _numbers(self.text)

    def has(self, quote):
        """Are these the work's own words? Exactly; or, for a copy with a slip in it, most of its runs of four words."""
        q = _norm(quote)
        words = q.split()
        if len(words) < 4:
            return False
        if " " + q + " " in self.norm:
            return True
        runs = [" ".join(words[i:i + 4]) for i in range(len(words) - 3)]
        hit = sum(1 for r in runs if " " + r + " " in self.norm)
        return len(words) >= 6 and hit >= max(2, int(len(runs) * 0.6 + 0.999))

    def figures_ok(self, text):
        """Every figure in the text is in the work. (A single digit is a count of points more often than a figure.)"""
        return all(n in self.numbers for n in _numbers(text) if len(n.replace(".", "")) > 1 or "." in n)


def _words(v, cap, clean):
    return str(clean(re.sub(r"\s+", " ", str(v or "")).strip()) or "").strip()[:cap]


def _sentences_ok(text, work):
    """A text without the sentences that carry a figure the work does not have."""
    keep = [s for s in re.split(r"(?<=[.!?])\s+", str(text or "").strip()) if s and work.figures_ok(s)]
    return " ".join(keep).strip()


def read_item(key, raw, work, clean):
    """One choice as the AI returned it → what can stand on a slide, or None when the work does not state it."""
    if not isinstance(raw, dict):
        return None
    cap = 6 if key in LISTS else 3
    what = [w for w in (_words(x, 260, clean) for x in (raw.get("what") if isinstance(raw.get("what"), list) else [raw.get("what")])) if w][:cap]
    what = [w for w in what if work.figures_ok(w)]
    if not what or not work.has(raw.get("src")):
        return None
    why = []
    for r in (raw.get("why") or [])[:5]:
        if isinstance(r, str):
            r = {"t": r, "src": ""}
        if not isinstance(r, dict):
            continue
        t = _words(r.get("t"), 240, clean)
        if not t or not work.figures_ok(t):
            continue
        why.append({"t": t, "own": bool(r.get("src")) and work.has(r.get("src"))})
    why = (sorted(why, key=lambda r: not r["own"]))[:4]
    say = _words(_sentences_ok(raw.get("say"), work), 420, clean)
    return {"key": key, "what": what, "why": why, "say": say, "ask": read_asks(raw.get("ask"), work, clean, 2)}


def read_asks(raw, work, clean, most):
    out = []
    for a in (raw or []):
        if not isinstance(a, dict):
            continue
        q, ans = _words(a.get("q"), 220, clean), _words(_sentences_ok(a.get("a"), work), 380, clean)
        if q and ans:
            out.append({"q": q, "a": ans, "own": bool(a.get("src")) and work.has(a.get("src"))})
    return out[:most]


def read_literature(raw, work, clean):
    if not isinstance(raw, dict):
        return None
    themes = []
    for t in (raw.get("themes") or [])[:3]:
        if not isinstance(t, dict) or not work.has(t.get("src")):
            continue
        h = _words(t.get("h"), 90, clean)
        pts = [p for p in (_words(x, 260, clean) for x in (t.get("points") or [])) if p and work.figures_ok(p)][:3]
        if h and pts:
            themes.append({"h": h, "points": pts})
    if not themes:
        return None
    return {"key": "literature", "themes": themes, "say": _words(_sentences_ok(raw.get("say"), work), 420, clean), "ask": read_asks(raw.get("ask"), work, clean, 2)}


def read_references(raw, work):
    out = []
    for r in (raw or [])[:14]:
        r = re.sub(r"\s+", " ", str(r or "")).strip()[:600]
        if len(r) > 20 and work.has(" ".join(r.split()[:9])) and r not in out:
            out.append(r)
    return out[:12]


def parse(raw):
    t = str(raw or "")
    a, b = t.find("{"), t.rfind("}")
    if a < 0 or b <= a:
        return None
    try:
        obj = json.loads(t[a:b + 1])
    except ValueError:
        return None
    return obj if isinstance(obj, dict) else None


# ---------------------------------------------------------------- the deck, in its fixed order
def _slide(header, sections, key, notes=None, marks=None):
    sections = [s for s in sections if s.get("subheading") or s.get("points")]
    flat = [x for s in sections for x in ([s["subheading"]] if s.get("subheading") else []) + list(s.get("points") or [])][:10]
    out = {"header": header, "sections": sections, "bullets": flat, "icon": "", "chart": None, "is_cta": False, "key": key}
    if notes and (notes.get("say") or notes.get("ask")):
        out["notes"] = notes
    if marks:
        out["marks"] = marks
    return out


def _item_slide(it):
    """A choice and its reasons on one slide: the choice named first (as the sub-heading when it is short), then why."""
    head, label, why_label, *_ = ITEMS[it["key"]]
    what, sections, marks = it["what"], [], []
    if len(what) == 1 and len(what[0]) <= 64 and it["key"] not in LISTS:
        sections.append({"subheading": what[0].rstrip("."), "points": []})
    else:
        sections.append({"subheading": label, "points": what})
    if it["why"] and why_label:
        sections.append({"subheading": why_label, "points": [r["t"] for r in it["why"]]})
        marks = [[len(sections) - 1, i] for i, r in enumerate(it["why"]) if not r["own"]]
    return _slide(head, sections, it["key"], {"say": it["say"], "ask": it["ask"]}, marks)


def build(a, b, work, clean, presenters=None, title_given=""):
    """The two answers → (title, kind, slides, what the work does not say, the questions to be ready for, how many
    reasons are suggestions). Raises ValueError when the text does not read as research work."""
    a, b = a if isinstance(a, dict) else {}, b if isinstance(b, dict) else {}
    items = {}
    for key, *_ in ITEMS_A:
        it = read_item(key, a.get(key), work, clean)
        if it:
            items[key] = it
    for key, *_ in ITEMS_B:
        it = read_item(key, b.get(key), work, clean)
        if it:
            items[key] = it
    lit = read_literature(a.get("literature"), work, clean)
    gap = read_item("gap", a.get("gap"), work, clean) if isinstance(a.get("gap"), dict) else None
    core = [k for k in ("problem", "general", "specific", "design", "population", "size", "sampling", "instrument", "analysis") if k in items]
    if len(core) < 3:
        raise ValueError("not research work")
    kind = "study" if (str(a.get("kind")).strip().lower() == "study" or "results" in items) and ("results" in items or "conclusion" in items) else "proposal"
    if kind == "proposal":
        for k in ("results", "discussion"):
            items.pop(k, None)
    else:
        for k in ("plan", "budget"):
            items.pop(k, None)
    title = _words(title_given or a.get("title"), 200, clean) or "Research work"
    people = [str(x).strip()[:80] for x in (presenters or []) if str(x).strip()][:10]

    slides = []
    # the objectives stand together on one heading: the general one, the specific ones, then why
    if "general" in items or "specific" in items:
        g, s = items.get("general"), items.get("specific")
        sections, marks = [], []
        if g:
            sections.append({"subheading": "General objective", "points": g["what"][:1]})
        if s:
            sections.append({"subheading": "Specific objectives", "points": s["what"]})
        why = (g or {}).get("why") or []
        if why:
            sections.append({"subheading": "Why these objectives", "points": [r["t"] for r in why]})
            marks = [[len(sections) - 1, i] for i, r in enumerate(why) if not r["own"]]
        notes = {"say": " ".join(x for x in ((g or {}).get("say"), (s or {}).get("say")) if x)[:600], "ask": ((g or {}).get("ask") or []) + ((s or {}).get("ask") or [])}
        objectives = _slide("Research Objectives", sections, "objectives", notes, marks)
    else:
        objectives = None
    for key in ORDER:
        if key == "general":
            if objectives:
                slides.append(objectives)
        elif key == "specific":
            continue
        elif key == "literature":
            if lit:
                slides.append(_slide("Literature Review", [{"subheading": t["h"], "points": t["points"]} for t in lit["themes"]], "literature",
                                     {"say": lit["say"], "ask": lit["ask"]}))
        elif key == "gap":
            if gap:
                sections = [{"subheading": "What is not yet known", "points": gap["what"]}]
                if gap["why"]:
                    sections.append({"subheading": "What this study adds", "points": [r["t"] for r in gap["why"]]})
                slides.append(_slide("The Gap This Study Fills", sections, "gap", {"say": gap["say"], "ask": gap["ask"]},
                                     [[1, i] for i, r in enumerate(gap["why"]) if not r["own"]]))
        elif key in items:
            slides.append(_item_slide(items[key]))

    have = set(items) | ({"literature"} if lit else set()) | ({"gap"} if gap else set())
    outline = [name for name, keys in GROUPS if any(k in have for k in keys)]
    refs = read_references(a.get("references"), work)
    plain = lambda header, lines, lk: [{"header": header, "bullets": lines[i:i + 6], "icon": "", "chart": None, "is_cta": False, "list": lk} for i in range(0, len(lines), 6)]
    deck = [{"kind": "cover", "header": title, "bullets": ["Research proposal defence" if kind == "proposal" else "Research defence"], "icon": "", "chart": None, "is_cta": False}]
    deck += plain("Presented by", people or ["Your name and index number"], "names")
    deck += plain("Presentation Outline", outline + (["References"] if refs else []), "outline")
    deck += slides
    if refs:
        deck += plain("References", refs, "refs")
    deck.append({"kind": "thanks", "header": "Thank you", "bullets": ["Your questions are welcome"], "icon": "", "chart": None, "is_cta": False})

    # what the work does not say, in the order a panel would notice it
    gaps = [ITEMS[k][3] for k in ORDER if k in ITEMS and k not in have and ITEMS[k][4 if kind == "proposal" else 5]
            and not (k == "general" and "specific" in have) and not (k == "specific" and "general" in have)]
    if not lit:
        gaps.insert(min(len(gaps), 3), "a literature review (what earlier studies found)")
    # the questions to be ready for: the ones on the slides, then the ones about the whole study
    asked, seen = [], set()
    for s in slides:
        for q in (s.get("notes") or {}).get("ask") or []:
            k = _norm(q["q"])
            if k and k not in seen:
                seen.add(k)
                asked.append(dict(q, on=s["header"]))
    for q in read_asks(b.get("general_asks"), work, clean, 4):
        k = _norm(q["q"])
        if k and k not in seen:
            seen.add(k)
            asked.append(dict(q, on=""))
    suggested = sum(len(s.get("marks") or []) for s in slides)
    return title, kind, deck, gaps, asked[:18], suggested


def notes_text(notes, marks_n=0):
    """A slide's presenter's notes, as they stand in the PowerPoint file."""
    if not isinstance(notes, dict):
        return ""
    out = []
    say = re.sub(r"\s+", " ", str(notes.get("say") or "")).strip()
    if say:
        out += ["WHAT TO SAY", say, ""]
    asks = [a for a in (notes.get("ask") or []) if isinstance(a, dict) and a.get("q") and a.get("a")]
    if asks:
        out.append("IF THEY ASK")
        for a in asks[:3]:
            out += ["Q: " + re.sub(r"\s+", " ", str(a["q"])).strip(), "A: " + re.sub(r"\s+", " ", str(a["a"])).strip()
                    + ("" if a.get("own") else "  (RamzAI's suggestion: make it yours.)"), ""]
    if marks_n:
        out.append(("One reason" if marks_n == 1 else f"{marks_n} reasons") + " on this slide " + ("is" if marks_n == 1 else "are")
                   + " RamzAI's suggestion, not from your work: check that " + ("it is" if marks_n == 1 else "each is") + " true of your study.")
    return "\n".join(out).strip()
