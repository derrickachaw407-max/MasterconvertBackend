"""6.6, round 55 (the owner, 21 Nov 2026 00:47: "Raise the app's validity. Do as much as a research tool can be."): RamzAI's
notes of a meeting say where each point came from ("[Ama 10:42]", "[Slide 3]"), and nothing held a point to that. Now:
  - a bracket must name a moment of the meeting: a line that person said or wrote then (the time within TIME_SLACK
    minutes: the AI rounds, a caption runs on) or a slide that was shown; a bracket that names no one and no time
    ("[inaudible]") is the point's own words;
  - the point's figures, dates and names must be in the lines it names (server/app.py _mn_rules), and an independent
    check then reads the point with only those lines (_v65_verify with MEET_VERIFY_SYSTEM);
  - a point whose brackets name no moment, or moments that don't hold its figures, is pointed to the one line of the
    meeting that holds it (the people it named first), when there is such a line; otherwise it goes;
  - a heading left with no point goes; the person is told how many points went (check_note, never inside the notes:
    the notes are emailed to everyone who came).
Pure Python: no AI, no database. Made-up people in tests/test_meetnotes66.py."""
import re

TIME_SLACK = 2                 # minutes between a bracket's time and the line's
NAME_ONLY_LINES = 3            # "[Ama]" stands for her lines that share most words with the point, at most this many
REPOINT_TRIES = 8              # lines tried, best first, when a point is pointed to where it was said
QUESTION_FOLLOW = 3            # under "Questions asked", the lines after a question (its answer) are read with it

_BRACKET = re.compile(r"\[([^\[\]\n]{1,160})\]")
_TIME = re.compile(r"(?<![\d:.])([01]?\d|2[0-3])[:.]([0-5]\d)(?![\d])(?:\s*([ap])\.?\s?m\b\.?)?", re.I)
_SLIDE = re.compile(r"\bslides?\s*(?:no\.?\s*|#\s*)?(\d{1,3})(?:\s*(?:[-–—]|to|and|&)\s*(\d{1,3}))?", re.I)
_SEP = re.compile(r"\s*(?:[;,/&+]|\band\b|\bthen\b)\s*", re.I)
_HEAD = re.compile(r"^\s*(?:#{1,6}\s+(.*?)\s*#*\s*|\*\*([^*]{1,80})\*\*\s*:?)\s*$")
_ITEM = re.compile(r"^(\s*)(?:[-*+•]|\d{1,3}[.)])\s+(.*)$")
_RULE = re.compile(r"^\s*(?:-{3,}|\*{3,}|_{3,})\s*$")
_WORD = re.compile(r"[^\W\d_][\w'’\-]*", re.U)
# words a bracket uses that are not a person's name
_NOT_NAME = {"chat", "captions", "caption", "said", "says", "say", "question", "questions", "asked", "asks", "answer",
             "answered", "live", "the", "at", "in", "from", "by", "also", "see", "meeting", "slide", "slides", "time",
             "am", "pm", "a", "m", "p", "written", "wrote", "message", "messages", "transcript", "spoken", "audio",
             "inaudible", "sic", "unclear", "crosstalk", "laughter", "silence", "music", "noise", "pause", "applause"}
_TITLES = {"dr", "mr", "mrs", "ms", "miss", "prof", "professor", "sir", "madam", "rev", "hon", "eng"}
_STOP = set("""a an the and or but if then so of to in on at by for with from into about over after before during under
between than that this these those there here it its is are was were be been being has have had do does did will would
shall should can could may might must not no yes we our you your they their them he she his her him i me my us also just
very more most less many much some any all each every other such own same only both either neither what which who whom
whose when where why how as up down out off again once further too said says say asked ask told tell going get got let
please okay ok yeah well like think know want need make made one two three four five six seven eight nine ten""".split())
_COMMON_CAPS = {"The", "This", "These", "Those", "There", "Their", "They", "When", "What", "Which", "While", "With", "Without",
                "From", "Each", "Every", "Some", "Most", "Many", "More", "Less", "And", "But", "For", "Not", "All", "Any", "Its",
                "Our", "Your", "Key", "Note", "Notes", "Also", "However", "Because", "Therefore", "Then", "After", "Before",
                "During", "Summary", "Decisions", "Decision", "Questions", "Question", "Next", "Steps", "Answer", "Agreed",
                "Action", "Actions", "Slide", "Slides", "Chat", "Host", "Everyone", "Nobody", "None", "Yes", "Okay"}
_PROPER = re.compile(r"(?<![.!?:]\s)(?<!^)\b([A-Z][a-zà-ɏ]{2,}(?:[-\s][A-Z][a-zà-ɏ]{2,})*)\b")


def minute(t):
    """Minutes into the day of "HH:MM" ("10:42", "9.05", "2:05 pm"), or None."""
    m = _TIME.search(str(t or ""))
    if not m:
        return None
    h, mi = int(m.group(1)), int(m.group(2))
    if m.group(3):
        ap = m.group(3).lower()
        if ap == "p" and h < 12:
            h += 12
        elif ap == "a" and h == 12:
            h = 0
    return h * 60 + mi


def _name_words(name):
    return [w for w in (x.lower().strip(".'’-") for x in _WORD.findall(str(name or ""))) if w and w not in _TITLES]


def name_fits(said_name, speaker):
    """A bracket's name is this speaker's: the same words, or some of them (the first name, the surname), with no word
    that the speaker's name doesn't have ("Ama" or "Ama Mensah" for "Ama Mensah"; not "Kofi" for "Ama Mensah")."""
    a, b = _name_words(said_name), _name_words(speaker)
    if not a or not b:
        return False
    return set(a) <= set(b) or set(b) <= set(a)


def make_lines(said, chat):
    """The meeting's lines, in the order the notes' material gave them. said, chat: [(time "HH:MM", name, text)]."""
    out = []
    for stream, items in (("said", said or []), ("chat", chat or [])):
        for k, (t, name, text) in enumerate(items):
            text = re.sub(r"\s+", " ", str(text or "")).strip()
            if text:
                out.append({"stream": stream, "k": k, "time": str(t or "")[:5], "minute": minute(t), "name": str(name or ""),
                            "text": text})
    return out


def label(line):
    return f"[{line['name']} {line['time']}]" if line.get("time") else f"[{line['name']}]"


def _speakers(lines):
    return sorted({l["name"] for l in lines if l.get("name")}, key=len, reverse=True)


def _names(text):
    got = []
    for part in _SEP.split(text or ""):
        ws = [w for w in _name_words(part) if w not in _NOT_NAME]
        if ws:
            got.append(" ".join(ws))
    return got


def _looks_like_names(s):
    """Bracket words with a capital that aren't a slide, a time or a common word: a name ("[Ama]", "[Esi Darko]"; not
    "[inaudible]", "[Agreed]")."""
    ws = [w for w in _WORD.findall(s) if w.lower() not in _NOT_NAME and w.lower() not in _TITLES]
    return bool(ws) and len(ws) <= 4 and all(w[:1].isupper() and w not in _COMMON_CAPS and w.lower() not in _STOP for w in ws)


def refs_of(inner, speakers):
    """The moments one bracket names: ("slide", n), ("at", [names], minute) for a person (or anyone, when no name is
    given) at a time, ("by", [names]) for a person. [] when the bracket names no moment: it is then the point's own
    words ("[inaudible]", "[sic]"). A name that is no one in the meeting still counts as a name: the bracket points to
    no one."""
    s = str(inner or "")
    out = []
    for m in _SLIDE.finditer(s):
        a = int(m.group(1))
        b = int(m.group(2) or a)
        if b < a or b > a + 30:
            b = a
        out += [("slide", n) for n in range(a, b + 1)]
    s = _SLIDE.sub(" ; ", s)
    times = list(_TIME.finditer(s))
    if not times:
        ns = _names(s)
        if ns and (any(any(name_fits(n, sp) for sp in speakers) for n in ns) or (not out and _looks_like_names(s))):
            out.append(("by", ns))
        return out
    time_first = not _names(s[:times[0].start()])
    for i, m in enumerate(times):
        if time_first:
            seg = s[m.end():times[i + 1].start() if i + 1 < len(times) else len(s)].split(";")[0]
        else:
            seg = s[times[i - 1].end() if i else 0:m.start()].split(";")[-1]
        ns = _names(seg)
        if not ns and out and out[-1][0] == "at":
            ns = list(out[-1][1])           # "[Ama 10:42, 10:45]": the same person twice
        out.append(("at", ns, minute(m.group(0))))
    return out


def _near(a, b):
    if a is None or b is None:
        return False
    d = abs(a - b) % 1440
    return min(d, 1440 - d) <= TIME_SLACK


def words(t):
    return {w for w in (x.lower().strip("'’-") for x in _WORD.findall(str(t or ""))) if len(w) >= 3 and w not in _STOP}


def overlap(a, b):
    return len(words(a) & words(b))


def match(ref, lines, slides, body=""):
    """The keys of the moments a ref names: ("l", i) a line, ("s", n) a slide."""
    if ref[0] == "slide":
        n = ref[1]
        return [("s", n)] if 1 <= n <= len(slides) and str(slides[n - 1] or "").strip() else []
    if ref[0] == "at":
        names, mi = ref[1], ref[2]
        return [("l", i) for i, l in enumerate(lines)
                if _near(mi, l["minute"]) and (not names or any(name_fits(n, l["name"]) for n in names))]
    names = ref[1]
    mine = [i for i, l in enumerate(lines) if any(name_fits(n, l["name"]) for n in names)]
    mine.sort(key=lambda i: -overlap(body, lines[i]["text"]))
    best = [i for i in mine[:NAME_ONLY_LINES] if overlap(body, lines[i]["text"]) > 0]
    return [("l", i) for i in (best or mine[:1])]


def quote_of(key, lines, slides, labelled=True):
    """A moment's words: with who and when for the independent reading; without them for the rules (a time in a label is
    no figure of what was said)."""
    if key[0] == "s":
        t = str(slides[key[1] - 1] or "")
        return f"[Slide {key[1]}] {t}" if labelled else t
    l = lines[key[1]]
    return f"{label(l)} {l['text']}" if labelled else l["text"]


def _quotes(keys, lines, slides, labelled):
    return " | ".join(quote_of(k, lines, slides, labelled) for k in keys)


def names_ok(body, material, extra=""):
    """Every name in the point (a word with a capital inside a sentence) is a word of the meeting: what was said, written
    and shown, the people's names and the meeting's title. A name heard a letter wrong in the captions ("Mensa" for
    "Mensah", five letters or more) is the same name."""
    low = (str(material or "") + " " + str(extra or "")).lower()
    have = {w.lower().strip("'’-") for w in _WORD.findall(low)}
    for m in _PROPER.finditer(str(body or "")):
        for w in m.group(1).split():
            w = w.strip("-")
            if not w or w in _COMMON_CAPS:
                continue
            lw = w.lower()
            if lw in have or re.search(r"(?<![a-zà-ɏ])" + re.escape(lw) + r"(?![a-zà-ɏ])", low):
                continue
            if len(lw) >= 5 and any(h[:1] == lw[:1] and _one_edit(lw, h) for h in have):
                continue
            return False
    return True


def _one_edit(a, b):
    if a == b:
        return True
    if abs(len(a) - len(b)) > 1:
        return False
    if len(a) == len(b):
        return sum(1 for x, y in zip(a, b) if x != y) == 1
    if len(a) > len(b):
        a, b = b, a
    i = j = d = 0
    while i < len(a) and j < len(b):
        if a[i] != b[j]:
            d += 1
            if d > 1:
                return False
            j += 1
        else:
            i += 1
            j += 1
    return True


def points(md):
    """The notes' points: [(line number, the section it is under, its words without the brackets, its brackets
    [(start, end, inside)])]. Headings, rules and empty lines are not points."""
    out = []
    section = ""
    for i, raw in enumerate(str(md or "").split("\n")):
        s = raw.rstrip()
        if not s.strip() or _RULE.match(s):
            continue
        h = _HEAD.match(s)
        if h:
            section = (h.group(1) or h.group(2) or "").strip()
            continue
        m = _ITEM.match(s)
        text = m.group(2) if m else s.strip()
        brackets = [(b.start(), b.end(), b.group(1)) for b in _BRACKET.finditer(s)]
        body = _BRACKET.sub(" ", text)
        body = re.sub(r"\s+([.,;:!?])", r"\1", re.sub(r"\s+", " ", body)).strip(" -–—")
        if body:
            out.append((i, section, body, brackets))
    return out


def plan(md, lines, slides, rules, material, extra=""):
    """What the check of the notes against the meeting does before the independent reading:
    {"verify": [(line number, "(under section) words", quotes)], "drop": {line: why}, "fix": {line: new line},
     "repointed": n, "points": n, "unsourced": n}. rules(body, words of the lines) -> [what those words don't hold]."""
    speakers = _speakers(lines)
    rows = str(md or "").split("\n")
    rep = {"verify": [], "drop": {}, "fix": {}, "repointed": 0, "points": 0, "unsourced": 0}

    def holder(body, prefer=()):
        """The one line of the meeting that holds the point (the people it named first), or None."""
        cands = sorted(range(len(lines)), key=lambda i: (
            0 if any(name_fits(n, lines[i]["name"]) for n in prefer) else 1, -overlap(body, lines[i]["text"])))
        for i in cands[:REPOINT_TRIES]:
            if overlap(body, lines[i]["text"]) >= 2 and not rules(body, lines[i]["text"]):
                return i
        return None

    for ln, section, body, brackets in points(md):
        rep["points"] += 1
        where = f"(under {section or 'the notes'}) {body}"
        if not names_ok(body, material, extra):
            rep["drop"][ln] = "name"            # a name the meeting never had: no moment of it can hold the point
            continue
        refs, ref_brackets = [], []
        for b in brackets:
            rs = refs_of(b[2], speakers)
            if rs:
                refs.append(rs)
                ref_brackets.append(b)
        if not refs:
            # no bracket names a moment: read with the lines that share most words with it, when any do
            best = [i for i in sorted(range(len(lines)), key=lambda i: -overlap(body, lines[i]["text"]))[:NAME_ONLY_LINES]
                    if overlap(body, lines[i]["text"]) >= 2]
            rep["unsourced"] += 1
            if not best:
                bad = rules(body, " | ".join(l["text"] for l in lines) + " | " + " | ".join(str(x) for x in slides))
                if bad:
                    rep["drop"][ln] = bad[0]
                continue                        # nothing to read it with; its figures and names are the meeting's
            keys = [("l", i) for i in best]
            bad = rules(body, _quotes(keys, lines, slides, False))
            if bad:
                one = holder(body)
                if one is None:
                    rep["drop"][ln] = bad[0]
                    continue
                keys = [("l", one)]
            rep["verify"].append((ln, where, _quotes(keys, lines, slides, True)))
            continue
        keys, firsts, named, stray = [], [], [], 0
        for rs in refs:
            for r in rs:
                got = match(r, lines, slides, body)
                if r[0] in ("at", "by"):
                    named += r[1]
                if not got:
                    stray += 1
                    continue
                keys += [k for k in got if k not in keys]

                def rank(k, r=r):
                    near = -abs((lines[k[1]]["minute"] or 0) - (r[2] or 0)) if k[0] == "l" and r[0] == "at" else 0
                    return (overlap(body, quote_of(k, lines, slides, False)), near)
                top = max(got, key=rank)
                if top not in firsts:
                    firsts.append(top)
        read = list(keys)
        if section and re.search(r"question", section, re.I):
            for k in keys:                      # a question's answer is in the lines after it
                if k[0] == "l":
                    st = lines[k[1]]["stream"]
                    read += [("l", j) for j in range(k[1] + 1, min(len(lines), k[1] + 1 + QUESTION_FOLLOW))
                             if lines[j]["stream"] == st and ("l", j) not in read]
        bad = rules(body, _quotes(read, lines, slides, False)) if keys else ["nowhere"]
        if bad:
            one = holder(body, named)
            if one is None:
                rep["drop"][ln] = bad[0]
                continue
            read = [("l", one)]
            rep["fix"][ln] = _rebracket(rows[ln], ref_brackets, label(lines[one]))
            rep["repointed"] += 1
        elif stray:
            # some of its brackets name no moment: they now name only the moments that hold it
            rep["fix"][ln] = _rebracket(rows[ln], ref_brackets, " ".join(
                (f"[Slide {k[1]}]" if k[0] == "s" else label(lines[k[1]])) for k in firsts[:4]))
            rep["repointed"] += 1
        rep["verify"].append((ln, where, _quotes(read, lines, slides, True)))
    return rep


def _rebracket(row, brackets, new):
    """The line with its brackets (those that name moments) replaced by one: new, where the last of them stood."""
    if not brackets:
        return row.rstrip() + " " + new
    out, last = [], 0
    for k, (a, b, _inner) in enumerate(brackets):
        out.append(row[last:a])
        if k == len(brackets) - 1:
            out.append(new)
        last = b
    out.append(row[last:])
    s = re.sub(r"(?<=\S) {2,}", " ", "".join(out).rstrip())
    return re.sub(r"\s+([.,;:!?])$", r"\1", s)


def rebuild(md, fix, drop):
    """The notes with the fixed lines, without the dropped ones; a heading left with nothing under it goes."""
    rows = str(md or "").split("\n")
    out = [fix.get(i, r) for i, r in enumerate(rows) if i not in drop]
    for _round in range(3):
        keep = [True] * len(out)
        for k, r in enumerate(out):
            if not _HEAD.match(r):
                continue
            body = False
            for nxt in out[k + 1:]:
                if _HEAD.match(nxt):
                    break
                if nxt.strip() and not _RULE.match(nxt):
                    body = True
                    break
            keep[k] = body
        new = [r for r, ok in zip(out, keep) if ok]
        if len(new) == len(out):
            break
        out = new
    return re.sub(r"\n{3,}", "\n\n", "\n".join(out)).strip()
