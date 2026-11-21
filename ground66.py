"""6.6, round 1 (the owner, 21 Nov 2026 00:47: "Raise the app's validity. Do as much as a research tool can be."):
what the app makes FROM THE PERSON'S OWN FILE (a summary's slides, study notes: a revision sheet, key terms, exam
questions, a mind map) says only what that file says. Before 6.6 the writer was only told so ("Use ONLY what is in the
material"); now every point is checked against the file's own sentences (server/app.py _mat_check: a check that names
the file's sentences behind each point, the rules of server/valid65.py on those sentences, then a second, independent
check), and what the file doesn't say is taken out.

Here only what needs no AI: the file cut into labelled sentences, the points read out of notes and slides, the notes and
slides put back together without the points that went, and the check by words alone (figures, dates and names must be
the file's own) that stands in when the AI can't check. The tests call these directly (tests/test_ground66.py)."""
import re

MAT_SENTS_MAX = 450            # the file's sentences shown to the check (20,000 characters are about 250)
MAT_SENT_CHARS = 420           # a longer line is cut where its sentences end
TOPIC_ONLY_WORDS = 60          # fewer words than this is a topic, not a file: nothing to check against

_SENT_END = re.compile(r"(?<=[.!?])[\"”’)]?\s+(?=[\"“‘(]?[A-Z0-9À-ɏ])")
_NOT_END = re.compile(r"(?:\bet al|\be\.g|\bi\.e|\bvs|\bcf|\bDr|\bNo|\bpp?|\bSt|\bFig|\bEq|\b[A-Z])\.$")


def split_material(text):
    """The file as the check reads it: its lines, each cut into sentences (a list, a table row or a heading is a line of
    its own), whitespace tidied; nothing is added or reworded."""
    out = []
    for raw in str(text or "").splitlines():
        line = re.sub(r"\s+", " ", raw).strip(" \t|")
        if not line or not re.search(r"[A-Za-z0-9À-ɏ]", line):
            continue
        parts, start = [], 0
        for m in _SENT_END.finditer(line):
            if _NOT_END.search(line[max(0, m.start() - 6):m.start()]):
                continue
            parts.append(line[start:m.start()].strip())
            start = m.end()
        parts.append(line[start:].strip())
        for p in parts:
            while len(p) > MAT_SENT_CHARS:
                cut = p.rfind(" ", 0, MAT_SENT_CHARS)
                cut = cut if cut > 80 else MAT_SENT_CHARS
                out.append(p[:cut].strip())
                p = p[cut:].strip()
            if p:
                out.append(p)
    return out[:MAT_SENTS_MAX]


def labelled(sents):
    return "\n".join(f"M{k + 1}: {t}" for k, t in enumerate(sents))


def labels_ok(raw, n_sents):
    """The material's sentences a check named, verified: "M4" (or 4), each a real sentence. None when any is not."""
    items = raw if isinstance(raw, list) else [raw]
    out = []
    for x in items[:4]:
        if isinstance(x, bool):
            return None
        if isinstance(x, int):
            k = x
        else:
            m = re.fullmatch(r"\s*M?\s*(\d{1,4})\s*", str(x or ""), re.I)
            if not m:
                return None
            k = int(m.group(1))
        if not 1 <= k <= n_sents:
            return None
        out.append(k - 1)
    return out or None


def is_topic_only(text):
    return len(re.findall(r"[A-Za-zÀ-ɏ]{2,}", str(text or ""))) < TOPIC_ONLY_WORDS


# ---------------------------------------------------------------- the points of study notes (Markdown)
_BULLET = re.compile(r"^(\s*)(?:[-*+•]|\d{1,3}[.)])\s+(.*)$")
_HEAD = re.compile(r"^\s*#{1,6}\s")
_GOOD_ANSWER = re.compile(r"^\s*(?:\*\*)?\s*a good answer\s*:?\s*(?:\*\*)?\s*:?\s*$", re.I)
_MARKS = re.compile(r"\(\s*\d{1,3}\s*marks?\s*\)", re.I)


def plain(t):
    t = re.sub(r"\*\*(.+?)\*\*", r"\1", str(t or ""))
    t = re.sub(r"(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?![\w*])", r"\1", t)
    t = re.sub(r"`([^`]+)`", r"\1", t)
    return re.sub(r"\s+", " ", t).strip()


def notes_points(md, kind):
    """The lines of study notes that state something, each [line number, its words, how deep, role]. Headings, a rule,
    "A good answer:" and the exam questions themselves state nothing of their own and are not points; in a mind map
    every item is a point (2 to 7 words)."""
    out = []
    lines = str(md or "").split("\n")
    for i, line in enumerate(lines):
        s = line.rstrip()
        if not s.strip() or _HEAD.match(s) or re.fullmatch(r"\s*(?:-{3,}|\*{3,}|_{3,})\s*", s) or _GOOD_ANSWER.match(s):
            continue
        m = _BULLET.match(s)
        depth = len(m.group(1).replace("\t", "  ")) if m else 0
        words = plain(m.group(2) if m else s)
        if not words:
            continue
        if kind == "questions" and (_MARKS.search(words) or words.rstrip().endswith("?")) and depth == 0:
            continue                      # an exam question asks; its "A good answer" points are what is checked
        out.append([i, words, depth, "point"])
    return out


def _depth(l):
    m = _BULLET.match(l)
    return len(m.group(1).replace("\t", "  ")) if m else 0


def _is_question(l, kind):
    if kind != "questions" or not l.strip() or _HEAD.match(l) or _depth(l):
        return False
    q = plain(_BULLET.match(l).group(2) if _BULLET.match(l) else l)
    return bool(_MARKS.search(q) or q.endswith("?"))


def _is_point(l, kind):
    return bool(l.strip()) and not _HEAD.match(l) and not _GOOD_ANSWER.match(l) and not _is_question(l, kind) \
        and not re.fullmatch(r"\s*(?:-{3,}|\*{3,}|_{3,})\s*", l)


def notes_without(md, drop, kind):
    """The notes without the lines in drop (line numbers), each with the deeper points under it (they only go on about
    it); then a heading, an exam question (with its "A good answer:") or a lone "A good answer:" left with no point under
    it goes too. The title (the first line) always stays."""
    lines = str(md or "").split("\n")
    gone = set()
    for i in sorted(drop):
        if i in gone or i >= len(lines):
            continue
        gone.add(i)
        if not _BULLET.match(lines[i]):
            continue
        d = _depth(lines[i])
        for j in range(i + 1, len(lines)):
            if not lines[j].strip():
                continue
            if _BULLET.match(lines[j]) and _depth(lines[j]) > d:
                gone.add(j)
                continue
            break
    out = [l for i, l in enumerate(lines) if i not in gone]
    for _round in range(3):                     # (a heading emptied by an emptied question goes on the next round)
        keep = [True] * len(out)
        for k, l in enumerate(out):
            if k == 0:
                continue
            h = re.match(r"^\s*(#{1,6})\s", l)
            if h:
                lvl = len(h.group(1))
                body = False
                for m in out[k + 1:]:
                    hm = re.match(r"^\s*(#{1,6})\s", m)
                    if hm and len(hm.group(1)) <= lvl:
                        break
                    if _is_point(m, kind) or _is_question(m, kind):
                        body = True
                        break
                keep[k] = body
            elif _is_question(l, kind):
                body = False
                for m in out[k + 1:]:
                    if _HEAD.match(m) or _is_question(m, kind):
                        break
                    if _is_point(m, kind):
                        body = True
                        break
                if not body:
                    keep[k] = False
                    for j in range(k + 1, len(out)):          # its "A good answer:" goes with it
                        if not out[j].strip():
                            continue
                        if _GOOD_ANSWER.match(out[j]):
                            keep[j] = False
                        break
            elif _GOOD_ANSWER.match(l):
                body = False
                for m in out[k + 1:]:
                    if _HEAD.match(m) or _is_question(m, kind) or _GOOD_ANSWER.match(m):
                        break
                    if _is_point(m, kind):
                        body = True
                        break
                keep[k] = keep[k] and body
        new = [l for l, ok in zip(out, keep) if ok]
        if len(new) == len(out):
            break
        out = new
    return re.sub(r"\n{3,}", "\n\n", "\n".join(out)).strip()


# ---------------------------------------------------------------- the points of a summary's slides
def slide_points(slides):
    """The points of a summary's content slides, each [key, its words]: ("b", slide, i) a bullet, ("p", slide, section,
    i) a sub-heading's point, ("h", slide) a heading that gives a figure (a heading that only names a topic states nothing)."""
    out = []
    for si, sl in enumerate(slides or []):
        h = plain(sl.get("header") or "")
        if re.search(r"\d", h):
            out.append([("h", si), h])
        for bi, b in enumerate(sl.get("bullets") or []):
            if plain(b):
                out.append([("b", si, bi), plain(b)])
        for ci, sec in enumerate(sl.get("sections") or []):
            for pi, p in enumerate(sec.get("points") or []):
                if plain(p):
                    out.append([("p", si, ci, pi), plain(p)])
    return out


def slides_without(slides, drop):
    """The slides without the points in drop (keys of slide_points); a slide whose heading's figure isn't the file's,
    or that is left with no point and no chart, goes."""
    drop = set(drop)
    out = []
    for si, sl in enumerate(slides or []):
        if ("h", si) in drop:
            continue
        sl = dict(sl)
        sl["bullets"] = [b for bi, b in enumerate(sl.get("bullets") or []) if ("b", si, bi) not in drop]
        if sl.get("sections"):
            secs = []
            for ci, sec in enumerate(sl["sections"]):
                pts = [p for pi, p in enumerate(sec.get("points") or []) if ("p", si, ci, pi) not in drop]
                if pts:
                    secs.append(dict(sec, points=pts))
            sl["sections"] = secs or None
        if ("c", si) in drop:
            sl["chart"] = None
        if sl.get("bullets") or sl.get("sections") or sl.get("chart"):
            out.append(sl)
    return out


# ---------------------------------------------------------------- the check by words alone
_NUM = re.compile(r"(?<![\w.])\d[\d,]*(?:\.\d+)?")
_PROPER = re.compile(r"(?<![.!?:]\s)(?<!^)\b([A-Z][a-zà-ɏ]{2,}(?:[-\s][A-Z][a-zà-ɏ]{2,})*)\b")
_COMMON_CAPS = {"The", "This", "These", "Those", "There", "Their", "They", "When", "What", "Which", "While", "With", "Without", "From",
                "Each", "Every", "Some", "Most", "Many", "More", "Less", "And", "But", "For", "Not", "All", "Any", "Its", "Our",
                "Your", "Key", "Remember", "Note", "Notes", "Example", "Examples", "Definition", "Types", "Causes", "Effects",
                "Signs", "Symptoms", "Treatment", "Summary", "Introduction", "Conclusion", "Overview", "Main", "Other", "First",
                "Second", "Third", "Finally", "Also", "However", "Because", "Therefore", "Then", "After", "Before", "During"}


def _num_tokens(t):
    return {re.sub(r"[,]", "", x).rstrip(".") for x in _NUM.findall(str(t or ""))}


def words_ok(point, material):
    """The check that needs no AI: every figure and date in the point is in the file, and every name (a word with a
    capital inside a sentence) is a word of the file."""
    mat = str(material or "")
    nums = _num_tokens(mat)
    if not _num_tokens(point) <= nums:
        return False
    low = mat.lower()
    for m in _PROPER.finditer(str(point or "")):
        w = m.group(1)
        if w in _COMMON_CAPS:
            continue
        if w.lower() not in low:
            return False
    return True


def chart_ok(chart, material):
    """A chart's figures and the figures in its labels are the file's own."""
    if not isinstance(chart, dict):
        return True
    nums = _num_tokens(material)
    vals = []
    for v in chart.get("values") or []:
        try:
            f = float(v)
        except (TypeError, ValueError):
            return False
        vals.append(("%d" % f) if f == int(f) else ("%g" % f))
    for lab in chart.get("labels") or []:
        vals += list(_num_tokens(lab))
    return all(v in nums for v in vals)
