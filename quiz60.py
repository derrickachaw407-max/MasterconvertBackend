"""6.0: much harder practice questions.

The owner, 15 Nov 2026 00:14, with screenshots of another app's physiotherapy practice questions: "Make the questions in
academic home extremely harder than these." What made those easy: the right option was the only sensible one; the wrong
ones were wrong by their WORDING ("only", "always", "as though this alone were a complete plan", "without reviewing red
flags", "identical on every examination"), padded with filler ("especially when clinic time is limited … in clinic
today") or of another kind than the right one. So here:
1. every level asks for more (LEVELS), and the rules (HARD_RULES) forbid what gives an answer away;
2. every multiple-choice question is read by code for what gives it away (flags): options of very different lengths,
   the key far the longest, absolute or careless words only in wrong options, "all of the above", two options alike;
3. an examiner (a second request that is NOT told the key) answers every multiple-choice question from the subject and
   the material, and says whether only one option can be defended, whether the wording gives the answer away, and how
   demanding the question really is;
4. a question that code or the examiner finds fault with is rewritten in one request; the rewrite is taken only when its
   key is one of its options and code finds nothing in it; otherwise the question stays as it was;
5. (6.6, round 48: the owner, 21 Nov 2026: "Raise the app's validity") a rewrite is answered again by an examiner who is
   not told its key: it is taken when that examiner chooses its key and finds one best answer. A question whose answer an
   examiner disputed (another option chosen, or more than one defensible) and that has no rewrite so confirmed is LEFT
   OUT: a practice question must not teach a wrong answer. Unless the examiner disputed more than half of the questions
   (then the examination itself is in doubt, and nothing is left out for it). A rewrite that only fixed the wording and
   is not confirmed gives way to the question as it was. No time for the second reading: the rewrite is taken as before;
6. (round 54) a short-answer question's model answer is checked against the material (KEYCHECK, told the answer); one it
   finds wrong is left out with its question, unless it finds most of them wrong.
Nothing here logs a word of anyone's material or questions: the caller logs counts only."""
import json
import random
import re
import time
import concurrent.futures as _futures

LEVELS = {
    "easy": ("DIFFICULTY: FOUNDATION, and still demanding. Every question needs the student to apply one idea correctly (to a "
             "short situation, an example, a result or a claim), never to recall a definition or a list. Every option is a real "
             "answer from this subject that a student who has not studied could not rule out. Mark each \"easy\"."),
    "medium": ("DIFFICULTY: EXAM STANDARD. Most questions open with a short realistic case or data (findings, figures, a result) "
               "and need two steps of reasoning: find what matters in the details, then apply the right principle. The options "
               "are close to each other: a student who only half knows the topic finds two or three of them attractive. Mark "
               "each \"medium\"."),
    "hard": ("DIFFICULTY: SPECIALIST, THE STANDARD OF A PROFESSIONAL LICENSING OR BOARD EXAM, the hardest there is. Every "
             "question is a full single-best-answer item: a vignette with specific findings (values, signs, results, timings, "
             "measurements, history) that must be interpreted, and three or more steps of reasoning, asking for the most likely "
             "cause, the best next step, the most appropriate choice, the expected finding, the exception, or what would change "
             "the decision. Every option looks right at first sight to a student who knows the topic well; only careful "
             "reasoning with the details shows the one best answer. Wrong options are right in a neighbouring situation, right "
             "but not first, or right for a condition that looks similar. Mark each \"hard\"."),
    "mixed": ("DIFFICULTY: weighted to the exam end. About one question in five is FOUNDATION (apply one idea to a short "
              "situation, never recall), two in five are EXAM STANDARD (a short case or data, two steps of reasoning, close "
              "options), two in five are SPECIALIST, the standard of a professional licensing exam (a vignette with findings "
              "to interpret, three or more steps, every option right at first sight). Mark each question \"easy\", \"medium\" "
              "or \"hard\" accordingly."),
}

HARD_RULES = (
    "\n\nHARD BY DESIGN: what makes practice questions easy is that the answer can be found without knowing the subject. "
    "Never let that happen:\n"
    "7. EVERY OPTION A REAL CONTENDER. The options of a question are all of one kind (all causes, all tests, all next "
    "steps, all values, all mechanisms), and each is something a competent person in the field would consider: right in "
    "a neighbouring situation, right but not first, right but not best here, or right for a condition that looks similar. "
    "A wrong option is wrong because of the subject, never because of how it is worded.\n"
    "8. NO WORDING THAT GIVES AN OPTION AWAY. In the options: no absolute words (always, never, only, solely, completely, "
    "entirely, identical, regardless, alone) unless the subject's own facts need them and the right option uses them as "
    "much as the wrong ones; no clause that makes an option careless or absurd (\"without checking …\", \"as though this "
    "alone were …\", \"despite warning signs\", \"ignoring …\"); no padding or filler; no \"all of the above\", \"none of the "
    "above\" or \"both A and B\". The options are within a few words of each other in length and alike in form.\n"
    "9. THE STEM CARRIES THE DIFFICULTY. The details that decide the answer are in the stem or the case (findings, numbers, "
    "history, setting); the question asks for the single BEST answer; the right option never repeats the stem's words.\n"
    "10. ONE BEST ANSWER, CHECKED. Before you write the key, answer the question yourself from the stem alone. If two "
    "options could be defended as best, or the answer can be guessed from the wording, rewrite the options until exactly "
    "one is best and none can be ruled out without knowing the subject."
)

_ABSOLUTE = re.compile(r"\b(?:always|never|only|solely|completely|entirely|totally|identical(?:ly)?|regardless|guaranteed?|impossible|"
                       r"permanently|in all cases|under any circumstances|without exception|alone|exclusively|invariably)\b", re.I)
_CARELESS = re.compile(r"\b(?:without (?:any|checking|reviewing|assessing|considering|monitoring|examining|asking|consent)|"
                       r"despite (?:clear|obvious|warning|the warning|red)|as though|as if this|ignoring|instead of (?:assessing|checking|examining)|"
                       r"and nothing else|no matter)\b", re.I)
_AOTA = re.compile(r"\b(?:all|none) of the (?:above|options|answers|choices)\b|\bboth (?:[a-d]|of these|of the above)\b|\bneither (?:[a-d]|of these)\b|"
                   r"^\s*[a-d] and [a-d]\s*$", re.I)


def _norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().lower().rstrip(".")


def is_mc(q):
    opts = q.get("options") or []
    return q.get("type") == "multiple_choice" and len(opts) >= 2 and any(_norm(o) == _norm(q.get("correct_answer")) for o in opts)


def is_tf(q):
    return sorted(_norm(o) for o in (q.get("options") or [])) == ["false", "true"]


def key_index(q):
    c = _norm(q.get("correct_answer"))
    for i, o in enumerate(q.get("options") or []):
        if _norm(o) == c:
            return i
    return -1


def flags(q):
    """What code can see gives a multiple-choice question's answer away (an empty list when nothing does)."""
    if not is_mc(q) or is_tf(q):
        return []
    opts = [str(o) for o in q["options"]]
    k = key_index(q)
    out = []
    lens = [len(o) for o in opts]
    if max(lens) > 2.2 * max(1, min(lens)) and max(lens) - min(lens) > 30:
        out.append("the options are of very different lengths")
    rest = sorted((l for i, l in enumerate(lens) if i != k), reverse=True)
    if rest and lens[k] >= 1.3 * rest[0] and lens[k] - rest[0] > 15:
        out.append("the right option is clearly the longest")
    wrong = [o for i, o in enumerate(opts) if i != k]
    abs_wrong = sum(len(_ABSOLUTE.findall(o)) for o in wrong)
    if abs_wrong >= 2 and not _ABSOLUTE.search(opts[k]) or sum(1 for o in wrong if _ABSOLUTE.search(o)) >= 2 and not _ABSOLUTE.search(opts[k]):
        out.append("absolute words (only, always, never …) are in wrong options and not in the right one")
    if any(_CARELESS.search(o) for o in wrong) and not _CARELESS.search(opts[k]):
        out.append("a wrong option is made careless or absurd by its wording")
    if any(_AOTA.search(o) for o in opts):
        out.append("an option is \"all of the above\", \"none of the above\" or \"both\"")
    if len(set(_norm(o) for o in opts)) < len(opts):
        out.append("two options say the same")
    return out


def _letters(n):
    return "ABCDEFGH"[:n]


def _as_item(q, i):
    item = {"i": i, "question": q.get("question")}
    if q.get("case_study") and q.get("case"):
        item["case"] = q["case"]
    item["options"] = {L: str(o) for L, o in zip(_letters(len(q["options"])), q["options"])}
    return item


EXAMINER = (
    "You are an examiner checking a multiple-choice exam before it is printed. You are NOT told the answer key. For each "
    "question, work out the single best option yourself from the subject (the MATERIAL below is what the exam is based on: "
    "its facts are the facts to go by), then judge the question:\n"
    "- \"answer\": the letter you choose;\n"
    "- \"only_one\": true when exactly one option can be defended as the best answer; false when two or more could be "
    "argued as best, or none is right;\n"
    "- \"giveaway\": true when a student who does not know the subject could rule options out or find the answer from the "
    "wording alone (an absurd, absolute, careless or padded option, an option of another kind, one far longer or shorter "
    "than the rest, the stem's words repeated in one option, a grammatical clue);\n"
    "- \"level\": \"foundation\" (one simple step), \"exam\" (a case or data, two steps) or \"specialist\" (a vignette with "
    "findings to interpret, three or more steps, every option right at first sight).\n"
    "Answer with JSON only: [{\"i\": 0, \"answer\": \"B\", \"only_one\": true, \"giveaway\": false, \"level\": \"exam\"}], one "
    "item for each question.")

REWRITER = (
    "You rewrite practice exam questions that an examiner found fault with, so that each becomes a fair, hard, single-best-"
    "answer question. Keep the subject and what each question tests (the MATERIAL below is what the exam is based on: never "
    "go against its facts). You may rewrite the stem, the case, the options and the explanation. Each question must end with "
    "exactly one best option; every option of one kind and a real contender (right in a neighbouring situation, right but "
    "not first, or right for something that looks similar); no absolute, careless or padded wording; no \"all/none of the "
    "above\"; options within a few words of each other in length; the deciding details in the stem; at the level asked. "
    "When the examiner chose another answer than the key, decide from the subject's facts which is right, and make the "
    "question so that only that one is right.\n"
    "Answer with JSON only: [{\"i\": 0, \"question\": \"…\", \"case\": \"…\" (only when the question has a case), \"options\": "
    "[four strings, no letters], \"correct_answer\": \"the exact text of the right option\", \"explanation\": \"one or two "
    "sentences: why it is right and why the most tempting wrong option is wrong\"}], one item for each question.")


def _json_list(text):
    m = re.search(r"\[.*\]", text or "", re.S)
    if not m:
        return []
    try:
        v = json.loads(m.group(0))
    except ValueError:
        return []
    return v if isinstance(v, list) else []


def _material(text, limit=24000):
    t = str(text or "")
    return t if len(t) <= limit else t[:limit] + "\n[…]"


def _want(level):
    return {"medium": {"foundation"}, "hard": {"foundation"}}.get(level, set())


# 6.6, round 54: a short-answer question's model answer, checked against the material (it was never read at all). The
# checker IS told the answer (a short answer can be right in many words, so an answer written blind can't be compared by
# code); it says whether that answer is right. A model answer it finds wrong is left out with its question, unless it
# finds most of them wrong (then the check itself is in doubt).
KEYCHECK = (
    "You check the answer key of short-answer practice questions before they are printed. Each QUESTION comes with its "
    "MODEL ANSWER. Using the MATERIAL below as the facts to go by (and the subject's established knowledge where the material "
    "says nothing), judge each model answer: \"ok\" when it is right (it may be brief), \"wrong\" when it states something "
    "false or answers another question, \"unsure\" when you can't tell. Answer with JSON only: [{\"i\": 0, \"v\": \"ok\"}, "
    "{\"i\": 1, \"v\": \"wrong\", \"why\": \"a few words\"}], one item for each question.")


def is_sa(q):
    return q.get("type") == "short_answer" and bool(str(q.get("correct_answer") or "").strip()) and bool(str(q.get("question") or "").strip())


def _check_keys(questions, sa, mat, ask, deadline, now):
    """The short-answer questions whose model answer the checker finds wrong ([] when none; None when it couldn't run)."""
    wrong, answered = [], 0
    for k in range(0, min(len(sa), 60), 20):
        batch = sa[k:k + 20]
        items = [dict({"i": i, "question": questions[i].get("question"), "model answer": questions[i].get("correct_answer")},
                      **({"case": questions[i]["case"]} if questions[i].get("case_study") and questions[i].get("case") else {})) for i in batch]
        user = "MATERIAL:\n<<<\n" + mat + "\n>>>\n\nQUESTIONS:\n" + json.dumps(items, ensure_ascii=False)
        got = _json_list(ask(KEYCHECK, user, 120 + 40 * len(batch), max(15, min(60, int(deadline - now() - 30)))))
        for x in got:
            if isinstance(x, dict) and isinstance(x.get("i"), int) and x["i"] in batch:
                answered += 1
                if str(x.get("v") or "").strip().lower() == "wrong":
                    wrong.append(x["i"])
    return wrong if answered else None


def harden(questions, difficulty, material, ask, deadline, now=time.time, standalone=lambda s: s):
    """questions as the quiz endpoint parsed them (options already in their final order). ask(system, user, max_tokens,
    timeout) -> the AI's text (it may raise). Returns (questions, stats). Never raises."""
    stats = {"mc": 0, "code": 0, "disagree": 0, "not_one": 0, "giveaway": 0, "low": 0, "asked": 0, "rewritten": 0, "kept": 0,
             "examined": 0, "failed": False, "confirmed": 0, "not_confirmed": 0, "left_out": 0, "sa": 0, "sa_wrong": 0}
    try:
        mat = _material(material)
        sa = [i for i, q in enumerate(questions) if is_sa(q)]
        stats["sa"] = len(sa)
        pool = _futures.ThreadPoolExecutor(max_workers=1)
        sa_job = pool.submit(_check_keys, questions, sa, mat, ask, deadline, now) if sa and deadline - now() > 25 else None
        drop = _harden_mc(questions, difficulty, mat, ask, deadline, now, standalone, stats)
        if sa_job is not None:
            try:
                wrong = sa_job.result(timeout=max(1, deadline - now() + 5))
            except Exception:
                wrong = None
                stats["failed"] = True
            if wrong:
                stats["sa_wrong"] = len(wrong)
                if len(wrong) <= max(1, len(sa) // 2):
                    drop = list(drop) + list(wrong)
        pool.shutdown(wait=False)
        return _left_out(questions, drop, stats), stats
    except Exception:
        stats["failed"] = True
        return questions, stats


def _harden_mc(questions, difficulty, mat, ask, deadline, now, standalone, stats):
    """The multiple-choice questions' checks (see the top of this file): rewrites put in place; returns the questions to
    leave out (round 48)."""
    mc = [i for i, q in enumerate(questions) if is_mc(q)]
    stats["mc"] = len(mc)
    if not mc:
        return []
    faults = {i: flags(questions[i]) for i in mc}
    stats["code"] = sum(1 for i in mc if faults[i])
    # ---- the examiner, in batches at once
    verdict = {}
    left = deadline - now()
    if left > 25:
        batches = [mc[k:k + 25] for k in range(0, len(mc), 25)][:8]

        def examine(batch):
            items = [_as_item(questions[i], i) for i in batch]
            user = "MATERIAL:\n<<<\n" + mat + "\n>>>\n\nQUESTIONS:\n" + json.dumps(items, ensure_ascii=False)
            return _json_list(ask(EXAMINER, user, 200 + 60 * len(batch), max(15, min(70, int(deadline - now() - 40)))))
        with _futures.ThreadPoolExecutor(max_workers=min(4, len(batches))) as pool:
            for f in [pool.submit(examine, b) for b in batches]:
                try:
                    for x in f.result():
                        if isinstance(x, dict) and isinstance(x.get("i"), int) and x["i"] in faults:
                            verdict[x["i"]] = x
                except Exception:
                    stats["failed"] = True
    stats["examined"] = len(verdict)
    want_up = _want(difficulty)
    disputed = set()                                 # (round 48: the examiner chose another answer, or found no one best)
    for i, v in verdict.items():
        q = questions[i]
        letter = str(v.get("answer") or "").strip().upper()[:1]
        k = key_index(q)
        if letter and letter in _letters(len(q["options"])) and _letters(len(q["options"])).index(letter) != k:
            faults[i].append(f"an examiner who was not told the key chose {letter}, not {_letters(len(q['options']))[k]}")
            stats["disagree"] += 1
            disputed.add(i)
        if v.get("only_one") is False:
            faults[i].append("more than one option can be defended as best, or none is right")
            stats["not_one"] += 1
            disputed.add(i)
        if v.get("giveaway") is True and not is_tf(q):
            faults[i].append("the wording gives the answer away")
            stats["giveaway"] += 1
        if str(v.get("level") or "").lower() in want_up:
            faults[i].append(f"it is easier than asked (asked: {'specialist' if difficulty == 'hard' else 'exam'} standard)")
            stats["low"] += 1
    todo = [i for i in mc if faults[i]]
    stats["asked"] = len(todo)
    # the examination is trusted to leave questions out only when it disputed at most half of what it read
    trusted = bool(verdict) and len(disputed) <= max(1, len(verdict) // 2)
    if not todo or deadline - now() < 30:
        drop = [i for i in todo if i in disputed] if trusted else []
        stats["kept"] = len(todo) - len(drop)
        return drop
    # ---- the rewrite, in batches at once
    level_word = {"easy": "foundation", "medium": "exam standard", "hard": "specialist (licensing-exam) standard"}.get(difficulty, "")

    def rewrite(batch):
        items = []
        for i in batch:
            q = questions[i]
            it = _as_item(q, i)
            it["key"] = _letters(len(q["options"]))[key_index(q)]
            it["explanation"] = q.get("explanation") or ""
            it["level asked"] = level_word or {"easy": "foundation", "medium": "exam standard", "hard": "specialist"}.get(q.get("difficulty") or "", "exam standard")
            it["faults"] = faults[i]
            if is_tf(q):
                it["form"] = "true or false: the options stay exactly [\"True\", \"False\"] and the question is a statement"
            items.append(it)
        user = "MATERIAL:\n<<<\n" + mat + "\n>>>\n\nQUESTIONS TO REWRITE:\n" + json.dumps(items, ensure_ascii=False)
        return _json_list(ask(REWRITER, user, 300 + 420 * len(batch), max(15, min(80, int(deadline - now() - 10)))))
    batches = [todo[k:k + 10] for k in range(0, len(todo), 10)][:6]
    got = {}
    with _futures.ThreadPoolExecutor(max_workers=min(4, len(batches))) as pool:
        for f in [pool.submit(rewrite, b) for b in batches]:
            try:
                for x in f.result():
                    if isinstance(x, dict) and isinstance(x.get("i"), int) and x["i"] in faults:
                        got[x["i"]] = x
            except Exception:
                stats["failed"] = True
    cand = {}
    for i in todo:
        new = _taken(questions[i], got.get(i), standalone)
        if new:
            cand[i] = new
    # ---- round 48: every rewrite answered again by an examiner who isn't told its key
    confirm = {}
    if cand and deadline - now() > 20:
        order = list(cand)
        batches2 = [order[k:k + 25] for k in range(0, len(order), 25)][:4]

        def reexamine(batch):
            items = [_as_item(cand[i], i) for i in batch]
            user = "MATERIAL:\n<<<\n" + mat + "\n>>>\n\nQUESTIONS:\n" + json.dumps(items, ensure_ascii=False)
            return _json_list(ask(EXAMINER, user, 200 + 60 * len(batch), max(10, min(45, int(deadline - now() - 5)))))
        with _futures.ThreadPoolExecutor(max_workers=min(4, len(batches2))) as pool:
            for f in [pool.submit(reexamine, b) for b in batches2]:
                try:
                    for x in f.result():
                        if isinstance(x, dict) and isinstance(x.get("i"), int) and x["i"] in cand:
                            q2 = cand[x["i"]]
                            letter = str(x.get("answer") or "").strip().upper()[:1]
                            ok = (letter in _letters(len(q2["options"])) and _letters(len(q2["options"])).index(letter) == key_index(q2)
                                  and x.get("only_one") is not False)
                            confirm[x["i"]] = ok
                except Exception:
                    stats["failed"] = True
    drop = []
    for i in todo:
        new, ok = cand.get(i), confirm.get(i)
        if new is not None and ok is not False:      # confirmed, or no second reading in time (as before)
            questions[i] = new
            stats["rewritten"] += 1
            if ok:
                stats["confirmed"] += 1
            continue
        if new is not None:
            stats["not_confirmed"] += 1
        if i in disputed and trusted:
            drop.append(i)                            # its answer stays in doubt: left out
        elif i in disputed and new is not None:
            questions[i] = new                        # (an examiner in doubt about most questions: as before)
            stats["rewritten"] += 1
        else:
            stats["kept"] += 1                        # only its wording was at fault: the question as it was
    return drop


def _left_out(questions, drop, stats):
    """The quiz without the questions whose answer stays in doubt (round 48), counted in stats["left_out"]."""
    if not drop:
        return questions
    stats["left_out"] += len(drop)
    gone = set(drop)
    return [q for j, q in enumerate(questions) if j not in gone]


def _taken(q, x, standalone):
    """The rewritten question, when it may replace the old one: its key is one of its options, it has as many options as
    before, a case where the old one had a case, and code finds nothing in it. Its options are shuffled."""
    if not isinstance(x, dict):
        return None
    opts = x.get("options")
    if not isinstance(opts, list) or len(opts) != len(q.get("options") or []) or not all(isinstance(o, str) and o.strip() for o in opts):
        return None
    opts = [standalone(re.sub(r"^\s*[A-Ha-h][).:]\s+", "", o).strip()) for o in opts]
    stem = standalone(str(x.get("question") or "").strip())
    key = standalone(re.sub(r"^\s*[A-Ha-h][).:]\s+", "", str(x.get("correct_answer") or "")).strip())
    if len(stem) < 12 or not key:
        return None
    new = dict(q)
    new.update({"question": stem, "options": opts, "correct_answer": key,
                "explanation": standalone(str(x.get("explanation") or "").strip()) or q.get("explanation")})
    if q.get("case_study") and q.get("case"):
        case = standalone(str(x.get("case") or "").strip())
        new["case"] = case[:2400] if len(case) > 40 else q["case"]
    if not is_mc(new) or flags(new):
        return None
    if is_tf(new):
        new["options"] = ["True", "False"]
    else:
        random.shuffle(new["options"])
    return new
