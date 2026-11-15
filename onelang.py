"""6.0: one language in everything RamzAI writes.

The owner, 15 Nov 2026, with a picture of a research proposal's Word file ("Institutional 管理 and administration"):
"Translate each and every word in the work well and appropriately. And correct all associated issues like that."

The AI sometimes lets a word of another language slip into a piece: a word in another script ("管理" in an English
sentence), or a letter of another alphabet that only looks Latin inside a Latin word ("dаta" with a Cyrillic "а").
What RamzAI hands over is read for both before anyone sees it:
- a letter that only LOOKS Latin, inside a Latin word, becomes that Latin letter; full-width punctuation ("，" "。")
  becomes ordinary punctuation in a piece that is not in Chinese, Japanese or Korean. Nothing is asked of the AI.
- a word in a script the piece doesn't use is put into the piece's language by one short request to the AI with only
  the sentences that hold such a word, the word marked. What comes back is taken only when nothing else in the sentence
  changed (its citations, numbers and other words), or when the AI kept the word because the sentence is about it (a
  word quoted as an example in a study of that language). When the AI can't answer, the word is taken out.
The scripts a piece may use: Latin always (names, citations, symbols); the script of its language when the language is
known, otherwise the script most of its letters are in; and every script in the person's own words and in the papers'
titles and authors. Greek letters standing alone (α, β, χ², μ, Δ) are symbols, never a slip.
Nothing here keeps or logs a word of anyone's text: the caller logs counts only."""
import difflib
import functools
import json
import re
import unicodedata

_SCRIPT_RANGES = (
    ("Han", "⺀-⿟々〇〡-〩〸-〻㐀-䶿一-鿿豈-﫿\U00020000-\U0003134f"),
    ("Kana", "ぁ-ゖゝ-ゟァ-ヺー-ヿㇰ-ㇿｦ-ﾝ"),
    ("Hangul", "ᄀ-ᇿㄱ-ㆎꥠ-꥿가-힣ힰ-퟿ﾠ-ￜ"),
    ("Cyrillic", "Ѐ-ԯᲀ-ᲈⷠ-ⷿꙀ-ꚟ"),
    ("Greek", "Ͱ-ͳͶͷͻ-ͽͿΆΈ-Ͽἀ-῾"),
    ("Arabic", "؀-؅ؐ-ؚ؝-يٖ-ٟ٪-ۓە-ۿݐ-ݿࡰ-ࣿﭐ-﷿ﹰ-ﻼ"),
    ("Hebrew", "֑-״יִ-ﭏ"),
    ("Devanagari", "ऀ-ॣ०-ॿ꣠-ꣿ"),
    ("Bengali", "ঀ-৿"), ("Gurmukhi", "਀-੿"), ("Gujarati", "઀-૿"), ("Oriya", "଀-୿"),
    ("Tamil", "஀-௿"), ("Telugu", "ఀ-౿"), ("Kannada", "ಀ-೿"), ("Malayalam", "ഀ-ൿ"),
    ("Sinhala", "඀-෿"), ("Thai", "฀-๿"), ("Lao", "຀-໿"), ("Tibetan", "ༀ-࿿"),
    ("Myanmar", "က-႟"), ("Georgian", "Ⴀ-ჿᲐ-Ჿ"), ("Ethiopic", "ሀ-᎟ⶀ-⷟"),
    ("Cherokee", "Ꭰ-᏿"), ("Khmer", "ក-៿"), ("Mongolian", "᠀-᢯"), ("Armenian", "Ա-֊"),
    ("Nko", "߀-߿"), ("Tifinagh", "ⴰ-⵿"), ("Vai", "ꔀ-꘿"), ("Thaana", "ހ-޿"),
    ("Syriac", "܀-ݏ"),
)
_SCRIPT_RX = re.compile("|".join(f"(?P<{name}>[{rng}])" for name, rng in _SCRIPT_RANGES))

# the scripts of languages that aren't written in Latin letters (a language not named here is Latin, or is judged
# by the script most of the piece is in)
LANG_SCRIPTS = {
    "chinese": {"Han"}, "mandarin": {"Han"}, "cantonese": {"Han"}, "zh": {"Han"},
    "japanese": {"Han", "Kana"}, "ja": {"Han", "Kana"}, "korean": {"Hangul", "Han"}, "ko": {"Hangul", "Han"},
    "russian": {"Cyrillic"}, "ru": {"Cyrillic"}, "ukrainian": {"Cyrillic"}, "uk": {"Cyrillic"}, "bulgarian": {"Cyrillic"},
    "serbian": {"Cyrillic"}, "belarusian": {"Cyrillic"}, "macedonian": {"Cyrillic"}, "kazakh": {"Cyrillic"},
    "kyrgyz": {"Cyrillic"}, "tajik": {"Cyrillic"}, "mongolian": {"Cyrillic", "Mongolian"},
    "arabic": {"Arabic"}, "ar": {"Arabic"}, "persian": {"Arabic"}, "farsi": {"Arabic"}, "fa": {"Arabic"}, "urdu": {"Arabic"},
    "pashto": {"Arabic"}, "dari": {"Arabic"}, "sindhi": {"Arabic"}, "uyghur": {"Arabic"},
    "hebrew": {"Hebrew"}, "he": {"Hebrew"}, "yiddish": {"Hebrew"},
    "hindi": {"Devanagari"}, "hi": {"Devanagari"}, "marathi": {"Devanagari"}, "nepali": {"Devanagari"},
    "sanskrit": {"Devanagari"}, "bengali": {"Bengali"}, "bangla": {"Bengali"}, "assamese": {"Bengali"},
    "punjabi": {"Gurmukhi", "Arabic"}, "gujarati": {"Gujarati"}, "odia": {"Oriya"}, "oriya": {"Oriya"}, "tamil": {"Tamil"},
    "telugu": {"Telugu"}, "kannada": {"Kannada"}, "malayalam": {"Malayalam"}, "sinhala": {"Sinhala"},
    "sinhalese": {"Sinhala"}, "thai": {"Thai"}, "lao": {"Lao"}, "tibetan": {"Tibetan"}, "burmese": {"Myanmar"},
    "myanmar": {"Myanmar"}, "georgian": {"Georgian"}, "amharic": {"Ethiopic"}, "tigrinya": {"Ethiopic"},
    "geez": {"Ethiopic"}, "khmer": {"Khmer"}, "armenian": {"Armenian"}, "greek": {"Greek"}, "el": {"Greek"},
    "cherokee": {"Cherokee"}, "nko": {"Nko"}, "tamazight": {"Tifinagh"}, "vai": {"Vai"}, "dhivehi": {"Thaana"},
    "syriac": {"Syriac"},
}
# languages written in Latin letters: for these the piece's own script is never judged from its letters
LATIN_LANGS = set("""english en french fr spanish es portuguese pt german de italian it dutch nl swahili sw kiswahili twi akan
fante ewe ga dagbani hausa ha yoruba yo igbo ig zulu zu xhosa xh afrikaans af somali so wolof oromo kinyarwanda kirundi
luganda shona sesotho setswana lingala malagasy polish pl czech slovak slovenian croatian bosnian romanian hungarian
finnish estonian latvian lithuanian swedish norwegian danish icelandic irish welsh basque catalan galician turkish tr
azerbaijani uzbek turkmen indonesian malay tagalog filipino vietnamese latin esperanto maltese albanian""".split())

_HOMO = {  # letters of another alphabet that look like Latin ones (inside a Latin word they are a slip, never meant)
    "а": "a", "е": "e", "о": "o", "р": "p", "с": "c", "у": "y", "х": "x", "і": "i", "ј": "j", "ѕ": "s", "ԁ": "d",
    "ԛ": "q", "ԝ": "w", "һ": "h", "ӏ": "l", "А": "A", "В": "B", "Е": "E", "К": "K", "М": "M", "Н": "H", "О": "O",
    "Р": "P", "С": "C", "Т": "T", "Х": "X", "І": "I", "Ј": "J", "Ѕ": "S", "Ү": "Y", "Ԛ": "Q", "Ԝ": "W",
    "ο": "o", "ν": "v", "ρ": "p", "ι": "i", "κ": "k", "υ": "u", "Α": "A", "Β": "B", "Ε": "E", "Ζ": "Z", "Η": "H",
    "Ι": "I", "Κ": "K", "Μ": "M", "Ν": "N", "Ο": "O", "Ρ": "P", "Τ": "T", "Υ": "Y", "Χ": "X",
}
_CJK = {"Han", "Kana", "Hangul"}
_WIDE_PUNCT = {"，": ", ", "。": ". ", "、": ", ", "：": ": ", "；": "; ", "！": "! ", "？": "? ", "（": " (", "）": ") ",
               "「": "“", "」": "”", "『": "“", "』": "”", "《": "“", "》": "”", "〈": "‘",
               "〉": "’", "【": "[", "】": "]", "　": " ", "～": "~", "・": " ", "．": ". ", "＂": "\""}
MARK_L, MARK_R = "⟦", "⟧"          # ⟦ ⟧ around a word that slipped in, in what the AI is sent


@functools.lru_cache(maxsize=8192)
def script_of(ch):
    """The script a letter belongs to ("Latin", "Han", "Cyrillic" …), None for anything that isn't a letter of one
    (digits, punctuation, spaces, symbols, marks of no script)."""
    o = ord(ch)
    if o < 0x80:
        return "Latin" if ch.isalpha() else None
    m = _SCRIPT_RX.match(ch)
    if m:
        return m.lastgroup
    if not ch.isalpha():
        return None
    n = unicodedata.name(ch, "")
    if "LATIN" in n or 0x250 <= o <= 0x2af or 0x1d00 <= o <= 0x1dbf:
        return "Latin"
    if n.startswith(("MODIFIER LETTER", "SUPERSCRIPT", "SUBSCRIPT")):
        return None
    return n.split(" ")[0].title() if n else None


def scripts_in(text, limit=60000):
    """How many letters of each script the text has."""
    out = {}
    for ch in str(text or "")[:limit]:
        if ch.isspace() or ch.isdigit():
            continue
        s = script_of(ch)
        if s:
            out[s] = out.get(s, 0) + 1
    return out


def lang_scripts(language):
    """The scripts a named language is written in, set() for a Latin one, None when the name isn't known."""
    s = re.sub(r"\([^)]*\)", " ", str(language or "")).strip().lower().replace("'", "").replace("’", "")
    if not s:
        return None
    for key in (s, s.split()[0] if s.split() else s, re.split(r"[-_]", s)[0]):
        if key in LANG_SCRIPTS:
            return set(LANG_SCRIPTS[key])
        if key in LATIN_LANGS:
            return set()
    for word in re.findall(r"[a-z]+", s):          # "Brazilian Portuguese", "Simplified Chinese", "Modern Standard Arabic"
        if word in LANG_SCRIPTS:
            return set(LANG_SCRIPTS[word])
        if word in LATIN_LANGS:
            return set()
    return None


def allowed_scripts(texts, language=None, context=""):
    """The scripts the pieces may use (see the module's note)."""
    allowed = {"Latin"}
    known = lang_scripts(language)
    if known is not None:
        allowed |= known
    else:
        counts = {}
        for t in texts:
            for k, v in scripts_in(t).items():
                counts[k] = counts.get(k, 0) + v
        total = sum(counts.values())
        if total:
            top = max(counts, key=counts.get)
            if counts[top] * 2 >= total:
                allowed.add(top)
    allowed |= set(scripts_in(context))
    if allowed & {"Han", "Kana"}:
        allowed |= {"Han", "Kana"} if "Kana" in allowed else set()
    if "Hangul" in allowed:
        allowed.add("Han")
    return allowed


def tidy(text, allowed):
    """What needs no AI: a look-alike letter inside a Latin word becomes the Latin letter, and full-width punctuation
    becomes ordinary punctuation in a piece that isn't Chinese, Japanese or Korean. Returns (text, how many changed)."""
    if not isinstance(text, str) or not text:
        return text, 0
    changed = 0

    def word(m):
        nonlocal changed
        w = m.group(0)
        sc = [script_of(c) for c in w]
        latin = sc.count("Latin")
        odd = [c for c, s in zip(w, sc) if s and s != "Latin" and s not in allowed]
        if not latin or not odd or any(c not in _HOMO for c in odd) or len(odd) > latin:
            return w
        changed += len(odd)
        return "".join(_HOMO[c] if c in odd else c for c in w)
    out = re.sub(r"[^\W\d_]+", word, text)
    if not allowed & _CJK and re.search("[　-〿！-～・]", out):
        def wide(m):
            nonlocal changed
            changed += 1
            c = m.group(0)
            if script_of(c):
                changed -= 1
                return c                  # a letter (々, 〇): a word of that script, read as one below
            if c in _WIDE_PUNCT:
                return _WIDE_PUNCT[c]
            return unicodedata.normalize("NFKC", c)
        out = re.sub("[　-〿！-～・]", wide, out)
        out = re.sub(r"[ \t]{2,}", " ", out)
        out = re.sub(r"[ \t]+(?=[,.;:!?)\]”’])", "", out)
        out = re.sub(r"(?<=[(\[“‘])[ \t]+", "", out)
        out = re.sub(r"(?m)[ \t]+$", "", out)
    return out, changed


def stray_runs(text, allowed):
    """Where words of a script the piece doesn't use stand: [(start, end)], a run taking in the spaces and marks
    between such letters ("管理 系统" is one run)."""
    runs, i, n = [], 0, len(text or "")
    while i < n:
        s = script_of(text[i])
        if not s or s in allowed:
            i += 1
            continue
        j = i + 1
        while j < n:
            c = text[j]
            sc = script_of(c)
            if sc and sc not in allowed:
                j += 1
                continue
            if sc is None and unicodedata.category(c).startswith("M"):
                j += 1
                continue
            if c in " ·-‐" and j + 1 < n and script_of(text[j + 1]) not in (None,) and script_of(text[j + 1]) not in allowed:
                j += 1
                continue
            break
        if s == "Greek" and set(script_of(c) for c in text[i:j] if script_of(c)) == {"Greek"} and j - i <= 2:
            i = j                         # α, β, χ, μ, Δ: symbols
            continue
        runs.append((i, j))
        i = j
    return runs


_SENT_END = re.compile(r"(?<=[.!?])[\"”’)]*\s+|\n+")


def _sentence_span(text, start, end):
    """The sentence a run stands in: (start, end) of it in the text."""
    a = 0
    for m in _SENT_END.finditer(text, 0, start):
        a = m.end()
    m = _SENT_END.search(text, end)
    b = m.start() if m else len(text)
    while b > a and text[b - 1] in " \t":
        b -= 1
    return a, b


def _numbers(s):
    return sorted(re.findall(r"\d+", s))


def _cites(s):
    return sorted(re.findall(r"\([^()]*\b(?:1[5-9]|20)\d\d[a-z]?\b[^()]*\)|\([^()]*\bn\.\s?d\.[^()]*\)|\[\d+(?:\s*[,–-]\s*\d+)*\]", s))


def _without(s, runs):
    """The sentence with its stray words taken out, tidied (what is left when nothing better can be put there)."""
    out, last = [], 0
    for a, b in runs:
        out.append(s[last:a])
        last = b
    out.append(s[last:])
    t = "".join(out)
    t = re.sub(r"\(\s*\)|\[\s*\]|“\s*”|\"\s*\"|‘\s*’", "", t)
    t = re.sub(r"[ \t]{2,}", " ", t)
    t = re.sub(r"[ \t]+(?=[,.;:!?)\]])", "", t)
    t = re.sub(r"(?<=[(\[])[ \t]+", "", t)
    t = re.sub(r",\s*,", ",", t)
    t = re.sub(r"^[ \t,;:]+", "", t)
    return t.strip() if s.strip() == s else t


def _take(orig, runs, fixed, allowed):
    """What the AI gave back for one sentence, if it may be used: ("fixed" | "kept", text) or None."""
    if not isinstance(fixed, str):
        return None
    fixed = fixed.replace(MARK_L, "").replace(MARK_R, "").strip()
    if not fixed:
        return None
    if fixed == orig.strip():
        return "kept", orig                # the sentence is about that word (an example of another language)
    if stray_runs(fixed, allowed) or "\n" in fixed:
        return None
    base = _without(orig, runs)
    if _numbers(fixed) != _numbers(base) or _cites(fixed) != _cites(orig):
        return None
    if len(fixed) > len(base) + 30 + 14 * len(runs) or len(fixed) < len(base) * 0.8:
        return None
    # every other word stays, in its order (a word next to it may change: "a" to "an"); a few words come in for each
    bw, fw = base.split(), fixed.split()
    kept = sum(b.size for b in difflib.SequenceMatcher(None, bw, fw, autojunk=False).get_matching_blocks())
    if kept < len(bw) - max(1, len(bw) // 12) or len(fw) - kept > 5 * len(runs) + 1:
        return None
    lead = re.match(r"\s*", orig).group(0)
    return "fixed", lead + fixed


def _ask_prompt(language):
    lang = (str(language).strip() if language else "") or "the language the rest of the sentence is written in"
    return (
        f"You correct slips in a finished piece of writing in {lang}. In each sentence below, the words between "
        f"{MARK_L} and {MARK_R} slipped in from another language or script by mistake. Replace each marked word with the "
        f"right word or words in {lang} for that place, so the sentence reads correctly and means what it was meant to "
        f"mean, and take the marks away. Change nothing else: every other word, citation, number and punctuation mark "
        f"stays exactly as it is. Only if a sentence is about the marked word itself (the word is quoted as an example of "
        f"that language, as in a study of the language), give the sentence back unchanged without the marks.\n"
        f"Answer with JSON only: a list like [{{\"i\": 0, \"s\": \"the corrected sentence\"}}], one item for each sentence.")


def fix_texts(texts, language=None, context="", ask=None, max_sentences=40):
    """Every text in one language. texts: a list of strings (anything else is passed through). ask(system, user,
    max_tokens) -> the AI's answer (it may raise); without it a stray word is taken out. Returns (texts, stats)."""
    texts = list(texts)
    stats = {"tidied": 0, "runs": 0, "sentences": 0, "fixed": 0, "kept": 0, "removed": 0, "ai": False, "ai_failed": False}
    live = [i for i, t in enumerate(texts) if isinstance(t, str) and t.strip()]
    if not live:
        return texts, stats
    allowed = allowed_scripts([texts[i] for i in live], language, context)
    jobs = []                                   # (text index, sentence start, sentence end, runs inside it, relative)
    for i in live:
        texts[i], n = tidy(texts[i], allowed)
        stats["tidied"] += n
        t = texts[i]
        runs = stray_runs(t, allowed)
        if not runs:
            continue
        stats["runs"] += len(runs)
        spans = []
        for a, b in runs:
            sa, sb = _sentence_span(t, a, b)
            if spans and sa < spans[-1][1]:
                spans[-1] = (spans[-1][0], max(spans[-1][1], sb), spans[-1][2] + [(a, b)])
            else:
                spans.append((sa, sb, [(a, b)]))
        for sa, sb, rs in spans:
            jobs.append((i, sa, sb, [(a - sa, b - sa) for a, b in rs]))
    if not jobs:
        return texts, stats
    stats["sentences"] = len(jobs)
    answers = {}
    asked = [k for k, (i, sa, sb, rs) in enumerate(jobs) if sb - sa <= 1500][:max_sentences]
    if ask and asked:
        items = []
        for k in asked:
            i, sa, sb, rs = jobs[k]
            s = texts[i][sa:sb]
            marked, last = [], 0
            for a, b in rs:
                marked.append(s[last:a] + MARK_L + s[a:b] + MARK_R)
                last = b
            marked.append(s[last:])
            items.append({"i": k, "s": "".join(marked)})
        stats["ai"] = True
        try:
            reply = ask(_ask_prompt(language), json.dumps(items, ensure_ascii=False),
                        min(4000, 300 + sum(len(x["s"]) for x in items) // 2))
            m = re.search(r"\[.*\]", reply or "", re.S)
            for x in json.loads(m.group(0)) if m else []:
                if isinstance(x, dict) and isinstance(x.get("i"), int):
                    answers[x["i"]] = x.get("s")
        except Exception:
            stats["ai_failed"] = True
    # put them back from the end of each text, so the places before stay right
    for k in sorted(range(len(jobs)), key=lambda k: (jobs[k][0], -jobs[k][1])):
        i, sa, sb, rs = jobs[k]
        t = texts[i]
        s = t[sa:sb]
        got = _take(s, rs, answers.get(k), allowed) if k in answers else None
        if got:
            stats[got[0]] += 1
            new = got[1]
        else:
            stats["removed"] += 1
            new = _without(s, rs)
        texts[i] = t[:sa] + new + t[sb:]
    return texts, stats


LIVE_SKIP = {"references", "papers", "sources", "searched", "thinking", "handoff", "url", "urls", "img", "image", "images",
             "query", "queries", "id", "style", "kind", "type", "layout", "template", "unsourced", "ref", "cite", "key",
             "akey", "doi", "pdf", "references_title", "style_name", "cite_marks", "parts", "outline2", "theme", "font",
             "presenters", "topic", "subtopic"}


def fix_tree(data, language=None, context="", ask=None):
    """Every string a person reads in an answer (not its references, sources or settings), in one language, with one
    request to the AI for all of them. Returns (data, stats)."""
    spots = []

    def walk(v, holder, key):
        if isinstance(v, dict):
            for k in list(v.keys()):
                if k not in LIVE_SKIP:
                    walk(v[k], v, k)
        elif isinstance(v, list):
            for j, x in enumerate(v):
                walk(x, v, j)
        elif isinstance(v, str) and holder is not None and len(v.strip()) > 1:
            spots.append((holder, key))
    walk(data, None, None)
    if not spots:
        return data, {"runs": 0}
    fixed, stats = fix_texts([h[k] for h, k in spots], language, context, ask)
    for (h, k), t in zip(spots, fixed):
        h[k] = t
    return data, stats


def context_text(obj, limit=60000):
    """The person's own words in a request (every string in it), for the scripts they use."""
    out, size = [], 0

    def walk(v):
        nonlocal size
        if size > limit:
            return
        if isinstance(v, dict):
            for x in v.values():
                walk(x)
        elif isinstance(v, list):
            for x in v:
                walk(x)
        elif isinstance(v, str) and not v.startswith("data:"):
            out.append(v[:20000])
            size += min(len(v), 20000)
    walk(obj)
    return " ".join(out)[:limit]
