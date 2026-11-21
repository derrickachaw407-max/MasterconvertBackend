"""6.6, round 8 (the owner, 21 Nov 2026 00:47: "Raise the app's validity. Do as much as a research tool can be."): the
statistics a piece of writing reports, worked out again.

A test's statistic and its degrees of freedom give its p-value: t, F, chi-square, r and z as APA writes them
("t(48) = 2.10, p = .041"). A p-value that is not what its statistic gives is said, and whether that changes the result's
significance at .05, as statcheck does (Nuijten, Hartgerink, van Assen, Epskamp & Wicherts, 2016, Behavior Research
Methods, 48, 1205-1226). Both figures are taken as rounded: a statistic written 2.10 is anything from 2.095 to 2.105, a
p written .041 anything from .0405 to .0415; only what no rounding explains is said. A test the writing calls one-tailed
may be one-tailed.

Also what no figure can be: a p-value of 0 or more than 1, a correlation past -1 or 1, R squared or Cronbach's alpha past 1,
a negative standard deviation; an estimate outside its own confidence interval, a ratio's interval below 0, a 95%
interval that disagrees with its own p-value (it holds the null value, 1 for a ratio and 0 for a difference, while p is
below .05, or the other way round, beyond what rounding explains); and a percentage that is not its count of its total
("123 of 412 (32.1%)").

Round 69: a sample size worked out again from its own formula (sample_check): Yamane's (Slovin's) and Cochran's (Kish's,
Fisher's), from the figures the writing gives them.

Pure Python (the server has no scipy): the regularized incomplete beta and gamma functions as Numerical Recipes (3rd ed.,
sections 6.2 and 6.4) computes them. Used by Check my citations (the person's writing) and by valid65.claim_problems (a
sentence RamzAI writes is not kept when its own figures disagree). Tests: tests/test_stats66.py (against scipy where the
workspace has it)."""
import logging
import math
import re

_log = logging.getLogger("academic-home")


def _fault(what, e):
    """6.6, round 128: a check that failed stays out of the person's way, and is said in the server's log by its kind and line
    alone (never its message: a message can quote the person's writing; the privacy notice says the logs never hold that).
    Before, these faults were passed over without a word, and a check could stop working unseen."""
    try:
        tb, line = getattr(e, "__traceback__", None), "?"
        while tb is not None:
            if tb.tb_frame.f_code.co_filename == __file__:
                line = tb.tb_lineno
            tb = tb.tb_next
        _log.error(f"Your statistics: the {what} check failed ({type(e).__name__} at stats66.py line {line})")
    except Exception:
        pass

_FPMIN = 1e-300


# ---------------------------------------------------------------- the distributions
def _betacf(a, b, x):
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c = 1.0
    d = 1.0 - qab * x / qap
    d = 1.0 / (d if abs(d) > _FPMIN else _FPMIN)
    h = d
    for m in range(1, 500):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d
        d = 1.0 / (d if abs(d) > _FPMIN else _FPMIN)
        c = 1.0 + aa / c
        c = c if abs(c) > _FPMIN else _FPMIN
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d
        d = 1.0 / (d if abs(d) > _FPMIN else _FPMIN)
        c = 1.0 + aa / c
        c = c if abs(c) > _FPMIN else _FPMIN
        de = d * c
        h *= de
        if abs(de - 1.0) < 1e-15:
            break
    return h


def ibeta(a, b, x):
    """The regularized incomplete beta function I_x(a, b)."""
    if x <= 0.0:
        return 0.0
    if x >= 1.0:
        return 1.0
    lbt = math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b) + a * math.log(x) + b * math.log1p(-x)
    if x < (a + 1.0) / (a + b + 2.0):
        return math.exp(lbt) * _betacf(a, b, x) / a
    return 1.0 - math.exp(lbt) * _betacf(b, a, 1.0 - x) / b


def gammq(a, x):
    """The regularized upper incomplete gamma function Q(a, x)."""
    if x <= 0.0:
        return 1.0
    gln = math.lgamma(a)
    if x < a + 1.0:
        ap, s = a, 1.0 / a
        d = s
        for _ in range(2000):
            ap += 1.0
            d *= x / ap
            s += d
            if abs(d) < abs(s) * 1e-16:
                break
        return max(0.0, 1.0 - s * math.exp(-x + a * math.log(x) - gln))
    b = x + 1.0 - a
    c = 1.0 / _FPMIN
    d = 1.0 / b
    h = d
    for i in range(1, 2000):
        an = -i * (i - a)
        b += 2.0
        d = an * d + b
        d = 1.0 / (d if abs(d) > _FPMIN else _FPMIN)
        c = b + an / c
        c = c if abs(c) > _FPMIN else _FPMIN
        de = d * c
        h *= de
        if abs(de - 1.0) < 1e-16:
            break
    return math.exp(-x + a * math.log(x) - gln) * h


def p_t(t, df):
    """Two-tailed p of Student's t."""
    t = abs(float(t))
    return ibeta(df / 2.0, 0.5, df / (df + t * t)) if df > 0 else None


def p_f(f, d1, d2):
    """Upper-tail p of F(d1, d2)."""
    f = float(f)
    if d1 <= 0 or d2 <= 0:
        return None
    return 1.0 if f <= 0 else ibeta(d2 / 2.0, d1 / 2.0, d2 / (d2 + d1 * f))


def p_chi2(x, df):
    """Upper-tail p of chi-square."""
    x = float(x)
    if df <= 0:
        return None
    return 1.0 if x <= 0 else gammq(df / 2.0, x / 2.0)


def p_z(z):
    """Two-tailed p of a standard normal z."""
    return math.erfc(abs(float(z)) / math.sqrt(2.0))


def p_r(r, df):
    """Two-tailed p of Pearson's r with df = n - 2 (through t)."""
    r = abs(float(r))
    if df <= 0 or r > 1:
        return None
    if r >= 1:
        return 0.0
    return p_t(r * math.sqrt(df / (1.0 - r * r)), df)


# ---------------------------------------------------------------- reading the writing
_NUM = r"[-−–]?\s?(?:\d+(?:\.\d+)?|\.\d+)"
_DF = r"\d+(?:\.\d+)?"
_CHI = r"(?:χ\s*(?:²|\^?2)|(?<![A-Za-z0-9])[Xx]\s*(?:²|\^?2)|(?<![A-Za-z])chi[-\s]?squared?(?:\s+test)?)"
_STATS = [
    ("t", re.compile(r"(?<![A-Za-z])t\s*\(\s*(?P<df>" + _DF + r")\s*\)\s*=\s*(?P<v>" + _NUM + ")")),
    ("t", re.compile(r"(?<![A-Za-z])t\s*=\s*(?P<v>" + _NUM + r")\s*[,;]\s*df\s*=\s*(?P<df>" + _DF + ")")),
    ("t", re.compile(r"(?<![A-Za-z])df\s*=\s*(?P<df>" + _DF + r")\s*[,;]\s*t\s*=\s*(?P<v>" + _NUM + ")")),
    ("F", re.compile(r"(?<![A-Za-z])F\s*\(\s*(?P<d1>" + _DF + r")\s*,\s*(?P<d2>" + _DF + r")\s*\)\s*=\s*(?P<v>" + _NUM + ")")),
    ("chi2", re.compile(_CHI + r"\s*\(\s*(?P<df>\d+)\s*(?:,\s*[Nn]\s*=\s*[\d,]+\s*)?\)\s*=\s*(?P<v>" + _NUM + ")")),
    ("chi2", re.compile(_CHI + r"\s*=\s*(?P<v>" + _NUM + r")\s*[,;]\s*df\s*=\s*(?P<df>\d+)")),
    ("r", re.compile(r"(?<![A-Za-z])r\s*\(\s*(?P<df>\d+)\s*\)\s*=\s*(?P<v>" + _NUM + ")")),
    ("z", re.compile(r"(?<![A-Za-z])[zZ]\s*=\s*(?P<v>" + _NUM + ")")),
    # 6.6, round 85: the other ways people write the degrees of freedom: "d.f. = 48", "with 2 degrees of freedom",
    # "F = 4.21, df = 2, 97" (or "df = (2, 97)"), "a chi-square value of 12.4"
    ("t", re.compile(r"(?<![A-Za-z])t\s*(?:-?\s*value\s+)?(?:=|of)\s*(?P<v>" + _NUM + r")\s*[,;]?\s*(?:with\s+)?(?:d\.\s?f\.|df)\s*(?:=|of)\s*(?P<df>" + _DF + ")")),
    ("t", re.compile(r"(?<![A-Za-z])t\s*(?:-?\s*value\s+)?(?:=|of)\s*(?P<v>" + _NUM + r")\s*[,;]?\s*with\s+(?P<df>" + _DF + r")\s+degrees?\s+of\s+freedom")),
    ("F", re.compile(r"(?<![A-Za-z])F\s*(?:-?\s*value\s+)?(?:=|of)\s*(?P<v>" + _NUM + r")\s*[,;]?\s*(?:with\s+)?(?:d\.\s?f\.|df)\s*(?:=|of)?\s*\(?\s*(?P<d1>"
                     + _DF + r")\s*(?:,|and)\s*(?P<d2>" + _DF + r")\s*\)?")),
    ("chi2", re.compile(_CHI + r"\s*(?:-?\s*value\s+)?(?:=|of)\s*(?P<v>" + _NUM + r")\s*[,;]?\s*(?:with\s+)?(?:d\.\s?f\.|df)\s*(?:=|of)\s*(?P<df>\d+)", re.I)),
    ("chi2", re.compile(_CHI + r"\s*(?:-?\s*value\s+)?(?:=|of)\s*(?P<v>" + _NUM + r")\s*[,;]?\s*with\s+(?P<df>\d+)\s+degrees?\s+of\s+freedom", re.I)),
]
_R_ALONE = re.compile(r"(?<![A-Za-z])r\s*=\s*(?P<v>" + _NUM + ")")
_N_OF = re.compile(r"(?<![A-Za-z])[nN]\s*=\s*(?P<n>\d[\d,]*)")
_P = re.compile(r"(?<![A-Za-z])[pP]s?\s*(?:-?\s*values?\s*)?(?P<rel><=|>=|≤|≥|=|<|>)\s*(?P<p>\d*\.?\d+)(?![\d.]*\s*%)")
_NS = re.compile(r"(?<![A-Za-z])(?:n\.\s?s\.?|ns|NS)(?![A-Za-z])")
_OTHER = re.compile(r"(?<![A-Za-z])(?:U|W|H|T|Q|K|V)\s*(?:\(\s*[\d.,\s]+\))?\s*=\s*[-−–]?\s?\d")   # a test this check can't work out
_ONE_TAIL = re.compile(r"\bone[-\s]?(?:tailed|sided)\b|\bdirectional\b", re.I)


def _num(s):
    s = re.sub(r"\s+", "", str(s)).replace("−", "-").replace("–", "-")
    return float(s)


def _bare(s):
    """A figure as written, without its spaces, its minus sign as one ("− 2.10" -> "−2.10")."""
    return re.sub(r"\s+", "", str(s or "")).replace("–", "−").replace("-", "−")


def _dec(s):
    s = re.sub(r"\s+", "", str(s))
    return len(s.split(".", 1)[1]) if "." in s else 0


def sentences(text):
    """The writing's sentences (a paragraph never runs into the next)."""
    out = []
    for para in re.split(r"\n\s*\n|\n", str(text or "")):
        para = para.strip()
        if not para:
            continue
        out += [s.strip() for s in re.split(r"(?<=[.!?])\s+(?=[A-Z“\"])", para) if s.strip()]
    return out


def _clean(s):
    """Markdown's italics off ("*t*(48)"), one kind of space."""
    return re.sub(r"[ \t   ]+", " ", re.sub(r"[*_]", "", str(s or "")))


def _fmt_p(p, lead):
    if p < 0.001:
        return "p < 0.001" if lead else "p < .001"
    s = f"{p:.3f}"
    return "p = " + (s if lead else s[1:] if s.startswith("0") else s)


def _stat_label(kind, m, g):
    v = _bare(g["v"])
    if kind == "r" and "df" not in m.groupdict():
        return f"r = {v} (n = {int(g['df']) + 2})"
    if kind == "t":
        return f"t({g['df']}) = {v}"
    if kind == "F":
        return f"F({g['d1']}, {g['d2']}) = {v}"
    if kind == "chi2":
        return f"χ²({g['df']}) = {v}"
    if kind == "r":
        return f"r({g['df']}) = {v}"
    return f"z = {v}"


def _p_of(kind, v, g, one=False):
    """The p a statistic gives, or None when it can't give one."""
    try:
        if kind == "t":
            p = p_t(v, float(g["df"]))
        elif kind == "F":
            p = p_f(v, float(g["d1"]), float(g["d2"]))
        elif kind == "chi2":
            p = p_chi2(v, float(g["df"]))
        elif kind == "r":
            p = p_r(v, float(g["df"]))
        else:
            p = p_z(v)
    except (ValueError, ZeroDivisionError, OverflowError):
        return None
    if p is None:
        return None
    return p / 2.0 if one and kind in ("t", "r", "z") else p


def _fits(rel, rep, dp, plo, phi):
    """Whether a p written as `rel rep` (dp decimals) can be the p a statistic gives (plo..phi, its rounding)."""
    if rel == "=":
        h = 0.5 * 10 ** -dp
        return plo <= rep + h + 1e-12 and phi >= rep - h - 1e-12
    if rel in ("<", "≤", "<="):
        return plo < rep + (1e-12 if rel != "<" else 0.0)
    return phi > rep - (1e-12 if rel != ">" else 0.0)


def _tests(s):
    """The tests one sentence reports: [(kind, the match, its p's match or the "ns" match or None)]."""
    found = []
    for kind, rx in _STATS:
        for m in rx.finditer(s):
            if any(a <= m.start() < b for _k, mm, _p in found for a, b in [(mm.start(), mm.end())]):
                continue
            found.append((kind, m, None))
    rs = [m for m in _R_ALONE.finditer(s) if not any(mm.start() <= m.start() < mm.end() for _k, mm, _p in found)]
    for m in rs:
        found.append(("r?", m, None))
    found.sort(key=lambda x: x[1].start())
    out = []
    for i, (kind, m, _x) in enumerate(found):
        stop = found[i + 1][1].start() if i + 1 < len(found) else len(s)
        win = s[m.end():min(stop, m.end() + 90)]
        other = _OTHER.search(win)
        if other:
            win = win[:other.start()]          # (its p is that test's)
        pm = _P.search(win)
        if not pm:
            ns = _NS.search(win[:24])
            out.append((kind, m, ns, True) if ns else (kind, m, None, False))
            continue
        out.append((kind, m, pm, False))
    return out


def _test_notes(s, counted):
    notes = []
    one = bool(_ONE_TAIL.search(s))
    for kind, m, pm, ns in _tests(s):
        if pm is None:
            continue
        g = m.groupdict()
        if kind == "r?":                     # "r = .23, n = 412": the degrees of freedom from the sample
            ns_ = [int(x.group("n").replace(",", "")) for x in _N_OF.finditer(s)]
            if len(set(ns_)) != 1 or ns_[0] < 4:
                continue
            kind, g = "r", {"df": str(ns_[0] - 2), "v": g["v"]}
        try:
            v = _num(g["v"])
        except ValueError:
            continue
        if kind == "r" and abs(v) >= 1:
            continue                            # (said by the ranges below)
        if kind in ("F", "chi2") and v < 0:
            continue
        h = 0.5 * 10 ** -_dec(g["v"])
        lo_v, hi_v = max(0.0, abs(v) - h), abs(v) + h
        if kind == "r":
            hi_v = min(hi_v, 0.999999)
        pv = _p_of(kind, abs(v), g, one=False)
        plo, phi = _p_of(kind, hi_v, g), _p_of(kind, lo_v, g)
        if pv is None or plo is None or phi is None:
            continue
        counted[0] += 1
        label = _stat_label(kind, m, g)
        if ns:
            rel, rep, dp, written = ">", 0.05, 2, "“ns” (p > .05)"
        else:
            rel = pm.group("rel")
            try:
                rep = float(pm.group("p"))
            except ValueError:
                continue
            if rep > 1 or rep == 0:
                continue                        # (said by the ranges below)
            dp = _dec(pm.group("p"))
            written = "p " + rel + " " + pm.group("p")
        ok = _fits(rel, rep, dp, plo, phi)
        if not ok and one and kind in ("t", "r", "z"):
            ok = _fits(rel, rep, dp, plo / 2.0, phi / 2.0)
        if ok:
            continue
        lead = (not ns) and pm.group("p").startswith("0")
        tails = " (two-tailed)" if kind in ("t", "r", "z") else ""
        note = f"{label} gives {_fmt_p(pv, lead)}{tails}, not {written}."
        sig_said = (not ns) and ((rel in ("<", "≤", "<=") and rep <= 0.05) or (rel == "=" and rep < 0.05))
        nonsig_said = ns or (rel in (">", "≥", ">=") and rep >= 0.05) or (rel == "=" and rep > 0.05)
        decision = (sig_said and plo >= 0.05) or (nonsig_said and phi < 0.05)
        if decision:
            note += " By its own figures, the result is " + ("not significant at .05." if sig_said else "significant at .05.")
        notes.append({"note": note, "kind": "test", "decision": decision})
    return notes


# ---------------------------------------------------------------- figures no study can have
_P_ANY = re.compile(r"(?<![A-Za-z])[pP]s?\s*(?:-?\s*values?\s*)?(?:=|<|≤)\s*(?P<p>\d*\.?\d+)(?![\d.]*\s*%)")
_R_ANY = re.compile(r"(?<![A-Za-z])r\s*(?:\(\s*\d+\s*\))?\s*=\s*(?P<v>" + _NUM + ")")
_R2 = re.compile(r"(?<![A-Za-z])(?:R\s*(?:²|\^?2)|R[-\s]squared|adjusted\s+R\s*(?:²|\^?2))\s*=\s*(?P<v>" + _NUM + ")")
_ALPHA = re.compile(r"Cronbach['’]?s?\s+(?:alpha|α|coefficient\s+alpha)(?:\s+(?:coefficient|value|score))?\s*(?:of|=|was|is|:)?\s*(?:=\s*)?(?P<v>" + _NUM + ")", re.I)
_SD = re.compile(r"(?<![A-Za-z])(?:SD|s\.d\.|standard\s+deviation)\s*(?:of|=|:)?\s*(?:=\s*)?(?P<v>[-−–]\s?\d*\.?\d+)")


def _range_notes(s):
    notes = []
    for m in _P_ANY.finditer(s):
        try:
            p = float(m.group("p"))
        except ValueError:
            continue
        if p == 0 and "<" not in m.group(0) and "≤" not in m.group(0):
            notes.append({"note": "A p-value is never 0: write p < .001.", "kind": "value", "decision": False})
        elif p == 0:                                  # 6.6, round 162: "p < .000"
            notes.append({"note": "No p-value is below 0: SPSS’s “.000” means p < .001, so write p < .001.", "kind": "value", "decision": False})
        elif p > 1:
            notes.append({"note": f"A p-value can’t be more than 1 (it is written {m.group('p')}).", "kind": "value", "decision": False})
    for m in _R_ANY.finditer(s):
        try:
            if abs(_num(m.group("v"))) > 1:
                notes.append({"note": f"A correlation lies between −1 and 1 (r is written {_bare(m.group('v'))}).", "kind": "value", "decision": False})
        except ValueError:
            pass
    for m in _R2.finditer(s):
        try:
            v = _num(m.group("v"))
            if v > 1 or v < 0:
                notes.append({"note": f"R² lies between 0 and 1 (it is written {_bare(m.group('v'))}).", "kind": "value", "decision": False})
        except ValueError:
            pass
    for m in _ALPHA.finditer(s):
        try:
            if _num(m.group("v")) > 1:
                notes.append({"note": f"Cronbach’s alpha can’t be more than 1 (it is written {_bare(m.group('v'))}).", "kind": "value", "decision": False})
        except ValueError:
            pass
    for m in _SD.finditer(s):
        notes.append({"note": "A standard deviation can’t be negative.", "kind": "value", "decision": False})
    return notes


# ---------------------------------------------------------------- an estimate and its confidence interval
_CI = re.compile(r"(?P<lvl>\d{2}(?:\.\d)?)\s*%\s*(?:CI|C\.I\.|CrI|confidence\s+intervals?)\s*(?:[:=,]\s*|\s+)?(?:of\s+|from\s+)?"
                 r"[\[(]?\s*(?P<lo>[-−]?\s?(?:\d+(?:\.\d+)?|\.\d+))\s*(?P<sep>,|;|\bto\b|[–—]|-)\s*(?P<hi>[-−]?\s?(?:\d+(?:\.\d+)?|\.\d+))\s*[\])]?")
_RATIO = (r"(?:(?:[aA]|adj\.?\s?|[Aa]djusted\s+|[Cc]rude\s+)?(?:OR|RR|HR|IRR|PR|POR|SIR|SMR)|"
          r"(?:[Aa]djusted\s+|[Cc]rude\s+)?(?:odds|risk|hazard|prevalence|rate|incidence\s+rate)\s+ratios?|relative\s+risks?)")
_DIFF = (r"(?:B|b|β|[Bb]eta|MD|SMD|[Mm]ean\s+difference|[Dd]ifference|Cohen['’]s\s+d|Hedges['’]?\s+g|d|g|Δ)")
_EST = re.compile(r"(?:(?<![A-Za-z])(?P<lab>" + _RATIO + "|" + _DIFF + r"|M|[Mm]ean|[Mm]edian|ARR|AR|r|ρ)\s*(?:[:=]\s*|\s+))?"
                  r"(?P<est>[-−]?\s?(?:\d+\.\d+|\.\d+|\d+))\s*(?:\(|,|;|:)?\s*$")
_NOT_EST = re.compile(r"(?:(?<![A-Za-z])[pP]s?|(?<![A-Za-z])t\s*\(\s*[\d.]+\s*\)|(?<![A-Za-z])F\s*\([^)]*\)|" + _CHI + r"\s*\([^)]*\)|"
                      r"(?<![A-Za-z])r\s*\(\s*\d+\s*\)|(?<![A-Za-z])[zZ]|(?<![A-Za-z])df|SE|SD|(?<![A-Za-z])[nN]|I\s*(?:²|\^?2))\s*[:=<>≤≥]\s*$")


_SIG_NOT = re.compile(r"\b(?:not|no)\s+(?:\w+\s+){0,2}?significan\w*|\bnon[-\s]?significan\w*|\bin-?significan\w*|"   # (6.6, round 161: "Income was
                      r"\bfailed\s+to\s+reach\s+significance", re.I)                                                        # significantly" is not "in…significant")
_SIG_YES = re.compile(r"(?<!clinically\s)(?<!practically\s)\bsignificant(?:ly)?\b(?!\s+(?:level|at\s+the))|"
                      r"\b(?:reached|achieved|attained|showed)\s+(?:statistical\s+)?significance\b", re.I)
_CLAUSE = re.compile(r",\s*(?:but|while|whereas|although|yet)\b|;|\bwhereas\b|\bwhile\b|\bbut\b", re.I)


def _sig_said(before):
    """What a result's own clause says of its significance: "sig", "nonsig" or "" (6.6, round 78)."""
    cut = list(_CLAUSE.finditer(before))
    t = before[cut[-1].end():] if cut else before
    if _SIG_NOT.search(t):
        return "nonsig"
    return "sig" if _SIG_YES.search(t) else ""


def _ci_notes(s):
    notes = []
    prev_end = 0
    for m in _CI.finditer(s):
        clause_from, prev_end = prev_end, m.end()             # (the clause of this interval runs from the last one's end)
        try:
            lo, hi = _num(m.group("lo")), _num(m.group("hi"))
        except ValueError:
            continue
        before = s[max(0, m.start() - 70):m.start()]
        em = _EST.search(before)
        est = lab = None
        if em and not _NOT_EST.search(before[:em.start("est")]) and (em.group("lab") or "." in em.group("est")):
            if not re.fullmatch(r"(?:19|20)\d{2}", em.group("est").strip()):
                try:
                    est = _num(em.group("est"))
                except ValueError:
                    est = None
                lab = em.group("lab")
        lvl = m.group("lvl")
        lo_s, hi_s = _bare(m.group("lo")), _bare(m.group("hi"))
        shown = f"{lo_s}–{hi_s}" if lo >= 0 else f"{lo_s} to {hi_s}"
        if lo > hi:
            notes.append({"note": f"The {lvl}% CI’s lower end ({lo_s}) is above its upper end ({hi_s}).", "kind": "ci", "decision": False})
            continue
        ratio = bool(lab and re.fullmatch(_RATIO, lab))
        diff = bool(lab and re.fullmatch(_DIFF, lab))
        if ratio and lo < 0:
            notes.append({"note": f"A ratio’s confidence interval can’t go below 0 (it is written {shown}).", "kind": "ci", "decision": False})
            continue
        if est is not None:
            tol = 0.5 * 10 ** -min(_dec(em.group("est")), _dec(m.group("lo")), _dec(m.group("hi")))
            if est < lo - tol or est > hi + tol:
                notes.append({"note": f"The estimate ({_bare(em.group('est'))}) lies outside its own {lvl}% CI ({shown}).", "kind": "ci", "decision": False})
                continue
        if not (ratio or diff) or abs(float(lvl) - 95.0) > 0.01:
            continue
        null = 1.0 if ratio else 0.0
        u = 10 ** -max(_dec(m.group("lo")), _dec(m.group("hi")))
        holds = lo < null - u and hi > null + u
        leaves = lo > null + u or hi < null - u
        win = s[m.end():m.end() + 60]
        pm = _P.search(win)
        if not pm:
            # 6.6, round 78: no p beside it: the words of its own clause. "Significantly associated (OR 1.8, 95% CI 0.9–3.6)"
            # says what its interval doesn't (it holds 1); "no significant difference (MD 2.1, 95% CI 0.4–3.8)" the other way
            said = _sig_said(s[clause_from:m.start()])
            what = "1" if ratio else "0"
            h = 0.5 * u                                    # (a lower end written 0.9 is 0.85 to 0.95: it holds 1)
            holds, leaves = lo < null - h and hi > null + h, lo > null + h or hi < null - h
            if said == "sig" and holds:
                notes.append({"note": f"Its 95% CI ({shown}) holds {what}, so it isn’t significant at .05, as the sentence says it is.",
                              "kind": "ci", "decision": True})
            elif said == "nonsig" and leaves:
                notes.append({"note": f"Its 95% CI ({shown}) leaves out {what}, so it is significant at .05, though the sentence says it isn’t.",
                              "kind": "ci", "decision": True})
            continue
        try:
            rep = float(pm.group("p"))
        except ValueError:
            continue
        rel = pm.group("rel")
        sig = (rel in ("<", "≤", "<=") and rep <= 0.05) or (rel == "=" and rep < 0.045)
        nonsig = (rel in (">", "≥", ">=") and rep >= 0.05) or (rel == "=" and rep > 0.055)
        what = "1" if ratio else "0"
        if holds and sig:
            notes.append({"note": f"Its 95% CI ({shown}) holds {what}, so its p can’t be below .05: one of the two is wrong.", "kind": "ci", "decision": True})
        elif leaves and nonsig:
            notes.append({"note": f"Its 95% CI ({shown}) leaves out {what}, so its p should be below .05: one of the two is wrong.", "kind": "ci", "decision": True})
    return notes


# ---------------------------------------------------------------- a count and its percentage
_PCTS = [
    re.compile(r"(?<![\d.,])(?P<x>\d[\d,]*)\s+(?:out\s+)?of\s+(?:the\s+|all\s+|a\s+total\s+of\s+)?(?P<y>\d[\d,]*)(?![\d.,])(?:\s+[A-Za-z][\w'’\-]*){0,4}\s*\(\s*(?P<z>\d+(?:\.\d+)?)\s*%\s*\)"),
    re.compile(r"(?<![\d.,])(?P<x>\d[\d,]*)\s*/\s*(?P<y>\d[\d,]*)(?![\d.,])\s*\(\s*(?P<z>\d+(?:\.\d+)?)\s*%\s*\)"),
    re.compile(r"\b[Oo]f\s+(?:the\s+|all\s+|a\s+total\s+of\s+)?(?P<y>\d[\d,]*)(?![\d.,])(?:\s+[A-Za-z][\w'’\-]*){0,6}\s*,\s*(?P<x>\d[\d,]*)(?![\d.,])\s*\(\s*(?P<z>\d+(?:\.\d+)?)\s*%\s*\)"),
]


def _pct_notes(s):
    notes, seen = [], set()
    for rx in _PCTS:
        for m in rx.finditer(s):
            try:
                x, y, z = int(m.group("x").replace(",", "")), int(m.group("y").replace(",", "")), float(m.group("z"))
            except ValueError:
                continue
            if y <= 0 or x > y or (x, y, z) in seen or 1900 <= x <= 2100 or 1900 <= y <= 2100:
                continue                        # ("of the 2019 cohort, 120 (30%)": a year, not a total)
            seen.add((x, y, z))
            real = 100.0 * x / y
            tol = 0.5 * 10 ** -_dec(m.group("z")) + 0.05
            if abs(real - z) > tol:
                d = max(1, _dec(m.group("z")))
                notes.append({"note": f"{x:,} of {y:,} is {real:.{d}f}%, not {m.group('z')}%.", "kind": "percent", "decision": False})
    return notes


# ---- 6.6, round 38: one group split in two ("38% were male and 72% female") can't make more than 100%
_SPLIT = re.compile(r"(?P<p>\d{1,3}(?:\.\d+)?)\s?%\s*(?:\)\s*)?(?:were\s+|was\s+|are\s+|is\s+)?(?P<w>males?|females?|men|women|boys|girls)\b(?!\s+(?:who|that|with|reported|had|were))",
                    re.I)
_SPLIT2 = re.compile(r"\b(?P<w>males?|females?|men|women|boys|girls)\s*\(\s*(?:n\s*=\s*[\d,]+\s*[,;]\s*)?(?P<p>\d{1,3}(?:\.\d+)?)\s?%\s*\)", re.I)
_SPLIT_SIDE = {"male": "m", "males": "m", "men": "m", "boys": "m", "female": "f", "females": "f", "women": "f", "girls": "f"}


def _split_notes(s):
    """"Of the 412 students, 38% were male and 72% female": the two parts of one group make more than 100% (fewer can be
    right: some answered neither, or were neither). Only where each part is said once, as the group's make-up ("38% were
    male"), never "38% of males reported…"."""
    sides, spans = {}, []
    for rx in (_SPLIT2, _SPLIT):
        for m in rx.finditer(s):
            if any(a <= m.start() < b or a < m.end() <= b for a, b in spans):
                continue
            if rx is _SPLIT and re.search(r"\bof\s*$", s[max(0, m.start() - 4):m.start()]):
                continue
            spans.append((m.start(), m.end()))
            sides.setdefault(_SPLIT_SIDE[m.group("w").lower()], []).append(m)
    if len(sides.get("m", [])) != 1 or len(sides.get("f", [])) != 1:
        return []
    a, b = sides["m"][0], sides["f"][0]
    try:
        x, y = float(a.group("p")), float(b.group("p"))
    except ValueError:
        return []
    tol = 0.5 * 10 ** -_dec(a.group("p")) + 0.5 * 10 ** -_dec(b.group("p"))
    if x + y <= 100 + tol + 1e-9:
        return []
    return [{"note": f"{a.group('p')}% {a.group('w').lower()} and {b.group('p')}% {b.group('w').lower()} make "
                     f"{x + y:g}%: the two parts of one group can’t make more than 100%.", "kind": "percent", "decision": False}]


# ---- 6.6, round 52: a change or a comparison said the other way from its own figures
# "rose from 30% to 25%", "fell from 20 to 35"; "Ghana's rate (25%) was higher than Nigeria's (38%)", "higher in Ghana (25%) than
# in Nigeria (38%)": the words and the figures disagree. Only in these tight shapes (never "38% higher than", a difference;
# never "more than 40%", a bound; never "38% of boys reported higher scores than girls (45%)", other things compared).
_UP = r"rose|risen|rising|increased|increasing|increases|grew|grown|growing|climbed|climbing|went\s+up|gone\s+up|jumped|doubled|tripled"
_DOWN = r"fell|fallen|falling|decreased|decreasing|decreases|declined|declining|dropped|dropping|went\s+down|gone\s+down|halved|reduced|reducing"
_ABOUT = r"(?:about\s+|around\s+|nearly\s+|almost\s+|roughly\s+|some\s+)?"
_FROMTO = re.compile(r"(?P<neg>\b(?:not|never|no)\s+(?:\w+\s+){0,2})?\b(?P<v>" + _UP + "|" + _DOWN + r")\b(?P<mid>[^.;:!?\d%]{0,40}?)\bfrom\s+"
                     + _ABOUT + r"(?P<x>\d+(?:\.\d+)?)\s?(?P<ux>%|percent\b|per\s+cent\b)?(?:\s+in\s+(?:19|20)\d{2})?\s*(?:\([^)]{0,40}\)\s*)?\s*to\s+"
                     + _ABOUT + r"(?P<y>\d+(?:\.\d+)?)\s?(?P<uy>%|percent\b|per\s+cent\b)?", re.I)
_HI, _LO = {"higher", "greater", "larger"}, {"lower", "smaller"}
_CMP_A = re.compile(r"(?P<x>\d+(?:\.\d+)?)\s?%\s*(?P<mid>\)?\s*,?\s*(?:(?:was|were|is|are)\s+)?(?:(?:much|far|slightly|considerably|substantially|"
                    r"significantly|markedly)\s+)?)(?P<c>higher|greater|larger|lower|smaller)\s+than\s+(?P<obj>[^%.;:\d()]{0,40}?\(?\s*|the\s+)"
                    r"(?P<y>\d+(?:\.\d+)?)\s?%", re.I)
_CMP_B = re.compile(r"\b(?P<c>higher|greater|larger|lower|smaller)\s+(?:in|among|for|at|within)\s+[^%.;:\d()]{1,40}\(\s*(?P<x>\d+(?:\.\d+)?)\s?%\s*\)\s*"
                    r"than\s+(?:in|among|for|at|within)\s+[^%.;:\d()]{1,40}\(\s*(?P<y>\d+(?:\.\d+)?)\s?%\s*\)", re.I)


def _order_notes(s):
    notes = []
    for m in _FROMTO.finditer(s):
        if m.group("neg"):
            continue
        x, y = float(m.group("x")), float(m.group("y"))
        if bool(m.group("ux")) != bool(m.group("uy")) and m.group("uy"):
            continue
        if not m.group("ux") and not m.group("uy") and (1900 <= x <= 2100 or 1900 <= y <= 2100):
            continue                                # (years: "from 2015 to 2020")
        up = re.fullmatch(_UP, m.group("v"), re.I) is not None
        u = "%" if (m.group("ux") or m.group("uy")) else ""
        if up and y <= x:
            notes.append({"note": f"“{m.group('v')} from {m.group('x')}{u} to {m.group('y')}{u}”: {m.group('y')}{u} is "
                                  f"{'the same as' if y == x else 'lower than'} {m.group('x')}{u}, so the figures don’t show a rise.",
                          "kind": "order", "decision": False})
        elif not up and y >= x:
            notes.append({"note": f"“{m.group('v')} from {m.group('x')}{u} to {m.group('y')}{u}”: {m.group('y')}{u} is "
                                  f"{'the same as' if y == x else 'higher than'} {m.group('x')}{u}, so the figures don’t show a fall.",
                          "kind": "order", "decision": False})
    for rx in (_CMP_A, _CMP_B):
        for m in rx.finditer(s):
            if rx is _CMP_A and not m.group("mid").strip(" ,"):
                continue                            # ("38% higher than …": a difference, not a level)
            before = s[max(0, m.start() - 30):m.start()].lower()
            if re.search(r"\b(?:not|no|never|compared|relative|versus|vs)\b", before + " " + (m.group("obj") if rx is _CMP_A else "")):
                continue
            x, y, c = float(m.group("x")), float(m.group("y")), m.group("c").lower()
            if (c in _HI and x <= y) or (c in _LO and x >= y):
                notes.append({"note": f"{m.group('x')}% is {'the same as' if x == y else ('lower' if x < y else 'higher')} than "
                                      f"{m.group('y')}%, not {c}: the sentence and its figures disagree.", "kind": "order", "decision": False})
    return notes


# ---------------------------------------------------------------- the whole check
# ---------------------------------------------------------------- 6.6, round 37: a coefficient with its standard error
_COEF = re.compile(r"(?<![A-Za-z])(?P<lab>B|b|β|[Bb]eta|[Cc]oefficient|[Ee]stimate)\s*=\s*(?P<v>" + _NUM + r")\s*"
                   r"(?:[,;]\s*|\(\s*)(?:SE|S\.E\.|[Ss]tandard\s+error)\s*=\s*(?P<se>\d*\.?\d+)\s*\)?")


def _coef_notes(s, counted):
    """"b = 0.45, SE = 0.40, p = .02": the estimate over its standard error is a z (or a t, whose p is larger still with
    few degrees of freedom); whatever the sample, the p can't be smaller than the z gives. Said only that way round: a p
    written smaller (more significant) than the figures can give at best."""
    notes = []
    for m in _COEF.finditer(s):
        win = s[m.end():m.end() + 60]
        other = _OTHER.search(win)
        if other:
            win = win[:other.start()]
        pm = _P.search(win)
        if not pm:
            continue
        try:
            v, se = abs(_num(m.group("v"))), float(m.group("se"))
            rep = float(pm.group("p"))
        except ValueError:
            continue
        if se <= 0 or rep <= 0 or rep > 1:
            continue
        hv, hs = 0.5 * 10 ** -_dec(m.group("v")), 0.5 * 10 ** -_dec(m.group("se"))
        if se - hs <= 0:
            continue
        z_hi = (v + hv) / (se - hs)
        p_best = p_z(z_hi)                       # the smallest p these figures can give, with any sample
        counted[0] += 1
        rel, dp = pm.group("rel"), _dec(pm.group("p"))
        if rel in ("<", "≤", "<="):
            bad = p_best >= rep + (1e-12 if rel != "<" else 0.0)
        elif rel == "=":
            bad = p_best > rep + 0.5 * 10 ** -dp + 1e-12
        else:
            bad = False
        if not bad:
            continue
        z = v / se
        lead = pm.group("p").startswith("0")
        best = _fmt_p(p_best, lead).replace("p = ", "").replace("p < ", "< ")
        note = (f"{_bare(m.group('lab'))} = {_bare(m.group('v'))} with SE = {m.group('se')} gives z = {z:.2f}, so p is at least "
                f"{best}, not p {rel} {pm.group('p')}.")
        sig_said = (rel in ("<", "≤", "<=") and rep <= 0.05) or (rel == "=" and rep < 0.05)
        decision = sig_said and p_best >= 0.05
        if decision:
            note += " By its own figures, the result is not significant at .05."
        notes.append({"note": note, "kind": "test", "decision": decision})
    return notes


# ---- 6.6, round 66 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a mean rated on a scale lies inside it, and
# its standard deviation can't be more than half the scale's range ("M = 5.6 on a 5-point scale", "SD = 2.8" on a 1–5 scale);
# a share of people can't be more than 100% ("112% of students"; a rise of 112% can be)
_WORDNUM = {"three": 3, "four": 4, "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10, "eleven": 11}
_SCALE_N = re.compile(r"\b(?P<n>\d{1,2}|three|four|five|six|seven|eight|nine|ten|eleven)[\s-]*points?\s+(?:likert[\s-]*(?:type\s+)?)?(?:rating\s+)?scales?\b", re.I)
_SCALE_RANGE = re.compile(r"\b(?:scale|scored|rated|ranging|ranged|range)\b[^.;]{0,30}?\bfrom\s+(?P<a>[01])\s*(?:\([^)]{1,30}\)\s*)?to\s+(?P<b>\d{1,2})\b"
                          r"|\bscale\s+of\s+(?P<c>[01])\s*(?:-|\u2013|to)\s*(?P<d>\d{1,2})\b|\b(?P<e>[01])\s*(?:-|\u2013)\s*(?P<f>\d{1,2})\s+(?:likert\s+)?scale\b", re.I)
_MEAN = re.compile(r"(?:(?<![A-Za-z])M\s*=\s*|\bmean\s+(?:score\s+|rating\s+|value\s+)?(?:of\s+|was\s+|=\s*|:\s*)|\bmean\s*=\s*)(?P<v>-?\d*\.?\d+)(?!\d|\.\d)"
                   r"(?!\s*(?:years?|yrs?|months?|weeks?|days?|hours?|hrs?|minutes?|mins?|kg|g\b|cm|mm|m\b|%|kcal|times|points?\s+(?:higher|lower)))", re.I)
_SD_ANY = re.compile(r"(?<![A-Za-z])(?:SD|s\.d\.|standard\s+deviation)\s*(?:of|=|:)?\s*(?:=\s*)?(?P<v>\d*\.?\d+)", re.I)
_SHARE_OVER = re.compile(r"(?<![\d.,])(?P<v>\d{3,}(?:\.\d+)?)\s?%\s+of\s+(?:the\s+|all\s+)?(?:[\w-]+\s+){0,2}?(?:students|participants|respondents|patients|"
                         r"nurses|people|persons|children|women|men|adults|pupils|teachers|households|users|mothers|workers|learners|"
                         r"infants|girls|boys|cases)\b", re.I)


def _scale_of(s):
    """(lowest, highest) of the rating scale the sentence names, or None."""
    m = _SCALE_RANGE.search(s)
    if m:
        a = m.group("a") or m.group("c") or m.group("e")
        b = m.group("b") or m.group("d") or m.group("f")
        lo, hi = int(a), int(b)
        return (lo, hi) if hi > lo else None
    m = _SCALE_N.search(s)
    if m:
        n = m.group("n").lower()
        n = _WORDNUM.get(n) or int(n)
        return (1, n) if 3 <= n <= 11 else None
    return None


def _scale_notes(s):
    notes = []
    sc = _scale_of(s)
    if sc:
        lo, hi = sc
        top = (hi - lo) / 2.0
        for m in _MEAN.finditer(s):
            try:
                v = float(m.group("v"))
            except ValueError:
                continue
            if v > hi + 1e-9 or v < lo - 1e-9:
                notes.append({"note": f"A mean on a {lo}–{hi} scale lies between {lo} and {hi} (it is written {m.group('v')}).", "kind": "value", "decision": False})
            sd = _SD_ANY.search(s[m.end():m.end() + 30])     # its own SD, right after it ("M = 3.6, SD = 0.9")
            if sd:
                try:
                    w = float(sd.group("v"))
                except ValueError:
                    continue
                if w > top * 1.1 + 0.05:        # (half the range is the most it can be; a little more for a small sample)
                    notes.append({"note": f"On a {lo}–{hi} scale a standard deviation can’t be more than about {top:g} (it is written {sd.group('v')}).",
                                  "kind": "value", "decision": False})
    for m in _SHARE_OVER.finditer(s):
        try:
            v = float(m.group("v"))
        except ValueError:
            continue
        if v > 100:
            notes.append({"note": f"A share of people can’t be more than 100% (it is written {m.group('v')}%).", "kind": "value", "decision": False})
    return notes


# 6.6, round 77 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): Cronbach's alpha called more reliable than its
# value. 0.70 is the value usually taken as acceptable (Nunnally, 1978; George & Mallery: 0.70 acceptable, 0.80 good, 0.90
# excellent): "α = 0.65, which shows good reliability" says more than its figure. Said only of the person's own writing, and
# not where the sentence names a lower standard of its own (0.60 for a new or exploratory scale) or a pilot. RamzAI's
# sentences that report a paper's own words are held to the paper, not to this (problems() doesn't run it).
_ALPHA_VAL = re.compile(r"(?:Cronbach['’]?s?\s+(?:alpha|α|coefficient\s+alpha)|(?<![A-Za-z])α)(?:\s+(?:coefficient|value|score))?\s*(?:\(\s*α\s*\)\s*)?"
                        r"(?:of|=|was|were|is|:)?\s*(?:=\s*)?(?P<v>0?\.\d+)(?!\d)", re.I)
_ALPHA_GOOD = r"(?:excellent|very\s+good|good|high|strong|acceptable|adequate|satisfactory|sufficient)"
_ALPHA_CLAIM = re.compile(r"\b" + _ALPHA_GOOD + r"\s+(?:\w+\s+){0,2}?(?:reliability|internal\s+consistency|consistency)\b|"
                          r"\b(?:reliability|internal\s+consistency|consistency|alpha|α)\b[^.;]{0,40}?\b(?:was|is|were|are|be|been|being|as|indicat\w*|"
                          r"show\w*|suggest\w*)\s+(?:\w+\s+){0,2}?" + _ALPHA_GOOD + r"\b|\b(?:highly\s+)?reliable\b", re.I)
_ALPHA_LOWER = re.compile(r"(?<![\d.])0?\.6\d?(?!\d)\s*(?:or\s+(?:more|above|higher)|and\s+above|as\s+(?:the\s+)?(?:cut|thresh|minimum))|"
                          r"\b(?:exploratory|newly\s+developed|new\s+(?:scale|instrument|tool|questionnaire)|pilot)\b|\bquestionable\b|\bpoor\b|"
                          r"\bbelow\s+(?:the\s+)?(?:usual|accepted|recommended)|\blow\b", re.I)


def _alpha_word_notes(s):
    notes = []
    if _ALPHA_LOWER.search(s):
        return notes
    for m in _ALPHA_VAL.finditer(s):
        try:
            v = float(m.group("v"))
        except ValueError:
            continue
        if not 0 < v < 0.70:
            continue
        if _ALPHA_CLAIM.search(s):
            notes.append({"note": f"Cronbach’s alpha of {m.group('v')} is below 0.70, the value usually taken as acceptable: it can’t be called "
                                  "acceptable or good unless a lower standard is named and explained (0.60 is sometimes accepted for a new scale).",
                          "kind": "value", "decision": False})
            break
    return notes


# 6.6, round 79: a correlation and the share of the variance it explains. r = .40 explains r² = 16% of the variance, not 40%;
# R² = .25 explains 25%. "r = .40, explaining 40% of the variance" mixes r with r², a common slip.
_VAR_SHARE = re.compile(r"(?P<pct>\d+(?:\.\d+)?)\s?%\s+of\s+(?:the\s+)?(?:total\s+)?(?:variance|variation|variability)", re.I)
_R_COEF = re.compile(r"(?<![A-Za-z])(?P<lab>[rR](?:\s*(?:²|\^?2))?|R\s*-?\s*squared)\s*(?:\(\s*\d+\s*\))?\s*=\s*(?P<v>[-−–]?\s?\d*\.\d+)")


def _variance_notes(s):
    notes = []
    vs = list(_VAR_SHARE.finditer(s))
    if not vs:
        return notes
    for m in _R_COEF.finditer(s):
        try:
            v = abs(_num(m.group("v")))
        except ValueError:
            continue
        if v > 1:
            continue
        sq = bool(re.search(r"²|\^?2$|squared", m.group("lab"), re.I)) or m.group("lab").startswith("R") and "2" in m.group("lab")
        share = v * 100 if sq else v * v * 100
        d = _dec(m.group("v"))
        for x in vs:
            said = float(x.group("pct"))
            tol = 0.5 * 10 ** -_dec(x.group("pct")) + (100 * 0.5 * 10 ** -d * (1 if sq else 2 * v + 0.5 * 10 ** -d))
            if abs(said - share) <= tol + 1e-9:
                continue
            if not sq and abs(said - v * 100) <= 0.5 * 10 ** -_dec(x.group("pct")) + 100 * 0.5 * 10 ** -d + 1e-9:
                notes.append({"note": f"r = {_bare(m.group('v'))} explains r² = {share:.0f}% of the variance, not {x.group('pct')}% (that is r itself).",
                              "kind": "value", "decision": False})
                return notes
    return notes


# 6.6, round 93: a correlation's strength in words, held to its own r in the same sentence (as round 22 holds RamzAI's to the
# papers' r): weak below .3, moderate .3 to .5, strong .5 and above, with the same room either side (strong from .4, weak
# below .4, moderate .2 to .7). "A strong positive correlation (r = .23)" says more than its figure.
_STR_SAID = re.compile(r"\b(strong(?:ly)?|weak(?:ly)?|moderate(?:ly)?)\s+(?:(?:and\s+)?(?:positive(?:ly)?|negative(?:ly)?|inverse(?:ly)?|"
                       r"significant(?:ly)?|statistically\s+significant)\s+)?(?:correlat\w*|associat\w*|relationship|relation)\b", re.I)
_R_STR = re.compile(r"(?<![A-Za-z])(?:r|ρ|rho|r_?s|rs|Pearson['’]?s\s+r|Spearman['’]?s\s+(?:rho|ρ))\s*(?:\(\s*\d+\s*\))?\s*=\s*(?P<v>[-−–]?\s?0?\.\d+)", re.I)


def _strength_notes(s):
    m = _STR_SAID.search(s)
    if not m:
        return []
    vals = []
    for x in _R_STR.finditer(s):
        try:
            vals.append((abs(_num(x.group("v"))), _bare(x.group("v"))))
        except ValueError:
            continue
    if not vals:
        return []
    w = m.group(1).lower()
    band = "strong" if w.startswith("strong") else "weak" if w.startswith("weak") else "moderate"
    ok = {"strong": lambda r: r >= 0.4, "weak": lambda r: r < 0.4, "moderate": lambda r: 0.2 <= r < 0.7}[band]
    if any(ok(r) for r, _t in vals):
        return []
    r, shown = max(vals)
    usual = "weak" if r < 0.3 else "moderate" if r < 0.5 else "strong"
    return [{"note": f"r = {shown} is usually called {usual} (below .3 weak, .3 to .5 moderate, .5 and above strong), not {band}.",
             "kind": "value", "decision": False}]



# 6.6, round 95: an effect size's word, held to its own value in the same sentence, as round 93 does for r. Cohen's (1988)
# usual words: d (and Hedges' g) 0.2 small, 0.5 medium, 0.8 large; η², partial η² and ω² .01 small, .06 medium, .14 large.
# A word stands with room either side ("a medium effect, d = 0.42"; "a large effect, η² = .12"); only a word that no value in
# the sentence comes near is said ("a large effect (d = 0.25)").
_ES_WORD = r"(negligible|trivial|small|medium|moderate|large|very\s+large|huge)"
_ES_SAID = re.compile(r"\b" + _ES_WORD + r"(?:\s*(?:-|–|to|-to-)\s*" + _ES_WORD + r")?(?:\s*-?\s*sized?)?\s+effects?(?:\s+sizes?)?\b|"
                      r"\beffects?(?:\s+sizes?)?\s+(?:was|were|is|are|of|indicat\w+|suggest\w*|showed)\s+(?:a\s+|an\s+)?(?:\w+\s+)?" + _ES_WORD + r"\b", re.I)
_D_VAL = re.compile(r"(?:(?<![A-Za-z])Cohen['’]?s\s+d|(?<![A-Za-z])Hedges['’]?\s*g|(?<![A-Za-z'’])[dg])\s*=\s*(?P<v>[-−–]?\s?\d*\.?\d+)")
_ETA_VAL = re.compile(r"(?P<lab>(?:partial\s+)?(?:η|eta)\s*(?:p\s*)?(?:²|\^?2|[-\s]?squared)(?:\s*p)?|(?:ω|omega)\s*(?:²|\^?2|[-\s]?squared))\s*=\s*(?P<v>0?\.\d+)", re.I)
_ES_BAND = {"d": [(0.2, "negligible"), (0.5, "small"), (0.8, "medium"), (99, "large")],
            "eta": [(0.01, "negligible"), (0.06, "small"), (0.14, "medium"), (99, "large")]}
_ES_OK = {"d": {"negligible": lambda v: v < 0.35, "small": lambda v: v < 0.65, "medium": lambda v: 0.35 <= v < 1.0, "large": lambda v: v >= 0.65},
          "eta": {"negligible": lambda v: v < 0.03, "small": lambda v: v < 0.10, "medium": lambda v: 0.03 <= v < 0.20, "large": lambda v: v >= 0.10}}
_ES_SCALE = {"d": "0.2 small, 0.5 medium, 0.8 large", "eta": ".01 small, .06 medium, .14 large"}


def _es_band(w):
    w = re.sub(r"\s+", " ", (w or "").lower())
    return {"trivial": "negligible", "moderate": "medium", "very large": "large", "huge": "large"}.get(w, w)


def _effect_notes(s):
    m = _ES_SAID.search(s)
    if not m:
        return []
    words = [_es_band(w) for w in m.groups() if w]
    vals = []
    for x in _D_VAL.finditer(s):
        try:
            v = abs(_num(x.group("v")))
        except ValueError:
            continue
        if v < 5:
            lab = x.group(0).split("=")[0].strip()
            vals.append(("d", v, (("g" if lab.lower().endswith("g") else "d") + " = " + _bare(x.group("v")))))
    for x in _ETA_VAL.finditer(s):
        try:
            v = abs(_num(x.group("v")))
        except ValueError:
            continue
        lab = x.group("lab").lower()
        name = "ω²" if ("ω" in lab or "omega" in lab) else ("partial η²" if ("partial" in lab or "p" in lab.replace("partial", "").replace("eta", "")) else "η²")
        vals.append(("eta", v, name + " = " + _bare(x.group("v"))))
    if not vals:
        return []
    if any(_ES_OK[k][w](v) for k, v, _t in vals for w in words if w in _ES_OK[k]):
        return []
    k, v, shown = max(vals, key=lambda t: t[1] / (0.8 if t[0] == "d" else 0.14))
    usual = next(name for top, name in _ES_BAND[k] if v < top)
    said = words[0] if len(words) == 1 else " to ".join(words)
    return [{"note": f"{shown} is usually called {usual} ({_ES_SCALE[k]}), not {said}.", "kind": "value", "decision": False}]


# 6.6, round 96: a percentage of a total the sentence states is a whole number of that total: 37.5% of 150 respondents would
# be 56.25 people, and no count gives it (56 is 37.3%, 57 is 38.0%). Only where the total and the percentage stand together
# with nothing between them that could be another total ("Of the 150 respondents, 37.5% were male and 62.5% female"; "A
# total of 150 students took part, and 37.5% were male"): a count, "of whom", "of those", a share "of the women" or "of
# their time" is passed by. The person's writing only (a paper may give a share of those who answered a question).
_TOTAL_WHO = (r"(?:[A-Za-z][\w'’\-]*\s+){0,3}?(?:respondents|participants|students|patients|pupils|nurses|midwives|women|men|mothers|"
              r"fathers|children|adolescents|people|persons|workers|teachers|households|questionnaires|cases|subjects|clients|staff|"
              r"employees|traders|farmers|members|adults|infants|individuals|interviewees|caregivers|parents|youths?|girls|boys|"
              r"learners|graduates|users|residents|athletes|pharmacists|doctors|physicians|officers)")
_TOTAL_AT = [re.compile(r"^(?:(?:and|but|so|thus|however|overall),?\s+)?(?:of|out\s+of|among|amongst)\s+(?:the\s+|all\s+|a\s+total\s+of\s+|these\s+)?"
                        r"(?P<n>\d{1,3}(?:,\d{3})?|\d{4,5})(?![\d.,%])\s+(?P<who>" + _TOTAL_WHO + r")\b[^,;:]{0,60}?,(?P<rest>[^;:]*)", re.I),
             re.compile(r"^(?:a\s+total\s+of\s+)?(?P<n>\d{1,3}(?:,\d{3})?|\d{4,5})(?![\d.,%])\s+(?P<who>" + _TOTAL_WHO + r")\s+(?:\w+\s+){0,2}?"
                        r"(?:took\s+part|participated|responded|were\s+(?:included|interviewed|surveyed|sampled|recruited|studied|enrolled|analy[sz]ed))"
                        r"\b[^,;:]{0,40}?(?:,\s*(?:and\s+)?|\s+and\s+)(?P<rest>[^;:]*)", re.I)]
_PCT_ONE = re.compile(r"(?<![\d.,])(?P<p>\d{1,3}(?:\.\d{1,2})?)\s?%")
_PCT_PASS = re.compile(r"^\s*\)?\s*(?:of\s+(?!them\b|these\b|those\b|the\s+total\b|all\b|the\s+sample\b|the\s+respondents\b|the\s+participants\b)|"
                       r"(?:increase|decrease|rise|fall|drop|more|less|fewer|higher|lower|points?|reduction|growth|change|CI\b|confidence|chance|"
                       r"probability|likelihood|response\s+rate|of\s+the\s+time))", re.I)
_PCT_STOP = re.compile(r"\b(?:of\s+whom|of\s+those|of\s+these\s+\w+|among\s+(?:those|them)|of\s+them\s+who)\b|(?<![\d.,])\d[\d,]*(?![\d.,]*\s?%)(?!\s*(?:years?|months?|weeks?|days?|hours?)\b)", re.I)


def _whole_notes(s):
    for rx in _TOTAL_AT:
        m = rx.match(s.strip())
        if not m:
            continue
        try:
            n = int(m.group("n").replace(",", ""))
        except ValueError:
            return []
        if n < 2 or 1900 <= n <= 2100:
            return []
        rest = m.group("rest")
        stop = _PCT_STOP.search(rest)
        scan = rest[:stop.start()] if stop else rest
        notes = []
        for x in _PCT_ONE.finditer(scan):
            if _PCT_PASS.match(scan[x.end():]) or re.search(r"\d\s*\(\s*$", scan[:x.start()]):
                continue
            p, d = float(x.group("p")), _dec(x.group("p"))
            if p > 100:
                continue
            tol = 0.5 * 10 ** -d + 0.05
            k = round(p * n / 100)
            if any(abs(100.0 * j / n - p) <= tol + 1e-9 for j in (k - 1, k, k + 1) if 0 <= j <= n):
                continue
            lo = int(p * n / 100)
            dd = max(1, d)
            would = f"{p * n / 100:.2f}".rstrip("0").rstrip(".")
            notes.append({"note": f"{x.group('p')}% of {n:,} would be {would}, and no whole number of them gives it ({lo:,} is "
                                  f"{100.0 * lo / n:.{dd}f}%, {lo + 1:,} is {100.0 * (lo + 1) / n:.{dd}f}%). Check the figure, or give the number "
                                  f"it is a share of (if some didn’t answer).", "kind": "percent", "decision": False})
            break
        return notes
    return []


# 6.6, round 97: a null hypothesis says there is no relationship or difference, an alternative hypothesis that there is one.
# "H0: There is a significant relationship between …" (or "Ha: There is no significant …") states the other one. Read only
# where the label is plain: H0, Ho, H₀ (H01, Ho1…) or "the null hypothesis"; Ha, Hₐ or "the alternative hypothesis" ("H1" is
# also how hypotheses are numbered, so it isn't read). The person's writing only.
_H_NULL = re.compile(r"(?:^|(?<=[\s(]))(?:H\s?[0oO₀](?:\s?\d{1,2}|[₁₂₃₄₅₆₇₈₉])?\s*(?:\)\s*)?[:–-]\s*|"
                     r"(?:the\s+)?null\s+hypothes[ie]s(?:\s+\d{1,2})?(?:\s*\(\s*H\s?[0oO₀]\d?\s*\))?"
                     r"\s*(?:[:.–-]\s*|\s+(?:states?|stated|was|is|reads?)\s+(?:that\s+)?))(?P<h>.{12,})", re.I)
_H_ALT = re.compile(r"(?:^|(?<=[\s(]))(?:H\s?[aAₐ](?:\s?\d{1,2})?\s*(?:\)\s*)?[:–-]\s*|"
                    r"(?:the\s+)?alternative\s+hypothes[ie]s(?:\s+\d{1,2})?(?:\s*\(\s*H\s?[aA1]\d?\s*\))?"
                    r"\s*(?:[:.–-]\s*|\s+(?:states?|stated|was|is|reads?)\s+(?:that\s+)?))(?P<h>.{12,})", re.I)
_H_EFFECT = re.compile(r"\bthere\s+(?:is|are|will\s+be|exists?|was)\s+(?:an?\s+)?(?:statistically\s+)?significant\b|"
                       r"\b(?:has|have|will\s+have|had)\s+an?\s+(?:statistically\s+)?significant\b|"
                       r"\b(?:is|are|was|were|will\s+be)\s+(?:statistically\s+)?significantly\b|"
                       r"\bsignificantly\s+(?:\w+\s+){0,2}?(?:affects?|influences?|predicts?|increases?|reduces?|decreases?|improves?|relates?|"
                       r"related|associated|correlated|differs?|different|higher|lower)\b", re.I)
_H_NEG = re.compile(r"\b(?:no|not|none|neither|nor|never|without)\b|n['’]t\b|=|\bequal\b|\bthe\s+same\b", re.I)
_H_NULLFORM = re.compile(r"\bthere\s+(?:is|are|will\s+be|was)\s+no\b|\bno\s+(?:statistically\s+)?significant\b|"
                         r"\b(?:does|do|did|will|is|are|was|were)\s+not\s+(?:\w+\s+)?(?:statistically\s+)?significant(?:ly)?\b|"
                         r"\bnot\s+(?:statistically\s+)?significantly\b", re.I)


def _hypothesis_notes(s):
    m = _H_NULL.search(s)
    if m:
        h = m.group("h")
        if _H_EFFECT.search(h) and not _H_NEG.search(h):
            return [{"note": "A null hypothesis says there is no relationship or difference (“There is no significant relationship between "
                             "… and …”); this one says there is one, which is the alternative hypothesis.", "kind": "hypothesis", "decision": False}]
        return []
    m = _H_ALT.search(s)
    if m and _H_NULLFORM.search(m.group("h")):
        return [{"note": "An alternative hypothesis says there is a relationship or difference; this one says there is none, which is the "
                         "null hypothesis.", "kind": "hypothesis", "decision": False}]
    return []



# 6.6, round 104 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a ratio's words and a correlation's sign, held to its own figure. "Late phone use increased the odds of poor
# sleep (OR = 0.62)" says the other way from its own figure (below 1 is lower odds); "a positive correlation (r = −.32)" too.
_RATIO_ONE = re.compile(r"(?<![A-Za-z])(?P<lab>a?OR|AOR|aOR|RR|aRR|HR|aHR|IRR|PR|odds\s+ratio|risk\s+ratio|hazard\s+ratio|relative\s+risk)"
                        r"\s*(?:[:=]\s*|of\s+|was\s+|\s)(?P<v>\d*\.?\d+)(?!\s*[-–]\s*\d)")
_RATIO_UP_W = re.compile(r"\b(?:increased?|increases|increasing|higher|greater|more\s+likely|elevated|raised|risk\s+factors?|twice|doubled?|"
                         r"tripled?|times\s+(?:more|higher|as\s+likely))\b", re.I)
_RATIO_DOWN_W = re.compile(r"\b(?:decreased?|decreases|decreasing|reduced?|reduces|reducing|lower|less\s+likely|protective|protected|"
                           r"fewer|halved)\b", re.I)
_RS_SIGN = re.compile(r"\b(?P<w>positive(?:ly)?|negative(?:ly)?|inverse(?:ly)?)\s+(?:and\s+)?(?:(?:statistically\s+)?significant(?:ly)?\s+)?"
                      r"(?:associat\w*|correlat\w*|relat\w*)|\b(?:associat\w*|correlat\w*|related)\s+(?:(?:statistically\s+)?significantly\s+)?"
                      r"(?P<w2>positively|negatively|inversely)\b", re.I)
_R_SIGNED = re.compile(r"(?<![A-Za-z])(?:r|ρ|rho|r_?s|rs)\s*(?:\(\s*\d+\s*\))?\s*=\s*(?P<v>[-−–]?\s?0?\.\d+)", re.I)


def _direction_notes(s):
    notes = []
    ratios = [m for m in _RATIO_ONE.finditer(s)]
    if len(ratios) == 1:
        m = ratios[0]
        try:
            v = float(m.group("v"))
        except ValueError:
            v = None
        lab = re.sub(r"\s+", " ", m.group("lab"))
        what = "risk" if re.search(r"RR|risk", lab) else "hazard" if re.search(r"HR|hazard", lab) else "rate" if "IRR" in lab else \
            "prevalence" if lab == "PR" else "odds"
        rest = s[:m.start()] + " " + s[m.end():]
        up, down = bool(_RATIO_UP_W.search(rest)), bool(_RATIO_DOWN_W.search(rest))
        if v is not None and 0 < v < 1 and up and not down:
            notes.append({"note": f"{lab} = {m.group('v')} is below 1: lower {what}, not higher.", "kind": "value", "decision": False})
        elif v is not None and v > 1 and down and not up:
            notes.append({"note": f"{lab} = {m.group('v')} is above 1: higher {what}, not lower.", "kind": "value", "decision": False})
    signs = {("+" if (x.group("w") or x.group("w2")).lower().startswith("pos") else "-") for x in _RS_SIGN.finditer(s)}
    rs = []
    for x in _R_SIGNED.finditer(s):
        try:
            rs.append((_num(x.group("v")), _bare(x.group("v"))))
        except ValueError:
            continue
    if len(signs) == 1 and len(rs) == 1 and rs[0][0] != 0:
        r, shown = rs[0]
        if (r < 0) != ("-" in signs):
            notes.append({"note": f"r = {shown} is {'negative' if r < 0 else 'positive'} (as one rises, the other {'falls' if r < 0 else 'rises'}), "
                                  f"not {'positive' if r < 0 else 'negative'}.", "kind": "value", "decision": False})
    return notes


# 6.6, round 112 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): "the majority" is more than half. "The majority of
# the respondents (45%) were female" (the largest group, not more than half: a slip examiners mark in many results chapters),
# "most … (38.5%)", "more than half (48%)", and "a minority (62%)" say what their own figure doesn't.
_MAJ_W = re.compile(r"\b(?P<w>(?:the\s+|a\s+)?(?:vast\s+|overwhelming\s+|large\s+|slight\s+|simple\s+|clear\s+)?majority|most|more\s+than\s+half|"
                    r"over\s+half|above\s+half|greater\s+part)\b(?!\s+(?:likely|common|commonly|frequent|frequently|often|important|"
                    r"significant|cited|reported|used|effective|prevalent|popular|preferred)\b)", re.I)
_MIN_W = re.compile(r"\b(?P<w>(?:a|the)\s+minority|less\s+than\s+half|fewer\s+than\s+half|under\s+half)\b", re.I)
_PCT_NEXT = re.compile(r"(?<![\d.])(?P<p>\d{1,3}(?:\.\d+)?)\s?%")


def _majority_notes(s):
    notes = []
    for rx, more in ((_MAJ_W, True), (_MIN_W, False)):
        for m in rx.finditer(s):
            rest = re.split(r"[;:]|\.(?!\d)|,\s*(?:but|while|whereas|and)\b|\b(?:but|whereas|while|of\s+(?:those|whom|these))\b",
                            s[m.end():m.end() + 80], flags=re.I)[0]
            pm = _PCT_NEXT.search(rest)
            if not pm or re.search(r"\b(?:of|among|amongst)\s+(?:the\s+|those\s+|these\s+)?$", rest[:pm.start()], re.I):
                continue                       # ("the majority of the 45% who answered": a share of a share)
            p = float(pm.group("p"))
            w = re.sub(r"\s+", " ", m.group("w")).strip()
            w = w[:1].lower() + w[1:]
            if more and p <= 50:
                notes.append({"note": f"{pm.group('p')}% is not more than half: “{w}” means more than 50% (say “the largest group”, or give the figure).",
                              "kind": "value", "decision": False})
                return notes
            if not more and p >= 50:
                notes.append({"note": f"{pm.group('p')}% is not less than half: it isn’t “{w}”.", "kind": "value", "decision": False})
                return notes
    return notes


# 6.6, round 113: "significant" held to its own p. "The difference was statistically significant (p = 0.06)", or "no significant
# association (p = 0.01)", says the other way from its p at .05 (or at the level the sentence names: "α = .10", "at the 10%
# level"). Each clause with one p on its own; "marginally", "a trend towards", "approached" are left alone.
_ALPHA_SET = re.compile(r"(?:α|alpha|significance\s+level|level\s+of\s+significance)\s*(?:=|of|was\s+set\s+at|set\s+at|at)\s*(?P<a>0?\.\d+)|"
                        r"\bat\s+(?:the\s+)?(?P<pc>\d{1,2}(?:\.\d+)?)\s*%\s+(?:level|significance)", re.I)
_MARGINAL = re.compile(r"\b(?:marginal(?:ly)?|borderline|approach(?:ed|ing)?|trend(?:ed|ing|s)?\s+(?:toward|towards)|nearly|almost|close\s+to)\b", re.I)


def _p_word_notes(s):
    am = _ALPHA_SET.search(s)
    alpha = float(am.group("a")) if am and am.group("a") else (float(am.group("pc")) / 100 if am and am.group("pc") else 0.05)
    dec = max(2, len(repr(alpha).split(".")[1]) if "." in repr(alpha) else 2)
    a_str = f"{alpha:.{dec}f}".lstrip("0") if alpha < 1 else str(alpha)
    notes = []
    bounds = [0] + [m.end() for m in _CLAUSE.finditer(s)] + [len(s)]
    for a, z in zip(bounds, bounds[1:]):
        c = s[a:z]
        ps = list(_P.finditer(c))
        if len(ps) != 1 or _MARGINAL.search(c):
            continue
        pm = ps[0]
        words = c[:pm.start()] + " " + c[pm.end():]
        said = "nonsig" if _SIG_NOT.search(words) else "sig" if _SIG_YES.search(words) else ""
        if not said:
            continue
        try:
            v = float(pm.group("p"))
        except ValueError:
            continue
        rel = pm.group("rel")
        tiny = 0.5 * 10 ** -_dec(pm.group("p"))          # (a p written .05 may be .045 to .055)
        sig = (rel in ("<", "≤", "<=") and v <= alpha) or (rel == "=" and v < alpha - tiny)
        nonsig = (rel in (">", "≥", ">=") and v >= alpha) or (rel == "=" and v > alpha + tiny)
        shown = (rel if rel != "=" else "=") + " " + pm.group("p")
        if said == "sig" and nonsig:
            notes.append({"note": f"p {shown} is above {a_str}: not significant at that level, though the sentence says it is.", "kind": "value", "decision": True})
        elif said == "nonsig" and sig:
            notes.append({"note": f"p {shown} is below {a_str}: significant at that level, though the sentence says it isn’t.", "kind": "value", "decision": True})
    return notes



# 6.6, round 155 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): the tests that check another test's assumptions,
# read the right way round. Their p answers "do the data depart from the assumption?": below .05 (or the level the sentence
# sets), they do. Tests of normality (Shapiro–Wilk, Kolmogorov–Smirnov, Lilliefors, Anderson–Darling…), of equal variances
# (Levene's, Bartlett's, Fligner's…), of a logistic model's fit (Hosmer and Lemeshow's) and Mauchly's test of sphericity:
# "normally distributed (Shapiro–Wilk p = .002)", "Levene's test (p = .003) showed equal variances" and "the model fitted
# well (Hosmer–Lemeshow p = .02)" say the opposite of their p, and so does "not normally distributed (p = .34)". Bartlett's
# test of sphericity (before a factor analysis) the other way: the data suit factor analysis only when its p is below .05.
# Only a clause that names one such test, gives one p and says what it found; a concession ("although", "despite"), other
# evidence (skewness, kurtosis, a Q–Q plot, a histogram, the central limit theorem) or a second test leaves it alone, and a
# significance word that already disagrees with the p is round 113's note.
_AS_NAMES = (
    ("norm", re.compile(r"\bshapiro\s*[-–—]?\s*wilks?[a-e]?\b|\bS\s*[-–]\s*W\s+(?:test|statistic)", re.I), "Shapiro–Wilk"),
    ("norm", re.compile(r"\bkolmogorov\s*[-–—]?\s*smirnov[a-e]?\b|\bK\s*[-–]\s*S\s+(?:test|statistic)", re.I), "Kolmogorov–Smirnov"),
    ("norm", re.compile(r"\blilliefors\b", re.I), "Lilliefors"),
    ("norm", re.compile(r"\banderson\s*[-–—]?\s*darling\b", re.I), "Anderson–Darling"),
    ("norm", re.compile(r"\bd['’]\s?agostino\b", re.I), "D’Agostino’s test"),
    ("norm", re.compile(r"\bjarque\s*[-–—]?\s*bera\b", re.I), "Jarque–Bera"),
    ("norm", re.compile(r"\btests?\s+(?:of|for)\s+normality\b|\bnormality\s+tests?\b", re.I), "The test of normality"),
    ("var", re.compile(r"\blevene['’]?s?\b", re.I), "Levene’s test"),
    ("var", re.compile(r"\bbartlett['’]?s?\s+test\s+(?:of|for)\s+(?:the\s+)?(?:homogeneity|equal(?:ity)?)\b", re.I), "Bartlett’s test"),
    ("var", re.compile(r"\bfligner\b", re.I), "Fligner’s test"),
    ("var", re.compile(r"\bhartley['’]?s?\s+(?:F\s*[-–]?\s*max|test)\b|\bF\s*[-–]?\s*max\s+test\b", re.I), "Hartley’s F-max"),
    ("var", re.compile(r"\bbrown\s*[-–—]?\s*forsythe\s+test\s+(?:of|for)\s+(?:the\s+)?(?:homogeneity|equal)", re.I), "The Brown–Forsythe test"),
    ("fit", re.compile(r"\bhosmer\s*(?:[-–—]|\band\b|&)?\s*lemeshow\b|\bH\s*[-–]\s*L\s+(?:test|statistic|goodness)", re.I), "Hosmer and Lemeshow’s test"),
    ("sph", re.compile(r"\bmauchly['’]?s?\b", re.I), "Mauchly’s test"),
    ("bart", re.compile(r"\bbartlett['’]?s?\s+test\s+of\s+sphericity\b", re.I), "Bartlett’s test of sphericity"),
)
_AS_BART = re.compile(r"\bbartlett['’]?s?\b", re.I)
_AS_FA = re.compile(r"\bKMO\b|\bkaiser\s*[-–—]?\s*meyer|\bsampling\s+adequacy\b|\bfactor(?:ab\w+)?\b|\bprincipal\s+components?\b", re.I)
_AS_SKIP = re.compile(r"\b(?:although|though|even\s+though|despite|in\s+spite|nevertheless|nonetheless|however|notwithstanding)\b|"
                      r"\bskewness\b|\bkurtosis\b|\bQ\s*[-–]?\s*Q\b|\bP\s*[-–]?\s*P\s+plots?\b|\bhistograms?\b|\bcentral\s+limit\b", re.I)
_AS_NEG = re.compile(r"\b(?:not|no|never|cannot|neither|nor|without|lack(?:ed|s|ing)?|fail(?:ed|s)?|unable|hardly)\b|n['’]t\b|"
                     r"\brather\s+than\b|\binstead\s+of\b", re.I)
_AS_W = r"(?:(?!not\b|no\b|never\b)\w+\s+)"          # (one word, never a "not")
_AS_SAYS = r"\b(?:indicat\w*|show\w*|suggest\w*|mean\w*|confirm\w*|support\w*|demonstrat\w*|impl\w+)\s+(?:that\s+)?"
_AS_VH = r"(?:homogeneity\s+of\s+(?:the\s+)?variances?|homogeneity|equal(?:ity\s+of)?\s+variances?)"
_AS_ASK = re.compile(r"\b(?:whether|if|to\s+(?:test|check|assess|examine|determine|see|confirm|establish|verify|ascertain|ensure)|"
                     r"null\s+hypothesis|assum(?:es|ing)\s+that|requires?|assumption\s+that)\b", re.I)
_AS_CLAIM = {
    "norm": (
        re.compile(r"\b(?:did|does|do)\s+not\s+(?:significantly\s+)?(?:deviate|depart|differ)\s+(?:significantly\s+)?from\s+(?:a\s+|the\s+)?normal(?:ity|\s+distribution)?\b|"
                   r"\bno\s+(?:significant\s+)?(?:deviation|departure)\s+from\s+(?:a\s+|the\s+)?normal\w*|"
                   r"\bnormality\s+(?:assumption\s+)?(?:was|is|were)\s+not\s+(?:violated|rejected)|\bno\s+violation\s+of\s+(?:the\s+)?(?:assumption\s+of\s+)?normality|"
                   r"\bnot\s+(?:significantly\s+)?non-?normal\w*", re.I),
        re.compile(r"\bnot\s+" + _AS_W + r"{0,2}?normal(?:ly)?\b|\bnon-?normal(?:ly|ity)?\b|\b(?:deviat|depart)\w*\s+(?:\w+\s+){0,2}?from\s+(?:a\s+|the\s+)?normal\w*|"
                   r"\bnormality\s+(?:\w+\s+){0,2}?(?:violated|not\s+met|rejected|not\s+assumed|not\s+satisfied|not\s+tenable|could\s+not\s+be\s+assumed)|"
                   r"\bviolat\w*\s+(?:\w+\s+){0,3}?normality|\bskewed\b|\babnormal(?:ly)?\s+distribut\w*|\bnot\s+follow\s+(?:a\s+|the\s+)?normal", re.I),
        re.compile(r"\bnormal(?:ly)?\s+distribut\w*|\bnormal\s+distribution|\bnormality\s+(?:assumption\s+)?(?:(?:was|is|were|could\s+be)\s+)?(?:assumed|met|satisfied|confirmed|"
                   r"established|tenable|upheld|supported)\b|\b(?:the\s+)?assumption\s+of\s+normality\s+(?:was|is)\s+(?:met|satisfied|tenable|upheld)\b|"
                   r"\bnormality\s+(?:held|holds)\b|\bapproximately\s+normal\b|\b(?:were|was|are|is)\s+normal\b|" + _AS_SAYS + r"normality\b", re.I)),
    "var": (
        re.compile(r"\b(?:homogeneity|equality)\s+of\s+(?:the\s+)?variances?\s+(?:assumption\s+)?(?:was|were|is)\s+not\s+violated|"
                   r"\bno\s+violation\s+of\s+(?:the\s+)?(?:assumption\s+of\s+)?(?:homogeneity|equal)|\bvariances?\s+(?:did|do|does)\s+not\s+(?:significantly\s+)?differ|"
                   r"\bno\s+(?:significant\s+)?difference\s+(?:between|in|among)\s+(?:the\s+)?(?:groups['’]?\s+)?variances|"
                   r"\bvariances?\s+(?:were|was)\s+(?:found\s+to\s+be\s+)?not\s+(?:significantly\s+)?different|"
                   r"\bassumption\s+of\s+" + _AS_VH + r"\s+(?:was|were|is)\s+not\s+violated", re.I),
        re.compile(r"\bunequal\s+variances?|\bequal\s+variances?\s+(?:were\s+|was\s+|could\s+)?not\s+(?:be\s+)?assumed|\bvariances?\s+(?:were|was|are|is)\s+not\s+(?:\w+\s+)?(?:equal|homogeneous|similar)|"
                   r"\bheterogene\w+|\b(?:homogeneity|equality)\s+of\s+(?:the\s+)?variances?\s+(?:\w+\s+){0,2}?(?:violated|not\s+met|not\s+assumed|rejected|not\s+satisfied|not\s+tenable)|"
                   r"\bviolat\w*\s+(?:\w+\s+){0,3}?(?:homogeneity|equality\s+of\s+variances?|equal\s+variances?)|\bvariances?\s+(?:differed|were\s+different|differ\s+significantly)|"
                   r"\bassumption\s+of\s+" + _AS_VH + r"\s+(?:was|were|is)\s+(?:violated|not\s+met|not\s+satisfied|not\s+tenable)", re.I),
        re.compile(r"\bequal\s+variances?\s+(?:(?:were|was|are|could\s+be)\s+)?assumed\b|\b(?:homogeneity|equality)\s+of\s+(?:the\s+)?variances?\s+(?:(?:assumption\s+)?(?:was|were|is|could\s+be)\s+)?"
                   r"(?:assumed|met|satisfied|confirmed|established|tenable|upheld|supported)\b|\bvariances?\s+(?:were|was|are|is)\s+" + _AS_W + r"?(?:equal|homogeneous|similar|comparable)\b|"
                   r"\bhomogeneous\s+variances?|\bhomogeneity\s+of\s+(?:the\s+)?variances?\s+(?:held|holds)\b|"
                   r"\bassumption\s+of\s+" + _AS_VH + r"\s+(?:was|were|is)\s+(?:met|satisfied|tenable|upheld)\b|"
                   + _AS_SAYS + r"(?:the\s+)?(?:groups['’]?\s+)?(?:had\s+)?(?:equal|homogeneous)\s+variances?\b|" + _AS_SAYS + r"(?:homogeneity|equality)\s+of\s+(?:the\s+)?variances?\b", re.I)),
    "fit": (
        re.compile(r"\bno\s+(?:evidence\s+of\s+)?(?:a\s+)?(?:poor|lack\s+of|significant\s+lack\s+of)\s+(?:model\s+)?fit|"
                   r"\bdid\s+not\s+(?:indicate|suggest|show)\s+(?:a\s+)?(?:poor|lack\s+of|bad)\s+(?:model\s+)?fit|\bnot\s+(?:a\s+)?poor(?:ly)?\s+fit", re.I),
        re.compile(r"\bpoor(?:ly)?\s+(?:model\s+)?fit\w*|\bfit\w*\s+(?:the\s+data\s+)?poorly|\black\s+of\s+(?:model\s+)?fit|\b(?:did|does|do)\s+not\s+fit\b|"
                   r"\b(?:inadequate|unsatisfactory|bad|unacceptable)\s+(?:model\s+)?fit|\bmisfit\b|\bnot\s+(?:a\s+)?(?:good|adequate|acceptable)\s+(?:model\s+)?fit|"
                   r"\bnot\s+fit\w*\s+(?:the\s+data\s+)?well", re.I),
        re.compile(r"\b(?:good|adequate|acceptable|satisfactory|excellent|reasonable|appropriate)\s+(?:model\s+)?fit\b|"
                   r"\bfit(?:ted|s)?\s+(?:the\s+(?:observed\s+)?data\s+)?(?:well|adequately|reasonably\s+well|satisfactorily)\b|\bwell[-\s]fitt(?:ed|ing)|\bwell[-\s]calibrated|"
                   r"\bfit\s+(?:was|is)\s+(?:good|adequate|acceptable|satisfactory)\b|\bmodel\s+fit(?:ted|s)?\s+the\s+(?:observed\s+)?data\b(?!\s+(?:poorly|badly))", re.I)),
    "sph": (
        re.compile(r"\bsphericity\s+(?:assumption\s+)?(?:was|were|is)\s+not\s+violated|\bno\s+violation\s+of\s+(?:the\s+)?(?:assumption\s+of\s+)?sphericity", re.I),
        re.compile(r"\bsphericity\s+(?:\w+\s+){0,2}?(?:violated|not\s+met|not\s+assumed|not\s+satisfied|not\s+tenable|could\s+not\s+be\s+assumed)|"
                   r"\bviolat\w*\s+(?:\w+\s+){0,3}?sphericity|\bnon-?sphericity", re.I),
        re.compile(r"\bsphericity\s+(?:(?:assumption\s+)?(?:was|were|is|could\s+be)\s+)?(?:assumed|met|satisfied|tenable|upheld)\b|"
                   r"\b(?:the\s+)?assumption\s+of\s+sphericity\s+(?:was|is)\s+(?:met|satisfied|tenable|upheld)\b|\bsphericity\s+(?:held|holds)\b", re.I)),
    "bart": (
        re.compile(r"\b(?:was|were|is)\s+not\s+an\s+identity\s+matrix|\bdiffer\w*\s+(?:significantly\s+)?from\s+an\s+identity\s+matrix", re.I),
        re.compile(r"(?!x)x", re.I),
        re.compile(r"\b(?:suitable|appropriate|adequate|fit|good|apt|ideal)\s+for\s+(?:a\s+|the\s+)?(?:\w+\s+)?(?:factor|principal\s+components?)\b|\bfactorab\w+|"
                   r"\bcorrelations?\s+(?:were|was|are|is)\s+(?:sufficient|large\s+enough|adequate)|\bsufficient\s+correlations?", re.I)),
}
_AS_PV = re.compile(r"(?<![A-Za-z])[pP]\s*-?\s*values?\s*(?:\(\s*(?P<p>0?\.\d+|0|1(?:\.0+)?)\s*\)|(?:of|was|is)\s+(?P<q>0?\.\d+|0|1(?:\.0+)?)(?![\d.]*\d))")
_AS_SIG = re.compile(r"(?<![A-Za-z])sig(?:\.|\s+value|(?=\s*[=:])|(?=\s+0?\.\d))\s*(?:\((?:1|2)-tailed\)\s*)?(?:=|:|of|was|is)?\s*(?P<p>0?\.\d+|0|1(?:\.0+)?)(?![\d.]*\d)", re.I)


_AS_OF = re.compile(r"\b(variances?)\s+(?:of|in|across|between|among)\s+(?:the\s+)?(?:\w+\s+){0,2}?(?:groups?|conditions|classes|schools|"
                    r"levels|categories|samples|wards|sexes)\b", re.I)


def _as_said(fam, c):
    """What a clause says the test found: "holds", "fails" or "" (both, or neither)."""
    yes_neg, no, yes = _AS_CLAIM[fam]
    if fam == "var":                                   # (6.6, round 156: "the variances of the two groups were equal")
        c = _AS_OF.sub(r"\1", c)
    said = set()
    rest = c
    for rx, what in ((yes_neg, "holds"), (no, "fails")):
        for m in list(rx.finditer(rest)):
            if _AS_ASK.search(rest[max(0, m.start() - 40):m.start()]):
                return ""                              # ("to check whether the data were normally distributed")
            said.add(what)
            rest = rest[:m.start()] + " " * (m.end() - m.start()) + rest[m.end():]
    for m in yes.finditer(rest):
        before = rest[max(0, m.start() - 28):m.start()]
        if _AS_NEG.search(m.group(0)) or _AS_NEG.search(before) or _AS_ASK.search(rest[max(0, m.start() - 40):m.start()]):
            return ""                                  # (a "holds" under a "not" this rule doesn't know, or asked, not found)
        said.add("holds")
    return said.pop() if len(said) == 1 else ""


def _assume_notes(s):
    am = _ALPHA_SET.search(s)
    alpha = float(am.group("a")) if am and am.group("a") else (float(am.group("pc")) / 100 if am and am.group("pc") else 0.05)
    dec = max(2, len(repr(alpha).split(".")[1]) if "." in repr(alpha) else 2)
    a_str = f"{alpha:.{dec}f}".lstrip("0") if alpha < 1 else str(alpha)
    notes = []
    bounds = [0] + [m.end() for m in _CLAUSE.finditer(s)] + [len(s)]
    for a, z in zip(bounds, bounds[1:]):
        c = s[a:z]
        named = [(fam, show) for fam, rx, show in _AS_NAMES if rx.search(c)]
        if not named and _AS_BART.search(c):
            if _AS_FA.search(c) and not re.search(r"\bvariances?\b|\bhomogene", c, re.I):
                named = [("bart", "Bartlett’s test")]
            elif re.search(r"\bvariances?\b|\bhomogene", c, re.I) and not _AS_FA.search(c):
                named = [("var", "Bartlett’s test")]
        fams = {f for f, _ in named}
        if len(fams) != 1 or _AS_SKIP.search(c):
            continue
        fam = fams.pop()
        show = named[0][1] if len(named) == 1 else {"norm": "The tests of normality", "var": "The tests of equal variances"}.get(fam, named[0][1])
        if fam == "sph" and _AS_BART.search(c):
            continue
        ps = [(m.start(), m.end(), m.group("rel"), m.group("p"), f"p {m.group('rel')} {m.group('p')}") for m in _P.finditer(c)]
        ps += [(m.start(), m.end(), "=", m.group("p"), f"Sig. = {m.group('p')}") for m in _AS_SIG.finditer(c)
               if not any(x[0] <= m.start() < x[1] for x in ps)]
        ps += [(m.start(), m.end(), "=", m.group("p") or m.group("q"), f"p = {m.group('p') or m.group('q')}") for m in _AS_PV.finditer(c)
               if not any(x[0] <= m.start() < x[1] or m.start() <= x[0] < m.end() for x in ps)]
        if len(ps) != 1:
            continue
        p0, p1, rel, p_t, shown = ps[0]
        try:
            v = float(p_t)
        except ValueError:
            continue
        tiny = 0.5 * 10 ** -_dec(p_t)
        sig = (rel in ("<", "≤", "<=") and v <= alpha) or (rel == "=" and v < alpha - tiny)
        nonsig = (rel in (">", "≥", ">=") and v >= alpha) or (rel == "=" and v > alpha + tiny)
        if not (sig or nonsig):
            continue
        words = c[:p0] + " " + c[p1:]
        word = "nonsig" if _SIG_NOT.search(words) else "sig" if _SIG_YES.search(words) else ""
        if (word == "sig" and nonsig) or (word == "nonsig" and sig):
            continue                                   # (round 113 says it: the p and its significance word disagree)
        said = _as_said(fam, words)
        if not said:
            continue
        at = rel != "=" and abs(v - alpha) < 1e-12        # ("p < .05" itself: "a p below .05 says …")
        lo = f"a p below {a_str}" if at else f"{shown} is below {a_str}, which"
        hi = f"a p above {a_str}" if at else f"{shown} is above {a_str}, which"
        note = None
        if fam == "bart":
            if nonsig and said == "holds":
                note = (f"{show}: {f'a p above {a_str} means' if at else f'{shown} is above {a_str}, so'} the correlations may be too small for "
                        f"a factor analysis; the test supports one only when its p is below {a_str}.")
        elif sig and said == "holds":
            note = {"norm": f"{show}: {lo} says the data depart from a normal distribution, not that they are normal.",
                    "var": f"{show}: {lo} says the groups’ variances differ, not that they are equal.",
                    "fit": f"{show}: {lo} says the model does not fit the data well, not that it does.",
                    "sph": f"{show}: {lo} says sphericity is violated, not that it holds."}[fam]
        elif nonsig and said == "fails":
            note = {"norm": f"{show}: {hi} finds no departure from a normal distribution, though the sentence says the data aren’t normal.",
                    "var": f"{show}: {hi} finds no difference between the groups’ variances, though the sentence says they differ.",
                    "fit": f"{show}: {hi} finds no lack of fit, though the sentence says the model fits poorly.",
                    "sph": f"{show}: {hi} finds no violation of sphericity, though the sentence says it is violated."}[fam]
        if note:
            notes.append({"note": note, "kind": "value", "decision": True})
    return notes



# 6.6, round 158 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a change between two percentages and the size the
# sentence gives it. From 20% to 30% is 10 percentage points, or a rise of 50% (of the 20%): "an increase of 10%" says the
# second, "an increase of 50 percentage points" the first, and "an increase of 12%" neither. The change and its size are read
# in one sentence ("from 20% in 2019 to 30% in 2023, an increase of 10 percentage points"; "fell by 15% from 80% to 65%"); a
# range ("ranged from 20% to 30%"), an interval, a figure written "about" (then only its kind is held) and a sentence with two
# changes are left alone.
_PP_FT = re.compile(r"\bfrom\s+(?P<a>\d+(?:\.\d+)?)\s?(?:%|percent\b|per\s+cent\b)(?:\s+(?:in|at|during|for|among)\s+[^,;%]{1,30}?)?\s*(?:\([^)]{0,40}\)\s*)?"
                    r",?\s+to\s+(?P<b>\d+(?:\.\d+)?)\s?(?:%|percent\b|per\s+cent\b)", re.I)
_PP_BY = re.compile(r"\b(?:(?:an?\s+)?(?:increase|rise|decrease|decline|drop|reduction|fall|change|growth|gain|jump|dip|improvement)\s+of|"
                    r"(?:increased|rose|risen|grew|grown|went\s+up|gone\s+up|decreased|fell|fallen|dropped|declined|reduced|went\s+down|gone\s+down|changed|"
                    r"improved|worsened|climbed|jumped)\s+(?:\w+\s+){0,2}?by)\s+(?P<ab>about\s+|around\s+|approximately\s+|nearly\s+|almost\s+|some\s+|roughly\s+|over\s+|"
                    r"more\s+than\s+|less\s+than\s+)?(?P<c>\d+(?:\.\d+)?)\s?(?P<u>%|percent\b|per\s+cent\b|percentage[-\s]points?\b|points?\b|pp\b|p\.p\.)", re.I)
_PP_SKIP = re.compile(r"\brang(?:e|ed|es|ing)\b|\bvar(?:ied|ying|ies)\b|\binterval\b|(?<![A-Za-z])CI\b|\bbetween\s+\d", re.I)


def _points_notes(s):
    fts, bys = list(_PP_FT.finditer(s)), list(_PP_BY.finditer(s))
    if len(fts) != 1 or len(bys) != 1 or _PP_SKIP.search(s):
        return []
    ft, by = fts[0], bys[0]
    if by.start() >= ft.start() and by.end() <= ft.end():
        return []
    a, b, c = float(ft.group("a")), float(ft.group("b")), float(by.group("c"))
    ha, hb, hc = (0.5 * 10 ** -_dec(x) for x in (ft.group("a"), ft.group("b"), by.group("c")))
    if a == b:
        return []
    diff = abs(b - a)
    rough = bool(by.group("ab"))
    fits_d = abs(c - diff) <= ha + hb + hc + 1e-9 or (rough and abs(c - diff) <= max(1.0, 0.1 * diff))
    fits_r = False
    if a - ha > 0:
        lo = min(abs(bb - aa) / aa * 100 for aa in (a - ha, a + ha) for bb in (b - hb, b + hb))
        hi = max(abs(bb - aa) / aa * 100 for aa in (a - ha, a + ha) for bb in (b - hb, b + hb))
        fits_r = lo - hc - 1e-9 <= c <= hi + hc + 1e-9 or (rough and lo * 0.9 - 1 <= c <= hi * 1.1 + 1)
    rel = diff / a * 100 if a > 0 else None
    pts = by.group("u").lower().startswith(("percentage", "point", "pp", "p.p"))
    dd = max(_dec(ft.group("a")), _dec(ft.group("b")))
    d_t = f"{diff:.{dd}f}"
    r_t = f"{rel:.1f}".rstrip("0").rstrip(".") if rel is not None else None
    if fits_d and fits_r:
        return []
    head = f"From {ft.group('a')}% to {ft.group('b')}% is {d_t} percentage point{'s' if d_t != '1' else ''}"
    share = f", or {r_t}% of the {ft.group('a')}%" if r_t else ""
    if pts and not fits_d and fits_r:
        note = f"{head}; {by.group('c')} is the change as a share of the {ft.group('a')}% ({r_t}%), not in percentage points."
    elif not pts and fits_d and not fits_r:
        note = (f"{head}{share}: “{by.group('c')}%” would mean {by.group('c')}% of the {ft.group('a')}%; write “{by.group('c')} percentage "
                f"point{'s' if by.group('c') != '1' else ''}”.")
    elif (pts and fits_d) or (not pts and fits_r):
        return []
    elif rough:
        return []                                   # ("about 10%": only its kind is held)
    else:
        note = f"{head}{share}, not {by.group('c')}{'%' if not pts else ' percentage points'}."
    return [{"note": note, "kind": "value", "decision": False}]



# 6.6, round 159 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a median and its interquartile range. A median
# always lies between its first and third quartiles, and the quartiles within the lowest and highest values: "a median age of
# 21 years (IQR 23–27)" can't be, nor "IQR 20–30" beside "range 22–28". Read where the median is written with its IQR
# ("25 (IQR 20–30)", "median (IQR) = 25 (20–30)", "Mdn = 25, IQR = 20–30"); a mean beside an IQR is not judged (a skewed
# mean can lie outside it).
_IQ_NUM = r"\d+(?:\.\d+)?"
_IQ_TO = r"\s*(?:%|[a-zA-Z]{1,8}\.?)?\s*(?:–|—|-|to|,)\s*"
_IQ_A = re.compile(r"(?P<m>" + _IQ_NUM + r")\s*(?:%|[a-zA-Z]{1,8}\.?)?\s*[\(\[]\s*(?:IQR|inter-?quartile\s+range)\s*(?:=|:|of)?\s*(?P<q1>" + _IQ_NUM + r")"
                   + _IQ_TO + r"(?P<q3>" + _IQ_NUM + r")", re.I)
_IQ_B = re.compile(r"\bmedian\s*[\(\[]\s*(?:IQR|inter-?quartile\s+range)\s*[\)\]]\s*(?:\w+\s+){0,3}?(?:=|:|was|were|of|is)?\s*(?P<m>" + _IQ_NUM + r")\s*(?:%|[a-zA-Z]{1,8}\.?)?\s*"
                   r"[\(\[]\s*(?P<q1>" + _IQ_NUM + r")" + _IQ_TO + r"(?P<q3>" + _IQ_NUM + r")\s*(?:%|[a-zA-Z]{1,8}\.?)?\s*[\)\]]", re.I)
_IQ_C = re.compile(r"(?:\bmedian(?:\s+(?:age|score|value|duration|time|number))?|(?<![A-Za-z])Mdn)\s*(?:=|:|was|of|is)\s*(?P<m>" + _IQ_NUM + r")\s*(?:%|[a-zA-Z]{1,8}\.?)?\s*"
                   r"[,;(]\s*(?:and\s+)?(?:an?\s+|the\s+)?(?:IQR|inter-?quartile\s+range)\s*(?:=|:|of|was)?\s*(?P<q1>" + _IQ_NUM + r")" + _IQ_TO + r"(?P<q3>" + _IQ_NUM + r")", re.I)


def _iqr_notes(s):
    found, notes = [], []
    for rx in (_IQ_B, _IQ_C, _IQ_A):
        for m in rx.finditer(s):
            if any(a <= m.start("q1") < z for a, z, *_ in found):
                continue
            if rx is _IQ_A and not re.search(r"\bmedian\b|(?<![A-Za-z])Mdn\b", s[max(0, m.start() - 80):m.start()], re.I):
                continue                               # ("25 (IQR …)" only where the 25 is a median)
            if rx is _IQ_A and re.search(r"\bmean\b|(?<![A-Za-z])M\s*=", s[max(0, m.start() - 25):m.start("q1")], re.I):
                continue
            found.append((m.start("q1"), m.end("q3"), m.group("m"), m.group("q1"), m.group("q3")))
    rms = [m for m in _RANGE_FT.finditer(s) if not any(m.start() <= x[0] < m.end() + 1 for x in found)]
    rm = rms[0] if len(rms) == 1 else None             # (one range besides the IQR: "range 18–27")
    for a, z, m_t, q1_t, q3_t in found:
        med, q1, q3 = float(m_t), float(q1_t), float(q3_t)
        h = 0.5 * 10 ** -max(_dec(m_t), _dec(q1_t), _dec(q3_t))
        if q1 > q3 + h:
            notes.append({"note": f"An interquartile range runs from the first quartile up to the third: “{q1_t} to {q3_t}” is the other way round.",
                          "kind": "value", "decision": False})
            continue
        if med < q1 - h or med > q3 + h:
            notes.append({"note": f"The median ({m_t}) lies outside its interquartile range ({q1_t} to {q3_t}): a median is always between its first "
                                  f"and third quartiles.", "kind": "value", "decision": False})
            continue
        if rm:
            try:
                lo, hi = sorted((float(rm.group("a")), float(rm.group("b"))))
            except ValueError:
                continue
            if lo != hi and not (1900 <= lo <= 2100 and 1900 <= hi <= 2100) and (q1 < lo - h or q3 > hi + h):
                notes.append({"note": f"The interquartile range ({q1_t} to {q3_t}) runs outside the range ({rm.group('a')} to {rm.group('b')}): the "
                                      f"quartiles lie within the lowest and highest values.", "kind": "value", "decision": False})
    return notes



# 6.6, round 160 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): the words a sentence gives five statistics, held to
# their usual scales (round 95 did the effect sizes). Cohen's kappa (Landis and Koch: .21 fair, .41 moderate, .61 substantial,
# .81 almost perfect; Altman's "good" from .61, Fleiss's "fair to good" .40 to .75, "excellent" above .75); the KMO (Kaiser:
# .50 miserable, .60 mediocre, .70 middling, .80 meritorious, .90 marvellous, below .50 unacceptable); a VIF or a tolerance
# (10 or more, or a tolerance of .10 or less, is multicollinearity by every rule; below 2.5, or a tolerance above .40, none by
# any); the Durbin–Watson statistic (2 means no autocorrelation; below 1 or above 3 is cause for concern); the area under a ROC
# curve (Hosmer and Lemeshow: .70 acceptable, .80 excellent, .90 outstanding). A word is told only where no common scale gives
# it to the figure; "not" before the word turns it round.
_SW_NOT = re.compile(r"\b(?:not|no|never|hardly|lacked|without)\s+(?:\w+\s+){0,2}$", re.I)
_SW_V = r"(?P<v>-?0?\.\d+|1(?:\.0+)?|0)"
_SW_KAPPA = re.compile(r"(?:\b(?:cohen['’]?s\s+|fleiss['’]?\s?s?\s+)?kappa\b|κ)\s*(?:coefficient\s*|value\s*|statistic\s*)?(?:\(\s*κ\s*\)\s*)?(?:=|of|was|:)\s*" + _SW_V, re.I)
_SW_KWORD = re.compile(r"\b(?P<w>poor|slight|fair|moderate|substantial|good|very\s+good|excellent|almost\s+perfect|perfect|strong|weak)\s+"
                       r"(?:inter-?rater\s+|intra-?rater\s+|inter-?observer\s+|level\s+of\s+|degree\s+of\s+)?(?:agreement|reliability|concordance)\b|"
                       r"\b(?:agreement|reliability|concordance)\s+(?:was|were|is)\s+(?:(?!almost\b|very\b|not\b|no\b|never\b)\w+\s+)?(?P<w2>poor|slight|fair|moderate|substantial|good|very\s+good|"
                       r"excellent|almost\s+perfect|perfect|strong|weak)\b", re.I)
_SW_KOK = {"poor": (-1, .40), "slight": (-1, .21), "fair": (.19, .76), "moderate": (.39, .81), "substantial": (.59, .81), "good": (.39, .81),
           "very good": (.79, 1.01), "excellent": (.74, 1.01), "almost perfect": (.79, 1.01), "perfect": (.995, 1.01), "strong": (.59, 1.01),
           "weak": (-1, .61)}
_SW_KMO = re.compile(r"(?:\bKMO\b|Kaiser[-–\s]*Meyer[-–\s]*Olkin)(?:\s+(?:measure|value|statistic|test|index))?(?:\s+of\s+sampling\s+adequacy)?\s*"
                     r"(?:\(\s*KMO\s*\)\s*)?(?:=|of|was|is|:)\s*(?P<v>0?\.\d+|1(?:\.0+)?)", re.I)
_SW_MWORD = re.compile(r"\b(?P<w>marvell?ous|meritorious|middling|mediocre|miserable|unacceptable|inadequate|insufficient|unsuitable|superb|great|"
                       r"excellent|good|adequate|acceptable|sufficient|suitable|appropriate|satisfactory)\b", re.I)
_SW_MOK = {"marvellous": (.89, 1.01), "marvelous": (.89, 1.01), "superb": (.89, 1.01), "meritorious": (.79, 1.01), "great": (.79, 1.01),
           "excellent": (.79, 1.01), "good": (.69, 1.01), "middling": (.69, .81), "mediocre": (.59, .71), "miserable": (.49, .61),
           "unacceptable": (-1, .51), "inadequate": (-1, .61), "insufficient": (-1, .61), "unsuitable": (-1, .61),
           "adequate": (.49, 1.01), "acceptable": (.49, 1.01), "sufficient": (.49, 1.01), "suitable": (.49, 1.01), "appropriate": (.49, 1.01),
           "satisfactory": (.49, 1.01)}
_SW_NEGOF = {"adequate": "inadequate", "acceptable": "unacceptable", "sufficient": "insufficient", "suitable": "unsuitable",
             "appropriate": "unsuitable", "satisfactory": "inadequate", "good": "inadequate"}
_SW_VIF = re.compile(r"\b(?:VIFs?|variance\s+inflation\s+factors?)(?:\s*\(\s*VIFs?\s*\))?\s*(?:values?\s*)?(?:ranged\s+from|of|=|was|were|:|between)\s*"
                     r"(?P<v>\d+(?:\.\d+)?)(?:\s*(?:to|and|–|-)\s*(?P<v2>\d+(?:\.\d+)?))?", re.I)
_SW_TOL = re.compile(r"\btolerances?(?:\s+values?)?\s*(?:ranged\s+from|of|=|was|were|:|between)\s*(?P<v>0?\.\d+|1(?:\.0+)?)(?:\s*(?:to|and|–|-)\s*(?P<v2>0?\.\d+|1(?:\.0+)?))?", re.I)
_SW_MCOK = re.compile(r"\b(?:no|absence\s+of|without|free\s+(?:of|from)|not\s+(?:a\s+)?(?:problem|concern|an\s+issue))\b[^.;]{0,40}?\bmulti-?collinearity\b|"
                      r"\bmulti-?collinearity\s+(?:was|is|were)\s+(?:not\s+(?:a\s+)?(?:problem|concern|present|an\s+issue|detected|found)|absent|"
                      r"(?:of\s+)?no\s+concern|unlikely|ruled\s+out)", re.I)
_SW_MCBAD = re.compile(r"\bmulti-?collinearity\s+(?:was|is|were)\s+(?:(?!not\b|no\b|never\b)\w+\s+)?(?:present|a\s+problem|a\s+concern|detected|found|high|serious|severe)\b|"
                       r"\b(?:high|serious|severe|strong)\s+multi-?collinearity\b|\bshowed\s+multi-?collinearity\b", re.I)
_SW_DW = re.compile(r"(?:\bdurbin[-–\s]*watson(?:\s+(?:statistic|test|value|d))?|(?<![A-Za-z])DW)\s*(?:\(\s*DW\s*\)\s*)?(?:=|of|was|is|:)\s*(?P<v>\d(?:\.\d+)?)", re.I)
_SW_DWOK = re.compile(r"\bno\s+(?:\w+\s+)?(?:auto-?correlation|serial\s+correlation)\b|\bindependen(?:t|ce)\s+(?:of\s+)?(?:the\s+)?(?:errors|residuals|observations)\b|"
                      r"\b(?:errors|residuals|observations)\s+(?:were|was|are)\s+independent\b|\bauto-?correlation\s+(?:was|is)\s+(?:not\s+(?:present|a\s+problem|a\s+concern)|absent)", re.I)
_SW_DWBAD = re.compile(r"\b(?:positive|negative|serious|strong)?\s*auto-?correlation\s+(?:was|is|were)\s+(?:(?!not\b|no\b|never\b)\w+\s+)?(?:present|detected|found|a\s+problem|a\s+concern)\b|"
                       r"\b(?:errors|residuals)\s+(?:were|was|are)\s+not\s+independent\b|\bshowed\s+(?:\w+\s+)?auto-?correlation\b", re.I)
_SW_AUC = re.compile(r"(?:\bAUC\b|\barea\s+under\s+(?:the\s+)?(?:ROC\s+|receiver\s+operating\s+characteristic\s+)?curve|(?<![A-Za-z])c[-\s]?statistic)\s*"
                     r"(?:\(\s*AUC\s*\)\s*)?(?:=|of|was|is|:)\s*(?P<v>0?\.\d+|1(?:\.0+)?)", re.I)
_SW_AWORD = re.compile(r"\b(?P<w>outstanding|excellent|very\s+good|good|acceptable|adequate|reasonable|fair|poor|weak|strong|no)\s+(?:\w+\s+)?"
                       r"(?:discrimina\w*|discriminatory|predictive\s+(?:ability|power|accuracy)|accuracy|diagnostic\s+accuracy)\b", re.I)
_SW_AOK = {"outstanding": (.89, 1.01), "excellent": (.79, 1.01), "very good": (.79, 1.01), "good": (.69, 1.01), "acceptable": (.69, 1.01),
           "adequate": (.69, 1.01), "reasonable": (.69, 1.01), "strong": (.79, 1.01), "fair": (.59, .81), "poor": (-1, .71), "weak": (-1, .71),
           "no": (-1, .56)}


def _sw_band(v, table):
    for top, name in table:
        if v < top:
            return name
    return table[-1][1]


def _scale_word_notes(s):
    notes = []
    bounds = [0] + [m.end() for m in _CLAUSE.finditer(s)] + [len(s)]
    for a, z in zip(bounds, bounds[1:]):
        c = s[a:z]
        # Cohen's kappa
        km, kw = list(_SW_KAPPA.finditer(c)), list(_SW_KWORD.finditer(c))
        if len(km) == 1 and len(kw) == 1:
            w = (kw[0].group("w") or kw[0].group("w2")).lower()
            w = re.sub(r"\s+", " ", w)
            v = float(km[0].group("v"))
            if not _SW_NOT.search(c[:kw[0].start()]) and w in _SW_KOK and not (_SW_KOK[w][0] <= v < _SW_KOK[w][1]):
                lk = _sw_band(v, [(0, "poor"), (.205, "slight"), (.405, "fair"), (.605, "moderate"), (.805, "substantial"), (9, "almost perfect")])
                notes.append({"note": f"κ = {km[0].group('v')} is {lk} agreement by Landis and Koch (.21 fair, .41 moderate, .61 substantial, .81 almost "
                                      f"perfect), not {w}.", "kind": "value", "decision": False})
        # the KMO
        mm = list(_SW_KMO.finditer(c))
        if len(mm) == 1:
            v = float(mm[0].group("v"))
            ws = [x for x in _SW_MWORD.finditer(c) if not (mm[0].start() <= x.start() < mm[0].end())]
            if len(ws) == 1:
                w = ws[0].group("w").lower()
                if _SW_NOT.search(c[:ws[0].start()]):
                    w = _SW_NEGOF.get(w, "")
                if w in _SW_MOK and not (_SW_MOK[w][0] <= v < _SW_MOK[w][1]):
                    kz = _sw_band(v, [(.495, "unacceptable"), (.595, "miserable"), (.695, "mediocre"), (.795, "middling"), (.895, "meritorious"), (9, "marvellous")])
                    said = ws[0].group("w").lower()
                    notes.append({"note": f"A KMO of {mm[0].group('v')} is {kz} in Kaiser’s terms (.50 miserable, .60 mediocre, .70 middling, .80 meritorious, "
                                          f".90 marvellous; below .50 unacceptable), not {('not ' if w != said else '') + said}.", "kind": "value", "decision": False})
        # a VIF or a tolerance
        ok = _SW_MCOK.search(c)
        bad = next((m for m in _SW_MCBAD.finditer(c) if not re.search(r"\b(?:no|not|without|absence\s+of)\s+$", c[max(0, m.start() - 15):m.start()], re.I)), None)
        if bool(ok) != bool(bad):
            vifs = [float(x) for m in _SW_VIF.finditer(c) for x in (m.group("v"), m.group("v2")) if x]
            tols = [float(x) for m in _SW_TOL.finditer(c) for x in (m.group("v"), m.group("v2")) if x]
            if ok and vifs and max(vifs) >= 10:
                notes.append({"note": f"A VIF of {max(vifs):g} means serious multicollinearity (10 or more by the usual rule, 5 by a stricter one), "
                                      f"not none.", "kind": "value", "decision": False})
            elif ok and tols and min(tols) <= 0.1:
                notes.append({"note": f"A tolerance of {min(tols):g} (a VIF of {1 / min(tols):.1f}) means serious multicollinearity (a tolerance of .10 "
                                      f"or less by the usual rule), not none.", "kind": "value", "decision": False})
            elif bad and vifs and max(vifs) < 2.5:
                notes.append({"note": f"VIFs of {max(vifs):g} or less are well below 5 and 10, the usual marks of multicollinearity: they show none.",
                              "kind": "value", "decision": False})
            elif bad and tols and min(tols) > 0.4:
                notes.append({"note": f"Tolerances of {min(tols):g} or more (VIFs of {1 / min(tols):.1f} or less) show no multicollinearity by the "
                                      f"usual rules.", "kind": "value", "decision": False})
        # the Durbin–Watson statistic
        dm = list(_SW_DW.finditer(c))
        if len(dm) == 1:
            v = float(dm[0].group("v"))
            ok = _SW_DWOK.search(c)
            bad = next((m for m in _SW_DWBAD.finditer(c) if not re.search(r"\b(?:no|not|without|absence\s+of)\s+$", c[max(0, m.start() - 15):m.start()], re.I)), None)
            if ok and not bad and (v < 1 or v > 3):
                notes.append({"note": f"A Durbin–Watson of {dm[0].group('v')} points to {'positively' if v < 2 else 'negatively'} correlated residuals "
                                      f"(below 1 or above 3 is the usual cause for concern; 2 means none), not independent ones.", "kind": "value", "decision": False})
            elif bad and not ok and 1.5 <= v <= 2.5:
                notes.append({"note": f"A Durbin–Watson of {dm[0].group('v')} is close to 2, which means no autocorrelation of the residuals.",
                              "kind": "value", "decision": False})
        # the area under a ROC curve
        am, aw = list(_SW_AUC.finditer(c)), list(_SW_AWORD.finditer(c))
        if len(am) == 1 and len(aw) == 1:
            w = re.sub(r"\s+", " ", aw[0].group("w").lower())
            v = float(am[0].group("v"))
            if not _SW_NOT.search(c[:aw[0].start()]) and w in _SW_AOK and not (_SW_AOK[w][0] <= v < _SW_AOK[w][1]):
                hl = _sw_band(v, [(.55, "no discrimination"), (.695, "poor"), (.795, "acceptable"), (.895, "excellent"), (9, "outstanding")])
                notes.append({"note": f"An AUC of {am[0].group('v')} is {hl} by Hosmer and Lemeshow’s scale (.70 acceptable, .80 excellent, .90 outstanding), "
                                      f"not {w}.", "kind": "value", "decision": False})
    return notes



# 6.6, round 161 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): what a test can say. A chi-square test (or Fisher's
# exact test) shows whether two things are associated, not which way: "a significant positive relationship between age and
# self-medication (χ² = 12.30, p = .002)" gives it a direction it can't show (a correlation, an odds ratio, a test for trend or
# the table's own percentages can). And a correlation shows that two things vary together, not that one causes the other:
# "r = .45, showing that stress causes poor sleep". A clause that also gives a statistic with a direction (r, rho, an odds
# ratio, gamma, tau, a coefficient, a linear-by-linear or trend test), or hedges ("may cause"), is left alone. The person's
# writing only (RamzAI's own claims are held to their papers by valid65's cause rule).
_DR_CHI = re.compile(r"(?:χ|chi)\s*(?:²|2|-?\s*square[d]?)\s*(?:\([^)]{0,30}\))?\s*(?:value\s+)?(?:=|of|was)\s*\d|\bfisher['’]?s?\s+exact\b", re.I)
_DR_DIRWORD = re.compile(r"\b(?P<w>positive|negative|direct|inverse|positively|negatively|inversely)\s+(?:and\s+)?(?:(?:statistically\s+)?significant(?:ly)?\s+|"
                         r"strong(?:ly)?\s+|weak(?:ly)?\s+|moderate(?:ly)?\s+)?(?:relationship|association|correlation|link|associated|related|correlated|linked)\b|"
                         r"\b(?:relationship|association|correlation)\s+(?:was|is)\s+(?:\w+\s+)?(?P<w2>positive|negative|inverse)\b", re.I)
_DR_HASDIR = re.compile(r"(?<![A-Za-z])r\s*(?:\(\s*\d+\s*\))?\s*=|\brho\b|ρ|(?<![A-Za-z])r\s?s\s*=|\b(?:a|c)?OR\s*=|\bAOR\b|\bCOR\b|\bodds\s+ratio|"
                         r"\bgamma\b|γ|\btau\b|τ|β|\bbeta\b|(?<![A-Za-z])B\s*=|\blinear[-\s]by[-\s]linear\b|\b(?:test\s+for\s+|for\s+)?trend\b|\bmantel\b|"
                         r"\bcorrelation\s+coefficient\b|\bspearman\b|\bpearson['’]?s?\s+r\b|\bkendall\b", re.I)
_DR_CORR = re.compile(r"(?<![A-Za-z])r\s*(?:\(\s*\d+\s*\))?\s*=\s*[-−–]?\s?\.?\d|\brho\s*=|ρ\s*=|(?<![A-Za-z])r\s?s\s*=|\bcorrelation\s+coefficient\b|"
                      r"\bpearson['’]?s?\s+(?:r\b|correlation)|\bspearman", re.I)
_DR_CAUSE = re.compile(r"\b(?P<v>causes|caused|cause\s+of|leads?\s+to|led\s+to|results?\s+in|resulted\s+in|determines|determined|"
                       r"is\s+responsible\s+for|was\s+responsible\s+for|produces|produced)\b", re.I)
_DR_SOFT = re.compile(r"\b(?:not|no|never|may|might|could|possibly|perhaps|probably|likely|suggest\w*|whether|if)\b|n['’]t\b", re.I)
_DR_DESIGN = re.compile(r"\b(?:regression|experiment\w*|randomi[sz]ed|trial|intervention|manipulat\w*|longitudinal|cohort|over\s+time)\b", re.I)


def _dir_notes(s):
    notes = []
    bounds = [0] + [m.end() for m in _CLAUSE.finditer(s)] + [len(s)]
    for a, z in zip(bounds, bounds[1:]):
        c = s[a:z]
        dm = _DR_DIRWORD.search(c)
        if dm and _DR_CHI.search(c) and not _DR_HASDIR.search(s) and not re.search(r"\b(?:not|no|nor|without)\s+(?:\w+\s+){0,2}$", c[:dm.start()], re.I):
            w = (dm.group("w") or dm.group("w2")).lower()
            test = "Fisher’s exact test" if re.search(r"fisher", c, re.I) and not re.search(r"χ|chi", c, re.I) else "A chi-square test"
            notes.append({"note": f"{test} shows whether two things are associated, not which way: “{w}” needs a correlation, an odds ratio or a test "
                                  f"for trend (or say which group had more, from the table).", "kind": "value", "decision": False})
            continue
        cm = _DR_CAUSE.search(c)
        if cm and _DR_CORR.search(c) and not _DR_DESIGN.search(c) and not _DR_SOFT.search(c[max(0, cm.start() - 30):cm.start()]):
            notes.append({"note": f"A correlation shows that two things vary together, not that one causes the other: “{cm.group('v').lower()}” "
                                  f"says more than it can.", "kind": "value", "decision": False})
    return notes



# 6.6, round 162 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): figures and readings no result can have. A ratio
# (an odds ratio, AOR, RR, HR, Exp(B)…) is never negative: "AOR = −1.20" is the coefficient B written as the ratio (its ratio
# is e^B = 0.30). "p < .000" (round 8's range check): no p is below 0. A result that isn't significant shows that no difference
# was found, not that there is none: "no significant difference (p = .34), which proves that males and females have the same
# knowledge" says more than the test can. And a negative corrected item-total correlation (in the writing or SPSS's Item-Total
# Statistics) means an item runs against the rest of its scale: check whether it should be reverse-scored.
_NG_RATIO = re.compile(r"(?<![A-Za-z])(?P<lab>[aAcC]?OR|AOR|COR|aRR|RR|aHR|HR|IRR|PR|Exp\s*\(\s*B\s*\)|odds\s+ratio|risk\s+ratio|hazard\s+ratio|"
                       r"relative\s+risk)\s*(?:[:=]\s*|of\s+|was\s+)(?P<v>[-−–]\s?\d*\.?\d+)")
_NG_PROOF = re.compile(r"\b(?P<w>prove[sd]?|proof|confirm(?:s|ed)?|establish(?:es|ed)?|demonstrat(?:es|ed))\b\s+(?:that\s+)?(?:there\s+(?:is|was|were)\s+no\b|"
                       r"[^.;,]{0,40}?\b(?:the\s+same|identical|equal|no\s+(?:difference|relationship|association|effect|influence)))", re.I)
_NG_SAME = re.compile(r"\b(?:were|was|are|is)\s+(?:exactly\s+|essentially\s+|therefore\s+|thus\s+)?(?P<w>the\s+same|identical)\b", re.I)
_NG_ASSUME = re.compile(r"\bvariances?\b|\blevene|\bnormal|\bsphericity|\bhosmer|\bshapiro|\bkolmogorov", re.I)
_NG_ITC = re.compile(r"corrected\s+item[-–\s]*(?:to[-–\s]*)?total\s+correlations?\s*(?:of|=|was|:)?\s*(?P<v>[-−–]\s?0?\.\d+)", re.I)


def _null_notes(s):
    notes = []
    for m in _NG_RATIO.finditer(s):
        try:
            v = -float(re.sub(r"^[-−–]\s?", "", m.group("v")))
        except ValueError:
            continue
        lab = re.sub(r"\s+", " ", m.group("lab"))
        e = math.exp(v) if v > -30 else 0.0
        notes.append({"note": f"{lab} is never negative: {m.group('v').replace('-', '−')} looks like the coefficient B written as the ratio "
                              f"(its ratio would be e^B = {e:.2f}).", "kind": "value", "decision": False})
    bounds = [0] + [m.end() for m in _CLAUSE.finditer(s)] + [len(s)]
    if not _NG_ASSUME.search(s):
        for a, z in zip(bounds, bounds[1:]):
            c = s[a:z]
            nonsig = bool(_SIG_NOT.search(c)) or any(
                (m.group("rel") in (">", "≥", ">=") and float(m.group("p")) >= 0.05) or (m.group("rel") == "=" and float(m.group("p")) > 0.055)
                for m in _P.finditer(c) if re.fullmatch(r"\d*\.?\d+", m.group("p")))
            if not nonsig:
                continue                              # (the claim and the result it leans on: one clause)
            m = _NG_PROOF.search(c) or _NG_SAME.search(c)
            if m and not re.search(r"\b(?:not|no|never)\s+(?:\w+\s+){0,2}$", c[:m.start()], re.I):
                w = re.sub(r"\s+", " ", m.group("w")).lower()
                notes.append({"note": f"A result that isn’t significant shows that no difference was found, not that there is none: “{w}” says more "
                                      f"than the test can (a confidence interval shows how large a difference may still be).", "kind": "value",
                              "decision": False})
                break
    for m in _NG_ITC.finditer(s):
        notes.append({"note": f"A negative corrected item-total correlation ({m.group('v').replace('-', '−')}) means the item runs against the rest of "
                              f"its scale: check whether it should be reverse-scored, then work out the alpha again.", "kind": "value", "decision": False})
    return notes


_NG_LTCELL = re.compile(r"^(?:p\s*)?[<≤]\s*0?\.0+$|^(?:p\s*)?[<≤]\s*0$", re.I)


def null_table_check(text):
    """Round 162 in tables: a p cell written "<.000", and SPSS's Item-Total Statistics with a negative Corrected Item-Total
    Correlation."""
    out, lines = [], str(text or "").split("\n")
    seen = set()
    for i, line in enumerate(lines):
        cells = _cells_keep(line)
        if len([c for c in cells if (c or "").strip()]) >= 2 and any(_NG_LTCELL.match((c or "").strip()) for c in cells[1:]):
            if "lt" not in seen:
                seen.add("lt")
                out.append({"s": line.strip()[:500], "note": "No p-value is below 0: SPSS’s “.000” means p < .001, so write <.001 in the table.",
                            "kind": "table", "decision": False})
        low = [(c or "").strip().lower() for c in cells]
        k = next((j for j, c in enumerate(low) if re.match(r"^corrected\s+item[-–\s]*total\s+correlation", c)), None)
        if k is None:
            continue
        neg = []
        for j in range(i + 1, min(len(lines), i + 60)):
            cs = _cells_keep(lines[j])
            if not any((c or "").strip() for c in cs):
                if neg or j > i + 1:
                    break
                continue
            if k < len(cs):
                v = _num_cell((cs[k] or "").replace("−", "-"))
                if v is not None and v < 0:
                    neg.append(((cs[0] or "").strip() or f"row {j - i}", (cs[k] or "").strip()))
        if neg:
            items = ", ".join(f"{lab} ({v.replace('-', '−')})" for lab, v in neg[:4]) + (" …" if len(neg) > 4 else "")
            out.append({"s": lines[i].strip()[:500], "kind": "table", "decision": False,
                        "note": f"Item-Total Statistics: {items} {'has a' if len(neg) == 1 else 'have'} negative corrected item-total "
                                f"correlation{'s' if len(neg) > 1 else ''}, so {'it runs' if len(neg) == 1 else 'they run'} against the rest of the "
                                f"scale: check whether {'it' if len(neg) == 1 else 'they'} should be reverse-scored, then work out the alpha again."})
    return out


# 6.6, round 114 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a mean (or a median) lies inside the range it
# describes, and an SD can't be more than the range allows. "Ages ranged from 18 to 25 years (M = 27.3)" can't be; nor can
# "18 to 25 (SD = 5.1)": values within 18 and 25 have an SD of at most half the range (a little more for a small sample:
# √(n/(n−1)) times it; with no n given, two values' worst case).
_RANGE_FT = re.compile(r"\b(?:rang(?:ed|ing|e)|aged|between)\s*(?:from\s+|between\s+|of\s+|:\s*)?(?P<a>\d+(?:\.\d+)?)\s*(?P<pa>%)?\s*(?:to|and|[–-])\s*"
                       r"(?P<b>\d+(?:\.\d+)?)\s*(?P<pb>%)?(?!\s*(?:years?\s+)?(?:old\s+)?(?:at|in)\s+(?:19|20)\d{2})", re.I)
_MEAN_V = re.compile(r"(?:(?<![A-Za-z])M\s*=\s*|\b(?:pooled\s+)?mean(?:\s+(?:age|score|value|prevalence))?\s*(?:of|was|=|:|\()?\s*(?:=\s*)?)(?P<v>\d+(?:\.\d+)?)(?P<pc>\s*%)?", re.I)
_MEDIAN_V = re.compile(r"(?:(?<![A-Za-z])Mdn\s*=\s*|\bmedian(?:\s+(?:age|score|value))?\s*(?:of|was|=|:|\()\s*(?:=\s*)?)(?P<v>\d+(?:\.\d+)?)", re.I)
_SD_V = re.compile(r"(?<![A-Za-z])(?:SD|s\.d\.|standard\s+deviation)\s*(?:of|=|:|was)?\s*(?:=\s*)?(?P<v>\d+(?:\.\d+)?)", re.I)
_N_V = re.compile(r"(?<![A-Za-z])[nN]\s*=\s*(?P<n>\d[\d,]*)")


def _minmax_notes(s):
    rm = next((m for m in _RANGE_FT.finditer(s)                    # (6.6, round 159: an interquartile range is not the range)
               if not re.search(r"inter-?quartile\s*$", s[max(0, m.start() - 16):m.start()], re.I)), None)
    if not rm:
        return []
    try:
        a, b = float(rm.group("a")), float(rm.group("b"))
    except ValueError:
        return []
    if a > b:
        a, b = b, a
    if a == b or (1900 <= a <= 2100 and 1900 <= b <= 2100):
        return []                          # (years: "from 2015 to 2021" is no range of values)
    notes = []
    pct = bool(rm.group("pa") or rm.group("pb"))
    ages = rm.group(0).lower().startswith("aged") or bool(re.match(r"\s*years?\b", s[rm.end():rm.end() + 12], re.I))
    for rx, what in ((_MEAN_V, "mean"), (_MEDIAN_V, "median")):
        for m in rx.finditer(s):
            if (bool(m.groupdict().get("pc")) if rx is _MEAN_V else False) != pct and rx is _MEAN_V:
                continue                       # (a mean of a percentage against a range of percentages, and nothing else)
            if ages and not ("age" in m.group(0).lower() or re.match(r"\s*years?\b", s[m.end():m.end() + 10], re.I)
                             or (m.group(0).strip().startswith("M") and 0 <= m.start() - rm.end() <= 60)):
                continue                       # (a range of ages and a mean of something else: "aged 18–25 … a mean of 27.3 hours")
            v = float(m.group("v"))
            if v < a - 0.05 or v > b + 0.05:
                u = "%" if pct else ""
                notes.append({"note": f"The {what} ({m.group('v')}{u}) lies outside the range it describes ({rm.group('a')}{u} to {rm.group('b')}{u}).",
                              "kind": "value", "decision": False})
                return notes
    sd = _SD_V.search(s)
    if sd:
        v = float(sd.group("v"))
        n = None
        nm = _N_V.search(s)
        if nm:
            try:
                n = int(nm.group("n").replace(",", ""))
            except ValueError:
                n = None
        k = (n / (n - 1)) ** 0.5 if n and n > 1 else 2 ** 0.5
        top = (b - a) / 2 * k
        if v > top + 0.5 * 10 ** -_dec(sd.group("v")) + 1e-9:
            notes.append({"note": f"An SD of {sd.group('v')} is more than values between {rm.group('a')} and {rm.group('b')} can have (at most "
                                  f"{top:.2f}{' with n = ' + str(n) if n else ''}).", "kind": "value", "decision": False})
    return notes


# 6.6, round 117 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a frequency table in the person's writing, each
# row's percentage worked out again from its count and the table's total. A results chapter pasted from Word comes as rows of
# cells (tabs, wide spaces or "|"): "Male  56  37.3", …, "Total  150  100.0". A row whose percentage isn't its count of the
# total, or rows whose counts don't add up to the total the table gives, are said; a table with no total is not judged.
_CELL_SPLIT = re.compile(r"\t+|\s{2,}|\s*\|\s*")
_INT_CELL = re.compile(r"^\d{1,6}$")
_PCT_CELL = re.compile(r"^(?P<p>\d{1,3}(?:\.\d{1,2})?)\s*%?$")


_STAT_HEAD = re.compile(r"^(?:mean|std\.?\s*deviation|s\.?d\.?|std\.?\s*error(?:\s+mean)?|minimum|maximum|median|variance|range|"
                        r"lower(?:\s+bound)?|upper(?:\s+bound)?|sig\.?(?:\s*\(2-tailed\))?|df|t|f|b|beta|wald|exp\s*\(\s*b\s*\)|correlation|"
                        r"sum\s+of\s+squares|mean\s+square|95%\s+confidence\s+interval.*)$", re.I)


def _table_rows(text):
    """[(rows, whole, n)]: the frequency tables of the text, row by row ((label, count, percent string, line)); whole is False
    when the line before the table looks like a row that couldn't be read (then its rows may not all be there); n is the
    total its caption gives ("(N = 150)") within the three lines above it, or None."""
    blocks, cur, before, last, recent, cap = [], [], "", "", [], None
    skip = set()                                 # 6.6, round 129: SPSS's Frequencies tables are read by spss_freq_check
    for h, _cols, rows_ in _spss_freq_tables(text):
        skip.add(h)
        skip.update(j for j, _l, _c in rows_)

    def close():
        if len(cur) >= 2:
            whole = not (re.search(r"\d", before) and not re.search(r"\btable\b|\bN\s*=|\bn\s*=", before, re.I))
            blocks.append((list(cur), whole, cap))
    zone = False                                 # 6.6, round 136: under a heading of statistics (Mean, Std. Deviation, Sig., df…)
    for idx, line in enumerate(str(text or "").split("\n")):
        cells = [c.strip() for c in _CELL_SPLIT.split(line.strip()) if c.strip()]
        n_stat = sum(1 for c in cells if _STAT_HEAD.match(c))
        if not line.strip():
            zone = False
        elif n_stat >= 2:
            zone = True                          # (no row under it is a frequency row: a mean is no percentage)
        elif zone and not n_stat and not re.search(r"\d", line):
            zone = False                         # (a title: the next table)
        row = None
        many = (sum(1 for c in cells[1:] if re.match(r"^\d+(?:\.\d+)?%?$", c)) >= 4      # 6.6, round 132: four counts or more
                and sum(1 for c in cells[1:] if re.match(r"^\d+$", c)) >= 3)                  # (a rating scale's columns) are no
        if idx not in skip and not zone and not many and len(cells) >= 3 and (re.search(r"[A-Za-z]", cells[0]) or re.match(r"^\d{1,3}\s*(?:[-–]|to)\s*\d{1,3}$|^[<>≤≥]\s*\d", cells[0])) \
                and not re.search(r"\d{4,}", cells[0]):
            if _INT_CELL.match(cells[1]) and _PCT_CELL.match(cells[2]):
                row = (cells[0], int(cells[1]), _PCT_CELL.match(cells[2]).group("p"), line.strip())
        if row:
            if not cur:
                before = last
                nm = next((re.search(r"\bN\s*=\s*(\d[\d,]*)", x) for x in reversed(recent[-3:]) if re.search(r"\bN\s*=\s*\d", x)), None)
                cap = int(nm.group(1).replace(",", "")) if nm else None
            cur.append(row)
        else:
            close()
            cur = []
        if line.strip():
            last = line.strip()
            recent.append(last)
    close()
    return blocks


def table_check(text):
    """Notes on the frequency tables: [{"s": the row, "note", "kind", "decision"}]."""
    out = []
    for rows, whole, cap in _table_rows(text):
        tot = [r for r in rows if re.match(r"^(?:total|overall|all)\b", r[0], re.I)]
        if len(tot) > 1:
            continue
        total = tot[0][1] if tot else cap
        if not total or total <= 0:
            continue                             # (no total: a question with several answers, not judged)
        body = [r for r in rows if not tot or r is not tot[0]]
        for label, n, p, line in body:
            real = 100.0 * n / total
            d = _dec(p)
            if abs(real - float(p)) > 0.5 * 10 ** -d + 0.05 and n <= total:
                out.append({"s": line[:500], "note": f"In the table, {n:,} of {total:,} is {real:.{max(1, d)}f}%, not {p}%.", "kind": "percent", "decision": False})
        if tot and whole:
            s_n = sum(r[1] for r in body)
            if s_n != total and len(body) <= 12:
                out.append({"s": tot[0][3][:500], "note": f"The table’s rows add up to {s_n:,}, not its total of {total:,} (a row missing, or a count wrong?).",
                            "kind": "percent", "decision": False})
    return out


# 6.6, round 118 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a table of item means (Likert items, "Mean  SD
# Decision"), each row held to the writing's own scale and criterion mean. A mean outside the scale the writing names ("a
# five-point Likert scale"), an SD wider than the scale allows, and a decision against the criterion it states ("a criterion
# mean of 2.50": an item with a mean of 2.40 marked "Agree") are said.
_CRITERION = re.compile(r"\b(?:criterion|cut-?\s?off|benchmark|decision)\s+(?:mean|point|value|score)?\s*(?:of|was|=|:|is)?\s*(?:set\s+at\s+)?(?P<c>\d\.\d{1,2})\b", re.I)
_DEC_POS = re.compile(r"^(?:agree[sd]?|strongly\s+agree|accept(?:ed)?|positive|high|yes|true|good|adequate|favou?rable|satisfactory|"
                      r"practi[cs]ed|aware|effective|significant)$", re.I)
_DEC_NEG = re.compile(r"^(?:disagree[sd]?|strongly\s+disagree|reject(?:ed)?|negative|low|no|false|poor|inadequate|unfavou?rable|"
                      r"unsatisfactory|not\s+practi[cs]ed|unaware|not\s+aware|ineffective|not\s+significant)$", re.I)
_DECIMAL = re.compile(r"^\d{1,2}\.\d{1,3}$")


_LIKERT_POS = re.compile(r"^(?:SA|strongly\s+agree|A|agree|always|very\s+often|often|usually)$", re.I)
_LIKERT_MID = re.compile(r"^(?:U|N|undecided|neutral|not\s+sure|uncertain|sometimes|occasionally)$", re.I)
_LIKERT_NEG = re.compile(r"^(?:SD|strongly\s+disagree|D|disagree|never|rarely|seldom)$", re.I)


def _likert_cols(header):
    """(first column index, weights) of the answer columns a header names ("SA A D SD", "Always … Never"), or None."""
    lik = lambda c: bool(_LIKERT_POS.match(c) or _LIKERT_MID.match(c) or _LIKERT_NEG.match(c))
    start = next((k for k, c in enumerate(header) if lik(c)), None)
    if start is None:
        return None
    end = start
    while end + 1 < len(header) and lik(header[end + 1]):
        end += 1
    k = end - start + 1                          # (the run of answer columns; "SD" after "Mean" is the standard deviation)
    if not 3 <= k <= 7:
        return None
    first_pos = bool(_LIKERT_POS.match(header[start]))
    return start, (list(range(k, 0, -1)) if first_pos else list(range(1, k + 1)))


def _mean_rows(text):
    """[(label, mean, sd or None, decision or "", line, mean as written, (counts, weights) or None)] for the rows of tables
    of item means."""
    out, header = [], None
    for line in str(text or "").split("\n"):
        cells = [c.strip() for c in _CELL_SPLIT.split(line.strip()) if c.strip()]
        if len(cells) >= 4 and sum(1 for c in cells if re.match(r"^\d+(?:\.\d+)?$", c)) == 0:
            header = cells                       # (a header: no numbers)
            continue
        if len(cells) < 3 or not re.search(r"[A-Za-z]{3}", cells[0]):
            continue
        dec = cells[-1] if (_DEC_POS.match(cells[-1]) or _DEC_NEG.match(cells[-1])) else ""
        nums = [c for c in (cells[1:-1] if dec else cells[1:]) if re.match(r"^\d+(?:\.\d+)?$", c)]
        decs = [c for c in nums if _DECIMAL.match(c)]
        if not decs:
            continue
        if len(decs) >= 2:
            mean, sd, shown = float(decs[-2]), float(decs[-1]), decs[-2]
        else:
            mean, sd, shown = float(decs[-1]), None, decs[-1]
        counts = None
        lk = _likert_cols(header) if header and len(header) == len(cells) else None
        if lk:
            at, w = lk
            got = cells[at:at + len(w)]
            if all(_INT_CELL.match(c) for c in got):
                counts = ([int(c) for c in got], w)
        out.append((cells[0], mean, sd, dec, line.strip(), shown, counts))
    return out


# 6.6, round 141: a scale scored the other way round ("1 = strongly agree … 4 = strongly disagree"): its means are worked out
# with the weights the writing gives (or, where it doesn't say, the ones its own means fit), and its decisions against a
# criterion are not judged (which side of it is agreement is then the writing's to say). A scale the writing names with
# another number of points than the table's answer columns is said once.
_LIK_REVERSED = re.compile(r"\b1\s*=\s*(?:strongly\s+agree|always|very\s+often|very\s+high|excellent)\b|"
                           r"\b(?:strongly\s+agree|always)\s*(?:=|\(|-|–)\s*1\b", re.I)


def mean_table_check(text):
    sc = _scale_of(str(text or ""))
    cm = _CRITERION.search(str(text or ""))
    crit = float(cm.group("c")) if cm else None
    rows = _mean_rows(text)
    if not sc and crit is None and not any(r[6] for r in rows):
        return []
    out = []

    def fits(counts, w, mean, shown):
        n_ = sum(counts)
        return n_ > 0 and abs(sum(a * b for a, b in zip(counts, w)) / n_ - mean) <= 0.5 * 10 ** -_dec(shown) + 0.006
    with_counts = [r for r in rows if r[6]]
    fit_n = sum(1 for r in with_counts if fits(r[6][0], r[6][1], r[1], r[5]))
    fit_r = sum(1 for r in with_counts if fits(r[6][0], r[6][1][::-1], r[1], r[5]))
    reverse = bool(_LIK_REVERSED.search(str(text or ""))) or fit_r > fit_n
    said_points = set()
    for label, mean, sd, dec, line, shown, counts in rows:
        if counts and sc and len(counts[1]) != sc[1] - sc[0] + 1 and len(counts[1]) not in said_points:
            said_points.add(len(counts[1]))
            out.append({"s": line[:500], "kind": "value", "decision": False,
                        "note": f"The writing names a {sc[1] - sc[0] + 1}-point scale, but this table has {len(counts[1])} answer columns."})
        if counts and sum(counts[0]) > 0:
            f, w = counts
            if reverse:
                w = w[::-1]
            n = sum(f)
            calc = sum(a * b for a, b in zip(f, w)) / n
            if abs(calc - mean) > 0.5 * 10 ** -_dec(shown) + 0.006:
                out.append({"s": line[:500], "note": f"From its counts the mean is {calc:.2f}, not {shown}.", "kind": "value", "decision": False})
                continue
        if sc:
            lo, hi = sc
            if mean > hi + 1e-9 or mean < lo - 1e-9:
                out.append({"s": line[:500], "note": f"A mean on a {lo}–{hi} scale lies between {lo} and {hi} (the table gives {shown}).",
                            "kind": "value", "decision": False})
                continue
            if sd is not None and sd > (hi - lo) / 2.0 * 1.1 + 0.05:
                out.append({"s": line[:500], "note": f"On a {lo}–{hi} scale a standard deviation can’t be more than about {(hi - lo) / 2.0:g} (the table gives {sd:g}).",
                            "kind": "value", "decision": False})
                continue
        if crit is not None and dec and not reverse and (sc is None or sc[0] <= crit <= sc[1]):
            pos = bool(_DEC_POS.match(dec))
            if pos and mean < crit - 1e-9:
                out.append({"s": line[:500], "note": f"A mean of {shown} is below the criterion mean of {cm.group('c')}, so it isn’t “{dec}”.",
                            "kind": "value", "decision": True})
            elif not pos and mean >= crit + 1e-9:
                out.append({"s": line[:500], "note": f"A mean of {shown} is at or above the criterion mean of {cm.group('c')}, so it isn’t “{dec}”.",
                            "kind": "value", "decision": True})
    return out


# 6.6, round 119 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a cross-table with its chi-square, worked out again
# from its own counts. Results chapters test "the association between socio-demographic characteristics and practice" with
# tables like "Male  30 (40.0)  45 (60.0)  3.24  0.072 / Female  20 (26.7)  55 (73.3)": the χ² is worked out again from the
# counts (Pearson's, or Yates's for a 2 × 2 table), and its p from the χ² and the table's degrees of freedom. A table that names
# Fisher's exact test, or whose counts can't all be read, is not judged.
_NP_CELL = re.compile(r"^(?P<n>\d{1,6})\s*(?:\(\s*\d{1,3}(?:\.\d{1,2})?\s*%?\s*\))?$")
_P_CELL = re.compile(r"^(?P<rel><=|>=|≤|≥|<|>|=)?\s*(?P<p>0?\.\d+|1(?:\.0+)?|0)\*{0,3}$")
_NUM_CELL = re.compile(r"^\d+(?:\.\d+)?\*{0,3}$")


def _xtab_blocks(text):
    """[(rows of counts, χ² as written, df as written or None, (rel, p as written), the first row's line)] for each
    variable of a cross-table."""
    out, cur, header = [], None, ""
    for line in str(text or "").split("\n"):
        cells = [c.strip() for c in _CELL_SPLIT.split(line.strip()) if c.strip()]
        if len(cells) >= 3 and not any(re.search(r"\d", c) for c in cells[1:]):
            header = line
        if (len(cells) >= 4 and re.search(r"[A-Za-z]", cells[0]) and not _NP_CELL.match(cells[0]) and not _NP_CELL.match(cells[1])
                and not _NUM_CELL.match(cells[1]) and _NP_CELL.match(cells[2])):
            cells = cells[1:]                    # (6.6, round 173: "Sex | Male | 40 (50.0) | …": the variable's name set aside)
        if len(cells) < 3 or not re.search(r"[A-Za-z0-9]", cells[0]):
            if cur:
                out.append(cur)
                cur = None
            continue
        k = 1
        while k < len(cells) and _NP_CELL.match(cells[k]):
            k += 1
        counts = [int(_NP_CELL.match(c).group("n")) for c in cells[1:k]]
        rest = cells[k:]
        if len(counts) >= 3 and counts[-1] == sum(counts[:-1]):
            counts = counts[:-1]                 # (a "Total" column)
        if re.match(r"^(?:total|overall|all)\b", cells[0], re.I):
            counts = []                          # (a "Total" row is not a category)
        if len(counts) < 2 or re.search(r"fisher", header, re.I):
            if cur:
                out.append(cur)
                cur = None
            continue
        if rest and len(rest) in (2, 3) and _NUM_CELL.match(rest[0]) and _P_CELL.match(rest[-1]):
            if cur:
                out.append(cur)
            df = rest[1] if len(rest) == 3 and re.match(r"^\d{1,2}$", rest[1]) else None
            pm = _P_CELL.match(rest[-1])
            cur = {"rows": [counts], "chi": rest[0].rstrip("*"), "df": df, "p": (pm.group("rel") or "=", pm.group("p")), "line": line.strip()}
        elif not rest and cur and len(counts) == len(cur["rows"][0]):
            cur["rows"].append(counts)
        else:
            if cur:
                out.append(cur)
            cur = None
    if cur:
        out.append(cur)
    return [b for b in out if len(b["rows"]) >= 2]


def _chi2_of(rows, yates=False):
    tot = sum(sum(r) for r in rows)
    cols = [sum(r[j] for r in rows) for j in range(len(rows[0]))]
    if tot <= 0 or any(c == 0 for c in cols) or any(sum(r) == 0 for r in rows):
        return None
    x = 0.0
    for r in rows:
        rt = sum(r)
        for j, o in enumerate(r):
            e = rt * cols[j] / tot
            d = abs(o - e) - (0.5 if yates else 0.0)
            x += max(d, 0.0) ** 2 / e
    return x


def xtab_check(text):
    out = []
    lines = str(text or "").split("\n")
    for b in _xtab_blocks(text):
        rows = b["rows"]
        df = (len(rows) - 1) * (len(rows[0]) - 1)
        if df <= 0:
            continue
        got = _ct_cochran(rows)                  # 6.6, round 157: Cochran's rule on the table it tested
        if got and (got[0] > 0.2 * got[1] + 1e-9 or got[2] < 1):
            at = next((i for i, x in enumerate(lines) if x.strip() == b["line"]), None)
            if at is None or not _ct_done_near(lines, at - 15, at + 25):
                out.append({"s": b["line"][:500], "note": _ct_rule_note("The table: ", rows, *got), "kind": "test", "decision": False})
        try:
            said = float(b["chi"])
        except ValueError:
            continue
        tol = 0.5 * 10 ** -_dec(b["chi"]) + 0.011 * max(1.0, said)
        options = [x for x in (_chi2_of(rows), _chi2_of(rows, True) if df == 1 else None) if x is not None]
        if not options:
            continue
        if b["df"] and int(b["df"]) != df:
            out.append({"s": b["line"][:500], "note": f"A table of {len(rows)} rows and {len(rows[0])} columns has {df} degree{'s' if df != 1 else ''} "
                                                       f"of freedom, not {b['df']}.", "kind": "test", "decision": False})
            continue
        if all(abs(x - said) > tol for x in options):
            out.append({"s": b["line"][:500], "note": f"From its counts χ²({df}) is {options[0]:.2f}" + (f" ({options[1]:.2f} with Yates's correction)" if len(options) > 1 else "")
                        + f", not {b['chi']}.", "kind": "test", "decision": False})
            continue
        rel, p_txt = b["p"]
        try:
            rep = float(p_txt)
        except ValueError:
            continue
        lo_x, hi_x = said - 0.5 * 10 ** -_dec(b["chi"]), said + 0.5 * 10 ** -_dec(b["chi"])
        plo, phi = p_chi2(max(hi_x, 0), df), p_chi2(max(lo_x, 0), df)
        if plo is None or phi is None:
            continue
        if not _fits(rel if rel in ("<", ">", "≤", "≥", "<=", ">=") else "=", rep, _dec(p_txt), plo, phi):
            real = p_chi2(said, df)
            out.append({"s": b["line"][:500], "note": f"χ²({df}) = {b['chi']} gives {_fmt_p(real, p_txt.startswith('0'))}, not p {rel if rel != '=' else '='} {p_txt}.",
                        "kind": "test", "decision": (real < 0.05) != (rep < 0.05 if rel in ("=", "<", "≤", "<=") else False)})
    return out


# 6.6, round 120 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a regression table, worked out again from its own
# columns. An SPSS table pasted from Word ("B  S.E.  Wald  df  Sig.  Exp(B)" for a logistic regression; "B  Std. Error  t  Sig."
# for a linear one): Wald is (B / S.E.)², Exp(B) is e to the B, Sig. is the p of the Wald on its df; t is B / S.E. Each is
# allowed the rounding of the figures it comes from. A row whose cells can't be matched to its header is not judged.
def _cells_keep(line):
    """A table row's cells, empty ones kept where the row is set out with tabs or "|" (so they stay under their headings)."""
    t = line.rstrip("\n")
    if "\t" in t:
        return [c.strip() for c in t.split("\t")]
    if "|" in t:
        parts = [c.strip() for c in t.strip().strip("|").split("|")]
        return parts
    return [c.strip() for c in re.split(r"\s{2,}", t.strip())]


_REG_COL = {"b": re.compile(r"^(?:unstandardi[sz]ed\s+)?(?:B|b|β\s*\(unstd\.?\)|coef(?:ficient)?\.?|estimate)$", re.I),
            "se": re.compile(r"^(?:S\.?\s?E\.?|std\.?\s*error|standard\s+error)$", re.I),
            "wald": re.compile(r"^wald$", re.I), "df": re.compile(r"^df$", re.I),
            "p": re.compile(r"^(?:sig\.?|p(?:-?\s*value)?|p\s*value)$", re.I),
            "expb": re.compile(r"^(?:exp\s*\(\s*B\s*\)|odds\s+ratio|OR|AOR|COR)$", re.I),
            "t": re.compile(r"^t(?:\s*-?\s*value)?$", re.I), "z": re.compile(r"^z(?:\s*-?\s*value)?$", re.I)}
_FLOAT = re.compile(r"^[-−–]?\d*\.?\d+$")


def _num_cell(c):
    c = str(c or "").strip().rstrip("*").replace("−", "-").replace("–", "-")
    if c.startswith("<"):
        return None
    return float(c) if _FLOAT.match(c) else None


def _iv(x, txt):
    h = 0.5 * 10 ** -_dec(txt.strip().rstrip("*").lstrip("-−–"))
    return x - h, x + h


def reg_table_check(text, counted=None):
    out, cols, judged = [], None, set()
    lines = str(text or "").split("\n")
    allc = [_cells_keep(l) for l in lines]
    anovas = _reg_anovas(text)                   # 6.6, round 127: a coefficient's Sig. from its t on its ANOVA's Residual df
    df_res, n_const, last = None, 0, -1
    for idx, line in enumerate(lines):
        cells = allc[idx]
        if not any(c for c in cells):
            continue
        heads = {}
        for k, c in enumerate(cells):
            for key, rx in _REG_COL.items():
                if rx.match(c) and key not in heads:
                    heads[key] = k
        if "b" in heads and "se" in heads and not any(k in heads for k in ("wald", "t", "z", "expb")):
            up = next((allc[j] for j in range(idx - 1, max(idx - 3, -1), -1) if any(c for c in allc[j])), None)
            for k, c in enumerate(up or []):      # 6.6, round 127: SPSS's heading on two lines ("… t  Sig." over "B  Std. Error  Beta")
                for key in ("t", "z", "p", "wald", "df"):
                    if key not in heads and k not in heads.values() and _REG_COL[key].match((c or "").strip()):
                        heads[key] = k
        if "b" in heads and "se" in heads and ("wald" in heads or "t" in heads or "z" in heads or "expb" in heads):
            cols = heads
            for j in range(idx, min(idx + 2, len(allc))):   # 6.6, round 143: the 95% interval's Lower and Upper (on the next line in SPSS)
                for k, c in enumerate(allc[j]):
                    for key, rx in (("lo", r"^(?:95%\s*(?:c\.?i\.?\s*)?)?lower(?:\s+bound)?$"), ("hi", r"^(?:95%\s*(?:c\.?i\.?\s*)?)?upper(?:\s+bound)?$")):
                        if key not in cols and re.match(rx, (c or "").strip(), re.I):
                            cols[key] = k
            above = [a_ for a_ in anovas if last < a_["line"] < idx]
            df_res = above[-1]["res"] if above and above[-1]["one"] else None
            n_const, last = 0, idx
            continue
        lk, lab = _row_label(cells)
        if not cols or len(cells) <= max(v for k_, v in cols.items() if k_ not in ("lo", "hi")) or lk is None or lk in cols.values():
            continue
        if re.match(r"^\s*\(?\s*constant\b", lab, re.I):
            n_const += 1                          # (a second constant is a second model: its df can't be told)
            if n_const > 1:
                df_res = None
            continue                              # (the constant: its Exp(B) is often left out)
        b_t, se_t = cells[cols["b"]], cells[cols["se"]]
        b, se = _num_cell(b_t), _num_cell(se_t)
        if b is None or se is None or se <= 0:
            continue
        blo, bhi = _iv(b, b_t)
        slo, shi = _iv(se, se_t)
        rats = [x / y for x in (blo, bhi) for y in (slo, shi) if y > 0]
        lo_r, hi_r = min(rats), max(rats)
        label = lab[:60]
        if "wald" in cols and cells[cols["wald"]]:
            w_t = cells[cols["wald"]]
            w = _num_cell(w_t)
            if w is not None:
                sq = [r * r for r in rats] + ([0.0] if lo_r < 0 < hi_r else [])
                wlo, whi = _iv(w, w_t)
                if whi < min(sq) - 1e-9 or wlo > max(sq) + 1e-9:
                    judged.add(idx)
                    out.append({"s": line.strip()[:500], "note": f"{label}: Wald is (B / S.E.)² = {(b / se) ** 2:.3f}, not {w_t}.", "kind": "test", "decision": False})
                    continue
                if "p" in cols and cells[cols["p"]]:
                    p_t = cells[cols["p"]].strip().rstrip("*")
                    pc = _p_cell(p_t)
                    d = int(cells[cols["df"]]) if "df" in cols and re.match(r"^\d{1,2}$", cells[cols["df"]] or "") else 1
                    if pc is not None:
                        rel, pv, dp = pc
                        plo, phi = p_chi2(max(whi, 0), d), p_chi2(max(wlo, 0), d)
                        judged.add(idx)
                        if plo is not None and not _fits(rel, pv, dp, plo, phi):
                            real = p_chi2(w, d)
                            judged.add(idx)
                            out.append({"s": line.strip()[:500], "note": f"{label}: a Wald of {w_t} on {d} df gives {_fmt_p(real, _p_lead(p_t))}, not Sig. {_p_said(p_t)}.",
                                        "kind": "test", "decision": (real < 0.05) != (pv < 0.05)})
                            continue
        if "expb" in cols and cells[cols["expb"]]:
            e_t = cells[cols["expb"]]
            e = _num_cell(e_t)
            if e is not None and e > 0:
                elo, ehi = _iv(e, e_t)
                if ehi < math.exp(blo) - 1e-9 or elo > math.exp(bhi) + 1e-9:
                    judged.add(idx)
                    out.append({"s": line.strip()[:500], "note": f"{label}: Exp(B) is e^B = {math.exp(b):.3f}, not {e_t}.", "kind": "test", "decision": False})
                    continue
        lo_t = (cells[cols["lo"]] if "lo" in cols and cols["lo"] < len(cells) else "").strip()
        hi_t = (cells[cols["hi"]] if "hi" in cols and cols["hi"] < len(cells) else "").strip()
        lo_v, hi_v = _num_cell(lo_t), _num_cell(hi_t)
        if lo_v is not None and hi_v is not None:       # 6.6, round 143: the 95% interval, from B and its S.E.
            crit = 1.959964 if ("wald" in cols or "expb" in cols or "z" in cols) else (_t_crit(df_res) if df_res else None)
            if crit is not None:
                ratio = "expb" in cols or "wald" in cols
                f_ = math.exp if ratio else (lambda x: x)
                ends = [(f_(x - crit * y), f_(x + crit * y)) for x in (blo, bhi) for y in (slo, shi)]
                lo_lo, lo_hi = min(e_[0] for e_ in ends), max(e_[0] for e_ in ends)
                hi_lo, hi_hi = min(e_[1] for e_ in ends), max(e_[1] for e_ in ends)
                hl, hh = 0.5 * 10 ** -_dec(lo_t.lstrip("-−")), 0.5 * 10 ** -_dec(hi_t.lstrip("-−"))
                if lo_v + hl < lo_lo - 1e-9 or lo_v - hl > lo_hi + 1e-9 or hi_v + hh < hi_lo - 1e-9 or hi_v - hh > hi_hi + 1e-9:
                    judged.add(idx)
                    what = "Exp(B)" if ratio else "B"
                    how = "e^(B ∓ 1.96 × S.E.)" if ratio else f"B ∓ t(.975, {df_res}) × its Std. Error" if crit != 1.959964 else "B ∓ 1.96 × S.E."
                    out.append({"s": line.strip()[:500], "kind": "test", "decision": False,
                                "note": f"{label}: the 95% interval for {what} is {how}, about {_like(f_(b - crit * se), lo_t)} to {_like(f_(b + crit * se), hi_t)}, "
                                        f"not {lo_t} to {hi_t}."})
                    continue
        for key in ("t", "z"):
            if key in cols and cells[cols[key]]:
                v_t = cells[cols[key]]
                v = _num_cell(v_t)
                if v is not None:
                    vlo, vhi = _iv(v, v_t)
                    if vhi < lo_r - 1e-9 or vlo > hi_r + 1e-9:
                        judged.add(idx)
                        out.append({"s": line.strip()[:500], "note": f"{label}: {key} is B / S.E. = {b / se:.2f}, not {v_t}.", "kind": "test", "decision": False})
                    elif "p" in cols and (key == "z" or df_res) and "wald" not in cols:
                        sig_t = (cells[cols["p"]] or "").strip()      # (not p_t: that is the t distribution's p)
                        pc = _p_cell(sig_t)
                        if pc is not None:                 # 6.6, round 127: its Sig. is t's p on the Residual's df (z's: the normal's)
                            rel, pv, dp = pc
                            a_lo = 0.0 if vlo <= 0 <= vhi else min(abs(vlo), abs(vhi))
                            a_hi = max(abs(vlo), abs(vhi))
                            pf = (lambda x: p_z(x)) if key == "z" else (lambda x: p_t_two(x, df_res))
                            plo, phi = pf(a_hi), pf(a_lo)
                            if plo is not None and phi is not None:
                                judged.add(idx)
                                if not _fits(rel, pv, dp, plo, phi):
                                    real = pf(abs(v))
                                    on = f" on {df_res} df" if key == "t" else ""
                                    out.append({"s": line.strip()[:500], "kind": "test",
                                                "note": f"{label}: {key} = {v_t}{on} gives {_fmt_p(real, _p_lead(sig_t))}, not Sig. {_p_said(sig_t)}.",
                                                "decision": (real < 0.05) != (pv < 0.05 if rel == "=" else rel in ("<", "≤"))})
                    elif "p" in cols and key == "t" and not df_res and "wald" not in cols and "expb" not in cols:
                        sig_t = (cells[cols["p"]] or "").strip()      # 6.6, round 149: no ANOVA table, so no df: the p any df allows
                        pc = _p_cell(sig_t)
                        if pc is not None:
                            rel, pv, dp = pc
                            a_lo = 0.0 if vlo <= 0 <= vhi else min(abs(vlo), abs(vhi))
                            a_hi = max(abs(vlo), abs(vhi))
                            least, most = p_z(a_hi), p_t_two(a_lo, 1)      # (df from 1 to as many as there can be)
                            if least is not None and most is not None:
                                judged.add(idx)
                                if not _fits(rel, pv, dp, least, most):
                                    lead = _p_lead(sig_t)
                                    if (pv if rel == "=" else pv) > most:
                                        say = f"a p of {_fmt_p(p_t_two(abs(v), 1), lead)[4:]} or less, whatever its df"
                                    else:
                                        say = f"a p of {_fmt_p(p_z(abs(v)), lead)[4:]} or more, whatever its df"
                                    out.append({"s": line.strip()[:500], "kind": "test", "decision": False,
                                                "note": f"{label}: t = {v_t} gives {say}, not Sig. {_p_said(sig_t)}."})
    if counted is not None:
        counted[0] += len(judged)
    return out


# 6.6, round 121 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a one-way ANOVA table (SPSS's "Sum of Squares  df
# Mean Square  F  Sig.", rows Between Groups, Within Groups, Total), worked out again: each Mean Square is its Sum of Squares
# over its df, F is the two Mean Squares' ratio, Sig. is F's p on its two dfs, and the Total row is the sum of the other two.
# Each figure is allowed the rounding of those it comes from; one note for a table (the first thing that doesn't add up).
_AN_COL = {"ss": re.compile(r"^(?:sum\s+of\s+squares|SS|type\s+III\s+sum\s+of\s+squares)$", re.I), "df": re.compile(r"^df$", re.I),
           "ms": re.compile(r"^(?:mean\s+squares?|MS)$", re.I), "f": re.compile(r"^F$"), "p": re.compile(r"^(?:sig\.?|p(?:-?\s*value)?)$", re.I)}


def anova_table_check(text, counted=None):
    out, cols, rows = [], None, {}
    n_judged = [0]

    def judge():
        if not cols or "between" not in rows or "within" not in rows:
            return
        b, w = rows["between"], rows["within"]

        def num(r, key):
            k = cols.get(key)
            if k is None or k >= len(r) or not r[k]:
                return None, None
            v = _num_cell(r[k])
            return (v, r[k].strip().rstrip("*")) if v is not None else (None, None)
        notes = []
        compared, flip = False, False
        ms = {}
        names = rows.get("_names") or ("Between groups", "Within groups")
        for name, r in ((names[0], b), (names[1], w)):
            ss, ss_t = num(r, "ss")
            d, d_t = num(r, "df")
            m, m_t = num(r, "ms")
            if ss is None or not d or d <= 0:
                return
            lo, hi = _iv(ss, ss_t)
            if m is not None:
                mlo, mhi = _iv(m, m_t)
                if mhi < lo / d - 1e-9 or mlo > hi / d + 1e-9:
                    notes.append(f"{name}: the Mean Square is its Sum of Squares over its df, {ss / d:.3f}, not {m_t}.")
                    break
            ms[name] = (lo / d, hi / d, ss / d)
        if not notes and len(ms) == 2:
            f, f_t = num(b, "f")
            (blo, bhi, bmid), (wlo, whi, wmid) = ms[names[0]], ms[names[1]]
            if f is not None and wlo > 0:
                flo_i, fhi_i = _iv(f, f_t)
                if fhi_i < blo / whi - 1e-9 or flo_i > bhi / wlo + 1e-9:
                    notes.append(f"F is the two Mean Squares’ ratio, {bmid / wmid:.3f}, not {f_t}.")
                else:
                    kp = cols.get("p")
                    p_t = (b[kp] or "").strip().rstrip("*") if kp is not None and kp < len(b) else ""
                    pc = _p_cell(p_t)
                    db, dw = num(b, "df")[0], num(w, "df")[0]
                    if pc is not None and db and dw:
                        rel, p, dp = pc
                        plo, phi = p_f(fhi_i, int(db), int(dw)), p_f(max(flo_i, 0), int(db), int(dw))
                        if plo is not None and phi is not None:
                            compared = True
                            if not _fits(rel, p, dp, plo, phi):
                                real = p_f(f, int(db), int(dw))
                                flip = (real < 0.05) != (p < 0.05 if rel == "=" else rel in ("<", "≤"))
                                notes.append(f"F({int(db)}, {int(dw)}) = {f_t} gives {_fmt_p(real, _p_lead(p_t))}, not Sig. {_p_said(p_t)}.")
        if not notes and "total" in rows:
            t = rows["total"]
            for key, what in (("ss", "Sum of Squares"), ("df", "df")):
                vb, tb = num(b, key)
                vw, tw = num(w, key)
                vt, tt = num(t, key)
                if None in (vb, vw, vt):
                    continue
                slack = 0.5 * 10 ** -_dec(tb) + 0.5 * 10 ** -_dec(tw) + 0.5 * 10 ** -_dec(tt) + 1e-9
                if abs(vb + vw - vt) > slack:
                    notes.append(f"The Total row’s {what} is the sum of the other two, {vb + vw:g}, not {tt}.")
                    break
        if notes or compared:
            n_judged[0] += 1
        if notes:
            out.append({"s": (rows.get("_line") or "")[:500], "note": notes[0], "kind": "test", "decision": flip and len(notes) == 1})

    for line in str(text or "").split("\n"):
        cells = _cells_keep(line)
        if not any(c for c in cells):
            continue
        heads = {}
        for k, c in enumerate(cells):
            for key, rx in _AN_COL.items():
                if rx.match(c) and key not in heads:
                    heads[key] = k
        if "ss" in heads and "df" in heads and ("f" in heads or "ms" in heads):
            judge()
            cols, rows = heads, {}
            continue
        if not cols:
            continue
        lab = _row_label(cells)[1].lower()       # 6.6, round 127: SPSS writes the model's number before "Regression"
        if re.match(r"^between(?:\s+groups?)?$|^regression$", lab):
            rows["between"], rows["_line"] = cells, line.strip()
            rows["_names"] = ("Regression", "Residual") if lab == "regression" else ("Between groups", "Within groups")
        elif re.match(r"^within(?:\s+groups?)?$|^(?:error|residual)$", lab):
            rows["within"] = cells
        elif re.match(r"^total$", lab):
            rows["total"] = cells
            judge()
            cols, rows = None, {}
    judge()
    if counted is not None:
        counted[0] += n_judged[0]
    return out


# 6.6, round 122 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a table comparing two groups' means with a t-test
# ("Group  n  Mean  SD  t  df  p": "Male 75 12.4 3.1 2.35 148 0.020 / Female 75 11.2 3.2"), worked out again: t from the two
# means, SDs and ns (pooled, or Welch's), allowed the rounding of each figure; its df no more than n₁ + n₂ − 2; its p from t
# and df. One note a table.
_TT_COL = {"n": re.compile(r"^(?:n|N|number|frequency|freq\.?)$"), "m": re.compile(r"^(?:mean|M|x̄|average)$", re.I),
           "sd": re.compile(r"^(?:SD|S\.D\.|std\.?\s*deviation|standard\s+deviation)$", re.I), "t": re.compile(r"^t(?:\s*-?\s*value)?$", re.I),
           "df": re.compile(r"^df$", re.I), "p": re.compile(r"^(?:sig\.?(?:\s*\(2-tailed\))?|p(?:-?\s*value)?)$", re.I)}


def t_table_check(text, counted=None):
    out, cols, rows = [], None, []
    n_judged = [0]

    def judge():
        if not cols or len(rows) != 2:
            return
        try:
            (c1, l1), (c2, l2) = rows
            n1, n2 = int(c1[cols["n"]]), int(c2[cols["n"]])
            m1_t, m2_t, s1_t, s2_t = c1[cols["m"]], c2[cols["m"]], c1[cols["sd"]], c2[cols["sd"]]
            m1, m2, s1, s2 = (float(x) for x in (m1_t, m2_t, s1_t, s2_t))
        except (ValueError, IndexError, TypeError):
            return
        src = c1 if c1[cols["t"]].strip() else c2
        t_t = (src[cols["t"]] or "").strip().rstrip("*")
        t = _num_cell(t_t)
        if t is None or n1 < 2 or n2 < 2:
            return
        h = lambda x: 0.5 * 10 ** -_dec(x)
        d_lo = abs(m1 - m2) - h(m1_t) - h(m2_t)
        d_hi = abs(m1 - m2) + h(m1_t) + h(m2_t)
        s1lo, s1hi, s2lo, s2hi = max(s1 - h(s1_t), 1e-9), s1 + h(s1_t), max(s2 - h(s2_t), 1e-9), s2 + h(s2_t)

        def se(a, b, pooled):
            if pooled:
                sp2 = ((n1 - 1) * a * a + (n2 - 1) * b * b) / (n1 + n2 - 2)
                return (sp2 * (1 / n1 + 1 / n2)) ** 0.5
            return (a * a / n1 + b * b / n2) ** 0.5
        lo = min(max(d_lo, 0) / se(s1hi, s2hi, pooled) for pooled in (True, False))
        hi = max(d_hi / se(s1lo, s2lo, pooled) for pooled in (True, False))
        tl, th = abs(t) - h(t_t), abs(t) + h(t_t)
        if th < lo - 1e-9 or tl > hi + 1e-9:
            mid = abs(m1 - m2) / se(s1, s2, True)
            n_judged[0] += 1
            out.append({"s": l1[:500], "note": f"From the two means, SDs and ns, t is about {mid:.2f}, not {t_t}.", "kind": "test", "decision": False})
            return
        df_t = (src[cols["df"]] or "").strip() if "df" in cols else ""
        df = _num_cell(df_t) if df_t else None
        if df is not None and df > n1 + n2 - 2 + 0.5:
            n_judged[0] += 1
            out.append({"s": l1[:500], "note": f"Two groups of {n1} and {n2} give at most {n1 + n2 - 2} degrees of freedom, not {df_t}.", "kind": "test", "decision": False})
            return
        if "p" in cols and df:
            p_t = (src[cols["p"]] or "").strip().rstrip("*")
            pc = _p_cell(p_t)
            if pc is not None:
                rel, pv, dp = pc
                plo, phi = p_t_two(th, df), p_t_two(max(tl, 0), df)
                n_judged[0] += 1
                if not _fits(rel, pv, dp, plo, phi):
                    real = p_t_two(abs(t), df)
                    out.append({"s": l1[:500], "note": f"t({df_t}) = {t_t} gives {_fmt_p(real, _p_lead(p_t))}, not p {_p_said(p_t)}.", "kind": "test",
                                "decision": (real < 0.05) != (pv < 0.05 if rel == "=" else rel in ("<", "≤"))})

    for line in str(text or "").split("\n"):
        cells = _cells_keep(line)
        if not any(c for c in cells):
            continue
        heads = {}
        for k, c in enumerate(cells):
            for key, rx in _TT_COL.items():
                if rx.match(c) and key not in heads:
                    heads[key] = k
        if all(k in heads for k in ("n", "m", "sd", "t")):
            judge()
            cols, rows = heads, []
            continue
        if cols and len(cells) > max(cols["n"], cols["m"], cols["sd"]) and re.search(r"[A-Za-z]", cells[0] or "") \
                and _INT_CELL.match((cells[cols["n"]] or "").strip()):
            rows.append(([c or "" for c in cells] + [""] * (max(cols.values()) + 1 - len(cells)), line.strip()))
            if len(rows) == 2:
                judge()
                rows = []
        elif cols and rows:
            judge()
            rows = []
    judge()
    if counted is not None:
        counted[0] += n_judged[0]
    return out


def p_t_two(t, df):
    return p_t(abs(t), df)


# 6.6, round 123 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): the response rate, worked out again. A results
# chapter opens with it: "Of the 160 questionnaires distributed, 150 were returned, giving a response rate of 95.8%" (150 of
# 160 is 93.8%). Said in the sentence (or the two sentences) that give the three figures.
_RR_OUT = re.compile(r"(?P<a>\d[\d,]*)\s+(?:questionnaires?|copies|forms|instruments|surveys?)\b[^.;]{0,60}?\b(?:were\s+|was\s+)?"
                     r"(?:distributed|administered|given\s+out|sent(?:\s+out)?|issued|handed\s+out|shared)", re.I)
_RR_OUT2 = re.compile(r"(?:distributed|administered|gave\s+out|sent\s+out|issued|handed\s+out)\s+(?P<a>\d[\d,]*)\s+(?:questionnaires?|copies|forms)", re.I)
_RR_BACK = re.compile(r"(?P<b>\d[\d,]*)\s+(?:(?:of\s+them|questionnaires?|copies|forms|were|was)\s+){0,3}?"
                      r"(?:returned|completed|retrieved|collected|received|filled|answered|valid|usable|analy[sz]ed)\b", re.I)
_RR_RATE = re.compile(r"response\s+rate\s+(?:of\s+|was\s+|=\s*|:\s*)?(?:about\s+)?(?P<r>\d{1,3}(?:\.\d+)?)\s?%", re.I)


def _rr_notes(text):
    out = []
    sents = sentences(text)
    for i, s in enumerate(sents):
        rm = _RR_RATE.search(s)
        if not rm:
            continue
        win = " ".join(sents[max(0, i - 1):i + 1])
        om = _RR_OUT.search(win) or _RR_OUT2.search(win)
        if not om:
            continue
        rest = win[om.end():] if om.end() < len(win) else win
        bm = _RR_BACK.search(rest) or _RR_BACK.search(win)
        if not bm:
            continue
        try:
            a, b, r = int(om.group("a").replace(",", "")), int(bm.group("b").replace(",", "")), float(rm.group("r"))
        except ValueError:
            continue
        if a <= 0 or b > a or a == b == 0 or 1900 <= a <= 2100:
            continue
        real = 100.0 * b / a
        if abs(real - r) > 0.5 * 10 ** -_dec(rm.group("r")) + 0.05:
            out.append({"s": s[:500], "note": f"{b:,} returned of {a:,} is a response rate of {real:.1f}%, not {rm.group('r')}%.", "kind": "percent", "decision": False})
    return out


# 6.6, round 124 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): SPSS's "Descriptive Statistics" table ("N  Minimum
# Maximum  Mean  Std. Deviation"), row by row: the mean lies between the minimum and the maximum, and the SD can't be more than
# values between them allow (half the range, times √(N/(N−1))), as round 114 holds a sentence.
_DS_COL = {"n": re.compile(r"^N$"), "min": re.compile(r"^(?:minimum|min\.?)$", re.I), "max": re.compile(r"^(?:maximum|max\.?)$", re.I),
           "m": re.compile(r"^(?:mean|M)$", re.I), "sd": re.compile(r"^(?:std\.?\s*deviation|SD|S\.D\.|standard\s+deviation)$", re.I)}


def desc_table_check(text):
    out, cols = [], None
    for line in str(text or "").split("\n"):
        cells = _cells_keep(line)
        if not any(c for c in cells):
            continue
        heads = {}
        for k, c in enumerate(cells):
            for key, rx in _DS_COL.items():
                if rx.match(c) and key not in heads:
                    heads[key] = k
        if all(k in heads for k in ("min", "max", "m")):
            cols = heads
            continue
        if not cols or len(cells) <= max(cols.values()) or not re.search(r"[A-Za-z]", cells[0] or ""):
            continue
        lo, hi, m = _num_cell(cells[cols["min"]]), _num_cell(cells[cols["max"]]), _num_cell(cells[cols["m"]])
        if None in (lo, hi, m) or hi < lo:
            continue
        label = cells[0].strip()[:60]
        m_t = cells[cols["m"]].strip()
        if m < lo - 0.5 * 10 ** -_dec(m_t) - 1e-9 or m > hi + 0.5 * 10 ** -_dec(m_t) + 1e-9:
            out.append({"s": line.strip()[:500], "note": f"{label}: the mean ({m_t}) lies outside its minimum and maximum ({cells[cols['min']].strip()} to {cells[cols['max']].strip()}).",
                        "kind": "value", "decision": False})
            continue
        if "sd" in cols:
            sd_t = (cells[cols["sd"]] or "").strip()
            sd = _num_cell(sd_t)
            n = int(cells[cols["n"]]) if "n" in cols and _INT_CELL.match((cells[cols["n"]] or "").strip()) else None
            if sd is not None:
                k = (n / (n - 1)) ** 0.5 if n and n > 1 else 2 ** 0.5
                top = (hi - lo) / 2 * k
                if sd > top + 0.5 * 10 ** -_dec(sd_t) + 1e-9:
                    out.append({"s": line.strip()[:500], "note": f"{label}: an SD of {sd_t} is more than values between {cells[cols['min']].strip()} and "
                                                                 f"{cells[cols['max']].strip()} can have (at most {top:.2f}).", "kind": "value", "decision": False})
    return out


# 6.6, round 125 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): SPSS's "Chi-Square Tests" table ("Value  df
# Asymptotic Significance (2-sided)"): the significance of Pearson's chi-square, the continuity correction, the likelihood ratio
# and the linear-by-linear association is the chi-square p of its value on its df. Fisher's exact test is not judged.
_SC_ROW = re.compile(r"^(?:pearson\s+chi-?\s?square|continuity\s+correction\w*|likelihood\s+ratio|linear-?\s?by-?\s?linear\s+association)$", re.I)


def spss_chi_check(text, counted=None):
    out, cols, n_judged = [], None, 0
    for line in str(text or "").split("\n"):
        cells = _cells_keep(line)
        if not any(c for c in cells):
            continue
        low = [c.lower() for c in cells]
        if any(c == "value" for c in low) and any(c == "df" for c in low) and any("sig" in c for c in low):
            cols = {"v": low.index("value"), "df": low.index("df"), "p": next(k for k, c in enumerate(low) if "sig" in c)}
            continue
        if not cols or not _SC_ROW.match((cells[0] or "").strip().rstrip("ab*").strip()) or len(cells) <= max(cols.values()):
            continue
        v_t, df_t, p_t = (cells[cols["v"]] or "").strip().rstrip("ab*"), (cells[cols["df"]] or "").strip(), (cells[cols["p"]] or "").strip().rstrip("*")
        v, pc = _num_cell(v_t), _p_cell(p_t)
        if v is None or pc is None or not re.match(r"^\d{1,3}$", df_t):
            continue
        rel, pv, dp = pc
        d = int(df_t)
        h = 0.5 * 10 ** -_dec(v_t)
        plo, phi = p_chi2(v + h, d), p_chi2(max(v - h, 0), d)
        if plo is None or phi is None:
            continue
        n_judged += 1
        if not _fits(rel, pv, dp, plo, phi):
            real = p_chi2(v, d)
            out.append({"s": line.strip()[:500], "note": f"{cells[0].strip().rstrip('ab*').strip()} {v_t} on {d} df gives {_fmt_p(real, _p_lead(p_t))}, not {_p_said(p_t, bare=True)}.",
                        "kind": "test", "decision": (real < 0.05) != (pv < 0.05 if rel == "=" else rel in ("<", "≤"))})
    if counted is not None:
        counted[0] += n_judged
    return out


# 6.6, round 126 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): SPSS's "Correlations" table. Under each
# variable stand three rows, "Pearson Correlation", "Sig. (2-tailed)" (or 1-tailed) and "N"; cell by cell, an r beyond −1 to
# 1 can't be, and its significance is the p of r on N − 2 df (as round 8 works out a sentence's r), allowed r's rounding.
# The matrix gives each pair twice: it is said once. A table whose rows don't line up under each other (cells set apart by
# spaces lose the empty ones) is left alone, and a pair is named only where the table's own headings or its square shape
# say which two variables it is. Spearman's and Kendall's tables are not judged (their p is worked out other ways).
# Every table check (rounds 119 to 126) now also counts the tests it worked out again, and reads a p written "<.001".
_PV_CELL = re.compile(r"^(?:p\s*)?(?P<rel><=|>=|<|>|≤|≥|=)?\s*(?P<v>0?\.\d+|[01](?:\.0+)?)$", re.I)
_CR_PEARSON = re.compile(r"^pearson(?:['’]s)?\s+correlation(?:\s+coefficient)?$", re.I)
_CR_SIG = re.compile(r"^sig\.?\s*\(\s*(?P<tails>[12])\s*-?\s*tailed\s*\)$", re.I)
_CR_ONES = ("1", "1.0", "1.00", "1.000")
_CR_COEF = re.compile(r"^correlation\s+coefficient$", re.I)       # 6.6, round 151: Spearman's rho (Kendall's tau has another p)
_CR_METHOD = re.compile(r"^(?:spearman['’]?s\s+rho|kendall['’]?s\s+tau(?:[_\s-]?[bc])?)$", re.I)


def _p_cell(c):
    """A table's p cell: (relation, value, decimals) or None (".000", "0.046**", "<.001", "< 0.001", "p = .03")."""
    c = re.sub(r"(?<=\d)[a-e]$", "", str(c or "").strip().rstrip("*").strip()).replace("−", "-")   # (SPSS: ".000b")
    m = _PV_CELL.match(c)
    if not m:
        return None
    rel = {"<=": "≤", ">=": "≥"}.get(m.group("rel") or "=", m.group("rel") or "=")
    return rel, float(m.group("v")), _dec(m.group("v"))


def _p_lead(p_t):
    """Whether the writing puts a 0 before a p's point (then the worked-out p is written the same way)."""
    return str(p_t or "").strip().lstrip("pP<>=≤≥ ").startswith("0")


def _p_said(p_t, bare=False):
    """A p cell as the note quotes it: "= 0.300" (or "0.300" when bare), "< .001"."""
    pc = _p_cell(p_t)
    v = re.sub(r"(?<=\d)[a-e]$", "", str(p_t or "").strip().rstrip("*").strip().lstrip("pP").strip().lstrip("<>=≤≥").strip())
    if not pc or pc[0] == "=":
        return v if bare else "= " + v
    return pc[0] + " " + v


def corr_table_check(text, counted=None):
    out, said, judged = [], set(), set()
    lines = [_cells_keep(l) for l in str(text or "").split("\n")]

    def label_at(cells, rx):
        return next((j for j, c in enumerate(cells[:3]) if rx.match((c or "").strip())), None)
    meth, cur = [], None                          # 6.6, round 151: which coefficient each row's block is ("r", "rho", "tau")
    for cells in lines:
        mm = next((c.strip() for c in cells[:2] if c and _CR_METHOD.match(c.strip())), None)
        if mm:
            cur = "rho" if mm.lower().startswith("spearman") else "tau"
        elif label_at(cells, _CR_PEARSON) is not None:
            cur = "r"
        elif not any((c or "").strip() for c in cells):
            cur = None
        meth.append(cur)

    def lab_k(i):
        k = label_at(lines[i], _CR_PEARSON)
        return k if k is not None or meth[i] != "rho" else label_at(lines[i], _CR_COEF)
    blocks = []
    for i, cells in enumerate(lines):
        if lab_k(i) is not None:
            if blocks and i - blocks[-1][-1] <= 4:
                blocks[-1].append(i)
            else:
                blocks.append([i])
    for block in blocks:
        rows = []                                  # (i, k, row name, Sig row, its offset, tails, N row, its offset)
        for i in block:
            cells, k = lines[i], lab_k(i)
            sig = nrow = None
            s_off = n_off = 0
            tails = 2
            for d in (1, 2, 3):
                if i + d >= len(lines) or lab_k(i + d) is not None:
                    break
                nxt = lines[i + d]
                ks = label_at(nxt, _CR_SIG)
                if ks is not None and sig is None:
                    sig, s_off, tails = nxt, ks - k, int(_CR_SIG.match(nxt[ks].strip()).group("tails"))
                    continue
                kn = next((j for j, c in enumerate(nxt[:3]) if (c or "").strip() == "N"), None)
                if kn is not None and nrow is None:
                    nrow, n_off = nxt, kn - k
            if sig is None or nrow is None:
                continue
            ok = True                              # the three rows must stand under each other
            for j in range(k + 1, len(cells)):
                r_t = (cells[j] or "").strip().rstrip("*").strip()
                if not r_t:
                    continue
                s_c = (sig[j + s_off] if 0 <= j + s_off < len(sig) else "").strip()
                n_c = (nrow[j + n_off] if 0 <= j + n_off < len(nrow) else "").strip()
                if r_t in _CR_ONES:
                    ok = ok and s_c in ("", ".")   # (the diagonal has no significance; SPSS writes "." for Spearman's)
                elif _num_cell(r_t) is not None:
                    ok = ok and _p_cell(s_c) is not None and bool(_INT_CELL.match(n_c.replace(",", "")))
            if ok:
                name = next((c.strip() for c in cells[:k] if c and re.search(r"[A-Za-z]", c) and not _CR_METHOD.match(c.strip())), "")
                rows.append((i, k, name, sig, s_off, tails, nrow, n_off))
        if not rows:
            continue
        # which two variables a cell is: the headings above the table, where they stand over the diagonal's 1s; else the
        # matrix's square shape (its columns are its rows' variables, in order); else no names
        diag = {}
        for i, k, name, *_ in rows:
            ones = [j for j in range(k + 1, len(lines[i])) if (lines[i][j] or "").strip().rstrip("*") in _CR_ONES]
            if len(ones) == 1 and name:
                diag[i] = ones[0]
        col_name = {}
        head = next((lines[block[0] - d] for d in (1, 2, 3) if block[0] - d >= 0
                     and sum(1 for c in lines[block[0] - d] if c and re.search(r"[A-Za-z]", c) and not re.search(r"\d", c)) >= 2), None)
        if head and diag and all(len(head) == len(lines[i]) for i in diag) \
                and all((head[j] or "").strip().lower() == name.lower() for i, k, name, *_ in rows if i in diag for j in [diag[i]]):
            col_name = {j: head[j].strip() for j in range(len(head)) if head[j] and re.search(r"[A-Za-z]", head[j])}
        elif len(rows) == len(block) and len(diag) == len(rows):
            k0 = rows[0][1]
            if all(len(lines[i]) - k - 1 == len(rows) and diag[i] - k - 1 == t for t, (i, k, *_r) in enumerate(rows)) \
                    and all(k == k0 for _i, k, *_r in rows):
                col_name = {k0 + 1 + t: name for t, (_i, _k, name, *_r) in enumerate(rows)}
        for i, k, name, sig, s_off, tails, nrow, n_off in rows:
            cells = lines[i]
            for j in range(k + 1, len(cells)):
                r_t = (cells[j] or "").strip().rstrip("*").strip()
                r = _num_cell(r_t)
                if r is None or r_t in _CR_ONES:
                    continue
                other = col_name.get(j, "")
                cell = lambda row, off: (row[j + off] if 0 <= j + off < len(row) else "").strip()
                pair = (frozenset((name.lower(), other.lower())) if name and other
                        else ("?", r_t, cell(sig, s_off), cell(nrow, n_off)))     # (unnamed: the twin cell by its figures)
                lead = f"{name} and {other}: " if name and other else ""
                if abs(r) > 1:
                    body = f"a correlation can’t be beyond −1 or 1 (the table gives {r_t})."
                    judged.add(pair)
                else:
                    p_t = (sig[j + s_off] or "").strip().rstrip("*").strip()
                    n_t = (nrow[j + n_off] or "").strip().replace(",", "")
                    rel, pv, dp = _p_cell(p_t)
                    df = int(n_t) - 2
                    if df < 2:
                        continue
                    h = 0.5 * 10 ** -_dec(r_t)
                    plo, phi = p_r(min(abs(r) + h, 0.999999), df), p_r(max(abs(r) - h, 0.0), df)
                    if plo is None or phi is None:
                        continue
                    if tails == 1:
                        plo, phi = plo / 2, phi / 2
                    judged.add(pair)
                    if _fits(rel, pv, dp, plo, phi):
                        continue
                    real = p_r(abs(r), df) / (2 if tails == 1 else 1)
                    sym_ = "rho" if meth[i] == "rho" else "r"
                    body = (f"{sym_} = {r_t} with N = {n_t} gives {_fmt_p(real, _p_lead(p_t))}{' one-tailed' if tails == 1 else ''}, "
                            f"not Sig. {_p_said(p_t)}.")
                note = (lead + body if lead else "In the table, " + body if body.startswith(("r = ", "rho = "))
                        else body[0].upper() + body[1:])          # (an r stays small: R is another statistic)
                key = (pair, body) if lead else ("?", body)        # (the matrix's twin cell is said once)
                if key in said:
                    continue
                said.add(key)
                out.append({"s": " ".join(c for c in cells if c)[:500], "note": note, "kind": "test",
                            "decision": abs(r) <= 1 and (real < 0.05) != (pv < 0.05 if rel == "=" else rel in ("<", "≤"))})
    if counted is not None:
        counted[0] += len(judged)
    return out



# 6.6, round 127 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): SPSS's linear regression output, its three
# tables read together. Model Summary: R Square is R²; the Adjusted R Square is what R², the number of predictors and the
# number of people give; and R Square is the Regression's share of the Total Sum of Squares in its ANOVA table. ANOVA: the
# Regression and Residual rows are worked out as round 121 works out Between and Within groups (SPSS puts the model's number
# before a row's name). Coefficients: SPSS's two-line heading ("Unstandardized Coefficients … t  Sig." over "B  Std. Error
# Beta") is read, and each row's Sig. is its t's p on the Residual's df. One model at a time: a table of two models (a
# hierarchical regression) is not tied to another table, since which df is whose can't be told.
_MS_COL = {"r": re.compile(r"^R$"), "r2": re.compile(r"^R\s*(?:Square|²|2|\^2)$", re.I),
           "adj": re.compile(r"^Adj(?:usted|\.)?\s*R\s*(?:Square|²|2|\^2)$", re.I)}
_MODEL_NO = re.compile(r"^\d{1,2}$")
_STEP_NO = re.compile(r"^(?:\d{1,2}|step\s*\d{1,2}\s*[a-e]?|model\s*\d{1,2})$", re.I)   # (SPSS: "1", "Step 1a")


def _row_label(cells):
    """A table row's label and where it stands: its first cell, or its second where the first is empty or a model's number."""
    first = (cells[0] or "").strip() if cells else ""
    if len(cells) > 1 and (not first or _STEP_NO.match(first)) and re.search(r"[A-Za-z]", cells[1] or ""):
        return 1, cells[1].strip()
    if first and re.search(r"[A-Za-z]", first):
        return 0, first
    return None, ""


def _like(x, txt):
    """x written as txt is written: its decimals, and a 0 before the point only where txt has one (6.6, round 142: a negative
    x too: "-.590" beside "-.799", not "-0.590")."""
    t = str(txt or "").strip().lstrip("-−")
    s = f"{abs(x):.{_dec(t)}f}"
    if not t.startswith("0") and s.startswith("0."):
        s = s[1:]
    return ("-" if x < 0 and float(s) != 0 else "") + s


def _reg_anovas(text):
    """The text's regression ANOVA tables: [{"line", "n_reg", "one" (one model, whole), "k", "res", "tot", "ss_reg", "ss_tot", texts}]."""
    found, state = [], {"cols": None, "cur": None}

    def close():
        cur = state["cur"]
        if cur and cur.get("n_reg"):
            cur["one"] = cur["n_reg"] == 1 and {"k", "res", "tot"} <= set(cur)      # (one model, whole)
            found.append(cur)
        state["cur"] = None
    for i, line in enumerate(str(text or "").split("\n")):
        cells = _cells_keep(line)
        if not any(c for c in cells):
            continue
        heads = {}
        for k, c in enumerate(cells):
            for key, rx in _AN_COL.items():
                if rx.match(c) and key not in heads:
                    heads[key] = k
        if "ss" in heads and "df" in heads:
            close()
            state["cols"], state["cur"] = heads, {"line": i, "n_reg": 0}
            continue
        cols, cur = state["cols"], state["cur"]
        if not cols or cur is None:
            continue

        def num(key):
            k = cols.get(key)
            if k is None or k >= len(cells) or not cells[k]:
                return "", None
            return cells[k].strip(), _num_cell(cells[k])
        _lk, lab = _row_label(cells)
        lab = lab.lower()
        if lab not in ("regression", "residual", "total"):
            close()                                # (the table has ended: a note under it, or another table)
            state["cols"] = None
            continue
        (_dt, d), (sst, ss) = num("df"), num("ss")
        if lab == "regression":
            cur["n_reg"] += 1
            if d is not None and d == int(d) and d > 0:
                cur["k"] = int(d)
            cur["ss_reg"], cur["ss_reg_t"] = ss, sst
        elif lab == "residual":
            if d is not None and d == int(d) and d > 0:
                cur["res"] = int(d)
        elif lab == "total":
            if d is not None and d == int(d) and d > 0:
                cur["tot"] = int(d)
            cur["ss_tot"], cur["ss_tot_t"] = ss, sst      # (a second model's rows may follow: the table goes on)
    close()
    return found


def model_summary_check(text):
    out, lines = [], str(text or "").split("\n")
    anovas = _reg_anovas(text)
    heads_at = []                                   # (line, its columns) of each Model Summary heading
    for i, line in enumerate(lines):
        cells = _cells_keep(line)
        heads = {}
        for k, c in enumerate(cells):
            for key, rx in _MS_COL.items():
                if rx.match((c or "").strip()) and key not in heads:
                    heads[key] = k
        if "r" in heads and "r2" in heads:
            heads_at.append((i, heads))
    for h, (i0, cols) in enumerate(heads_at):
        rows = []
        for i in range(i0 + 1, len(lines)):
            cells = _cells_keep(lines[i])
            if not any(c for c in cells):
                if rows:
                    break
                continue
            if not _MODEL_NO.match((cells[0] or "").strip()) or len(cells) <= max(cols.values()):
                break
            rows.append((i, cells))
        nxt = heads_at[h + 1][0] if h + 1 < len(heads_at) else len(lines)
        mine = [a for a in anovas if i0 < a["line"] < nxt]
        tie = mine[0] if len(rows) == 1 and len(mine) == 1 and mine[0]["one"] else None    # its own ANOVA: the one after it
        for i, cells in rows:
            m = cells[0].strip()
            r_t = re.sub(r"[a-e]$", "", (cells[cols["r"]] or "").strip())
            r2_t = (cells[cols["r2"]] or "").strip()
            r, r2 = _num_cell(r_t), _num_cell(r2_t)
            if r is None or r2 is None:
                continue
            note = None
            if r < 0 or r > 1:
                note = f"Model {m}: R lies between 0 and 1 (the table gives {r_t})."
            else:
                rlo, rhi = _iv(r, r_t)
                qlo, qhi = _iv(r2, r2_t)
                if qhi < max(rlo, 0) ** 2 - 1e-12 or qlo > min(rhi, 1) ** 2 + 1e-12:
                    note = f"Model {m}: R Square is R², {_like(r * r, r2_t)}, not {r2_t}."
                elif tie and tie.get("ss_reg") is not None and tie.get("ss_tot"):
                    a_lo, a_hi = _iv(tie["ss_reg"], tie["ss_reg_t"])
                    t_lo, t_hi = _iv(tie["ss_tot"], tie["ss_tot_t"])
                    if t_lo > 0 and (qhi < a_lo / t_hi - 1e-12 or qlo > a_hi / t_lo + 1e-12):
                        note = (f"Model {m}: R Square is the Regression’s share of the Total Sum of Squares in its ANOVA table, "
                                f"{tie['ss_reg_t']} / {tie['ss_tot_t']} = {_like(tie['ss_reg'] / tie['ss_tot'], r2_t)}, not {r2_t}.")
                if note is None and tie and "adj" in cols and cells[cols["adj"]]:
                    adj_t = (cells[cols["adj"]] or "").strip()
                    adj = _num_cell(adj_t)
                    k, res, n = tie["k"], tie["res"], tie["tot"] + 1
                    if adj is not None and res > 0 and n - k - 1 == res:
                        f = lambda q: 1 - (1 - q) * (n - 1) / res
                        alo, ahi = _iv(adj, adj_t)
                        if ahi < f(qlo) - 1e-12 or alo > f(qhi) + 1e-12:
                            note = (f"Model {m}: with {k} predictor{'' if k == 1 else 's'} and N = {n}, the Adjusted R Square is "
                                    f"{_like(f(r2), adj_t)}, not {adj_t}.")
            if note:
                out.append({"s": " ".join(c for c in cells if c)[:500], "note": note, "kind": "value", "decision": False})
    return out



# 6.6, round 129 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): SPSS's Frequencies table ("Frequency  Percent
# Valid Percent  Cumulative Percent", its rows under "Valid" and "Missing", a valid Total and a Total of all). Round 117 read it
# as a plain frequency table: it missed the first row (after "Valid"), and where answers were missing it divided by the valid
# total, so it told a right "Percent" (of everyone) that it was wrong. Now such a table is read for what it is: each Percent
# is its count of everyone, each Valid Percent its count of the valid answers, each Cumulative Percent the valid answers so
# far; the valid rows add up to the valid Total, and the valid and missing answers to the Total of all. One note a row.
_SF_HEAD = re.compile(r"^valid\s+percent$", re.I)


def _spss_freq_tables(text):
    """[(header line, its columns, [(line index, line, cells)])]: the SPSS Frequencies tables of the text."""
    lines = str(text or "").split("\n")
    found, i = [], 0
    while i < len(lines):
        cells = _cells_keep(lines[i])
        low = [(c or "").strip().lower() for c in cells]
        if any(_SF_HEAD.match(c) for c in low) and "frequency" in low and "percent" in low:
            cols = {"f": low.index("frequency"), "p": low.index("percent"), "vp": next(k for k, c in enumerate(low) if _SF_HEAD.match(c))}
            cols["cp"] = next((k for k, c in enumerate(low) if re.match(r"^cumulative\s+percent$", c)), None)
            rows, j = [], i + 1
            while j < len(lines):
                c2 = _cells_keep(lines[j])
                if not any((c or "").strip() for c in c2):
                    if rows:
                        break
                    j += 1
                    continue
                if not any(_INT_CELL.match((c or "").strip().replace(",", "")) for c in c2[1:]):
                    break                          # (a line with no count: the table has ended)
                rows.append((j, lines[j].strip(), c2))
                j += 1
            found.append((i, cols, rows))
            i = j
            continue
        i += 1
    return found


def spss_freq_check(text):
    out = []
    for _h, cols, rows in _spss_freq_tables(text):
        def cell(c2, k):
            return (c2[k] or "").strip() if k is not None and k < len(c2) else ""
        part, valid, missing, vtotal, gtotal, mtotal = "valid", [], [], None, None, None
        for _j, line, c2 in rows:
            first = cell(c2, 0).lower()
            label = cell(c2, 1) or cell(c2, 0)
            if first == "valid":
                part = "valid"
            elif first == "missing":
                part = "missing"
            f_t = cell(c2, cols["f"]).replace(",", "")
            if not _INT_CELL.match(f_t):
                continue
            row = {"line": line, "label": label, "f": int(f_t), "p": cell(c2, cols["p"]), "vp": cell(c2, cols["vp"]), "cp": cell(c2, cols["cp"])}
            if first == "total" and not cell(c2, 1):
                gtotal = row                        # (the Total of all: "Total", nothing, then its count)
            elif label.lower() == "total":
                if part == "valid":
                    vtotal = row
                else:
                    mtotal = row
            elif part == "valid":
                valid.append(row)
            else:
                missing.append(row)
        if len(valid) < 2:
            continue
        def off(stated, real):
            v = _num_cell(stated)
            return v is not None and abs(v - real) > 0.5 * 10 ** -_dec(stated) + 0.05

        def best(cands, key, rows_):
            """The denominator most of the rows' percents fit (a wrong Total then gets one note, not one on every row)."""
            cands = [c for c in dict.fromkeys(cands) if c and c > 0]
            return max(cands, key=lambda c: (sum(1 for r in rows_ if r[key] and not off(r[key], 100.0 * r["f"] / c)), c == cands[0])) if cands else None
        v_sum, miss_sum = sum(r["f"] for r in valid), sum(r["f"] for r in missing)
        v_n = best(([vtotal["f"]] if vtotal else []) + [v_sum], "vp", valid)
        miss_n = mtotal["f"] if mtotal else miss_sum
        g_n = best(([gtotal["f"]] if gtotal else []) + [v_n + miss_n if (missing or mtotal) else v_n], "p", valid + missing)
        if not g_n or not v_n:
            continue
        notes = []
        if vtotal and v_sum != vtotal["f"]:
            notes.append((vtotal["line"], f"The valid rows add up to {v_sum:,}, not their Total of {vtotal['f']:,}."))
        if gtotal and (missing or mtotal) and v_n + miss_n != gtotal["f"]:
            notes.append((gtotal["line"], f"The valid answers ({v_n:,}) and the missing ones ({miss_n:,}) add up to {v_n + miss_n:,}, not the Total of {gtotal['f']:,}."))
        pc = lambda x, t: f"{x:.{max(1, _dec(t))}f}%"
        run, clean = 0, not notes                  # (a Cumulative Percent is judged only while the rows above it are right)
        for r in valid:
            run += r["f"]
            lab = r["label"][:60]
            before = len(notes)
            if r["p"] and off(r["p"], 100.0 * r["f"] / g_n):
                notes.append((r["line"], f"{lab}: its Percent is {r['f']:,} of all {g_n:,}, {pc(100.0 * r['f'] / g_n, r['p'])}, not {r['p']}%."))
            elif r["vp"] and off(r["vp"], 100.0 * r["f"] / v_n):
                notes.append((r["line"], f"{lab}: its Valid Percent is {r['f']:,} of the {v_n:,} valid answers, {pc(100.0 * r['f'] / v_n, r['vp'])}, not {r['vp']}%."))
            elif clean and r["cp"] and off(r["cp"], 100.0 * run / v_n):
                notes.append((r["line"], f"{lab}: its Cumulative Percent is the {run:,} valid answers so far of {v_n:,}, "
                                         f"{pc(100.0 * run / v_n, r['cp'])}, not {r['cp']}%."))
            clean = clean and len(notes) == before
        for r in missing + ([vtotal] if vtotal and vtotal["f"] == v_sum and (missing or mtotal) else []):
            if r["p"] and off(r["p"], 100.0 * r["f"] / g_n):
                notes.append((r["line"], f"{r['label'][:60]}: its Percent is {r['f']:,} of all {g_n:,}, {pc(100.0 * r['f'] / g_n, r['p'])}, not {r['p']}%."))
        for line, note in notes:
            out.append({"s": line[:500], "note": note, "kind": "percent", "decision": False})
    return out



# 6.6, round 130 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): the words of a results chapter against its own
# tables. A table is made again and the sentence under it is not ("The majority of the respondents, 94 (63.7%), were female"
# where the table's Female row gives 94 and 62.7%). A sentence is held to a table's row only where it is plainly the same
# figure: its count is the row's count, or the sentence names the table ("Table 4.1 shows…"); the row's label must be in one
# table only, and not a word every table has (Yes, No, Total, Agree…). A share the text rounds ("63%" for 62.7%) stands.
_TT_GENERIC = {"yes", "no", "total", "other", "others", "none", "all", "agree", "disagree", "strongly agree", "strongly disagree",
               "neutral", "undecided", "never", "rarely", "sometimes", "often", "always", "not sure", "don't know", "do not know",
               "missing", "system", "valid", "uncertain", "very often", "occasionally", "seldom", "true", "false", "good", "poor",
               "high", "low", "moderate", "fair", "average", "excellent", "very good", "n/a", "na"}
_TT_CAPTION = re.compile(r"\btable\s+(\d{1,2}(?:\.\d{1,2}){0,2})\b", re.I)
_TT_PCT = r"(?P<p>\d{1,3}(?:\.\d{1,2})?)\s*%"
_TT_N = r"(?:n\s*=\s*)?(?P<c>\d{1,3}(?:,\d{3})*|\d+)"
_TT_WHO = (r"(?:\s+(?:respondents|participants|students|nurses|midwives|women|men|patients|workers|staff|mothers|children|pupils|"
           r"adolescents|teachers|members|households|cases))?")                            # 6.6, round 138: "female respondents (94, 62.7%)"


def _tt_tables(text):
    """[{"no": its number or "", "rows": [(label, count, [percent strings]) …], "lines": set of its line indexes}]."""
    lines = str(text or "").split("\n")
    out, used = [], set()

    def caption(i):
        for d in (1, 2, 3, 4):
            if i - d >= 0:
                m = _TT_CAPTION.search(lines[i - d])
                if m:
                    return m.group(1)
        return ""
    for h, cols, rows in _spss_freq_tables(text):
        got = []
        for j, _line, c2 in rows:
            lab = (c2[1] if len(c2) > 1 and (c2[1] or "").strip() else c2[0] if c2 else "").strip()
            f_t = (c2[cols["f"]] if cols["f"] < len(c2) else "").strip().replace(",", "")
            if lab and _INT_CELL.match(f_t):
                pcts = [x for x in ((c2[cols["p"]] if cols["p"] < len(c2) else ""), (c2[cols["vp"]] if cols["vp"] < len(c2) else "")) if _PCT_CELL.match(x.strip())]
                got.append((lab, int(f_t), [x.strip().rstrip("%").strip() for x in pcts]))
        out.append({"no": caption(h), "rows": got, "lines": {h} | {j for j, _l, _c in rows}})
    start = 0
    for rows, _whole, _cap in _table_rows(text):
        first = next((i for i in range(start, len(lines)) if lines[i].strip() == rows[0][3]), None)
        if first is None:
            continue
        idx = set(range(first, first + len(rows)))
        start = first + len(rows)
        out.append({"no": caption(first), "rows": [(lab, n, [p]) for lab, n, p, _l in rows], "lines": idx})
    return out


def text_table_check(text):
    out, said = [], set()
    tables = _tt_tables(text)
    if not tables:
        return out
    where = {}
    for t in tables:
        for lab, n, pcts in t["rows"]:
            key = re.sub(r"\s+", " ", lab.strip().lower())
            if key in _TT_GENERIC or len(key) < 3 or not re.search(r"[a-z]", key) or re.search(r"\d", key):
                continue
            where.setdefault(key, []).append((t, lab, n, pcts))
    unique = {k: v[0] for k, v in where.items() if len(v) == 1}
    if not unique:
        return out
    skip = set().union(*(t["lines"] for t in tables))
    lines = str(text or "").split("\n")
    prose = "\n".join("" if i in skip else ln for i, ln in enumerate(lines))
    for sent in sentences(prose):
        if "\t" in sent or "%" not in sent:
            continue
        named = {m.group(1) for m in _TT_CAPTION.finditer(sent)}
        low = sent.lower()
        for key, (t, lab, n, pcts) in unique.items():
            if key not in low:
                continue
            L = r"\b" + r"\s+".join(re.escape(w) for w in key.split()) + r"(?:s|es)?\b"
            pats = (L + _TT_WHO + r"\s*\(\s*(?:" + _TT_N + r"\s*[,;]\s*)?" + _TT_PCT + r"\s*\)",       # female (94, 63.7%)
                    L + _TT_WHO + r"\s*\(\s*" + _TT_PCT + r"\s*[,;]\s*n\s*=\s*(?P<c>\d[\d,]*)\s*\)",       # female (63.7%, n = 94)
                    _TT_PCT + r"\s*\)?\s*(?:\(\s*" + _TT_N + r"\s*\)\s*)?(?:of\s+(?:the\s+)?(?:[a-z]+\s+){0,2})?,?\s*(?:were|was|are|is)\s+(?:[a-z]+\s+)?" + L,
                    r"(?P<c>\d[\d,]*)\s*\(\s*" + _TT_PCT + r"\s*\)\s*,?\s*(?:of\s+(?:the\s+)?(?:[a-z]+\s+){0,2})?,?\s*(?:were|was|are|is)\s+(?:[a-z]+\s+)?" + L,
                    L + r"\s+(?:respondents\s+|participants\s+|students\s+)?(?:constituted|accounted\s+for|made\s+up|formed|represented|comprised)\s+"
                    + r"(?:(?P<c>\d[\d,]*)\s*\(\s*)?" + _TT_PCT,
                    r"(?P<c>\d[\d,]*)" + _TT_WHO + r"\s*,?\s*(?:representing|constituting|forming|making\s+up|being)\s+" + _TT_PCT
                    + r"\s*,?\s*(?:of\s+(?:the\s+)?(?:[a-z]+\s+){0,2})?(?:were|was|are|is)\s+(?:[a-z]+\s+)?" + L)
            for rx in pats:
                for m in re.finditer(rx, low):
                    c_t, p_t = m.groupdict().get("c"), m.group("p")
                    count = int(c_t.replace(",", "")) if c_t else None
                    same = (count is not None and count == n) or (t["no"] and t["no"] in named)
                    if not same:
                        continue
                    pv, d = float(p_t), _dec(p_t)
                    fits = lambda x: abs(pv - float(x)) <= 0.5 * 10 ** -d + 0.05 + 1e-9
                    if any(fits(x) for x in pcts) and (count is None or count == n):
                        continue
                    tname = f"Table {t['no']}" if t["no"] else "the table"
                    if count is not None:
                        note = f"The text gives {count:,} ({p_t}%) for {lab}; {tname} gives {n:,} ({pcts[0]}%)."
                    else:
                        note = f"The text gives {p_t}% for {lab}; {tname} gives {pcts[0]}%."
                    other = [ol for ol, _on, ops in t["rows"] if ol != lab and not re.match(r"^total$", ol, re.I)
                             and any(abs(pv - float(x)) <= 0.5 * 10 ** -d + 0.05 + 1e-9 for x in ops)]
                    if len(other) == 1:
                        note = note[:-1] + f" ({p_t}% is {other[0]}’s)."
                    if (sent, note) in said:
                        continue
                    said.add((sent, note))
                    out.append({"s": sent[:500], "note": note, "kind": "percent", "decision": False})
    return out



# 6.6, round 131 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): SPSS's Crosstabulation ("Sex * Self-medication
# Crosstabulation": under each row, "Count", "% within Sex", "% within Self-medication", "% of Total", "Expected Count") and
# the "Chi-Square Tests" table that follows it. Each row's counts add up to its Total, and each column's to the Total row's;
# each percentage is its count over the row's, the column's or the whole table's total, as its label says; each Expected
# Count is the row's total times the column's over all; and the Chi-Square Tests below hold the table's own Pearson chi-square
# (Yates's for 2 × 2, and the likelihood ratio), its df and its N of Valid Cases. One note a row.
_CT_TITLE = re.compile(r"^(?P<r>.+?)\s*\*\s*(?P<c>.+?)\s+crosstabulation\b", re.I)
_CT_TEST_ROW = re.compile(r"^(pearson\s+chi-?\s?square|continuity\s+correction|likelihood\s+ratio|n\s+of\s+valid\s+cases)\s*[a-e]?$", re.I)
_CT_STAT = re.compile(r"^(?:count|expected\s+count|%\s+within\s+.+|%\s+of\s+total|residual|std\.?\s+residual|adjusted\s+residual)$", re.I)


def _ct_num(c):
    c = (c or "").strip().replace(",", "")
    m = re.match(r"^(-?\d+(?:\.\d+)?)\s*%?$", c)
    return float(m.group(1)) if m else None


def _crosstabs(text):
    """The SPSS Crosstabulations of the text: [{"rv", "cv", "names", "blocks": [{"label", "rows": {stat: (label, texts, line)}}],
    "end"}]."""
    lines = str(text or "").split("\n")
    found = []
    for i, line in enumerate(lines):
        m = _CT_TITLE.match(line.strip())
        if not m:
            continue
        rv, cv = m.group("r").strip().lower(), m.group("c").strip().lower()
        blocks, k_stat, names, j, heads = [], None, [], i + 1, []
        while j < len(lines) and j < i + 80:
            cells = _cells_keep(lines[j])
            if not any((c or "").strip() for c in cells):
                if blocks:
                    break
                j += 1
                continue
            ks = next((k for k in range(min(4, len(cells))) if _CT_STAT.match((cells[k] or "").strip())), None)
            if ks is None:
                if blocks:
                    break                          # (a line that isn't a row: the table has ended)
                heads.append(cells)
                j += 1
                continue
            if k_stat is None:
                k_stat = ks
            if ks != k_stat:
                break
            stat_label = (cells[ks] or "").strip()
            stat = re.sub(r"\s+", " ", stat_label.lower())
            lab = next(((cells[k] or "").strip() for k in range(ks - 1, -1, -1) if (cells[k] or "").strip()), "")
            if (cells[0] or "").strip().lower() == "total":
                lab = "Total"
            if stat == "count" or not blocks:
                blocks.append({"label": lab, "rows": {}})
            texts = [(v or "").strip() for v in cells[ks + 1:]]
            while texts and not texts[-1]:
                texts = texts[:-1]
            blocks[-1]["rows"][stat] = (stat_label, texts, lines[j].strip())
            j += 1
        if len(blocks) < 3 or k_stat is None:
            continue
        width = len(blocks[0]["rows"].get("count", ("", [], ""))[1])
        for hcells in reversed(heads):
            got = [(hcells[k] or "").strip() for k in range(k_stat + 1, k_stat + width)] if len(hcells) >= k_stat + width else []
            if got and all(got) and not any(_ct_num(g) is not None for g in got):
                names = got
                break
        found.append({"rv": rv, "cv": cv, "names": names, "blocks": blocks, "end": j})
    return found


def _ct_tests(text, start, obs, n):
    """Notes on the Chi-Square Tests that follow a crosstab (within 30 lines, before another crosstab), held to its counts."""
    lines, notes = str(text or "").split("\n"), []
    df = (len(obs) - 1) * (len(obs[0]) - 1)
    tot = sum(sum(r) for r in obs)
    cs = [sum(r[j] for r in obs) for j in range(len(obs[0]))]
    want_of = {"pearson chi-square": _chi2_of(obs), "continuity correction": _chi2_of(obs, True) if df == 1 else None}
    try:
        want_of["likelihood ratio"] = 2.0 * sum(o * math.log(o / (sum(r) * cs[j] / tot)) for r in obs for j, o in enumerate(r) if o > 0)
    except (ValueError, ZeroDivisionError):
        pass
    for i in range(start, min(len(lines), start + 30)):
        if _CT_TITLE.match(lines[i].strip()):
            break
        low = [(x or "").strip().lower() for x in _cells_keep(lines[i])]
        if "value" not in low or "df" not in low:
            continue
        kv, kd = low.index("value"), low.index("df")
        for i3 in range(i + 1, min(len(lines), i + 10)):
            c3 = _cells_keep(lines[i3])
            mt = _CT_TEST_ROW.match((c3[0] or "").strip()) if c3 else None
            if not mt:
                continue                           # (SPSS's footnote letter read past: "Continuity Correctionb")
            name = re.sub(r"\s+", " ", mt.group(1)).strip()
            if name.lower().startswith("n of valid cases"):
                nv = next((_ct_num(x) for x in c3[1:] if _ct_num(x) is not None), None)
                if nv is not None and int(nv) != int(n):
                    notes.append((lines[i3].strip(), f"N of Valid Cases is the crosstab’s {int(n):,}, not {int(nv):,}."))
                break
            want = want_of.get(re.sub(r"chi-?\s?square", "chi-square", name.lower()))
            v_t = re.sub(r"[a-e]$", "", (c3[kv] if kv < len(c3) else "").strip())
            d_t = (c3[kd] if kd < len(c3) else "").strip()
            if want is None or _ct_num(v_t) is None:
                continue
            if name.lower() != "continuity correction" and re.match(r"^\d{1,3}$", d_t) and int(d_t) != df:
                notes.append((lines[i3].strip(), f"{name}: a crosstab of {len(obs)} rows and {len(obs[0])} columns has {df} df, not {d_t}."))
            elif abs(_ct_num(v_t) - want) > 0.5 * 10 ** -_dec(v_t) + 1e-9:
                notes.append((lines[i3].strip(), f"{name}: the crosstab’s counts give {want:.{_dec(v_t)}f}, not {v_t}."))
        break
    return notes


# 6.6, round 157 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a chi-square test on small expected counts. SPSS's
# footnote under its Chi-Square Tests ("a. 2 cells (33.3%) have expected count less than 5. The minimum expected count is
# 3.45.") worked out again from the crosstab's own counts; and Cochran's rule, which that footnote is there for: Pearson's
# chi-square is trusted only when no more than 20% of the cells have an expected count below 5 and none is below 1. A
# crosstab tested with it (SPSS's, with its Chi-Square Tests; or a typed cross-table with its χ²) whose counts break the rule
# is told so, with the usual way (Fisher's exact test for a 2 × 2 table; joining small categories, or an exact test, for a
# larger one), unless the writing beside it says it took one.
_CT_FOOT = re.compile(r"^(?P<l>[a-e])\.\s+(?P<c>\d+)\s+cells?\s+\((?P<pc>\d+(?:\.\d+)?)%\)\s+ha(?:ve|s)\s+(?:an\s+)?expected\s+count\s+less\s+than\s+5\.\s+"
                      r"the\s+minimum\s+expected\s+count\s+is\s+(?P<m>\d+(?:\.\d+)?)\.?$", re.I)
_CT_DONE = re.compile(r"\bfisher|\bexact\s+(?:test|p\b|sig|significance|method)|\bmonte\s+carlo\b|"
                      r"\b(?:merged|combined|collapsed|regrouped|recoded|pooled|joined|grouped\s+together)\b", re.I)


def _ct_cochran(obs):
    """(cells with an expected count below 5, cells, the smallest expected count) of a table of counts; None when a row or a
    column is empty."""
    tot = sum(sum(r) for r in obs)
    cols = [sum(r[j] for r in obs) for j in range(len(obs[0]))]
    if tot <= 0 or any(c == 0 for c in cols) or any(sum(r) == 0 for r in obs):
        return None
    exp = [sum(r) * cols[j] / tot for r in obs for j in range(len(cols))]
    return sum(1 for e in exp if e < 5 - 1e-12), len(exp), min(exp)


def _ct_rule_note(name, obs, k5, cells, mn):
    two = len(obs) == 2 and len(obs[0]) == 2
    way = "Fisher’s exact test is the usual way for a 2 × 2 table" if two else "joining small categories, or an exact test, is the usual way"
    small = " and the smallest is below 1" if mn < 1 else ""
    return (f"{name}{k5} of its {cells} cells ({100.0 * k5 / cells:.1f}%) have an expected count below 5{small}. Pearson’s chi-square is "
            f"trusted only when no more than 20% do and none is below 1 (Cochran’s rule): {way}.")


def _ct_done_near(lines, a, z):
    """Whether the prose around a table (not the table itself) says it took Fisher's exact test or joined categories."""
    for i in range(max(0, a), min(len(lines), z)):
        if _ts_kind_line(lines[i]) == "P" and _CT_DONE.search(lines[i]):
            return True
    return False


def _ct_expect(text, t, obs):
    """Round 157's notes on an SPSS crosstab: its Chi-Square Tests' footnote, and Cochran's rule."""
    lines, notes = str(text or "").split("\n"), []
    got = _ct_cochran(obs)
    if not got:
        return notes
    k5, cells, mn = got
    title = next((j for j in range(min(t["end"], len(lines)) - 1, -1, -1) if _CT_TITLE.match(lines[j].strip())), None)
    m = _CT_TITLE.match(lines[title].strip()) if title is not None else None
    name = f"{m.group('r').strip()} * {m.group('c').strip()}: " if m else ""
    tested, foot, z = False, None, min(len(lines), t["end"] + 30)
    for i in range(t["end"], z):
        if _CT_TITLE.match(lines[i].strip()):
            z = i
            break
        low = [(x or "").strip().lower() for x in _cells_keep(lines[i])]
        if "value" in low and "df" in low:
            tested = True
        fm = _CT_FOOT.match(lines[i].strip())
        if fm and foot is None:
            foot = (fm, lines[i].strip())
    if foot:
        fm, line = foot
        c, pc, mm = int(fm.group("c")), float(fm.group("pc")), fm.group("m")
        if c != k5:
            notes.append((line, f"{name}the footnote counts {c} cell{'s' if c != 1 else ''} with an expected count below 5; the crosstab’s counts "
                                f"give {k5} of its {cells}."))
        elif abs(pc - 100.0 * c / cells) > 0.5 * 10 ** -_dec(fm.group("pc")) + 1e-9:
            notes.append((line, f"{name}the footnote’s {fm.group('pc')}% should be {c} of {cells} cells, {100.0 * c / cells:.1f}%."))
        elif abs(float(mm) - mn) > 0.5 * 10 ** -_dec(mm) + 1e-9:
            notes.append((line, f"{name}the crosstab’s counts give a minimum expected count of {mn:.{max(1, _dec(mm))}f}, not {mm}."))
    if (tested or foot) and (k5 > 0.2 * cells + 1e-9 or mn < 1) and not _ct_done_near(lines, (title or t["end"]) - 15, z + 15):
        notes.append(((foot[1] if foot else lines[title].strip() if title is not None else ""), _ct_rule_note(name, obs, k5, cells, mn)))
    return notes


def spss_crosstab_check(text):
    out = []
    for t in _crosstabs(text):
        body = [b for b in t["blocks"] if b["label"].lower() != "total" and "count" in b["rows"]]
        tot = next((b for b in t["blocks"] if b["label"].lower() == "total" and "count" in b["rows"]), None)
        if len(body) < 2 or not tot:
            continue
        counts = [[_ct_num(x) for x in b["rows"]["count"][1]] for b in body]
        w = len(counts[0])
        if w < 3 or any(len(c) != w or any(x is None or x != int(x) or x < 0 for x in c) for c in counts):
            continue                                 # (every row: its counts, then its Total)
        tcounts = [_ct_num(x) for x in tot["rows"]["count"][1]]
        if len(tcounts) != w or any(x is None for x in tcounts):
            continue
        names = t["names"] if len(t["names"]) == w - 1 else [f"column {j + 1}" for j in range(w - 1)]
        notes = []
        for b, c in zip(body, counts):
            if abs(sum(c[:-1]) - c[-1]) > 1e-9:
                notes.append((b["rows"]["count"][2], f"{b['label']}: its counts add up to {int(sum(c[:-1])):,}, not its Total of {int(c[-1]):,}."))
        for j in range(w if not notes else w - 1):      # (a row that doesn't add up already explains the Total column)
            s_j = sum(c[j] for c in counts)
            if abs(s_j - tcounts[j]) > 1e-9:
                col = names[j] if j < w - 1 else "Total"
                notes.append((tot["rows"]["count"][2], f"Total: the {col} column adds up to {int(s_j):,}, not {int(tcounts[j]):,}."))
        if not notes:                                # (percentages and the tests only on counts that add up)
            n = tcounts[-1]
            for b in body + [tot]:
                c = [_ct_num(x) for x in b["rows"]["count"][1]]
                for stat, (stat_label, texts, line) in b["rows"].items():
                    if stat == "count" or len(texts) != w:
                        continue
                    if stat == "expected count":
                        want, what, unit = [c[-1] * tcounts[j] / n for j in range(w - 1)] + [c[-1]], "its row’s Total times its column’s, over all", ""
                    elif stat == "% of total":
                        want, what, unit = [100.0 * c[j] / n for j in range(w)], f"its count of all {int(n):,}", "%"
                    elif stat.startswith("% within ") and stat[9:].strip() == t["rv"]:
                        want, what, unit = [100.0 * c[j] / c[-1] if c[-1] else None for j in range(w)], f"its count of the row’s {int(c[-1]):,}", "%"
                    elif stat.startswith("% within ") and stat[9:].strip() == t["cv"]:
                        want, what, unit = [100.0 * c[j] / tcounts[j] if tcounts[j] else None for j in range(w)], "its count of its column’s total", "%"
                    else:
                        continue
                    for j, (tx, x) in enumerate(zip(texts, want)):
                        v = _ct_num(tx)
                        if v is None or x is None:
                            continue
                        d = _dec(tx.strip().rstrip("%").strip())
                        if abs(v - x) > 0.5 * 10 ** -d + 0.05 + 1e-9:
                            col = names[j] if j < w - 1 else "Total"
                            notes.append((line, f"{b['label']}: its “{stat_label}” under {col} is {what}, {x:.{max(1, d)}f}{unit}, not {tx}."))
                            break
            notes += _ct_tests(text, t["end"], [[int(x) for x in c[:-1]] for c in counts], n)
            try:                                 # 6.6, round 157: its footnote, and Cochran's rule
                notes += _ct_expect(text, t, [[int(x) for x in c[:-1]] for c in counts])
            except Exception as e:
                _fault("expected counts", e)
        seen = set()
        for line, note in notes:
            if note in seen:
                continue
            seen.add(note)
            out.append({"s": line[:500], "note": note, "kind": "test" if " give " in note or " df, " in note or "Cochran" in note or "footnote" in note
                        or "expected count" in note else "percent", "decision": False})
    return out



# 6.6, round 132 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): the stars of an APA correlation table. Its
# note says what they mean ("*p < .05. **p < .01."), its caption or note gives N, and each starred r must reach the level its
# stars claim (r's p on N − 2 df, two-tailed unless the note says one-tailed, allowed r's rounding): ".12*" with N = 150 is
# p = .14, not below .05. A star left off is not judged (a table may star only some), nor a table whose N it doesn't give.
_ST_LEGEND = re.compile(r"(?P<s>\*{1,3}|†)\s*p\s*(?P<rel><|≤|<=)\s*(?P<v>0?\.\d+)")
_ST_N = re.compile(r"\bN\s*=\s*(?P<n>\d{1,3}(?:,\d{3})*|\d+)\b")
_ST_ROW = re.compile(r"^\s*(?P<i>\d{1,2})\s*[.)]\s*(?P<lab>.*\S)\s*$")
_ST_DIAG = ("—", "–", "-", "1", "1.00", "1.0", "")


def star_table_check(text):
    out, lines = [], str(text or "").split("\n")
    for h, line in enumerate(lines):
        cells = _cells_keep(line)
        nums = [(k, int(c)) for k, c in enumerate(cells) if re.match(r"^\d{1,2}$", (c or "").strip())]
        if len(nums) < 2 or [v for _k, v in nums] != list(range(1, len(nums) + 1)):
            continue                                   # (a heading "1 2 3 …": the matrix's columns)
        col_at = {v: k for k, v in nums}
        rows, j = {}, h + 1
        while j < len(lines) and j < h + 40:
            c2 = _cells_keep(lines[j])
            m = _ST_ROW.match((c2[0] or "").strip()) if c2 else None
            if not m:
                if rows:
                    break
                j += 1
                continue
            rows[int(m.group("i"))] = (m.group("lab").strip(), c2, lines[j].strip())
            j += 1
        if len(rows) < 2:
            continue
        around = "\n".join(lines[max(0, h - 5):min(len(lines), j + 6)])
        legend = {}
        for lm in _ST_LEGEND.finditer(around):
            legend.setdefault(lm.group("s"), float(lm.group("v")))
        nm = _ST_N.search(around)
        if not legend or not nm:
            continue
        n = int(nm.group("n").replace(",", ""))
        one_tailed = bool(re.search(r"one[-\s]tailed", around, re.I)) and not re.search(r"two[-\s]tailed", around, re.I)
        if n < 4:
            continue
        for i, (lab, c2, row_line) in sorted(rows.items()):
            for v, k in col_at.items():
                if v >= i or k >= len(c2) or v not in rows:
                    continue
                cell = (c2[k] or "").strip()
                mm = re.match(r"^(?P<r>[-−–]?\.\d{1,3}|[-−–]?0\.\d{1,3})(?P<s>\*{1,3}|†)?$", cell)
                if not mm or not mm.group("s") or mm.group("s") not in legend:
                    continue
                r_t = mm.group("r").replace("−", "-").replace("–", "-")
                r = abs(float(r_t))
                h_ = 0.5 * 10 ** -_dec(r_t.lstrip("-"))
                p_min = p_r(min(r + h_, 0.999999), n - 2)
                if p_min is None:
                    continue
                if one_tailed:
                    p_min /= 2
                lvl = legend[mm.group("s")]
                if p_min >= lvl - 1e-12:
                    real = p_r(r, n - 2) / (2 if one_tailed else 1)
                    out.append({"s": row_line[:500], "kind": "test", "decision": True,
                                "note": f"{rows[v][0]} and {lab}: r = {mm.group('r')} with N = {n} gives {_fmt_p(real, mm.group('r').lstrip('-−–').startswith('0'))}, "
                                        f"not below {_st_level(lvl)}, though it carries “{mm.group('s')}”."})
    return out


def _st_level(v):
    """A star's level as a note writes it: .05, .01, .001."""
    s = f"{v:.3f}".rstrip("0")
    return s[1:] if s.startswith("0") else s



# 6.6, round 133 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): SPSS's independent-samples t-test, its two
# tables read together. Group Statistics: each Std. Error Mean is the SD over √N. Independent Samples Test: Levene's Sig. is
# its F's p on (1, n₁ + n₂ − 2); the Mean Difference is the two means' difference; with equal variances assumed, t comes from
# the pooled SD and df is n₁ + n₂ − 2, without, from each group's own SD with Welch's df; each Sig. (2-tailed) is t's p on its
# df; the Std. Error Difference is what the SDs and Ns give; the 95% confidence interval is the Mean Difference ± t(.975, df)
# times it. Every figure is allowed the rounding of those it comes from; one note a row.
_TG_COL = {"n": re.compile(r"^N$"), "m": re.compile(r"^mean$", re.I), "sd": re.compile(r"^std\.?\s*deviation$", re.I),
           "sem": re.compile(r"^std\.?\s*error\s+mean$", re.I)}
_TI_COL = {"f": re.compile(r"^F$"), "t": re.compile(r"^t$"), "df": re.compile(r"^df$", re.I), "p": re.compile(r"^sig\.?\s*\(2-tailed\)$", re.I),
           "md": re.compile(r"^mean\s+difference$", re.I), "sed": re.compile(r"^std\.?\s*error\s+difference$", re.I),
           "lo": re.compile(r"^lower$", re.I), "hi": re.compile(r"^upper$", re.I)}


def _t_crit(df, alpha=0.05):
    """t such that its two-tailed p is alpha (bisection on p_t)."""
    lo, hi = 0.0, 1000.0
    for _ in range(80):
        mid = (lo + hi) / 2
        if p_t(mid, df) > alpha:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def _tt_groups(lines):
    """Group Statistics rows by variable: {variable (lower): [(group, n, mean text, sd text, sem text, line)]}."""
    out, cols, var = {}, None, ""
    for line in lines:
        cells = _cells_keep(line)
        heads = {}
        for k, c in enumerate(cells):
            for key, rx in _TG_COL.items():
                if rx.match((c or "").strip()) and key not in heads:
                    heads[key] = k
        if {"n", "m", "sd", "sem"} <= set(heads):          # (SPSS's Group Statistics; its Descriptives have "Std. Error")
            cols, var = (heads if heads["n"] >= 2 else None), ""     # (6.6, round 146: no group column before N = One-Sample Statistics)
            continue
        if not cols or not any((c or "").strip() for c in cells):
            if cols and out:
                cols = None
            continue
        first = (cells[0] or "").strip()
        if first:
            var = first
        group = (cells[1] or "").strip() if len(cells) > 1 else ""
        n_t = (cells[cols["n"]] if cols["n"] < len(cells) else "").strip().replace(",", "")
        if not var or not group or not _INT_CELL.match(n_t):
            cols = None if out else cols
            continue
        get = lambda key: (cells[cols[key]] if key in cols and cols[key] < len(cells) else "").strip()
        out.setdefault(var.lower(), []).append((group, int(n_t), get("m"), get("sd"), get("sem"), line.strip()))
    return out


def spss_ttest_check(text, counted=None):
    out, lines, n_judged = [], str(text or "").split("\n"), 0
    groups = _tt_groups(lines)
    if not groups:
        return out
    h_of = lambda t: 0.5 * 10 ** -_dec(t.lstrip("-−"))
    for var, rows in groups.items():                         # Std. Error Mean = SD / √N
        for g, n, m_t, sd_t, sem_t, line in rows:
            sd, sem = _num_cell(sd_t), _num_cell(sem_t)
            if sd is None or sem is None or n < 2:
                continue
            lo, hi = (sd - h_of(sd_t)) / math.sqrt(n), (sd + h_of(sd_t)) / math.sqrt(n)
            if sem + h_of(sem_t) < lo - 1e-12 or sem - h_of(sem_t) > hi + 1e-12:
                out.append({"s": line[:500], "kind": "value", "decision": False,
                            "note": f"{g}: the Std. Error Mean is its SD over √N, {_like(sd / math.sqrt(n), sem_t)}, not {sem_t}."})
    cols, var_now = None, ""
    for i, raw in enumerate(lines):
        cells = _cells_keep(raw)
        heads = {}
        for k, c in enumerate(cells):
            for key, rx in _TI_COL.items():
                if rx.match((c or "").strip()) and key not in heads:
                    heads[key] = k
        if {"t", "df", "p", "md"} <= set(heads):
            cols, var_now = heads, ""
            nxt = _cells_keep(lines[i + 1]) if i + 1 < len(lines) else []
            for k, c in enumerate(nxt):                      # (the interval's Lower and Upper: a line of their own)
                for key in ("lo", "hi"):
                    if key not in cols and _TI_COL[key].match((c or "").strip()):
                        cols[key] = k
            continue
        if not cols:
            continue
        lab = " ".join((c or "").strip().lower() for c in cells[:2])
        assumed, welch = "equal variances assumed" in lab, "equal variances not assumed" in lab
        if not (assumed or welch):
            if any((c or "").strip() for c in cells) and not re.search(r"\b(?:lower|upper)\b", " ".join(cells).lower()):
                cols = None                                  # (a line that isn't one of its rows: the table has ended)
            continue
        if (cells[0] or "").strip() and "equal variances" not in (cells[0] or "").lower():
            var_now = (cells[0] or "").strip().lower()
        rows = groups.get(var_now, [])
        if len(rows) != 2:
            continue
        (g1, n1, m1_t, s1_t, _e1, _l1), (g2, n2, m2_t, s2_t, _e2, _l2) = rows
        m1, m2, s1, s2 = (_num_cell(x) for x in (m1_t, m2_t, s1_t, s2_t))
        if None in (m1, m2, s1, s2) or n1 < 2 or n2 < 2:
            continue
        get = lambda key: (cells[cols[key]] if key in cols and cols[key] < len(cells) else "").strip()
        line = raw.strip()
        name = "Equal variances assumed" if assumed else "Equal variances not assumed"
        note, flip = None, False
        # the Mean Difference
        md_t = get("md")
        md = _num_cell(md_t)
        d_lo, d_hi = (m1 - h_of(m1_t)) - (m2 + h_of(m2_t)), (m1 + h_of(m1_t)) - (m2 - h_of(m2_t))
        if md is not None and (md + h_of(md_t) < d_lo - 1e-9 or md - h_of(md_t) > d_hi + 1e-9):
            note = f"{name}: the Mean Difference is {m1_t} − {m2_t} = {_like(m1 - m2, md_t)}, not {md_t}."
        # the standard error and t from the SDs and Ns
        corners = [(a, b) for a in (max(s1 - h_of(s1_t), 1e-9), s1 + h_of(s1_t)) for b in (max(s2 - h_of(s2_t), 1e-9), s2 + h_of(s2_t))]
        if assumed:
            se_of = lambda a, b: math.sqrt(((n1 - 1) * a * a + (n2 - 1) * b * b) / (n1 + n2 - 2) * (1 / n1 + 1 / n2))
        else:
            se_of = lambda a, b: math.sqrt(a * a / n1 + b * b / n2)
        ses = [se_of(a, b) for a, b in corners]
        sed_t = get("sed")
        sed = _num_cell(sed_t)
        if note is None and sed is not None and (sed + h_of(sed_t) < min(ses) - 1e-9 or sed - h_of(sed_t) > max(ses) + 1e-9):
            note = f"{name}: the Std. Error Difference the SDs and Ns give is {_like(se_of(s1, s2), sed_t)}, not {sed_t}."
        t_t = get("t")
        t = _num_cell(t_t)
        if note is None and t is not None:
            t_lo = (min(abs(d_lo), abs(d_hi)) if d_lo * d_hi > 0 else 0.0) / max(ses)    # 6.6, round 136: the smaller |difference|
            t_hi = max(abs(d_lo), abs(d_hi)) / min(ses)                                    # (it took the larger where the means fell)
            if abs(t) + h_of(t_t) < t_lo - 1e-9 or abs(t) - h_of(t_t) > t_hi + 1e-9 or (t * (m1 - m2) < 0 and d_lo * d_hi > 0):
                note = f"{name}: from the two means, SDs and Ns, t is about {(m1 - m2) / se_of(s1, s2):.3f}, not {t_t}."
        df_t = get("df")
        df = _num_cell(df_t)
        if note is None and df is not None:
            if assumed and abs(df - (n1 + n2 - 2)) > 1e-9:
                note = f"{name}: two groups of {n1} and {n2} have {n1 + n2 - 2} df, not {df_t}."
            elif welch:
                ws = []
                for a, b in corners:
                    v1, v2 = a * a / n1, b * b / n2
                    ws.append((v1 + v2) ** 2 / (v1 * v1 / (n1 - 1) + v2 * v2 / (n2 - 1)))
                if df + h_of(df_t) < min(ws) - 1e-6 or df - h_of(df_t) > max(ws) + 1e-6:
                    v1, v2 = s1 * s1 / n1, s2 * s2 / n2
                    note = f"{name}: Welch's df from the SDs and Ns is {(v1 + v2) ** 2 / (v1 * v1 / (n1 - 1) + v2 * v2 / (n2 - 1)):.3f}, not {df_t}."
        p_t_ = get("p")
        pc = _p_cell(p_t_)
        if note is None and t is not None and df and pc is not None:
            rel, pv, dp = pc
            plo, phi = p_t(abs(t) + h_of(t_t), df), p_t(max(abs(t) - h_of(t_t), 0.0), df)
            n_judged += 1
            if not _fits(rel, pv, dp, plo, phi):
                real = p_t(abs(t), df)
                flip = (real < 0.05) != (pv < 0.05 if rel == "=" else rel in ("<", "≤"))
                note = f"{name}: t = {t_t} on {df_t} df gives {_fmt_p(real, _p_lead(p_t_))}, not Sig. (2-tailed) {_p_said(p_t_)}."
        lo_t, hi_t = get("lo"), get("hi")
        lo_v, hi_v = _num_cell(lo_t), _num_cell(hi_t)
        if note is None and None not in (md, sed, df, lo_v, hi_v) and df > 0:
            tc = _t_crit(df)
            slack = h_of(md_t) + tc * h_of(sed_t)
            for v, t_x, want, end, sign in ((lo_v, lo_t, md - tc * sed, "lower", "−"), (hi_v, hi_t, md + tc * sed, "upper", "+")):
                if abs(v - want) > slack + h_of(t_x) + 1e-9:
                    note = (f"{name}: the 95% interval’s {end} end is about {_like(want, t_x)} (the Mean Difference {sign} t(.975, {df_t}) "
                            f"× the Std. Error Difference), not {t_x}.")
                    break
        if assumed and note is None:                        # Levene's Sig.: F's p on (1, n₁ + n₂ − 2)
            f_t = get("f")
            f = _num_cell(f_t)
            k_sig = next((k for k in range(cols["f"] + 1, cols["t"]) if k < len(cells) and _p_cell(cells[k])), None) if "f" in cols else None
            if f is not None and k_sig is not None:
                lp_t = cells[k_sig].strip()
                rel, pv, dp = _p_cell(lp_t)
                plo, phi = p_f(f + h_of(f_t), 1, n1 + n2 - 2), p_f(max(f - h_of(f_t), 0), 1, n1 + n2 - 2)
                n_judged += 1
                if plo is not None and phi is not None and not _fits(rel, pv, dp, plo, phi):
                    note = f"Levene's test: F = {f_t} on (1, {n1 + n2 - 2}) gives {_fmt_p(p_f(f, 1, n1 + n2 - 2), _p_lead(lp_t))}, not Sig. {_p_said(lp_t)}."
        if note:
            out.append({"s": line[:500], "note": note, "kind": "test", "decision": flip})
    if counted is not None:
        counted[0] += n_judged
    return out



# 6.6, round 134 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): SPSS's one-way ANOVA, its Descriptives and its
# ANOVA table read together. Descriptives: each Std. Error is the SD over √N; each 95% interval for the mean is the mean ±
# t(.975, N − 1) times it; the Total row's N is the groups' Ns added and its mean their weighted mean. ANOVA: the Between
# Groups Sum of Squares is what the groups' means and Ns give (Σ nᵢ(mᵢ − m)²), the Within Groups one what their SDs and Ns give
# (Σ (nᵢ − 1)sᵢ²), and their df are k − 1 and N − k. Each figure is allowed the rounding of those it comes from.
_AD_COL = {"n": re.compile(r"^N$"), "m": re.compile(r"^mean$", re.I), "sd": re.compile(r"^std\.?\s*deviation$", re.I),
           "se": re.compile(r"^std\.?\s*error$", re.I), "lo": re.compile(r"^lower\s+bound$", re.I), "hi": re.compile(r"^upper\s+bound$", re.I)}


def _ad_tables(lines):
    """SPSS's Descriptives tables: [(header line, [(label, n, mean text, sd text, se text, lo text, hi text, line)], end)]."""
    found, i = [], 0
    while i < len(lines):
        cells = _cells_keep(lines[i])
        heads = {}
        for k, c in enumerate(cells):
            for key, rx in _AD_COL.items():
                if rx.match((c or "").strip()) and key not in heads:
                    heads[key] = k
        if not {"n", "m", "sd", "se"} <= set(heads):
            i += 1
            continue
        nxt = _cells_keep(lines[i + 1]) if i + 1 < len(lines) else []
        start = i + 1
        if any(_AD_COL["lo"].match((c or "").strip()) for c in nxt):
            for k, c in enumerate(nxt):
                for key in ("lo", "hi"):
                    if key not in heads and _AD_COL[key].match((c or "").strip()):
                        heads[key] = k
            start = i + 2
        rows, j = [], start
        while j < len(lines):
            c2 = _cells_keep(lines[j])
            get = lambda key: (c2[heads[key]] if key in heads and heads[key] < len(c2) else "").strip()
            lab = (c2[0] or "").strip() if c2 else ""
            if not lab or not _INT_CELL.match(get("n").replace(",", "")):
                break
            rows.append((lab, int(get("n").replace(",", "")), get("m"), get("sd"), get("se"), get("lo"), get("hi"), lines[j].strip()))
            j += 1
            if lab.lower() == "total":
                break
        if len(rows) >= 3:
            found.append((i, rows, j))
        i = max(j, i + 1)
    return found


def spss_anova_desc_check(text):
    out, lines = [], str(text or "").split("\n")
    h_of = lambda t: 0.5 * 10 ** -_dec(str(t).lstrip("-−"))
    for _h, rows, end in _ad_tables(lines):
        groups = [r for r in rows if r[0].lower() != "total"]
        total = next((r for r in rows if r[0].lower() == "total"), None)
        notes, doubt, total_ok = [], False, True
        if total and len(groups) >= 2 and all(_num_cell(g[2]) is not None for g in groups):
            n_sum = sum(g[1] for g in groups)
            if total[1] != n_sum:
                notes.append((total[7], f"Total: the groups’ Ns add up to {n_sum:,}, not {total[1]:,}."))
                total_ok = False
        for lab, n, m_t, sd_t, se_t, lo_t, hi_t, line in rows:
            m, sd, se = _num_cell(m_t), _num_cell(sd_t), _num_cell(se_t)
            if None in (m, sd) or n < 2 or (lab.lower() == "total" and not total_ok):
                continue
            lo_v, hi_v = _num_cell(lo_t), _num_cell(hi_t)
            if lo_v is not None and hi_v is not None and abs((lo_v + hi_v) / 2 - m) > h_of(lo_t) + h_of(m_t) + 1e-9:
                notes.append((line, f"{lab}: its mean ({m_t}) isn’t the middle of its 95% interval ({lo_t} to {hi_t}, whose middle is {_like((lo_v + hi_v) / 2, m_t)})."))
                doubt = doubt or lab.lower() != "total"
                total_ok = total_ok and lab.lower() != "total"
                continue
            if se is not None:
                lo_, hi_ = (sd - h_of(sd_t)) / math.sqrt(n), (sd + h_of(sd_t)) / math.sqrt(n)
                if se + h_of(se_t) < lo_ - 1e-12 or se - h_of(se_t) > hi_ + 1e-12:
                    notes.append((line, f"{lab}: the Std. Error is its SD over √N, {_like(sd / math.sqrt(n), se_t)}, not {se_t}."))
                    continue
                if lo_v is not None and hi_v is not None:
                    tc = _t_crit(n - 1)
                    slack = h_of(m_t) + tc * h_of(se_t)
                    for v, tx, want, end_, sign in ((lo_v, lo_t, m - tc * se, "lower", "−"), (hi_v, hi_t, m + tc * se, "upper", "+")):
                        if abs(v - want) > slack + h_of(tx) + 1e-9:
                            notes.append((line, f"{lab}: the 95% interval’s {end_} bound is about {_like(want, tx)} (the mean {sign} t(.975, {n - 1}) "
                                                f"× its Std. Error), not {tx}."))
                            break
        if total and total_ok and not doubt and len(groups) >= 2 and all(_num_cell(g[2]) is not None for g in groups):
            n_sum = sum(g[1] for g in groups)
            wm = sum(g[1] * _num_cell(g[2]) for g in groups) / n_sum
            slack = sum(g[1] * h_of(g[2]) for g in groups) / n_sum + h_of(total[2])
            tm = _num_cell(total[2])
            if tm is not None and abs(tm - wm) > slack + 1e-9:
                notes.append((total[7], f"Total: the mean of all is the groups’ means weighted by their Ns, {_like(wm, total[2])}, not {total[2]}."))
        # ---- the ANOVA table that follows (within 12 lines)
        if not doubt and len(groups) >= 2 and all(_num_cell(g[2]) is not None and _num_cell(g[3]) is not None for g in groups):
            ns = [g[1] for g in groups]
            ms = [_num_cell(g[2]) for g in groups]
            sds = [_num_cell(g[3]) for g in groups]
            N, k = sum(ns), len(groups)
            grand = sum(n * m for n, m in zip(ns, ms)) / N
            ssb = sum(n * (m - grand) ** 2 for n, m in zip(ns, ms))
            ssb_tol = sum(2 * n * abs(m - grand) * h_of(g[2]) + n * h_of(g[2]) ** 2 for n, m, g in zip(ns, ms, groups))
            ssw = sum((n - 1) * s * s for n, s in zip(ns, sds))
            ssw_tol = sum(2 * (n - 1) * s * h_of(g[3]) + (n - 1) * h_of(g[3]) ** 2 for n, s, g in zip(ns, sds, groups))
            cols = None
            for i2 in range(end, min(len(lines), end + 12)):
                c2 = _cells_keep(lines[i2])
                heads = {}
                for kk, c in enumerate(c2):
                    for key, rx in _AN_COL.items():
                        if rx.match((c or "").strip()) and key not in heads:
                            heads[key] = kk
                if "ss" in heads and "df" in heads:
                    cols = heads
                    continue
                if _ad_tables(lines[i2:i2 + 3]) or (cols is None and _CT_TITLE.match(lines[i2].strip())):
                    break
                if not cols:
                    continue
                lab = _row_label(c2)[1].lower()
                kind = "between" if re.match(r"^between(?:\s+groups?)?$", lab) else "within" if re.match(r"^within(?:\s+groups?)?$", lab) else None
                if not kind:
                    if lab == "total":
                        break
                    continue
                ss_t = (c2[cols["ss"]] if cols["ss"] < len(c2) else "").strip()
                df_t = (c2[cols["df"]] if cols["df"] < len(c2) else "").strip()
                ss, dfv = _num_cell(ss_t), _num_cell(df_t)
                want_df = k - 1 if kind == "between" else N - k
                want_ss, tol = (ssb, ssb_tol) if kind == "between" else (ssw, ssw_tol)
                name = "Between Groups" if kind == "between" else "Within Groups"
                if dfv is not None and abs(dfv - want_df) > 1e-9:
                    notes.append((lines[i2].strip(), f"{name}: {k} groups of {N:,} people have {want_df} df, not {df_t}."))
                elif ss is not None and abs(ss - want_ss) > tol + h_of(ss_t) + 1e-9:
                    what = "means and Ns" if kind == "between" else "SDs and Ns"
                    notes.append((lines[i2].strip(), f"{name}: the Descriptives’ {what} give a Sum of Squares of about {want_ss:.{max(1, _dec(ss_t))}f}, not {ss_t}."))
        for line, note in notes:
            out.append({"s": line[:500], "note": note, "kind": "test" if "Sum of Squares" in note or " df, " in note else "value", "decision": False})
    return out



# 6.6, round 135 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): SPSS's paired-samples t-test (before and after
# a training, say), its three tables read together. Paired Samples Correlations: each Sig. is r's p on N − 2 df. Paired Samples
# Test: the Mean of the differences is the two means' difference; their Std. Deviation is what the two SDs and their correlation
# give (√(s₁² + s₂² − 2rs₁s₂)); the Std. Error Mean is that over √N; t is the mean over it, on N − 1 df; its Sig. t's p; the 95%
# interval the mean ± t(.975, N − 1) times the Std. Error. (Each pair's own Std. Error Means: round 133.) Rounding allowed.
_PS_COL = {"m": re.compile(r"^mean$", re.I), "sd": re.compile(r"^std\.?\s*deviation$", re.I), "sem": re.compile(r"^std\.?\s*error\s+mean$", re.I),
           "lo": re.compile(r"^lower$", re.I), "hi": re.compile(r"^upper$", re.I), "t": re.compile(r"^t$"), "df": re.compile(r"^df$", re.I),
           "p": re.compile(r"^sig\.?\s*\(2-tailed\)$", re.I), "n": re.compile(r"^N$"), "r": re.compile(r"^correlation$", re.I),
           "sig": re.compile(r"^sig\.?$", re.I)}
_PS_PAIR = re.compile(r"^pair\s*\d{1,2}$", re.I)


def _ps_read(lines):
    """{"stats": {pair: [(name, mean, n, sd)]}, "corr": {pair: (n, r text, sig text, line)}, "test": {pair: (cells by key, line)}}."""
    got = {"stats": {}, "corr": {}, "test": {}}
    kind, cols, pair = None, {}, ""
    for i, raw in enumerate(lines):
        cells = _cells_keep(raw)
        heads = {}
        for k, c in enumerate(cells):
            for key, rx in _PS_COL.items():
                if rx.match((c or "").strip()) and key not in heads:
                    heads[key] = k
        low = raw.lower()
        if {"m", "n", "sd", "sem"} <= set(heads) and "t" not in heads:
            kind, cols, pair = "stats", heads, ""
            continue
        if {"n", "r", "sig"} <= set(heads):
            kind, cols, pair = "corr", heads, ""
            continue
        if "paired differences" in low or ({"t", "df", "p"} <= set(heads) and kind == "test-head"):
            kind, cols, pair = "test-head", dict(heads), ""
            continue
        if kind == "test-head":
            if not _PS_PAIR.match((cells[0] or "").strip() if cells else ""):
                for k, v in heads.items():
                    cols.setdefault(k, v)
                continue
            kind = "test"
        if kind is None:
            continue
        first = (cells[0] or "").strip() if cells else ""
        if _PS_PAIR.match(first):
            pair = first.lower()
        elif first or not any((c or "").strip() for c in cells):
            kind = None if first else kind
            continue
        if not pair:
            continue
        get = lambda key: (cells[cols[key]] if key in cols and cols[key] < len(cells) else "").strip()
        name = (cells[1] or "").strip() if len(cells) > 1 else ""
        if kind == "stats":
            m, sd, n_t = _num_cell(get("m")), _num_cell(get("sd")), get("n").replace(",", "")
            if m is not None and sd is not None and _INT_CELL.match(n_t):
                got["stats"].setdefault(pair, []).append((name, get("m"), int(n_t), get("sd")))
        elif kind == "corr":
            if _INT_CELL.match(get("n").replace(",", "")):
                got["corr"][pair] = (int(get("n").replace(",", "")), get("r"), get("sig"), raw.strip())
        elif kind == "test":
            got["test"][pair] = ({k: get(k) for k in ("m", "sd", "sem", "lo", "hi", "t", "df", "p")}, raw.strip(), name)
    return got


def spss_paired_check(text, counted=None):
    out, n_judged = [], 0
    got = _ps_read(str(text or "").split("\n"))
    h_of = lambda t: 0.5 * 10 ** -_dec(str(t).lstrip("-−"))
    for pair, (n, r_t, sig_t, line) in got["corr"].items():
        r, pc = _num_cell(r_t), _p_cell(sig_t)
        if r is None or pc is None or n < 4 or abs(r) > 1:
            continue
        rel, pv, dp = pc
        plo, phi = p_r(min(abs(r) + h_of(r_t), 0.999999), n - 2), p_r(max(abs(r) - h_of(r_t), 0.0), n - 2)
        if plo is None or phi is None:
            continue
        n_judged += 1
        if not _fits(rel, pv, dp, plo, phi):
            real = p_r(abs(r), n - 2)
            out.append({"s": line[:500], "kind": "test", "decision": (real < 0.05) != (pv < 0.05 if rel == "=" else rel in ("<", "≤")),
                        "note": f"{pair.title()}: a correlation of {r_t} with N = {n} gives {_fmt_p(real, _p_lead(sig_t))}, not Sig. {_p_said(sig_t)}."})
    for pair, (v, line, name) in got["test"].items():
        st = got["stats"].get(pair, [])
        note = None
        md, sdd, sem, t, df = (_num_cell(v[k]) for k in ("m", "sd", "sem", "t", "df"))
        n = st[0][2] if len(st) == 2 and st[0][2] == st[1][2] else (got["corr"].get(pair, (None,))[0])
        if len(st) == 2 and md is not None:
            (_a, m1_t, _n1, s1_t), (_b, m2_t, _n2, s2_t) = st
            m1, m2 = _num_cell(m1_t), _num_cell(m2_t)
            if abs(md - (m1 - m2)) > h_of(v["m"]) + h_of(m1_t) + h_of(m2_t) + 1e-9:
                note = f"{pair.title()}: the Mean of the differences is {m1_t} − {m2_t} = {_like(m1 - m2, v['m'])}, not {v['m']}."
            elif sdd is not None and pair in got["corr"] and _num_cell(got["corr"][pair][1]) is not None:
                r_t = got["corr"][pair][1]
                r, s1, s2 = _num_cell(r_t), _num_cell(s1_t), _num_cell(s2_t)
                vals = []
                for a in (s1 - h_of(s1_t), s1 + h_of(s1_t)):
                    for b in (s2 - h_of(s2_t), s2 + h_of(s2_t)):
                        for rr in (r - h_of(r_t), r + h_of(r_t)):
                            q = a * a + b * b - 2 * rr * a * b
                            if q >= 0:
                                vals.append(math.sqrt(q))
                if vals and (sdd + h_of(v["sd"]) < min(vals) - 1e-9 or sdd - h_of(v["sd"]) > max(vals) + 1e-9):
                    q = s1 * s1 + s2 * s2 - 2 * r * s1 * s2
                    note = (f"{pair.title()}: the two SDs ({s1_t}, {s2_t}) and their correlation ({r_t}) give the differences an SD of about "
                            f"{_like(math.sqrt(max(q, 0)), v['sd'])}, not {v['sd']}.")
        if note is None and n and sdd is not None and sem is not None:
            lo_, hi_ = (sdd - h_of(v["sd"])) / math.sqrt(n), (sdd + h_of(v["sd"])) / math.sqrt(n)
            if sem + h_of(v["sem"]) < lo_ - 1e-12 or sem - h_of(v["sem"]) > hi_ + 1e-12:
                note = f"{pair.title()}: the Std. Error Mean of the differences is their SD over √N, {_like(sdd / math.sqrt(n), v['sem'])}, not {v['sem']}."
        if note is None and n and df is not None and abs(df - (n - 1)) > 1e-9:
            note = f"{pair.title()}: {n} pairs have {n - 1} df, not {v['df']}."
        if note is None and md is not None and sem and t is not None:
            lo_t = (abs(md) - h_of(v["m"])) / (sem + h_of(v["sem"]))
            hi_t = (abs(md) + h_of(v["m"])) / max(sem - h_of(v["sem"]), 1e-12)
            if abs(t) + h_of(v["t"]) < lo_t - 1e-9 or abs(t) - h_of(v["t"]) > hi_t + 1e-9 or (t * md < 0):
                note = f"{pair.title()}: t is the Mean over its Std. Error, about {md / sem:.3f}, not {v['t']}."
        pc = _p_cell(v["p"])
        if note is None and t is not None and df and pc is not None:
            rel, pv, dp = pc
            plo, phi = p_t(abs(t) + h_of(v["t"]), df), p_t(max(abs(t) - h_of(v["t"]), 0.0), df)
            n_judged += 1
            if not _fits(rel, pv, dp, plo, phi):
                real = p_t(abs(t), df)
                note = f"{pair.title()}: t = {v['t']} on {v['df']} df gives {_fmt_p(real, _p_lead(v['p']))}, not Sig. (2-tailed) {_p_said(v['p'])}."
        lo_v, hi_v = _num_cell(v["lo"]), _num_cell(v["hi"])
        if note is None and None not in (md, sem, df, lo_v, hi_v) and df > 0:
            tc = _t_crit(df)
            slack = h_of(v["m"]) + tc * h_of(v["sem"])
            for val, tx, want, end_, sign in ((lo_v, v["lo"], md - tc * sem, "lower", "−"), (hi_v, v["hi"], md + tc * sem, "upper", "+")):
                if abs(val - want) > slack + h_of(tx) + 1e-9:
                    note = (f"{pair.title()}: the 95% interval’s {end_} end is about {_like(want, tx)} (the Mean {sign} t(.975, {v['df']}) × its Std. Error), "
                            f"not {tx}.")
                    break
        if note:
            out.append({"s": line[:500], "note": note, "kind": "test", "decision": False})
    if counted is not None:
        counted[0] += n_judged
    return out



# 6.6, round 142 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): after a one-way ANOVA, SPSS's Test of Homogeneity
# of Variances and its Multiple Comparisons (post hoc), held to the Descriptives and the ANOVA table above them. Levene's df are
# k − 1 and N − k and its Sig. its F's p. Each comparison's Mean Difference (I − J) is the two groups' means' difference; its Std.
# Error is √(MSwithin (1/nᵢ + 1/nⱼ)); its Sig. is LSD's t p on N − k df, Bonferroni's that times the comparisons (at most 1),
# Scheffé's from F = t²/(k − 1) (Tukey's and Games–Howell's Sig. need the studentized range: not judged); its star marks a
# Sig. below .05; and the pair turned round (J − I) holds the same figures with the other sign. One note a row.
_PH_COL = {"md": re.compile(r"^mean\s+difference\s*\(\s*i\s*[-−–]\s*j\s*\)$", re.I), "se": re.compile(r"^std\.?\s*error$", re.I),
           "p": re.compile(r"^sig\.?$", re.I)}
_PH_METHOD = re.compile(r"^(tukey\s+hsd|tukey|lsd|bonferroni|scheff[eé]|games[-\s]howell|dunnett\s*t?|sidak|hochberg)\b", re.I)
_LV_HEAD = re.compile(r"^levene\s+statistic$", re.I)


def _ph_groups(lines):
    """The Descriptives above: {group label (lower): (n, mean, mean text)}, k, N, and MSwithin with its text (from the ANOVA)."""
    tabs = _ad_tables(lines)
    if not tabs:
        return None
    _h, rows, end = tabs[-1]
    groups = {r[0].strip().lower(): (r[1], _num_cell(r[2]), r[2]) for r in rows if r[0].lower() != "total" and _num_cell(r[2]) is not None}
    if len(groups) < 2:
        return None
    N, k = sum(g[0] for g in groups.values()), len(groups)
    msw = msw_t = None
    cols = None
    for i in range(end, min(len(lines), end + 14)):
        c2 = _cells_keep(lines[i])
        heads = {}
        for kk, c in enumerate(c2):
            for key, rx in _AN_COL.items():
                if rx.match((c or "").strip()) and key not in heads:
                    heads[key] = kk
        if "ss" in heads and "df" in heads and "ms" in heads:
            cols = heads
            continue
        if cols and re.match(r"^within(?:\s+groups?)?$", _row_label(c2)[1].lower()):
            msw_t = (c2[cols["ms"]] if cols["ms"] < len(c2) else "").strip()
            msw = _num_cell(msw_t)
            break
    return {"groups": groups, "k": k, "N": N, "msw": msw, "msw_t": msw_t, "end": end}


def spss_posthoc_check(text, counted=None):
    out, lines, n_judged = [], str(text or "").split("\n"), 0
    h_of = lambda t: 0.5 * 10 ** -_dec(str(t).strip().lstrip("-−").rstrip("*"))
    info = None
    for i, raw in enumerate(lines):
        cells = _cells_keep(raw)
        # ---- Test of Homogeneity of Variances
        if any(_LV_HEAD.match((c or "").strip()) for c in cells):
            info = _ph_groups(lines[:i])
            if not info:
                continue
            low = [(c or "").strip().lower() for c in cells]
            kl, k1, k2, ks = low.index("levene statistic"), (low.index("df1") if "df1" in low else None), (low.index("df2") if "df2" in low else None), next((k for k, c in enumerate(low) if c.startswith("sig")), None)
            for j in range(i + 1, min(len(lines), i + 6)):
                c2 = _cells_keep(lines[j])
                g = lambda k: (c2[k] if k is not None and k < len(c2) else "").strip()
                if _num_cell(g(kl)) is None:
                    continue
                note = None
                if k1 is not None and _num_cell(g(k1)) is not None and int(_num_cell(g(k1))) != info["k"] - 1:
                    note = f"Levene's test: {info['k']} groups give df1 = {info['k'] - 1}, not {g(k1)}."
                elif k2 is not None and _num_cell(g(k2)) is not None and int(_num_cell(g(k2))) != info["N"] - info["k"]:
                    note = f"Levene's test: {info['N']:,} people in {info['k']} groups give df2 = {info['N'] - info['k']}, not {g(k2)}."
                elif ks is not None and _p_cell(g(ks)):
                    f, f_t = _num_cell(g(kl)), g(kl)
                    rel, pv, dp = _p_cell(g(ks))
                    plo, phi = p_f(f + h_of(f_t), info["k"] - 1, info["N"] - info["k"]), p_f(max(f - h_of(f_t), 0), info["k"] - 1, info["N"] - info["k"])
                    if plo is not None and phi is not None:
                        n_judged += 1
                        if not _fits(rel, pv, dp, plo, phi):
                            note = (f"Levene's test: {f_t} on ({info['k'] - 1}, {info['N'] - info['k']}) gives "
                                    f"{_fmt_p(p_f(f, info['k'] - 1, info['N'] - info['k']), _p_lead(g(ks)))}, not Sig. {_p_said(g(ks))}.")
                if note:
                    out.append({"s": lines[j].strip()[:500], "note": note, "kind": "test", "decision": False})
                break
            continue
        # ---- Multiple Comparisons
        heads = {}
        for k, c in enumerate(cells):
            for key, rx in _PH_COL.items():
                if rx.match((c or "").strip()) and key not in heads:
                    heads[key] = k
        if "md" not in heads or "se" not in heads:
            continue
        info = _ph_groups(lines[:i])
        if not info:
            continue
        method = next((_PH_METHOD.match(lines[j].strip()).group(1).lower() for j in range(max(0, i - 4), i) if _PH_METHOD.match(lines[j].strip())), "")
        k, N, msw = info["k"], info["N"], info["msw"]
        comps = k * (k - 1) // 2
        seen, noted, i_lab = {}, set(), ""
        for j in range(i + 1, min(len(lines), i + 4 + k * k)):
            c2 = _cells_keep(lines[j])
            if not any((c or "").strip() for c in c2):
                break
            a, b = (c2[0] or "").strip() if c2 else "", (c2[1] or "").strip() if len(c2) > 1 else ""
            if a:
                i_lab = a
            g1, g2 = info["groups"].get(i_lab.lower()), info["groups"].get(b.lower())
            if not g1 or not g2:
                continue
            md_raw = (c2[heads["md"]] if heads["md"] < len(c2) else "").strip()
            md_t = md_raw.rstrip("*").strip()
            se_t = (c2[heads["se"]] if heads["se"] < len(c2) else "").strip()
            p_t_ = (c2[heads["p"]] if "p" in heads and heads["p"] < len(c2) else "").strip()
            md, se = _num_cell(md_t), _num_cell(se_t)
            if md is None or se is None or se <= 0:
                continue
            note = None
            (n1, m1, m1_t), (n2, m2, m2_t) = g1, g2
            if abs(md - (m1 - m2)) > h_of(md_t) + h_of(m1_t) + h_of(m2_t) + 1e-9:
                note = f"{i_lab} − {b}: the Mean Difference is {m1_t} − {m2_t} = {_like(m1 - m2, md_t)}, not {md_t}."
            elif msw is not None:
                lo_ = math.sqrt(max(msw - h_of(info["msw_t"]), 0) * (1 / n1 + 1 / n2))
                hi_ = math.sqrt((msw + h_of(info["msw_t"])) * (1 / n1 + 1 / n2))
                if not method.startswith("games") and (se + h_of(se_t) < lo_ - 1e-9 or se - h_of(se_t) > hi_ + 1e-9):
                    note = f"{i_lab} − {b}: the Std. Error is √(MS within × (1/{n1} + 1/{n2})), {_like(math.sqrt(msw * (1 / n1 + 1 / n2)), se_t)}, not {se_t}."
            pc = _p_cell(p_t_)
            if note is None and pc is not None and method in ("lsd", "bonferroni", "scheffe", "scheffé"):
                rel, pv, dp = pc
                t_lo, t_hi = max(abs(md) - h_of(md_t), 0) / (se + h_of(se_t)), (abs(md) + h_of(md_t)) / max(se - h_of(se_t), 1e-12)
                if method == "lsd":
                    pf = lambda t: p_t(t, N - k)
                elif method == "bonferroni":
                    pf = lambda t: min(1.0, comps * p_t(t, N - k))
                else:
                    pf = lambda t: p_f(t * t / (k - 1), k - 1, N - k)
                plo, phi = pf(t_hi), pf(t_lo)
                n_judged += 1
                if plo is not None and phi is not None and not _fits(rel, pv, dp, plo, phi):
                    name = {"lsd": "LSD", "bonferroni": "Bonferroni", "scheffe": "Scheffé", "scheffé": "Scheffé"}[method]
                    note = f"{i_lab} − {b}: {name}’s Sig. for a difference of {md_t} with a Std. Error of {se_t} is about {_fmt_p(pf(abs(md) / se), _p_lead(p_t_))}, not {p_t_}."
            if note is None and pc is not None and md_raw.endswith("*") != (pc[1] < 0.05 if pc[0] == "=" else pc[0] in ("<", "≤")) \
                    and not (pc[0] == "=" and 0.045 <= pc[1] <= 0.055):
                note = (f"{i_lab} − {b}: the star on {md_t} marks a difference significant at .05, but its Sig. is {p_t_}." if md_raw.endswith("*") else
                        f"{i_lab} − {b}: a Sig. of {p_t_} is below .05, but {md_t} carries no star (SPSS stars it).")
            key = frozenset((i_lab.lower(), b.lower()))
            if note is None and key in seen and key not in noted:          # (a pair already said is not said again turned round)
                md0, se0, p0 = seen[key]
                if abs(md + md0) > h_of(md_t) * 2 + 1e-9 or se_t != se0 or (p_t_ and p0 and p_t_ != p0):
                    note = f"{i_lab} − {b}: the same pair turned round holds {md0:g} with Std. Error {se0} and Sig. {p0}; here {md_t}, {se_t}, {p_t_}."
            seen.setdefault(key, (md, se_t, p_t_))
            if note:
                noted.add(key)
                out.append({"s": lines[j].strip()[:500], "note": note, "kind": "test", "decision": False})
    if counted is not None:
        counted[0] += n_judged
    return out



# 6.6, round 144 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a typed table of odds ratios (or risk, prevalence,
# hazard, rate ratios) with their 95% intervals, as public-health results chapters set them out: "Female | 2.31 (1.20–4.45) | 0.012
# | 1.98 (1.01–3.88) | 0.046" (the crude and the adjusted ratio), or "AOR | 95% CI | p" in three columns. A ratio lies inside its own
# interval, whose lower end comes first; an interval that leaves out 1 goes with p < .05, one that holds 1 with p ≥ .05. As the
# sentence check (_ci_notes) does, only where it plainly does (an end a unit or more from 1, a p away from .05), and only where the
# interval's own p, worked out from its ends as Wald's is (log scale), is itself well away from .05 (below .03 or above .08): an
# interval of another kind (profile likelihood, exact) and a test of another kind can then still disagree a little near the line.
# A reference row ("1.00", "Ref") is passed by. RR, PR and HR are also respiratory rate, pulse rate and heart rate: those columns
# are read only where their heading (or the next one) names the interval.
_OR_ODDS = re.compile(r"(?<![A-Za-z])(?:[ac]?OR|COR|AOR)(?![A-Za-z])")
_OR_ODDS_I = re.compile(r"\b(?:(?:crude|adjusted|unadjusted)\s+)?odds\s+ratios?\b", re.I)
_OR_OTHER = re.compile(r"(?<![A-Za-z])(?:[ac]?RR|ARR|CRR|[ac]?PR|APR|CPR|[ac]?HR|AHR|CHR|[ac]?IRR|AIRR)(?![A-Za-z])")
_OR_OTHER_I = re.compile(r"\b(?:(?:crude|adjusted|unadjusted)\s+)?(?:risk|prevalence|hazard|(?:incidence\s+)?rate)\s+ratios?\b", re.I)
_OR_CIHEAD = re.compile(r"\bC\.?\s?I\.?\b|95\s*%|\binterval\b", re.I)
_OR_PHEAD = re.compile(r"^\s*(?:p|p[\s-]*values?|sig\.?(?:\s*\(2-tailed\))?|significance)\s*[^\w\s]{0,2}\s*$", re.I)
_OR_CELL = re.compile(r"^(?P<e>\d+(?:\.\d+)?)\s*\*{0,3}\s*[(\[]\s*(?P<l>\d+(?:\.\d+)?)\s*(?:[-–—]|,|\bto\b)\s*(?P<u>\d+(?:\.\d+)?)\s*[)\]]\s*\*{0,3}$")
_CI_CELL = re.compile(r"^[(\[]?\s*(?P<l>\d+(?:\.\d+)?)\s*(?:[-–—]|,|\bto\b)\s*(?P<u>\d+(?:\.\d+)?)\s*[)\]]?$")
_OR_TITLE = re.compile(r"^\s*table\s+\d", re.I)
_OR_REF = re.compile(r"^(?:ref\.?|reference|referent|baseline|[-–—])$", re.I)


def _or_name(head):
    """The ratio a heading names, as the note says it: "AOR", "COR", "odds ratio"…"""
    m = _OR_ODDS.search(head) or _OR_OTHER.search(head) or _OR_ODDS_I.search(head) or _OR_OTHER_I.search(head)
    return m.group(0).strip() if m else "ratio"


def _or_cols(cells):
    """The columns of a heading row that hold a ratio (with the column of its p, if any): {col: (name, p_col or None)}."""
    heads = [(c or "").strip() for c in cells]
    found = []
    for k, h in enumerate(heads):
        if not h or len(h) > 60 or _OR_CELL.match(h):
            continue                                  # (a heading is short; a sentence is not one)
        odds = bool(_OR_ODDS.search(h) or _OR_ODDS_I.search(h))
        other = bool(_OR_OTHER_I.search(h)) or (bool(_OR_OTHER.search(h)) and (bool(_OR_CIHEAD.search(h)) or
                                               (k + 1 < len(heads) and bool(_OR_CIHEAD.search(heads[k + 1])) and not _OR_ODDS.search(heads[k + 1]))))
        if odds or other:
            found.append(k)
    out = {}
    for i, k in enumerate(found):
        stop = found[i + 1] if i + 1 < len(found) else len(heads)
        pc = next((j for j in range(k + 1, stop) if _OR_PHEAD.match(heads[j])), None)
        out[k] = (_or_name(heads[k]), pc)
    return out


def _or_wald_p(lo, hi, hl, hu):
    """The p an interval's ends give, as Wald's (on the log scale), over their rounding: (smallest, largest)."""
    ps = []
    for a in (lo - hl, lo + hl):
        for b in (hi - hu, hi + hu):
            if a <= 0 or b <= 0 or b <= a:
                continue
            la, lb = math.log(a), math.log(b)
            z = 1.959964 * (la + lb) / (lb - la)
            ps.append(p_z(abs(z)))
    return (min(ps), max(ps)) if ps else None


def or_table_check(text, counted=None):
    out, cols, n_judged = [], None, 0
    for raw in str(text or "").split("\n"):
        cells = _cells_keep(raw)
        filled = [c for c in cells if (c or "").strip()]
        if not filled or _OR_TITLE.match(raw):
            cols = None
            continue
        if all(re.fullmatch(r":?-{2,}:?", c.strip()) for c in filled):
            continue                                  # (a Markdown table's line under its heading)
        if len(filled) >= 2 and not any(_OR_CELL.match((c or "").strip()) or _CI_CELL.match((c or "").strip()) for c in cells):
            found = _or_cols(cells)
            if found:
                cols = found
                continue
            if len(filled) >= 2 and not any(re.search(r"\d", c) or _OR_REF.match(c.strip()) for c in filled):
                cols = None                           # (another table's heading; a reference row "Male | Ref" is not one)
                continue
        if not cols:
            continue
        lab = _row_label(cells)[1] or (cells[0] or "").strip()
        for k, (name, pcol) in cols.items():
            if k >= len(cells):
                continue
            c = (cells[k] or "").strip()
            m = _OR_CELL.match(c)
            if m:
                e_t, l_t, u_t = m.group("e"), m.group("l"), m.group("u")
            else:
                m2 = _CI_CELL.match((cells[k + 1] if k + 1 < len(cells) else "").strip())
                if not (m2 and _num_cell(c) is not None and re.fullmatch(r"\d+(?:\.\d+)?\*{0,3}", c)):
                    continue
                e_t, l_t, u_t = c.rstrip("*"), m2.group("l"), m2.group("u")
            e, lo, hi = float(e_t), float(l_t), float(u_t)
            he, hl, hu = (0.5 * 10 ** -_dec(x) for x in (e_t, l_t, u_t))
            who = f"{lab} ({name})" if lab else name
            shown = f"{l_t}–{u_t}"
            if lo > hi:
                out.append({"s": raw.strip()[:500], "note": f"{who}: its 95% CI’s lower end ({l_t}) is above its upper end ({u_t}).",
                            "kind": "ci", "decision": False})
                continue
            tol = max(he, hl, hu)
            if e < lo - tol or e > hi + tol:
                out.append({"s": raw.strip()[:500], "note": f"{who}: the ratio ({e_t}) lies outside its own 95% CI ({shown}).",
                            "kind": "ci", "decision": False})
                continue
            if pcol is None or pcol >= len(cells):
                continue
            p_txt = (cells[pcol] or "").strip()
            pc = _p_cell(p_txt)
            if pc is None:
                continue
            rel, pv, _dp = pc
            u = 10 ** -max(_dec(l_t), _dec(u_t))
            holds = lo < 1.0 - u and hi > 1.0 + u
            leaves = lo > 1.0 + u or hi < 1.0 - u
            wp = _or_wald_p(lo, hi, hl, hu)
            if wp is None or not (holds or leaves):
                continue
            if leaves and wp[1] >= 0.03 or holds and wp[0] <= 0.08:
                continue                              # (near the line: not judged)
            n_judged += 1
            sig = (rel in ("<", "≤") and pv <= 0.05) or (rel == "=" and pv < 0.045)
            nonsig = (rel in (">", "≥") and pv >= 0.05) or (rel == "=" and pv > 0.055)
            if holds and sig:
                out.append({"s": raw.strip()[:500], "note": f"{who}: its 95% CI ({shown}) holds 1, so its p can’t be below .05 (the table "
                            f"gives {_p_said(p_txt, bare=True)}): one of the two is wrong.", "kind": "ci", "decision": True})
            elif leaves and nonsig:
                out.append({"s": raw.strip()[:500], "note": f"{who}: its 95% CI ({shown}) leaves out 1, so its p should be below .05 (the "
                            f"table gives {_p_said(p_txt, bare=True)}): one of the two is wrong.", "kind": "ci", "decision": True})
    if counted is not None:
        counted[0] += n_judged
    return out



# 6.6, round 145 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): SPSS's binary logistic regression output, the tables
# rounds 120 and 143 left: Variables not in the Equation (each Score's Sig. from χ² on its df), Omnibus Tests of Model Coefficients
# (Sig. from its Chi-square and df), Hosmer and Lemeshow Test (the same), Classification Table (each Percentage Correct from its own
# counts, and the Overall Percentage), and Model Summary read with them: Cox & Snell R Square = 1 − e^(−χ²/N) and Nagelkerke's =
# Cox & Snell's ÷ (1 − e^(−(−2LL + χ²)/N)), with the model chi-square of the Omnibus table and N from the Classification Table (or
# the Case Processing Summary); the model chi-square and −2 Log likelihood add up to the constant-only model's −2 Log likelihood,
# −2(n₀ ln(n₀/N) + n₁ ln(n₁/N)). Nagelkerke's is never below Cox & Snell's. All over the rounding SPSS shows.
_LG_TITLE = re.compile(r"^(?P<t>omnibus\s+tests?\s+of\s+model\s+coefficients|model\s+summary|hosmer\s+(?:and|&)\s+lemeshow\s+test|"
                       r"classification\s+table|variables\s+not\s+in\s+the\s+equation|case\s+processing\s+summary|"
                       r"variables\s+in\s+the\s+equation|contingency\s+table\s+for\s+hosmer\s+(?:and|&)\s+lemeshow\s+test|"
                       r"iteration\s+history|dependent\s+variable\s+encoding|categorical\s+variables\s+codings|"
                       r"block\s+\d+\s*:[^\t|]*)(?:[a-e](?:,[a-e])*)?$", re.I)
_LG_NUM = re.compile(r"^-?\d*\.?\d+$")
_LG_INT = re.compile(r"^\d+$")


def _lg_nums(cells):
    """A row's numbers in order, as (text, value), footnote letters off ("439.839a")."""
    out = []
    for c in cells:
        t = re.sub(r"(?<=\d)[a-e]$", "", (c or "").strip()).replace("−", "-")
        if _LG_NUM.match(t):
            out.append((t, float(t)))
    return out


def _lg_p(label, stat_t, df, p_t, what):
    """A note where a χ² on df can't give the p written beside it, or None (and whether it was judged)."""
    pc = _p_cell(p_t)
    if pc is None or df < 1:
        return None, False
    rel, pv, dp = pc
    v = float(stat_t)
    h = 0.5 * 10 ** -_dec(stat_t)
    plo, phi = p_chi2(v + h, df), p_chi2(max(v - h, 0), df)
    if plo is None or phi is None:
        return None, False
    if _fits(rel, pv, dp, plo, phi):
        return None, True
    real = p_chi2(v, df)
    return ({"note": f"{what}{label}: {stat_t} on {df} df gives {_fmt_p(real, _p_lead(p_t))}, not {_p_said(p_t, bare=True)}.",
             "kind": "test", "decision": (real < 0.05) != (pv < 0.05 if rel == "=" else rel in ("<", "≤"))}, True)


def spss_logit_check(text, counted=None):
    out, mode, step, n_judged, sum_ok = [], None, None, 0, False
    n_cases, chi_model, summary, tables, cls = None, {}, [], [], None

    def end_cls():
        nonlocal cls
        if cls and len(cls["rows"]) == 2 and cls.get("overall") is not None:
            tables.append(cls)
        cls = None
    for raw in str(text or "").split("\n"):
        cells = _cells_keep(raw)
        filled = [(c or "").strip() for c in cells if (c or "").strip()]
        if not filled:
            continue
        one = " ".join(filled)
        tm = _LG_TITLE.match(one) if len(filled) == 1 else None
        if tm:
            end_cls()
            t = tm.group("t").lower()
            mode = ("omni" if t.startswith("omnibus") else "summary" if t.startswith("model summary") else
                    "hl" if t.startswith("hosmer") else "class" if t.startswith("classification") else
                    "score" if t.startswith("variables not") else "cases" if t.startswith("case processing") else None)
            step, sum_ok = None, False
            if mode == "class":
                cls = {"rows": [], "overall": None, "step": None}
            continue
        if len(filled) == 1 and (re.match(r"^[a-e]\.\s", one) or len(one) > 60):
            if not re.match(r"^[a-e]\.\s", one):
                end_cls()
                mode = None                          # (prose: the output has ended)
            continue
        low = [c.lower() for c in filled]
        sm = next((re.match(r"^step\s*(\d{1,2})\s*[a-e]?$", c, re.I) for c in filled if re.match(r"^step\s*\d{1,2}\s*[a-e]?$", c, re.I)), None)
        if sm:
            step = int(sm.group(1))
        nums = _lg_nums(cells)
        if mode == "cases":
            if any(re.match(r"^included\s+in\s+analysis$", c, re.I) for c in filled) and nums and _LG_INT.match(nums[0][0]):
                n_cases = int(nums[0][1])
        elif mode == "summary":
            if any("log likelihood" in c for c in low):
                sum_ok = any("nagelkerke" in c for c in low) and any("cox" in c for c in low)
                continue
            if sum_ok and len(nums) >= 4 and _LG_INT.match(nums[0][0]) and len(filled) == len(nums):
                summary.append({"step": int(nums[0][1]), "ll": nums[1], "cs": nums[2], "nk": nums[3], "s": raw.strip()[:500]})
        elif mode in ("omni", "hl", "score"):
            if len(nums) < 3 or not _LG_INT.match(nums[-2][0]):
                continue
            stat_t, df, p_t = nums[-3][0], int(nums[-2][1]), nums[-1][0]
            if mode == "omni":
                lab = next((c for c in filled if c.lower() in ("step", "block", "model")), None)
                if lab is None:
                    continue
                if lab.lower() == "model":
                    chi_model[step if step is not None else 1] = (stat_t, float(stat_t))
                what, label = "Omnibus Tests of Model Coefficients", f" ({lab})"
            elif mode == "hl":
                what, label = "Hosmer and Lemeshow Test", ""
                if len(nums) >= 4 and _LG_INT.match(nums[0][0]):
                    label = f" (step {nums[0][0]})"
            else:
                texts = [c for c in filled if not _LG_NUM.match(re.sub(r"(?<=\d)[a-e]$", "", c))]
                lab = texts[-1] if texts else ""
                what, label = "Variables not in the Equation, ", lab
            note, judged = _lg_p(label, stat_t, df, p_t, what)
            n_judged += 1 if judged else 0
            if note:
                out.append(dict(note, s=raw.strip()[:500]))
        elif mode == "class" and cls is not None:
            if any(re.match(r"^overall\s+percentage$", c, re.I) for c in filled):
                if len(nums) == 1:
                    cls["overall"] = nums[0]
                    cls["step"] = step
                    cls["s_over"] = raw.strip()[:500]
                end_cls()
                cls = {"rows": [], "overall": None, "step": None}
                continue
            if len(nums) == 3 and _LG_INT.match(nums[0][0]) and _LG_INT.match(nums[1][0]):
                texts = [c for c in filled if not _LG_NUM.match(re.sub(r"(?<=\d)[a-e]$", "", c))]
                cls["rows"].append({"a": int(nums[0][1]), "b": int(nums[1][1]), "pct": nums[2], "lab": texts[-1] if texts else "",
                                    "s": raw.strip()[:500]})
    end_cls()
    # the Classification Tables, row by row
    for tb in tables:
        r1, r2 = tb["rows"]
        where = f"Classification Table (step {tb['step']})" if tb["step"] is not None else "Classification Table"
        for r, right in ((r1, r1["a"]), (r2, r2["b"])):
            tot = r["a"] + r["b"]
            if tot <= 0:
                continue
            want = 100.0 * right / tot
            t, v = r["pct"]
            if abs(want - v) > 0.5 * 10 ** -_dec(t) + 1e-9:
                out.append({"s": r["s"], "note": f"{where}, {r['lab']}: {right} of its {tot} are predicted right, so its Percentage Correct is "
                            f"{_like(want, t)}, not {t}.", "kind": "table", "decision": False})
        n = r1["a"] + r1["b"] + r2["a"] + r2["b"]
        t, v = tb["overall"]
        if n > 0:
            want = 100.0 * (r1["a"] + r2["b"]) / n
            if abs(want - v) > 0.5 * 10 ** -_dec(t) + 1e-9:
                out.append({"s": tb.get("s_over", ""), "note": f"{where}: {r1['a'] + r2['b']} of the {n} are predicted right, so its Overall "
                            f"Percentage is {_like(want, t)}, not {t}.", "kind": "table", "decision": False})
    # the Model Summary, read with the Omnibus test and the Classification Table
    totals = {}
    for tb in tables:
        r1, r2 = tb["rows"]
        totals[tb["step"]] = (r1["a"] + r1["b"], r2["a"] + r2["b"])
    for row in summary:
        (ll_t, ll), (cs_t, cs), (nk_t, nk) = row["ll"], row["cs"], row["nk"]
        hc, hn = 0.5 * 10 ** -_dec(cs_t), 0.5 * 10 ** -_dec(nk_t)
        if nk + hn < cs - hc - 1e-9:
            out.append({"s": row["s"], "note": f"Model Summary, step {row['step']}: Nagelkerke R Square ({nk_t}) is never below Cox & Snell’s "
                        f"({cs_t}): it is Cox & Snell’s scaled up so that it can reach 1.", "kind": "table", "decision": False})
            continue
        if cs - hc >= 1.0 or nk - hn > 1.0:
            out.append({"s": row["s"], "note": f"Model Summary, step {row['step']}: Cox & Snell R Square stays below 1 and Nagelkerke’s is at most 1 (the table gives {cs_t} and {nk_t}).",
                        "kind": "table", "decision": False})
            continue
        split = totals.get(row["step"]) or next((v for k, v in totals.items() if k not in (0,)), None) or totals.get(0)
        n = sum(split) if split else None
        if n and n_cases and n != n_cases:
            continue                                  # (tables from two runs: not judged)
        n = n or n_cases
        chi = chi_model.get(row["step"])
        if not n or not chi:
            continue
        chi_t, x = chi
        hx, hl_ = 0.5 * 10 ** -_dec(chi_t), 0.5 * 10 ** -_dec(ll_t)
        if split and min(split) > 0:                  # (first: the two add up to the constant-only model's −2LL)
            n0, n1 = split
            ll0 = -2 * (n0 * math.log(n0 / n) + n1 * math.log(n1 / n))
            if abs((ll + x) - ll0) > hx + hl_ + 1e-6:
                out.append({"s": row["s"], "note": f"Model Summary, step {row['step']}: its −2 Log likelihood ({ll_t}) and the model chi-square "
                            f"({chi_t}) add up to {ll + x:.3f}, but with {n0} and {n1} in the two outcome groups the constant-only model’s −2 Log "
                            f"likelihood is {ll0:.3f}: one of the two is wrong.", "kind": "table", "decision": False})
                continue
        cs_lo, cs_hi = 1 - math.exp(-(x - hx) / n), 1 - math.exp(-(x + hx) / n)
        if cs_hi < cs - hc - 1e-9 or cs_lo > cs + hc + 1e-9:
            want = 1 - math.exp(-x / n)
            out.append({"s": row["s"], "note": f"Model Summary, step {row['step']}: with the model chi-square ({chi_t}) and N = {n}, Cox & Snell "
                        f"R Square is 1 − e^(−{chi_t}/{n}) ≈ {_like(want, cs_t)}, not {cs_t}.", "kind": "table", "decision": False})
            continue
        nks = []
        for xx in (x - hx, x + hx):
            for L in (ll - hl_, ll + hl_):
                c_ = 1 - math.exp(-xx / n)
                m_ = 1 - math.exp(-(L + xx) / n)
                if m_ > 0:
                    nks.append(c_ / m_)
        if nks and (max(nks) < nk - hn - 1e-9 or min(nks) > nk + hn + 1e-9):
            want = (1 - math.exp(-x / n)) / (1 - math.exp(-(ll + x) / n))
            out.append({"s": row["s"], "note": f"Model Summary, step {row['step']}: with its −2 Log likelihood ({ll_t}), the model chi-square "
                        f"({chi_t}) and N = {n}, Nagelkerke R Square is about {_like(want, nk_t)}, not {nk_t}.", "kind": "table", "decision": False})
    if counted is not None:
        counted[0] += n_judged
    return out



# 6.6, round 146 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): two SPSS outputs that hold one variable against a
# value. The one-sample t-test (One-Sample Statistics, One-Sample Test with its "Test Value = 3"): each Std. Error Mean = SD / √N,
# df = N − 1, Mean Difference = Mean − Test Value, t = Mean Difference / Std. Error Mean, Sig. (2-tailed) from t on its df, and the
# 95% interval of the difference = Mean Difference ∓ t(.975, df) × Std. Error Mean. The chi-square goodness of fit (each variable's
# Observed N / Expected N / Residual table, then Test Statistics with a column for each): Residual = Observed − Expected, the
# Expected add up to the Total, Chi-Square = Σ Residual² / Expected, df = categories − 1, Asymp. Sig. from them, and its footnote's
# smallest expected frequency. All over the rounding SPSS shows.
_OS_TITLE = re.compile(r"^one-sample\s+(?P<k>statistics|test)\s*$", re.I)
_OS_TV = re.compile(r"^test\s+value\s*=\s*(?P<v>-?\d*\.?\d+)\s*$", re.I)
_OS_COL = {"n": re.compile(r"^N$"), "m": re.compile(r"^mean$", re.I), "sd": re.compile(r"^std\.?\s*deviation$", re.I),
           "sem": re.compile(r"^std\.?\s*error\s+mean$", re.I), "t": re.compile(r"^t$"), "df": re.compile(r"^df$", re.I),
           "p": re.compile(r"^sig\.?\s*\(2-tailed\)$", re.I), "md": re.compile(r"^mean\s+difference$", re.I),
           "lo": re.compile(r"^lower$", re.I), "hi": re.compile(r"^upper$", re.I)}


def _os_heads(cells):
    heads = {}
    for k, c in enumerate(cells):
        for key, rx in _OS_COL.items():
            if rx.match((c or "").strip()) and key not in heads:
                heads[key] = k
    return heads


def spss_onesample_check(text, counted=None):
    out, n_judged = [], 0
    kind, cols, tv, stats_, tests = None, {}, None, {}, []
    for raw in str(text or "").split("\n"):
        cells = _cells_keep(raw)
        filled = [(c or "").strip() for c in cells if (c or "").strip()]
        if not filled:
            continue
        if len(filled) == 1:
            tm, vm = _OS_TITLE.match(filled[0]), _OS_TV.match(filled[0])
            if tm:
                kind, cols = ("stats" if tm.group("k").lower() == "statistics" else "test"), {}
                if kind == "test":
                    tv = None
                continue
            if vm and kind == "test":
                tv = vm.group("v")
                continue
            if not re.match(r"^[a-e]\.\s", filled[0]):
                kind = None
            continue
        if kind is None:
            continue
        heads = _os_heads(cells)
        if kind == "stats" and {"n", "m", "sd", "sem"} <= set(heads):
            cols = heads
            continue
        if kind == "test" and ({"t", "df", "p"} <= set(heads) or {"lo", "hi"} <= set(heads)):
            cols.update(heads)
            continue
        name = (cells[0] or "").strip()
        if not name or not re.search(r"[A-Za-z]", name) or not cols:
            continue
        get = lambda key: (cells[cols[key]] or "").strip() if key in cols and cols[key] < len(cells) else ""
        if kind == "stats":
            n_t = get("n")
            if re.fullmatch(r"\d+", n_t) and all(_num_cell(get(k)) is not None for k in ("m", "sd", "sem")):
                stats_[name.lower()] = {"n": int(n_t), "m": get("m"), "sd": get("sd"), "sem": get("sem"), "s": raw.strip()[:500], "name": name}
        elif tv is not None and {"t", "df", "p"} <= set(cols):
            tests.append({"name": name, "t": get("t"), "df": get("df"), "p": get("p"), "md": get("md"), "lo": get("lo"), "hi": get("hi"),
                          "tv": tv, "s": raw.strip()[:500]})
    h_of = lambda t: 0.5 * 10 ** -_dec(str(t).lstrip("-−"))
    for st in stats_.values():                      # each Std. Error Mean
        n, sd, sem = st["n"], _num_cell(st["sd"]), _num_cell(st["sem"])
        if n >= 2 and sd is not None and sem is not None:
            lo_, hi_ = (sd - h_of(st["sd"])) / math.sqrt(n), (sd + h_of(st["sd"])) / math.sqrt(n)
            if sem + h_of(st["sem"]) < lo_ - 1e-12 or sem - h_of(st["sem"]) > hi_ + 1e-12:
                out.append({"s": st["s"], "kind": "table", "decision": False,
                            "note": f"{st['name']}: its Std. Error Mean is its SD over √N, {st['sd']} / √{n} = {_like(sd / math.sqrt(n), st['sem'])}, not {st['sem']}."})
                st["bad"] = True
    for te in tests:
        st = stats_.get(te["name"].lower())
        t, df_t, md, lo_v, hi_v = _num_cell(te["t"]), te["df"], _num_cell(te["md"]), _num_cell(te["lo"]), _num_cell(te["hi"])
        tvv = float(te["tv"])
        note = None
        if st and re.fullmatch(r"\d+", df_t) and int(df_t) != st["n"] - 1:
            note = f"{te['name']}: N = {st['n']} gives {st['n'] - 1} df, not {df_t}."
        m = _num_cell(st["m"]) if st else None
        if note is None and st and md is not None and m is not None:
            if abs(md - (m - tvv)) > h_of(te["md"]) + h_of(st["m"]) + 1e-9:
                note = f"{te['name']}: the Mean Difference is its Mean less the Test Value, {st['m']} − {te['tv']} = {_like(m - tvv, te['md'])}, not {te['md']}."
        sem_t = st["sem"] if st and not st.get("bad") else None
        sem = _num_cell(sem_t) if sem_t else None
        if note is None and t is not None and md is not None and sem:
            lo_t = (abs(md) - h_of(te["md"])) / (sem + h_of(sem_t))
            hi_t = (abs(md) + h_of(te["md"])) / max(sem - h_of(sem_t), 1e-12)
            sd_ok = True
            if st:                                    # (the SD and N bound the Std. Error Mean too)
                sd = _num_cell(st["sd"])
                s_lo, s_hi = max(sem - h_of(sem_t), (sd - h_of(st["sd"])) / math.sqrt(st["n"])), min(sem + h_of(sem_t), (sd + h_of(st["sd"])) / math.sqrt(st["n"]))
                if s_lo <= s_hi and s_lo > 0:
                    lo_t, hi_t = max(lo_t, (abs(md) - h_of(te["md"])) / s_hi), min(hi_t, (abs(md) + h_of(te["md"])) / s_lo)
                sd_ok = s_lo <= s_hi
            if sd_ok and (abs(t) + h_of(te["t"]) < lo_t - 1e-9 or abs(t) - h_of(te["t"]) > hi_t + 1e-9
                          or (t * md < 0 and abs(md) > h_of(te["md"]) and abs(t) > h_of(te["t"]))):
                note = f"{te['name']}: t is the Mean Difference over its Std. Error Mean, {te['md']} / {sem_t} ≈ {md / sem:.3f}, not {te['t']}."
        pc = _p_cell(te["p"])
        if note is None and t is not None and re.fullmatch(r"\d+", df_t) and int(df_t) > 0 and pc is not None:
            rel, pv, dp = pc
            d = int(df_t)
            plo, phi = p_t(abs(t) + h_of(te["t"]), d), p_t(max(abs(t) - h_of(te["t"]), 0.0), d)
            n_judged += 1
            if not _fits(rel, pv, dp, plo, phi):
                real = p_t(abs(t), d)
                out.append({"s": te["s"], "kind": "test", "decision": (real < 0.05) != (pv < 0.05 if rel == "=" else rel in ("<", "≤")),
                            "note": f"{te['name']}: t = {te['t']} on {d} df gives {_fmt_p(real, _p_lead(te['p']))}, not Sig. (2-tailed) {_p_said(te['p'])}."})
                continue
        if note is None and None not in (md, lo_v, hi_v) and sem and re.fullmatch(r"\d+", df_t) and int(df_t) > 0:
            tc = _t_crit(int(df_t))
            slack = h_of(te["md"]) + tc * h_of(sem_t)
            for val, tx, want, end_, sign in ((lo_v, te["lo"], md - tc * sem, "lower", "−"), (hi_v, te["hi"], md + tc * sem, "upper", "+")):
                if abs(val - want) > slack + h_of(tx) + 1e-9:
                    note = (f"{te['name']}: the 95% interval’s {end_} end is about {_like(want, tx)} (the Mean Difference {sign} t(.975, {df_t}) × its Std. "
                            f"Error Mean), not {tx}.")
                    break
        if note:
            out.append({"s": te["s"], "note": note, "kind": "test", "decision": False})
    if counted is not None:
        counted[0] += n_judged
    return out


_GF_HEAD = re.compile(r"^observed\s+n$", re.I)
_GF_FOOT = re.compile(r"^(?P<l>[a-e])\.\s+(?P<c>\d+)\s+cells?\s+\((?P<pc>\d+(?:\.\d+)?)%\)\s+have\s+expected\s+frequencies\s+less\s+than\s+5\.\s+"
                      r"the\s+minimum\s+expected\s+cell\s+frequency\s+is\s+(?P<m>\d+(?:\.\d+)?)\.?$", re.I)


def spss_gof_check(text, counted=None):
    out, n_judged = [], 0
    lines = str(text or "").split("\n")
    tables, cur, prev_title = [], None, None
    ts, foot = None, {}
    for raw in lines:
        cells = _cells_keep(raw)
        filled = [(c or "").strip() for c in cells if (c or "").strip()]
        if not filled:
            cur = None
            continue
        low = [c.lower() for c in filled]
        if any(_GF_HEAD.match(c) for c in filled) and "expected n" in low and "residual" in low:
            ix = {k: next(j for j, c in enumerate(cells) if (c or "").strip().lower() == k) for k in ("observed n", "expected n", "residual")}
            cur = {"title": prev_title or "", "ix": ix, "rows": [], "total": None}
            tables.append(cur)
            continue
        if ts is not None and ts["cols"] is None and not (cells[0] or "").strip():
            ts["cols"] = [(c or "").strip() for c in cells[1:]]
            continue                                  # (Test Statistics' heading: a variable's name over each column)
        if len(filled) == 1:
            fm = _GF_FOOT.match(filled[0])
            if fm:
                foot[fm.group("l").lower()] = fm
                continue
            if re.match(r"^test\s+statistics(?:[a-e](?:,[a-e])*)?$", filled[0], re.I):
                ts = {"cols": None, "rows": {}}
                cur = None
                continue
            prev_title = filled[0]
            cur = None
            continue
        if cur is not None:
            get = lambda key: (cells[cur["ix"][key]] or "").strip() if cur["ix"][key] < len(cells) else ""
            lab = (cells[0] or "").strip()
            o_t, e_t, r_t = get("observed n"), get("expected n"), get("residual")
            if lab.lower() == "total":
                if re.fullmatch(r"\d+", o_t):
                    cur["total"] = int(o_t)
                cur = None
            elif re.fullmatch(r"\d+", o_t) and _num_cell(e_t) is not None and _num_cell(r_t) is not None:
                cur["rows"].append((lab, int(o_t), e_t, r_t, raw.strip()[:500]))
            continue
        if ts is not None:
            lab = (cells[0] or "").strip().lower()
            if lab in ("chi-square", "df", "asymp. sig.", "asymptotic significance"):
                ts["rows"]["sig" if "sig" in lab else lab] = ([(c or "").strip() for c in cells[1:]], raw.strip()[:500])
    h_of = lambda t: 0.5 * 10 ** -_dec(str(t).lstrip("-−"))
    tables = [tb for tb in tables if len(tb["rows"]) >= 2]
    for tb in tables:
        tb["chi"], tb["bad"] = None, False
        for lab, o, e_t, r_t, s in tb["rows"]:
            e, r = _num_cell(e_t), _num_cell(r_t)
            if abs(r - (o - e)) > h_of(e_t) + h_of(r_t) + 1e-9:
                out.append({"s": s, "kind": "table", "decision": False,
                            "note": f"{tb['title'] or 'Goodness of fit'}, {lab}: its Residual is Observed − Expected, {o} − {e_t} = {_like(o - e, r_t)}, not {r_t}."})
                tb["bad"] = True
        n = tb["total"] if tb["total"] is not None else sum(r[1] for r in tb["rows"])
        es = [(_num_cell(r[2]), h_of(r[2])) for r in tb["rows"]]
        if abs(sum(e for e, _ in es) - n) > sum(h for _, h in es) + 1e-9:
            out.append({"s": tb["rows"][0][4], "kind": "table", "decision": False,
                        "note": f"{tb['title'] or 'Goodness of fit'}: its Expected N add up to {sum(e for e, _ in es):.1f}, but they share out its {n} cases."})
            tb["bad"] = True
        if not tb["bad"]:
            lo_, hi_ = 0.0, 0.0
            for (lab, o, e_t, r_t, s), (e, h) in zip(tb["rows"], es):
                cand = [(o - x) ** 2 / x for x in (e - h, e + h) if x > 0]
                if not cand:
                    break
                d_lo = 0.0 if e - h <= o <= e + h else min(cand)
                lo_, hi_ = lo_ + d_lo, hi_ + max(cand)
            else:
                tb["chi"] = (lo_, hi_, sum((o - e) ** 2 / e for (lab, o, e_t, r_t, s), (e, h) in zip(tb["rows"], es) if e > 0))
    if ts and ts["cols"] and "chi-square" in ts["rows"]:
        names = ts["cols"]
        order = []
        for j, name in enumerate(names):
            tb = next((x for x in tables if x["title"] and x["title"].lower() == name.lower()), None)
            if tb is None and len(names) == len(tables):
                tb = tables[j]
            order.append(tb)
        chis, s_chi = ts["rows"]["chi-square"]
        dfs = ts["rows"].get("df", ([], ""))[0]
        sigs, s_sig = ts["rows"].get("sig", ([], ""))
        for j, tb in enumerate(order):
            if j >= len(chis):
                break
            if tb is None:
                continue                              # (6.6, round 147: a Kruskal–Wallis or Friedman "Chi-Square" is round 147's)
            c_t = re.sub(r"(?<=\d)[a-e]$", "", chis[j])
            fl = chis[j][-1].lower() if chis[j] and chis[j][-1].isalpha() else None
            x = _num_cell(c_t)
            if x is None:
                continue
            name = names[j] or (tb["title"] if tb else "")
            if tb is not None and tb["chi"] is not None:
                lo_, hi_, want = tb["chi"]
                if x + h_of(c_t) < lo_ - 1e-9 or x - h_of(c_t) > hi_ + 1e-9:
                    out.append({"s": s_chi, "kind": "test", "decision": False,
                                "note": f"{name}: Chi-Square is Σ Residual² / Expected over its {len(tb['rows'])} categories, about {want:.3f}, not {c_t}."})
                    continue
            d_t = dfs[j] if j < len(dfs) else ""
            if tb is not None and re.fullmatch(r"\d+", d_t) and int(d_t) != len(tb["rows"]) - 1:
                out.append({"s": s_chi, "kind": "test", "decision": False,
                            "note": f"{name}: {len(tb['rows'])} categories give {len(tb['rows']) - 1} df, not {d_t}."})
                continue
            p_t_ = sigs[j] if j < len(sigs) else ""
            if re.fullmatch(r"\d+", d_t) and int(d_t) > 0:
                note, judged = _lg_p("", c_t, int(d_t), p_t_, name)
                n_judged += 1 if judged else 0
                if note:
                    out.append(dict(note, s=s_sig))
                    continue
            fm = foot.get(fl) if fl else None
            if fm and tb is not None:
                es = [(_num_cell(r[2]), r[2]) for r in tb["rows"]]
                mn = min(es, key=lambda z: z[0])
                if abs(float(fm.group("m")) - mn[0]) > h_of(fm.group("m")) + h_of(mn[1]) + 1e-9:
                    out.append({"s": fm.group(0)[:500], "kind": "table", "decision": False,
                                "note": f"{name}: its smallest Expected N is {mn[1]}, so the footnote’s minimum expected cell frequency is {mn[1]}, not {fm.group('m')}."})
                elif int(fm.group("c")) != sum(1 for e, _ in es if e < 5):
                    k5 = sum(1 for e, _ in es if e < 5)
                    out.append({"s": fm.group(0)[:500], "kind": "table", "decision": False,
                                "note": f"{name}: {k5} of its {len(es)} expected frequencies are below 5, not {fm.group('c')}."})
    if counted is not None:
        counted[0] += n_judged
    return out



# 6.6, round 147 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): SPSS's tests on ranks, each Ranks table read with its
# Test Statistics. Mann–Whitney: each Mean Rank = Sum of Ranks / N; the two Sums = N(N + 1)/2; U = a group's Sum less n(n + 1)/2;
# W = a group's Sum; |Z| at least what U gives without ties (ties only make it larger); Asymp. Sig. from Z. Kruskal–Wallis: the
# groups' N add up to the Total; Σ N × Mean Rank = N(N + 1)/2; df = groups − 1; H at least what the Mean Ranks give without ties;
# Asymp. Sig. from H on its df. Wilcoxon signed ranks: Negative + Positive + Ties = Total; each Mean Rank = Sum / N; the two Sums =
# m(m + 1)/2 for the m that aren't ties; |Z| at least what the Sums give without ties; Asymp. Sig. from Z. Friedman: the Mean Ranks
# add up to k(k + 1)/2; df = k − 1; Chi-Square at least what the Mean Ranks give without ties; its Sig. (Round 146's goodness of fit
# now judges only the columns of its own tables: a Kruskal–Wallis or Friedman "Chi-Square" is judged here, once.)
_NP_RANKS = re.compile(r"^ranks$", re.I)
_NP_TS = re.compile(r"^test\s+statistics(?:[a-e](?:,[a-e])*)?$", re.I)
_NP_ROW = {"u": re.compile(r"^mann-?\s?whitney\s+u$", re.I), "w": re.compile(r"^wilcoxon\s+w$", re.I), "z": re.compile(r"^z$", re.I),
           "p2": re.compile(r"^asymp(?:\.|totic)\s+sig(?:\.|nificance)\s*\(2-tailed\)$", re.I),
           "h": re.compile(r"^kruskal-?\s?wallis\s+h$", re.I), "chi": re.compile(r"^chi-?\s?square$", re.I), "df": re.compile(r"^df$", re.I),
           "p": re.compile(r"^asymp(?:\.|totic)\s+sig(?:\.|nificance)$", re.I), "n": re.compile(r"^N$")}


def _np_cell(c):
    """A cell's text without its footnote letter ("79a", "-4.090b")."""
    return re.sub(r"(?<=\d)[a-e]$", "", (c or "").strip()).replace("−", "-")


def _np_read(lines):
    blocks, ts_list, cur, cols, kind = [], [], None, None, None
    for raw in lines:
        cells = _cells_keep(raw)
        filled = [(c or "").strip() for c in cells if (c or "").strip()]
        if not filled:
            continue
        if len(filled) == 1 and _NP_RANKS.match(filled[0]):
            kind, cols, cur = "ranks", None, None
            continue
        if len(filled) == 1 and _NP_TS.match(filled[0]):
            kind, cur = "ts", {"cols": None, "rows": {}}
            ts_list.append(cur)
            continue
        if kind == "ts" and cur is not None and cur["cols"] is None and not (cells[0] or "").strip():
            cur["cols"] = [(c or "").strip() for c in cells[1:]]
            continue                                  # (Test Statistics' heading, even with one column)
        if kind == "ranks" and any((c or "").strip().lower() == "mean rank" for c in cells):
            low = [(c or "").strip().lower() for c in cells]
            cols = {"n": low.index("n") if "n" in low else None, "mr": low.index("mean rank"),
                    "sum": low.index("sum of ranks") if "sum of ranks" in low else None}
            continue                                  # (the Ranks heading; Friedman's has "Mean Rank" alone)
        if len(filled) == 1:
            if not re.match(r"^[a-e]\.\s", filled[0]):
                kind, cur = None, None                # (another table's title)
            continue
        if kind == "ranks":
            if cols is None:
                continue
            first = (cells[0] or "").strip()
            if cols["n"] is None:                     # (Friedman: a Mean Rank for each condition)
                if first and len(cells) > cols["mr"] and _num_cell(_np_cell(cells[cols["mr"]])) is not None:
                    if not blocks or blocks[-1]["type"] != "friedman" or blocks[-1].get("done"):
                        blocks.append({"type": "friedman", "rows": [], "name": "", "s": raw.strip()[:500]})
                    blocks[-1]["rows"].append((first, _np_cell(cells[cols["mr"]])))
                continue
            grp = (cells[1] or "").strip() if len(cells) > 1 else ""
            if first and grp:
                cur = {"type": "groups", "name": first, "rows": [], "total": None, "s": raw.strip()[:500]}
                blocks.append(cur)
            if cur is None or not grp:
                continue
            get = lambda key: _np_cell(cells[cols[key]]) if cols.get(key) is not None and cols[key] < len(cells) else ""
            n_t = get("n")
            if grp.lower() == "total":
                cur["total"] = int(n_t) if re.fullmatch(r"\d+", n_t) else None
                cur = None
                continue
            if re.fullmatch(r"\d+", n_t):
                cur["rows"].append({"g": grp, "n": int(n_t), "mr": get("mr"), "sum": get("sum"), "s": raw.strip()[:500]})
        elif kind == "ts" and cur is not None:
            lab = (cells[0] or "").strip()
            for key, rx in _NP_ROW.items():
                if rx.match(lab):
                    cur["rows"][key] = ([_np_cell(c) for c in cells[1:]], raw.strip()[:500], lab)
                    break
    for b in blocks:
        if b["type"] == "friedman":
            b["done"] = True
    return blocks, ts_list


def spss_np_check(text, counted=None):
    out, n_judged = [], 0
    blocks, ts_list = _np_read(str(text or "").split("\n"))
    if not blocks:
        return out
    h_of = lambda t: 0.5 * 10 ** -_dec(str(t).lstrip("-−"))
    cols = []                                         # (each Test Statistics column: its name and its rows)
    for ts in ts_list:
        names = ts["cols"] if ts["cols"] is not None else ([""] if ts["rows"] else [])     # (Friedman's: no heading)
        for j, name in enumerate(names):
            vals = {k: (v[0][j] if j < len(v[0]) else "", v[1], v[2]) for k, v in ts["rows"].items()}
            cols.append((name, vals))
    used = set()
    for bi, b in enumerate(blocks):
        name = b["name"] or (b["rows"][0][0] if b["type"] == "friedman" and b["rows"] else "")
        col = next((c for k, c in enumerate(cols) if k not in used and c[0].lower() == (b["name"] or "").lower() and b["name"]), None)
        if col is None and len(cols) == len(blocks):
            col = cols[bi]
        if col is not None:
            used.add(cols.index(col))
        vals = col[1] if col else {}
        val = lambda k: vals[k][0] if k in vals else ""
        sline = lambda k: vals[k][1] if k in vals else b["s"]
        note = None
        if b["type"] == "friedman":
            rows = b["rows"]
            k = len(rows)
            mrs = [(_num_cell(t), t) for _g, t in rows]
            if k < 2 or any(m is None for m, _ in mrs):
                continue
            tot, slack = sum(m for m, _ in mrs), sum(h_of(t) for _, t in mrs)
            if abs(tot - k * (k + 1) / 2) > slack + 1e-9:
                out.append({"s": b["s"], "kind": "table", "decision": False,
                            "note": f"Friedman test: its {k} Mean Ranks add up to {tot:.{_dec(mrs[0][1])}f}, but the Mean Ranks of {k} occasions always add "
                                    f"up to {k} × {k + 1} / 2 = {k * (k + 1) / 2:g}."})
                continue
            df_t = val("df")
            if re.fullmatch(r"\d+", df_t) and int(df_t) != k - 1:
                out.append({"s": sline("df"), "kind": "test", "decision": False, "note": f"Friedman test: {k} conditions give {k - 1} df, not {df_t}."})
                continue
            n_t, c_t = val("n"), val("chi")
            if re.fullmatch(r"\d+", n_t) and _num_cell(c_t) is not None:
                n = int(n_t)
                c = (k + 1) / 2
                lo = 12 * n / (k * (k + 1)) * sum(max(abs(m - c) - h_of(t), 0) ** 2 for m, t in mrs)
                if _num_cell(c_t) + h_of(c_t) < lo - 1e-6:
                    out.append({"s": sline("chi"), "kind": "test", "decision": False,
                                "note": f"Friedman test: its Mean Ranks with N = {n} give a Chi-Square of at least {lo:.3f} (ties only make it larger), not {c_t}."})
                    continue
                if re.fullmatch(r"\d+", df_t) and int(df_t) > 0:
                    note, judged = _lg_p("", c_t, int(df_t), val("p"), "Friedman test")
                    n_judged += 1 if judged else 0
                    if note:
                        out.append(dict(note, s=sline("p")))
            continue
        rows = b["rows"]
        if len(rows) < 2:
            continue
        wil = {r["g"].lower() for r in rows} >= {"negative ranks", "positive ranks"}
        sums = [r for r in rows if r["g"].lower() in ("negative ranks", "positive ranks")] if wil else rows
        if len(sums) == 2 and all(_num_cell(r["sum"]) is not None for r in sums):      # (first: the Sums of Ranks add up)
            m = sum(r["n"] for r in sums)
            got_, want = sum(_num_cell(r["sum"]) for r in sums), m * (m + 1) / 2
            if abs(got_ - want) > sum(h_of(r["sum"]) for r in sums) + 1e-9:
                which = f"the {m} ranks that aren’t ties" if wil else f"the ranks 1 to {m}"
                out.append({"s": b["s"], "kind": "table", "decision": False,
                            "note": f"{b['name']}: its two Sums of Ranks add up to {got_:g}, but {which} add up to {m} × {m + 1} / 2 = {want:g}."})
                continue
        # each Mean Rank against its Sum of Ranks
        for r in rows:
            mr, sm = _num_cell(r["mr"]), _num_cell(r["sum"])
            if mr is not None and sm is not None and r["n"] > 0 and abs(mr - sm / r["n"]) > h_of(r["mr"]) + h_of(r["sum"]) / r["n"] + 1e-9:
                note = (f"{b['name']}, {r['g']}: its Mean Rank is its Sum of Ranks over N, {r['sum']} / {r['n']} = {_like(sm / r['n'], r['mr'])}, "
                        f"not {r['mr']}.")
                out.append({"s": r["s"], "kind": "table", "decision": False, "note": note})
                break
        if note:
            continue
        if wil:
            neg = next(r for r in rows if r["g"].lower() == "negative ranks")
            pos = next(r for r in rows if r["g"].lower() == "positive ranks")
            ties = next((r for r in rows if r["g"].lower() == "ties"), None)
            if b["total"] is not None and ties is not None and neg["n"] + pos["n"] + ties["n"] != b["total"]:
                out.append({"s": b["s"], "kind": "table", "decision": False,
                            "note": f"{b['name']}: Negative Ranks, Positive Ranks and Ties add up to {neg['n'] + pos['n'] + ties['n']}, not its Total {b['total']}."})
                continue
            m = neg["n"] + pos["n"]
            sn, sp = _num_cell(neg["sum"]), _num_cell(pos["sum"])
            if m < 1 or sn is None or sp is None:
                continue
            want = m * (m + 1) / 2
            if abs(sn + sp - want) > h_of(neg["sum"]) + h_of(pos["sum"]) + 1e-9:
                out.append({"s": b["s"], "kind": "table", "decision": False,
                            "note": f"{b['name']}: the {m} ranks that aren’t ties add up to {m} × {m + 1} / 2 = {want:g}, but its Sums of Ranks add up to {sn + sp:g}."})
                continue
            z_t = val("z")
            z = _num_cell(z_t)
            if z is not None:
                z0 = abs(sn - m * (m + 1) / 4) / math.sqrt(m * (m + 1) * (2 * m + 1) / 24)
                zlo = max(z0 - (h_of(neg["sum"]) + h_of(pos["sum"])) / math.sqrt(m * (m + 1) * (2 * m + 1) / 24), 0)
                if abs(z) + h_of(z_t) < zlo - 1e-9:
                    out.append({"s": sline("z"), "kind": "test", "decision": False,
                                "note": f"{b['name']}: its Sums of Ranks give a Z of at least {zlo:.3f} in size (ties only make it larger), not {z_t}."})
                    continue
                p_t_ = val("p2")
                pc = _p_cell(p_t_)
                if pc is not None:
                    rel, pv, dp = pc
                    plo, phi = p_z(abs(z) + h_of(z_t)), p_z(max(abs(z) - h_of(z_t), 0))
                    n_judged += 1
                    if not _fits(rel, pv, dp, plo, phi):
                        real = p_z(abs(z))
                        out.append({"s": sline("p2"), "kind": "test", "decision": (real < 0.05) != (pv < 0.05 if rel == "=" else rel in ("<", "≤")),
                                    "note": f"{b['name']}: Z = {z_t} gives {_fmt_p(real, _p_lead(p_t_))}, not Asymp. Sig. (2-tailed) {_p_said(p_t_)}."})
            continue
        ns = [r["n"] for r in rows]
        n = sum(ns)
        if b["total"] is not None and b["total"] != n:
            out.append({"s": b["s"], "kind": "table", "decision": False,
                        "note": f"{b['name']}: its groups’ N add up to {n}, not its Total {b['total']}."})
            continue
        mrs = [(_num_cell(r["mr"]), r["mr"]) for r in rows]
        if any(m_ is None for m_, _ in mrs):
            continue
        tot = sum(r["n"] * m_ for r, (m_, _) in zip(rows, mrs))
        if abs(tot - n * (n + 1) / 2) > sum(r["n"] * h_of(t) for r, (_, t) in zip(rows, mrs)) + 1e-6:
            out.append({"s": b["s"], "kind": "table", "decision": False,
                        "note": f"{b['name']}: each group’s N × Mean Rank add up to {tot:.2f}, but the ranks 1 to {n} add up to {n * (n + 1) / 2:g}."})
            continue
        if "u" in vals and len(rows) == 2:              # Mann–Whitney
            r1, r2 = rows
            s1, s2 = _num_cell(r1["sum"]), _num_cell(r2["sum"])
            u_t, w_t, z_t = val("u"), val("w"), val("z")
            u, w, z = _num_cell(u_t), _num_cell(w_t), _num_cell(z_t)
            if s1 is not None and s2 is not None:
                if abs(s1 + s2 - n * (n + 1) / 2) > h_of(r1["sum"]) + h_of(r2["sum"]) + 1e-9:
                    out.append({"s": b["s"], "kind": "table", "decision": False,
                                "note": f"{b['name']}: its two Sums of Ranks add up to {s1 + s2:g}, but the ranks 1 to {n} add up to {n * (n + 1) / 2:g}."})
                    continue
                u1, u2 = s1 - r1["n"] * (r1["n"] + 1) / 2, s2 - r2["n"] * (r2["n"] + 1) / 2
                if u is not None and min(abs(u - u1), abs(u - u2)) > h_of(u_t) + max(h_of(r1["sum"]), h_of(r2["sum"])) + 1e-9:
                    out.append({"s": sline("u"), "kind": "test", "decision": False,
                                "note": f"{b['name']}: Mann-Whitney U is a group’s Sum of Ranks less n(n + 1)/2, {_like(min(u1, u2), u_t)} (or "
                                        f"{_like(max(u1, u2), u_t)}), not {u_t}."})
                    continue
                if w is not None and min(abs(w - s1), abs(w - s2)) > h_of(w_t) + max(h_of(r1["sum"]), h_of(r2["sum"])) + 1e-9:
                    out.append({"s": sline("w"), "kind": "test", "decision": False,
                                "note": f"{b['name']}: Wilcoxon W is one group’s Sum of Ranks ({r1['sum']} or {r2['sum']}), not {w_t}."})
                    continue
            if u is not None and z is not None:
                n1, n2 = r1["n"], r2["n"]
                sd0 = math.sqrt(n1 * n2 * (n + 1) / 12)
                zlo = max(abs(u - n1 * n2 / 2) - h_of(u_t), 0) / sd0
                if abs(z) + h_of(z_t) < zlo - 1e-9:
                    out.append({"s": sline("z"), "kind": "test", "decision": False,
                                "note": f"{b['name']}: U = {u_t} with groups of {n1} and {n2} gives a Z of at least {zlo:.3f} in size (ties only make "
                                        f"it larger), not {z_t}."})
                    continue
                p_t_ = val("p2")
                pc = _p_cell(p_t_)
                if pc is not None:
                    rel, pv, dp = pc
                    plo, phi = p_z(abs(z) + h_of(z_t)), p_z(max(abs(z) - h_of(z_t), 0))
                    n_judged += 1
                    if not _fits(rel, pv, dp, plo, phi):
                        real = p_z(abs(z))
                        out.append({"s": sline("p2"), "kind": "test", "decision": (real < 0.05) != (pv < 0.05 if rel == "=" else rel in ("<", "≤")),
                                    "note": f"{b['name']}: Z = {z_t} gives {_fmt_p(real, _p_lead(p_t_))}, not Asymp. Sig. (2-tailed) {_p_said(p_t_)}."})
            continue
        hk = "h" if "h" in vals else "chi" if "chi" in vals else None
        if hk is None:
            continue
        k = len(rows)
        df_t = val("df")
        if re.fullmatch(r"\d+", df_t) and int(df_t) != k - 1:
            out.append({"s": sline("df"), "kind": "test", "decision": False, "note": f"{b['name']}: {k} groups give {k - 1} df, not {df_t}."})
            continue
        h_t = val(hk)
        h = _num_cell(h_t)
        if h is None:
            continue
        c = (n + 1) / 2
        lo = 12 / (n * (n + 1)) * sum(r["n"] * max(abs(m_ - c) - h_of(t), 0) ** 2 for r, (m_, t) in zip(rows, mrs))
        if h + h_of(h_t) < lo - 1e-6:
            out.append({"s": sline(hk), "kind": "test", "decision": False,
                        "note": f"{b['name']}: its Mean Ranks give {'H' if hk == 'h' else 'a Chi-Square'} of at least {lo:.3f} (ties only make it larger), not {h_t}."})
            continue
        if re.fullmatch(r"\d+", df_t) and int(df_t) > 0:
            note, judged = _lg_p("", h_t, int(df_t), val("p"), b["name"])
            n_judged += 1 if judged else 0
            if note:
                out.append(dict(note, s=sline("p")))
    if counted is not None:
        counted[0] += n_judged
    return out



# 6.6, round 148 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): SPSS's reliability analysis, read together (nearly
# every thesis reports Cronbach's alpha). Scale Statistics: Std. Deviation = √Variance; Mean = the items' means added. Reliability
# Statistics: alpha = k/(k − 1) × (1 − Σ sᵢ² / Variance) from the Item Statistics' Std. Deviations; the alpha on standardized items
# = k r̄ / (1 + (k − 1) r̄) from Summary Item Statistics' mean inter-item correlation; N of Items = the items listed. Item-Total
# Statistics, item by item: Scale Mean if Item Deleted = the Mean less the item's; Corrected Item-Total Correlation =
# (V − V₋ᵢ − sᵢ²) / (2 sᵢ √V₋ᵢ); Cronbach's Alpha if Item Deleted = (k − 1)/(k − 2) × (1 − (Σ s² − sᵢ²) / V₋ᵢ). Summary Item
# Statistics: the mean, least and greatest item mean, Range and Maximum / Minimum. All over SPSS's rounding. Where one wrong figure
# would make many disagree (an item's Std. Deviation, the Variance), one note says so instead of many.
_RL_TITLE = re.compile(r"^(?P<t>reliability\s+statistics|item\s+statistics|summary\s+item\s+statistics|item-total\s+statistics|scale\s+statistics|"
                       r"case\s+processing\s+summary|inter-item\s+(?:correlation|covariance)\s+matrix|anova|hotelling's\s+t-squared\s+test|"
                       r"intraclass\s+correlation\s+coefficient)(?:[a-e](?:,[a-e])*)?$", re.I)


def _rl_read(lines):
    got = {"rel": None, "items": [], "summary": {}, "itemtotal": [], "scale": None, "valid": None}
    mode, heads = None, None
    for raw in lines:
        cells = _cells_keep(raw)
        filled = [(c or "").strip() for c in cells if (c or "").strip()]
        if not filled:
            continue
        if len(filled) == 1:
            tm = _RL_TITLE.match(filled[0])
            if tm:
                t = tm.group("t").lower()
                mode = {"reliability statistics": "rel", "item statistics": "items", "summary item statistics": "summary",
                        "item-total statistics": "itemtotal", "scale statistics": "scale", "case processing summary": "cases"}.get(t)
                heads = None
                continue
            if not re.match(r"^[a-e]\.\s", filled[0]):
                mode = None
            continue
        low = [(c or "").strip().lower() for c in cells]
        if mode in ("rel", "scale"):
            if any("alpha" in c or c == "variance" or c == "n of items" for c in low) and not any(_num_cell(c) is not None for c in filled):
                heads = low
                continue
            if heads and all(_num_cell(c) is not None for c in filled):
                row = {h: (cells[j] or "").strip() for j, h in enumerate(heads) if j < len(cells) and h}
                if mode == "rel":
                    got["rel"] = (row, raw.strip()[:500])
                else:
                    got["scale"] = (row, raw.strip()[:500])
            continue
        if mode in ("items", "summary", "itemtotal"):
            if not (cells[0] or "").strip() and not any(_num_cell(c) is not None for c in filled):
                heads = low
                continue
            if heads is None or not (cells[0] or "").strip():
                continue
            row = {h: (cells[j] or "").strip() for j, h in enumerate(heads) if j < len(cells) and h}
            row["_name"], row["_s"] = (cells[0] or "").strip(), raw.strip()[:500]
            if mode == "items":
                got["items"].append(row)
            elif mode == "summary":
                got["summary"][row["_name"].lower()] = row
            else:
                got["itemtotal"].append(row)
            continue
        if mode == "cases" and any(c == "valid" for c in low):
            nums = [c for c in filled if re.fullmatch(r"\d+", c)]
            if nums:
                got["valid"] = int(nums[0])
    return got


def spss_rel_check(text, counted=None):
    out = []
    got = _rl_read(str(text or "").split("\n"))
    if not (got["rel"] or got["scale"] or got["itemtotal"]):
        return out
    h_of = lambda t: 0.5 * 10 ** -_dec(str(t).lstrip("-−"))
    num = lambda r, k: _num_cell(r.get(k, "")) if r else None
    items = [r for r in got["items"] if num(r, "mean") is not None and num(r, "std. deviation") is not None]
    k = len(items)
    rel, rel_s = got["rel"] if got["rel"] else ({}, "")
    sc, sc_s = got["scale"] if got["scale"] else ({}, "")
    V_t, V = sc.get("variance", ""), num(sc, "variance")
    v_ok = V is not None and V > 0
    # the Scale Statistics: SD = √Variance; Mean = Σ item means
    sd = num(sc, "std. deviation")
    if v_ok and sd is not None:
        lo_, hi_ = math.sqrt(max(V - h_of(V_t), 0)), math.sqrt(V + h_of(V_t))
        if sd + h_of(sc["std. deviation"]) < lo_ - 1e-9 or sd - h_of(sc["std. deviation"]) > hi_ + 1e-9:
            out.append({"s": sc_s, "kind": "table", "decision": False,
                        "note": f"Scale Statistics: its Std. Deviation is the square root of its Variance, √{V_t} = {_like(math.sqrt(V), sc['std. deviation'])}, "
                                f"not {sc['std. deviation']}."})
            v_ok = False                              # (which of the two is wrong can't be told: nothing more rests on them)
    mean_ok = False
    sm = num(sc, "mean")
    if sm is not None and k >= 2:
        tot = sum(num(r, "mean") for r in items)
        slack = sum(h_of(r["mean"]) for r in items) + h_of(sc["mean"])
        if abs(tot - sm) > slack + 1e-9:
            out.append({"s": sc_s, "kind": "table", "decision": False,
                        "note": f"Scale Statistics: its Mean is the items’ means added, about {_like(tot, sc['mean'])} from Item Statistics, not {sc['mean']}."})
        else:
            mean_ok = True
    # N of Items
    n_bad = False
    for where, row, s_ in (("Reliability Statistics", rel, rel_s), ("Scale Statistics", sc, sc_s)):
        n_t = row.get("n of items", "")
        if k >= 2 and re.fullmatch(r"\d+", n_t) and int(n_t) != k:
            out.append({"s": s_, "kind": "table", "decision": False, "note": f"{where}: N of Items is {n_t}, but Item Statistics lists {k} items."})
            k, n_bad = 0, True                        # (the lists don't agree: nothing that counts items is judged)
            break
    sds = [(num(r, "std. deviation"), h_of(r["std. deviation"])) for r in items]
    A_lo, A_hi = sum(max(s_ - h, 0) ** 2 for s_, h in sds), sum((s_ + h) ** 2 for s_, h in sds)
    # each item's alpha if deleted (needs no Variance)
    by_name = {r["_name"].lower(): r for r in items}
    it_notes, it_bad, it_good = [], 0, []
    for r in got["itemtotal"]:
        it = by_name.get(r["_name"].lower())
        if it is None or k < 3:
            continue
        vi_t, vi = r.get("scale variance if item deleted", ""), num(r, "scale variance if item deleted")
        ai_t, ai = r.get("cronbach's alpha if item deleted", ""), num(r, "cronbach's alpha if item deleted")
        if vi is None or ai is None or vi <= 0:
            continue
        s_i, h_i = num(it, "std. deviation"), h_of(it["std. deviation"])
        rest_lo = sum(max(s_ - h, 0) ** 2 for (s_, h), x in zip(sds, items) if x is not it)
        rest_hi = sum((s_ + h) ** 2 for (s_, h), x in zip(sds, items) if x is not it)
        f = (k - 1) / (k - 2)
        a_lo, a_hi = f * (1 - rest_hi / max(vi - h_of(vi_t), 1e-12)), f * (1 - rest_lo / (vi + h_of(vi_t)))
        if not (ai + h_of(ai_t) < a_lo - 1e-9 or ai - h_of(ai_t) > a_hi + 1e-9):
            it_good.append(it)
        else:
            it_bad += 1
            want = f * (1 - sum(s_ ** 2 for (s_, h), x in zip(sds, items) if x is not it) / vi)
            it_notes.append({"s": r["_s"], "kind": "table", "decision": False,
                             "note": f"Item-Total Statistics, {r['_name']}: the alpha of the other {k - 1} items is {k - 1}/{k - 2} × (1 − their Σ s² / {vi_t}) "
                                     f"≈ {_like(want, ai_t)}, not {ai_t}."})
    # Cronbach's Alpha
    a_t, a = rel.get("cronbach's alpha", ""), num(rel, "cronbach's alpha")
    alpha_bad = False
    if v_ok and a is not None and k >= 2:
        f = k / (k - 1)
        lo_, hi_ = f * (1 - A_hi / max(V - h_of(V_t), 1e-12)), f * (1 - A_lo / (V + h_of(V_t)))
        if a + h_of(a_t) < lo_ - 1e-9 or a - h_of(a_t) > hi_ + 1e-9:
            alpha_bad = True
            want = f * (1 - sum(s_ ** 2 for s_, _ in sds) / V)
            if it_bad >= 2:                           # (the alphas if deleted disagree too: the items' Std. Deviations are in question)
                one = it_good[0] if len(it_good) == 1 else None
                need = (1 - a * (k - 1) / k) * V - sum(s_ ** 2 for (s_, h), x in zip(sds, items) if x is not one) if one else None
                if one is not None and need and need > 0:
                    out.append({"s": one["_s"], "kind": "table", "decision": False,
                                "note": f"Item Statistics, {one['_name']}: its Std. Deviation ({one['std. deviation']}) agrees neither with Cronbach’s Alpha "
                                        f"({a_t}) nor with the other items’ alphas if deleted; about {_like(math.sqrt(need), one['std. deviation'])} would fit "
                                        f"them all: it is likely mistyped."})
                else:
                    out.append({"s": rel_s, "kind": "table", "decision": False,
                                "note": f"Item Statistics: its Std. Deviations agree neither with Cronbach’s Alpha ({a_t}) nor with {it_bad} of the alphas "
                                        f"if an item is deleted: one of them is likely mistyped (with them Cronbach’s Alpha would be {_like(want, a_t)})."})
                it_notes = []
            else:
                out.append({"s": rel_s, "kind": "table", "decision": False,
                            "note": f"Reliability Statistics: with the items’ Std. Deviations and the scale’s Variance ({V_t}), Cronbach’s Alpha is "
                                    f"{k}/{k - 1} × (1 − Σ s² / {V_t}) ≈ {_like(want, a_t)}, not {a_t}."})
    out += it_notes[:3]
    # the alpha on standardized items, from the mean inter-item correlation
    st_t, st = rel.get("cronbach's alpha based on standardized items", ""), num(rel, "cronbach's alpha based on standardized items")
    iic = got["summary"].get("inter-item correlations")
    nk = None if n_bad else int(rel["n of items"]) if re.fullmatch(r"\d+", rel.get("n of items", "")) else (k or None)
    if st is not None and iic and num(iic, "mean") is not None and nk and nk >= 2:
        rb_t, rb = iic["mean"], num(iic, "mean")
        g = lambda x: nk * x / (1 + (nk - 1) * x) if 1 + (nk - 1) * x > 0 else None
        ends = [g(rb - h_of(rb_t)), g(rb + h_of(rb_t))]
        if None not in ends and (st + h_of(st_t) < min(ends) - 1e-9 or st - h_of(st_t) > max(ends) + 1e-9):
            out.append({"s": rel_s, "kind": "table", "decision": False,
                        "note": f"Reliability Statistics: with a mean inter-item correlation of {rb_t} over {nk} items, the alpha based on standardized items is "
                                f"{nk} × {rb_t} / (1 + {nk - 1} × {rb_t}) ≈ {_like(g(rb), st_t)}, not {st_t}."})
    # item by item: Scale Mean if Item Deleted; the Corrected Item-Total Correlation
    for r in got["itemtotal"]:
        it = by_name.get(r["_name"].lower())
        if it is None:
            continue
        smi_t, smi = r.get("scale mean if item deleted", ""), num(r, "scale mean if item deleted")
        if mean_ok and smi is not None:
            want = sm - num(it, "mean")
            if abs(smi - want) > h_of(smi_t) + h_of(sc["mean"]) + h_of(it["mean"]) + 1e-9:
                out.append({"s": r["_s"], "kind": "table", "decision": False,
                            "note": f"Item-Total Statistics, {r['_name']}: its Scale Mean if Item Deleted is the scale’s Mean less the item’s, "
                                    f"{sc['mean']} − {it['mean']} = {_like(want, smi_t)}, not {smi_t}."})
                continue
        ri_t, ri = r.get("corrected item-total correlation", ""), num(r, "corrected item-total correlation")
        vi_t, vi = r.get("scale variance if item deleted", ""), num(r, "scale variance if item deleted")
        if not v_ok or alpha_bad or ri is None or vi is None or vi <= 0:
            continue
        s_i, h_i = num(it, "std. deviation"), h_of(it["std. deviation"])
        vals = []
        for VV in (V - h_of(V_t), V + h_of(V_t)):
            for vv in (vi - h_of(vi_t), vi + h_of(vi_t)):
                for ss in (s_i - h_i, s_i + h_i):
                    if ss > 0 and vv > 0:
                        vals.append((VV - vv - ss * ss) / (2 * ss * math.sqrt(vv)))
        if vals and (ri + h_of(ri_t) < min(vals) - 1e-9 or ri - h_of(ri_t) > max(vals) + 1e-9):
            want = (V - vi - s_i * s_i) / (2 * s_i * math.sqrt(vi))
            out.append({"s": r["_s"], "kind": "table", "decision": False,
                        "note": f"Item-Total Statistics, {r['_name']}: the scale’s Variance ({V_t}), its Scale Variance if Item Deleted ({vi_t}) and its "
                                f"Std. Deviation ({it['std. deviation']}) give a Corrected Item-Total Correlation of about {_like(want, ri_t)}, not {ri_t}."})
    # Summary Item Statistics
    for key, what in (("item means", "Item Means"), ("inter-item correlations", "Inter-Item Correlations")):
        row = got["summary"].get(key)
        if not row:
            continue
        mn, mx = num(row, "minimum"), num(row, "maximum")
        note = None
        if key == "item means" and k >= 2 and mn is not None and mx is not None:
            ms = [(num(r, "mean"), h_of(r["mean"])) for r in items]
            lo_m, hi_m = min(m for m, _ in ms), max(m for m, _ in ms)
            avg = sum(m for m, _ in ms) / k
            mean_ = num(row, "mean")
            if abs(mn - lo_m) > h_of(row["minimum"]) + ms[0][1] + 1e-9:
                note = f"Summary Item Statistics, {what}: its Minimum is the smallest item mean, {_like(lo_m, row['minimum'])}, not {row['minimum']}."
            elif abs(mx - hi_m) > h_of(row["maximum"]) + ms[0][1] + 1e-9:
                note = f"Summary Item Statistics, {what}: its Maximum is the largest item mean, {_like(hi_m, row['maximum'])}, not {row['maximum']}."
            elif mean_ is not None and abs(mean_ - avg) > h_of(row["mean"]) + ms[0][1] + 1e-9:
                note = f"Summary Item Statistics, {what}: its Mean is the average of the item means, {_like(avg, row['mean'])}, not {row['mean']}."
        rg = num(row, "range")
        if note is None and rg is not None and mn is not None and mx is not None:
            if abs(rg - (mx - mn)) > h_of(row["range"]) + h_of(row["minimum"]) + h_of(row["maximum"]) + 1e-9:
                note = (f"Summary Item Statistics, {what}: its Range is Maximum − Minimum, {row['maximum']} − {row['minimum']} = "
                        f"{_like(mx - mn, row['range'])}, not {row['range']}.")
        ratio = num(row, "maximum / minimum")
        if note is None and ratio is not None and mn is not None and mx is not None and mn - h_of(row["minimum"]) > 0:
            r_lo = (mx - h_of(row["maximum"])) / (mn + h_of(row["minimum"]))
            r_hi = (mx + h_of(row["maximum"])) / (mn - h_of(row["minimum"]))
            if ratio + h_of(row["maximum / minimum"]) < r_lo - 1e-9 or ratio - h_of(row["maximum / minimum"]) > r_hi + 1e-9:
                note = (f"Summary Item Statistics, {what}: its Maximum / Minimum is {row['maximum']} / {row['minimum']} ≈ {_like(mx / mn, row['maximum / minimum'])}, "
                        f"not {row['maximum / minimum']}.")
        if note:
            out.append({"s": row["_s"], "kind": "table", "decision": False, "note": note})
    return out



# 6.6, round 149 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): three more SPSS outputs. Tests of Between-Subjects
# Effects (a two-way ANOVA or an ANCOVA): each Mean Square = SS / df, F = MS / the Error's MS, Sig. from F, Partial Eta Squared =
# SS / (SS + the Error's SS), Noncent. Parameter = F × df, the Corrected Model and the Error adding up to the Corrected Total (SS
# and df), Total's df one more, and the footnote's R Squared and Adjusted R Squared. Collinearity Statistics: VIF = 1 / Tolerance.
# Skewness and kurtosis: their Std. Errors depend on N alone (Descriptive Statistics, Frequencies' Statistics, Explore's
# Descriptives); Explore's Descriptives also give Variance = SD², Range = Maximum − Minimum, the mean's Std. Error = SD / √N and its
# 95% interval = Mean ∓ t(.975, N − 1) × that Std. Error.
_GL_TITLE = re.compile(r"^tests\s+of\s+between-?\s?subjects\s+effects(?:[a-e](?:,[a-e])*)?$", re.I)
_GL_COL = {"ss": re.compile(r"^type\s+(?:i{1,3}|iv)\s+sum\s+of\s+squares$", re.I), "df": re.compile(r"^df$", re.I),
           "ms": re.compile(r"^mean\s+square$", re.I), "f": re.compile(r"^F$"), "p": re.compile(r"^sig\.?$", re.I),
           "eta": re.compile(r"^partial\s+eta\s+squared$", re.I), "nc": re.compile(r"^noncent\.?\s+parameter$", re.I)}
_GL_R2 = re.compile(r"R\s+Squared\s*=\s*(?P<r>-?\d*\.\d+)\s*\(\s*Adjusted\s+R\s+Squared\s*=\s*(?P<a>-?\d*\.\d+)\s*\)", re.I)


def spss_glm_check(text, counted=None):
    out, n_judged = [], 0
    lines = str(text or "").split("\n")
    i = 0
    while i < len(lines):
        if not _GL_TITLE.match(lines[i].strip()):
            i += 1
            continue
        cols, rows, foot, j = None, {}, None, i + 1
        while j < len(lines):
            raw = lines[j]
            cells = _cells_keep(raw)
            filled = [(c or "").strip() for c in cells if (c or "").strip()]
            j += 1
            if not filled:
                if rows:
                    break
                continue
            if len(filled) == 1:
                m = _GL_R2.search(filled[0])
                if m:
                    foot = (m, raw.strip()[:500])
                    continue
                if re.match(r"^(?:dependent\s+variable|[a-e]\.\s)", filled[0], re.I):
                    continue
                if rows:
                    break
                continue
            heads = {}
            for k, c in enumerate(cells):
                for key, rx in _GL_COL.items():
                    if rx.match((c or "").strip()) and key not in heads:
                        heads[key] = k
            if {"ss", "df"} <= set(heads):
                cols = heads
                continue
            if cols is None:
                continue
            src = (cells[0] or "").strip()
            get = lambda key: re.sub(r"(?<=\d)[a-e]$", "", (cells[cols[key]] or "").strip()) if key in cols and cols[key] < len(cells) else ""
            if src:
                rows[src.lower()] = {"src": src, **{key: get(key) for key in cols}, "s": raw.strip()[:500]}
        i = j
        if not rows or "error" not in rows:
            continue
        h_of = lambda t: 0.5 * 10 ** -_dec(str(t).lstrip("-−"))
        num = lambda r, k: _num_cell(r.get(k, ""))
        er = rows["error"]
        sse, dfe, mse = num(er, "ss"), er.get("df", ""), num(er, "ms")
        if sse is None or not re.fullmatch(r"\d+", dfe) or int(dfe) < 1:
            continue
        dfe = int(dfe)
        if mse is not None and abs(mse - sse / dfe) > h_of(er["ms"]) + h_of(er["ss"]) / dfe + 1e-9:
            out.append({"s": er["s"], "kind": "test", "decision": False,
                        "note": f"Tests of Between-Subjects Effects, Error: its Mean Square is its Sum of Squares over its df, {er['ss']} / {dfe} = "
                                f"{_like(sse / dfe, er['ms'])}, not {er['ms']}."})
            mse_t = None
        else:
            mse_t = er.get("ms") or None
        mse_lo = (mse - h_of(mse_t)) if mse_t else (sse - h_of(er["ss"])) / dfe
        mse_hi = (mse + h_of(mse_t)) if mse_t else (sse + h_of(er["ss"])) / dfe
        for key, r in rows.items():
            if key in ("error", "total", "corrected total"):
                continue
            ss, df_t, ms, f = num(r, "ss"), r.get("df", ""), num(r, "ms"), num(r, "f")
            if ss is None or not re.fullmatch(r"\d+", df_t) or int(df_t) < 1:
                continue
            d = int(df_t)
            where = f"Tests of Between-Subjects Effects, {r['src']}"
            if ms is not None and abs(ms - ss / d) > h_of(r["ms"]) + h_of(r["ss"]) / d + 1e-9:
                out.append({"s": r["s"], "kind": "test", "decision": False,
                            "note": f"{where}: its Mean Square is its Sum of Squares over its df, {r['ss']} / {d} = {_like(ss / d, r['ms'])}, not {r['ms']}."})
                continue
            if f is not None and ms is not None and mse_lo > 0:
                f_lo, f_hi = (ms - h_of(r["ms"])) / mse_hi, (ms + h_of(r["ms"])) / mse_lo
                if f + h_of(r["f"]) < f_lo - 1e-9 or f - h_of(r["f"]) > f_hi + 1e-9:
                    out.append({"s": r["s"], "kind": "test", "decision": False,
                                "note": f"{where}: F is its Mean Square over the Error’s, {r['ms']} / {mse_t or _like(sse / dfe, r['ms'])} ≈ "
                                        f"{_like(ms / ((mse_lo + mse_hi) / 2), r['f'])}, not {r['f']}."})
                    continue
            pc = _p_cell(r.get("p", ""))
            if f is not None and pc is not None:
                rel, pv, dp = pc
                plo, phi = p_f(f + h_of(r["f"]), d, dfe), p_f(max(f - h_of(r["f"]), 0.0), d, dfe)
                if plo is not None and phi is not None:
                    n_judged += 1
                    if not _fits(rel, pv, dp, plo, phi):
                        real = p_f(f, d, dfe)
                        out.append({"s": r["s"], "kind": "test", "decision": (real < 0.05) != (pv < 0.05 if rel == "=" else rel in ("<", "≤")),
                                    "note": f"{where}: F = {r['f']} on {d} and {dfe} df gives {_fmt_p(real, _p_lead(r['p']))}, not Sig. {_p_said(r['p'])}."})
                        continue
            eta = num(r, "eta")
            if eta is not None:
                ends = []
                for a in (ss - h_of(r["ss"]), ss + h_of(r["ss"])):
                    for b in (sse - h_of(er["ss"]), sse + h_of(er["ss"])):
                        if a + b > 0:
                            ends.append(a / (a + b))
                if ends and (eta + h_of(r["eta"]) < min(ends) - 1e-9 or eta - h_of(r["eta"]) > max(ends) + 1e-9):
                    out.append({"s": r["s"], "kind": "test", "decision": False,
                                "note": f"{where}: its Partial Eta Squared is SS / (SS + the Error’s SS), {r['ss']} / ({r['ss']} + {er['ss']}) ≈ "
                                        f"{_like(ss / (ss + sse), r['eta'])}, not {r['eta']}."})
                    continue
            nc = num(r, "nc")
            if nc is not None and f is not None and abs(nc - f * d) > h_of(r["nc"]) + d * h_of(r["f"]) + 1e-9:
                out.append({"s": r["s"], "kind": "test", "decision": False,
                            "note": f"{where}: its Noncent. Parameter is F × df, {r['f']} × {d} = {_like(f * d, r['nc'])}, not {r['nc']}."})
        cm, ct, tt = rows.get("corrected model"), rows.get("corrected total"), rows.get("total")
        if cm and ct and num(cm, "ss") is not None and num(ct, "ss") is not None:
            tot = num(cm, "ss") + sse
            if abs(tot - num(ct, "ss")) > h_of(cm["ss"]) + h_of(er["ss"]) + h_of(ct["ss"]) + 1e-9:
                out.append({"s": ct["s"], "kind": "test", "decision": False,
                            "note": f"Tests of Between-Subjects Effects: the Corrected Model’s and the Error’s Sums of Squares add up to "
                                    f"{_like(tot, ct['ss'])}, not the Corrected Total’s {ct['ss']}."})
            elif re.fullmatch(r"\d+", cm.get("df", "")) and re.fullmatch(r"\d+", ct.get("df", "")) and int(cm["df"]) + dfe != int(ct["df"]):
                out.append({"s": ct["s"], "kind": "test", "decision": False,
                            "note": f"Tests of Between-Subjects Effects: the Corrected Model’s and the Error’s df add up to {int(cm['df']) + dfe}, not the "
                                    f"Corrected Total’s {ct['df']}."})
            elif tt and re.fullmatch(r"\d+", tt.get("df", "")) and re.fullmatch(r"\d+", ct.get("df", "")) and int(tt["df"]) != int(ct["df"]) + 1:
                out.append({"s": tt["s"], "kind": "test", "decision": False,
                            "note": f"Tests of Between-Subjects Effects: Total has N df and Corrected Total N − 1, so {int(ct['df']) + 1}, not {tt['df']}."})
            elif foot and re.fullmatch(r"\d+", ct.get("df", "")):
                m, s_ = foot
                r2_t, ar_t = m.group("r"), m.group("a")
                ssm, sst = num(cm, "ss"), num(ct, "ss")
                dft = int(ct["df"])
                if sst > 0:
                    r2 = ssm / sst
                    slack = (h_of(cm["ss"]) + h_of(ct["ss"])) / sst * 2
                    if abs(float(r2_t) - r2) > h_of(r2_t) + slack + 1e-9:
                        out.append({"s": s_, "kind": "test", "decision": False,
                                    "note": f"Tests of Between-Subjects Effects: R Squared is the Corrected Model’s Sum of Squares over the Corrected Total’s, "
                                            f"{cm['ss']} / {ct['ss']} ≈ {_like(r2, r2_t)}, not {r2_t}."})
                    elif dft > 0:
                        ar = 1 - (sse / dfe) / (sst / dft)
                        if abs(float(ar_t) - ar) > h_of(ar_t) + slack * 2 + 1e-9:
                            out.append({"s": s_, "kind": "test", "decision": False,
                                        "note": f"Tests of Between-Subjects Effects: Adjusted R Squared is 1 − the Error’s Mean Square / (the Corrected "
                                                f"Total’s SS / its df), 1 − {_like(sse / dfe, er.get('ms') or '0.000')} / ({ct['ss']} / {dft}) ≈ {_like(ar, ar_t)}, "
                                                f"not {ar_t}."})
    if counted is not None:
        counted[0] += n_judged
    return out


def collin_check(text):
    """Collinearity Statistics: VIF = 1 / Tolerance (Tolerance at most 1)."""
    out, cols = [], None
    for raw in str(text or "").split("\n"):
        cells = _cells_keep(raw)
        low = [(c or "").strip().lower() for c in cells]
        if "tolerance" in low and "vif" in low:
            cols = (low.index("tolerance"), low.index("vif"))
            continue
        if not any(low):
            cols = None
            continue
        if cols is None or max(cols) >= len(cells):
            continue
        t_t, v_t = (cells[cols[0]] or "").strip(), (cells[cols[1]] or "").strip()
        t, v = _num_cell(t_t), _num_cell(v_t)
        if t is None or v is None:
            continue
        lab = _row_label(cells)[1] or "A predictor"
        h_t, h_v = 0.5 * 10 ** -_dec(t_t), 0.5 * 10 ** -_dec(v_t)
        if t - h_t > 1 or v + h_v < 1:
            out.append({"s": raw.strip()[:500], "kind": "value", "decision": False,
                        "note": f"{lab}: a Tolerance is at most 1 and a VIF at least 1 (the table gives {t_t} and {v_t})."})
            continue
        if t - h_t > 0 and (v + h_v < 1 / (t + h_t) - 1e-9 or v - h_v > 1 / (t - h_t) + 1e-9):
            out.append({"s": raw.strip()[:500], "kind": "value", "decision": False,
                        "note": f"{lab}: VIF is 1 / Tolerance, 1 / {t_t} ≈ {_like(1 / t, v_t)}, not {v_t}."})
    return out


def _se_skew(n):
    return math.sqrt(6.0 * n * (n - 1) / ((n - 2) * (n + 1) * (n + 3))) if n > 2 else None


def _se_kurt(n):
    s = _se_skew(n)
    return math.sqrt(4.0 * (n * n - 1) * s * s / ((n - 3) * (n + 5))) if s and n > 3 else None


def _se_note(lab, what, n, se_t):
    se = _num_cell(se_t)
    want = (_se_skew if what == "skewness" else _se_kurt)(n)
    if se is None or want is None:
        return None
    if abs(se - want) > 0.5 * 10 ** -_dec(se_t) + 1e-9:
        return (f"{lab}: with N = {n}, the Std. Error of its {what} is {_like(want, se_t)} (it depends on N alone), not {se_t}.")
    return None


def skew_se_check(text):
    """The Std. Errors of skewness and kurtosis, from N (SPSS's Descriptive Statistics, Frequencies' Statistics, Explore's
    Descriptives, with its Variance, Range, the mean's Std. Error and interval)."""
    out = []
    lines = str(text or "").split("\n")
    # Descriptive Statistics: "N … Skewness [Std. Error] Kurtosis [Std. Error]"
    cols = None
    for k, raw in enumerate(lines):
        cells = _cells_keep(raw)
        low = [(c or "").strip().lower() for c in cells]
        if "skewness" in low and "n" in low and k + 1 < len(lines):
            nxt = [(c or "").strip().lower() for c in _cells_keep(lines[k + 1])]
            sk, ku = low.index("skewness"), (low.index("kurtosis") if "kurtosis" in low else None)
            se_sk = sk + 1 if sk + 1 < len(nxt) and nxt[sk + 1] == "std. error" else None
            se_ku = ku + 1 if ku is not None and ku + 1 < len(nxt) and nxt[ku + 1] == "std. error" else None
            cols = (low.index("n"), se_sk, se_ku) if (se_sk or se_ku) else None
            continue
        if not any(low):
            cols = None
            continue
        if cols is None or not re.search(r"[A-Za-z]", (cells[0] or "")) or "statistic" in low:
            continue
        n_t = (cells[cols[0]] or "").strip() if cols[0] < len(cells) else ""
        if not re.fullmatch(r"\d+", n_t):
            continue
        for idx, what in ((cols[1], "skewness"), (cols[2], "kurtosis")):
            if idx is not None and idx < len(cells) and (cells[idx] or "").strip():
                note = _se_note((cells[0] or "").strip(), what, int(n_t), (cells[idx] or "").strip())
                if note:
                    out.append({"s": raw.strip()[:500], "kind": "value", "decision": False, "note": note})
                    break
    # Frequencies' Statistics: a column for each variable (its values the row's last cells); rows N Valid, Std. Error of Skewness / Kurtosis
    names, valid = None, None
    for raw in lines:
        cells = _cells_keep(raw)
        first = (cells[0] or "").strip().lower()
        if first == "statistics" and len([c for c in cells if (c or "").strip()]) == 1:
            names, valid = [], None
            continue
        if names is None:
            continue
        if not any((c or "").strip() for c in cells):
            names = None
            continue
        if names == [] and not first:
            names = [(c or "").strip() for c in cells if (c or "").strip()]
            continue
        k_ = len(names)
        if not k_:
            continue
        if first in ("n", "") and len(cells) > 1 and (cells[1] or "").strip().lower() == "valid":
            valid = [(c or "").strip() for c in cells[-k_:]]
            continue
        m = re.match(r"^std\.?\s+error\s+of\s+(skewness|kurtosis)$", first)
        if m and valid and len(cells) > k_:
            for j, se_t in enumerate((c or "").strip() for c in cells[-k_:]):
                n_t = valid[j] if j < len(valid) else ""
                if re.fullmatch(r"\d+", n_t) and se_t:
                    note = _se_note(names[j], m.group(1), int(n_t), se_t)
                    if note:
                        out.append({"s": raw.strip()[:500], "kind": "value", "decision": False, "note": note})
    # Explore's Descriptives, with N from its Case Processing Summary
    ns, cur, stats_ = {}, None, {}
    mode = None
    for raw in lines:
        cells = _cells_keep(raw)
        filled = [(c or "").strip() for c in cells if (c or "").strip()]
        if not filled:
            continue
        if len(filled) == 1 and filled[0].lower() in ("case processing summary", "descriptives"):
            mode, cur = filled[0].lower(), None
            continue
        if mode == "case processing summary" and (cells[0] or "").strip() and len(cells) > 1 and re.fullmatch(r"\d+", (cells[1] or "").strip()):
            ns[(cells[0] or "").strip().lower()] = int((cells[1] or "").strip())
        elif mode == "descriptives" and len(cells) >= 4:
            if (cells[0] or "").strip():
                cur = (cells[0] or "").strip()
            lab = ((cells[1] or "").strip() + " " + (cells[2] or "").strip()).strip().lower()
            if cur and lab:
                stats_.setdefault(cur, {})[lab] = ((cells[3] or "").strip(), (cells[4] or "").strip() if len(cells) > 4 else "", raw.strip()[:500])
    for var, st in stats_.items():
        n = ns.get(var.lower())
        g = lambda key: st.get(key, ("", "", ""))
        sd_t, var_t = g("std. deviation")[0], g("variance")[0]
        sd, vv = _num_cell(sd_t), _num_cell(var_t)
        if sd is not None and vv is not None:
            h = 0.5 * 10 ** -_dec(sd_t)
            if vv + 0.5 * 10 ** -_dec(var_t) < (sd - h) ** 2 - 1e-9 or vv - 0.5 * 10 ** -_dec(var_t) > (sd + h) ** 2 + 1e-9:
                out.append({"s": g("variance")[2], "kind": "value", "decision": False,
                            "note": f"{var}: its Variance is its Std. Deviation squared, {sd_t}² ≈ {_like(sd * sd, var_t)}, not {var_t}."})
        mn, mx, rg = _num_cell(g("minimum")[0]), _num_cell(g("maximum")[0]), _num_cell(g("range")[0])
        if None not in (mn, mx, rg):
            slack = sum(0.5 * 10 ** -_dec(t) for t in (g("minimum")[0], g("maximum")[0], g("range")[0]))
            if abs(rg - (mx - mn)) > slack + 1e-9:
                out.append({"s": g("range")[2], "kind": "value", "decision": False,
                            "note": f"{var}: its Range is Maximum − Minimum, {g('maximum')[0]} − {g('minimum')[0]} = {_like(mx - mn, g('range')[0])}, not {g('range')[0]}."})
        if not n:
            continue
        m_t, se_t = g("mean")[0], g("mean")[1]
        m, se = _num_cell(m_t), _num_cell(se_t)
        if sd is not None and se is not None and n > 1:
            h = 0.5 * 10 ** -_dec(sd_t)
            lo_, hi_ = (sd - h) / math.sqrt(n), (sd + h) / math.sqrt(n)
            if se + 0.5 * 10 ** -_dec(se_t) < lo_ - 1e-9 or se - 0.5 * 10 ** -_dec(se_t) > hi_ + 1e-9:
                out.append({"s": g("mean")[2], "kind": "value", "decision": False,
                            "note": f"{var}: the Std. Error of its Mean is its Std. Deviation over √N, {sd_t} / √{n} = {_like(sd / math.sqrt(n), se_t)}, not {se_t}."})
                se = None
        lo_t = g("95% confidence interval for mean lower bound")[0]
        hi_t = g("upper bound")[0] or g("95% confidence interval for mean upper bound")[0]
        if m is not None and se is not None and n > 1:
            tc = _t_crit(n - 1)
            for tx, want, end_, sign in ((lo_t, m - tc * se, "lower", "−"), (hi_t, m + tc * se, "upper", "+")):
                v = _num_cell(tx)
                if v is not None and abs(v - want) > 0.5 * 10 ** -_dec(tx) + 0.5 * 10 ** -_dec(m_t) + tc * 0.5 * 10 ** -_dec(se_t) + 1e-9:
                    out.append({"s": g("mean")[2], "kind": "value", "decision": False,
                                "note": f"{var}: the 95% interval’s {end_} bound is its Mean {sign} t(.975, {n - 1}) × its Std. Error, about {_like(want, tx)}, not {tx}."})
                    break
        for key, what in (("skewness", "skewness"), ("kurtosis", "kurtosis")):
            se_k = g(key)[1]
            if se_k:
                note = _se_note(var, what, n, se_k)
                if note:
                    out.append({"s": g(key)[2], "kind": "value", "decision": False, "note": note})
    return out



# 6.6, round 150 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): SPSS's factor analysis, read together (theses that
# build or test a questionnaire). KMO and Bartlett's Test: KMO between 0 and 1; Bartlett's df = p(p − 1)/2 for p items; its Sig.
# from its Chi-Square. Total Variance Explained: each % of Variance = Total / p × 100; each Cumulative % the running sum; the
# eigenvalues falling and adding up to p; with principal components, the Extraction totals the same as the Initial; an orthogonal
# rotation's totals adding up to the extraction's. Communalities: with principal components, Initial 1.000; each Extraction the
# sum of its row's squared loadings (in the Component Matrix, and in the Rotated one too); together, the extraction's total. Each
# column of a component matrix: its squared loadings add up to its component's total. One mistyped loading or communality is
# named, not said twice. All over SPSS's rounding.
_FA_TITLE = re.compile(r"^(?P<t>kmo\s+and\s+bartlett'?s\s+test|communalities|total\s+variance\s+explained|component\s+matrix|factor\s+matrix|"
                       r"rotated\s+component\s+matrix|rotated\s+factor\s+matrix|pattern\s+matrix|structure\s+matrix|component\s+transformation\s+matrix|"
                       r"factor\s+transformation\s+matrix|component\s+score\s+coefficient\s+matrix|correlation\s+matrix|descriptive\s+statistics|"
                       r"component\s+correlation\s+matrix|factor\s+correlation\s+matrix|anti-image\s+matrices|goodness-of-fit\s+test)(?:[a-e](?:,[a-e])*)?$", re.I)


def _fa_read(lines):
    got = {"kmo": {}, "comm": [], "tve": [], "mat": [], "rot": []}
    mode = None
    for raw in lines:
        cells = _cells_keep(raw)
        filled = [(c or "").strip() for c in cells if (c or "").strip()]
        if not filled:
            continue
        if len(filled) == 1:
            tm = _FA_TITLE.match(filled[0])
            if tm:
                t = tm.group("t").lower()
                mode = ("kmo" if t.startswith("kmo") else "comm" if t == "communalities" else "tve" if t.startswith("total variance") else
                        "mat" if t in ("component matrix", "factor matrix") else "rot" if t in ("rotated component matrix", "rotated factor matrix") else None)
                continue
            if not re.match(r"^(?:[a-e]\.\s|extraction\s+method|rotation\s+method|component|factor)", filled[0], re.I):
                mode = None
            continue
        if re.match(r"^\s*(?:extraction|rotation)\s+method", raw, re.I):
            continue
        nums = [(c, _num_cell(c)) for c in filled if _num_cell(c) is not None]
        if mode == "kmo":
            low = raw.lower()
            if nums:
                key = ("kmo" if "kaiser" in low or "kmo" in low else "chi" if "chi-square" in low else "df" if re.search(r"(?:^|\t)df(?:\t|$)", low)
                       else "sig" if "sig" in low else None)
                if key:
                    got["kmo"][key] = (nums[-1][0], raw.strip()[:500])
        elif mode == "comm":
            lab = (cells[0] or "").strip()
            if lab and len(nums) == 2 and re.search(r"[A-Za-z]", lab):
                got["comm"].append((lab, nums[0][0], nums[1][0], raw.strip()[:500]))
        elif mode == "tve":
            lab = (cells[0] or "").strip()
            if re.fullmatch(r"\d{1,3}", lab) and len(nums) >= 4:
                got["tve"].append((int(lab), [c for c, _ in nums[1:]], raw.strip()[:500]))
        elif mode in ("mat", "rot"):
            lab = (cells[0] or "").strip()
            if lab and re.search(r"[A-Za-z]", lab) and nums and len(nums) == len(filled) - 1:
                got[mode].append((lab, [c for c, _ in nums], raw.strip()[:500]))
    return got


def spss_fa_check(text, counted=None):
    out, n_judged = [], 0
    got = _fa_read(str(text or "").split("\n"))
    if not (got["kmo"] or got["tve"] or got["comm"]):
        return out
    h_of = lambda t: 0.5 * 10 ** -_dec(str(t).lstrip("-−"))
    p = len(got["tve"]) or len(got["comm"]) or None
    kmo = got["kmo"]
    if "kmo" in kmo:
        v = _num_cell(kmo["kmo"][0])
        if v is not None and (v < 0 or v - h_of(kmo["kmo"][0]) > 1):
            out.append({"s": kmo["kmo"][1], "kind": "value", "decision": False,
                        "note": f"KMO and Bartlett's Test: the Kaiser-Meyer-Olkin measure lies between 0 and 1 (the table gives {kmo['kmo'][0]})."})
    if "df" in kmo and re.fullmatch(r"\d+", kmo["df"][0]):
        d = int(kmo["df"][0])
        if p and d != p * (p - 1) // 2:
            out.append({"s": kmo["df"][1], "kind": "test", "decision": False,
                        "note": f"KMO and Bartlett's Test: with {p} items Bartlett's test has p(p − 1)/2 = {p * (p - 1) // 2} df, not {d}."})
        elif not p and not any(q * (q - 1) // 2 == d for q in range(2, 300)):
            out.append({"s": kmo["df"][1], "kind": "test", "decision": False,
                        "note": f"KMO and Bartlett's Test: Bartlett's test has p(p − 1)/2 df for p items, and {d} is no such number."})
        elif "chi" in kmo and "sig" in kmo and d > 0:
            note, judged = _lg_p("", kmo["chi"][0], d, kmo["sig"][0], "KMO and Bartlett's Test, Bartlett's test")
            n_judged += 1 if judged else 0
            if note:
                out.append(dict(note, s=kmo["sig"][1]))
    pca = bool(got["comm"]) and all(abs((_num_cell(i) or 0) - 1.0) < 1e-9 for _l, i, _e, _s in got["comm"])
    ext_tot, rot_tot, orth, tot_bad = [], [], False, set()
    if got["tve"] and p:
        rows = sorted(got["tve"], key=lambda r: r[0])
        cum = {0: [0.0, 0.0], 3: [0.0, 0.0], 6: [0.0, 0.0]}
        tot_init = 0.0
        prev = None
        noted = False
        for i_, (comp, vals, s_) in enumerate(rows):
            groups = [0, 3] + ([6] if len(vals) >= 9 else [])
            for g in groups:
                if g + 2 >= len(vals):
                    continue
                t_t, pc_t, cu_t = vals[g], vals[g + 1], vals[g + 2]
                t_, pc_, cu_ = _num_cell(t_t), _num_cell(pc_t), _num_cell(cu_t)
                want = 100.0 * t_ / p
                if abs(pc_ - want) > 100.0 * h_of(t_t) / p + h_of(pc_t) + 1e-9 and not noted:
                    what = {0: "Initial Eigenvalues", 3: "Extraction Sums of Squared Loadings", 6: "Rotation Sums of Squared Loadings"}[g]
                    chain_ok = abs(cu_ - (cum[g][0] + pc_)) <= cum[g][1] + h_of(pc_t) + h_of(cu_t) + 1e-9
                    if chain_ok:                        # (its % agrees with the Cumulative %: the Total is the figure in question)
                        tot_bad.add((g, comp - 1))
                        out.append({"s": s_, "kind": "table", "decision": False,
                                    "note": f"Total Variance Explained, {what}, component {comp}: its Total ({t_t}) doesn’t fit its % of Variance ({pc_t}), "
                                            f"which with {p} items is a Total of {_like(pc_ * p / 100.0, t_t)}."})
                    else:
                        out.append({"s": s_, "kind": "table", "decision": False,
                                    "note": f"Total Variance Explained, {what}, component {comp}: its % of Variance is its Total over the {p} items × 100, "
                                            f"{t_t} / {p} × 100 ≈ {_like(want, pc_t)}, not {pc_t}."})
                    noted = True
                cum[g][0] += pc_
                cum[g][1] += h_of(pc_t)
                if abs(cu_ - cum[g][0]) > cum[g][1] + h_of(cu_t) + 1e-9 and not noted:
                    out.append({"s": s_, "kind": "table", "decision": False,
                                "note": f"Total Variance Explained, component {comp}: its Cumulative % is the running sum of the % of Variance, "
                                        f"{_like(cum[g][0], cu_t)}, not {cu_t}."})
                    noted = True
                if abs(cu_ - cum[g][0]) <= cum[g][1] + h_of(cu_t) + 1e-9:
                    cum[g] = [cu_, h_of(cu_t)]       # (carry on from the figure as written)
                if g == 3:
                    ext_tot.append((t_, t_t))
                    if pca and abs(t_ - _num_cell(vals[0])) > h_of(t_t) + h_of(vals[0]) + 1e-9 and not noted:
                        out.append({"s": s_, "kind": "table", "decision": False,
                                    "note": f"Total Variance Explained, component {comp}: with principal components, the Extraction total is the Initial "
                                            f"eigenvalue, {vals[0]}, not {t_t}."})
                        noted = True
                if g == 6:
                    rot_tot.append((t_, t_t))
                    orth = True
            e0 = _num_cell(vals[0])
            tot_init += e0
            if prev is not None and e0 > prev[0] + h_of(prev[1]) + h_of(vals[0]) + 1e-9 and not noted:
                out.append({"s": s_, "kind": "table", "decision": False,
                            "note": f"Total Variance Explained: the eigenvalues fall from one component to the next, but component {comp}'s ({vals[0]}) "
                                    f"is above the one before ({prev[1]})."})
                noted = True
            prev = (e0, vals[0])
        if len(rows) == p and not noted and abs(tot_init - p) > p * h_of(rows[0][1][0]) + 1e-9:
            out.append({"s": rows[-1][2], "kind": "table", "decision": False,
                        "note": f"Total Variance Explained: the {p} eigenvalues of {p} items add up to {p}, but these add up to {tot_init:.3f}."})
        if orth and ext_tot and rot_tot and len(ext_tot) == len(rot_tot) and not noted:
            se_, sr_ = sum(t for t, _ in ext_tot), sum(t for t, _ in rot_tot)
            if abs(se_ - sr_) > sum(h_of(t) for _, t in ext_tot) + sum(h_of(t) for _, t in rot_tot) + 1e-9:
                out.append({"s": rows[0][2], "kind": "table", "decision": False,
                            "note": f"Total Variance Explained: a rotation shares out the same variance, so the Rotation totals add up to the Extraction’s "
                                    f"({se_:.3f}), not {sr_:.3f}."})
    # Communalities against the loadings, and each column against its total
    comm = {l.lower(): (l, i, e, s_) for l, i, e, s_ in got["comm"]}
    comm_noted = set()
    for key, mat, tots, name in (("mat", got["mat"], ext_tot, "Component Matrix"), ("rot", got["rot"], rot_tot if orth else [], "Rotated Component Matrix")):
        if not mat:
            continue
        k = len(mat[0][1])
        if any(len(v) != k for _l, v, _s in mat):
            continue
        bad_rows = []
        for lab, vals, s_ in mat:
            c = comm.get(lab.lower())
            if not c or lab.lower() in comm_noted:
                continue
            e_t, e = c[2], _num_cell(c[2])
            lo_ = sum(max(abs(_num_cell(x)) - h_of(x), 0) ** 2 for x in vals)
            hi_ = sum((abs(_num_cell(x)) + h_of(x)) ** 2 for x in vals)
            if e + h_of(e_t) < lo_ - 1e-9 or e - h_of(e_t) > hi_ + 1e-9:
                bad_rows.append((lab, vals, s_, c))
        bad_cols = []
        if len(tots) == k and len(mat) == p:
            for j in range(k):
                lo_ = sum(max(abs(_num_cell(v[j])) - h_of(v[j]), 0) ** 2 for _l, v, _s in mat)
                hi_ = sum((abs(_num_cell(v[j])) + h_of(v[j])) ** 2 for _l, v, _s in mat)
                t_, t_t = tots[j]
                if t_ + h_of(t_t) < lo_ - 1e-9 or t_ - h_of(t_t) > hi_ + 1e-9:
                    bad_cols.append(j)
        bad_cols = [j for j in bad_cols if (3 if key == "mat" else 6, j) not in tot_bad]     # (a Total already said is not said again)
        if len(bad_rows) == 1 and len(bad_cols) == 1:      # (one loading in question: name it)
            lab, vals, s_, c = bad_rows[0]
            j = bad_cols[0]
            rest = sum(_num_cell(x) ** 2 for q, x in enumerate(vals) if q != j)
            need = _num_cell(c[2]) - rest
            fit = f"; about {_like(math.copysign(math.sqrt(need), _num_cell(vals[j])), vals[j])} would fit both" if need > 0 else ""
            out.append({"s": s_, "kind": "table", "decision": False,
                        "note": f"{name}, {lab}: its loading on component {j + 1} ({vals[j]}) fits neither its communality ({c[2]}) nor component "
                                f"{j + 1}'s total ({tots[j][1]}){fit}."})
            continue
        if len(bad_rows) == 1 and not bad_cols and key == "mat":
            lab, vals, s_, c = bad_rows[0]
            want = sum(_num_cell(x) ** 2 for x in vals)
            out.append({"s": c[3], "kind": "table", "decision": False,
                        "note": f"Communalities, {lab}: its Extraction is the sum of its squared loadings, {_like(want, c[2])}, not {c[2]}."})
            comm_noted.add(lab.lower())
            continue
        for lab, vals, s_, c in bad_rows[:2]:
            want = sum(_num_cell(x) ** 2 for x in vals)
            out.append({"s": s_, "kind": "table", "decision": False,
                        "note": f"{name}, {lab}: its squared loadings add up to {_like(want, c[2])}, but its communality is {c[2]}."})
        for j in bad_cols[:2]:
            want = sum(_num_cell(v[j]) ** 2 for _l, v, _s in mat)
            out.append({"s": mat[0][2], "kind": "table", "decision": False,
                        "note": f"{name}: component {j + 1}'s squared loadings add up to {want:.3f}, but its total is {tots[j][1]}."})
    if (pca and got["comm"] and ext_tot and len(got["comm"]) == p and not any(n_["note"].startswith("Communalities") for n_ in out)
            and not any(g == 3 for g, _c in tot_bad)):
        se_ = sum(_num_cell(e) for _l, _i, e, _s in got["comm"])
        st_ = sum(t for t, _ in ext_tot)
        slack = sum(h_of(e) for _l, _i, e, _s in got["comm"]) + sum(h_of(t) for _, t in ext_tot)
        if abs(se_ - st_) > slack + 1e-9 and not any("loading" in n_["note"] for n_ in out):
            out.append({"s": got["comm"][0][3], "kind": "table", "decision": False,
                        "note": f"Communalities: the Extraction communalities add up to {se_:.3f}, but the components kept explain {st_:.3f} in Total "
                                f"Variance Explained; the two are the same variance."})
    if counted is not None:
        counted[0] += n_judged
    return out



# 6.6, round 151 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): Crosstabs' Symmetric Measures, read with its
# Chi-Square Tests: the Contingency Coefficient = |Phi| / √(1 + Phi²); Cramér's V = |Phi| / √(min(rows, columns) − 1), so (Phi / V)²
# is a whole number and V is never above |Phi|; |Phi| = √(χ² / N) with the Pearson Chi-Square and N of Valid Cases; their
# Approximate Significance is the Pearson Chi-Square's.
_SY_ROW = {"phi": re.compile(r"^phi$", re.I), "v": re.compile(r"^cram[eé]r['’]?s\s+v$", re.I),
           "c": re.compile(r"^contingency\s+coefficient$", re.I), "n": re.compile(r"^n\s+of\s+valid\s+cases$", re.I)}


def symm_check(text):
    out = []
    lines = str(text or "").split("\n")
    pearson, sym, vcol, pcol, s_line = None, {}, None, None, ""
    mode = None
    for raw in lines:
        cells = _cells_keep(raw)
        filled = [(c or "").strip() for c in cells if (c or "").strip()]
        if not filled:
            continue
        low = [(c or "").strip().lower() for c in cells]
        if len(filled) == 1 and re.match(r"^symmetric\s+measures(?:[a-e](?:,[a-e])*)?$", filled[0], re.I):
            mode, vcol, pcol = "sym", None, None
            continue
        if len(filled) == 1 and re.match(r"^chi-?\s?square\s+tests(?:[a-e](?:,[a-e])*)?$", filled[0], re.I):
            mode = "chi"
            continue
        if mode == "chi" and any(_SC_ROW.match(c) and c.lower().startswith("pearson") for c in filled):
            nums = [c for c in filled[1:] if _num_cell(re.sub(r"(?<=\d)[a-e]$", "", c)) is not None or _p_cell(c) is not None]
            if len(nums) >= 3:
                pearson = (re.sub(r"(?<=\d)[a-e]$", "", nums[0]), nums[1], nums[2])
            continue
        if mode == "chi" and filled[0].lower() == "n of valid cases":
            nv = [c for c in filled[1:] if re.fullmatch(r"\d+", c)]
            if nv:
                sym.setdefault("_n_chi", nv[0])
            continue
        if mode == "sym":
            if "value" in low:
                vcol = low.index("value")
                pcol = next((j for j, c in enumerate(low) if "sig" in c), None)
                continue
            if vcol is None:
                continue
            lab = next((k for k, rx in _SY_ROW.items() for c in filled[:3] if rx.match(c)), None)
            if lab is None:
                if len(filled) == 1:
                    mode = None
                continue
            v_t = (cells[vcol] or "").strip() if vcol < len(cells) else ""
            p_t = (cells[pcol] or "").strip() if pcol is not None and pcol < len(cells) else ""
            if lab == "n":
                sym["n"] = v_t
            elif _num_cell(v_t) is not None:
                sym[lab] = (v_t, p_t, raw.strip()[:500])
    if "phi" not in sym:
        return out
    h_of = lambda t: 0.5 * 10 ** -_dec(str(t).lstrip("-−"))
    ph_t, ph_p, ph_s = sym["phi"]
    ph = abs(_num_cell(ph_t))
    hp = h_of(ph_t)
    n_t = sym.get("n") or sym.get("_n_chi")
    if pearson and n_t and re.fullmatch(r"\d+", n_t):
        x_t, _df, xp_t = pearson
        x, n = _num_cell(x_t), int(n_t)
        if x is not None and n > 0:
            lo_, hi_ = math.sqrt(max(x - h_of(x_t), 0) / n), math.sqrt((x + h_of(x_t)) / n)
            if ph + hp < lo_ - 1e-9 or ph - hp > hi_ + 1e-9:
                out.append({"s": ph_s, "kind": "test", "decision": False,
                            "note": f"Symmetric Measures: |Phi| is √(χ² / N), √({x_t} / {n}) = {_like(math.sqrt(x / n), ph_t)}, not {ph_t}."})
                return out
    if "c" in sym:
        c_t, _cp, c_s = sym["c"]
        c = _num_cell(c_t)
        g = lambda x: x / math.sqrt(1 + x * x)
        if c + h_of(c_t) < g(max(ph - hp, 0)) - 1e-9 or c - h_of(c_t) > g(ph + hp) + 1e-9:
            out.append({"s": c_s, "kind": "test", "decision": False,
                        "note": f"Symmetric Measures: the Contingency Coefficient is |Phi| / √(1 + Phi²), {ph_t} gives {_like(g(ph), c_t)}, not {c_t}."})
            return out
    if "v" in sym:
        v_t, _vp, v_s = sym["v"]
        v = _num_cell(v_t)
        if v - h_of(v_t) > ph + hp + 1e-9:
            out.append({"s": v_s, "kind": "test", "decision": False,
                        "note": f"Symmetric Measures: Cramér's V is |Phi| divided by √(the smaller of rows and columns, less 1), so it is never above "
                                f"|Phi| ({ph_t}), not {v_t}."})
            return out
        if v - h_of(v_t) > 0:
            k_lo = ((max(ph - hp, 0)) / (v + h_of(v_t))) ** 2
            k_hi = ((ph + hp) / (v - h_of(v_t))) ** 2
            dfp = int(pearson[1]) if pearson and re.fullmatch(r"\d+", pearson[1]) else None
            ks = [k for k in range(1, 50) if dfp is None or (dfp % k == 0 and k * k <= dfp)]   # (min(r, c) − 1 divides df, its square at most df)
            if not any(k_lo - 1e-9 <= k <= k_hi + 1e-9 for k in ks):
                out.append({"s": v_s, "kind": "test", "decision": False,
                            "note": f"Symmetric Measures: Cramér's V is |Phi| / √1, √2, √3… (the smaller of rows and columns, less 1); with Phi = {ph_t}, "
                                    f"V = {v_t} fits none of them."})
                return out
    if pearson and n_t and re.fullmatch(r"\d+", n_t):
        pc = _p_cell(pearson[2])
        xp_t = pearson[2]
        if pc is not None:
            for key, name in (("phi", "Phi"), ("v", "Cramér's V"), ("c", "the Contingency Coefficient")):
                if key in sym and sym[key][1]:
                    sp = _p_cell(sym[key][1])
                    if sp is not None and abs(sp[1] - pc[1]) > 0.5 * 10 ** -sp[2] + 0.5 * 10 ** -pc[2] + 1e-9 and sp[0] == "=" and pc[0] == "=":
                        out.append({"s": sym[key][2], "kind": "test", "decision": False,
                                    "note": f"Symmetric Measures: the Approximate Significance of {name} is the Pearson Chi-Square's ({xp_t}), not {sym[key][1]}."})
                        break
    return out



# 6.6, rounds 152 to 154 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): the words of a results chapter against the SPSS
# output beside them. A figure a sentence reports (Cronbach's alpha, KMO, Nagelkerke's or Cox & Snell's R², R², the share
# correctly classified, χ², t, F, r or rho, U, H, Z, Cramér's V, phi, an odds ratio) is looked for among the same statistic's
# figures in the table the sentence names ("Table 4.3"), else in the tables just above and below its paragraph; it must be one
# of them, rounded or cut to the sentence's decimals (a t or Z may drop its minus). Only where those tables hold that statistic at
# all: a sentence about a table not pasted, a pilot study, or another study (a citation in it) is left alone, and so is α = .05.
_TS_CITE = re.compile(r"\((?:[^()]{0,80}?,\s*)?(?:19|20)\d{2}[a-z]?(?:,\s*p\.?\s*\d+)?\)|[A-Z][A-Za-z'’\-]+(?:\s+(?:et\s+al\.|and\s+[A-Z][A-Za-z'’\-]+))?\s*\((?:19|20)\d{2}[a-z]?\)"
                      r"|\[\s*\d+(?:\s*[,–-]\s*\d+)*\s*\]")
_TS_OTHER = re.compile(r"\b(?:pilot|pre-?test(?:ing)?|previous(?:ly)?|earlier\s+stud|other\s+stud|prior\s+stud|literature|reported\s+by|according\s+to)\b", re.I)
_TS_NUM = r"(?P<v>[-−–]?\d*\.?\d+)"
_TS_TALK = [
    ("alpha", "Cronbach’s alpha", re.compile(r"cronbach['’]?s\s+(?:alpha|α)(?:\s+(?:coefficient|value|score))?\s*(?:\(\s*α\s*\)\s*)?"
                                              r"(?:\s*(?:for|of|in|on)\s+(?:the\s+|this\s+|each\s+)?(?:[A-Za-z-]+\s+){0,3}?[A-Za-z-]+\s+(?=was|is|were|stood|yielded|=))?"
                                              r"(?:of|=|was|is|were|:|stood\s+at|yielded)?\s*(?:=\s*)?(?:a\s+value\s+of\s+)?(?:α\s*=\s*)?" + _TS_NUM, re.I)),
    ("alpha", "Cronbach’s alpha", re.compile(r"(?<![A-Za-z])α\s*=\s*" + _TS_NUM)),
    ("kmo", "KMO", re.compile(r"(?:\bKMO\b|Kaiser[-–\s]*Meyer[-–\s]*Olkin)(?:\s+(?:measure|value|statistic|test|index))?(?:\s+of\s+sampling\s+adequacy)?\s*(?:\(\s*KMO\s*\)\s*)?(?:of|=|was|is|:)?\s*(?:=\s*)?" + _TS_NUM, re.I)),
    ("nagel", "Nagelkerke’s R²", re.compile(r"Nagelkerke(?:['’]s)?\s+(?:pseudo[-\s]*)?R\s*(?:²|2|\^2|-?\s*square[d]?)?\s*(?:of|=|was|is|:)?\s*(?:=\s*)?" + _TS_NUM, re.I)),
    ("nagel%", "Nagelkerke’s R²", re.compile(r"(?P<v>\d{1,2}(?:\.\d+)?)\s*%\s*\(\s*Nagelkerke", re.I)),
    ("cox", "Cox and Snell’s R²", re.compile(r"Cox\s*(?:&|and)\s*Snell(?:['’]s)?\s+(?:pseudo[-\s]*)?R\s*(?:²|2|\^2|-?\s*square[d]?)?\s*(?:of|=|was|is|:)?\s*(?:=\s*)?" + _TS_NUM, re.I)),
    ("cox%", "Cox and Snell’s R²", re.compile(r"(?P<v>\d{1,2}(?:\.\d+)?)\s*%\s*\(\s*Cox", re.I)),
    ("overall", "the share correctly classified", re.compile(r"correctly\s+(?:classified|predicted)\s+(?:about\s+|approximately\s+|an?\s+overall\s+)?(?P<v>\d{1,3}(?:\.\d+)?)\s*%", re.I)),
    ("overall", "the share correctly classified", re.compile(r"(?P<v>\d{1,3}(?:\.\d+)?)\s*%\s+of\s+(?:the\s+|all\s+)?(?:cases|respondents|participants|students)\s+(?:were\s+)?correctly\s+(?:classified|predicted)", re.I)),
    ("overall", "the share correctly classified", re.compile(r"overall\s+(?:classification|prediction|predictive)\s+(?:accuracy|rate)\s+(?:of|was|=)\s*(?P<v>\d{1,3}(?:\.\d+)?)\s*%", re.I)),
    ("r2", "R²", re.compile(r"(?<![A-Za-z&])(?<!Snell\s)(?<!kerke\s)(?:adjusted\s+)?R\s*(?:²|\^2|2(?=\s*=))\s*=\s*" + _TS_NUM)),
    ("chi", "χ²", re.compile(r"(?:χ|chi)\s*(?:²|2|-?\s*square[d]?)\s*(?:\(\s*(?P<df>\d+)\s*(?:,\s*N\s*=\s*\d[\d,]*)?\s*\))?\s*(?:value\s+)?(?:=|of|was)\s*" + _TS_NUM, re.I)),
    ("t", "t", re.compile(r"(?<![A-Za-z])t\s*(?:\(\s*(?P<df>\d+(?:\.\d+)?)\s*\))?\s*=\s*" + _TS_NUM)),
    ("F", "F", re.compile(r"(?<![A-Za-z])F\s*(?:\(\s*(?P<d1>\d+)\s*,\s*(?P<d2>\d+(?:\.\d+)?)\s*\))?\s*=\s*" + _TS_NUM)),
    ("r", "r", re.compile(r"(?<![A-Za-z])r\s*(?:\(\s*\d+\s*\))?\s*=\s*" + _TS_NUM)),
    ("r", "rho", re.compile(r"(?:\brho\b|ρ|(?<![A-Za-z])r\s?s(?![A-Za-z]))\s*(?:\(\s*\d+\s*\))?\s*=\s*" + _TS_NUM, re.I)),
    ("U", "U", re.compile(r"(?<![A-Za-z])U\s*=\s*" + _TS_NUM)),
    ("H", "H", re.compile(r"(?<![A-Za-z])H\s*(?:\(\s*(?P<df>\d+)\s*(?:,\s*N\s*=\s*\d+\s*)?\))?\s*=\s*" + _TS_NUM)),
    ("beta", "β", re.compile(r"(?:β|\b[Bb]eta\b)\s*=\s*" + _TS_NUM)),                # 6.6, round 153: a coefficient
    ("b", "B", re.compile(r"(?<!Exp\()(?<![A-Za-z])[Bb]\s*=\s*" + _TS_NUM)),
    ("p", "p", re.compile(r"(?<![A-Za-z])[pP]\s*(?:-?\s*value\s*)?(?:=|of|was|is)\s*(?P<v>0?\.\d+|0|1(?:\.0+)?)(?![\d.]*\d)")),   # 6.6, round 154
    ("Z", "Z", re.compile(r"(?<![A-Za-z])[zZ]\s*=\s*" + _TS_NUM)),
    ("V", "Cramér’s V", re.compile(r"Cram[eé]r['’]?s\s+V\s*(?:of|=|was|is|:)?\s*(?:=\s*)?" + _TS_NUM, re.I)),
    ("phi", "phi", re.compile(r"(?:\bphi\b|φ)\s*(?:coefficient\s*)?(?:of|=|was|is|:)?\s*(?:=\s*)?" + _TS_NUM, re.I)),
    ("or", "the odds ratio", re.compile(r"(?:(?<![A-Za-z])(?:[ac]?OR|AOR|COR)|odds\s+ratio)\s*(?:of|=|was|is|:)?\s*(?:=\s*)?(?P<v>\d+(?:\.\d+)?)(?!\s*%)")),
]
_TS_LABEL = {"pearson chi-square": "chi", "continuity correction": "chi", "likelihood ratio": "chi", "approx. chi-square": "chi", "chi-square": "chi",
             "kruskal-wallis h": "H", "mann-whitney u": "U", "z": "Z", "phi": "phi", "cramer's v": "V", "cramér's v": "V",
             "pearson correlation": "r", "correlation coefficient": "r", "overall percentage": "overall"}
_TS_HEAD = {"t": "t", "f": "F", "chi-square": "chi", "r square": "r2", "adjusted r square": "r2", "nagelkerke r square": "nagel",
            "cox & snell r square": "cox", "cronbach's alpha": "alpha", "cronbach's alpha based on standardized items": "alpha", "exp(b)": "or",
            "score": "chi", "beta": "beta", "b": "b", "correlation": "r"}


def _ts_kind_line(raw):
    cells = _cells_keep(raw)
    filled = [(c or "").strip() for c in cells if (c or "").strip()]
    if len(filled) >= 2 and ("\t" in raw or "|" in raw):
        return "T"
    if not filled:
        return "B"
    one = filled[0]
    if re.match(r"^table\s+\d", one, re.I):               # ("Table 4.3: Knowledge…" is a caption; "Table 4.3 shows…" is prose)
        verb = re.search(r"\b(?:shows?|presents?|indicates?|reveals?|summari[sz]es?|displays?|gives?|illustrates?|depicts?|reports?|contains?)\b", one, re.I)
        return "C" if len(one) < 140 and not verb else "P"
    if re.match(r"^(?:[a-e]\.\s|dependent\s+variable|extraction\s+method|rotation\s+method|\*\*?\.\s)", one, re.I) or (len(one) < 70 and not re.search(r"[.!?]$", one)):
        return "C"
    return "P"


_TS_SIG = re.compile(r"^(?:sig\.?(?:\s*\((?:1|2)-tailed\))?|p(?:-?\s*value)?|asymp(?:\.|totic)\s+sig(?:\.|nificance)(?:\s*\((?:1|2)-(?:tailed|sided)\))?"
                     r"|approx(?:\.|imate)\s+sig(?:\.|nificance)|exact\s+sig\.?(?:\s*\((?:1|2)-(?:tailed|sided)\))?)$", re.I)
_TS_STAT_BEFORE = re.compile(r"(?:(?<![A-Za-z])[tFHUzZrWBb]\s*(?:\([^)]*\))?\s*=|χ|chi\b|chi-|rho|β|cram[eé]r|\bOR\s*=|AOR|odds\s+ratio)", re.I)


def _ts_p_fits(v_t, cands):
    """Whether a sentence's "p = v" is one of the table's Sig. cells ("<.001" holds any p = .000)."""
    v, d = float(v_t), _dec(v_t)
    for c_t in cands:
        pc = _p_cell(c_t)
        if pc is None:
            continue
        rel, c, dp = pc
        if rel == "=" and abs(v - c) <= 0.5 * max(10 ** -d, 10 ** -dp) + 1e-9:
            return True
        if rel in ("<", "≤") and v < c + 1e-12:
            return True
    return False


def _ts_head_kind(h):
    for cand in (h, re.sub(r"(?<=[a-z)])[a-e](?:,[a-e])*$", "", h)):
        if cand in _TS_HEAD:
            return _TS_HEAD[cand]
    if re.match(r"^(?:[ac]?or|aor|cor|adjusted\s+odds\s+ratio|crude\s+odds\s+ratio|odds\s+ratio)\b", h):
        return "or"
    return None


_TS_COUNT = re.compile(r"^\d[\d,]*(?:\.\d+)?\s*\(\s*\d+(?:\.\d+)?\s*%?\s*\)$")


def _ts_values(rows):
    """{kind: [value texts]} of a cluster of table lines; and (6.6, round 153) for t, F, χ² and H, {"df:<kind>": [(value, its df)]}:
    a t's df on its own row, an F's own df with the Error's (Within Groups, Residual), a χ²'s df column, or a transposed table's
    df row (Test Statistics; KMO's Bartlett test) by column."""
    vals, heads, last_head = {}, None, False
    pend_f, err_df, trans = [], None, {}
    norm = lambda c: re.sub(r"(?<=[a-z])[a-e](?:,[a-e])*$", "", c) if re.sub(r"(?<=[a-z])[a-e](?:,[a-e])*$", "", c) in _TS_LABEL else c

    def flush():
        nonlocal pend_f, err_df, trans
        if err_df is not None:
            for v, d1 in pend_f:
                vals.setdefault("df:F", []).append((v, (d1, err_df)))
        for kk in ("chi", "H"):
            for j, v in trans.get(kk, {}).items():
                d = trans.get("df", {}).get(j)
                if d is not None and re.fullmatch(r"\d+", d):
                    vals.setdefault("df:" + kk, []).append((v, d))
        pend_f, err_df, trans = [], None, {}
    for raw in rows:
        cells = _cells_keep(raw)
        filled = [(c or "").strip() for c in cells if (c or "").strip()]
        if len(filled) < 2:
            continue                                  # (a title, a sub-heading such as "Age (years)", a footnote: the headings stand)
        low = [(c or "").strip().lower() for c in cells]
        txt = lambda i: re.sub(r"(?<=\d)[a-e]$", "", (cells[i] or "").strip().rstrip("*")).replace("−", "-")
        nums = [i for i, c in enumerate(cells) if _num_cell(txt(i)) is not None]
        data = nums or any(_OR_CELL.match(c) or _p_cell(c) or _TS_COUNT.match(c) for c in filled)   # (6.6, round 163: "34 (57.6)" is data)
        if not data:
            rest = [c for c in low[1:]]
            named = [c for c in rest if c and c not in ("ref", "ref.", "reference", "referent", "-", "—")]
            if len(named) >= 2 or (named and len(rest) <= 2):      # (a heading; "Male | Ref | | Ref" is a reference row)
                if not last_head:
                    flush()
                new = {i: c for i, c in enumerate(low) if c}           # (a heading: most of its cells named)
                heads = {**(heads or {}), **new} if last_head else new
                last_head = True
            continue                                  # (a reference row "18-20 | Ref", a sub-heading: the headings stand)
        last_head = False
        dfcol = next((i for i, h in (heads or {}).items() if h == "df"), None)
        if any(c == "df" for c in low[:2]) and not any(_TS_SIG.match(c) for c in low[:2] if c):   # (a transposed table's df row; KMO's "df")
            trans.setdefault("df", {}).update({i: txt(i) for i in nums})
            continue
        if "kaiser-meyer-olkin" in raw.lower():
            vals.setdefault("kmo", []).append(txt(nums[-1]))
            continue
        if heads:                                     # 6.6, round 154: each Sig. cell, "<.001" too
            for i, c in enumerate(cells):
                if _TS_SIG.match(heads.get(i, "")) and _p_cell((c or "").strip().rstrip("*")):
                    vals.setdefault("p", []).append(re.sub(r"(?<=\d)[a-e]$", "", (c or "").strip().rstrip("*")))
        if any(_TS_SIG.match(re.sub(r"(?<=[a-z.)])[a-e]$", "", c)) for c in low[:2] if c):
            vals.setdefault("p", []).extend(re.sub(r"(?<=\d)[a-e]$", "", (c or "").strip().rstrip("*")) for c in cells[1:] if _p_cell((c or "").strip().rstrip("*")))
            continue
        k = next((_TS_LABEL[norm(c)] for c in low[:3] if c and norm(c) in _TS_LABEL), None)
        if k == "r":
            vals.setdefault("r", []).extend(txt(i) for i in nums if txt(i) not in _CR_ONES)
            continue
        if k == "overall":
            vals.setdefault("overall", []).append(txt(nums[-1]))
            continue
        if k:
            vcol = next((i for i, h in (heads or {}).items() if h == "value"), None)
            got = [txt(vcol)] if vcol is not None and vcol in nums else [txt(i) for i in nums]
            vals.setdefault(k, []).extend(got)
            if k == "H":
                vals.setdefault("chi", []).extend(got)
            if k in ("chi", "H"):
                if vcol is not None and dfcol is not None and dfcol in nums:
                    vals.setdefault("df:" + k, []).append((got[0], txt(dfcol)))
                else:
                    for kk in ((k, "chi") if k == "H" else (k,)):
                        trans.setdefault(kk, {}).update({i: txt(i) for i in nums})
            continue
        if dfcol is not None and dfcol in nums and any(c in ("within groups", "residual", "error") for c in low[:3]):
            err_df = txt(dfcol)
        if heads:
            for i in nums:
                kk = _ts_head_kind(heads.get(i, ""))
                if kk:
                    vals.setdefault(kk, []).append(txt(i))
                    if dfcol is not None and dfcol in nums and kk in ("t", "chi"):
                        vals.setdefault("df:" + kk, []).append((txt(i), txt(dfcol)))
                    if dfcol is not None and dfcol in nums and kk == "F" and dfcol < i:
                        pend_f.append((txt(i), txt(dfcol)))
        for c in filled:                              # (a typed table of ratios: "2.31 (1.20–4.45)")
            m = _OR_CELL.match(c)
            if m:
                vals.setdefault("or", []).append(m.group("e"))
    flush()
    return vals


def _ts_match(v_t, cands, kind):
    """Whether a sentence's figure is one of the table's, rounded or cut to the sentence's decimals."""
    v = float(v_t.replace("−", "-").replace("–", "-"))
    d = _dec(v_t.lstrip("-−–"))
    unit = 10 ** -d
    for c_t in cands:
        c = _num_cell(c_t)
        if c is None:
            continue
        for a, b in ((v, c), (abs(v), abs(c))) if kind in ("t", "Z") else ((v, c),):
            if abs(a - b) <= 0.5 * max(unit, 10 ** -_dec(c_t.lstrip("-"))) + 1e-9:
                return True                           # (the table's figure rounded to the text's decimals)
            if _dec(c_t.lstrip("-")) > d and (a >= 0) == (b >= 0) and int(abs(b) * 10 ** d + 1e-9) == round(abs(a) * 10 ** d):
                return True                           # (cut, not rounded: 58.71 for 58.717)
    return False


def text_stat_check(text):
    out, said = [], set()
    lines = str(text or "").split("\n")
    kinds = [_ts_kind_line(raw) for raw in lines]
    # clusters of table lines (with their titles, captions and footnotes) and paragraphs of prose
    blocks, cur = [], None
    for i, k in enumerate(kinds):
        typ = "table" if k in ("T", "C") else "prose" if k == "P" else None
        if typ is None:
            continue
        if cur and cur["type"] == typ:
            cur["idx"].append(i)
        else:
            cur = {"type": typ, "idx": [i]}
            blocks.append(cur)
    tables = [b for b in blocks if b["type"] == "table" and any(kinds[i] == "T" for i in b["idx"])]
    if not tables:
        return out
    for b in tables:
        b["vals"] = _ts_values([lines[i] for i in b["idx"]])
        cap = next((m.group(1) for i in b["idx"] for m in [_TT_CAPTION.search(lines[i])] if m and kinds[i] == "C"), None)
        b["no"] = cap
    for bi, b in enumerate(blocks):
        if b["type"] != "prose":
            continue
        near = [x for x in (blocks[bi - 1] if bi > 0 else None, blocks[bi + 1] if bi + 1 < len(blocks) else None) if x is not None and "vals" in x]
        para = " ".join(lines[i].strip() for i in b["idx"])
        for sent in sentences(para):
            if _TS_CITE.search(sent) or _TS_OTHER.search(sent) or _NC_SENT.search(sent):
                continue                              # (6.6, round 166: a question, a hypothesis, an expectation is not a finding)
            named = {m.group(1) for m in _TT_CAPTION.finditer(sent)}
            if named:
                pool = [t for t in tables if t["no"] in named]
                if not pool:
                    continue                          # (a table that isn't pasted: nothing to hold the sentence to)
            else:
                pool = near
            if not pool:
                continue
            taken = []
            for kind, name, rx in _TS_TALK:
                for m in rx.finditer(sent):
                    if any(a <= m.start("v") < z for a, z in taken):
                        continue
                    taken.append((m.start("v"), m.end("v")))
                    base = kind.rstrip("%")
                    v_t = m.group("v")
                    if kind == "p":
                        before = sent[max(0, m.start() - 60):m.start()]
                        before = re.split(r"[;.](?:\s|$)", before)[-1]
                        cands = [c for t in pool for c in t["vals"].get("p", [])]
                        if not cands or _TS_STAT_BEFORE.search(before) or _ts_p_fits(v_t, cands):
                            continue
                        if abs(float(v_t) - round(float(v_t), 3)) < 1e-12 and float(v_t) in (0.05, 0.01, 0.001, 0.1) and re.search(
                                r"\b(?:level|set|threshold|cut-?\s?off|consider|deem|accept|alpha|criterion)", sent, re.I):
                            continue                  # ("significance was set at p = 0.05": the level, not a result)
                        where = f"Table {next(iter(named))}" if named and len(named) == 1 else "the table beside it"
                        uniq = list(dict.fromkeys(cands))
                        shown = " or ".join(uniq[:3]) + (", among others" if len(uniq) > 3 else "")
                        note = f"The text gives p = {v_t}, but {where} gives Sig. {shown}."
                        if (sent, note) not in said:
                            said.add((sent, note))
                            out.append({"s": sent[:500], "note": note, "kind": "value", "decision": False})
                        continue
                    if _TH_AFTER.match(sent[m.end("v"):]) or _TH_BEFORE.search(sent[max(0, m.start("v") - 40):m.start("v")]):
                        continue                      # (6.6, round 166: "an alpha of 0.70 or more was considered acceptable": a threshold)
                    if kind == "alpha" and "cronbach" not in sent.lower() and v_t.lstrip("0") in (".05", ".01", ".001", ".1", ".10"):
                        continue                      # (α = .05: the significance level)
                    if kind == "alpha" and "cronbach" not in sent.lower() and not re.search(r"reliab|internal\s+consistency", sent, re.I):
                        continue
                    if kind == "r" and name == "r" and re.search(r"(?<![A-Za-z])r\s*s\s*=", sent):
                        continue
                    cands = [c for t in pool for c in t["vals"].get(base, [])]
                    if not cands:
                        continue
                    if kind.endswith("%"):
                        cands = [f"{100 * _num_cell(c):.{max(_dec(c) - 2, 0)}f}" for c in cands if _num_cell(c) is not None]
                    if kind == "r2" and re.search(r"(?:nagelkerke|cox|snell)\W{0,20}$", sent[:m.start()], re.I):
                        continue
                    if _ts_match(v_t, cands, base):
                        pairs = [(v, d) for t in pool for v, d in t["vals"].get("df:" + base, []) if _ts_match(v_t, [v], base)]
                        gd = m.groupdict()
                        said_df = (gd.get("d1"), gd.get("d2")) if base == "F" else gd.get("df")
                        if pairs and said_df and said_df[0] if base == "F" else pairs and said_df:
                            def same(d):
                                if base == "F":
                                    return int(said_df[0]) == int(float(d[0])) and _ts_match(said_df[1], [d[1]], "df")
                                return _ts_match(said_df, [d], "df")
                            if not any(same(d) for _v, d in pairs):
                                dfs = list(dict.fromkeys(f"{d[0]} and {d[1]}" if base == "F" else d for _v, d in pairs))
                                where = f"Table {next(iter(named))}" if named and len(named) == 1 else "the table beside it"
                                shown_ = f"{base}({', '.join(said_df) if base == 'F' else said_df})" if base != "chi" else f"χ²({said_df})"
                                note = f"The text gives {shown_}, but {where} has this {name} on {' or '.join(dfs[:2])} df."
                                if (sent, note) not in said:
                                    said.add((sent, note))
                                    out.append({"s": sent[:500], "note": note, "kind": "value", "decision": False})
                        continue
                    if base == "r" and re.search(r"\b(?:negative|inverse)", sent, re.I) and _ts_match(v_t.lstrip("-−–"), [c.lstrip("-") for c in cands], base):
                        continue
                    where = f"Table {next(iter(named))}" if named and len(named) == 1 else "the table beside it"
                    uniq = list(dict.fromkeys(cands))
                    shown = uniq[0] if len(uniq) == 1 else " or ".join(uniq[:3]) + (", among others" if len(uniq) > 3 else "")
                    unit_ = "%" if kind in ("nagel%", "cox%", "overall") else ""
                    note = f"The text gives {name} as {v_t}{unit_}, but {where} gives {shown}{unit_ if len(uniq) == 1 else ''}."
                    if (sent, note) in said:
                        continue
                    said.add((sent, note))
                    out.append({"s": sent[:500], "note": note, "kind": "value", "decision": False})
    return out



# 6.6, round 156 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): round 155's readings held to the SPSS output beside
# them when the sentence gives no p of its own. "The knowledge scores were normally distributed (Table 4.6)" under Tests of
# Normality whose Sig. are .014 and .038; "equal variances were assumed" beside an Independent Samples Test whose Levene's Sig.
# is .035 (or a Test of Homogeneity of Variances); "the model fitted the data well" beside a Hosmer and Lemeshow Test with Sig.
# .002; "sphericity was assumed" beside Mauchly's Test of Sphericity with Sig. .007; "the data were suitable for factor
# analysis" beside a KMO and Bartlett's Test whose Bartlett's Sig. is .214. The sentence's table: the one it names, else the
# output just above or below it (as round 152). The rows it means: those it names, or all when it speaks of the data, the
# scores or the variables (or the table has one); a row whose two tests disagree is not judged.
_AT_TITLE = (
    ("norm", re.compile(r"^tests?\s+of\s+normality[a-e]?$", re.I), "Tests of Normality"),
    ("ist", re.compile(r"^independent\s+samples?\s+test[a-e]?$", re.I), "Independent Samples Test"),
    ("hov", re.compile(r"^tests?\s+of\s+homogeneity\s+of\s+variances?[a-e]?$", re.I), "Test of Homogeneity of Variances"),
    ("hl", re.compile(r"^hosmer\s+(?:and|&)\s+lemeshow\s+test[a-e]?$", re.I), "Hosmer and Lemeshow Test"),
    ("mau", re.compile(r"^mauchly['’]?s\s+test\s+of\s+sphericity[a-e]?$", re.I), "Mauchly’s Test of Sphericity"),
    ("kmo", re.compile(r"^kmo\s+and\s+bartlett['’]?s\s+test[a-e]?$", re.I), "KMO and Bartlett’s Test"),
)
_AT_SIGH = re.compile(r"^(?:sig\.?|p(?:-?\s*value)?)$", re.I)
_AT_GENERAL = re.compile(r"\b(?:the|all(?:\s+the)?|these|both(?:\s+the)?)\s+(?:data|variables|scores|distributions|dependent\s+variables?|measures|outcomes)\b|"
                         r"\b(?:each|every)\s+(?:variable|score|measure)\b|\ball\s+(?:variables|scores|measures)\b", re.I)
_AT_OTHERFIT = re.compile(r"\bomnibus\b|\bnagelkerke\b|\bcox\b|\bR\s*(?:²|2|-?\s*squared?)\b|\bclassif\w*|\bpredict\w*", re.I)
_AT_STOP = {"total", "score", "scores", "based", "group", "groups", "equal", "mean", "step", "the", "and", "for", "time", "item", "items"}


def _at_num(c):
    pc = _p_cell(c)
    return pc if pc and pc[1] <= 1 else None


def _at_find(rows):
    """The assumption tests of a cluster of table lines: [{"fam", "title", "rows": [(label, [(test, Sig. text)])]}]."""
    found, cells = [], [_cells_keep(r) for r in rows]
    filled = [[(c or "").strip() for c in cs if (c or "").strip()] for cs in cells]
    for i, f in enumerate(filled):
        if len(f) != 1:
            continue
        kind = next(((k, show) for k, rx, show in _AT_TITLE if rx.match(f[0])), None)
        if not kind:
            continue
        k, show = kind
        tb = {"fam": {"norm": "norm", "ist": "var", "hov": "var", "hl": "fit", "mau": "sph", "kmo": "bart"}[k], "title": show, "rows": []}
        end = min(len(rows), i + 40)
        for j in range(i + 1, end):                    # (the table ends at the next title)
            if len(filled[j]) == 1 and any(rx.match(filled[j][0]) for _k, rx, _s in _AT_TITLE):
                end = j
                break
        if k == "norm":
            names, sigs, h = {}, [], None
            for j in range(i + 1, min(end, i + 5)):
                for q, c in enumerate(cells[j]):
                    nm = next((sh for fam, rx, sh in _AS_NAMES if fam == "norm" and rx.search(c or "")), None)
                    if nm:
                        names[q] = nm
                    if _AT_SIGH.match((c or "").strip()):
                        sigs.append(q)
                        h = j
                if names and sigs:
                    break
            if not names or not sigs:
                continue
            cols = []
            for q in sigs:
                left = [x for x in names if x <= q]
                if left:
                    cols.append((q, names[max(left)]))
            label = ""
            for j in range(h + 1, end):
                f = filled[j]
                if not f or len(f) == 1:
                    if tb["rows"]:
                        break
                    continue
                if re.match(r"^(?:[a-e*]\.|\*+\.)\s", f[0]):
                    break
                pairs = [(t, cells[j][q]) for q, t in cols if q < len(cells[j]) and _at_num(cells[j][q])]
                if not pairs:
                    continue
                first = min(names)
                lab = " ".join(c for c in cells[j][:first] if c and not _num_cell(c))
                label = lab or label
                tb["rows"].append((label, pairs))
        elif k == "ist":
            lev, sp = None, None
            for j in range(i + 1, min(end, i + 5)):
                for q, c in enumerate(cells[j]):
                    if lev is None and re.match(r"^levene", c or "", re.I):
                        lev = q
                    elif lev is not None and sp is None and q >= lev and re.match(r"^sig\.?$", (c or "").strip(), re.I):
                        sp = q
            if sp is None:
                continue
            label = ""
            for j in range(i + 1, end):
                cs = cells[j]
                if not any(re.match(r"^equal\s+variances\s+assumed$", c or "", re.I) for c in cs):
                    continue
                label = (cs[0] or "").strip() or label
                if sp < len(cs) and _at_num(cs[sp]):
                    tb["rows"].append((label, [("Levene’s test", cs[sp])]))
        elif k in ("hov", "hl", "mau"):
            sp, var, h = None, "", None
            for j in range(i + 1, min(end, i + 5)):
                if len(filled[j]) == 1 and sp is None:
                    var = re.sub(r"^measure\s*:\s*", "", filled[j][0], flags=re.I)
                    continue
                for q, c in enumerate(cells[j]):
                    if sp is None and re.match(r"^sig\.?$", (c or "").strip(), re.I):
                        sp, h = q, j
                if sp is not None:
                    break
            if sp is None:
                continue
            for j in range(h + 1, end):
                cs = cells[j]
                if not filled[j] or len(filled[j]) == 1:
                    if tb["rows"]:
                        break                          # (a blank line or the next table's title ends its rows)
                    continue
                if sp < len(cs) and _at_num(cs[sp]):
                    lab = (cs[0] or "").strip()
                    if k == "hov":
                        tb["rows"].append((var or lab, [("Levene’s test", cs[sp])]))
                        break                          # (Based on Mean: the first row)
                    tb["rows"].append((lab, [(("Hosmer and Lemeshow’s test" if k == "hl" else "Mauchly’s test"), cs[sp])]))
            if k == "hl" and tb["rows"]:
                tb["rows"] = tb["rows"][-1:]           # (the last step)
        else:                                          # (KMO and Bartlett's Test: its "Sig." row)
            for j in range(i + 1, min(end, i + 6)):
                cs = [(c or "").strip() for c in cells[j]]
                if any(re.match(r"^sig\.?$", c, re.I) for c in cs):
                    v = next((c for c in reversed(cs) if _at_num(c)), None)
                    if v:
                        tb["rows"].append(("Bartlett’s test of sphericity", [("Bartlett’s test of sphericity", v)]))
                    break
        if tb["rows"]:
            found.append(tb)
    return found


def _at_named(c, rows):
    low = c.lower()
    hit = []
    for r in rows:
        lab = r[0].strip().rstrip("*").strip().lower()
        if not lab:
            continue
        if re.search(r"(?<![a-z])" + re.escape(lab) + r"(?![a-z])", low):
            hit.append(r)
            continue
        w = next((x for x in re.findall(r"[a-z][a-z'-]+", lab) if len(x) >= 4 and x not in _AT_STOP), None)
        if w and re.search(r"(?<![a-z])" + re.escape(w) + r"(?![a-z])", low):
            hit.append(r)
    return hit


def assume_table_check(text):
    out, said = [], set()
    lines = str(text or "").split("\n")
    kinds = [_ts_kind_line(raw) for raw in lines]
    blocks, cur = [], None
    for i, k in enumerate(kinds):
        typ = "table" if k in ("T", "C") else "prose" if k == "P" else None
        if typ is None:
            continue
        if cur and cur["type"] == typ:
            cur["idx"].append(i)
        else:
            cur = {"type": typ, "idx": [i]}
            blocks.append(cur)
    tables = [b for b in blocks if b["type"] == "table"]
    for b in tables:
        b["at"] = _at_find(lines[b["idx"][0]:b["idx"][-1] + 1])
        b["no"] = next((m.group(1) for i in b["idx"] for m in [_TT_CAPTION.search(lines[i])] if m and kinds[i] == "C"), None)
    if not any(b["at"] for b in tables):
        return out
    for bi, b in enumerate(blocks):
        if b["type"] != "prose":
            continue
        near = [x for x in (blocks[bi - 1] if bi > 0 else None, blocks[bi + 1] if bi + 1 < len(blocks) else None) if x is not None and x.get("at")]
        para = " ".join(lines[i].strip() for i in b["idx"])
        for sent in sentences(para):
            if _TS_CITE.search(sent) or _TS_OTHER.search(sent) or _NC_SENT.search(sent):
                continue                              # (6.6, round 166: a question, a hypothesis, an expectation is not a finding)
            s = _clean(sent)
            named = {m.group(1) for m in _TT_CAPTION.finditer(s)}
            pool = [t for t in tables if t["no"] in named and t.get("at")] if named else near
            ats = [a for t in pool for a in t["at"]]
            if not ats:
                continue
            am = _ALPHA_SET.search(s)
            alpha = float(am.group("a")) if am and am.group("a") else (float(am.group("pc")) / 100 if am and am.group("pc") else 0.05)
            a_str = f"{alpha:.{max(2, len(repr(alpha).split('.')[1]))}f}".lstrip("0")
            where = f"Table {next(iter(named))}" if named and len(named) == 1 and len(pool) == 1 else None
            bounds = [0] + [m.end() for m in _CLAUSE.finditer(s)] + [len(s)]
            for a, z in zip(bounds, bounds[1:]):
                c = s[a:z]
                if _AS_SKIP.search(c) or _P.search(c) or _AS_SIG.search(c) or _AS_PV.search(c) or _NC_CLAUSE.search(c):
                    continue                           # (a p of its own: rounds 154 and 155)
                for fam in dict.fromkeys(t["fam"] for t in ats):
                    if fam == "fit" and _AT_OTHERFIT.search(c):
                        continue
                    claim = _as_said(fam, c)
                    if not claim:
                        continue
                    tests = [sh for f_, rx, sh in _AS_NAMES if f_ == "norm" and rx.search(c)] if fam == "norm" else []
                    for tb in (t for t in ats if t["fam"] == fam):
                        rows = tb["rows"]
                        if fam in ("norm", "var", "sph") and len(rows) > 1:
                            hit = _at_named(c, rows)
                            rows = hit or (rows if _AT_GENERAL.search(c) else [])
                        if not rows:
                            continue
                        judged = []
                        for lab, pairs in rows:
                            use = [(t, v) for t, v in pairs if not tests or t in tests] or (pairs if not tests else [])
                            if not use:
                                continue
                            ps = [_at_num(v) for _t, v in use]
                            lo = all((rel in ("<", "≤") and v_ <= alpha) or (rel == "=" and v_ < alpha - 0.5 * 10 ** -d) for rel, v_, d in ps)
                            hi = all((rel in (">", "≥") and v_ >= alpha) or (rel == "=" and v_ > alpha + 0.5 * 10 ** -d) for rel, v_, d in ps)
                            judged.append((lab, use, "fails" if lo else "holds" if hi else None))
                        if not judged:
                            continue
                        if fam == "bart":
                            bad = [x for x in judged if x[2] == "holds"] if claim == "holds" else []
                        elif claim == "holds":
                            bad = [x for x in judged if x[2] == "fails"]
                        else:
                            bad = judged if all(x[2] == "holds" for x in judged) else []
                        if not bad:
                            continue
                        many = len(tb["rows"]) > 1

                        def sigs(use, named_tests):
                            vs = [_p_said(v, bare=True) for _t, v in use]
                            if len(use) > 1 or (fam == "norm" and named_tests):
                                return " and ".join(f"{v_} ({t})" for (t, _v), v_ in zip(use, vs))
                            return vs[0]
                        txt = "; ".join(f"Sig. {sigs(use, fam == 'norm')}" + (f" for {lab}" if many and lab else "") for lab, use, _j in bad[:3])
                        txt += " …" if len(bad) > 3 else ""
                        place = where or f"The {tb['title']} table beside it"
                        lev = "Levene’s test " if fam == "var" else ""
                        if fam == "bart":
                            note = (f"{place} gives Bartlett’s test {txt}, above {a_str}: the correlations may be too small for a factor analysis, "
                                    "though the sentence says the data suit one.")
                        elif claim == "holds":
                            note = {"norm": f"{place} gives {txt}, below {a_str}: a departure from a normal distribution, not normal data as the sentence says.",
                                    "var": f"{place} gives {lev}{txt}, below {a_str}: the variances differ, not equal as the sentence says.",
                                    "fit": f"{place} gives {txt}, below {a_str}: the model does not fit the data well, though the sentence says it does.",
                                    "sph": f"{place} gives {txt}, below {a_str}: sphericity is violated, though the sentence says it holds."}[fam]
                        else:
                            note = {"norm": f"{place} gives {txt}, above {a_str}: no departure from a normal distribution, though the sentence says the data aren’t normal.",
                                    "var": f"{place} gives {lev}{txt}, above {a_str}: no difference between the variances, though the sentence says they differ.",
                                    "fit": f"{place} gives {txt}, above {a_str}: no lack of fit, though the sentence says the model fits poorly.",
                                    "sph": f"{place} gives {txt}, above {a_str}: no violation of sphericity, though the sentence says it is violated."}[fam]
                        if (sent, note) in said:
                            continue
                        said.add((sent, note))
                        out.append({"s": sent[:500], "note": note, "kind": "value", "decision": True})
    return out



# 6.6, round 163 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a table's summary held to the table. "Age, sex and
# level of study were significantly associated with self-medication" under a table whose p for sex is .340; "residence was not
# significantly associated with it" where residence's p is below .001. The table: the one the sentence names, else the one just
# above or below it (as round 152). Each variable's p's: a typed table's p column (a variable's heading row and its categories
# under it, or a variable in the first column and its categories in the second), or SPSS's Variables in the Equation. Only a
# clause that says "significant" (or "not significant") and names a variable of the table, with no p of its own (round 154's);
# a variable with several p's (its categories) is significant where any is, not significant where none is.
_SS_PHEAD = re.compile(r"^(?:sig\.?(?:\s*\((?:1|2)-(?:tailed|sided)\))?|p(?:-?\s*values?)?|asymp(?:\.|totic)\s+sig(?:\.|nificance)(?:\s*\((?:1|2)-(?:tailed|sided)\))?)$", re.I)
_SS_STOPLAB = {"total", "overall", "overall statistics", "block", "constant", "step", "model", "valid", "missing", "variable", "variables", "characteristic", "characteristics", "factor",
               "factors", "category", "categories", "ref", "reference"}


def _ss_lab(c):
    return re.sub(r"\(\d\)$|:$", "", (c or "").strip()).strip()


def _ss_isnum(c):
    """A figure cell (a count, "30 (40.0)", a p, an odds ratio with its interval, "Ref"), not a label ("18–20", "24 and above")."""
    c = (c or "").strip()
    return bool(c) and (_num_cell(c) is not None or bool(re.match(r"^\d[\d,]*(?:\.\d+)?\s*\(", c)) or bool(re.match(r"^[<>≤≥=]\s*0?\.\d", c))
                        or c.lower() in ("ref", "ref.", "reference", "1.00 (ref)", "1 (ref)", "—", "-", "–"))


def _ss_map(rows):
    """{a variable's label (lower case): [its p cells]} of a cluster of table lines."""
    cells = [_cells_keep(r) for r in rows]
    out, cur, mode, pcols = {}, None, None, []    # mode: "head" (a heading row, its categories under it), "two" (variable | category), "step"
    for cs in cells:
        filled = [(c or "").strip() for c in cs if (c or "").strip()]
        if not filled:
            continue
        ks = [j for j, c in enumerate(cs) if _SS_PHEAD.match((c or "").strip())]
        two = [j for j in ks if re.search(r"\((?:1|2)-(?:tailed|sided)\)", cs[j] or "")]
        ks = two if two and len(two) < len(ks) else ks      # (6.6, round 167: Levene's "Sig." beside the t-test's "Sig. (2-tailed)")
        if ks:
            pcols, cur, mode = ks, None, None     # (each table's own heading: its p columns)
            continue
        if not pcols:
            continue
        if len(filled) == 1 and not re.search(r"\d", filled[0]) and len(filled[0]) <= 60 and (cs[0] or "").strip():
            cur, mode = _ss_lab(filled[0]), "head"
            continue
        labs = []
        for c in cs:
            if _ss_isnum(c):
                break
            labs.append((c or "").strip())
        ne = [x for x in labs if x]
        if not ne:
            continue
        if re.match(r"^step\s*\d+[a-z]?$", ne[0], re.I):
            mode, var = "step", (ne[1] if len(ne) > 1 else None)
        elif mode == "step" and not labs[0]:
            var = ne[0]
        elif len(ne) >= 2 and labs[0]:
            cur, mode = _ss_lab(ne[0]), "two"
            var = cur
        elif not labs[0] and mode == "two":
            var = cur
        elif labs[0] and mode == "head":
            var = cur
        else:
            var = ne[0]
        var = _ss_lab(var or "")
        if not var or var.lower() in _SS_STOPLAB:
            continue
        for k in pcols:
            pv = (cs[k] or "").strip() if k < len(cs) else ""
            pc = _p_cell(pv)
            if pc and pc[1] <= 1:
                out.setdefault(var.lower(), []).append(pv)
    return out


def _ss_said(c, lab):
    """Whether the clause names the variable (its label, or its first word of five letters or more)."""
    low = c.lower()
    if re.search(r"(?<![a-z])" + re.escape(lab) + r"(?![a-z])", low):
        return True
    w = re.findall(r"[a-z][a-z'-]+", lab)
    return bool(w and len(w[0]) >= 5 and w[0] not in {"level", "total", "other", "years", "number", "score", "group", "status", "place",
                                                        "about", "before", "after", "having"} and re.search(r"(?<![a-z])" + re.escape(w[0]) + r"(?![a-z])", low))


def sig_summary_check(text):
    out, said = [], set()
    lines = str(text or "").split("\n")
    kinds = [_ts_kind_line(raw) for raw in lines]
    blocks, cur = [], None
    for i, k in enumerate(kinds):
        typ = "table" if k in ("T", "C") else "prose" if k == "P" else None
        if typ is None:
            continue
        if cur and cur["type"] == typ:
            cur["idx"].append(i)
        else:
            cur = {"type": typ, "idx": [i]}
            blocks.append(cur)
    tables = [b for b in blocks if b["type"] == "table"]
    for b in tables:
        b["ss"] = _ss_map(lines[b["idx"][0]:b["idx"][-1] + 1])
        b["no"] = next((m.group(1) for i in b["idx"] for m in [_TT_CAPTION.search(lines[i])] if m and kinds[i] == "C"), None)
    if not any(b["ss"] for b in tables):
        return out
    for bi, b in enumerate(blocks):
        if b["type"] != "prose":
            continue
        near = [x for x in (blocks[bi - 1] if bi > 0 else None, blocks[bi + 1] if bi + 1 < len(blocks) else None) if x is not None and x.get("ss")]
        para = " ".join(lines[i].strip() for i in b["idx"])
        for sent in sentences(para):
            if _TS_CITE.search(sent) or _TS_OTHER.search(sent) or _NC_SENT.search(sent):
                continue                              # (6.6, round 166: a question, a hypothesis, an expectation is not a finding)
            st = _clean(sent)
            named = {m.group(1) for m in _TT_CAPTION.finditer(st)}
            pool = [t for t in tables if t["no"] in named and t.get("ss")] if named else near
            if len(pool) != 1:
                continue                              # (one table to hold the sentence to)
            mp = pool[0]["ss"]
            am = _ALPHA_SET.search(st)
            alpha = float(am.group("a")) if am and am.group("a") else 0.05
            a_str = f"{alpha:.2f}".lstrip("0")
            where = f"Table {pool[0]['no']}" if pool[0]["no"] else "The table beside it"
            bounds = [0] + [m.end() for m in _CLAUSE.finditer(st)] + [len(st)]
            for a, z in zip(bounds, bounds[1:]):
                c = st[a:z]
                if _P.search(c) or _AS_SIG.search(c) or _AS_PV.search(c) or _AS_SKIP.search(c) or _NC_CLAUSE.search(c):
                    continue
                nonsig = bool(_SIG_NOT.search(c))
                sig = not nonsig and bool(_SIG_YES.search(c))
                if not (sig or nonsig) or re.search(r"\b(?:normal|variances?|fit|sphericity|levene|shapiro|hosmer|omnibus)\b", c, re.I):
                    continue
                for lab, ps in mp.items():
                    if not _ss_said(c, lab):
                        continue
                    pcs = [_p_cell(x) for x in ps]
                    lo = [(rel in ("<", "≤") and v <= alpha) or (rel == "=" and v < alpha - 0.5 * 10 ** -d) for rel, v, d in pcs]
                    hi = [(rel in (">", "≥") and v >= alpha) or (rel == "=" and v > alpha + 0.5 * 10 ** -d) for rel, v, d in pcs]
                    sp_ = [_p_said(x) for x in ps[:3]]
                    its = [x[2:] for x in sp_] if all(x.startswith("= ") for x in sp_) else ["p " + x for x in sp_]
                    lead = "p = " if all(x.startswith("= ") for x in sp_) else ""
                    shown = lead + (", ".join(its) + ", among others" if len(ps) > 3 else its[0] if len(its) == 1 else ", ".join(its[:-1]) + " and " + its[-1])
                    name = lab[:1].upper() + lab[1:]
                    if sig and all(hi):
                        note = f"{where} gives {name} {shown} (above {a_str}): not significant, though the sentence says it is."
                    elif nonsig and all(lo):
                        note = f"{where} gives {name} {shown} (below {a_str}): significant, though the sentence says it isn’t."
                    else:
                        continue
                    if (sent, note) not in said:
                        said.add((sent, note))
                        out.append({"s": sent[:500], "note": note, "kind": "value", "decision": True})
    return out



# 6.6, round 164 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a mean and an SD the writing gives, held to the
# table beside it (or the one it names): "the mean knowledge score was 3.45 (SD = 0.62)" under a Descriptive Statistics table
# whose Knowledge row gives 3.54 and .62; "Group A had a mean of 11.59" under Group Statistics with 10.59. Only where the
# sentence names a row of the table (its label, or its first word of five letters or more) and gives one mean (or "M ± SD");
# the figure must be the named row's, rounded or cut to the sentence's decimals. An SD is read only after a mean column (in an
# item table "SD" before "Mean" is Strongly Disagree).
_MN_MEANH = re.compile(r"^(?:mean|m|mean\s+score|mean\s+\(?x̄\)?|x̄|average)$", re.I)
_MN_SDH = re.compile(r"^(?:std\.?\s*deviation|standard\s+deviation|s\.?d\.?|std\.?\s*dev\.?)$", re.I)
_MN_SAY = re.compile(r"(?:(?<![A-Za-z])M\s*=\s*|\b(?:mean|average)\b(?!\s+(?:differences?|squares?|ranks?)\b)[^.;:()=]{0,80}?(?:\b(?:of|was|is|were)\b|=|:|\()\s*"
                     r"(?:=\s*)?(?:M\s*=\s*)?|\baveraged\s+)(?P<v>\d+(?:\.\d+)?)"
                     r"(?!\s*%)|(?P<pm>\d+(?:\.\d+)?)\s*(?:±|\+/-|\+-)\s*(?P<pms>\d+(?:\.\d+)?)", re.I)
_MN_SDSAY = re.compile(r"(?<![A-Za-z])(?:SD|s\.d\.|standard\s+deviation)\s*(?:of|=|:|was)?\s*(?:=\s*)?(?P<v>\d*\.?\d+)", re.I)


def _mn_map(rows):
    """[(its labels, mean text, SD text or None)] of a cluster of table lines: each row of a table with a Mean column (a row
    under a two-column label, "Score | Group A" then " | Group B", keeps the first one: ["Score", "Group B"])."""
    out, mcol, scol, prev = [], None, None, []
    for raw in rows:
        cs = _cells_keep(raw)
        low = [(c or "").strip().lower() for c in cs]
        filled = [x for x in low if x]
        if not filled or len(filled) == 1:
            continue                                  # (a title or a sub-heading: the heading stands)
        if not any(_ss_isnum(c) for c in cs):          # (a heading: a new table's columns)
            mcol = next((j for j, c in enumerate(low) if c and _MN_MEANH.match(c)), None)
            scol = next((j for j, c in enumerate(low) if mcol is not None and j > mcol and c and _MN_SDH.match(c)), None)
            continue
        if mcol is None:
            continue
        labs = []
        for c in cs:
            if _ss_isnum(c):
                break
            labs.append((c or "").strip())
        m_t = (cs[mcol] or "").strip() if mcol < len(cs) else ""
        if _num_cell(m_t) is None:
            continue
        s_t = (cs[scol] or "").strip() if scol is not None and scol < len(cs) else None
        s_t = s_t if s_t and _num_cell(s_t) is not None else None
        names = [x for x in labs if x]
        if labs and not labs[0] and len(prev) >= 2:
            names = [prev[0]] + names
        prev = names if (labs and labs[0]) else prev
        names = [x for x in names if x.lower() not in _SS_STOPLAB and not x.lower().startswith("valid n")]
        if names:
            out.append((names, m_t, s_t))
    return out


def _mn_fits(v_t, c_t):
    """Whether the sentence's figure is the table's, rounded (or cut) to the sentence's decimals."""
    try:
        v, c = float(v_t), float(c_t.replace("−", "-"))
    except ValueError:
        return True
    d = _dec(v_t)
    if abs(v - c) <= 0.5 * 10 ** -d + 0.5 * 10 ** -_dec(c_t) + 1e-9:
        return True
    return _dec(c_t) > d and abs(math.floor(c * 10 ** d + 1e-9) / 10 ** d - v) < 1e-9


def mean_text_check(text):
    out, said = [], set()
    lines = str(text or "").split("\n")
    kinds = [_ts_kind_line(raw) for raw in lines]
    blocks, cur = [], None
    for i, k in enumerate(kinds):
        typ = "table" if k in ("T", "C") else "prose" if k == "P" else None
        if typ is None:
            continue
        if cur and cur["type"] == typ:
            cur["idx"].append(i)
        else:
            cur = {"type": typ, "idx": [i]}
            blocks.append(cur)
    tables = [b for b in blocks if b["type"] == "table"]
    for b in tables:
        b["mn"] = _mn_map(lines[b["idx"][0]:b["idx"][-1] + 1])
        b["no"] = next((m.group(1) for i in b["idx"] for m in [_TT_CAPTION.search(lines[i])] if m and kinds[i] == "C"), None)
    if not any(b["mn"] for b in tables):
        return out
    for bi, b in enumerate(blocks):
        if b["type"] != "prose":
            continue
        near = [x for x in (blocks[bi - 1] if bi > 0 else None, blocks[bi + 1] if bi + 1 < len(blocks) else None) if x is not None and x.get("mn")]
        para = " ".join(lines[i].strip() for i in b["idx"])
        for sent in sentences(para):
            if _TS_CITE.search(sent) or _TS_OTHER.search(sent) or _NC_SENT.search(sent):
                continue                              # (6.6, round 166: a question, a hypothesis, an expectation is not a finding)
            st = _clean(sent)
            named = {m.group(1) for m in _TT_CAPTION.finditer(st)}
            pool = [t for t in tables if t["no"] in named and t.get("mn")] if named else near
            if len(pool) != 1:
                continue
            rows = pool[0]["mn"]
            where = f"Table {pool[0]['no']}" if named and pool[0]["no"] else "The table beside it"
            bounds = [0] + [m.end() for m in _CLAUSE.finditer(st)] + [len(st)]
            for a, z in zip(bounds, bounds[1:]):
                c = st[a:z]
                ms = list(_MN_SAY.finditer(c))
                if len(ms) != 1:
                    continue
                m = ms[0]
                m_t = m.group("v") or m.group("pm")
                pm2 = re.match(r"\s*(?:±|\+/-|\+-)\s*(?P<s>\d+(?:\.\d+)?)", c[m.end():])          # ("averaged 2.47 ± 0.65")
                s_t = m.group("pms") or (pm2.group("s") if pm2 else None) or next((x.group("v") for x in _MN_SDSAY.finditer(c)), None)
                hit = [r for r in rows if _ss_said(c, r[0][-1].lower())]          # (the row's own name: "Group B")
                hit = hit or [r for r in rows if any(_ss_said(c, x.lower()) for x in r[0])]
                if not hit or len(hit) > 3:
                    continue
                if not any(_mn_fits(m_t, r[1]) for r in hit):
                    shown = " or ".join(dict.fromkeys(f"{r[1]} for {r[0][-1]}" for r in hit))
                    note = f"{where} gives a mean of {shown}, not {m_t}."
                elif s_t and any(r[2] for r in hit) and not any(r[2] and _mn_fits(s_t, r[2]) for r in hit if _mn_fits(m_t, r[1])):
                    shown = " or ".join(dict.fromkeys(f"{r[2]} for {r[0][-1]}" for r in hit if r[2] and _mn_fits(m_t, r[1])))
                    if not shown:
                        continue
                    note = f"{where} gives an SD of {shown}, not {s_t}."
                else:
                    continue
                if (sent, note) not in said:
                    said.add((sent, note))
                    out.append({"s": sent[:500], "note": note, "kind": "value", "decision": False})
    return out



# 6.6, round 165 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): "most", "the highest proportion" and "the most
# common" held to the frequency table beside the sentence (or the one it names): "the highest proportion of respondents were
# aged 20–24 (45.3%)" where 25–29 has 48.0%; "the least common was Level 400" where Level 300 has fewer. A typed table (a
# variable's heading row and its categories, or variable | category) or SPSS's Frequencies (its title the variable, "Valid"
# rows its categories); Total, Missing and System rows aside. Only a clause that names one category of a variable; a tie (to
# the table's decimals) stands.
_FQ_CNT = re.compile(r"^(?:frequency|freq\.?|n|count|number|no\.?|number\s+of\s+respondents|respondents)(?:\s*\(\s*n\s*\))?$", re.I)
_FQ_PCT = re.compile(r"^(?:percent(?:age)?|%|per\s*cent|percentage\s*\(\s*%\s*\)|percent\s*\(\s*%\s*\)|valid\s+percent)$", re.I)
_FQ_SKIP = {"valid", "missing", "total", "system", "overall", "valid total"}
_SP_MAX = re.compile(r"\b(?:most\s+(?:of\s+(?:the|our)\s+)?(?:respondents|participants|students|people|nurses|midwives|women|men|mothers|patients|"
                     r"pupils|workers|of\s+them)|the\s+majority|majority\s+of|the\s+(?:highest|largest|greatest|biggest)\s+(?:proportion|number|percentage|"
                     r"share|group|category|frequency)|the\s+most\s+common|the\s+commonest|the\s+predominant|the\s+modal|most\s+frequent(?:ly)?)\b", re.I)
_SP_MIN = re.compile(r"\b(?:the\s+(?:lowest|smallest|least)\s+(?:proportion|number|percentage|share|group|category|frequency)|the\s+fewest|"
                     r"the\s+least\s+(?:common|frequent))\b", re.I)


def _fq_key(x):
    return re.sub(r"\s*(?:–|—|-|to)\s*", "-", re.sub(r"\s+", " ", (x or "").strip().lower())).rstrip("*").strip()


def _fq_map(rows):
    """{variable: [(category, value text, is a percent)]} of a cluster of table lines with a count or percent column."""
    cells = [_cells_keep(r) for r in rows]
    out, title, cur, mode, col, pct, single = {}, None, None, None, None, False, None
    for cs in cells:
        filled = [(c or "").strip() for c in cs if (c or "").strip()]
        if not filled:
            continue
        if len(filled) == 1:
            if not re.search(r"\d", filled[0]) or re.match(r"^table\s+\d", filled[0], re.I):
                single = re.sub(r"^table\s+[\d.]+\s*[:.]?\s*", "", filled[0], flags=re.I)
                if col is None:
                    title = single
                else:
                    cur, mode = _ss_lab(filled[0]), "head"
            continue
        if not any(_ss_isnum(c) for c in cs):          # (a heading; the line just above it, a title, names its variable)
            title = single or title
            low = [(c or "").strip().lower() for c in cs]
            pj = next((j for j, c in enumerate(low) if c and _FQ_PCT.match(c)), None)
            cj = next((j for j, c in enumerate(low) if c and _FQ_CNT.match(c)), None)
            col, pct = (pj, True) if pj is not None else ((cj, False) if cj is not None else (None, False))
            cur, mode, single = None, None, None
            continue
        if col is None:
            continue
        single = None
        labs = []
        for c in cs:
            if _ss_isnum(c):
                break
            labs.append((c or "").strip())
        ne = [x for x in labs if x and x.lower() not in _FQ_SKIP]
        if not ne or any(x.lower() in ("total", "missing", "system") for x in labs if x):
            continue
        if mode == "two" and labs and not labs[0]:
            var, cat = cur, ne[-1]
        elif len(ne) >= 2:
            cur, mode = _ss_lab(ne[0]), "two"
            var, cat = cur, ne[-1]
        elif mode == "head":
            var, cat = cur, ne[0]
        else:
            var, cat = (title or "the table"), ne[0]       # (SPSS's Frequencies: its title is the variable)
        v_t = (cs[col] or "").strip() if col < len(cs) else ""
        v_t = re.sub(r"\s*%$", "", re.sub(r"^\d[\d,]*\s*\(\s*([\d.]+)\s*%?\s*\)$", r"\1", v_t) if pct else v_t)
        if _num_cell(v_t) is None:
            continue
        out.setdefault(var or "the table", []).append((cat, v_t, pct))
    return {k: v for k, v in out.items() if len(v) >= 2}


def superlative_table_check(text):
    out, said = [], set()
    lines = str(text or "").split("\n")
    kinds = [_ts_kind_line(raw) for raw in lines]
    blocks, cur = [], None
    for i, k in enumerate(kinds):
        typ = "table" if k in ("T", "C") else "prose" if k == "P" else None
        if typ is None:
            continue
        if cur and cur["type"] == typ:
            cur["idx"].append(i)
        else:
            cur = {"type": typ, "idx": [i]}
            blocks.append(cur)
    tables = [b for b in blocks if b["type"] == "table"]
    for b in tables:
        b["fq"] = _fq_map(lines[b["idx"][0]:b["idx"][-1] + 1])
        b["no"] = next((m.group(1) for i in b["idx"] for m in [_TT_CAPTION.search(lines[i])] if m and kinds[i] == "C"), None)
    if not any(b["fq"] for b in tables):
        return out
    for bi, b in enumerate(blocks):
        if b["type"] != "prose":
            continue
        near = [x for x in (blocks[bi - 1] if bi > 0 else None, blocks[bi + 1] if bi + 1 < len(blocks) else None) if x is not None and x.get("fq")]
        para = " ".join(lines[i].strip() for i in b["idx"])
        for sent in sentences(para):
            if _TS_CITE.search(sent) or _TS_OTHER.search(sent) or _NC_SENT.search(sent):
                continue                              # (6.6, round 166: a question, a hypothesis, an expectation is not a finding)
            st = _clean(sent)
            named = {m.group(1) for m in _TT_CAPTION.finditer(st)}
            pool = [t for t in tables if t["no"] in named and t.get("fq")] if named else near
            if len(pool) != 1:
                continue
            where = f"Table {pool[0]['no']}" if named and pool[0]["no"] else "The table beside it"
            bounds = [0] + [m.end() for m in _CLAUSE.finditer(st)] + [len(st)]
            for a, z in zip(bounds, bounds[1:]):
                c = st[a:z]
                mx, mn = _SP_MAX.search(c) or _SP_HALF.search(c), _SP_MIN.search(c)
                if bool(mx) == bool(mn) or _NC_CLAUSE.search(c) or re.search(r"\b(?:not|no|never)\s+(?:\w+\s+){0,2}$", c[:(mx or mn).start()], re.I):
                    continue
                low = _fq_key(c)
                hit_of = {var: [x for x in cats if re.search(r"(?<![a-z0-9])" + re.escape(_fq_key(x[0])) + r"(?:s|es)?(?![a-z0-9])", low)]
                          for var, cats in pool[0]["fq"].items()}
                if sum(1 for h_ in hit_of.values() if h_) > 1:
                    continue                          # (6.6, round 169: "the majority of the female respondents were single")
                for var, cats in pool[0]["fq"].items():
                    hit = hit_of[var]
                    if len(hit) != 1:
                        continue
                    cat, v_t, pct = hit[0]
                    v = float(v_t)
                    h = 0.5 * 10 ** -_dec(v_t)
                    others = [(c2, float(w), w) for c2, w, _p in cats if c2 != cat]
                    if mx:
                        top = max(others, key=lambda x: x[1])
                        if top[1] > v + 2 * h + 1e-9:
                            u = "%" if pct else ""
                            note = f"{where} gives {top[0]} {top[2]}{u}, more than {cat} ({v_t}{u}): the largest group is {top[0]}."
                        elif _SP_HALF.search(c) and not re.search(r"\d\s*%", c):
                            tot = sum(float(w) for _c2, w, _p in cats)
                            if pct and v < 50.0 - 1e-9:
                                note = f"{where} gives {cat} {v_t}%: less than half, so not a majority."
                            elif not pct and tot > 0 and 2 * v <= tot:
                                note = f"{where} gives {cat} {v_t} of {tot:g} ({100.0 * v / tot:.1f}%): not more than half, so not a majority."
                            else:
                                continue
                        else:
                            continue
                    else:
                        bot = min(others, key=lambda x: x[1])
                        if bot[1] < v - 2 * h - 1e-9:
                            u = "%" if pct else ""
                            note = f"{where} gives {bot[0]} {bot[2]}{u}, less than {cat} ({v_t}{u}): the smallest group is {bot[0]}."
                        else:
                            continue
                    if (sent, note) not in said:
                        said.add((sent, note))
                        out.append({"s": sent[:500], "note": note, "kind": "value", "decision": False})
    return out



# 6.6, round 166 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): (1) a sentence that is not a finding is not held to
# the table beside it. Rounds 152, 156, 163, 164 and 165 read "Ho1: There is no significant association between age and
# self-medication.", "Is there a significant association between sex and self-medication?", "The study sought to determine
# whether…", "It was hypothesised that…" and "were the scores normally distributed?" as findings, and told a right results
# chapter it was wrong (FOUND by this round's test; not live: 6.6 is unpublished). A question, a hypothesis, an expectation or a
# prediction is left alone (_NC_SENT), and in a clause, what a study sets out to find (_NC_CLAUSE). (2) A threshold is not a
# result: "A Cronbach's alpha of 0.70 or more was considered acceptable" was read as an alpha of 0.70 (round 152).
_NC_SENT = re.compile(r"\?|\bhypothes[ie]s\b|\bhypothesi[sz]ed\b|(?<![A-Za-z])H\s?[0oO₀aA₁1-9]\d?\s*:|"
                      r"\b(?:expected|predicted|anticipated)\s+(?:that|to)\b|\bwould\s+(?:be|have|show|find|differ|vary)\b", re.I)   # ("X predicted Y" is a finding)
_NC_CLAUSE = re.compile(r"\bwhether\b|\bif\b|\bto\s+(?:test|determine|examine|assess|establish|find\s+out|investigate|explore|ascertain)\b", re.I)
_TH_AFTER = re.compile(r"^\s*%?\s*(?:(?:or|and)\s+(?:more|above|higher|greater|over|better|beyond|below|less|lower|under)\b|\+|"
                       r"(?:was|is|were|are)\s+(?:taken|used|set|considered|regarded|deemed)\s+as\s+(?:the\s+)?(?:cut|threshold|benchmark|criterion|minimum))", re.I)
_TH_BEFORE = re.compile(r"(?:\bat\s+least|\babove|\bbelow|\bgreater\s+than|\bmore\s+than|\bless\s+than|\bhigher\s+than|\blower\s+than|\bover|\bunder|"
                        r"\bexceed(?:s|ing|ed)?|\bminimum(?:\s+of)?|\bmaximum(?:\s+of)?|\bthreshold(?:\s+of)?|\bcut-?\s?off(?:\s+(?:of|point\s+of))?|"
                        r"\bbenchmark(?:\s+of)?|\brecommended(?:\s+value)?(?:\s+of)?|≥|≤|>=|<=|>|<)\s*(?:the\s+|a\s+|an\s+|value\s+of\s+)?$", re.I)


# (3) What a paragraph says of a correlation, held to the SPSS Correlations table beside it (or the one it names): "knowledge
# was significantly correlated with practice" where their Sig. (2-tailed) is .144; "there was no relationship between knowledge
# and stress" where it is .000; "a positive correlation between knowledge and stress" where r is −.302; "a strong correlation"
# where r is .120 (round 93's words: weak below .3, moderate .3 to .5, strong .5 and above, with its room). Pearson's,
# Spearman's and Kendall's matrices (with two methods in one table, only a sentence that names its method). The pair: the two
# variables the clause names, by their labels (or the label without "score", "level" …). A clause with an r or a p of its own
# is rounds 152 and 154's; a question, a hypothesis or an aim is left alone.
_CM_COEF = re.compile(r"^(?:pearson\s+correlation|correlation\s+coefficient)$", re.I)
_CM_SIG = re.compile(r"^sig\.\s*\((?:1|2)-tailed\)$", re.I)
_CM_WORD = re.compile(r"\b(?:correlat\w*|relationships?|related|relat(?:es|ing)|associations?|associated|link(?:s|ed)?)\b", re.I)
_CM_NONE = re.compile(r"\bno\s+(?:statistically\s+)?(?:significant\s+)?(?:correlation|relationship|association|link|relation)s?\b|"
                      r"\bnot\s+(?:statistically\s+)?(?:significantly\s+)?(?:correlated|related|associated|linked)\b|"
                      r"\b(?:did|does|do)\s+not\s+(?:significantly\s+)?(?:correlate|relate)\b", re.I)
_CM_METH = (("rho", re.compile(r"spearman|\brho\b|ρ", re.I)), ("tau-b", re.compile(r"kendall|\btau\b|τ", re.I)), ("r", re.compile(r"pearson", re.I)))
_CM_OWN = re.compile(r"(?<![A-Za-z])(?:r|r_?s|rs|ρ|rho|τ|tau(?:[-_]?b)?)\s*(?:\(\s*\d+\s*\))?\s*=", re.I)
_CM_GENERIC = re.compile(r"\s*\([^)]*\)$|\s+(?:scores?|levels?|index|scale|total|mean|status)$", re.I)


def _cm_pairs(rows):
    """({(row label, column label) in lower case: {method: (coefficient, Sig.)}}, {label in lower case: as the table writes it}) of
    the SPSS Correlations matrices in a cluster of lines. The cells are lined up with the names over the columns; a row pasted
    without its last empty cells is filled out where its own diagonal shows where it stands."""
    cells = [_cells_keep(r) for r in rows]
    text = "\n".join(rows).lower()
    meth = "rho" if "spearman" in text else "tau-b" if "kendall" in text else "r"
    out, disp, head = {}, {}, None
    for i, cs in enumerate(cells):
        low = [(c or "").strip() for c in cs]
        if any(re.match(r"^spearman", c, re.I) for c in low[:3]):
            meth = "rho"
        elif any(re.match(r"^kendall", c, re.I) for c in low[:3]):
            meth = "tau-b"
        k = next((j for j, c in enumerate(low[:4]) if _CM_COEF.match(c)), None)
        if k is None:
            names = [c for c in low if c]
            if len(names) >= 2 and not any(_ss_isnum(c) for c in names) and not any(_CM_SIG.match(c) or c.lower() in ("n", "correlations") for c in names):
                head = names
            continue
        m_here = "r" if low[k].lower().startswith("pearson") else meth
        name = next((c for c in reversed(low[:k]) if c and re.search(r"[A-Za-z]", c) and not re.match(r"^(?:spearman|kendall)", c, re.I)), "")
        if not name or not head:
            continue
        sig = None
        for d in (1, 2):
            if i + d < len(cells):
                nx = [(c or "").strip() for c in cells[i + d]]
                ks = next((j for j, c in enumerate(nx[:4]) if _CM_SIG.match(c)), None)
                if ks is not None:
                    sig = nx[ks + 1:]
                    break
        if sig is None:
            continue
        hl = [x.lower() for x in head]
        jd = hl.index(name.lower()) if name.lower() in hl else None

        def fit(xs, is_r):
            xs = list(xs)
            while len(xs) > len(head) and not xs[-1]:
                xs.pop()
            if len(xs) < len(head) and jd is not None and ((is_r and jd < len(xs) and xs[jd].rstrip("*") in _CR_ONES) or (not is_r and len(xs) <= jd + 1)):
                xs += [""] * (len(head) - len(xs))
            return xs if len(xs) == len(head) else None
        rv, sv = fit(low[k + 1:], True), fit(sig, False)
        if rv is None or sv is None:
            continue
        for j, col in enumerate(head):
            r_t = rv[j].rstrip("*").strip()
            if j == jd or not r_t or r_t in _CR_ONES or _num_cell(r_t) is None or not _p_cell(sv[j]):
                continue
            out.setdefault((name.lower(), col.lower()), {})[m_here] = (r_t, sv[j])
            disp.setdefault(name.lower(), name)
            disp.setdefault(col.lower(), col)
    return out, disp


def _cm_find(low, lab):
    """Where a clause (in lower case) names a variable: its label, or the label without "score", "level", "(years)"…"""
    forms = [lab]
    bare = _CM_GENERIC.sub("", lab).strip()
    if bare != lab and len(bare) >= 3:
        forms.append(bare)
        lead = re.sub(r"^(?:total|mean|overall)\s+", "", bare)
        if lead != bare and len(lead) >= 3:
            forms.append(lead)
    for f in forms:
        m = re.search(r"(?<![a-z0-9])" + re.escape(f) + r"(?:s|es)?(?![a-z0-9])", low)
        if m:
            return m.span()
    return None


def corr_summary_check(text):
    out, said = [], set()
    lines = str(text or "").split("\n")
    kinds = [_ts_kind_line(raw) for raw in lines]
    blocks, cur = [], None
    for i, k in enumerate(kinds):
        typ = "table" if k in ("T", "C") else "prose" if k == "P" else None
        if typ is None:
            continue
        if cur and cur["type"] == typ:
            cur["idx"].append(i)
        else:
            cur = {"type": typ, "idx": [i]}
            blocks.append(cur)
    tables = [b for b in blocks if b["type"] == "table"]
    for b in tables:
        b["cm"], b["disp"] = _cm_pairs(lines[b["idx"][0]:b["idx"][-1] + 1])
        b["no"] = next((m.group(1) for i in b["idx"] for m in [_TT_CAPTION.search(lines[i])] if m and kinds[i] == "C"), None)
    if not any(b["cm"] for b in tables):
        return out
    for bi, b in enumerate(blocks):
        if b["type"] != "prose":
            continue
        near = [x for x in (blocks[bi - 1] if bi > 0 else None, blocks[bi + 1] if bi + 1 < len(blocks) else None) if x is not None and x.get("cm")]
        para = " ".join(lines[i].strip() for i in b["idx"])
        for sent in sentences(para):
            if _TS_CITE.search(sent) or _TS_OTHER.search(sent) or _NC_SENT.search(sent):
                continue
            st = _clean(sent)
            named = {m.group(1) for m in _TT_CAPTION.finditer(st)}
            pool = [t for t in tables if t["no"] in named and t.get("cm")] if named else near
            if len(pool) != 1:
                continue
            pairs, disp = pool[0]["cm"], pool[0]["disp"]
            where = f"Table {pool[0]['no']}" if named and pool[0]["no"] else "The table beside it"
            labels = sorted({x for ab in pairs for x in ab})
            meths = [m for m, rx in _CM_METH if rx.search(st)]
            am = _ALPHA_SET.search(st)
            alpha = float(am.group("a")) if am and am.group("a") else (float(am.group("pc")) / 100 if am and am.group("pc") else 0.05)
            a_str = f"{alpha:.{max(2, len(repr(alpha).split('.')[1]))}f}".lstrip("0")
            bounds = [0] + [m.end() for m in _CLAUSE.finditer(st)] + [len(st)]
            for a, z in zip(bounds, bounds[1:]):
                c = st[a:z]
                if not _CM_WORD.search(c) or _P.search(c) or _AS_SIG.search(c) or _AS_PV.search(c) or _CM_OWN.search(c) or _NC_CLAUSE.search(c):
                    continue
                low = c.lower()
                found = {lab: sp for lab in labels for sp in [_cm_find(low, lab)] if sp}
                found = {lab: sp for lab, sp in found.items()
                         if not any(o != lab and o2[0] <= sp[0] and sp[1] <= o2[1] and o2 != sp for o, o2 in found.items())}
                if len(found) != 2:
                    continue
                x1, x2 = sorted(found, key=lambda lab: found[lab][0])
                ent = dict(pairs.get((x2, x1), {}))
                ent.update(pairs.get((x1, x2), {}))
                if meths:
                    ent = {m: v for m, v in ent.items() if m in meths}
                if len(ent) != 1:
                    continue                          # (no pair, or two methods and the sentence names neither)
                meth, (r_t, s_t) = next(iter(ent.items()))
                rel, pv, d = _p_cell(s_t)
                sig_t = (rel in ("<", "≤") and pv <= alpha) or (rel == "=" and pv < alpha - 0.5 * 10 ** -d)
                non_t = (rel in (">", "≥") and pv >= alpha) or (rel == "=" and pv > alpha + 0.5 * 10 ** -d)
                nonsig = bool(_SIG_NOT.search(c) or _CM_NONE.search(c))
                sig = not nonsig and bool(_SIG_YES.search(c))
                try:
                    r = float(r_t.replace("−", "-").replace("–", "-"))
                except ValueError:
                    continue
                signs = {("+" if (x.group("w") or x.group("w2")).lower().startswith("pos") else "-") for x in _RS_SIGN.finditer(c)}
                pair_t = f"{disp.get(x1, x1)} and {disp.get(x2, x2)}"
                notes = []
                if sig and non_t:
                    notes.append((f"{where} gives {meth} = {r_t} (Sig. {s_t}) for {pair_t}: not significant at {a_str}, though the sentence says it is.", True))
                elif nonsig and sig_t:
                    notes.append((f"{where} gives {meth} = {r_t} (Sig. {s_t}) for {pair_t}: significant at {a_str}, though the sentence says it isn’t.", True))
                if len(signs) == 1 and r != 0 and (r < 0) != ("-" in signs):
                    notes.append((f"{where} gives {meth} = {r_t} for {pair_t}: {'negative' if r < 0 else 'positive'}, not {'positive' if r < 0 else 'negative'}.", False))
                sm = _STR_SAID.search(c)
                if sm and meth in ("r", "rho") and not nonsig:
                    w = sm.group(1).lower()
                    band = "strong" if w.startswith("strong") else "weak" if w.startswith("weak") else "moderate"
                    ok = {"strong": lambda v: v >= 0.4, "weak": lambda v: v < 0.4, "moderate": lambda v: 0.2 <= v < 0.7}[band]
                    if not ok(abs(r)):
                        usual = "weak" if abs(r) < 0.3 else "moderate" if abs(r) < 0.5 else "strong"
                        notes.append((f"{where} gives {meth} = {r_t} for {pair_t}: usually called {usual} (below .3 weak, .3 to .5 moderate, "
                                      f".5 and above strong), not {band}.", False))
                for note, dec in notes:
                    if (sent, note) not in said:
                        said.add((sent, note))
                        out.append({"s": sent[:500], "note": note, "kind": "value", "decision": dec})
    return out



# 6.6, round 167 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): which group a paragraph says scored higher, held to
# the group means beside it (SPSS's Group Statistics, Paired Samples Statistics or one-way Descriptives; or the one it names):
# "females had higher knowledge scores than males" where Female's mean is 3.12 and Male's 3.45; "scores after the programme were
# higher than before" where After's mean is lower; "Level 300 had the highest mean score" where Level 400's is higher. Only a
# clause that names two groups of one table around "higher/greater/more/lower/less … than" (or one group and "the highest/lowest
# mean", "scored highest"), and the table's variable (by its label, its main word, or, for a table of one variable called "Score",
# a word about scores or means). "Better", "worse" and "more likely" are left alone (higher is not always better); a difference
# within the table's rounding stands. A question, a hypothesis or an aim is not a finding (round 166).
_GC_UP = re.compile(r"\b(?:higher|greater|larger|bigger|more)\b", re.I)
_GC_DOWN = re.compile(r"\b(?:lower|smaller|less|fewer)\b", re.I)
_GC_LIKELY = re.compile(r"\b(?:more|less)\s+likely\b|\b(?:more|less)\s+(?:often|frequently)\b", re.I)
_GC_THAN = re.compile(r"\bthan\b|\bcompared\s+(?:with|to)\b|\brelative\s+to\b", re.I)
_GC_AFTER = re.compile(r"^\s*(?:(?:those|that|the\s+ones?)\s+(?:of|in|among|for)\s+|(?:in|among|for|of|by|with)\s+)?(?:the\s+)?", re.I)
_GS_MAX = re.compile(r"\b(?:highest|largest|greatest)\s+(?:\w+\s+){0,2}?(?:means?|averages?|scores?|mean\s+scores?)\b|\bscored\s+(?:the\s+)?highest\b|"
                     r"\b(?:means?|averages?|scores?)\s+(?:\w+\s+){0,2}?(?:was|were|is|are)\s+(?:the\s+)?highest\b", re.I)
_GS_MIN = re.compile(r"\b(?:lowest|smallest|least)\s+(?:\w+\s+){0,2}?(?:means?|averages?|scores?|mean\s+scores?)\b|\bscored\s+(?:the\s+)?lowest\b|"
                     r"\b(?:means?|averages?|scores?)\s+(?:\w+\s+){0,2}?(?:was|were|is|are)\s+(?:the\s+)?lowest\b", re.I)
_GC_SCOREW = re.compile(r"\b(?:scores?|scored|means?|averages?|averaged|marks?)\b", re.I)
_GC_STOP = {"yes", "no", "none", "other", "others", "total", "all", "missing", "valid", "both", "male/female", "n/a", "na"}


_GC_REPORT = re.compile(r"^report$", re.I)
_GC_CAPMEAN = re.compile(r"\b(?:means?|mean\s+scores?|scores?|averages?)\b", re.I)


def _gc_more(lines, caption=None):
    """6.6, round 171: [(variable or None, [(group, value text)], "a mean rank of" / "a mean of")] from the Ranks of a Mann–Whitney or
    Kruskal–Wallis test, SPSS's Report, and group means typed under a title about means."""
    out = []
    blocks, _ts = _np_read(lines)
    for b in blocks:
        if b["type"] == "groups":
            gs = [(r["g"], r["mr"]) for r in b["rows"] if _num_cell(r["mr"]) is not None]
            if len(gs) >= 2:
                out.append((b["name"], gs, "a mean rank of"))
    cells = [_cells_keep(l) for l in lines]
    i = 0
    while i < len(lines):
        filled = [(c or "").strip() for c in cells[i] if (c or "").strip()]
        if len(filled) == 1 and _GC_REPORT.match(filled[0]):          # (SPSS's Report)
            var, j = None, i + 1
            while j < len(lines) and len([c for c in cells[j] if (c or "").strip()]) == 1:
                var = var or next((c or "").strip() for c in cells[j] if (c or "").strip())
                j += 1
            if j < len(lines):
                low = [(c or "").strip().lower() for c in cells[j]]
                mcol = next((k for k, c in enumerate(low) if c == "mean"), None)
                if mcol is not None and "n" in low:
                    gs, k = [], j + 1
                    while k < len(lines) and any((c or "").strip() for c in cells[k]):
                        lab = (cells[k][0] or "").strip()
                        m_t = (cells[k][mcol] or "").strip() if mcol < len(cells[k]) else ""
                        if lab and lab.lower() != "total" and _num_cell(m_t) is not None:
                            gs.append((lab, m_t))
                        k += 1
                    if len(gs) >= 2:
                        out.append((var, gs, "a mean of"))
                    i = k
                    continue
        i += 1
    if caption and _GC_CAPMEAN.search(caption):                  # (group means typed under a title about means)
        var = _xt_outcome(caption)
        txt = [l for l in lines if not _TT_CAPTION.search(l)]
        if not any(re.match(r"^\s*(?:group\s+statistics|descriptives|report|descriptive\s+statistics|ranks)\s*$", l, re.I) for l in txt):
            rows = _mn_map(txt)
            if len(rows) >= 2 and all(len(r[0]) == 1 for r in rows):
                out.append((var, [(r[0][0], r[1]) for r in rows if r[0][0].lower() != "total"], "a mean of"))
    return out


def _gc_sets(lines, caption=None):
    """The groups and their means in a cluster of table lines: [(variable or None, [(group, mean text)], "a mean of")]."""
    sets = []
    for var, rows in _tt_groups(lines).items():                  # (Group Statistics; Paired Samples Statistics: "Pair 1")
        gs = [(g, m) for g, _n, m, _sd, _se, _l in rows if _num_cell(m) is not None]
        shown = next((cs[0].strip() for *_x, l in rows for cs in [_cells_keep(l)] if cs and cs[0].strip().lower() == var), var)   # (as the table writes it)
        if len(gs) >= 2:
            sets.append((None if re.match(r"^pair\s*\d+$", var) else shown, gs, "a mean of"))
    for i, rows, _j in _ad_tables(lines):                          # (one-way Descriptives: the variable on its own line above)
        var = None
        for back in (1, 2):
            if i - back >= 0:
                cs = [c for c in _cells_keep(lines[i - back]) if (c or "").strip()]
                if len(cs) == 1 and not re.search(r"\d", cs[0]) and cs[0].strip().lower() not in ("descriptives", "descriptive statistics"):
                    var = cs[0].strip()
                    break
        gs = [(lab, m) for lab, _n, m, _sd, _se, _lo, _hi, _l in rows if lab.lower() != "total" and _num_cell(m) is not None]
        if len(gs) >= 2:
            sets.append((var, gs, "a mean of"))
    sets += [x for x in _gc_more(lines, caption) if len(x[1]) >= 2]          # (6.6, round 171)
    return sets


def _gc_find(low, lab):
    """Where a clause (in lower case) names a group ("Female" in "females"; never "male" in "female")."""
    m = re.search(r"(?<![a-z0-9])" + re.escape(lab.lower()) + r"(?:s|es)?(?![a-z0-9])", low)
    return m.span() if m else None


def _gc_var_ok(c, var):
    """Whether the clause is about the table's variable: its label (or main word), or a word about scores where the variable is
    only a score."""
    if var is None:
        return bool(_GC_SCOREW.search(c))
    bare = _CM_GENERIC.sub("", var.strip().lower()).strip()
    bare = re.sub(r"^(?:total|mean|overall)\s+", "", bare)
    if not bare or bare in ("score", "scores", "mean", "total"):
        return bool(_GC_SCOREW.search(c))
    return bool(_cm_find(c.lower(), var.strip().lower()))


def group_compare_check(text):
    out, said = [], set()
    lines = str(text or "").split("\n")
    kinds = [_ts_kind_line(raw) for raw in lines]
    blocks, cur = [], None
    for i, k in enumerate(kinds):
        typ = "table" if k in ("T", "C") else "prose" if k == "P" else None
        if typ is None:
            continue
        if cur and cur["type"] == typ:
            cur["idx"].append(i)
        else:
            cur = {"type": typ, "idx": [i]}
            blocks.append(cur)
    tables = [b for b in blocks if b["type"] == "table"]
    for b in tables:
        cap_ = next((lines[i] for i in b["idx"] if kinds[i] == "C" and _TT_CAPTION.search(lines[i])), None)
        b["gc"] = _gc_sets(lines[b["idx"][0]:b["idx"][-1] + 1], cap_)
        b["no"] = next((m.group(1) for i in b["idx"] for m in [_TT_CAPTION.search(lines[i])] if m and kinds[i] == "C"), None)
    if not any(b["gc"] for b in tables):
        return out
    for bi, b in enumerate(blocks):
        if b["type"] != "prose":
            continue
        near = [x for x in (blocks[bi - 1] if bi > 0 else None, blocks[bi + 1] if bi + 1 < len(blocks) else None) if x is not None and x.get("gc")]
        para = " ".join(lines[i].strip() for i in b["idx"])
        for sent in sentences(para):
            if _TS_CITE.search(sent) or _TS_OTHER.search(sent) or _NC_SENT.search(sent):
                continue
            st = _clean(sent)
            named = {m.group(1) for m in _TT_CAPTION.finditer(st)}
            pool = [t for t in tables if t["no"] in named and t.get("gc")] if named else near
            if len(pool) != 1:
                continue
            where = f"Table {pool[0]['no']}" if named and pool[0]["no"] else "The table beside it"
            sets = pool[0]["gc"]
            bounds = [0] + [m.end() for m in _CLAUSE.finditer(st)] + [len(st)]
            for a, z in zip(bounds, bounds[1:]):
                c = st[a:z]
                if _NC_CLAUSE.search(c) or _GC_LIKELY.search(c) or re.search(r"\b(?:not|no|never|neither|nor)\b", c, re.I):
                    continue
                low = c.lower()
                for var, gs, what in sets:
                    labs = [(g, m_t) for g, m_t in gs if g.strip().lower() not in _GC_STOP and len(g.strip()) > 2 and re.search(r"[a-z]", g.lower())]
                    hits = {g: (sp, m_t) for g, m_t in labs for sp in [_gc_find(low, g)] if sp}
                    hits = {g: v for g, v in hits.items() if not any(o != g and o2[0][0] <= v[0][0] and v[0][1] <= o2[0][1] and o2[0] != v[0] for o, o2 in hits.items())}
                    if not hits or not _gc_var_ok(c, var):
                        continue
                    if var is None and sum(1 for v2, gs2, _w in sets if v2 is None and all(any(g == x for x, _m in gs2) for g in hits)) > 1:
                        continue                      # (two tables of these groups, neither naming its variable)
                    var_t = f" ({var})" if var else ""
                    note = None
                    th = _GC_THAN.search(c)
                    if len(hits) == 2 and th:
                        after = [g for g, (sp, _m) in hits.items() if sp[0] >= th.end() and _GC_AFTER.match(c[th.end():sp[0]]) and
                                 _GC_AFTER.match(c[th.end():sp[0]]).end() == sp[0] - th.end()]
                        before = [g for g, (sp, _m) in hits.items() if sp[1] <= th.start()]
                        if len(after) != 1 or len(before) != 1:
                            continue
                        g1, g2 = before[0], after[0]
                        win = c[max(0, th.start() - 80):th.start()]
                        up, down = list(_GC_UP.finditer(win)), list(_GC_DOWN.finditer(win))
                        if bool(up) == bool(down):
                            continue
                        m1_t, m2_t = hits[g1][1], hits[g2][1]
                        m1, m2 = float(m1_t.replace("−", "-")), float(m2_t.replace("−", "-"))
                        h = 0.5 * 10 ** -_dec(m1_t) + 0.5 * 10 ** -_dec(m2_t)
                        if up and m1 < m2 - h - 1e-9:
                            note = f"{where} gives {g1} {what} {m1_t} and {g2} {m2_t}{var_t}: {g1} is lower, not higher."
                        elif down and m1 > m2 + h + 1e-9:
                            note = f"{where} gives {g1} {what} {m1_t} and {g2} {m2_t}{var_t}: {g1} is higher, not lower."
                    elif len(hits) == 1 and not th and len(labs) >= 2:
                        mx, mn = _GS_MAX.search(c), _GS_MIN.search(c)
                        if bool(mx) == bool(mn):
                            continue
                        g, (sp, m_t) = next(iter(hits.items()))
                        v = float(m_t.replace("−", "-"))
                        others = [(o, float(t.replace("−", "-")), t) for o, t in gs if o != g and o.lower() != "total"]
                        if not others:
                            continue
                        h = 0.5 * 10 ** -_dec(m_t)
                        if mx:
                            top = max(others, key=lambda x: x[1])
                            if top[1] > v + 2 * h + 1e-9:
                                note = f"{where} gives {top[0]} {what} {top[2]}, higher than {g} ({m_t}){var_t}: the highest {what[2:-3]} is {top[0]}’s."
                        else:
                            bot = min(others, key=lambda x: x[1])
                            if bot[1] < v - 2 * h - 1e-9:
                                note = f"{where} gives {bot[0]} {what} {bot[2]}, lower than {g} ({m_t}){var_t}: the lowest {what[2:-3]} is {bot[0]}’s."
                    if note and (sent, note) not in said:
                        said.add((sent, note))
                        out.append({"s": sent[:500], "note": note, "kind": "value", "decision": False})
    return out



# 6.6, round 168 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): what a paragraph says of an SPSS crosstab, held to
# it. (1) "Sex was significantly associated with self-medication" where the Pearson Chi-Square's Sig. under "Sex * Self-medication
# Crosstabulation" is .330 (round 163 reads typed tables and logistic output; SPSS's crosstab names its variables only in its
# title). (2) "Females were more likely than males to self-medicate", "self-medication was more common among males than among
# females", "a higher proportion of rural than urban students…" where the crosstab's own counts give the other group the larger
# share. Only a clause that names the crosstab's row variable (for 1) or two of its row categories around "than" (for 2), and its
# column variable (or, for 2, one of its categories: "good knowledge"); a test named in the clause (Pearson's, Fisher's, the
# likelihood ratio) or a p of its own is round 163's and 154's; with Fisher's exact test beside it and the two disagreeing, the
# clause is not judged. Shares within half a percentage point stand. A question, a hypothesis or an aim is not a finding.
_XA_GENERIC = re.compile(r"\s*\([^)]*\)$|\s+(?:groups?|categor(?:y|ies)|status|levels?)$|\s+of\s+(?:the\s+)?(?:respondents?|participants?|students?|"
                         r"nurses?|patients?|mothers?|women|men)$", re.I)
_XA_UP = re.compile(r"\bmore\s+likely\b|\b(?:higher|greater|larger)\s+(?:proportions?|percentages?|shares?|rates?|prevalences?|odds)\b|"
                    r"\bmore\s+(?:common|prevalent|frequent|widespread|usual)\b|\b(?:higher|greater)\s+(?:among|in|for)\b", re.I)
_XA_DOWN = re.compile(r"\bless\s+likely\b|\b(?:lower|smaller)\s+(?:proportions?|percentages?|shares?|rates?|prevalences?|odds)\b|"
                      r"\bless\s+(?:common|prevalent|frequent|widespread|usual)\b|\b(?:lower|smaller)\s+(?:among|in|for)\b", re.I)
_XA_TEST = re.compile(r"\bpearson|\bfisher|\blikelihood\s+ratio|\bcontinuity\s+correction|\byates", re.I)
_XA_YES = {"yes", "no"}


def _xa_stem(w):
    w = w.lower()
    for suf in ("ations", "ation", "ating", "ated", "ates", "ate", "ions", "ion", "ings", "ing", "ers", "er", "ed", "es", "s", "e"):
        if w.endswith(suf) and len(w) - len(suf) >= 4:
            return w[:-len(suf)]
    return w


def _xa_where(c, label):
    """Where a clause names a crosstab's variable: its label (or the label without "group", "of respondents"…), or a word with
    the stem of the label's longest word ("self-medicate" for "Self-medication"); None if it doesn't."""
    low = c.lower()
    sp = _cm_find(low, label.lower())
    if sp:
        return sp
    bare = _XA_GENERIC.sub("", label.strip().lower()).strip()
    if bare and bare != label.strip().lower() and len(bare) >= 3:
        m = re.search(r"(?<![a-z0-9])" + re.escape(bare) + r"(?:s|es)?(?![a-z0-9])", low)
        if m:
            return m.span()
    words = [w for w in re.findall(r"[a-z][a-z'-]+", bare) if len(w) >= 5]
    if not words:
        return None
    st = _xa_stem(max(words, key=len))
    if len(st) < 4:
        return None
    return next((m.span() for m in re.finditer(r"[a-z][a-z'-]+", low) if _xa_stem(m.group(0)) == st), None)


def _xa_says(c, label):
    return _xa_where(c, label) is not None


# (the words that may stand between a comparison and the outcome it is about: "more likely to practise self-medication", "to
# have good knowledge", "a higher proportion of females than males self-medicated"; "to report side effects of self-medication"
# is about something else)
_XA_FILL = {"practise", "practice", "practising", "practicing", "practised", "practiced", "engage", "engaged", "engaging", "in", "have", "has",
            "had", "having", "report", "reported", "reporting", "use", "used", "using", "be", "being", "do", "did", "show", "showed", "possess",
            "possessed", "demonstrate", "demonstrated", "the", "a", "an", "were", "was", "is", "are", "to", "of"}
_XA_FILL |= {"respondents", "participants", "students", "nurses", "midwives", "patients", "mothers", "women", "men", "people", "pupils",
             "workers", "individuals", "children", "adolescents", "those", "ones", "counterparts"}
_XA_SUBJ = _XA_FILL | {"prevalence", "rate", "rates", "level", "levels", "proportion", "proportions", "percentage", "practice", "of", "overall"}


def _xa_fill_ok(txt, allowed=_XA_FILL):
    return all(w in allowed for w in re.findall(r"[a-z][a-z'-]*", txt.lower()))


def _xa_sets(text):
    """SPSS crosstabs with their chi-square tests: [{"rv", "cv" (as SPSS writes them), "cats" (column names), "groups":
    [(row category, [counts])], "p" (Pearson's Sig.), "fisher" (its Exact Sig. (2-sided), if any), "at" (its title's line)}]."""
    lines = str(text or "").split("\n")
    out = []
    for ct in _crosstabs(text):
        names = list(ct["names"])
        groups = []
        for bl in ct["blocks"]:
            if bl["label"].lower() == "total":
                continue
            cnt = bl["rows"].get("count")
            if not cnt:
                continue
            vals = [_ct_num(v) for v in cnt[1]]
            if any(v is None for v in vals):
                continue
            groups.append((bl["label"], vals))
        if len(groups) < 2:
            continue
        k = len(groups[0][1])
        if len(names) >= k:
            names = names[:k]
        elif len(names) == k - 1:
            vals_total = all(abs(v[-1] - sum(v[:-1])) < 1e-9 for _g, v in groups)
            if vals_total:
                groups = [(g, v[:-1]) for g, v in groups]
                k -= 1
        if len(names) != k or any(len(v) != k for _g, v in groups):
            continue
        if names[-1].lower() == "total":
            names, groups = names[:-1], [(g, v[:-1]) for g, v in groups]
        at = next((i for i in range(min(ct["end"], len(lines)) - 1, -1, -1) if _CT_TITLE.match(lines[i].strip())), None)
        mt = _CT_TITLE.match(lines[at].strip()) if at is not None else None
        rv, cv = (mt.group("r").strip(), mt.group("c").strip()) if mt else (ct["rv"], ct["cv"])
        p_t, f_t = None, None
        for i in range(ct["end"], min(len(lines), ct["end"] + 30)):
            if _CT_TITLE.match(lines[i].strip()):
                break
            low = [(x or "").strip().lower() for x in _cells_keep(lines[i])]
            if "value" not in low or "df" not in low:
                continue
            sig_k = [j for j, x in enumerate(low) if re.match(r"^asymp(?:\.|totic)\s+sig(?:\.|nificance)", x)]
            ex_k = [j for j, x in enumerate(low) if re.match(r"^exact\s+sig(?:\.|nificance)\s*\(2-sided\)", x)]
            for i3 in range(i + 1, min(len(lines), i + 10)):
                c3 = _cells_keep(lines[i3])
                first = re.sub(r"[a-e]$", "", (c3[0] or "").strip()).lower() if c3 else ""
                if first.startswith("pearson chi") and sig_k and sig_k[0] < len(c3) and _p_cell(c3[sig_k[0]]):
                    p_t = c3[sig_k[0]].strip()
                elif first.startswith("fisher") and ex_k and ex_k[0] < len(c3) and _p_cell(c3[ex_k[0]]):
                    f_t = c3[ex_k[0]].strip()
            break
        out.append({"rv": rv, "cv": cv, "cats": names, "groups": groups, "p": p_t, "fisher": f_t, "at": at})
    return out


def _xa_share_note(c, x, where):
    """(judged, note): which group of a cross-table a clause says has the larger share of the outcome, held to its counts. x:
    {"cv": the outcome's name (None: only a category named in the clause, not Yes/No), "cats": the outcome's categories,
    "groups": [(row category, [counts])]}."""
    low = c.lower()
    th = _GC_THAN.search(c)
    up, down = _XA_UP.search(c), _XA_DOWN.search(c)
    if not th or bool(up) == bool(down) or re.search(r"\b(?:not|no|never|neither|nor)\b", c, re.I):
        return False, None
    labs = [g for g, _v in x["groups"] if g.strip().lower() not in _GC_STOP and len(g.strip()) > 2]
    hits = {g: sp for g in labs for sp in [_gc_find(low, g)] if sp}
    hits = {g: sp for g, sp in hits.items() if not any(o != g and o2[0] <= sp[0] and sp[1] <= o2[1] and o2 != sp for o, o2 in hits.items())}
    after = [g for g, sp in hits.items() if sp[0] >= th.end() and _GC_AFTER.match(c[th.end():sp[0]]) and
             _GC_AFTER.match(c[th.end():sp[0]]).end() == sp[0] - th.end()]
    before = [g for g, sp in hits.items() if sp[1] <= th.start()]
    if len(hits) != 2 or len(after) != 1 or len(before) != 1:
        return False, None
    cats = [(j, cat, _gc_find(low, cat)) for j, cat in enumerate(x["cats"]) if cat.strip().lower() not in _XA_YES and _gc_find(low, cat)]
    if len(cats) == 1:
        j, cat, ow = cats[0]
    elif not cats and {cat.strip().lower() for cat in x["cats"]} == _XA_YES and x.get("cv") and _xa_where(c, x["cv"]):
        j = next(i for i, cat in enumerate(x["cats"]) if cat.strip().lower() == "yes")
        cat, ow = x["cats"][j], _xa_where(c, x["cv"])
    else:
        return False, None
    g1, g2 = before[0], after[0]
    cmp_ = (up or down)
    if ow[0] < cmp_.start():                      # (the outcome is the subject: "self-medication was more common among …")
        if not _xa_fill_ok(c[:ow[0]], _XA_SUBJ):
            return False, None
    else:                                          # ("… more likely than males to self-medicate", "… than males self-medicated")
        ends = [m.end() for m in re.finditer(r"\bto\b", c[cmp_.end():ow[0]], re.I)]
        anchor = max([cmp_.end() + e for e in ends] + [hits[g2][1] if hits[g2][1] <= ow[0] else -1])
        if anchor < 0 or not _xa_fill_ok(c[anchor:ow[0]]):
            return False, None
    v1, v2 = dict(x["groups"])[g1], dict(x["groups"])[g2]
    if sum(v1) <= 0 or sum(v2) <= 0:
        return False, None
    p1, p2 = 100.0 * v1[j] / sum(v1), 100.0 * v2[j] / sum(v2)
    what = f"“{cat}” ({x['cv']})" if x.get("cv") else f"“{cat}”"
    if up and p1 < p2 - 0.5:
        return True, f"{where} gives {g1} {p1:.1f}% and {g2} {p2:.1f}% for {what}: lower among {g1}, not higher."
    if down and p1 > p2 + 0.5:
        return True, f"{where} gives {g1} {p1:.1f}% and {g2} {p2:.1f}% for {what}: higher among {g1}, not lower."
    return True, None


def crosstab_text_check(text):
    out, said = [], set()
    lines = str(text or "").split("\n")
    sets = _xa_sets(text)
    if not sets:
        return out
    kinds = [_ts_kind_line(raw) for raw in lines]
    blocks, cur = [], None
    for i, k in enumerate(kinds):
        typ = "table" if k in ("T", "C") else "prose" if k == "P" else None
        if typ is None:
            continue
        if cur and cur["type"] == typ:
            cur["idx"].append(i)
        else:
            cur = {"type": typ, "idx": [i]}
            blocks.append(cur)
    tables = [b for b in blocks if b["type"] == "table"]
    for b in tables:
        a_, z_ = b["idx"][0], b["idx"][-1]
        b["xa"] = [x for x in sets if x["at"] is not None and a_ <= x["at"] <= z_]
        b["no"] = next((m.group(1) for i in b["idx"] for m in [_TT_CAPTION.search(lines[i])] if m and kinds[i] == "C"), None)
    for bi, b in enumerate(blocks):
        if b["type"] != "prose":
            continue
        near = [x for x in (blocks[bi - 1] if bi > 0 else None, blocks[bi + 1] if bi + 1 < len(blocks) else None) if x is not None and x.get("xa")]
        para = " ".join(lines[i].strip() for i in b["idx"])
        for sent in sentences(para):
            if _TS_CITE.search(sent) or _TS_OTHER.search(sent) or _NC_SENT.search(sent):
                continue
            st = _clean(sent)
            named = {m.group(1) for m in _TT_CAPTION.finditer(st)}
            pool = [t for t in tables if t["no"] in named and t.get("xa")] if named else near
            if len(pool) != 1:
                continue
            where = f"Table {pool[0]['no']}" if named and pool[0]["no"] else "The table beside it"
            am = _ALPHA_SET.search(st)
            alpha = float(am.group("a")) if am and am.group("a") else (float(am.group("pc")) / 100 if am and am.group("pc") else 0.05)
            a_str = f"{alpha:.{max(2, len(repr(alpha).split('.')[1]))}f}".lstrip("0")
            bounds = [0] + [m.end() for m in _CLAUSE.finditer(st)] + [len(st)]
            for a, z in zip(bounds, bounds[1:]):
                c = st[a:z]
                if _NC_CLAUSE.search(c) or _P.search(c) or _AS_SIG.search(c) or _AS_PV.search(c) or _XA_TEST.search(c):
                    continue
                low = c.lower()
                xs = [x for x in pool[0]["xa"] if _xa_says(c, x["rv"])]
                note, dec = None, False
                # (1) significance
                nonsig = bool(_SIG_NOT.search(c) or _CM_NONE.search(c))
                sig = not nonsig and bool(_SIG_YES.search(c))
                ow = _xa_where(c, xs[0]["cv"]) if len(xs) == 1 else None
                anchor = max([m.end() for m in re.finditer(r"\b(?:with|and|between|to|on)\b", c[:ow[0]], re.I)] or [-1]) if ow else -1
                if (sig or nonsig) and ow and anchor >= 0 and _xa_fill_ok(c[anchor:ow[0]]) and xs[0]["p"]:
                    x = xs[0]
                    rel, pv, d = _p_cell(x["p"])
                    lo = (rel in ("<", "≤") and pv <= alpha) or (rel == "=" and pv < alpha - 0.5 * 10 ** -d)
                    hi = (rel in (">", "≥") and pv >= alpha) or (rel == "=" and pv > alpha + 0.5 * 10 ** -d)
                    if x["fisher"]:
                        fr, fv, fd = _p_cell(x["fisher"])
                        flo = (fr in ("<", "≤") and fv <= alpha) or (fr == "=" and fv < alpha - 0.5 * 10 ** -fd)
                        fhi = (fr in (">", "≥") and fv >= alpha) or (fr == "=" and fv > alpha + 0.5 * 10 ** -fd)
                        lo, hi = lo and flo, hi and fhi
                    if sig and hi:
                        note, dec = (f"{where} gives the Pearson chi-square for {x['rv']} and {x['cv']} Sig. {x['p']} (above {a_str}): not significant, "
                                     f"though the sentence says it is."), True
                    elif nonsig and lo:
                        note, dec = (f"{where} gives the Pearson chi-square for {x['rv']} and {x['cv']} Sig. {x['p']} (below {a_str}): significant, "
                                     f"though the sentence says it isn’t."), True
                # (2) which group has the larger share
                if note is None:
                    for x in pool[0]["xa"]:
                        judged, note = _xa_share_note(c, x, where)
                        if judged:
                            break
                if note and (sent, note) not in said:
                    said.add((sent, note))
                    out.append({"s": sent[:500], "note": note, "kind": "value", "decision": dec})
    return out



# 6.6, round 169 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): round 168's shares for the cross-tables people type
# (Word tables as theses set them out: "Variable | Category | Yes n (%) | No n (%) | χ² | p-value", a variable's categories under
# it): "females were more likely than males to self-medicate" where the table's counts give males the larger share. The outcome:
# a category the clause names ("good knowledge"), or, for Yes/No columns, the outcome the heading above them or the table's title
# names ("Factors associated with self-medication"). The same guards as round 168.
_XT_HEAD = re.compile(r"^(?P<cat>[A-Za-z][^()]*?)\s*,?\s*(?:n\s*\(\s*%\s*\)|\(\s*n\s*,?\s*%\s*\)|f\s*\(\s*%\s*\)|freq(?:uency)?\.?\s*\(\s*%\s*\)|"
                      r"\(\s*n\s*=\s*\d+\s*\)\s*n\s*\(\s*%\s*\))$", re.I)
_XT_CELL = re.compile(r"^(?P<n>\d[\d,]*)\s*\(\s*\d+(?:\.\d+)?\s*%?\s*\)$")
_XT_CAP = (re.compile(r"\b(?:associated\s+with|association\s+(?:between\s+.+?\s+and|with)|relationship\s+between\s+.+?\s+and|"
                      r"factors\s+(?:influencing|affecting|that\s+influence|determining)|determinants\s+of|predictors\s+of)\s+"
                      r"(?P<o>[A-Za-z][A-Za-z\s'-]{2,40}?)(?:\s+(?:among|in|at|by|for|of)\s+.*)?\s*$", re.I),
           re.compile(r"^(?P<o>[A-Za-z][A-Za-z\s'-]{2,40}?)\s+(?:of\s+(?:the\s+)?(?:respondents|participants|students)\s+)?(?:by|across|according\s+to)\b", re.I))
_XT_NOT_O = re.compile(r"^(?:the\s+)?(?:respondents?|participants?|students?|study|sample|background(?:\s+characteristics)?|characteristics|"
                       r"socio-?\s?demographic(?:\s+characteristics)?|demographic(?:\s+characteristics)?|variables?|results?|cross-?tabulation)$", re.I)


def _xt_outcome(caption):
    """The outcome a table's title names ("Factors associated with self-medication" → "self-medication"), or None."""
    cap = re.sub(r"^table\s+[\d.]+\s*[:.]?\s*", "", (caption or "").strip(), flags=re.I)
    for rx in _XT_CAP:
        m = rx.search(cap)
        if m and not _XT_NOT_O.match(m.group("o").strip()):
            return m.group("o").strip()
    return None


def _xt_sets(rows, caption=None):
    """The typed cross-tables in a cluster of lines: [{"cv", "cats", "groups"}] (one per variable)."""
    cells = [_cells_keep(r) for r in rows]
    out, cols, cv, var, mode, groups = [], None, None, None, None, {}

    def flush():
        for v, gs in groups.items():
            if len(gs) >= 2:
                out.append({"cv": cv, "cats": [c for _j, c in cols], "groups": gs, "var": v})
        groups.clear()
    for i, cs in enumerate(cells):
        low = [(c or "").strip() for c in cs]
        hd = [(j, _XT_HEAD.match(c).group("cat").strip()) for j, c in enumerate(low) if c and _XT_HEAD.match(c)]
        if len(hd) >= 2:
            flush()
            cols, var, mode = hd, None, None
            cv = None
            if i > 0:                                 # (a heading above the Yes/No columns names the outcome)
                up_ = [(cells[i - 1][j] or "").strip() if j < len(cells[i - 1]) else "" for j, _c in hd]
                if up_[0] and not any(up_[1:]) and re.search(r"[A-Za-z]{4}", up_[0]):
                    cv = up_[0]
            if cv is None and caption:
                cv = _xt_outcome(caption)
            continue
        if not cols:
            continue
        filled = [c for c in low if c]
        if not filled:
            continue
        vals = []
        for j, _c in cols:
            m = _XT_CELL.match(low[j]) if j < len(low) else None
            vals.append(float(m.group("n").replace(",", "")) if m else None)
        labs = []
        for c in low:
            if _ss_isnum(c) or _XT_CELL.match(c):
                break
            labs.append(c)
        ne = [x for x in labs if x]
        if any(v is None for v in vals):
            if len(filled) == 1 and not re.search(r"\d", filled[0]) and low[0]:
                var, mode = _ss_lab(filled[0]), "head"         # (a variable's own heading row)
            continue
        if not ne:
            continue
        if len(ne) >= 2 and labs and labs[0]:
            var, mode, cat = _ss_lab(ne[0]), "two", ne[-1]
        else:
            cat = ne[-1]
        if not var or cat.lower() in ("total", "overall"):
            continue
        groups.setdefault(var, []).append((cat, vals))
    flush()
    return out


def typed_share_check(text):
    out, said = [], set()
    lines = str(text or "").split("\n")
    kinds = [_ts_kind_line(raw) for raw in lines]
    blocks, cur = [], None
    for i, k in enumerate(kinds):
        typ = "table" if k in ("T", "C") else "prose" if k == "P" else None
        if typ is None:
            continue
        if cur and cur["type"] == typ:
            cur["idx"].append(i)
        else:
            cur = {"type": typ, "idx": [i]}
            blocks.append(cur)
    tables = [b for b in blocks if b["type"] == "table"]
    for b in tables:
        cap = next((lines[i] for i in b["idx"] if kinds[i] == "C" and _TT_CAPTION.search(lines[i])), None)
        b["xt"] = _xt_sets(lines[b["idx"][0]:b["idx"][-1] + 1], cap)
        b["no"] = next((m.group(1) for i in b["idx"] for m in [_TT_CAPTION.search(lines[i])] if m and kinds[i] == "C"), None)
    if not any(b["xt"] for b in tables):
        return out
    for bi, b in enumerate(blocks):
        if b["type"] != "prose":
            continue
        near = [x for x in (blocks[bi - 1] if bi > 0 else None, blocks[bi + 1] if bi + 1 < len(blocks) else None) if x is not None and x.get("xt")]
        para = " ".join(lines[i].strip() for i in b["idx"])
        for sent in sentences(para):
            if _TS_CITE.search(sent) or _TS_OTHER.search(sent) or _NC_SENT.search(sent):
                continue
            st = _clean(sent)
            named = {m.group(1) for m in _TT_CAPTION.finditer(st)}
            pool = [t for t in tables if t["no"] in named and t.get("xt")] if named else near
            if len(pool) != 1:
                continue
            where = f"Table {pool[0]['no']}" if named and pool[0]["no"] else "The table beside it"
            bounds = [0] + [m.end() for m in _CLAUSE.finditer(st)] + [len(st)]
            for a, z in zip(bounds, bounds[1:]):
                c = st[a:z]
                if _NC_CLAUSE.search(c) or _P.search(c) or re.search(r"\d\s*%", c):
                    continue
                for x in pool[0]["xt"]:
                    judged, note = _xa_share_note(c, x, where)
                    if judged:
                        if note and (sent, note) not in said:
                            said.add((sent, note))
                            out.append({"s": sent[:500], "note": note, "kind": "value", "decision": False})
                        break
    return out


# (2) "The majority of the respondents were female", "most respondents were Christians", "more than half were married": more than
# half, not only the largest group (round 165 holds the largest group; round 112 a figure in the sentence). And a clause that
# names categories of two variables ("the majority of the female respondents were single") is not read as either.
_SP_HALF = re.compile(r"\b(?:the\s+)?majority\b|\bmost\s+(?:of\s+(?:the|our)\s+)?(?:respondents|participants|students|people|nurses|midwives|women|men|"
                      r"mothers|patients|pupils|workers|of\s+them)\b|\b(?:more|over)\s+than\s+half\b|\bover\s+half\b", re.I)



# 6.6, round 170 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): which way a predictor goes, in words, held to the
# coefficient table beside it: "knowledge positively predicted practice" where the Coefficients give Knowledge B = −.244 (a.
# Dependent Variable: Practice score); "stress was negatively associated with sleep quality" where its B is positive. SPSS's
# Coefficients (the outcome from its footnote) and coefficient tables people type (the outcome from the title: "Predictors of
# practice"); logistic regression's Variables in the Equation by a variable's own row (a category's row, "Sex(1)", is left
# alone: which category it is can't be told). Only a clause that names a predictor of the table, the outcome, and one direction
# ("positively/negatively predicted", "a positive/negative predictor", "inversely related"); a coefficient of its own (round
# 104's), a negated direction, a question, a hypothesis or an aim is left alone.
_RC_SIGN = re.compile(r"\b(?P<w>positive(?:ly)?|negative(?:ly)?|inverse(?:ly)?)\s+(?:and\s+)?(?:(?:statistically\s+)?significant(?:ly)?\s+)?"
                      r"(?:predict\w*|associat\w*|correlat\w*|relat\w*|influenc\w*|effects?|impacts?|contribut\w*|affect\w*)|"
                      r"\b(?:predict\w*|associat\w*|related|influenc\w*|affect\w*|contribut\w*)\s+(?:(?:statistically\s+)?significantly\s+)?(?:and\s+)?"
                      r"(?P<w2>positively|negatively|inversely)\b", re.I)
_RC_OWN = re.compile(r"(?<![A-Za-z])(?:B|b|β|beta|Beta|AOR|OR|Exp\s*\(\s*B\s*\))\s*=", re.I)
_RC_NEG = re.compile(r"\b(?:not|no|never)\s+(?:\w+\s+)?(?:positive|negative|inverse)|\b(?:not|n['’]t)\s+(?:\w+\s+)?(?:predict|associat|relat|influenc|affect)", re.I)
_RC_DV = re.compile(r"^[a-e]\.\s*dependent\s+variable\s*:\s*(?P<dv>.+?)\s*$", re.I)


def _rc_tables(lines, caption=None):
    """The coefficient tables in a cluster of lines: [{"dv", "rows": {label (lower): (label, B text)}, "kind"}]."""
    allc = [_cells_keep(l) for l in lines]
    out, cur = [], None
    for idx, cells in enumerate(allc):
        if not any(c for c in cells):
            continue
        heads = {}
        for k, c in enumerate(cells):
            for key, rx in _REG_COL.items():
                if rx.match((c or "").strip()) and key not in heads:
                    heads[key] = k
        if "b" in heads and "se" in heads and not any(k in heads for k in ("wald", "t", "z", "expb")):
            up = next((allc[j] for j in range(idx - 1, max(idx - 3, -1), -1) if any(c for c in allc[j])), None)
            for k, c in enumerate(up or []):      # (SPSS's heading on two lines)
                for key in ("t", "z", "p", "wald", "df"):
                    if key not in heads and k not in heads.values() and _REG_COL[key].match((c or "").strip()):
                        heads[key] = k
        if "b" in heads and "se" in heads and any(k in heads for k in ("wald", "t", "z", "expb")):
            cur = {"cols": heads, "rows": {}, "kind": "logit" if ("wald" in heads or "expb" in heads) else "linear",
                   "dv": _xt_outcome(caption) if caption else None}
            out.append(cur)
            continue
        if cur is None:
            continue
        first = (cells[0] or "").strip()
        m = _RC_DV.match(first)
        if m:
            cur["dv"] = m.group("dv")
            continue
        lk, lab = _row_label(cells)
        if lk is None or re.match(r"^\s*\(?\s*constant\b", lab, re.I) or re.search(r"\(\d+\)\s*$", lab) or len(cells) <= cur["cols"]["b"]:
            continue
        b_t = (cells[cur["cols"]["b"]] or "").strip()
        if _num_cell(b_t) is None or b_t.lower().startswith("ref"):
            continue
        cur["rows"].setdefault(lab.lower(), (lab, b_t))
    return [t for t in out if t["rows"]]


def coef_sign_check(text):
    out, said = [], set()
    lines = str(text or "").split("\n")
    kinds = [_ts_kind_line(raw) for raw in lines]
    blocks, cur = [], None
    for i, k in enumerate(kinds):
        typ = "table" if k in ("T", "C") else "prose" if k == "P" else None
        if typ is None:
            continue
        if cur and cur["type"] == typ:
            cur["idx"].append(i)
        else:
            cur = {"type": typ, "idx": [i]}
            blocks.append(cur)
    tables = [b for b in blocks if b["type"] == "table"]
    for b in tables:
        cap = next((lines[i] for i in b["idx"] if kinds[i] == "C" and _TT_CAPTION.search(lines[i])), None)
        b["rc"] = [t for t in _rc_tables(lines[b["idx"][0]:b["idx"][-1] + 1], cap) if t["dv"]]
        b["no"] = next((m.group(1) for i in b["idx"] for m in [_TT_CAPTION.search(lines[i])] if m and kinds[i] == "C"), None)
    if not any(b["rc"] for b in tables):
        return out
    for bi, b in enumerate(blocks):
        if b["type"] != "prose":
            continue
        near = [x for x in (blocks[bi - 1] if bi > 0 else None, blocks[bi + 1] if bi + 1 < len(blocks) else None) if x is not None and x.get("rc")]
        para = " ".join(lines[i].strip() for i in b["idx"])
        for sent in sentences(para):
            if _TS_CITE.search(sent) or _TS_OTHER.search(sent) or _NC_SENT.search(sent):
                continue
            st = _clean(sent)
            named = {m.group(1) for m in _TT_CAPTION.finditer(st)}
            pool = [t for t in tables if t["no"] in named and t.get("rc")] if named else near
            if len(pool) != 1:
                continue
            where = f"Table {pool[0]['no']}" if named and pool[0]["no"] else "The table beside it"
            bounds = [0] + [m.end() for m in _CLAUSE.finditer(st)] + [len(st)]
            for a, z in zip(bounds, bounds[1:]):
                c = st[a:z]
                if _NC_CLAUSE.search(c) or _RC_OWN.search(c) or _RC_NEG.search(c) or _P.search(c):
                    continue
                signs = {("+" if (x.group("w") or x.group("w2")).lower().startswith("pos") else "-") for x in _RC_SIGN.finditer(c)}
                polar = {("+" if x.group(0).lower().startswith("pos") else "-") for x in re.finditer(r"\b(?:positive|negative|inverse)(?:ly)?\b", c, re.I)}
                if len(signs) != 1 or len(polar) != 1:
                    continue                          # ("knowledge negatively and stress positively predicted practice": not read)
                low = c.lower()
                for t in pool[0]["rc"]:
                    dvw = _xa_where(c, t["dv"])
                    if not dvw:
                        continue
                    found = {k_: sp for k_ in t["rows"] for sp in [_cm_find(low, k_)] if sp and not (sp[0] < dvw[1] and dvw[0] < sp[1])}
                    found = {k_: sp for k_, sp in found.items() if not any(o != k_ and o2[0] <= sp[0] and sp[1] <= o2[1] and o2 != sp for o, o2 in found.items())}
                    for k_ in sorted(found, key=lambda x: found[x][0]):
                        lab, b_t = t["rows"][k_]
                        bv = _num_cell(b_t)
                        if bv is None or bv == 0 or (bv < 0) == ("-" in signs):
                            continue
                        note = f"{where} gives {lab} B = {b_t} for {t['dv']}: {'negative' if bv < 0 else 'positive'}, not {'positive' if bv < 0 else 'negative'}."
                        if (sent, note) not in said:
                            said.add((sent, note))
                            out.append({"s": sent[:500], "note": note, "kind": "value", "decision": False})
                    if found:
                        break
    return out



# 6.6, round 172 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): round 166's reading of a correlation in words, held
# to the correlation table people type the APA way ("Variable | M | SD | 1 | 2 | 3", rows "1. Knowledge", the lower triangle, its
# note "*p < .05. **p < .01."): "knowledge was significantly correlated with practice" where r = .12 with N = 150 (p = .144);
# "no significant relationship" where r = −.30**; "positive" where r is negative; "strong" where r = .12. Significance from r and
# N (the caption's or the note's "N = …", allowed r's rounding), or, with no N, only a starred r said not to be significant.
def _apa_mats(lines):
    """APA correlation tables in a list of lines: [{"pairs": {(label, label) lower: (r text, stars)}, "disp", "legend", "n", "one"}]."""
    out = []
    for h, line in enumerate(lines):
        cells = _cells_keep(line)
        nums = [(k, int(c)) for k, c in enumerate(cells) if re.match(r"^\d{1,2}$", (c or "").strip())]
        if len(nums) < 2 or [v for _k, v in nums] != list(range(1, len(nums) + 1)):
            continue
        col_at = {v: k for k, v in nums}
        rows, j = {}, h + 1
        while j < len(lines) and j < h + 40:
            c2 = _cells_keep(lines[j])
            m = _ST_ROW.match((c2[0] or "").strip()) if c2 else None
            if not m:
                if rows:
                    break
                j += 1
                continue
            rows[int(m.group("i"))] = (m.group("lab").strip(), c2)
            j += 1
        if len(rows) < 2:
            continue
        around = "\n".join(lines[max(0, h - 5):min(len(lines), j + 6)])
        legend = {}
        for lm in _ST_LEGEND.finditer(around):
            legend.setdefault(lm.group("s"), float(lm.group("v")))
        nm = _ST_N.search(around)
        n = int(nm.group("n").replace(",", "")) if nm else None
        one = bool(re.search(r"one[-\s]tailed", around, re.I)) and not re.search(r"two[-\s]tailed", around, re.I)
        pairs, disp = {}, {}
        for i, (lab, c2) in rows.items():
            for v, k in col_at.items():
                if v == i or k >= len(c2) or v not in rows:
                    continue
                mm = re.match(r"^(?P<r>[-−–]?\.\d{1,3}|[-−–]?0\.\d{1,3})(?P<s>\*{1,3}|†)?$", (c2[k] or "").strip())
                if not mm:
                    continue
                a_, b_ = lab.lower(), rows[v][0].lower()
                pairs[(a_, b_)] = (mm.group("r").replace("−", "-").replace("–", "-"), mm.group("s") or "")
                disp[a_], disp[b_] = lab, rows[v][0]
        if pairs:
            out.append({"pairs": pairs, "disp": disp, "legend": legend, "n": n if n and n >= 4 else None, "one": one, "at": (h, j)})
    return out


def apa_corr_check(text):
    out, said = [], set()
    lines = str(text or "").split("\n")
    mats = _apa_mats(lines)
    if not mats:
        return out
    kinds = [_ts_kind_line(raw) for raw in lines]
    blocks, cur = [], None
    for i, k in enumerate(kinds):
        typ = "table" if k in ("T", "C") else "prose" if k == "P" else None
        if typ is None:
            continue
        if cur and cur["type"] == typ:
            cur["idx"].append(i)
        else:
            cur = {"type": typ, "idx": [i]}
            blocks.append(cur)
    tables = [b for b in blocks if b["type"] == "table"]
    for b in tables:
        a_, z_ = b["idx"][0], b["idx"][-1]
        b["apa"] = [m for m in mats if a_ <= m["at"][0] <= z_]
        b["no"] = next((m.group(1) for i in b["idx"] for m in [_TT_CAPTION.search(lines[i])] if m and kinds[i] == "C"), None)
        if b["no"] is None:                           # ("Table 2" on its own line, the title under it)
            b["no"] = next((m.group(1) for i in b["idx"] for m in [re.match(r"^\s*table\s+(\d+(?:\.\d+)?)\s*$", lines[i], re.I)] if m), None)
    for bi, b in enumerate(blocks):
        if b["type"] != "prose":
            continue
        near = [x for x in (blocks[bi - 1] if bi > 0 else None, blocks[bi + 1] if bi + 1 < len(blocks) else None) if x is not None and x.get("apa")]
        para = " ".join(lines[i].strip() for i in b["idx"])
        for sent in sentences(para):
            if _TS_CITE.search(sent) or _TS_OTHER.search(sent) or _NC_SENT.search(sent):
                continue
            st = _clean(sent)
            named = {m.group(1) for m in _TT_CAPTION.finditer(st)}
            pool = [t for t in tables if t["no"] in named and t.get("apa")] if named else near
            if len(pool) != 1 or len(pool[0]["apa"]) != 1:
                continue
            mt = pool[0]["apa"][0]
            pairs, disp = mt["pairs"], mt["disp"]
            where = f"Table {pool[0]['no']}" if named and pool[0]["no"] else "The table beside it"
            labels = sorted({x for ab in pairs for x in ab})
            am = _ALPHA_SET.search(st)
            alpha = float(am.group("a")) if am and am.group("a") else (float(am.group("pc")) / 100 if am and am.group("pc") else 0.05)
            a_str = f"{alpha:.{max(2, len(repr(alpha).split('.')[1]))}f}".lstrip("0")
            bounds = [0] + [m.end() for m in _CLAUSE.finditer(st)] + [len(st)]
            for a, z in zip(bounds, bounds[1:]):
                c = st[a:z]
                if not _CM_WORD.search(c) or _P.search(c) or _AS_SIG.search(c) or _AS_PV.search(c) or _CM_OWN.search(c) or _NC_CLAUSE.search(c):
                    continue
                low = c.lower()
                found = {lab: sp for lab in labels for sp in [_cm_find(low, lab)] if sp}
                found = {lab: sp for lab, sp in found.items() if not any(o != lab and o2[0] <= sp[0] and sp[1] <= o2[1] and o2 != sp for o, o2 in found.items())}
                if len(found) != 2:
                    continue
                x1, x2 = sorted(found, key=lambda lab: found[lab][0])
                pr = pairs.get((x1, x2)) or pairs.get((x2, x1))
                if not pr:
                    continue
                r_t, stars = pr
                r = float(r_t)
                hh = 0.5 * 10 ** -_dec(r_t.lstrip("-"))
                nonsig = bool(_SIG_NOT.search(c) or _CM_NONE.search(c))
                sig = not nonsig and bool(_SIG_YES.search(c))
                pair_t = f"{disp.get(x1, x1)} and {disp.get(x2, x2)}"
                notes = []
                if (sig or nonsig) and mt["n"]:
                    n = mt["n"]
                    p_lo = p_r(min(abs(r) + hh, 0.999999), n - 2)
                    p_hi = p_r(max(abs(r) - hh, 0.0), n - 2)
                    p_at = p_r(abs(r), n - 2)
                    if None not in (p_lo, p_hi, p_at):
                        if mt["one"]:
                            p_lo, p_hi, p_at = p_lo / 2, p_hi / 2, p_at / 2
                        shown = _fmt_p(p_at, False)
                        if sig and p_lo > alpha:
                            notes.append((f"{where} gives r = {r_t} for {pair_t}; with N = {n} that is {shown}: not significant at {a_str}, though the sentence "
                                          f"says it is.", True))
                        elif nonsig and p_hi < alpha:
                            notes.append((f"{where} gives r = {r_t} for {pair_t}; with N = {n} that is {shown}: significant at {a_str}, though the sentence "
                                          f"says it isn’t.", True))
                elif nonsig and stars and stars in mt["legend"] and mt["legend"][stars] <= alpha + 1e-12:
                    notes.append((f"{where} gives r = {r_t}{stars} for {pair_t}, starred p < {_st_level(mt['legend'][stars])}: significant, though the "
                                  f"sentence says it isn’t.", True))
                signs = {("+" if (x.group("w") or x.group("w2")).lower().startswith("pos") else "-") for x in _RS_SIGN.finditer(c)}
                if len(signs) == 1 and r != 0 and (r < 0) != ("-" in signs):
                    notes.append((f"{where} gives r = {r_t} for {pair_t}: {'negative' if r < 0 else 'positive'}, not {'positive' if r < 0 else 'negative'}.", False))
                sm = _STR_SAID.search(c)
                if sm and not nonsig:
                    w = sm.group(1).lower()
                    band = "strong" if w.startswith("strong") else "weak" if w.startswith("weak") else "moderate"
                    ok = {"strong": lambda v: v >= 0.4, "weak": lambda v: v < 0.4, "moderate": lambda v: 0.2 <= v < 0.7}[band]
                    if not ok(abs(r)):
                        usual = "weak" if abs(r) < 0.3 else "moderate" if abs(r) < 0.5 else "strong"
                        notes.append((f"{where} gives r = {r_t} for {pair_t}: usually called {usual} (below .3 weak, .3 to .5 moderate, .5 and above "
                                      f"strong), not {band}.", False))
                for note, dec in notes:
                    if (sent, note) not in said:
                        said.add((sent, note))
                        out.append({"s": sent[:500], "note": note, "kind": "value", "decision": dec})
    return out



def problems(sentence):
    """What one sentence's own figures get wrong: [{"note", "kind", "decision"}] (empty when nothing)."""
    s = _clean(sentence)
    counted = [0]
    return (_test_notes(s, counted) + _coef_notes(s, counted) + _range_notes(s) + _ci_notes(s) + _pct_notes(s) + _split_notes(s) + _order_notes(s)
            + _scale_notes(s) + _variance_notes(s) + _strength_notes(s) + _effect_notes(s) + _direction_notes(s)
            + _majority_notes(s) + _p_word_notes(s) + _minmax_notes(s) + _assume_notes(s) + _points_notes(s) + _iqr_notes(s) + _scale_word_notes(s) + _null_notes(s))


def check(text):
    """The writing's statistics: ([{"s": the sentence, "note", "kind", "decision"}], how many tests were worked out again)."""
    out, seen = [], set()
    counted = [0]
    for sent in sentences(text):
        s = _clean(sent)
        if not re.search(r"\d", s):
            found = _hypothesis_notes(s)           # 6.6, round 97: "Ho: There is a significant relationship…" has no figure
        else:
            found = (_test_notes(s, counted) + _coef_notes(s, counted) + _range_notes(s) + _ci_notes(s) + _pct_notes(s) + _split_notes(s)
                     + _order_notes(s) + _scale_notes(s) + _alpha_word_notes(s) + _variance_notes(s) + _strength_notes(s)
                     + _effect_notes(s) + _whole_notes(s) + _hypothesis_notes(s) + _direction_notes(s) + _majority_notes(s)
                     + _p_word_notes(s) + _minmax_notes(s) + _assume_notes(s) + _points_notes(s) + _iqr_notes(s) + _scale_word_notes(s) + _dir_notes(s) + _null_notes(s))
        for n in found:
            key = (s, n["note"])
            if key in seen:
                continue
            seen.add(key)
            out.append(dict(n, s=sent[:500]))
    try:                                         # 6.6, round 117: its frequency tables, row by row
        out += table_check(text)
    except Exception as e:
        _fault("frequency-table", e)
    try:                                         # 6.6, round 118: its tables of item means
        out += mean_table_check(text)
    except Exception as e:
        _fault("item-mean table", e)
    try:                                         # 6.6, round 119: its cross-tables' chi-square, worked out again
        found = xtab_check(text)
        counted[0] += len(_xtab_blocks(text))
        out += found
    except Exception as e:
        _fault("cross-table", e)
    try:                                         # 6.6, round 120: its regression tables, worked out again from their columns
        out += reg_table_check(text, counted)
    except Exception as e:
        _fault("regression table", e)
    try:                                         # 6.6, round 121: its ANOVA tables
        out += anova_table_check(text, counted)
    except Exception as e:
        _fault("ANOVA table", e)
    try:                                         # 6.6, round 122: its tables of two groups' means with a t-test
        out += t_table_check(text, counted)
    except Exception as e:
        _fault("two-group t-test table", e)
    try:                                         # 6.6, round 123: its response rate
        out += _rr_notes(text)
    except Exception as e:
        _fault("response-rate", e)
    try:                                         # 6.6, round 124: SPSS's Descriptive Statistics tables
        out += desc_table_check(text)
    except Exception as e:
        _fault("descriptive-statistics table", e)
    try:                                         # 6.6, round 125: SPSS's Chi-Square Tests tables
        out += spss_chi_check(text, counted)
    except Exception as e:
        _fault("chi-square tests table", e)
    try:                                         # 6.6, round 126: SPSS's Correlations tables
        out += corr_table_check(text, counted)
    except Exception as e:
        _fault("correlation table", e)
    try:                                         # 6.6, round 127: SPSS's Model Summary, with its ANOVA
        out += model_summary_check(text)
    except Exception as e:
        _fault("model summary", e)
    try:                                         # 6.6, round 129: SPSS's Frequencies tables
        out += spss_freq_check(text)
    except Exception as e:
        _fault("SPSS frequencies", e)
    try:                                         # 6.6, round 130: the words against the tables
        out += text_table_check(text)
    except Exception as e:
        _fault("text-and-table", e)
    try:                                         # 6.6, round 131: SPSS's Crosstabulation, with its Chi-Square Tests
        out += spss_crosstab_check(text)
    except Exception as e:
        _fault("crosstab", e)
    try:                                         # 6.6, round 132: the stars of an APA correlation table
        out += star_table_check(text)
    except Exception as e:
        _fault("table-stars", e)
    try:                                         # 6.6, round 133: SPSS's independent-samples t-test
        out += spss_ttest_check(text, counted)
    except Exception as e:
        _fault("independent-samples t-test", e)
    try:                                         # 6.6, round 134: SPSS's one-way ANOVA, its Descriptives with its ANOVA
        out += spss_anova_desc_check(text)
    except Exception as e:
        _fault("ANOVA descriptives", e)
    try:                                         # 6.6, round 135: SPSS's paired-samples t-test
        out += spss_paired_check(text, counted)
    except Exception as e:
        _fault("paired t-test", e)
    try:                                         # 6.6, round 142: Levene's test and the post hoc comparisons after an ANOVA
        out += spss_posthoc_check(text, counted)
    except Exception as e:
        _fault("post hoc", e)
    try:                                         # 6.6, round 144: a table of ratios with their 95% intervals
        out += or_table_check(text, counted)
    except Exception as e:
        _fault("ratio table", e)
    try:                                         # 6.6, round 145: SPSS's logistic regression output, read together
        out += spss_logit_check(text, counted)
    except Exception as e:
        _fault("logistic output", e)
    try:                                         # 6.6, round 146: a one-sample t-test; a chi-square goodness of fit
        out += spss_onesample_check(text, counted)
    except Exception as e:
        _fault("one-sample t-test", e)
    try:
        out += spss_gof_check(text, counted)
    except Exception as e:
        _fault("goodness of fit", e)
    try:                                         # 6.6, round 147: tests on ranks (Mann–Whitney, Kruskal–Wallis, Wilcoxon, Friedman)
        out += spss_np_check(text, counted)
    except Exception as e:
        _fault("tests on ranks", e)
    try:                                         # 6.6, round 148: the reliability analysis (Cronbach's alpha), read together
        out += spss_rel_check(text, counted)
    except Exception as e:
        _fault("reliability", e)
    try:                                         # 6.6, round 149: Tests of Between-Subjects Effects; collinearity; skewness's Std. Errors
        out += spss_glm_check(text, counted)
    except Exception as e:
        _fault("between-subjects effects", e)
    try:
        out += collin_check(text)
    except Exception as e:
        _fault("collinearity", e)
    try:
        out += skew_se_check(text)
    except Exception as e:
        _fault("skewness and kurtosis", e)
    try:                                         # 6.6, round 150: the factor analysis, read together
        out += spss_fa_check(text, counted)
    except Exception as e:
        _fault("factor analysis", e)
    try:                                         # 6.6, round 151: Crosstabs' Symmetric Measures
        out += symm_check(text)
    except Exception as e:
        _fault("symmetric measures", e)
    try:                                         # 6.6, round 152: the words against the output beside them
        out += text_stat_check(text)
    except Exception as e:
        _fault("text against output", e)
    try:                                         # 6.6, round 156: an assumption read against the output beside it
        out += assume_table_check(text)
    except Exception as e:
        _fault("assumption against output", e)
    try:                                         # 6.6, round 162: "<.000" in a table; a negative item-total correlation
        out += null_table_check(text)
    except Exception as e:
        _fault("impossible cells", e)
    try:                                         # 6.6, round 163: a table's summary held to the table
        out += sig_summary_check(text)
    except Exception as e:
        _fault("table summary", e)
    try:                                         # 6.6, round 164: a mean and an SD held to the table's row
        out += mean_text_check(text)
    except Exception as e:
        _fault("means against the table", e)
    try:                                         # 6.6, round 165: "most", "the highest proportion" held to the frequency table
        out += superlative_table_check(text)
    except Exception as e:
        _fault("largest group", e)
    try:                                         # 6.6, round 166: a correlation's summary held to the Correlations table
        out += corr_summary_check(text)
    except Exception as e:
        _fault("correlation summary", e)
    try:                                         # 6.6, round 167: which group scored higher, held to the group means
        out += group_compare_check(text)
    except Exception as e:
        _fault("group comparison", e)
    try:                                         # 6.6, round 168: a crosstab's association and its shares, held to it
        out += crosstab_text_check(text)
    except Exception as e:
        _fault("crosstab in words", e)
    try:                                         # 6.6, round 169: a typed cross-table's shares, held to its counts
        out += typed_share_check(text)
    except Exception as e:
        _fault("typed shares", e)
    try:                                         # 6.6, round 170: which way a predictor goes, held to its coefficient
        out += coef_sign_check(text)
    except Exception as e:
        _fault("coefficient signs", e)
    try:                                         # 6.6, round 172: a correlation in words, held to an APA correlation table
        out += apa_corr_check(text)
    except Exception as e:
        _fault("APA correlations in words", e)
    return out, counted[0]


# ---------------------------------------------------------------- 6.6, round 69: a sample size worked out again
# (the owner, 21 Nov 2026 00:47: "Raise the app's validity. Do as much as a research tool can be.")
# Yamane's (1967) formula, n = N / (1 + N e²) (also called Slovin's), and Cochran's (1977), n₀ = Z² p (1 − p) / e² (also
# written Z²pq/d² and called Kish's or Fisher's), worked out again from the figures the writing gives them. Said only when
# the writing's own working never reaches what the formula gives (rounding of its steps allowed; a sample made larger for
# non-response after it, or Cochran's made smaller for a known population, stands), or when it states a sample smaller than
# Yamane's formula gives for the population and margin it names. Nothing is said where its figures aren't all there, or
# where the size came from somewhere else (Krejcie and Morgan's table, a power analysis, a calculator).
_NUMR = r"\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?|\.\d+"
_SIZE_YAMANE = re.compile(r"\bYamane\b|\bSlovin\b", re.I)
_SIZE_COCHRAN = re.compile(r"\bCochran\b|\bKish\b|\bFisher['’]?s?\s+(?:et\s+al\.?\s+)?formula\b", re.I)
_SIZE_ELSE = re.compile(r"\bKrejcie\b|\bG\s?\*\s?Power\b|\bpower\s+analysis\b|\bOpenEpi\b|\bRaosoft\b|\bEpi\s?Info\b|\bcalculator\b", re.I)
_SIZE_HEAD = re.compile(r"^\s*(?:(?:\d+(?:\.\d+)+\.?|\d+\.)\s+[A-Z][^=\n]{2,80}|[A-Z][A-Z ,&’'-]{4,80}|(?:[A-Z][a-z]+(?:\s+(?:of|and|the|in|for|to)\b)?\s+){1,7}[A-Z][a-z]+)(?<![.:;,])\s*$")
_SIZE_CALC = re.compile(r"^\s*(?:(?:n|s)\s*[₀0o]?\s*|\(?\s*n\s*\)?\s*)?(?:=|≈|≅)|^\s*(?:therefore|thus|hence|so),?\s+(?:the\s+)?(?:n\s*[₀0]?|sample\s+size)\b", re.I)
_EQ_NUM = re.compile(r"(?:=|≈|≅)\s*(?P<v>" + _NUMR + r")(?![\d.,]*\d)(?!\s*(?:[/*+^²(\[]|-\s*\d|x\s*\d))")
_YAM_SUB = re.compile(r"(?P<a>" + _NUMR + r")\s*/\s*[(\[]?\s*1\s*\+\s*(?P<b>" + _NUMR + r")\s*\*?\s*[(\[]?\s*(?P<e>\d*\.\d+|\d{1,2}(?:\.\d+)?\s*%)\s*[)\]]?"
                      r"\s*(?P<sq>²|\^\s*2|\*\*\s*2|2(?![\d.,]))?")
_Z_DEF = re.compile(r"(?<![A-Za-z])Z\s*(?:[(\[][^)\]=]{0,12}[)\]]\s*)?(?:=|is|of)\s*\(?\s*(?P<v>1\.645|1\.65|1\.96|2\.33|2\.326|2\.576|2\.58)(?!\d)")
_P_DEF = re.compile(r"(?<![A-Za-z])[pP]\s*(?:=|is)\s*(?P<v>\d*\.\d+|\d{1,2}(?:\.\d+)?\s*%)(?![\d.,]*\d)|"
                    r"\b(?:[Pp]roportion|[Pp]revalence)\s*(?:[(\[]\s*[pP]\s*[)\]]\s*)?(?:of|=|is|was|:)\s*(?P<w>\d*\.\d+|\d{1,2}(?:\.\d+)?\s*%)(?![\d.,]*\d)")
_E_DEF = re.compile(r"(?<![A-Za-z])(?P<l>e|d)\s*(?:=|is)\s*(?P<v>\d*\.\d+|\d{1,2}(?:\.\d+)?\s*%)(?![\d.,]*\d)|"
                    r"\b(?:margin\s+of\s+error|level\s+of\s+precision|precision\s+level|desired\s+precision|error\s+margin|degree\s+of\s+(?:precision|accuracy))"
                    r"\s*(?:[(\[]\s*(?:e|d)\s*[)\]]\s*)?(?:of|=|is|at|was|:)?\s*(?:\(\s*)?(?P<w>\d*\.\d+|\d{1,2}(?:\.\d+)?\s*%)(?![\d.,]*\d)", re.I)
_N_DEF = re.compile(r"(?<![A-Za-z])N\s*(?:=|is)\s*(?P<v>" + _NUMR + r")(?![\d.,]*\d)(?!\s*[/*+^(])")
_POP_DEF = re.compile(r"\b(?:target\s+|study\s+|accessible\s+)?population(?:\s+size)?\s*(?:[(\[]\s*N\s*[)\]]\s*)?(?:of|=|is|was|:)\s*(?:estimated\s+(?:to\s+be|at)\s+)?(?:about\s+|approximately\s+|around\s+)?"
                      r"(?P<v>" + _NUMR + r")(?![\d.,]*\d)", re.I)
_STATED = re.compile(r"\bsample\s+size\s*(?:[(\[]\s*n\s*[)\]]\s*)?(?:of|is|was|will\s+be|becomes|became|=|:)\s*(?:about\s+|approximately\s+)?(?P<v>" + _NUMR + r")(?![\d.,]*\d)"
                     r"|(?<![A-Za-z])n\s*=\s*(?P<w>" + _NUMR + r")(?![\d.,]*\d)(?!\s*(?:[/*+^²(\[]|-\s*\d))", re.I)


def _size_num(s):
    s = str(s).replace(",", "").strip()
    if s.endswith("%"):
        return float(s[:-1].strip()) / 100.0
    return float(s)


def _size_fmt(x):
    """A figure as a person writes it: 1,200; 0.05; 384.16."""
    if abs(x - round(x)) < 1e-9:
        return f"{int(round(x)):,}"
    return f"{x:,.2f}".rstrip("0").rstrip(".") if abs(x) >= 1 else f"{x:g}"


def _size_fmt4(x):
    return f"{x:,.4f}".rstrip("0").rstrip(".")


def _size_ok(v, exacts):
    """Whether a written result is one of the formula's values as a person rounds it (down, up, to the nearest, or to one
    or two decimals)."""
    for x in exacts:
        if x <= 0:
            continue
        if any(abs(v - c) < 1e-9 for c in (math.floor(x), math.ceil(x), round(x))) or abs(v - round(x, 1)) < 0.051 or abs(v - round(x, 2)) < 0.0051:
            return True
    return False


def _size_windows(text):
    """The passages that work out a sample size: from a heading (or 900 characters before the formula's name) to the next
    heading (or 1,500 characters after it)."""
    t = str(text or "")
    lines = t.split("\n")
    starts, pos = [], 0
    for ln in lines:
        starts.append(pos)
        pos += len(ln) + 1
    heads = [starts[i] for i, ln in enumerate(lines) if ln.strip() and _SIZE_HEAD.match(ln) and len(ln.strip()) <= 90]
    out = []
    for m in list(_SIZE_YAMANE.finditer(t)) + list(_SIZE_COCHRAN.finditer(t)):
        a = max([h for h in heads if h <= m.start()] + [max(0, m.start() - 900)])
        b = min([h for h in heads if h > m.start()] + [m.end() + 1500, len(t)])
        a = t.rfind("\n", 0, a) + 1 if a > 0 and t[a - 1] != "\n" else a
        out.append((a, b))
    out.sort()
    merged = []
    for a, b in out:
        if merged and a <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(b, merged[-1][1]))
        else:
            merged.append((a, b))
    return [t[a:b] for a, b in merged]


def _size_norm(w):
    return (w.replace("÷", "/").replace("×", "*").replace("·", "*").replace("−", "-").replace("–", "-").replace("[", "(").replace("]", ")")
            .replace("{", "(").replace("}", ")").replace("⁰", "0").replace("₀", "0").replace(" ", " ").replace(" ", " "))


def _size_chain(lines, i, start, used=None):
    """The working that follows a formula: the rest of its line and the lines after it that go on working ("= …", "n = …").
    used: the lines it takes are added to it (a step that writes the formula's figures again is the same working)."""
    chain = [lines[i][start:]]
    took = [i]
    for j in range(i + 1, min(len(lines), i + 12)):
        ln = lines[j]
        if not ln.strip():
            continue
        if _SIZE_CALC.match(ln):
            chain.append(ln)
            took.append(j)
        else:
            break
    if used is not None:
        used.update(took)
    return chain


def _size_results(chain, skip=()):
    """The figures the working arrives at (after "=" or "≈", not the start of another sum), 10 or more."""
    got = []
    for ln in chain:
        for m in _EQ_NUM.finditer(ln):
            try:
                v = _size_num(m.group("v"))
            except ValueError:
                continue
            if v >= 10 and not any(abs(v - s) < 1e-9 for s in skip):
                got.append((v, m.group("v")))
    return got


def _yamane_notes(w, orig):
    notes, n = [], 0
    name = "Slovin’s formula" if re.search(r"\bSlovin\b", w, re.I) and not re.search(r"\bYamane\b", w, re.I) else "Yamane’s formula"
    lines, shown = w.split("\n"), orig.split("\n")
    used = set()
    for i, ln in enumerate(lines):
        if i in used:
            continue
        for m in _YAM_SUB.finditer(ln):
            try:
                a, b, ev = _size_num(m.group("a")), _size_num(m.group("b")), _size_num(m.group("e"))
            except ValueError:
                continue
            if abs(a - b) > 1e-9 or a < 10 or a > 1e8:
                continue
            e = ev if (m.group("sq") or ev >= 0.01) else math.sqrt(ev)
            if not 0.005 <= e <= 0.2:
                continue
            N = a
            d = 1 + N * e * e
            exacts = [N / d, N / round(d, 2), N / round(d, 1), N / round(d, 3)]
            got = _size_results(_size_chain(lines, i, m.end(), used), skip=(N,))
            if not got:
                continue
            n += 1
            if any(_size_ok(v, exacts) for v, _ in got):
                break
            notes.append({"note": f"{name} with N = {_size_fmt(N)} and e = {_size_fmt(e)} gives {_size_fmt(N)} ÷ (1 + {_size_fmt(N)} × {_size_fmt(e)}²) = "
                                  f"{_size_fmt(N)} ÷ {_size_fmt4(d)} = {_size_fmt(round(N / d, 2))}" + (f" ({math.ceil(N / d - 1e-9):,} when rounded up)" if abs(N / d - round(N / d)) > 1e-6 else "")
                                  + f"; the working gives {got[-1][1]}.", "kind": "value", "decision": False, "s": shown[i].strip()[:500]})
            return notes, n
    if n or _SIZE_ELSE.search(w):
        return notes, n
    # no working written out: the population, the margin and the sample the writing names
    Ns = {_size_num(x.group("v")) for x in _N_DEF.finditer(w)} | {_size_num(x.group("v")) for x in _POP_DEF.finditer(w)}
    es = set()
    for x in _E_DEF.finditer(w):
        try:
            es.add(_size_num(x.group("v") or x.group("w")))
        except (ValueError, TypeError):
            pass
    stated = []
    for x in _STATED.finditer(w):
        try:
            v = _size_num(x.group("v") or x.group("w"))
        except (ValueError, TypeError):
            continue
        if v >= 10 and v not in Ns:
            stated.append((v, x.group("v") or x.group("w"), x.start()))
    if len(Ns) != 1 or len(es) != 1 or not stated:
        return notes, n
    N, e = next(iter(Ns)), next(iter(es))
    if not (10 <= N <= 1e8 and 0.005 <= e <= 0.2):
        return notes, n
    d = 1 + N * e * e
    low = math.floor(N / round(d, 1) if round(d, 1) > 0 else N / d)
    n += 1
    low = min(low, math.floor(N / d))
    if all(v < low for v, _, _ in stated):
        v, txt, at = stated[0]
        s0 = w.rfind("\n", 0, at) + 1
        s1 = w.find("\n", at)
        notes.append({"note": f"{name} with N = {_size_fmt(N)} and e = {_size_fmt(e)} gives {_size_fmt(N)} ÷ (1 + {_size_fmt(N)} × {_size_fmt(e)}²) = "
                              f"{_size_fmt(round(N / d, 2))}: a sample of {txt} is smaller than that.", "kind": "value", "decision": False,
                      "s": orig[s0:s1 if s1 > 0 else len(orig)].strip()[:500]})
    return notes, n


def _cochran_notes(w, orig):
    notes, n = [], 0
    name = ("Cochran’s formula" if re.search(r"\bCochran\b", w, re.I) else "Kish’s formula" if re.search(r"\bKish\b", w, re.I) else "Fisher’s formula")
    zs = {float(x.group("v")) for x in _Z_DEF.finditer(w)}
    if not zs and re.search(r"(?<![\d.])1\.96(?!\d)", w):
        zs = {1.96}
    ps = set()
    for x in _P_DEF.finditer(w):
        try:
            v = _size_num(x.group("v") or x.group("w"))
        except (ValueError, TypeError):
            continue
        if 0 < v < 1:
            ps.add(v)
    es = set()
    for x in _E_DEF.finditer(w):
        try:
            v = _size_num(x.group("v") or x.group("w"))
        except (ValueError, TypeError):
            continue
        if 0.005 <= v <= 0.2:
            es.add(v)
    if len(zs) != 1 or len(ps) != 1 or len(es) != 1:
        return notes, n
    z, p, e = next(iter(zs)), next(iter(ps)), next(iter(es))
    letter = "d" if any((x.group("l") or "").lower() == "d" for x in _E_DEF.finditer(w)) else "e"
    lines, shown = w.split("\n"), orig.split("\n")
    zw = re.compile(r"(?<![\d.])(?:" + re.escape(f"{z:g}") + r"|" + re.escape(f"{z * z:.4f}".rstrip("0")) + r"|" + re.escape(f"{z * z:.2f}") + r")(?!\d)")
    Ns = {_size_num(x.group("v")) for x in _N_DEF.finditer(w)} | {_size_num(x.group("v")) for x in _POP_DEF.finditer(w)}
    for i, ln in enumerate(lines):
        mz = zw.search(ln)
        if not mz or "/" not in ln[mz.start():] or not re.search(r"=|≈", ln):
            continue
        if re.match(r"\s*(?:where|z\s*=|z\s+is)\b", ln, re.I):
            continue
        q = p * (1 - p)
        z2s = {z * z, round(z * z, 2), round(z * z, 3), round(z * z, 4)}
        base = [zz * q / (e * e) for zz in z2s] + [zz * round(q, 2) / (e * e) for zz in z2s]
        exacts = list(base)
        for N in Ns:
            if N >= 10:
                exacts += [x / (1 + (x - 1) / N) for x in base] + [x / (1 + x / N) for x in base]
        got = _size_results(_size_chain(lines, i, mz.start()), skip=tuple(Ns))
        if not got:
            continue
        n += 1
        if any(_size_ok(v, exacts) for v, _ in got):
            return notes, n
        x = z * z * q / (e * e)
        notes.append({"note": f"{name} with Z = {z:g}, p = {_size_fmt(p)} and {letter} = {_size_fmt(e)} gives {z:g}² × {_size_fmt(p)} × {_size_fmt(round(1 - p, 4))} ÷ "
                              f"{_size_fmt(e)}² = {_size_fmt(round(x, 2))}" + (f" ({math.ceil(x - 1e-9):,} when rounded up)" if abs(x - round(x)) > 1e-6 else "")
                              + f"; the working gives {got[-1][1]}.", "kind": "value", "decision": False, "s": shown[i].strip()[:500]})
        return notes, n
    return notes, n


# 6.6, round 139 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): Krejcie and Morgan's (1970) table, worked out
# again. Its sizes come from one formula, s = χ²NP(1 − P) / (d²(N − 1) + χ²P(1 − P)) with χ² = 3.841, P = .50, d = .05; a
# writer may take the size of the population the table lists just above or just below theirs. A stated sample smaller than all
# of these is said (a larger one stands: a sample made larger for non-response is right).
_KM_N = (10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 90, 95, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200,
         210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 320, 340, 360, 380, 400, 420, 440, 460, 480, 500, 550, 600, 650, 700, 750, 800,
         850, 900, 950, 1000, 1100, 1200, 1300, 1400, 1500, 1600, 1700, 1800, 1900, 2000, 2200, 2400, 2600, 2800, 3000, 3500, 4000, 4500,
         5000, 6000, 7000, 8000, 9000, 10000, 15000, 20000, 30000, 40000, 50000, 75000, 1000000)
_KM_NAME = re.compile(r"\bKrejcie\b", re.I)
_KM_SAMPLE = re.compile(r"\bsample\s+(?:size\s+)?(?:\(\s*n\s*\)\s*)?(?:of|=|is|was|will\s+be|:)\s*(?:about\s+|approximately\s+)?(?P<v>" + _NUMR + r")(?![\d.,]*\d)"
                        r"|(?<![A-Za-z])n\s*=\s*(?P<w>" + _NUMR + r")(?![\d.,]*\d)|(?P<x>" + _NUMR + r")\s+(?:respondents|participants|students|nurses)\s+"
                        r"(?:were|was|will\s+be)\s+(?:selected|sampled|chosen|used|recruited)", re.I)


def _km_size(N):
    """Krejcie and Morgan's (1970) sample size for a population of N."""
    x2p = 3.841 * 0.25
    return x2p * N / (0.0025 * (N - 1) + x2p)


def _krejcie_notes(text):
    t = str(text or "")
    notes, n = [], 0
    seen = set()
    for m in _KM_NAME.finditer(t):
        a = max(t.rfind("\n\n", 0, m.start()), t.rfind("\n", 0, max(0, m.start() - 600)), 0)
        b = t.find("\n\n", m.end())
        b = len(t) if b < 0 else min(b, m.end() + 700)
        w = _size_norm(t[a:b])
        if (a, b) in seen:
            continue
        seen.add((a, b))
        pm = _POP_DEF.search(w) or _N_DEF.search(w)
        sm = next((x for x in _KM_SAMPLE.finditer(w) if not (pm and x.start() <= pm.start() < x.end())), None)
        if not pm or not sm:
            continue
        try:
            N = _size_num(pm.group("v"))
            s_t = next(g for g in (sm.group("v"), sm.group("w"), sm.group("x")) if g)
            s = _size_num(s_t)
        except (ValueError, StopIteration, IndexError):
            continue
        if N < 10 or N != int(N) or s <= 0 or s >= N:
            continue
        n += 1
        exact = _km_size(N)
        below = max([x for x in _KM_N if x <= N] or [10])
        above = min([x for x in _KM_N if x >= N] or [1000000])
        allowed = [math.floor(exact), math.ceil(exact), round(_km_size(below)), round(_km_size(above)), math.floor(_km_size(below))]
        if s < min(allowed) - 1e-9:
            sent = next((x for x in sentences(t[a:b]) if s_t in x), t[a:b][:300])
            notes.append({"s": sent[:500], "kind": "size", "decision": False,
                          "note": f"Krejcie and Morgan’s table gives {round(exact):,} for a population of {int(N):,} (its formula: χ² = 3.841, "
                                  f"P = .50, d = .05), not {s_t}: a smaller sample than the table asks."})
    return notes, n


def sample_check(text):
    """The sample sizes the writing works out, worked out again: ([{"s", "note", "kind", "decision"}], how many were)."""
    notes, n = [], 0
    for orig in _size_windows(text):
        w = _size_norm(orig)
        for f, kw in ((_yamane_notes, _SIZE_YAMANE), (_cochran_notes, _SIZE_COCHRAN)):
            if not kw.search(w):
                continue
            try:
                got, k = f(w, orig)
            except (ValueError, ZeroDivisionError, OverflowError):
                continue
            notes += got
            n += k
    try:                                         # 6.6, round 139: Krejcie and Morgan's table
        got, k = _krejcie_notes(text)
        notes += got
        n += k
    except Exception as e:
        _fault("Krejcie and Morgan", e)
    return notes, n
