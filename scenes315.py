"""3.15 — Slide Studio's designs: each its own look, and Morph designs whose shapes move from slide to slide.

The user: "I don't want the templates after the first 12 to look the same. Add many more templates, too." and "add
more than 30 extremely cool Morph transitions ... I don't want them to look the same."

The first 12 (the classic student designs) stay as they are. The 21 after them that were one layout in many colours
each get their own shapes and their own layout (where the title goes, where the words go, what's on the cover, a
section slide's look); 17 new designs come for students and tutors alike; and 33 Morph designs: named shapes that move,
turn, grow and change colour from one slide to the next (PowerPoint's Morph matches shapes named "!!..."; the app's
slide show does the same).

One source for both: the server draws these shapes into the .pptx (slidestudio), the app draws the same shapes on its
slides (written into the app at build time). A shape:

    {"k": kind, "x", "y", "w", "h" (inches, on a 13.333 x 7.5 slide), "f": fill (hex), "a": opacity, "r": turn
     (degrees, clockwise), "rr": corner radius (inches), "g": [hex, hex, angle] a gradient (angle anticlockwise from →),
     "l": [hex, points] an outline, "th": a ring's thickness (inches), "n": its name (Morph), "tx"/"ts"/"tc"/"tb": words
     in it (size in points, colour, bold)}

kinds: rect, oval, tri (point up), rtri (right angle bottom left), para, hex, diam, ring, star, plus.

audit() reads every design the way a person would: each place words go, against whatever is behind it there, at AAA
(7:1; 4.5:1 for large words).
"""
import math

W, H = 13.333, 7.5


# ---------------------------------------------------------------- colours
def _hex(c):
    return c.lstrip('#').upper()


def mix(a, b, t):
    """a → b by t (0 … 1)."""
    a, b = _hex(a), _hex(b)
    pa = [int(a[i:i + 2], 16) for i in (0, 2, 4)]
    pb = [int(b[i:i + 2], 16) for i in (0, 2, 4)]
    return ''.join('%02X' % round(pa[i] + (pb[i] - pa[i]) * t) for i in range(3))


def lum(c):
    c = _hex(c)
    v = [int(c[i:i + 2], 16) / 255 for i in (0, 2, 4)]
    v = [x / 12.92 if x <= 0.03928 else ((x + 0.055) / 1.055) ** 2.4 for x in v]
    return 0.2126 * v[0] + 0.7152 * v[1] + 0.0722 * v[2]


def contrast(a, b):
    la, lb = sorted((lum(a), lum(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def readable(c, on, need=7.0):
    """c made darker (or lighter) until it reads on `on`."""
    c = _hex(c)
    if contrast(c, on) >= need:
        return c
    dark = lum(on) > 0.18
    for k in range(1, 41):
        cc = mix(c, '000000' if dark else 'FFFFFF', k / 40)
        if contrast(cc, on) >= need:
            return cc
    return '000000' if dark else 'FFFFFF'


class P(dict):
    """A palette by role: ink (dark: titles, panels), ink2, accent, accent2, soft (a light tint), bg, and the colours
    words take: text (points, 12:1), emph (highlighted words, 7:1), muted (small lines, 7:1), onink (words on ink)."""
    def __init__(self, ink, accent, soft, bg='FFFFFF', ink2=None, emph=None, text=None, muted=None, accent2=None):
        ink, accent, soft, bg = _hex(ink), _hex(accent), _hex(soft), _hex(bg)
        ink = readable(ink, bg, 7.0)
        ink2 = readable(ink2 or mix(ink, accent, 0.4), 'FFFFFF', 4.6)
        super().__init__(ink=ink, ink2=ink2, accent=accent, accent2=_hex(accent2 or mix(accent, soft, 0.5)),
                         soft=soft, bg=bg, emph=readable(emph or accent, bg, 7.0), text=readable(text or '1F2937', bg, 12.0),
                         muted=readable(muted or '4B5563', bg, 7.0), white='FFFFFF', tint=mix(soft, 'FFFFFF', 0.55),
                         deep=mix(ink, '000000', 0.3), onink=readable(mix(soft, 'FFFFFF', 0.4), ink2, 7.0))


def C(p, c):
    """A colour by its role in the palette (or a hex)."""
    return p[c] if c in p else _hex(c)


# ---------------------------------------------------------------- shapes
def sh(k, x, y, w, h, f, a=None, r=None, rr=None, g=None, l=None, th=None, n=None, tx=None, ts=None, tc=None, tb=None):
    d = {'k': k, 'x': round(x, 3), 'y': round(y, 3), 'w': round(max(0.01, w), 3), 'h': round(max(0.01, h), 3), 'f': f}
    for key, v in (('a', a), ('r', r), ('rr', rr), ('g', g), ('l', l), ('th', th), ('n', n), ('tx', tx), ('ts', ts), ('tc', tc), ('tb', tb)):
        if v is not None:
            d[key] = round(v, 3) if isinstance(v, float) else v
    return d


def R(*a, **o): return sh('rect', *a, **o)
def O(*a, **o): return sh('oval', *a, **o)
def Tri(*a, **o): return sh('tri', *a, **o)
def RT(*a, **o): return sh('rtri', *a, **o)
def Hx(*a, **o): return sh('hex', *a, **o)
def Dm(*a, **o): return sh('diam', *a, **o)
def Rg(*a, **o): return sh('ring', *a, **o)
def St(*a, **o): return sh('star', *a, **o)
def Pl(*a, **o): return sh('plus', *a, **o)
def Pa(*a, **o): return sh('para', *a, **o)


def circ(cx, cy, d, f, **o):
    return O(cx - d / 2, cy - d / 2, d, d, f, **o)


def full(f, **o):
    return R(0, 0, W, H, f, **o)


def auto_tc(fill, p):
    """Words inside a shape: the first of white, ink, black that reads on it at 4.5:1 (they're large)."""
    for c in ('FFFFFF', p['ink'], p['deep'], '000000'):
        if contrast(c, fill) >= 4.5:
            return c
    return max(('FFFFFF', '000000'), key=lambda c: contrast(c, fill))


def resolve(shapes, p):
    out = []
    for s in shapes:
        s = dict(s)
        s['f'] = C(p, s['f'])
        if 'g' in s:
            s['g'] = [C(p, s['g'][0]), C(p, s['g'][1]), s['g'][2]]
        if 'l' in s:
            s['l'] = [C(p, s['l'][0]), s['l'][1]]
        if 'tx' in s:
            s['tc'] = auto_tc(s['f'], p)
        else:
            s.pop('tc', None)
        out.append(s)
    return out


# ---------------------------------------------------------------- where the words go
def G(title=(0.9, 0.5, 11.5, 0.95), body=(0.9, 1.75, 11.5, 4.85), tc=None):
    g = {'title': list(title), 'body': list(body), 'foot': [body[0], 6.9, min(9.5, body[2]), 0.36],
         'quote': [body[0] + 0.6, body[1] - 0.15, body[2] - 1.2, min(body[3], 4.6)]}
    if tc:
        g['tc'] = tc
    return g


def CG(t, s, m, tc='ink', sc='text', mc='muted', al='left', pic=None, pk='round', pr=0.25, dark=False, tp=None):
    """A cover: the title (tp: its size, if larger than usual), the line under it, the small line; where a picture
    goes (and its shape); their colours."""
    g = {'t': list(t), 's': list(s), 'm': list(m), 'tc': tc, 'sc': sc, 'mc': mc, 'al': al, 'dark': dark,
         'p': list(pic or (8.4, 1.0, 4.2, 5.4)), 'pk': pk, 'pr': pr}
    if tp:
        g['tp'] = tp
    return g


def DARK(t, s, m, **o):
    return CG(t, s, m, tc='white', sc='onink', mc='onink', dark=True, **o)


def SG(t, s, tc='white', sc='onink', al='left', dark=True):
    return {'t': list(t), 's': list(s), 'tc': tc, 'sc': sc, 'al': al, 'dark': dark}


def SGL(t, s, al='left'):
    """A section slide whose words sit on the light colour."""
    return SG(t, s, tc='ink', sc='text', al=al, dark=False)


def LEFT(w=7.0, y=1.4):
    """The usual cover: words on the left, large; the picture (if any) on the right."""
    return (0.9, y, w, 2.5), (0.9, y + 2.6, w, 0.7), (0.9, y + 3.3, w, 0.5)


# ======================================================================== designs
# Each: its cover, an ordinary slide, a section (and the end) as shapes; where the words go; where a cover picture goes.

def m_waves(p):
    cover = [full('bg'), O(-2.5, 5.0, 10, 5.2, 'soft'), O(3.8, 5.5, 12, 5.6, 'accent', a=0.85), O(-1.2, 6.2, 9.5, 4.2, 'ink'),
             O(8.6, 5.9, 7.5, 4.5, 'ink2', a=0.9), circ(12.1, 0.9, 0.9, 'accent', a=0.9)]
    content = [full('bg'), O(10.6, -2.5, 4.6, 3.6, 'soft', a=0.85), O(11.7, -1.95, 3.4, 2.7, 'accent', a=0.55)]
    section = [full('ink'), O(-3, 5.0, 11, 5.6, 'ink2'), O(4.5, 5.6, 12, 5.2, 'accent', a=0.9), O(-1.5, 6.4, 9, 3.6, 'soft', a=0.35)]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 9.6, 0.95), body=(0.9, 1.75, 11.5, 4.6)),
                cg=CG(*LEFT(7.4, 0.8), pic=(8.7, 0.8, 3.9, 3.9), pk='circle'), sg=SG((0.9, 0.9, 10.5, 2.3), (0.9, 3.25, 10.5, 0.8)))


def m_rings(p):
    def rings(cx, cy, k=1.0):
        return [Rg(cx - 3.6 * k, cy - 3.6 * k, 7.2 * k, 7.2 * k, 'accent', a=0.35, th=0.14 * k), Rg(cx - 2.6 * k, cy - 2.6 * k, 5.2 * k, 5.2 * k, 'soft', a=0.5, th=0.12 * k),
                Rg(cx - 1.6 * k, cy - 1.6 * k, 3.2 * k, 3.2 * k, 'accent', a=0.8, th=0.1 * k), circ(cx, cy, 1.2 * k, 'accent')]
    cover = [full('ink')] + rings(10.6, 3.75)
    content = [full('bg')] + rings(13.0, 0.2, 0.4)
    section = [full('ink')] + rings(1.6, 6.4, 0.8)
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.4, 0.95)),
                cg=DARK(*LEFT(5.9), pic=(9.0, 2.15, 3.2, 3.2), pk='circle'), sg=SG((4.6, 2.0, 7.9, 2.2), (4.6, 4.3, 7.9, 0.8)))


def m_sidebar(p):
    cover = [full('bg'), R(0, 0, 4.6, H, 'ink'), R(4.6, 0, 0.18, H, 'accent'), circ(2.3, 5.9, 2.2, 'ink2'), circ(2.3, 5.9, 1.0, 'accent', a=0.9)]
    content = [full('bg'), R(0, 0, 0.42, H, 'ink'), R(0.42, 0, 0.08, H, 'accent')]
    section = [full('bg'), R(0, 0, 6.2, H, 'ink'), R(6.2, 0, 0.18, H, 'accent'), circ(10.0, 3.75, 3.4, 'soft', a=0.7)]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(1.1, 0.5, 11.3, 0.95), body=(1.1, 1.75, 11.3, 4.85)),
                cg=CG((5.4, 1.5, 7.2, 2.5), (5.4, 4.1, 7.2, 0.7), (5.4, 4.8, 7.2, 0.5), pic=(0.5, 0.6, 3.6, 3.6), pr=0.3),
                sg=SG((0.8, 1.9, 5.0, 2.3), (0.8, 4.3, 5.0, 0.9)))


def m_banner(p):
    cover = [full('bg'), R(0, 0, W, 4.4, 'ink', g=['ink', 'deep', 0]), R(0, 4.4, W, 0.12, 'accent'), R(9.4, 0.75, 3.1, 3.1, 'accent', a=0.3, rr=0.4, r=12)]
    content = [full('bg'), R(0, 0, W, 1.45, 'ink'), R(0, 1.45, W, 0.07, 'accent')]
    section = [full('bg'), R(0, 2.2, W, 3.1, 'ink'), R(0, 5.3, W, 0.1, 'accent')]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.25, 11.5, 0.95), body=(0.9, 1.9, 11.5, 4.7), tc='white'),
                cg=CG((0.9, 1.2, 8.0, 2.6), (0.9, 4.75, 8.0, 0.7), (0.9, 5.45, 8.0, 0.5), tc='white', pic=(9.2, 0.75, 3.4, 3.4), pr=0.3),
                sg=SG((0.9, 2.45, 11.5, 1.8), (0.9, 4.25, 11.5, 0.8)))


def m_frame(p):
    cover = [full('bg'), R(0.35, 0.35, W - 0.7, H - 0.7, 'bg', l=['ink', 2.25]), R(0.55, 0.55, W - 1.1, H - 1.1, 'bg', l=['accent', 1]),
             R(5.67, 5.6, 2.0, 0.08, 'accent')]
    content = [full('bg'), R(0.3, 0.3, W - 0.6, H - 0.6, 'bg', l=['soft', 1.5]), R(0.9, 1.45, 1.4, 0.06, 'accent')]
    section = [full('ink'), R(0.45, 0.45, W - 0.9, H - 0.9, 'ink', l=['accent', 1.5])]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.5, 11.5, 0.9), body=(0.9, 1.75, 11.5, 4.7)),
                cg=CG((1.4, 1.5, 10.5, 2.4), (1.4, 4.0, 10.5, 0.7), (1.4, 4.75, 10.5, 0.5), al='center', pic=(8.6, 1.2, 3.6, 5.0), pr=0.0),
                sg=SG((1.4, 2.0, 10.5, 2.0), (1.4, 4.1, 10.5, 0.8), al='center'))


def m_diagonal(p):
    cover = [full('bg'), R(0, 0, 5.0, H, 'ink'), Pa(3.0, 0, 4.0, H, 'ink'), Pa(6.3, 0, 2.0, H, 'accent', a=0.95), circ(11.3, 6.3, 2.6, 'soft', a=0.8)]
    content = [full('bg'), Pa(11.4, 0, 2.6, 0.5, 'accent'), Pa(12.3, 0, 1.4, 0.5, 'ink')]
    section = [full('bg'), R(0, 0, 6.0, H, 'ink'), Pa(4.0, 0, 4.0, H, 'ink'), Pa(7.4, 0, 1.6, H, 'accent')]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.65, 10.0, 0.95)),
                cg=DARK((0.7, 0.8, 5.5, 2.4), (0.7, 3.3, 5.4, 0.7), (0.7, 4.0, 5.3, 0.5), pic=(8.6, 0.7, 4.1, 3.6), pr=0.2),
                sg=SG((0.8, 2.0, 5.6, 2.2), (0.8, 4.3, 5.4, 0.8)))


def m_bubbles(p):
    dots = [circ(10.8, 1.7, 2.6, 'accent', a=0.85), circ(12.5, 3.5, 1.3, 'soft'), circ(9.6, 4.6, 1.0, 'ink2', a=0.8), circ(11.5, 5.7, 2.0, 'soft', a=0.8),
            circ(9.0, 2.6, 0.55, 'accent'), circ(12.7, 0.75, 0.7, 'ink2', a=0.7)]
    cover = [full('bg')] + dots
    content = [full('bg'), circ(12.75, 0.4, 1.4, 'accent', a=0.75), circ(11.85, 0.3, 0.55, 'soft'), circ(12.35, 1.35, 0.4, 'ink2', a=0.6)]
    section = [full('ink')] + [dict(d, f='accent' if i % 2 else 'soft', a=0.55) for i, d in enumerate(dots)]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.4, 0.95)),
                cg=CG(*LEFT(7.4), pic=(9.2, 0.9, 3.4, 3.4), pk='circle'), sg=SG((0.9, 1.9, 7.4, 2.3), (0.9, 4.3, 7.4, 0.8)))


def m_stripes(p):
    def st(x0, n, w, f, a):
        return [R(x0 + i * w * 2, -1, w, H + 2, f, a=a, r=18) for i in range(n)]
    cover = [full('ink')] + st(7.6, 4, 0.55, 'accent', 0.9) + st(8.15, 4, 0.55, 'ink2', 0.9) + [R(0, 6.7, W, 0.8, 'accent')]
    content = [full('bg')] + [R(12.0 + i * 0.3, -0.2, 0.15, 1.6, 'accent' if i % 2 == 0 else 'soft', r=18) for i in range(5)]
    section = [full('bg')] + st(-1.2, 3, 0.6, 'ink', 1) + st(-0.6, 3, 0.6, 'accent', 0.85)
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.8, 0.95)),
                cg=DARK(*LEFT(6.0, 1.1), pic=(8.0, 0.9, 4.4, 5.2), pr=0.2),
                sg=SGL((6.0, 2.0, 6.6, 2.2), (6.0, 4.3, 6.6, 0.8)))


def m_blocks(p):
    cover = [full('bg'), R(8.0, 0, 5.333, 3.75, 'ink'), R(8.0, 3.75, 2.67, 3.75, 'accent'), R(10.67, 3.75, 2.67, 3.75, 'soft'), circ(9.33, 1.9, 1.5, 'accent2')]
    content = [full('bg'), R(12.2, 0, 1.133, 0.6, 'ink'), R(12.75, 0.6, 0.583, 0.6, 'accent'), R(12.2, 0.6, 0.55, 0.6, 'soft')]
    section = [full('bg'), R(0, 0, 6.67, H, 'ink'), R(6.67, 0, 3.3, 3.75, 'accent'), R(6.67, 3.75, 3.3, 3.75, 'soft')]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.8, 0.95)),
                cg=CG(*LEFT(6.6, 1.3), pic=(8.0, 0, 5.333, 3.75), pr=0.0), sg=SG((0.8, 2.0, 5.5, 2.2), (0.8, 4.3, 5.5, 0.9)))


def m_arch(p):
    cover = [full('bg'), R(7.6, 1.2, 5.0, 6.3, 'soft', rr=2.5), R(7.6, 3.7, 5.0, 3.8, 'soft'), R(8.3, 2.1, 3.6, 5.4, 'accent', rr=1.8),
             R(8.3, 3.9, 3.6, 3.6, 'accent'), circ(10.1, 0.75, 0.6, 'ink')]
    content = [full('bg'), R(12.1, -0.8, 1.0, 2.4, 'soft', rr=0.5), R(12.5, -0.8, 0.55, 1.75, 'accent', rr=0.27)]
    section = [full('ink'), R(1.0, 1.4, 4.4, 6.1, 'ink2', rr=2.2), R(1.0, 3.6, 4.4, 3.9, 'ink2'), R(1.6, 2.4, 3.2, 5.1, 'accent', rr=1.6), R(1.6, 4.0, 3.2, 3.5, 'accent')]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.9, 0.95)),
                cg=CG(*LEFT(6.3, 1.3), pic=(8.3, 2.1, 3.6, 5.4), pk='arch'), sg=SG((6.2, 2.0, 6.4, 2.2), (6.2, 4.3, 6.4, 0.8)))


def m_dots(p):
    def grid(x0, y0, cols, rows, step, d, f, a):
        return [circ(x0 + c * step, y0 + r * step, d, f, a=a) for r in range(rows) for c in range(cols)]
    cover = [full('bg')] + grid(8.6, 1.0, 7, 9, 0.62, 0.16, 'accent', 0.75) + [circ(11.0, 3.6, 3.2, 'ink', a=0.92)]
    content = [full('bg')] + grid(11.8, 0.35, 4, 3, 0.38, 0.1, 'accent', 0.8)
    section = [full('ink')] + grid(0.9, 0.7, 19, 2, 0.64, 0.14, 'soft', 0.5)
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.6, 0.95)),
                cg=CG(*LEFT(7.2), pic=(9.4, 2.0, 3.2, 3.2), pk='circle'), sg=SG((0.9, 2.6, 11.5, 2.2), (0.9, 4.9, 11.5, 0.8)))


def m_hexes(p):
    def hx(cx, cy, s, f, a=None):
        return Hx(cx - s / 2, cy - s * 0.433, s, s * 0.866, f, a=a)

    def cluster(x, y, s):
        return [hx(x, y, s, 'accent'), hx(x + s * 0.76, y + s * 0.44, s, 'soft'), hx(x, y + s * 0.88, s, 'ink2', 0.85),
                hx(x + s * 0.76, y - s * 0.44, s, 'ink', 0.9), hx(x - s * 0.76, y + s * 0.44, s, 'soft', 0.7)]
    cover = [full('bg')] + cluster(10.6, 3.2, 1.9)
    content = [full('bg')] + cluster(12.45, 0.4, 0.55)
    section = [full('ink')] + cluster(2.4, 3.2, 1.6)
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.5, 0.95)),
                cg=CG(*LEFT(7.0), pic=(8.6, 1.4, 4.0, 3.46), pk='hex'), sg=SG((5.4, 2.0, 7.2, 2.2), (5.4, 4.3, 7.2, 0.8)))


def m_notebook(p):
    def lines(y0, n, x0=0, w=W, f='soft', a=0.9):
        return [R(x0, y0 + i * 0.5, w, 0.025, f, a=a) for i in range(n)]
    cover = [full('bg')] + lines(1.0, 13) + [R(1.5, 0, 0.04, H, 'accent'), R(1.6, 0, 0.02, H, 'accent', a=0.6)] + \
        [circ(0.6, 0.9 + i * 1.4, 0.36, 'ink') for i in range(5)]
    content = [full('bg')] + lines(1.55, 11, 0.0, W, 'soft', 0.75) + [R(0.62, 0, 0.035, H, 'accent', a=0.85)]
    section = [full('ink')] + [circ(0.6, 0.9 + i * 1.4, 0.36, 'bg') for i in range(5)] + [R(1.4, 0, 0.05, H, 'accent')]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.95, 0.5, 11.5, 0.95), body=(0.95, 1.72, 11.5, 4.9)),
                cg=CG((2.1, 1.3, 9.8, 2.5), (2.1, 3.95, 9.8, 0.7), (2.1, 4.65, 9.8, 0.5), sc=readable(p['text'], p['soft'], 7.0), mc=readable(p['muted'], p['soft'], 7.0),
                      pic=(8.6, 1.1, 3.8, 4.6), pr=0.1),
                sg=SG((2.1, 2.0, 10.0, 2.2), (2.1, 4.3, 10.0, 0.8)))


def m_grid(p):
    def gl(x0, y0, w, h, step, f, a):
        return [R(x0 + i * step, y0, 0.012, h, f, a=a) for i in range(int(w / step) + 1)] + [R(x0, y0 + j * step, w, 0.012, f, a=a) for j in range(int(h / step) + 1)]
    cover = [full('ink')] + gl(0, 0, W, H, 0.5, 'soft', 0.22) + [Rg(8.6, 1.3, 4.0, 4.0, 'accent', th=0.05), R(10.585, 0.9, 0.03, 4.8, 'accent'), R(8.2, 3.285, 4.8, 0.03, 'accent')]
    content = [full('bg')] + gl(11.1, 0.3, 1.9, 1.15, 0.26, 'accent', 0.35)
    section = [full('ink')] + gl(0, 0, W, H, 1.0, 'soft', 0.2)
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.0, 0.95)),
                cg=DARK(*LEFT(7.2), pic=(8.85, 1.55, 3.5, 3.5), pk='circle'), sg=SG((0.9, 2.0, 11.0, 2.2), (0.9, 4.3, 11.0, 0.8)))


def m_tabs(p):
    cover = [full('bg'), R(0.9, 0, 1.2, 1.5, 'accent', rr=0.15), R(2.3, 0, 1.2, 1.1, 'soft', rr=0.15), R(3.7, 0, 1.2, 0.8, 'ink2', rr=0.15),
             R(0, 6.2, W, 1.3, 'ink'), R(0, 6.12, W, 0.08, 'accent')]
    content = [full('bg'), R(0.9, 0, 0.9, 0.34, 'accent', rr=0.1), R(1.95, 0, 0.9, 0.22, 'soft', rr=0.1)]
    section = [full('ink'), R(0.9, 0, 1.6, 1.6, 'accent', rr=0.2), R(2.7, 0, 1.6, 1.2, 'soft', rr=0.2), R(4.5, 0, 1.6, 0.9, 'ink2', rr=0.2)]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.6, 11.5, 0.95)),
                cg=CG((0.9, 1.9, 11.0, 2.4), (0.9, 4.35, 11.0, 0.7), (0.9, 6.45, 11.0, 0.5), mc='onink', pic=(8.6, 1.7, 3.8, 4.0), pr=0.2),
                sg=SG((0.9, 2.6, 11.0, 2.0), (0.9, 4.7, 11.0, 0.8)))


def m_halo(p):
    cover = [full('ink'), circ(W / 2, 9.4, 11, 'ink2'), circ(W / 2, 9.4, 8.2, 'accent', a=0.85), circ(W / 2, 9.4, 5.4, 'soft', a=0.9)]
    content = [full('bg'), circ(W / 2, -2.75, 4.4, 'soft', a=0.8), circ(W / 2, -2.75, 3.0, 'accent', a=0.5)]
    section = [full('ink'), circ(W / 2, 3.75, 6.8, 'ink2'), Rg(W / 2 - 3.4, 0.35, 6.8, 6.8, 'accent', th=0.08)]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(1.4, 0.6, 10.5, 0.95), body=(0.9, 1.75, 11.5, 4.85)),
                cg=DARK((1.2, 0.6, 10.9, 2.0), (1.2, 2.65, 10.9, 0.65), (1.2, 3.25, 10.9, 0.45), al='center', pic=(5.27, 4.0, 2.8, 2.8), pk='circle'),
                sg=SG((3.9, 2.5, 5.5, 1.6), (3.9, 4.1, 5.5, 0.7), al='center'))


def m_ribbon(p):
    cover = [full('bg'), R(-2, 5.0, 18, 1.2, 'accent', r=-6), R(-2, 5.8, 18, 0.45, 'ink', r=-6), R(-2, 0.5, 18, 0.16, 'soft', r=-6)]
    content = [full('bg'), R(10.8, 0.18, 3.4, 0.42, 'accent', r=-6), R(11.1, 0.58, 3.1, 0.14, 'ink', r=-6)]
    section = [full('ink'), R(-2, 5.3, 18, 0.95, 'accent', r=-6), R(-2, 6.05, 18, 0.32, 'soft', r=-6)]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.75, 9.8, 0.95), body=(0.9, 1.9, 11.5, 4.6)),
                cg=CG((0.9, 1.15, 10.5, 2.2), (0.9, 3.4, 8.5, 0.7), (0.9, 4.1, 8.5, 0.5), pic=(9.0, 0.9, 3.6, 3.4), pr=0.2),
                sg=SG((0.9, 1.3, 11.0, 2.2), (0.9, 3.5, 11.0, 0.8)))


def m_polaroid(p):
    cover = [full('soft'), R(7.0, 1.6, 4.4, 5.1, 'white', r=-5, a=0.7), R(7.6, 0.9, 4.6, 5.4, 'white', r=4), R(7.95, 1.25, 3.9, 3.9, 'accent', r=4, a=0.85)]
    content = [full('bg'), R(12.0, 0.25, 1.0, 1.15, 'soft', r=6), R(12.12, 0.36, 0.76, 0.76, 'accent', r=6, a=0.7)]
    section = [full('ink'), R(1.0, 1.2, 4.2, 4.9, 'white', r=-5), R(1.3, 1.5, 3.6, 3.6, 'accent', r=-5)]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.7, 0.95)),
                cg=CG(*LEFT(6.0), tc=readable(p['ink'], p['soft'], 7.0), sc=readable(p['text'], p['soft'], 7.0), mc=readable(p['muted'], p['soft'], 7.0),
                      pic=(7.95, 1.25, 3.9, 3.9), pr=0.0),
                sg=SG((6.0, 2.0, 6.6, 2.2), (6.0, 4.3, 6.6, 0.8)))


def m_bauhaus(p):
    cover = [full('bg'), circ(10.0, 2.2, 3.2, 'accent'), R(7.9, 3.7, 2.6, 2.6, 'ink'), Tri(10.4, 4.2, 2.6, 2.3, 'soft'), R(11.7, 0.45, 1.1, 1.1, 'ink2', r=45)]
    content = [full('bg'), circ(12.55, 0.6, 0.85, 'accent'), R(11.6, 0.28, 0.5, 0.5, 'ink'), Tri(12.05, 1.0, 0.55, 0.5, 'soft')]
    section = [full('ink'), circ(2.6, 3.75, 3.4, 'accent'), R(4.3, 4.2, 2.2, 2.2, 'soft'), Tri(1.0, 0.6, 2.0, 1.8, 'ink2')]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.5, 0.95)),
                cg=CG(*LEFT(6.6), pic=(8.4, 0.6, 3.2, 3.2), pk='circle'), sg=SG((7.2, 2.0, 5.4, 2.2), (7.2, 4.3, 5.4, 0.8)))


def m_memphis(p):
    conf = [St(8.7, 0.8, 0.7, 0.7, 'accent', r=12), Pl(11.7, 1.2, 0.6, 0.6, 'ink2', r=20), circ(12.4, 4.0, 0.5, 'soft'), Tri(9.1, 4.9, 0.7, 0.6, 'ink', r=-18),
            R(10.5, 6.0, 0.9, 0.22, 'accent', r=30, rr=0.11), circ(8.5, 3.0, 0.35, 'ink'), St(12.0, 5.9, 0.5, 0.5, 'soft', r=-8), R(11.0, 2.6, 0.7, 0.7, 'accent', a=0.4, r=20)]
    cover = [full('bg'), circ(10.5, 3.3, 2.6, 'soft', a=0.55)] + conf
    content = [full('bg'), St(12.35, 0.35, 0.45, 0.45, 'accent', r=12), Pl(11.7, 0.55, 0.35, 0.35, 'ink2', r=20), circ(12.8, 1.15, 0.22, 'soft')]
    section = [full('ink')] + [dict(c, x=c['x'] - 7.6) for c in conf]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.5, 0.95)),
                cg=CG(*LEFT(7.0), pic=(9.2, 2.0, 2.6, 2.6), pk='circle'), sg=SG((6.6, 2.0, 6.0, 2.2), (6.6, 4.3, 6.0, 0.8)))


def m_leaves(p):
    def leaf(x, y, w, h, r, f, a=None):
        return O(x, y, w, h, f, r=r, a=a)
    cover = [full('bg'), leaf(9.6, -0.8, 1.6, 4.2, 35, 'accent'), leaf(11.2, 0.2, 1.3, 3.6, 65, 'ink2', 0.9), leaf(10.4, 3.2, 1.4, 3.8, -30, 'soft'),
             leaf(8.9, 4.5, 1.1, 3.0, 20, 'accent', 0.7), leaf(-0.4, 5.6, 1.4, 3.4, -55, 'soft', 0.8)]
    content = [full('bg'), leaf(12.25, -0.6, 0.65, 1.8, 35, 'accent', 0.85), leaf(12.75, -0.2, 0.5, 1.4, 70, 'soft')]
    section = [full('ink'), leaf(9.0, 0.4, 2.0, 5.2, 30, 'accent', 0.85), leaf(10.8, 1.8, 1.6, 4.4, 70, 'soft', 0.7)]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.9, 0.95)),
                cg=CG(*LEFT(7.2), pic=(8.6, 1.0, 3.6, 4.8), pk='leaf'), sg=SG((0.9, 2.0, 7.4, 2.2), (0.9, 4.3, 7.4, 0.8)))


def m_chevrons(p):
    def cv(x, y, s, f, a=None):
        return Dm(x, y, s, s, f, a=a)
    cover = [full('ink')] + [cv(7.8 + i * 1.2, 2.3, 2.6, 'accent' if i % 2 else 'soft', 0.95 - i * 0.12) for i in range(4)]
    content = [full('bg')] + [cv(11.7 + i * 0.42, 0.35, 0.65, 'accent' if i % 2 else 'soft') for i in range(3)]
    section = [full('bg')] + [cv(-1.6 + i * 1.2, 2.1, 3.2, 'ink' if i % 2 else 'accent', 0.95) for i in range(4)]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.6, 0.95)),
                cg=DARK(*LEFT(6.5), pic=(8.6, 1.2, 4.0, 5.0), pr=0.2),
                sg=SGL((6.6, 2.0, 6.0, 2.2), (6.6, 4.3, 6.0, 0.8)))


def m_card(p):
    cover = [full('soft'), R(1.0, 0.9, 11.33, 5.7, 'white', rr=0.35), R(1.0, 0.9, 11.33, 0.18, 'accent', rr=0.09), circ(11.4, 1.9, 1.2, 'accent', a=0.25)]
    content = [full('tint'), R(0.5, 0.35, 12.33, 6.45, 'white', rr=0.3)]
    section = [full('ink'), R(1.6, 1.6, 10.13, 4.3, 'ink2', rr=0.35), R(1.6, 5.9, 10.13, 0.12, 'accent')]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(1.0, 0.6, 11.3, 0.95), body=(1.0, 1.8, 11.3, 4.7)),
                cg=CG((1.7, 2.0, 9.9, 2.2), (1.7, 4.25, 9.9, 0.7), (1.7, 4.95, 9.9, 0.5), al='center', pic=(8.3, 1.5, 3.6, 4.6), pr=0.25),
                sg=SG((2.2, 2.2, 8.9, 1.9), (2.2, 4.15, 8.9, 0.8), al='center'))


def m_halftone(p):
    ht = [circ(9.2 + c * 0.55, 0.6 + r * 0.55, 0.42 * (1 - (abs(c - 3) + abs(r - 4)) / 9), 'accent', a=0.9) for r in range(11) for c in range(8)]
    cover = [full('bg')] + [d for d in ht if d['w'] > 0.04]
    content = [full('bg')] + [circ(11.8 + c * 0.3, 0.35 + r * 0.3, max(0.03, 0.2 * (1 - (c + r) / 7)), 'accent', a=0.85) for r in range(4) for c in range(5)]
    section = [full('ink')] + [dict(d, x=d['x'] - 8.6, f='soft', a=0.45) for d in ht if d['w'] > 0.04]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.6, 0.95)),
                cg=CG(*LEFT(7.6), pic=(9.0, 1.0, 3.6, 5.0), pr=0.2), sg=SG((5.9, 2.0, 6.7, 2.2), (5.9, 4.3, 6.7, 0.8)))


def m_sun(p):
    def rays(cx, cy, L, n, f, a):
        return [R(cx - L / 2, cy - 0.04, L, 0.08, f, a=a, r=i * 180 / n) for i in range(n)]
    cover = [full('bg')] + rays(11.0, 7.6, 9.0, 12, 'soft', 0.9) + [circ(11.0, 7.6, 4.2, 'accent'), circ(11.0, 7.6, 2.6, 'ink2')]
    content = [full('bg')] + rays(13.1, 0.1, 2.4, 8, 'soft', 0.9) + [circ(13.1, 0.1, 1.3, 'accent')]
    section = [full('ink')] + rays(W / 2, 8.0, 12.0, 14, 'ink2', 0.45) + [circ(W / 2, 8.0, 4.8, 'accent')]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.6, 0.95)),
                cg=CG(*LEFT(6.6, 1.0), pic=(8.8, 0.7, 3.6, 3.6), pk='circle'), sg=SG((1.0, 1.3, 11.3, 2.0), (1.0, 3.3, 11.3, 0.8), al='center'))


def m_split(p):
    cover = [full('bg'), R(0, 0, 6.66, H, 'ink'), circ(6.66, 3.75, 2.2, 'accent'), circ(6.66, 3.75, 1.1, 'bg')]
    content = [full('bg'), R(0, 0, W, 0.16, 'ink'), R(0, 0.16, 3.0, 0.06, 'accent')]
    section = [full('soft'), R(0, 0, 6.66, H, 'ink')]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 11.5, 0.95)),
                cg=CG((0.8, 1.6, 4.6, 2.6), (8.0, 2.4, 4.6, 0.9), (8.0, 3.3, 4.6, 0.6), tc='white', dark=True, pic=(8.0, 4.1, 4.6, 2.7), pr=0.2),
                sg=SG((0.8, 2.0, 5.2, 2.2), (7.4, 3.0, 5.2, 0.9), sc='text'))


def m_corners(p):
    cover = [full('bg'), R(0.7, 0.7, 1.6, 0.1, 'accent'), R(0.7, 0.7, 0.1, 1.6, 'accent'), R(11.03, 6.7, 1.6, 0.1, 'accent'), R(12.53, 5.2, 0.1, 1.6, 'accent'),
             circ(11.0, 2.0, 2.4, 'soft', a=0.7)]
    content = [full('bg'), R(0.5, 0.36, 0.75, 0.07, 'accent'), R(0.5, 0.36, 0.07, 0.75, 'accent'), R(12.08, 7.0, 0.75, 0.07, 'soft'), R(12.76, 6.32, 0.07, 0.75, 'soft')]
    section = [full('ink'), R(1.0, 1.0, 2.0, 0.1, 'accent'), R(1.0, 1.0, 0.1, 2.0, 'accent'), R(10.33, 6.4, 2.0, 0.1, 'accent'), R(12.23, 4.5, 0.1, 2.0, 'accent')]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.62, 11.3, 0.95)),
                cg=CG((1.2, 1.6, 8.2, 2.5), (1.2, 4.2, 8.2, 0.7), (1.2, 4.9, 8.2, 0.5), pic=(9.0, 1.0, 3.5, 4.6), pr=0.1),
                sg=SG((1.6, 2.3, 10.0, 2.0), (1.6, 4.3, 10.0, 0.8)))


def m_stack(p):
    cover = [full('bg'), R(8.4, 1.4, 4.2, 4.8, 'soft', rr=0.3, r=-8), R(8.2, 1.2, 4.2, 4.8, 'accent', rr=0.3, r=-3, a=0.9), R(8.0, 1.0, 4.2, 4.8, 'ink', rr=0.3, r=3)]
    content = [full('bg'), R(12.05, 0.28, 0.8, 0.95, 'soft', rr=0.1, r=-8), R(12.0, 0.25, 0.8, 0.95, 'accent', rr=0.1, r=4, a=0.9)]
    section = [full('soft'), R(1.2, 1.2, 10.9, 5.1, 'white', rr=0.3, r=-2), R(1.0, 1.0, 10.9, 5.1, 'ink', rr=0.3, r=1)]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.8, 0.95)),
                cg=CG(*LEFT(6.6), pic=(8.0, 1.0, 4.2, 4.8), pr=0.3), sg=SG((1.8, 2.1, 9.6, 2.0), (1.8, 4.1, 9.6, 0.8)))


def m_timeline(p):
    cover = [full('ink'), R(0.9, 5.6, 11.5, 0.05, 'soft', a=0.6)] + [circ(0.9 + i * 2.875, 5.62, 0.5 if i == 0 else 0.34, 'accent' if i == 0 else 'soft') for i in range(5)]
    content = [full('bg'), R(0.45, 0.3, 0.04, 6.9, 'soft'), circ(0.47, 0.98, 0.26, 'accent')]
    section = [full('ink'), R(0, 3.75, W, 0.05, 'soft', a=0.5), circ(2.0, 3.77, 0.7, 'accent')]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.95, 0.5, 11.4, 0.95), body=(0.95, 1.75, 11.4, 4.85)),
                cg=DARK((0.9, 1.0, 10.5, 2.5), (0.9, 3.6, 10.5, 0.7), (0.9, 4.3, 10.5, 0.5), pic=(9.4, 0.8, 3.2, 3.8), pr=0.2),
                sg=SG((3.0, 1.2, 9.4, 2.2), (3.0, 4.4, 9.4, 0.8)))


def m_tiles(p):
    tl = [R(7.9 + c * 1.35, 0.8 + r * 1.35, 1.2, 1.2, ('accent', 'soft', 'ink2', 'tint')[(r + c * 2) % 4], rr=0.18) for r in range(4) for c in range(4) if (r + c) % 3 != 2]
    cover = [full('bg')] + tl
    content = [full('bg')] + [R(11.8 + c * 0.4, 0.28 + r * 0.4, 0.34, 0.34, ('accent', 'soft', 'ink2')[(r + c) % 3], rr=0.06) for r in range(2) for c in range(3)]
    section = [full('ink')] + [dict(t, x=t['x'] - 7.1) for t in tl]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.6, 0.95)),
                cg=CG(*LEFT(6.6), pic=(9.25, 2.15, 2.55, 2.55), pr=0.18), sg=SG((6.6, 2.0, 6.0, 2.2), (6.6, 4.3, 6.0, 0.8)))


def m_mountain(p):
    cover = [full('bg', g=['tint', 'bg', 270]), Tri(5.0, 3.4, 7.6, 4.1, 'ink2'), Tri(8.0, 3.9, 6.4, 3.6, 'ink'), Tri(2.6, 4.9, 6.2, 2.6, 'accent', a=0.9),
             circ(10.9, 1.4, 1.1, 'accent')]
    content = [full('bg'), Tri(11.6, 0.3, 1.45, 0.95, 'ink2'), Tri(12.25, 0.55, 1.1, 0.7, 'accent', a=0.9)]
    section = [full('ink'), Tri(-0.5, 4.3, 7.0, 3.2, 'ink2'), Tri(3.6, 5.0, 6.0, 2.5, 'accent', a=0.85), circ(11.0, 1.3, 1.2, 'soft', a=0.8)]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.6, 0.95)),
                cg=CG((0.9, 0.7, 7.8, 2.0), (0.9, 2.75, 6.6, 0.6), (0.9, 3.35, 6.6, 0.45), pic=(9.6, 0.5, 3.0, 2.6), pr=0.2),
                sg=SG((0.9, 0.9, 9.2, 2.2), (0.9, 3.1, 9.2, 0.8)))


def m_lens(p):
    cover = [full('ink', g=['ink', 'deep', 315]), circ(10.0, 3.4, 4.6, 'accent', a=0.45), circ(11.5, 2.4, 3.2, 'soft', a=0.4), circ(10.8, 5.3, 2.2, 'ink2', a=0.6)]
    content = [full('bg'), circ(12.7, 0.55, 1.7, 'accent', a=0.25), circ(13.05, 1.3, 1.0, 'soft', a=0.6)]
    section = [full('ink', g=['ink', 'deep', 315]), circ(3.2, 3.8, 5.0, 'accent', a=0.35), circ(5.0, 2.6, 3.0, 'soft', a=0.3)]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.8, 0.95)),
                cg=DARK(*LEFT(7.0), pic=(8.6, 1.3, 3.8, 3.8), pk='circle'), sg=SG((6.4, 2.0, 6.2, 2.2), (6.4, 4.3, 6.2, 0.8)))


def m_ticket(p):
    cover = [full('ink'), R(0.9, 1.1, 11.53, 5.3, 'bg', rr=0.25), circ(0.9, 3.75, 0.9, 'ink'), circ(12.43, 3.75, 0.9, 'ink'),
             R(9.4, 1.1, 3.03, 5.3, 'accent', rr=0.25), R(9.4, 1.1, 1.0, 5.3, 'accent'), R(9.3, 1.1, 0.05, 5.3, 'soft')]
    content = [full('bg'), R(12.1, 0, 0.9, 1.2, 'accent', rr=0.08), circ(12.55, 1.2, 0.34, 'bg')]
    section = [full('accent'), R(1.2, 1.4, 10.93, 4.7, 'ink', rr=0.25), circ(1.2, 3.75, 0.8, 'accent'), circ(12.13, 3.75, 0.8, 'accent')]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.9, 0.95)),
                cg=CG((1.6, 1.6, 7.3, 2.5), (1.6, 4.2, 7.3, 0.7), (1.6, 4.9, 7.3, 0.5), pic=(9.7, 1.4, 2.45, 4.7), pr=0.15),
                sg=SG((2.0, 2.2, 9.3, 1.9), (2.0, 4.1, 9.3, 0.8)))


def m_tape(p):
    cover = [full('soft'), R(0.8, 0.9, 11.7, 5.7, 'bg', r=-1.5), R(5.4, 0.55, 2.5, 0.65, 'accent', r=-4, a=0.85), R(10.6, 6.0, 2.2, 0.6, 'ink2', r=8, a=0.8)]
    content = [full('bg'), R(11.6, 0.15, 1.6, 0.4, 'accent', r=-6, a=0.8)]
    section = [full('ink'), R(1.0, 1.5, 11.33, 4.5, 'bg', r=1.5), R(5.4, 1.2, 2.5, 0.6, 'accent', r=-4, a=0.9)]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.6, 10.5, 0.95)),
                cg=CG((1.5, 1.6, 10.3, 2.5), (1.5, 4.2, 10.3, 0.7), (1.5, 4.9, 10.3, 0.5), pic=(8.6, 1.4, 3.4, 4.4), pr=0.0),
                sg=SGL((1.9, 2.3, 9.5, 1.8), (1.9, 4.15, 9.5, 0.8)))


def m_pills(p):
    pl = [R(7.4 + (i % 2) * 0.6, 0.8 + i * 1.15, 5.0 - (i % 3) * 0.8, 0.8, ('accent', 'soft', 'ink2', 'tint', 'accent')[i], rr=0.4) for i in range(5)]
    cover = [full('bg')] + pl
    content = [full('bg'), R(11.0, 0.3, 1.9, 0.32, 'accent', rr=0.16), R(11.6, 0.75, 1.3, 0.32, 'soft', rr=0.16)]
    section = [full('ink')] + [dict(x, x=x['x'] - 6.6) for x in pl]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 9.9, 0.95)),
                cg=CG(*LEFT(6.2), pic=(8.2, 1.2, 4.2, 4.2), pk='circle'), sg=SG((7.2, 2.0, 5.4, 2.2), (7.2, 4.3, 5.4, 0.8)))


def m_quarter(p):
    cover = [full('bg'), circ(W, H, 9.4, 'ink'), circ(W, H, 6.4, 'accent'), circ(W, H, 3.4, 'soft')]
    content = [full('bg'), circ(W, 0, 2.4, 'soft', a=0.9), circ(W, 0, 1.3, 'accent')]
    section = [full('ink'), circ(0, H, 9.0, 'ink2'), circ(0, H, 5.6, 'accent', a=0.9)]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.6, 0.95)),
                cg=CG((0.9, 0.8, 7.4, 2.4), (0.9, 3.25, 6.0, 0.7), (0.9, 3.95, 6.0, 0.5), pic=(9.0, 0.6, 3.6, 3.6), pk='circle'),
                sg=SG((5.4, 1.2, 7.2, 2.2), (5.4, 3.5, 7.2, 0.8)))


def m_magazine(p):
    cover = [full('bg'), R(0, 0, W, 0.55, 'ink'), R(0.9, 6.35, 11.53, 0.04, 'ink'), R(0.9, 6.45, 2.0, 0.12, 'accent')]
    content = [full('bg'), R(0.9, 0.28, 11.53, 0.03, 'ink'), R(0.9, 1.55, 11.53, 0.015, 'muted', a=0.7)]
    section = [full('bg'), R(0, 0, W, 4.65, 'ink'), R(0.9, 5.0, 2.2, 0.14, 'accent')]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.5, 11.53, 0.95), body=(0.9, 1.8, 11.53, 4.8)),
                cg=CG((0.9, 1.9, 11.5, 3.0), (0.9, 5.0, 11.5, 0.6), (0.9, 5.6, 11.5, 0.5), pic=(8.8, 1.0, 3.6, 5.0), pr=0.0, tp=68),
                sg=SG((0.9, 1.2, 11.5, 2.4), (0.9, 5.3, 11.5, 0.8), sc='text'))


def m_glow(p):
    cover = [full('ink'), circ(11.2, 5.4, 7.0, 'accent', a=0.35), circ(12.2, 1.6, 3.8, 'ink2', a=0.8), circ(1.4, -1.0, 3.2, 'accent2', a=0.3)]
    content = [full('bg', g=['tint', 'bg', 270])]
    section = [full('ink'), circ(2.0, 5.8, 6.0, 'accent', a=0.35), circ(11.4, 1.4, 3.6, 'ink2', a=0.8)]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 11.5, 0.95)),
                cg=DARK(*LEFT(7.6), pic=(8.8, 1.4, 3.8, 3.8), pk='circle'), sg=SG((4.6, 2.0, 7.9, 2.2), (4.6, 4.3, 7.9, 0.8)))


def m_steps(p):
    st = [R(7.4 + i * 1.18, 6.3 - i * 0.92, 5.933 - i * 1.18, 0.92, ('ink', 'ink2', 'accent', 'soft', 'tint')[i]) for i in range(5)]
    cover = [full('bg')] + st
    content = [full('bg')] + [R(12.0 + i * 0.27, 1.0 - i * 0.27, 1.333 - i * 0.27, 0.27, ('ink', 'accent', 'soft')[i]) for i in range(3)]
    section = [full('ink')] + [dict(s, x=s['x'] - 7.4, f=('ink2', 'accent', 'soft', 'tint', 'bg')[i]) for i, s in enumerate(st)]
    return dict(decor=dict(cover=cover, content=content, section=section), geo=G(title=(0.9, 0.55, 10.9, 0.95)),
                cg=CG((0.9, 0.9, 6.4, 2.4), (0.9, 3.35, 6.2, 0.7), (0.9, 4.05, 6.2, 0.5), pic=(8.6, 0.5, 4.0, 2.7), pr=0.2),
                sg=SG((6.8, 0.9, 5.8, 2.2), (6.8, 3.2, 5.8, 0.8)))


# ======================================================================== Morph designs
# A motif with a step: the same named shapes, placed differently on each slide, so they glide when the slide changes
# (the cover is step 0; the steps repeat). The words keep to their side; the shapes move on theirs.

def mo_puzzle(p, i):
    out = [full('bg', n='bg')]
    for t in range(4):
        on = t == i % 4
        s = 1.6 if on else 1.1
        cx = 1.7 + t * 2.0
        out.append(R(cx - s / 2 + 0.4, 1.15 - s / 2 + 0.1, s, s, 'accent' if on else 'soft', rr=0.22 if on else 0.16, r=0 if on else 8 * (t - 1.5),
                     n='tile%d' % t, tx=str(t + 1), ts=32 if on else 22, tb=True))
    return out


def mo_orbit(p, i):
    out = [full('bg', n='bg'), Rg(8.1, 0.85, 5.8, 5.8, 'soft', a=0.7, th=0.03, n='path')]
    letters = 'OCEAN'
    for t in range(5):
        slot = (t - i) % 5
        ang = math.radians(-90 + slot * 72)
        on = slot == 0
        d = 1.45 if on else 0.95
        cx, cy = 11.0 + 2.9 * math.cos(ang), 3.75 + 2.9 * math.sin(ang)
        out.append(circ(cx, cy, d, ('accent', 'ink2', 'soft', 'accent2', 'ink')[t], n='o%d' % t, tx=letters[t], ts=28 if on else 18, tb=True))
    return out


def mo_wheel(p, i):
    rot = -90 * i
    out = [full('bg', n='bg'), circ(10.6, 3.75, 5.4, 'ink', n='wheel'), circ(10.6, 3.75, 2.8, 'bg', n='hub')]
    for t in range(4):
        out.append(R(10.6 - 2.7, 3.75 - 0.05, 5.4, 0.1, 'bg', r=rot + t * 45, n='spoke%d' % t))
    for t in range(4):
        ang = math.radians(rot + t * 90 - 90)
        on = (t + i) % 4 == 0
        out.append(circ(10.6 + 2.05 * math.cos(ang), 3.75 + 2.05 * math.sin(ang), 0.95 if on else 0.65, 'accent' if on else 'soft', n='dot%d' % t,
                        tx='✓' if on else None, ts=22 if on else None, tb=True if on else None))
    return out


def mo_wavestrips(p, i):
    out = [full('bg', n='bg')]
    for t in range(9):
        h = 0.5 + 1.6 * (0.5 + 0.5 * math.sin((t + i * 1.6) * 0.8))
        out.append(R(0.9 + t * 1.3, 0, 1.05, h, ('accent', 'ink2', 'soft')[t % 3], n='s%d' % t))
    return out


def mo_spiral(p, i):
    out = [full('bg', n='bg')] + [circ(0.55, 0.62 + k * 0.78, 0.3, 'ink', n='ring%d' % k) for k in range(9)]
    for t in range(4):
        on = t == i % 4
        out.append(R(12.1 if on else 12.6, 0.9 + t * 1.45, 1.6 if on else 1.1, 1.1, ('accent', 'soft', 'ink2', 'accent2')[t], rr=0.12, n='tab%d' % t,
                     tx=str(t + 1), ts=22, tb=True))
    out.append(R(1.0, 0, 0.04, H, 'accent', n='margin'))
    return out


def mo_carousel(p, i):
    out = [full('bg', n='bg')]
    for t in range(5):
        slot = (t - i) % 5
        on = slot == 2
        s = 1.9 if on else (1.3 if slot in (1, 3) else 0.9)
        cx = 2.2 + slot * 2.25
        out.append(R(cx - s / 2 + 0.4, 6.0 - s / 2 - (0.15 if on else 0), s, s, ('accent', 'ink2', 'soft', 'accent2', 'ink')[t], rr=0.25, n='al%d' % t, a=1 if on else 0.88))
    return out


def mo_zoom(p, i):
    k = i % 4
    d = (1.0, 3.0, 5.4, 9.0)[k]
    return [full('bg', n='bg'), circ(12.6, 0.6, d, 'soft', n='zoom', a=0.85), circ(12.6, 0.6, d * 0.45, 'accent', n='core', a=0.9)]


def mo_panels(p, i):
    ws = [(1.0, 0.5, 0.5), (0.5, 1.0, 0.5), (0.5, 0.5, 1.0)][i % 3]
    out = [full('bg', n='bg')]
    x = W - sum(ws) * 2.2
    for t, w in enumerate(ws):
        out.append(R(x, 0, w * 2.2, H, ('ink', 'accent', 'soft')[t], n='pan%d' % t))
        x += w * 2.2
    return out


def mo_stairs(p, i):
    out = [full('bg', n='bg')]
    for t in range(5):
        on = t == i % 5
        out.append(R(8.8 + t * 0.85, 6.9 - (t + 1) * 0.85, 0.85, (t + 1) * 0.85, 'accent' if on else ('soft' if t % 2 else 'tint'), n='st%d' % t))
    k = i % 5
    out.append(circ(8.8 + k * 0.85 + 0.425, 6.9 - (k + 1) * 0.85 - 0.45, 0.55, 'ink', n='ball'))
    return out


def mo_bubbles(p, i):
    out = [full('bg', n='bg')]
    for t in range(7):
        y = 6.6 - ((t * 1.7 + i * 1.1) % 6.4)
        d = 0.5 + (t % 3) * 0.4
        out.append(circ(9.6 + (t * 1.37) % 3.2, y, d, ('accent', 'soft', 'ink2')[t % 3], a=0.85, n='b%d' % t))
    return out


def mo_pendulum(p, i):
    ang = math.radians((-20, 0, 20, 0)[i % 4])
    px, py, L = 11.2, -0.2, 4.9
    bx, by = px + L * math.sin(ang), py + L * math.cos(ang)
    return [full('bg', n='bg'), R((px + bx) / 2 - 0.03, (py + by) / 2 - L / 2, 0.06, L, 'ink', r=-math.degrees(ang), n='rod'),
            circ(bx, by, 1.5, 'accent', n='bob'), circ(px, py, 0.5, 'ink', n='pivot')]


def mo_ripple(p, i):
    out = [full('bg', n='bg')]
    for t in range(4):
        d = 1.2 + ((t + i) % 4) * 1.7
        out.append(Rg(11.4 - d / 2, 5.4 - d / 2, d, d, 'accent' if t % 2 else 'ink2', th=0.09, a=0.9 - 0.18 * ((t + i) % 4), n='rp%d' % t))
    out.append(circ(11.4, 5.4, 0.7, 'accent', n='drop'))
    return out


def mo_gears(p, i):
    return [full('bg', n='bg'), Hx(9.6, 1.0, 3.0, 2.6, 'ink2', r=30 * i, n='g1'), circ(11.1, 2.3, 0.9, 'bg', n='g1h'),
            Hx(11.3, 3.3, 2.1, 1.82, 'accent', r=-45 * i, n='g2'), circ(12.35, 4.21, 0.55, 'bg', n='g2h'),
            Hx(9.9, 4.8, 1.45, 1.26, 'soft', r=60 * i, n='g3')]


def mo_compass(p, i):
    ang = (0, 90, 200, 300)[i % 4]
    return [full('bg', n='bg'), Rg(8.7, 1.05, 5.3, 5.3, 'ink', th=0.12, n='rim'), circ(11.35, 3.7, 4.3, 'tint', n='face'),
            Dm(11.35 - 0.33, 3.7 - 2.0, 0.66, 4.0, 'accent', r=ang, n='needle'), circ(11.35, 3.7, 0.45, 'ink', n='pin')]


def mo_cards(p, i):
    out = [full('bg', n='bg')]
    for t in range(4):
        on = t == i % 4
        out.append(R(9.4 + t * 0.32, 1.4 + (0 if on else 0.4), 2.7, 3.8, ('accent', 'ink2', 'soft', 'accent2')[t], rr=0.2,
                     r=(t - 1.5) * 8 if not on else -2, n='card%d' % t, a=1 if on else 0.9))
    return out


def mo_spot(p, i):
    pos = [(10.9, 2.0), (11.6, 4.9), (10.3, 4.6), (12.0, 1.5)][i % 4]
    return [full('bg', n='bg'), circ(pos[0], pos[1], 3.4, 'soft', a=0.75, n='spot'), circ(pos[0] + 0.4, pos[1] - 0.3, 1.1, 'accent', a=0.85, n='glint'),
            circ(9.4, 6.4, 0.3, 'ink', n='d1'), circ(12.6, 6.6, 0.22, 'accent', n='d2')]


def mo_bars(p, i):
    out = [full('bg', n='bg')]
    for t in range(7):
        h = 0.6 + 2.5 * (0.5 + 0.5 * math.sin(t * 1.3 + i * 1.9))
        out.append(R(9.0 + t * 0.58, 6.6 - h, 0.4, h, ('accent', 'ink2', 'soft')[t % 3], rr=0.08, n='bar%d' % t))
    out.append(R(8.8, 6.62, 4.2, 0.05, 'ink', n='axis'))
    return out


def mo_cascade(p, i):
    out = [full('bg', n='bg')]
    for r in range(3):
        for c in range(3):
            on = (r + c) % 3 == i % 3
            s = 0.9 if on else 0.58
            out.append(R(10.3 + c * 0.95 + (0.9 - s) / 2, 0.25 + r * 0.95 + (0.9 - s) / 2, s, s, 'accent' if on else 'soft', rr=0.1, n='c%d%d' % (r, c)))
    return out


def mo_mountains(p, i):
    s = (0, 1.2, 2.4, 1.2)[i % 4]
    return [full('bg', n='bg'), Tri(8.3 - s * 0.3, 4.6, 6.0, 2.9, 'soft', n='m1'), Tri(9.6 + s * 0.3, 4.1, 5.2, 3.4, 'ink2', n='m2'),
            Tri(8.6 - s * 0.6, 5.2, 4.6, 2.3, 'accent', a=0.9, n='m3'), circ(11.9 - s * 0.9, 1.3 + s * 0.3, 0.95, 'accent', n='sun')]


def mo_progress(p, i):
    k = i % 5
    out = [full('bg', n='bg'), R(0.9, 6.55, 11.5, 0.06, 'soft', n='track'), R(0.9, 6.55, max(0.06, 11.5 * k / 4), 0.06, 'accent', n='fill')]
    for t in range(5):
        out.append(circ(0.9 + t * 2.875, 6.58, 0.46 if t == k else 0.28, 'accent' if t <= k else 'soft', n='pt%d' % t))
    return out


def mo_ribbon(p, i):
    y1, y2, ang = [(0.18, 0.6, -3), (6.38, 6.14, 1.2), (0.12, 0.54, -2), (6.4, 6.16, -1.2)][i % 4]
    return [full('bg', n='bg'), R(-2, y1, 17.5, 0.36, 'accent', r=ang, n='rib1'), R(-2, y2, 17.5, 0.12, 'ink2', r=ang, n='rib2')]


def mo_hive(p, i):
    def hx(cx, cy, s, f, n):
        return Hx(cx - s / 2, cy - s * 0.433, s, s * 0.866, f, n=n)
    out = [full('bg', n='bg')]
    cells = [(0, 0), (0.76, 0.44), (0, 0.88), (0.76, -0.44), (-0.76, 0.44), (-0.76, -0.44), (0, -0.88)]
    for t, (dx, dy) in enumerate(cells):
        on = t == i % 7
        out.append(hx(11.1 + dx * 1.45, 3.7 + dy * 1.45, 1.4 if on else 1.2, 'accent' if on else ('soft' if t % 2 else 'tint'), 'h%d' % t))
    return out


def mo_split(p, i):
    lw = (4.0, 1.2, 2.6)[i % 3]
    return [full('bg', n='bg'), R(W - lw, 0, lw, H, 'ink', n='half'), R(W - lw - 0.18, 0, 0.18, H, 'accent', n='edge'),
            circ(W - lw - 0.09, 3.75, 1.2, 'accent', a=0.9, n='knob')]


def mo_kaleido(p, i):
    out = [full('bg', n='bg')]
    for t in range(6):
        ang = t * 60 + i * 30
        a = math.radians(ang)
        out.append(Tri(11.1 + 1.2 * math.sin(a) - 0.95, 3.75 - 1.2 * math.cos(a) - 0.95, 1.9, 1.9, ('accent', 'soft', 'ink2')[(t + i) % 3], r=ang, n='k%d' % t, a=0.9))
    return out


def mo_planet(p, i):
    a = math.radians(-30 + 80 * (i % 4))
    return [full('bg', n='bg'), circ(11.1, 4.1, 3.2, 'ink2', n='planet'), Rg(8.6, 3.4, 5.0, 1.4, 'accent', th=0.08, r=-15, n='band'),
            circ(11.1 + 2.6 * math.cos(a), 4.1 + 2.6 * math.sin(a) * 0.55, 0.65, 'accent', n='moon')]


def mo_snake(p, i):
    out = [full('bg', n='bg')]
    for t in range(8):
        x, y = 8.9 + t * 0.58, 3.75 + 1.5 * math.sin(t * 0.75)
        on = t == i % 8
        out.append(circ(x, y, 0.6 if on else 0.34, 'accent' if on else 'soft', n='sn%d' % t))
    return out


def mo_blinds(p, i):
    op = (0.2, 0.7, 0.45, 0.9)[i % 4]
    return [full('bg', n='bg')] + [R(9.5, 0.25 + t * 1.0, 3.8, 1.0 * op, ('ink', 'accent')[t % 2], n='bl%d' % t, a=0.95) for t in range(7)]


def mo_papers(p, i):
    out = [full('bg', n='bg')]
    for t in range(3):
        on = t == i % 3
        out.append(R(9.1 + t * 0.25, 1.2 + t * 0.2, 3.5, 4.5, ('soft', 'accent', 'ink2')[t], rr=0.2, r=(-10, 4, 12)[(t + i) % 3], n='pp%d' % t, a=1 if on else 0.9))
    return out


def mo_sunburst(p, i):
    rot = 15 * i
    out = [full('bg', n='bg')] + [R(11.8 - 2.9, 1.2 - 0.05, 5.8, 0.1, 'soft', r=rot + k * 22.5, n='ray%d' % k) for k in range(8)]
    out.append(circ(11.8, 1.2, 1.7, 'accent', n='sunc'))
    return out


def mo_confetti(p, i):
    out = [full('bg', n='bg')]
    kinds = [St, Pl, O, Tri, Dm]
    for t in range(10):
        x = 9.0 + ((t * 1.91 + i * 0.83) % 3.9)
        y = 0.4 + ((t * 2.37 + i * 1.47) % 6.3)
        s = 0.35 + (t % 3) * 0.17
        out.append(kinds[t % 5](x, y, s, s, ('accent', 'ink2', 'soft', 'accent2')[t % 4], r=(t * 37 + i * 50) % 360, n='cf%d' % t))
    return out


def mo_sweep(p, i):
    pos = [(10.7, -1.3), (11.6, 4.7), (9.4, 4.9), (12.0, 1.2)][i % 4]
    return [full('bg', n='bg'), R(pos[0], pos[1], 3.4, 3.4, 'accent', r=45 + 30 * i, rr=0.3, a=0.9, n='sq'),
            R(pos[0] + 0.75, pos[1] + 0.75, 1.9, 1.9, 'bg', r=45 + 30 * i, rr=0.2, n='sqh')]


def mo_lenses(p, i):
    out = [full('bg', n='bg')]
    for t in range(4):
        x = 9.2 + ((t + i) % 4) * 1.0
        out.append(circ(x + 0.9, 1.2 + t * 1.5, 1.8 - t * 0.28, ('accent', 'soft', 'ink2', 'accent2')[t], a=0.55, n='ln%d' % t))
    return out


def mo_bricks(p, i):
    out = [full('bg', n='bg')]
    for t in range(6):
        row, col = t // 2, t % 2
        on = t == i % 6
        out.append(R(9.7 + col * 1.7 + (0.8 if row % 2 else 0) - (0.2 if on else 0), 0.9 + row * 1.05 - (0.12 if on else 0), 1.55 + (0.4 if on else 0),
                     0.9 + (0.24 if on else 0), 'accent' if on else ('soft' if t % 2 else 'tint'), rr=0.08, n='br%d' % t))
    return out


# (key, name, motif, steps, palette, where the title and the points go)
TL = (0.9, 0.6, 7.6, 0.95), (0.9, 1.8, 7.6, 4.8)           # words on the left, the shapes on the right
MORPHS = [
    ('puzzle', 'Puzzle Steps', mo_puzzle, 4, P('1E293B', 'F97316', 'FED7AA'), ((0.9, 2.55, 11.5, 0.95), (0.9, 3.65, 11.5, 3.0))),
    ('orbit', 'Orbit Letters', mo_orbit, 5, P('0C4A6E', '0EA5E9', 'BAE6FD', accent2='F59E0B'), ((0.9, 0.6, 6.6, 0.95), (0.9, 1.8, 6.6, 4.8))),
    ('wheel', 'Turning Wheel', mo_wheel, 4, P('1F2937', '10B981', 'A7F3D0'), ((0.9, 0.6, 6.6, 0.95), (0.9, 1.8, 6.6, 4.8))),
    ('wavestrips', 'Wave Strips', mo_wavestrips, 6, P('312E81', '6366F1', 'C7D2FE'), ((0.9, 2.45, 11.5, 0.95), (0.9, 3.55, 11.5, 3.1))),
    ('spiral', 'Spiral Notebook', mo_spiral, 4, P('3F2E1E', 'D97706', 'FDE68A'), ((1.4, 0.6, 10.2, 0.95), (1.4, 1.8, 10.2, 4.9))),
    ('carousel', 'Album Carousel', mo_carousel, 5, P('4C1D95', 'EC4899', 'FBCFE8', accent2='8B5CF6'), ((0.9, 0.55, 11.5, 0.95), (0.9, 1.65, 11.5, 3.2))),
    ('zoom', 'Zoom Circle', mo_zoom, 4, P('134E4A', '14B8A6', 'CCFBF1'), ((0.9, 0.6, 7.0, 0.95), (0.9, 1.8, 7.0, 4.8))),
    ('panels', 'Sliding Panels', mo_panels, 3, P('172554', 'F43F5E', 'FECDD3'), ((0.9, 0.6, 7.2, 0.95), (0.9, 1.8, 7.2, 4.8))),
    ('stairs', 'Climbing Stairs', mo_stairs, 5, P('14532D', '22C55E', 'BBF7D0'), TL),
    ('bubbles', 'Rising Bubbles', mo_bubbles, 6, P('0F3460', '38BDF8', 'BAE6FD'), TL),
    ('pendulum', 'Pendulum', mo_pendulum, 4, P('3B0764', 'A855F7', 'E9D5FF'), TL),
    ('ripple', 'Ripples', mo_ripple, 4, P('083344', '06B6D4', 'A5F3FC'), ((0.9, 0.6, 7.2, 0.95), (0.9, 1.8, 7.2, 4.8))),
    ('gears', 'Turning Gears', mo_gears, 6, P('292524', 'EAB308', 'FEF08A'), TL),
    ('compass', 'Compass', mo_compass, 4, P('1E3A8A', 'EF4444', 'DBEAFE'), ((0.9, 0.6, 7.4, 0.95), (0.9, 1.8, 7.4, 4.8))),
    ('cards', 'Card Fan', mo_cards, 4, P('7C2D12', 'FB923C', 'FFEDD5', accent2='F43F5E'), TL),
    ('spotlight', 'Spotlight', mo_spot, 4, P('111827', 'F59E0B', 'FDE68A'), TL),
    ('bars', 'Moving Bars', mo_bars, 6, P('1E1B4B', '22D3EE', 'C7D2FE'), TL),
    ('cascade', 'Tile Cascade', mo_cascade, 3, P('4A044E', 'D946EF', 'F5D0FE'), ((0.9, 0.6, 8.9, 0.95), (0.9, 3.3, 11.5, 3.3))),
    ('mountains', 'Mountain Parallax', mo_mountains, 4, P('1C1917', 'F97316', 'E7E5E4'), ((0.9, 0.6, 7.0, 0.95), (0.9, 1.8, 6.6, 4.8))),
    ('progress', 'Progress Track', mo_progress, 5, P('0F172A', '6366F1', 'E0E7FF'), ((0.9, 0.6, 11.5, 0.95), (0.9, 1.8, 11.5, 4.3))),
    ('ribbon', 'Swinging Ribbon', mo_ribbon, 4, P('881337', 'F43F5E', 'FFE4E6'), ((0.9, 1.0, 11.5, 0.95), (0.9, 2.15, 11.5, 3.8))),
    ('hive', 'Honeycomb', mo_hive, 7, P('422006', 'F59E0B', 'FEF3C7'), TL),
    ('split', 'Split Reveal', mo_split, 3, P('0B1324', '3B82F6', 'DBEAFE'), ((0.9, 0.6, 7.3, 0.95), (0.9, 1.8, 7.3, 4.8))),
    ('kaleido', 'Kaleidoscope', mo_kaleido, 4, P('1E1B4B', 'F472B6', 'C4B5FD', accent2='34D399'), TL),
    ('planet', 'Planet and Moon', mo_planet, 4, P('020617', '8B5CF6', 'DDD6FE'), TL),
    ('snake', 'Snake Path', mo_snake, 8, P('052E16', '84CC16', 'D9F99D'), TL),
    ('blinds', 'Window Blinds', mo_blinds, 4, P('0C0A09', '0EA5E9', 'E0F2FE'), ((0.9, 0.6, 8.0, 0.95), (0.9, 1.8, 8.0, 4.8))),
    ('papers', 'Paper Shuffle', mo_papers, 3, P('3F3F46', 'E11D48', 'F4F4F5'), TL),
    ('sunburst', 'Sunburst', mo_sunburst, 6, P('7C2D12', 'F59E0B', 'FEF3C7'), ((0.9, 0.6, 7.6, 0.95), (0.9, 1.8, 7.6, 4.8))),
    ('confetti', 'Confetti Drift', mo_confetti, 5, P('1F2937', 'EC4899', 'FDE68A', accent2='22D3EE'), TL),
    ('sweep', 'Square Sweep', mo_sweep, 4, P('0B3B2E', '10B981', 'D1FAE5'), ((0.9, 0.6, 7.6, 0.95), (0.9, 1.8, 7.6, 4.8))),
    ('lenses', 'Sliding Lenses', mo_lenses, 4, P('172554', '60A5FA', 'E0E7FF', accent2='F472B6'), TL),
    ('bricks', 'Brick Wall', mo_bricks, 6, P('450A0A', 'DC2626', 'FEE2E2'), ((0.9, 0.6, 8.0, 0.95), (0.9, 1.8, 8.0, 4.8))),
]

# a cover picture on a Morph design whose shapes run along the top or the bottom: clear of them (the rest: on the right)
MORPH_PIC = {'puzzle': (8.6, 2.6, 4.0, 4.0), 'wavestrips': (8.6, 2.6, 4.0, 4.0), 'carousel': (8.6, 0.6, 4.0, 3.9), 'progress': (8.6, 0.8, 4.0, 5.2),
             'ribbon': (8.6, 1.1, 4.0, 4.8), 'cascade': (8.6, 3.35, 4.0, 3.3), 'spiral': (8.0, 1.0, 3.8, 5.4)}

# ======================================================================== the designs, by template
# (id, group, name, palette, motif, title font). The first 12 student designs are not here: they stay as they were.
REDONE = [
    ('student-modern-neon-mint', 'student', 'Neon Mint', P('0F3D3E', '10B981', 'A7F3D0', bg='F0FDF9'), m_hexes, 'light'),
    ('student-modern-violet-pop', 'student', 'Violet Pop', P('2E1065', 'A855F7', 'E9D5FF'), m_memphis, 'light'),
    ('student-modern-coral-wave', 'student', 'Coral Wave', P('7C2D12', 'FB7185', 'FECDD3', bg='FFF7F5'), m_waves, 'light'),
    ('student-modern-midnight-study', 'student', 'Midnight Study', P('0F172A', '6366F1', 'C7D2FE'), m_grid, 'light'),
    ('student-modern-sunrise', 'student', 'Sunrise', P('9A3412', 'F59E0B', 'FED7AA', bg='FFFBF5'), m_sun, 'light'),
    ('student-modern-aqua-photo', 'student', 'Aqua Photo', P('0E7490', '22D3EE', 'CFFAFE', bg='F5FEFF'), m_polaroid, 'light'),
    ('student-modern-mono-bold', 'student', 'Mono Bold', P('18181B', 'FACC15', 'FEF08A'), m_bauhaus, 'grotesk'),
    ('student-modern-rosewood', 'student', 'Rosewood', P('4C0519', 'F472B6', 'FBCFE8', bg='FFF5F8'), m_leaves, 'serif'),
    ('tutor-modern-royal-purple', 'tutor', 'Royal Purple', P('3B0764', 'A78BFA', 'DDD6FE'), m_frame, 'serif'),
    ('tutor-modern-sunset-coral', 'tutor', 'Sunset Coral', P('9A3412', 'EA580C', 'FFEDD5', bg='FFF8F3'), m_diagonal, 'light'),
    ('tutor-modern-ocean-teal', 'tutor', 'Ocean Teal', P('134E4A', '2DD4BF', '99F6E4', bg='F7FBFA'), m_rings, 'light'),
    ('tutor-modern-graphite-gold', 'tutor', 'Graphite Gold', P('1C1917', 'CA8A04', 'FEF3C7'), m_ticket, 'serif'),
    ('tutor-modern-forest-sage', 'tutor', 'Forest Sage', P('14532D', '4D7C0F', 'D9F99D'), m_mountain, 'light'),
    ('tutor-modern-midnight-indigo', 'tutor', 'Midnight Indigo', P('1E1B4B', '6366F1', 'C7D2FE'), m_glow, 'light'),
    ('tutor-modern-terracotta', 'tutor', 'Terracotta', P('7C2D12', 'C2410C', 'FED7AA', bg='FFFBF7'), m_arch, 'serif'),
    ('tutor-modern-sky-minimal', 'tutor', 'Sky Minimal', P('0C4A6E', '0284C7', 'E0F2FE'), m_card, 'light'),
    ('tutor-modern-berry-plum', 'tutor', 'Berry Plum', P('4A044E', 'C026D3', 'F5D0FE'), m_bubbles, 'light'),
    ('tutor-lecture-maroon', 'tutor', 'Lecture Maroon', P('450A0A', 'B91C1C', 'FECACA'), m_banner, 'serif'),
    ('tutor-lecture-forest', 'tutor', 'Lecture Forest', P('052E16', '15803D', 'BBF7D0'), m_sidebar, 'light'),
    ('tutor-lecture-purple', 'tutor', 'Lecture Purple', P('2E1065', '7C3AED', 'DDD6FE'), m_timeline, 'light'),
    ('tutor-lecture-slate', 'tutor', 'Lecture Slate', P('0F172A', '475569', 'E2E8F0', accent2='F59E0B'), m_split, 'grotesk'),
]
# new: for students and tutors alike ("student-scene-…" and "tutor-scene-…")
NEW = [
    ('tiles', 'Mosaic', P('1E3A8A', 'F97316', 'FED7AA', accent2='14B8A6'), m_tiles, 'grotesk'),
    ('ribbon', 'Ribbon', P('831843', 'DB2777', 'FCE7F3'), m_ribbon, 'light'),
    ('dots', 'Dot Grid', P('064E3B', '059669', 'D1FAE5'), m_dots, 'light'),
    ('stripes', 'Bold Stripes', P('0F172A', 'F43F5E', 'FFE4E6'), m_stripes, 'grotesk'),
    ('blocks', 'Colour Blocks', P('1E293B', '0EA5E9', 'E0F2FE', accent2='FBBF24'), m_blocks, 'light'),
    ('pills', 'Rounded Pills', P('3730A3', '818CF8', 'E0E7FF', accent2='F472B6'), m_pills, 'light'),
    ('halftone', 'Halftone', P('7F1D1D', 'EF4444', 'FEE2E2'), m_halftone, 'grotesk'),
    ('lens', 'Soft Lens', P('312E81', 'A78BFA', 'EDE9FE'), m_lens, 'light'),
    ('notebook', 'Notebook', P('1E3A5F', 'E11D48', 'DBEAFE'), m_notebook, 'light'),
    ('scrapbook', 'Scrapbook', P('7C2D12', 'FB923C', 'FFEDD5', bg='FFFBF7'), m_tape, 'serif'),
    ('chevrons', 'Chevrons', P('4A044E', 'C026D3', 'F5D0FE'), m_chevrons, 'light'),
    ('stack', 'Card Stack', P('0E7490', '22D3EE', 'CFFAFE', bg='F5FEFF'), m_stack, 'light'),
    ('corners', 'Corners', P('4C0519', 'E11D48', 'FFE4E6', bg='FFF5F7'), m_corners, 'serif'),
    ('quarter', 'Quarter Circles', P('0C2D57', '2563EB', 'D3E7FF', accent2='56CCF2'), m_quarter, 'light'),
    ('editorial', 'Editorial', P('1F2328', 'DC2626', 'E6E8EB'), m_magazine, 'serif'),
    ('halo', 'Halo', P('3B2F0B', 'D97706', 'FDE68A'), m_halo, 'serif'),
    ('staircase', 'Staircase', P('7C2D12', 'EA580C', 'FFEDD5', bg='FFFBF7'), m_steps, 'serif'),
    ('tabs', 'Index Tabs', P('134E4A', '0D9488', 'CCFBF1'), m_tabs, 'light'),
]
TITLE_FONTS = {'light': 'Calibri Light', 'serif': 'Georgia', 'grotesk': 'Trebuchet MS'}


def _base(group, name, p, tf, sk):
    return {'group': group, 'style': 'modern', 'name': name, 'cover': 'scene', 'chrome': 'scene', 'media': 'large', 'tfont': TITLE_FONTS[tf], 'sk': sk,
            'ink': p['ink'], 'ink2': p['ink2'], 'accent': p['accent'], 'emph': p['emph'], 'soft': p['soft'], 'bg': p['bg'], 'text': p['text'], 'muted': p['muted']}


def _colours(g, p):
    out = dict(g)
    for k in ('tc', 'sc', 'mc'):
        if k in out:
            out[k] = C(p, out[k])
    return out


def scenes():
    """{scene key: {decor | states, geo, cg, sg}} and [(template id, group, name, palette, title font, scene key, morph)]."""
    S, L = {}, []
    for id_, group, name, p, motif, tf in REDONE:
        m = motif(p)
        S[id_] = {'decor': {k: resolve(v, p) for k, v in m['decor'].items()}, 'geo': _colours(m['geo'], p), 'cg': _colours(m['cg'], p), 'sg': _colours(m['sg'], p)}
        L.append((id_, group, name, p, tf, id_, False))
    for key, name, p, motif, tf in NEW:
        m = motif(p)
        sk = 'scene-' + key
        S[sk] = {'decor': {k: resolve(v, p) for k, v in m['decor'].items()}, 'geo': _colours(m['geo'], p), 'cg': _colours(m['cg'], p), 'sg': _colours(m['sg'], p)}
        for group in ('student', 'tutor'):
            L.append(('%s-%s' % (group, sk), group, name, p, tf, sk, False))
    for key, name, fn, K, p, (title, body) in MORPHS:
        sk = 'morph-' + key
        ty = title[1] if title[1] > 1.5 else 1.3
        x, w = title[0], title[2]
        S[sk] = {'states': [resolve(fn(p, i), p) for i in range(K)], 'geo': _colours(G(title=title, body=body), p),
                 'cg': _colours(CG((x, ty, w, 2.3), (x, ty + 2.4, w, 0.7), (x, ty + 3.1, w, 0.5), pic=MORPH_PIC.get(key)), p),
                 'sg': _colours(SG((x, max(1.4, ty), w, 2.2), (x, max(1.4, ty) + 2.3, w, 0.8)), p)}
        for group in ('student', 'tutor'):
            L.append(('%s-%s' % (group, sk), group, name, p, 'light', sk, True))
    return S, L


def build():
    """Every design as a template, in the server's form (hex without #)."""
    S, L = scenes()
    T = {}
    for id_, group, name, p, tf, sk, morph in L:
        t = _base(group, name, p, tf, sk)
        t.update(S[sk])
        if morph:
            t['morph'] = True
        T[id_] = t
    return T


def decor_for(t, kind, index):
    """The shapes on a slide: a design's own for its kind; a Morph design's step (a section's on its dark colour)."""
    if t.get('states'):
        st = t['states'][index % len(t['states'])]
        if kind in ('section', 'end'):
            st = [dict(s, f=t['ink']) if s.get('n') == 'bg' else s for s in st]
        return st
    d = t.get('decor') or {}
    if index == 0:
        return d.get('cover') or []
    if kind in ('section', 'end'):
        return d.get('section') or []
    return d.get('content') or []


def cover_boxes(t, picture):
    """The cover's title, line under it and small line (and the picture's box) — narrowed beside a picture."""
    cg = t['cg']
    boxes = [list(cg['t']), list(cg['s']), list(cg['m'])]
    al = cg.get('al', 'left')
    pb = cg['p'] if picture else None
    if pb:
        for b in boxes:
            if b[0] < pb[0] and b[0] + b[2] > pb[0] - 0.3 and b[1] < pb[1] + pb[3] and b[1] + b[3] > pb[1]:
                b[2] = max(2.0, pb[0] - 0.3 - b[0])
                al = 'left'
    return boxes, al, pb


# ======================================================================== the readability check
def _local(s, px, py):
    cx, cy = s['x'] + s['w'] / 2, s['y'] + s['h'] / 2
    a = math.radians(s.get('r') or 0)
    dx, dy = px - cx, py - cy
    u = dx * math.cos(a) + dy * math.sin(a)
    v = -dx * math.sin(a) + dy * math.cos(a)
    return u + s['w'] / 2, v + s['h'] / 2


STAR = []
for _i in range(10):
    _a = math.radians(-90 + _i * 36)
    _r = 0.5 if _i % 2 == 0 else 0.5 * 0.382
    STAR.append((round(0.5 + _r * math.cos(_a), 4), round((0.5 + _r * math.sin(_a)) / 0.9045, 4)))


def _in_poly(pts, x, y):
    inside = False
    for i in range(len(pts)):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % len(pts)]
        if (y1 > y) != (y2 > y) and x < (x2 - x1) * (y - y1) / (y2 - y1) + x1:
            inside = not inside
    return inside


def covers(s, px, py):
    """Is (px, py) inside shape s (as drawn: the same outlines the app and PowerPoint use)."""
    u, v = _local(s, px, py)
    w, h, k = s['w'], s['h'], s['k']
    if u < 0 or v < 0 or u > w or v > h:
        return False
    ss = min(w, h)
    if k == 'rect':
        rr = min(s.get('rr') or 0, ss / 2)
        if rr:
            cx = rr if u < rr else w - rr if u > w - rr else None
            cy = rr if v < rr else h - rr if v > h - rr else None
            if cx is not None and cy is not None and (u - cx) ** 2 + (v - cy) ** 2 > rr ** 2:
                return False
        return True
    if k == 'oval':
        return ((u - w / 2) / (w / 2)) ** 2 + ((v - h / 2) / (h / 2)) ** 2 <= 1
    if k == 'ring':
        th = s.get('th') or ss * 0.25
        if ((u - w / 2) / (w / 2)) ** 2 + ((v - h / 2) / (h / 2)) ** 2 > 1:
            return False
        iw, ih = w / 2 - th, h / 2 - th
        return iw <= 0 or ih <= 0 or ((u - w / 2) / iw) ** 2 + ((v - h / 2) / ih) ** 2 >= 1
    if k == 'tri':
        return abs(u - w / 2) <= (v / h) * (w / 2)
    if k == 'rtri':
        return u / w <= v / h
    if k == 'para':
        x1 = 0.25 * ss
        return x1 * (1 - v / h) <= u <= w - x1 * v / h
    if k == 'hex':
        x1 = 0.25 * w
        e = x1 * abs(h / 2 - v) / (h / 2)
        return e <= u <= w - e
    if k == 'diam':
        return abs(u - w / 2) / (w / 2) + abs(v - h / 2) / (h / 2) <= 1
    if k == 'plus':
        x1 = 0.25 * ss
        return (x1 <= u <= w - x1) or (x1 <= v <= h - x1)
    if k == 'star':
        return _in_poly(STAR, u / w, v / h)
    return False


def _fill_at(s, px, py):
    if 'g' not in s:
        return s['f']
    c1, c2, ang = s['g']
    a = math.radians(ang)
    dx, dy = math.cos(a), -math.sin(a)
    corners = [(s['x'], s['y']), (s['x'] + s['w'], s['y']), (s['x'], s['y'] + s['h']), (s['x'] + s['w'], s['y'] + s['h'])]
    proj = [cx * dx + cy * dy for cx, cy in corners]
    lo, hi = min(proj), max(proj)
    t = 0 if hi == lo else ((px * dx + py * dy) - lo) / (hi - lo)
    return mix(c1, c2, max(0, min(1, t)))


def colour_at(shapes, px, py):
    c = 'FFFFFF'
    for s in shapes:
        if covers(s, px, py):
            a = s.get('a', 1)
            c = mix(c, _fill_at(s, px, py), a) if a < 1 else _fill_at(s, px, py)
    return c


def worst(shapes, box, colour, step=0.12):
    x, y, w, h = box
    lo = 99
    nx, ny = max(2, int(w / step)), max(2, int(h / step))
    for i in range(nx + 1):
        for j in range(ny + 1):
            c = colour_at(shapes, x + w * (0.02 + 0.96 * i / nx), y + h * (0.04 + 0.92 * j / ny))
            lo = min(lo, contrast(colour, c))
    return lo


def audit(T=None):
    """[(template, slide, words, contrast, needed)] for every place words go that doesn't read at AAA."""
    T = T or build()
    bad = []

    def check(tid, slide, what, shapes, box, colour, need):
        v = worst(shapes, box, colour)
        if v < need:
            bad.append((tid, slide, what, round(v, 2), need))

    for tid, t in T.items():
        if tid.startswith('tutor-') and t.get('sk', '').startswith(('scene-', 'morph-')):
            continue          # the same design as its student twin
        cover = decor_for(t, 'content', 0)
        cg = t['cg']
        for pic in (False, True):
            (tb, sb, mb), al, pb = cover_boxes(t, pic)
            tag = 'cover' + (' with a picture' if pic else '')
            check(tid, tag, 'title', cover, tb, cg['tc'], 4.5)
            check(tid, tag, 'line under it', cover, sb, cg['sc'], 4.5)
            check(tid, tag, 'small line', cover, mb, cg['mc'], 7.0)
        g = t['geo']
        steps = len(t['states']) if t.get('states') else 1
        for i in range(1, steps + 1):
            content = decor_for(t, 'content', i)
            check(tid, 'slide %d' % (i + 1), 'title', content, g['title'], g.get('tc', t['ink']), 4.5)
            check(tid, 'slide %d' % (i + 1), 'points', content, g['body'], t['text'], 7.0)
            check(tid, 'slide %d' % (i + 1), 'footer', content, g['foot'], t['muted'], 7.0)
            check(tid, 'slide %d' % (i + 1), 'quote', content, g['quote'], t['ink'], 4.5)
            sec = decor_for(t, 'section', i)
            check(tid, 'section %d' % (i + 1), 'title', sec, t['sg']['t'], t['sg']['tc'], 4.5)
            check(tid, 'section %d' % (i + 1), 'line under it', sec, t['sg']['s'], t['sg']['sc'], 4.5)
    return bad


TEMPLATES = build()

if __name__ == '__main__':
    import json
    import sys
    print(len(TEMPLATES), 'designs;', sum(1 for t in TEMPLATES.values() if t.get('morph')), 'Morph;', len(json.dumps(scenes()[0])) // 1024, 'KB of shapes')
    if '--audit' in sys.argv:
        for b in audit():
            print(b)
