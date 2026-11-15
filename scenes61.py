"""6.1 — fifty more Slide Studio designs (the owner, 15 Nov 2026 02:19: "add 50 more templates").

The same form as scenes315 (one source for both: the server draws these shapes into the .pptx, the app draws the same
shapes on its slides; tools/scenes61_app.py writes the app's copy). Each design is for students and tutors alike, with
its own cover, ordinary slide and section slide. The words of an ordinary slide never sit on a shape: its shapes stay in
the top-right corner (the title stops short of it) or under the last line, clear of the footer. Every place words go
reads at AAA (scenes315.audit).

Groups, as the Design gallery shows them: subjects (lab, maths, clinic, books…), calm, bold, and dark & dramatic (for
the dramatic transitions).
"""
import math
from scenes315 import (W, H, P, R, O, Tri, RT, Hx, Dm, Rg, St, Pl, Pa, circ, full, G, CG, DARK, SG, SGL, LEFT, mix, readable,
                       resolve, _colours, _base)

TG = (0.9, 0.55, 10.0, 0.95)          # an ordinary slide's title, short of the corner where its shapes are


def _d(cover, content, section, cg, sg, geo=None):
    return dict(decor=dict(cover=cover, content=content, section=section), geo=geo or G(title=TG), cg=cg, sg=sg)


SG_L = ((0.9, 2.0, 7.0, 2.2), (0.9, 4.3, 7.0, 0.8))      # a section's words on the left
SG_R = ((6.4, 2.0, 6.2, 2.2), (6.4, 4.3, 6.2, 0.8))      # …on the right


# ======================================================================== subjects
def s_blueprint(p):
    def plan(x, y, k, f, w=1.25):
        return [R(x, y, 4.2 * k, 3.0 * k, 'ink', l=[f, w]), R(x + 2.6 * k, y, 1.6 * k, 1.7 * k, 'ink', l=[f, w]),
                R(x, y + 1.9 * k, 2.2 * k, 1.1 * k, 'ink', l=[f, w]), Rg(x + 0.5 * k, y + 0.35 * k, 1.1 * k, 1.1 * k, f, th=0.03, a=0.8),
                R(x, y - 0.35 * k, 4.2 * k, 0.02, f, a=0.8), Tri(x - 0.05, y - 0.42 * k, 0.14, 0.14, f, r=-90), Tri(x + 4.2 * k - 0.09, y - 0.42 * k, 0.14, 0.14, f, r=90)]
    grid = [R(i * 0.5, 0, 0.012, H, 'soft', a=0.16) for i in range(27)] + [R(0, j * 0.5, W, 0.012, 'soft', a=0.16) for j in range(16)]
    cover = [full('ink')] + grid + plan(8.0, 2.0, 1.1, 'soft')
    content = [full('bg'), R(11.25, 0.35, 1.7, 1.05, 'bg', l=['ink2', 1]), R(12.25, 0.35, 0.7, 0.55, 'bg', l=['ink2', 1]), R(11.25, 0.18, 1.7, 0.015, 'accent')]
    section = [full('ink')] + grid + plan(1.0, 2.2, 1.0, 'soft')
    return _d(cover, content, section, DARK(*LEFT(6.6), pic=(8.0, 1.5, 4.6, 3.3), pr=0.0), SG(*SG_R))


def s_chalk(p):
    wood = '7C4A1E'
    cover = [full('ink'), R(0, 6.75, W, 0.75, wood), R(0, 6.75, W, 0.08, 'accent', a=0.35), O(8.3, 1.1, 4.4, 1.6, 'white', a=0.05), O(9.0, 3.6, 3.6, 1.2, 'white', a=0.05),
             R(8.6, 2.0, 3.6, 0.05, 'white', a=0.7, r=-3), R(8.9, 2.5, 2.6, 0.05, 'white', a=0.6, r=-3), circ(11.7, 4.6, 1.0, 'accent', a=0.85),
             R(9.0, 4.4, 1.6, 0.05, 'white', a=0.55), R(9.0, 4.75, 1.1, 0.05, 'white', a=0.55), R(3.1, 6.6, 0.9, 0.14, 'white', rr=0.07), R(4.3, 6.62, 0.6, 0.12, 'accent', rr=0.06)]
    content = [full('bg'), R(0, 0, W, 0.16, 'ink'), R(11.6, 0.42, 1.3, 0.05, 'ink2', r=-3), R(11.8, 0.7, 0.9, 0.05, 'ink2', r=-3), circ(12.75, 1.1, 0.32, 'accent')]
    section = [full('ink'), R(0, 6.75, W, 0.75, wood), O(1.0, 1.4, 5.6, 2.2, 'white', a=0.05), R(8.4, 6.6, 1.0, 0.14, 'white', rr=0.07)]
    return _d(cover, content, section, DARK(*LEFT(6.8), pic=(8.4, 1.0, 4.0, 4.6), pr=0.1),
              SG((0.9, 1.6, 11.0, 2.2), (0.9, 3.9, 11.0, 0.8)), geo=G(title=TG, body=(0.9, 1.75, 11.5, 4.85)))


def s_graph(p):
    def grid(x0, y0, w, h, s, a):
        return [R(x0 + i * s, y0, 0.01, h, 'soft', a=a) for i in range(int(w / s) + 1)] + [R(x0, y0 + j * s, w, 0.01, 'soft', a=a) for j in range(int(h / s) + 1)]
    pts = [(8.2, 5.6), (9.1, 5.0), (9.9, 5.25), (10.7, 3.9), (11.5, 3.2), (12.3, 1.7)]
    curve = []
    for (x1, y1), (x2, y2) in zip(pts, pts[1:]):
        L = math.hypot(x2 - x1, y2 - y1)
        curve.append(R((x1 + x2) / 2 - L / 2, (y1 + y2) / 2 - 0.03, L, 0.06, 'accent', r=math.degrees(math.atan2(y2 - y1, x2 - x1))))
    dots = [circ(x, y, 0.2, 'ink') for x, y in pts]
    cover = [full('bg')] + grid(7.6, 0.8, 5.2, 5.6, 0.4, 0.9) + [R(7.6, 6.4, 5.2, 0.03, 'ink'), R(7.6, 0.8, 0.03, 5.6, 'ink')] + curve + dots
    content = [full('bg')] + grid(11.3, 0.3, 1.6, 1.2, 0.2, 0.9) + [R(11.35, 1.12, 0.6, 0.04, 'accent', r=-28), R(11.85, 0.82, 0.6, 0.04, 'accent', r=-40)]
    section = [full('ink')] + [R(i * 0.4, 0, 0.01, H, 'soft', a=0.18) for i in range(34)] + [R(0, j * 0.4, W, 0.01, 'soft', a=0.18) for j in range(19)]
    return _d(cover, content, section, CG(*LEFT(6.3), pic=(8.0, 1.2, 4.4, 4.6), pr=0.0), SG((0.9, 2.0, 11.0, 2.2), (0.9, 4.3, 11.0, 0.8)))


def s_news(p):
    col = lambda x, y, n, w: [R(x, y + i * 0.22, w * (0.6 if i == n - 1 else 1), 0.07, 'soft') for i in range(n)]
    cover = [full('bg'), R(0.9, 0.55, 11.53, 0.09, 'ink'), R(0.9, 0.7, 11.53, 0.025, 'ink'), R(0.9, 6.5, 11.53, 0.025, 'ink'),
             R(8.4, 1.6, 0.02, 4.6, 'soft'), R(10.5, 1.6, 0.02, 4.6, 'soft')] + col(8.6, 1.7, 18, 1.75) + col(10.7, 1.7, 14, 1.7) + [R(10.7, 4.95, 1.7, 1.2, 'accent', a=0.85)]
    content = [full('bg'), R(0.9, 0.3, 11.53, 0.05, 'ink'), R(0.9, 1.55, 11.53, 0.02, 'ink')]
    section = [full('ink'), R(0.9, 1.55, 11.53, 0.06, 'white'), R(0.9, 1.68, 11.53, 0.02, 'white'), R(0.9, 5.6, 11.53, 0.02, 'white')]
    return _d(cover, content, section, CG((0.9, 1.3, 7.2, 2.7), (0.9, 4.1, 7.2, 0.7), (0.9, 4.8, 7.2, 0.5), pic=(8.6, 1.7, 3.8, 3.0), pr=0.0),
              SG((0.9, 2.0, 11.0, 2.2), (0.9, 4.3, 11.0, 0.8)), geo=G(title=(0.9, 0.5, 11.5, 0.95)))


def s_shelf(p):
    cols = ['ink', 'accent', 'ink2', 'soft', 'accent2', 'ink', 'accent', 'soft', 'ink2', 'accent2', 'ink']
    hs = [3.2, 2.7, 3.5, 2.4, 3.0, 2.2, 3.3, 2.8, 2.5, 3.4, 2.9]
    books, x = [], 7.5
    for i, (c, h) in enumerate(zip(cols, hs)):
        w = 0.42 + (i % 3) * 0.08
        books.append(R(x, 6.2 - h, w, h, c, rr=0.04, r=(-6 if i == 7 else 0)))
        books.append(R(x + 0.06, 6.2 - h + 0.35, w - 0.12, 0.05, 'white', a=0.6))
        x += w + 0.04
    cover = [full('bg')] + books + [R(7.2, 6.2, 6.1, 0.22, 'ink'), R(7.2, 6.42, 6.1, 0.08, 'accent')]
    content = [full('bg')] + [R(11.45 + i * 0.27, 0.35 + (i % 2) * 0.18, 0.22, 1.0 - (i % 2) * 0.18, cols[i], rr=0.03) for i in range(5)] + [R(11.35, 1.35, 1.6, 0.07, 'ink')]
    section = [full('ink')] + [dict(b, x=b['x'] - 6.4) for b in books] + [R(0.8, 6.2, 6.1, 0.22, 'accent')]
    return _d(cover, content, section, CG(*LEFT(6.0), pic=(8.0, 0.8, 4.4, 2.6), pr=0.15), SG((7.6, 2.0, 5.0, 2.2), (7.6, 4.3, 5.0, 0.8)))


def s_grad(p):
    def cap(cx, cy, k, f, g):
        return [Dm(cx - 2.0 * k, cy - 0.75 * k, 4.0 * k, 1.5 * k, f), R(cx - 1.15 * k, cy + 0.3 * k, 2.3 * k, 0.9 * k, f, rr=0.12 * k),
                circ(cx, cy, 0.2 * k, g), R(cx + 0.05 * k, cy - 0.04 * k, 1.6 * k, 0.06 * k, g), R(cx + 1.62 * k, cy - 0.04 * k, 0.06 * k, 1.3 * k, g),
                O(cx + 1.5 * k, cy + 1.2 * k, 0.32 * k, 0.5 * k, g)]
    sparks = [St(8.1, 0.8, 0.45, 0.45, 'accent', a=0.9), St(12.2, 1.2, 0.3, 0.3, 'accent', a=0.8), St(11.6, 5.6, 0.35, 0.35, 'accent', a=0.7), circ(8.6, 5.3, 0.14, 'accent')]
    cover = [full('ink')] + cap(10.3, 3.0, 1.15, 'ink2', 'accent') + sparks
    content = [full('bg')] + cap(12.2, 0.75, 0.27, 'ink', 'accent')
    section = [full('ink')] + cap(3.6, 3.2, 1.0, 'ink2', 'accent') + [dict(s, x=s['x'] - 7.2) for s in sparks]
    return _d(cover, content, section, DARK(*LEFT(6.4), pic=(8.6, 1.2, 3.8, 3.8), pk='circle'), SG(*SG_R))


def s_helix(p):
    def helix(x0, y0, length, amp, n, horizontal=True, k=1.0):
        out = []
        for i in range(n):
            t = i / (n - 1)
            ph = t * math.pi * 3
            a, b = math.sin(ph) * amp, -math.sin(ph) * amp
            if horizontal:
                x = x0 + t * length
                out.append(R(x - 0.015, y0 + min(a, b), 0.03, abs(a - b) + 0.001, 'soft', a=0.9))
                out.append(circ(x, y0 + a, 0.22 * k, 'accent'))
                out.append(circ(x, y0 + b, 0.22 * k, 'ink2'))
            else:
                y = y0 + t * length
                out.append(R(x0 + min(a, b), y - 0.015, abs(a - b) + 0.001, 0.03, 'soft', a=0.9))
                out.append(circ(x0 + a, y, 0.22 * k, 'accent'))
                out.append(circ(x0 + b, y, 0.22 * k, 'ink2'))
        return out
    cover = [full('bg')] + helix(10.3, 0.7, 6.1, 1.2, 15, horizontal=False)
    content = [full('bg')] + helix(11.4, 0.75, 1.5, 0.28, 7, k=0.5)
    section = [full('ink')] + helix(1.0, 6.0, 11.3, 0.55, 22)
    return _d(cover, content, section, CG(*LEFT(7.2), pic=(8.0, 2.0, 1.4, 1.4), pk='circle'), SG((0.9, 1.6, 11.0, 2.2), (0.9, 3.9, 11.0, 0.8)))


def s_circuit(p):
    def traces(x, y, k):
        return [R(x, y, 3.6 * k, 0.04, 'accent'), R(x + 3.6 * k, y, 0.04, 1.4 * k, 'accent'), R(x + 3.6 * k, y + 1.4 * k, 1.2 * k, 0.04, 'accent'),
                R(x + 0.8 * k, y + 2.2 * k, 3.0 * k, 0.04, 'soft', a=0.8), R(x + 0.8 * k, y + 2.2 * k, 0.04, 1.8 * k, 'soft', a=0.8),
                R(x + 0.8 * k, y + 4.0 * k, 3.8 * k, 0.04, 'accent', a=0.9), R(x + 2.4 * k, y + 0.9 * k, 1.6 * k, 1.1 * k, 'ink2', rr=0.08, l=['accent', 1]),
                Rg(x - 0.12, y - 0.12, 0.28, 0.28, 'accent', th=0.05), Rg(x + 4.8 * k - 0.14, y + 1.4 * k - 0.14, 0.28, 0.28, 'accent', th=0.05),
                Rg(x + 0.8 * k - 0.14, y + 4.0 * k - 0.14, 0.28, 0.28, 'soft', th=0.05), circ(x + 4.6 * k, y + 4.02 * k, 0.2, 'accent')]
    cover = [full('ink')] + traces(7.9, 1.2, 1.05)
    content = [full('bg'), R(11.3, 0.5, 1.2, 0.03, 'accent'), R(12.5, 0.5, 0.03, 0.7, 'accent'), R(11.7, 1.2, 0.83, 0.03, 'ink2'),
               Rg(11.2, 0.42, 0.18, 0.18, 'accent', th=0.04), Rg(11.6, 1.12, 0.18, 0.18, 'ink2', th=0.04)]
    section = [full('ink')] + traces(0.9, 1.4, 1.0)
    return _d(cover, content, section, DARK(*LEFT(6.4), pic=(8.6, 1.0, 3.8, 3.4), pr=0.1), SG(*SG_R))


def s_atom(p):
    def atom(cx, cy, k):
        return [Rg(cx - 2.4 * k, cy - 0.8 * k, 4.8 * k, 1.6 * k, 'accent', th=0.05 * k + 0.01, r=a) for a in (0, 60, 120)] + \
               [circ(cx, cy, 0.55 * k, 'ink2'), circ(cx + 2.38 * k, cy, 0.22 * k, 'accent'), circ(cx - 1.2 * k, cy - 1.75 * k, 0.22 * k, 'accent'),
                circ(cx - 1.18 * k, cy + 1.72 * k, 0.22 * k, 'soft')]
    cover = [full('bg')] + atom(10.4, 3.6, 1.2)
    content = [full('bg')] + atom(12.15, 0.85, 0.3)
    section = [full('ink')] + atom(3.4, 3.75, 1.1)
    return _d(cover, content, section, CG(*LEFT(6.6), pic=(8.9, 2.1, 3.0, 3.0), pk='circle'), SG(*SG_R))


def s_clinic(p):
    def plus(x, y, s, f, a=None):
        return Pl(x, y, s, s, f, a=a)
    cover = [full('bg'), R(8.0, 0.9, 4.6, 5.6, 'soft', rr=0.5), plus(9.1, 2.0, 2.4, 'accent'), plus(11.6, 0.6, 0.9, 'ink2', 0.9), plus(8.4, 5.2, 0.7, 'ink2', 0.8),
             R(8.6, 5.0, 3.4, 0.07, 'ink2', rr=0.035), R(8.6, 5.3, 2.4, 0.07, 'ink2', rr=0.035, a=0.7)]
    content = [full('bg'), plus(12.2, 0.4, 0.7, 'accent'), plus(11.65, 0.9, 0.35, 'ink2', 0.8)]
    section = [full('ink'), plus(1.4, 1.6, 3.6, 'accent', 0.9), plus(4.6, 4.6, 1.2, 'soft', 0.8)]
    return _d(cover, content, section, CG(*LEFT(6.3), pic=(8.4, 1.3, 3.8, 3.3), pr=0.25), SG(*SG_R))


def s_pulse(p):
    seg = [(7.2, 3.6), (8.6, 3.6), (8.9, 3.2), (9.2, 3.6), (9.6, 3.6), (9.85, 4.3), (10.2, 1.5), (10.6, 5.0), (10.9, 3.6), (11.3, 3.6), (11.65, 3.0), (12.0, 3.6), (W, 3.6)]
    line = []
    for (x1, y1), (x2, y2) in zip(seg, seg[1:]):
        L = math.hypot(x2 - x1, y2 - y1)
        line.append(R((x1 + x2) / 2 - L / 2, (y1 + y2) / 2 - 0.045, L, 0.09, 'accent', r=math.degrees(math.atan2(y2 - y1, x2 - x1))))
    cover = [full('ink')] + [R(7.2, 1.0 + j * 0.65, 6.1, 0.01, 'soft', a=0.15) for j in range(9)] + line
    small = [(11.2, 0.95), (11.7, 0.95), (11.85, 0.55), (12.05, 1.3), (12.2, 0.95), (12.95, 0.95)]
    content = [full('bg')]
    for (x1, y1), (x2, y2) in zip(small, small[1:]):
        L = math.hypot(x2 - x1, y2 - y1)
        content.append(R((x1 + x2) / 2 - L / 2, (y1 + y2) / 2 - 0.025, L, 0.05, 'accent', r=math.degrees(math.atan2(y2 - y1, x2 - x1))))
    section = [full('ink')] + [dict(s, y=s['y'] + 2.0) for s in line]
    return _d(cover, content, section, DARK(*LEFT(5.9), pic=(8.6, 0.8, 3.8, 2.0), pr=0.1), SG((0.9, 0.7, 11.0, 1.9), (0.9, 2.6, 11.0, 0.7)))


def s_cell(p):
    def cell(cx, cy, k):
        return [O(cx - 2.6 * k, cy - 2.0 * k, 5.2 * k, 4.0 * k, 'soft', l=['accent', 2]), circ(cx + 0.3 * k, cy - 0.1 * k, 1.5 * k, 'accent', a=0.85),
                circ(cx + 0.45 * k, cy - 0.2 * k, 0.5 * k, 'ink2'), O(cx - 1.9 * k, cy + 0.6 * k, 1.0 * k, 0.45 * k, 'ink2', r=-20, a=0.85),
                O(cx + 1.2 * k, cy + 1.1 * k, 0.9 * k, 0.4 * k, 'accent', r=25), circ(cx - 1.4 * k, cy - 1.0 * k, 0.35 * k, 'accent', a=0.8),
                circ(cx + 1.7 * k, cy - 1.1 * k, 0.25 * k, 'ink2', a=0.8)]
    cover = [full('bg')] + cell(10.2, 3.6, 1.0)
    content = [full('bg')] + cell(12.15, 0.85, 0.27)
    section = [full('ink')] + cell(3.6, 3.75, 1.0)
    return _d(cover, content, section, CG(*LEFT(6.3), pic=(8.6, 2.0, 3.2, 3.2), pk='circle'), SG(*SG_R))


def s_maths(p):
    def sym(kind, x, y, s, f, a=None):
        if kind == '+':
            return [Pl(x, y, s, s, f, a=a)]
        if kind == 'x':
            return [Pl(x, y, s, s, f, a=a, r=45)]
        if kind == '-':
            return [R(x, y + s * 0.4, s, s * 0.2, f, a=a)]
        if kind == '=':
            return [R(x, y + s * 0.25, s, s * 0.17, f, a=a), R(x, y + s * 0.58, s, s * 0.17, f, a=a)]
        return [R(x, y + s * 0.42, s, s * 0.16, f, a=a), circ(x + s / 2, y + s * 0.18, s * 0.2, f, a=a), circ(x + s / 2, y + s * 0.82, s * 0.2, f, a=a)]
    lay = [('+', 8.2, 1.0, 1.3, 'accent'), ('x', 10.4, 0.9, 1.1, 'ink2'), ('-', 11.9, 1.3, 0.9, 'soft'), ('/', 8.4, 3.2, 1.1, 'ink2'), ('=', 10.2, 3.1, 1.4, 'accent'),
           ('+', 12.0, 3.4, 0.7, 'soft'), ('x', 8.9, 5.2, 0.8, 'soft'), ('-', 10.6, 5.5, 1.2, 'ink2'), ('/', 12.2, 5.4, 0.7, 'accent')]
    cover = [full('bg')] + [s for k, x, y, z, f in lay for s in sym(k, x, y, z, f)]
    content = [full('bg')] + sym('+', 11.5, 0.4, 0.5, 'accent') + sym('=', 12.25, 0.45, 0.5, 'ink2') + sym('x', 11.9, 0.95, 0.4, 'soft')
    section = [full('ink')] + [s for k, x, y, z, f in lay for s in sym(k, x - 7.4, y, z, 'soft' if f == 'ink2' else f)]
    return _d(cover, content, section, CG(*LEFT(6.4), pic=(8.4, 1.5, 4.0, 3.8), pr=0.2), SG(*SG_R))


def s_flask(p):
    def flask(x, y, k):
        return [R(x + 1.15 * k, y, 0.7 * k, 1.5 * k, 'soft', l=['ink', 2]), Tri(x, y + 1.3 * k, 3.0 * k, 2.9 * k, 'soft', l=['ink', 2]),
                Tri(x + 0.42 * k, y + 2.35 * k, 2.16 * k, 1.85 * k, 'accent', a=0.9), R(x + 1.0 * k, y - 0.1 * k, 1.0 * k, 0.18 * k, 'ink', rr=0.05),
                circ(x + 1.5 * k, y + 2.9 * k, 0.22 * k, 'white', a=0.85), circ(x + 1.25 * k, y + 3.45 * k, 0.14 * k, 'white', a=0.8),
                circ(x + 1.8 * k, y + 3.5 * k, 0.12 * k, 'white', a=0.8), circ(x + 1.55 * k, y - 0.55 * k, 0.24 * k, 'accent', a=0.6),
                circ(x + 1.9 * k, y - 1.0 * k, 0.16 * k, 'accent', a=0.45)]
    cover = [full('bg')] + flask(8.9, 1.6, 1.2)
    content = [full('bg')] + flask(11.75, 0.42, 0.3)
    section = [full('ink')] + flask(2.0, 1.7, 1.1)
    return _d(cover, content, section, CG(*LEFT(6.6), pic=(8.6, 1.2, 3.8, 4.0), pr=0.2), SG(*SG_R))


def s_contour(p):
    def rings(cx, cy, k, n):
        return [Rg(cx - (0.5 + i * 0.55) * k * 1.2, cy - (0.5 + i * 0.55) * k, (1.0 + i * 1.1) * k * 1.2, (1.0 + i * 1.1) * k, 'accent' if i % 2 else 'ink2',
                   th=0.03, a=max(0.35, 0.95 - i * 0.08), r=-12) for i in range(n)]
    cover = [full('bg')] + rings(11.2, 3.6, 0.85, 6) + [circ(11.2, 3.6, 0.3, 'accent')]
    content = [full('bg')] + rings(12.3, 0.85, 0.32, 4)
    section = [full('ink')] + [dict(s, f='soft') for s in rings(3.2, 3.75, 0.9, 5)]
    return _d(cover, content, section, CG(*LEFT(6.2), pic=(8.9, 1.9, 3.4, 3.0), pk='circle'), SG(*SG_R))


def s_book(p):
    def book(x, y, k):
        page = lambda px, r: [R(px, y, 2.4 * k, 3.2 * k, 'white', rr=0.08 * k, r=r, l=['ink2', 1.25])]
        lines = [R(x + 0.3 * k, y + (0.5 + i * 0.32) * k, 1.8 * k, 0.05 * k + 0.005, 'soft', r=-4) for i in range(8)] + \
                [R(x + 2.75 * k, y + (0.5 + i * 0.32) * k, 1.8 * k, 0.05 * k + 0.005, 'soft', r=4) for i in range(8)]
        return [R(x - 0.2 * k, y + 0.25 * k, 5.25 * k, 3.25 * k, 'accent', rr=0.12 * k)] + page(x, -4) + page(x + 2.45 * k, 4) + lines + \
               [R(x + 2.4 * k, y + 0.05 * k, 0.06 * k, 3.25 * k, 'ink2'), R(x + 3.9 * k, y - 0.2 * k, 0.25 * k, 1.4 * k, 'accent2')]
    cover = [full('bg')] + book(7.9, 1.9, 1.0)
    content = [full('bg')] + book(11.45, 0.45, 0.27)
    section = [full('ink')] + book(1.2, 2.0, 1.0)
    return _d(cover, content, section, CG(*LEFT(6.3), pic=(8.0, 1.7, 4.8, 3.4), pr=0.1), SG(*SG_R))


def s_pitch(p):
    def field(x, y, w, h, f, t):
        return [R(x, y, w, h, 'ink2', l=[f, t]), R(x + w / 2 - 0.01, y, 0.02, h, f), Rg(x + w / 2 - h * 0.18, y + h / 2 - h * 0.18, h * 0.36, h * 0.36, f, th=0.025),
                R(x, y + h * 0.28, w * 0.13, h * 0.44, 'ink2', l=[f, t]), R(x + w * 0.87, y + h * 0.28, w * 0.13, h * 0.44, 'ink2', l=[f, t]), circ(x + w / 2, y + h / 2, 0.08, f)]
    stripes = [R(7.6 + i * 0.72, 0.9, 0.36, 5.7, 'ink', a=0.16) for i in range(8)]
    cover = [full('ink'), R(7.6, 0.9, 5.2, 5.7, 'ink2')] + stripes + field(7.8, 1.1, 4.8, 5.3, 'white', 1.5)
    content = [full('bg'), R(11.3, 0.35, 1.65, 1.05, 'ink2')] + field(11.38, 0.42, 1.5, 0.9, 'white', 0.75)[1:]
    section = [full('ink'), R(0.9, 0.9, 5.2, 5.7, 'ink2')] + field(1.1, 1.1, 4.8, 5.3, 'white', 1.5)
    return _d(cover, content, section, DARK(*LEFT(6.0), pic=(8.0, 1.5, 4.2, 4.0), pr=0.1), SG(*SG_R))


def s_parchment(p):
    def orn(x, y, s, f):
        return [Dm(x, y, s, s, f), circ(x + s / 2, y + s / 2, s * 0.35, 'bg'), circ(x + s / 2, y + s / 2, s * 0.16, f)]
    cover = [full('bg'), R(0.45, 0.45, W - 0.9, H - 0.9, 'bg', l=['accent', 1.25]), R(0.6, 0.6, W - 1.2, H - 1.2, 'bg', l=['ink', 0.5])] + \
        orn(0.25, 0.25, 0.6, 'accent') + orn(W - 0.85, 0.25, 0.6, 'accent') + orn(0.25, H - 0.85, 0.6, 'accent') + orn(W - 0.85, H - 0.85, 0.6, 'accent') + \
        [R(5.37, 5.55, 2.6, 0.04, 'accent'), Dm(6.52, 5.42, 0.3, 0.3, 'accent')]
    content = [full('bg')] + orn(0.12, 0.12, 0.34, 'accent') + orn(W - 0.46, 0.12, 0.34, 'accent') + [R(0.9, 1.52, 2.0, 0.03, 'accent')]
    section = [full('ink'), R(0.45, 0.45, W - 0.9, H - 0.9, 'ink', l=['accent', 1.25])] + orn(0.25, 0.25, 0.6, 'accent') + orn(W - 0.85, H - 0.85, 0.6, 'accent')
    return _d(cover, content, section, CG((1.4, 1.5, 10.5, 2.4), (1.4, 3.95, 10.5, 0.7), (1.4, 4.6, 10.5, 0.5), al='center', pic=(8.8, 1.3, 3.4, 4.0), pr=0.05),
              SG((1.4, 2.0, 10.5, 2.0), (1.4, 4.1, 10.5, 0.8), al='center'), geo=G(title=(0.9, 0.55, 11.5, 0.95)))


def s_equalizer(p):
    hs = [1.4, 2.6, 3.8, 2.2, 4.6, 3.1, 1.8, 3.6, 2.4, 1.2]
    bars = [R(7.9 + i * 0.5, 6.3 - h, 0.34, h, 'accent', rr=0.17, g=['accent', 'accent2', 90]) for i, h in enumerate(hs)]
    cover = [full('bg')] + bars + [R(7.7, 6.4, 5.2, 0.04, 'ink')]
    content = [full('bg')] + [R(11.55 + i * 0.24, 1.35 - h * 0.25, 0.15, h * 0.25, 'accent', rr=0.075) for i, h in enumerate(hs[:6])]
    section = [full('ink')] + [dict(b, x=b['x'] - 7.0) for b in bars]
    return _d(cover, content, section, CG(*LEFT(6.6), pic=(8.6, 0.9, 3.8, 2.4), pr=0.2), SG(*SG_R))


# ======================================================================== calm
def s_hairline(p):
    cover = [full('bg'), R(0.9, 4.55, 6.2, 0.015, 'ink'), circ(7.25, 4.555, 0.16, 'accent')]
    content = [full('bg'), R(0.9, 1.55, 1.2, 0.015, 'ink'), circ(2.2, 1.555, 0.08, 'accent')]
    section = [full('bg'), R(0.9, 4.45, 11.5, 0.015, 'ink'), circ(12.45, 4.455, 0.16, 'accent')]
    return _d(cover, content, section, CG((0.9, 1.6, 10.5, 2.7), (0.9, 4.75, 9.5, 0.7), (0.9, 5.45, 9.5, 0.5), pic=(9.4, 1.0, 3.0, 3.2), pr=0.0),
              SGL((0.9, 2.1, 11.5, 2.2), (0.9, 4.65, 11.5, 0.8)), geo=G(title=(0.9, 0.55, 11.5, 0.95)))


def s_clouds(p):
    def cloud(x, y, k, f, a=None):
        return [O(x, y + 0.5 * k, 2.6 * k, 1.1 * k, f, a=a), O(x + 0.5 * k, y, 1.4 * k, 1.3 * k, f, a=a), O(x + 1.2 * k, y + 0.2 * k, 1.2 * k, 1.1 * k, f, a=a)]
    cover = [full('bg'), R(0, 0, W, H, 'tint')] + cloud(8.2, 0.8, 1.6, 'white') + cloud(10.0, 3.1, 1.2, 'soft') + cloud(7.6, 4.8, 1.4, 'white', 0.95) + \
        [circ(12.2, 1.0, 1.1, 'accent', a=0.6)]
    content = [full('bg')] + cloud(11.4, 0.45, 0.55, 'soft') + [circ(12.8, 0.45, 0.4, 'accent', a=0.5)]
    section = [full('ink')] + cloud(1.0, 1.5, 1.9, 'ink2') + cloud(3.0, 4.2, 1.3, 'ink2', 0.8)
    return _d(cover, content, section, CG(*LEFT(6.3), sc='text', pic=(8.4, 1.2, 4.0, 3.6), pr=0.35), SG(*SG_R))


def s_papercut(p):
    layers = [('soft', 0), ('accent2', 0.55), ('accent', 1.05), ('ink2', 1.6), ('ink', 2.1)]
    def hills(x0, y0, w, k):
        out = []
        for i, (f, dy) in enumerate(layers):
            out.append(O(x0 - 1.0 + (i % 2) * 0.8, y0 + dy * k, w * 0.7, 3.2 * k, f, l=['white', 1.5]))
            out.append(O(x0 + w * 0.35 - (i % 2) * 0.6, y0 + dy * k + 0.3 * k, w * 0.75, 3.0 * k, f, l=['white', 1.5]))
        return out
    cover = [full('bg')] + hills(5.6, 4.4, 9.0, 1.0)
    content = [full('bg')] + [O(11.4 + i * 0.15, 0.3 + i * 0.22, 1.5 - i * 0.12, 0.6, f, l=['white', 0.75]) for i, (f, dy) in enumerate(layers[:4])]
    section = [full('ink')] + [dict(s, y=s['y'] - 0.3) for s in hills(-1.5, 4.6, 9.0, 1.0)]
    return _d(cover, content, section, CG(*LEFT(7.6, 0.5), pic=(9.0, 0.7, 3.4, 2.8), pr=0.3), SG((0.9, 0.8, 11.0, 2.0), (0.9, 2.9, 11.0, 0.8)))


def s_origami(p):
    def crane(x, y, k):
        return [Tri(x, y + 1.0 * k, 3.0 * k, 2.2 * k, 'accent', g=['accent', 'ink2', 30]), RT(x + 1.5 * k, y + 1.0 * k, 1.5 * k, 2.2 * k, 'ink2', a=0.9),
                Tri(x + 2.2 * k, y, 1.6 * k, 2.4 * k, 'soft', r=35, g=['soft', 'accent', 60]), Tri(x - 0.7 * k, y + 0.4 * k, 1.4 * k, 2.0 * k, 'accent2', r=-40, a=0.95),
                RT(x + 0.3 * k, y + 2.6 * k, 2.4 * k, 0.8 * k, 'ink', a=0.85)]
    cover = [full('bg')] + crane(8.6, 1.4, 1.25)
    content = [full('bg')] + crane(11.8, 0.25, 0.32)
    section = [full('ink')] + crane(1.8, 1.6, 1.2)
    return _d(cover, content, section, CG(*LEFT(6.6), pic=(8.6, 1.4, 3.8, 3.8), pr=0.2), SG(*SG_R))


def s_dunes(p):
    cover = [full('bg'), circ(11.4, 1.7, 1.5, 'accent', a=0.85), O(5.0, 4.6, 10.5, 4.2, 'soft', g=['soft', 'accent', 90]), O(-2.0, 5.3, 11.0, 4.0, 'accent', g=['accent', 'ink2', 90], a=0.95),
             O(6.8, 5.9, 9.0, 3.0, 'ink2', g=['ink2', 'ink', 90])]
    content = [full('bg'), O(11.0, 0.65, 2.6, 1.3, 'soft', g=['soft', 'accent', 90]), O(11.9, 0.95, 2.0, 1.0, 'accent', a=0.85), circ(12.85, 0.35, 0.35, 'accent', a=0.8)]
    section = [full('ink'), circ(2.0, 1.6, 1.3, 'accent', a=0.9), O(-2.5, 5.1, 9.5, 4.0, 'ink2', a=0.9), O(4.0, 5.6, 11.0, 3.6, 'accent', a=0.6)]
    return _d(cover, content, section, CG(*LEFT(7.4, 0.7), pic=(8.6, 0.9, 3.4, 2.8), pr=0.25), SG((3.6, 1.4, 9.0, 2.2), (3.6, 3.7, 9.0, 0.8)))


def s_cafe(p):
    def ring(cx, cy, d, a, th=0.07):
        return Rg(cx - d / 2, cy - d / 2, d, d, 'accent', th=th, a=a)
    cover = [full('bg'), ring(10.3, 2.6, 3.4, 0.35), ring(11.3, 4.4, 2.6, 0.22, 0.05), circ(9.6, 4.6, 2.6, 'ink'), circ(9.6, 4.6, 2.05, 'ink2'),
             O(10.85, 4.15, 0.9, 1.0, 'ink', l=['ink', 2]), O(9.35, 3.25, 0.9, 0.45, 'soft', a=0.35, r=-15)]
    content = [full('bg'), ring(12.3, 0.85, 1.1, 0.4, 0.05), ring(11.85, 1.1, 0.7, 0.25, 0.04)]
    section = [full('ink'), ring(3.3, 3.6, 4.2, 0.5, 0.09), ring(4.6, 5.0, 2.6, 0.3, 0.06)]
    return _d(cover, content, section, CG(*LEFT(6.2), pic=(8.6, 0.9, 3.6, 2.6), pr=0.2), SG(*SG_R))


def s_scales(p):
    def scale_rows(x0, y0, cols, rows, s, fs):
        out = []
        for r in range(rows):
            for c in range(cols):
                x = x0 + c * s + (s / 2 if r % 2 else 0)
                y = y0 + r * s * 0.5
                f = fs[(r + c) % len(fs)]
                out += [circ(x, y, s, f), Rg(x - s * 0.36, y - s * 0.36, s * 0.72, s * 0.72, 'white', th=0.03, a=0.8), Rg(x - s * 0.2, y - s * 0.2, s * 0.4, s * 0.4, 'white', th=0.03, a=0.8)]
        return out
    cover = [full('bg')] + scale_rows(7.6, 5.0, 8, 6, 0.85, ['accent', 'ink2', 'soft'])
    content = [full('bg')] + scale_rows(11.6, 0.55, 3, 2, 0.45, ['accent', 'soft'])
    section = [full('ink')] + scale_rows(-0.2, 5.0, 17, 6, 0.85, ['ink2', 'accent'])
    return _d(cover, content, section, CG(*LEFT(7.4, 0.6), pic=(9.0, 0.7, 3.4, 3.0), pr=0.25), SG((0.9, 1.2, 11.0, 2.2), (0.9, 3.5, 11.0, 0.8)))


def s_lattice(p):
    def lat(x0, y0, cols, rows, s, f, a):
        return [Dm(x0 + c * s * 0.5, y0 + r * s * 0.5 + (s * 0.25 if c % 2 else 0), s * 0.5, s * 0.5, 'soft', l=[f, 2.0], a=a) for r in range(rows) for c in range(cols)]
    cover = [full('bg')] + lat(8.0, 0.6, 10, 12, 1.0, 'accent', 0.95) + [Dm(10.0, 2.6, 1.0, 1.0, 'accent'), Dm(11.0, 4.1, 0.5, 0.5, 'ink2'), Dm(9.0, 4.85, 0.5, 0.5, 'ink2')]
    content = [full('bg')] + lat(11.5, 0.3, 6, 4, 0.6, 'accent', 0.85)
    section = [full('ink')] + [dict(s, f='ink', l=['soft', 1.25]) for s in lat(0.6, 0.6, 12, 12, 1.0, 'soft', 0.5)]
    return _d(cover, content, section, CG(*LEFT(6.6), pic=(8.8, 1.6, 3.6, 3.6), pk='circle'), SG(*SG_R))


def s_frosted(p):
    cover = [full('soft', g=['soft', 'bg', 45]), circ(10.4, 2.0, 3.4, 'accent', a=0.55), circ(12.0, 5.2, 2.6, 'accent2', a=0.45),
             R(8.4, 1.4, 4.0, 2.6, 'white', a=0.55, rr=0.3, l=['white', 1.5]), R(9.2, 3.5, 3.5, 2.4, 'white', a=0.6, rr=0.3, l=['white', 1.5])]
    content = [full('bg'), circ(12.3, 0.7, 1.0, 'accent', a=0.45), R(11.4, 0.4, 1.4, 0.95, 'white', a=0.6, rr=0.15, l=['soft', 1])]
    section = [full('ink'), circ(3.0, 2.6, 3.6, 'accent', a=0.45), circ(5.0, 5.0, 2.4, 'accent2', a=0.4), R(1.2, 2.0, 4.6, 3.2, 'white', a=0.12, rr=0.35, l=['white', 1])]
    return _d(cover, content, section, CG(*LEFT(6.2), tc=readable(p['ink'], mix(p['soft'], p['bg'], 0.3), 7.0), sc=readable(p['text'], p['soft'], 7.0),
                                          mc=readable(p['muted'], p['soft'], 7.0), pic=(8.6, 1.6, 3.6, 2.2), pr=0.3), SG(*SG_R))


def s_brush(p):
    cover = [full('bg'), R(7.4, 1.0, 6.6, 1.3, 'accent', rr=0.65, r=-8, a=0.9), R(8.2, 2.6, 5.6, 1.0, 'ink2', rr=0.5, r=-8, a=0.85),
             R(7.0, 4.1, 6.8, 1.5, 'soft', rr=0.75, r=-8), R(9.6, 5.4, 3.6, 0.6, 'accent', rr=0.3, r=-8, a=0.7), circ(8.4, 6.2, 0.3, 'accent', a=0.7)]
    content = [full('bg'), R(11.0, 0.45, 2.4, 0.42, 'accent', rr=0.21, r=-8, a=0.9), R(11.6, 0.95, 1.6, 0.3, 'soft', rr=0.15, r=-8)]
    section = [full('ink'), R(-1.0, 1.3, 7.6, 1.5, 'accent', rr=0.75, r=-8, a=0.85), R(-0.6, 3.2, 6.6, 1.1, 'ink2', rr=0.55, r=-8)]
    return _d(cover, content, section, CG(*LEFT(6.0), pic=(8.2, 1.9, 3.8, 3.0), pr=0.3), SG(*SG_R))


def s_contour_free(p):          # (not used: kept for the record)
    return s_contour(p)


# ======================================================================== bold
def s_woven(p):
    cs = ['accent', 'ink', 'accent2', 'soft', 'C2410C', 'accent', 'ink', 'accent2']
    def band(x0, y0, w, h, cols, rows):
        out = []
        cw, rh = w / cols, h / rows
        for r in range(rows):
            for c in range(cols):
                f = cs[(r * 3 + c) % len(cs)] if (r + c) % 2 == 0 else cs[(r + c * 2 + 1) % len(cs)]
                out.append(R(x0 + c * cw, y0 + r * rh, cw + 0.005, rh + 0.005, f))
        return out
    cover = [full('bg')] + band(8.4, 0, 4.933, H, 4, 12) + [R(8.4, 0, 0.08, H, 'ink')]
    content = [full('bg')] + band(0, 0, W, 0.16, 24, 1)
    section = [full('ink')] + band(0, 5.6, W, 1.9, 12, 3)
    return _d(cover, content, section, CG(*LEFT(6.6), pic=(1.0, 4.2, 2.4, 2.4), pr=0.2), SG((0.9, 1.3, 11.0, 2.2), (0.9, 3.6, 11.0, 0.8)),
              geo=G(title=(0.9, 0.55, 11.5, 0.95)))


def s_dusk(p):
    hills = [O(-1.5, 5.4, 9.0, 3.6, 'ink'), O(6.0, 5.0, 9.5, 4.0, 'deep'), O(3.0, 6.0, 8.0, 3.0, 'ink2')]
    cover = [full('ink'), R(0, 4.0, W, 3.5, 'accent', g=['accent', 'ink', 90]), circ(10.2, 5.6, 3.0, 'soft', a=0.95), circ(10.2, 5.6, 2.2, 'accent2', a=0.6)] + hills
    content = [full('bg'), circ(12.3, 1.1, 0.9, 'accent'), R(11.4, 1.1, 1.9, 0.45, 'bg'), R(11.45, 1.1, 1.8, 0.04, 'ink')]
    section = [full('ink'), R(0, 4.0, W, 3.5, 'accent2', g=['accent2', 'ink', 90]), circ(3.2, 6.0, 3.2, 'accent', a=0.9)] + hills
    return _d(cover, content, section, DARK((0.9, 0.7, 7.2, 2.4), (0.9, 3.15, 7.2, 0.7), (0.9, 3.85, 7.2, 0.5), pic=(8.6, 0.8, 3.8, 2.6), pr=0.2),
              SG((5.6, 0.9, 7.0, 2.2), (5.6, 3.2, 7.0, 0.8)))


def s_swiss(p):
    grid = [R(0.9 + i * 2.31, 0, 0.012, H, 'soft', a=0.9) for i in range(6)]
    cover = [full('bg')] + grid + [R(8.2, 0, 5.133, 4.3, 'accent'), R(0, 6.9, 8.2, 0.6, 'ink'), circ(10.7, 2.15, 1.6, 'bg')]
    content = [full('bg'), R(12.4, 0, 0.933, 0.933, 'accent'), R(0.9, 1.52, 0.9, 0.06, 'accent')]
    section = [full('ink'), R(0, 0, 5.0, 4.3, 'accent'), circ(2.5, 2.15, 1.5, 'ink')]
    return _d(cover, content, section, CG((0.9, 1.1, 7.0, 2.8), (0.9, 4.45, 7.0, 0.7), (0.9, 5.15, 7.0, 0.5), pic=(8.6, 4.6, 3.6, 2.6), pr=0.0),
              SG((5.6, 1.8, 7.0, 2.2), (5.6, 4.1, 7.0, 0.8)), geo=G(title=(0.9, 0.55, 11.0, 0.95)))


def s_construct(p):
    cover = [full('bg'), R(6.8, 2.6, 8.8, 0.9, 'accent', r=-30), R(8.4, 4.2, 6.6, 0.4, 'ink', r=-30), circ(11.4, 1.6, 2.0, 'ink'), circ(11.4, 1.6, 0.8, 'accent2'),
             R(9.0, 5.9, 3.6, 0.14, 'accent', r=-30)]
    content = [full('bg'), R(11.0, 0.55, 2.6, 0.3, 'accent', r=-30), circ(12.75, 1.1, 0.42, 'ink')]
    section = [full('ink'), R(-1.4, 2.8, 9.0, 0.9, 'accent', r=-30), R(-0.6, 4.4, 7.0, 0.4, 'soft', r=-30), circ(1.9, 1.5, 1.6, 'accent2')]
    return _d(cover, content, section, CG(*LEFT(5.8), pic=(10.2, 3.6, 2.6, 2.6), pk='circle'), SG(*SG_R))


def s_confetti(p):
    import random
    rnd = random.Random(61)
    cols = ['accent', 'accent2', 'ink2', 'soft', 'ink']
    def bits(x0, x1, y0, y1, n, k=1.0):
        out = []
        for i in range(n):
            x, y = rnd.uniform(x0, x1), rnd.uniform(y0, y1)
            f = cols[i % len(cols)]
            if i % 3 == 0:
                out.append(circ(x, y, 0.22 * k, f))
            else:
                out.append(R(x, y, 0.38 * k, 0.13 * k, f, r=rnd.uniform(-70, 70), rr=0.04 * k))
        return out
    streamers = [R(8.4, 0.9 + i * 1.6, 4.6, 0.09, cols[i], r=-14 + i * 9, rr=0.045, a=0.85) for i in range(4)]
    cover = [full('bg')] + streamers + bits(7.8, 13.0, 0.4, 7.0, 46)
    content = [full('bg')] + bits(11.3, 13.0, 0.25, 1.35, 9, 0.7)
    section = [full('ink')] + bits(0.4, 13.0, 0.3, 1.4, 30) + bits(0.4, 13.0, 5.9, 7.2, 30)
    return _d(cover, content, section, CG(*LEFT(6.3), pic=(8.6, 1.8, 3.8, 3.4), pr=0.25), SG((0.9, 2.0, 11.0, 2.2), (0.9, 4.3, 11.0, 0.8)))


def s_skyline(p):
    import random
    rnd = random.Random(7)
    def city(x0, x1, base, k, fb, fw):
        out, x = [], x0
        while x < x1:
            w = rnd.uniform(0.55, 1.05) * k
            h = rnd.uniform(1.4, 4.3) * k
            out.append(R(x, base - h, w, h, fb))
            for wy in range(int(h / (0.42 * k))):
                for wx in range(int(w / (0.3 * k))):
                    if rnd.random() < 0.45:
                        out.append(R(x + 0.1 * k + wx * 0.3 * k, base - h + 0.2 * k + wy * 0.42 * k, 0.12 * k, 0.16 * k, fw, a=0.95))
            x += w + 0.05 * k
        return out
    cover = [full('bg'), circ(11.6, 1.4, 1.2, 'accent', a=0.5)] + city(7.4, W, 6.9, 1.0, 'ink', 'accent') + [R(7.2, 6.9, 6.2, 0.6, 'ink2')]
    content = [full('bg')] + city(11.3, 13.0, 1.4, 0.26, 'ink2', 'accent')
    section = [full('ink'), circ(10.8, 1.5, 1.4, 'accent', a=0.85)] + city(0, W, 7.5, 1.0, 'ink2', 'accent')
    return _d(cover, content, section, CG(*LEFT(6.0), pic=(8.2, 0.7, 3.6, 2.1), pr=0.2), SG((0.9, 0.9, 9.0, 2.2), (0.9, 3.2, 9.0, 0.8)))


def s_bunting(p):
    cols = ['accent', 'ink2', 'accent2', 'soft']
    def flags(x0, x1, y0, sag, n, s):
        out = []
        for i in range(n):
            t = i / max(1, n - 1)
            x = x0 + t * (x1 - x0)
            y = y0 + sag * 4 * t * (1 - t)
            out.append(Tri(x - s / 2, y, s, s * 1.1, cols[i % len(cols)], r=180))
        out.insert(0, R(x0, y0 + sag * 0.55, x1 - x0, 0.03, 'ink', a=0.7))
        return out
    cover = [full('bg')] + flags(7.4, 13.1, 0.4, 0.9, 9, 0.6) + flags(8.6, 13.4, 2.1, 0.6, 7, 0.5)
    content = [full('bg')] + flags(11.2, 13.1, 0.25, 0.35, 5, 0.3)
    section = [full('ink')] + flags(0.2, 13.1, 0.4, 1.2, 15, 0.6)
    return _d(cover, content, section, CG(*LEFT(6.4, 1.6), pic=(8.6, 3.4, 3.8, 3.2), pr=0.25), SG((0.9, 2.6, 11.0, 2.2), (0.9, 4.9, 11.0, 0.8)))


def s_prism(p):
    spec = ['EF4444', 'F97316', 'FACC15', '22C55E', '3B82F6', '8B5CF6']
    def prism(cx, cy, k):
        out = [R(cx - 3.5 * k, cy + 0.2 * k, 2.1 * k, 0.06 * k + 0.01, 'soft', r=-8)]
        for i, c in enumerate(spec):
            out.append(R(cx + 0.35 * k, cy + (0.05 + i * 0.13) * k, 3.6 * k, 0.12 * k, c, r=8 + i * 4))
        out.append(Tri(cx - 1.4 * k, cy - 1.6 * k, 2.8 * k, 2.6 * k, 'soft', a=0.92, l=['white', 1.5]))
        return out
    cover = [full('ink')] + prism(10.6, 3.3, 1.0)
    content = [full('bg'), R(11.0, 0.95, 0.9, 0.03, 'ink2', r=-8)] + [R(12.25, 0.75 + i * 0.05, 0.75, 0.045, c, r=8 + i * 4) for i, c in enumerate(spec)] + \
        [Tri(11.75, 0.38, 0.75, 0.7, 'ink2')]
    section = [full('ink')] + prism(4.6, 3.6, 1.0)
    return _d(cover, content, section, DARK(*LEFT(6.0), pic=(8.6, 0.6, 3.6, 2.2), pr=0.2), SG((6.4, 0.7, 6.2, 1.8), (6.4, 2.55, 6.2, 0.7)))


def s_bullseye(p):
    def target(cx, cy, d, k):
        out = [circ(cx, cy, d * (1 - i * 0.2), 'accent' if i % 2 == 0 else 'white') for i in range(5)]
        out += [R(cx - 0.02, cy - 0.06, 3.4 * k, 0.08 * k + 0.02, 'ink', r=-25), Tri(cx + 3.0 * k, cy - 1.6 * k, 0.6 * k, 0.6 * k, 'ink', r=-25)]
        return out
    cover = [full('bg')] + target(10.6, 3.8, 4.6, 0.9)
    content = [full('bg')] + target(12.4, 0.85, 1.0, 0.2)
    section = [full('ink')] + target(3.4, 3.75, 4.8, 0.9)
    return _d(cover, content, section, CG(*LEFT(6.3), pic=(8.6, 1.2, 3.8, 3.8), pk='circle'), SG(*SG_R))


def s_speech(p):
    def bubble(x, y, w, h, f, tail=1, a=None):
        return [R(x, y, w, h, f, rr=min(w, h) * 0.3, a=a), Tri(x + (w * 0.2 if tail > 0 else w * 0.65), y + h - 0.05, 0.5, 0.45, f, r=180 + 20 * tail, a=a)]
    cover = [full('bg')] + bubble(8.0, 0.9, 3.8, 1.7, 'accent') + bubble(9.4, 3.0, 3.4, 1.6, 'ink2', -1) + bubble(7.8, 5.0, 2.8, 1.2, 'soft') + \
        [circ(8.9 + i * 0.5, 1.75, 0.22, 'white') for i in range(3)]
    content = [full('bg')] + bubble(11.3, 0.35, 1.4, 0.65, 'accent') + bubble(12.0, 0.95, 1.0, 0.45, 'soft', -1)
    section = [full('ink')] + bubble(1.0, 1.4, 4.4, 2.0, 'accent') + bubble(2.4, 4.0, 3.6, 1.7, 'ink2', -1)
    return _d(cover, content, section, CG(*LEFT(6.2), pic=(8.6, 0.8, 3.8, 3.2), pr=0.3), SG(*SG_R))


def s_pixels(p):
    shape = ["..XX..", ".XXXX.", "XXXXXX", "XX..XX", "XXXXXX", ".X..X."]
    cols = ['accent', 'ink2', 'accent2', 'soft']
    def px(x0, y0, s):
        out = []
        for r, row in enumerate(shape):
            for c, ch in enumerate(row):
                if ch == 'X':
                    out.append(R(x0 + c * s, y0 + r * s, s * 0.9, s * 0.9, cols[(r + c) % len(cols)], rr=s * 0.08))
        return out
    trail = [R(7.6 + i * 0.4, 6.4 - (i % 3) * 0.4, 0.32, 0.32, cols[i % 4], a=0.4 + (i % 4) * 0.15) for i in range(8)]
    cover = [full('bg')] + px(9.0, 1.4, 0.6) + trail
    content = [full('bg')] + px(11.75, 0.3, 0.18)
    section = [full('ink')] + px(1.4, 1.4, 0.65)
    return _d(cover, content, section, CG(*LEFT(6.6), pic=(8.6, 1.2, 3.8, 3.8), pr=0.2), SG(*SG_R))


def s_retro(p):
    arcs = ['accent', 'accent2', 'soft', 'ink2']
    def rainbow(cx, cy, d, th):
        return [Rg(cx - (d - i * th * 2) / 2, cy - (d - i * th * 2) / 2, d - i * th * 2, d - i * th * 2, arcs[i % 4], th=th) for i in range(5)]
    cover = [full('bg')] + rainbow(W, H, 9.6, 0.55)
    content = [full('bg')] + rainbow(W, 0, 3.0, 0.16)
    section = [full('ink')] + rainbow(0, H, 9.0, 0.5)
    return _d(cover, content, section, CG(*LEFT(6.4, 0.9), pic=(9.4, 0.7, 3.0, 2.6), pr=0.3), SG((5.4, 1.4, 7.2, 2.2), (5.4, 3.7, 7.2, 0.8)),
              geo=G(title=(0.9, 0.55, 10.4, 0.95)))


def s_liquid(p):
    cover = [full('bg'), O(7.4, -1.2, 6.8, 5.6, 'accent', g=['accent', 'accent2', 30], a=0.85), O(9.6, 3.2, 4.6, 4.6, 'accent2', g=['accent2', 'soft', 60], a=0.75),
             O(7.0, 4.4, 3.4, 3.2, 'ink2', a=0.6)]
    content = [full('bg'), O(11.4, -0.6, 2.4, 1.9, 'accent', g=['accent', 'accent2', 30], a=0.8), O(12.4, 0.7, 0.9, 0.9, 'accent2', a=0.6)]
    section = [full('ink'), O(-1.4, -1.0, 7.4, 6.2, 'accent', g=['accent', 'accent2', 30], a=0.7), O(2.6, 3.6, 4.6, 4.4, 'accent2', a=0.55)]
    return _d(cover, content, section, CG(*LEFT(6.0), pic=(8.8, 1.4, 3.4, 3.4), pk='circle'), SG(*SG_R))


# ======================================================================== dark and dramatic (for the dramatic transitions)
def s_starry(p):
    import random
    rnd = random.Random(3)
    def sky(x0, x1, y0, y1, n):
        out = []
        for i in range(n):
            x, y = rnd.uniform(x0, x1), rnd.uniform(y0, y1)
            if i % 7 == 0:
                out.append(St(x, y, 0.28, 0.28, 'accent', a=0.95))
            else:
                out.append(circ(x, y, rnd.choice([0.05, 0.07, 0.1]), 'white', a=rnd.choice([0.5, 0.7, 0.9])))
        return out
    cover = [full('ink', g=['ink2', 'ink', 90])] + sky(7.6, W, 0, H, 44) + sky(0.2, 7.4, 5.5, H, 12) + [circ(10.9, 1.9, 1.9, 'soft', a=0.95), circ(11.45, 1.62, 1.7, 'ink', a=0.94)]
    content = [full('bg')] + [St(12.2, 0.35, 0.32, 0.32, 'accent'), St(11.7, 0.95, 0.2, 0.2, 'ink2'), circ(12.75, 1.15, 0.1, 'ink2'), circ(11.5, 0.45, 0.08, 'accent')]
    section = [full('ink')] + sky(0, W, 0, 1.75, 26) + sky(0, W, 5.4, H, 26)
    return _d(cover, content, section, DARK(*LEFT(6.8), pic=(8.6, 3.4, 3.4, 3.4), pk='circle'), SG((0.9, 2.0, 11.0, 2.2), (0.9, 4.3, 11.0, 0.8)))


def s_aurora(p):
    cover = [full('ink'), Pa(5.0, 0.4, 9.0, 1.4, 'accent', g=['accent', 'accent2', 0], a=0.6, r=-10), Pa(6.5, 1.7, 8.0, 1.0, 'accent2', g=['accent2', 'soft', 0], a=0.45, r=-14),
             Pa(7.8, 2.5, 5.8, 0.6, 'soft', a=0.3, r=-8), O(-2, 6.2, 18, 3.0, 'deep'), Tri(8.2, 5.0, 1.4, 2.0, 'deep'), Tri(9.3, 4.6, 1.8, 2.4, 'deep'), Tri(11.0, 5.2, 1.3, 1.8, 'deep')]
    content = [full('bg'), Pa(11.0, 0.3, 2.4, 0.4, 'accent', g=['accent', 'accent2', 0], a=0.8, r=-10), Pa(11.4, 0.85, 1.9, 0.25, 'accent2', a=0.6, r=-14)]
    section = [full('ink'), Pa(-1.0, 0.6, 10.0, 1.3, 'accent', g=['accent', 'accent2', 0], a=0.5, r=-8), Pa(0.4, 2.0, 8.0, 0.8, 'accent2', a=0.35, r=-12)]
    return _d(cover, content, section, DARK((0.9, 3.0, 7.0, 2.2), (0.9, 5.25, 7.0, 0.6), (0.9, 5.85, 7.0, 0.45), pic=(9.0, 3.2, 3.0, 2.6), pr=0.2),
              SG((0.9, 3.2, 11.0, 2.0), (0.9, 5.2, 11.0, 0.8)))


def s_neon(p):
    cover = [full('ink'), Rg(8.6, 0.9, 3.6, 3.6, 'accent', th=0.07), Rg(8.45, 0.75, 3.9, 3.9, 'accent', th=0.12, a=0.25),
             R(9.6, 4.6, 3.0, 1.8, 'ink', rr=0.3, l=['soft', 3]), R(9.5, 4.5, 3.2, 2.0, 'ink', rr=0.35, l=['soft', 1], a=0.5),
             Tri(7.8, 4.4, 1.4, 1.3, 'ink', l=['accent2', 2.5])]
    content = [full('bg'), Rg(11.95, 0.35, 0.95, 0.95, 'accent', th=0.05), R(11.3, 0.8, 0.9, 0.5, 'bg', rr=0.1, l=['soft', 1.5])]
    section = [full('ink'), Rg(1.2, 1.6, 4.2, 4.2, 'accent', th=0.08), Rg(1.0, 1.4, 4.6, 4.6, 'accent', th=0.14, a=0.22), R(3.8, 4.2, 2.6, 1.6, 'ink', rr=0.3, l=['soft', 3])]
    return _d(cover, content, section, DARK(*LEFT(6.6), pic=(8.95, 1.25, 2.9, 2.9), pk='circle'), SG(*SG_R))


def s_deco(p):
    def fan(cx, cy, d, n):
        return [Rg(cx - (d - i * d / n) / 2, cy - (d - i * d / n) / 2, d - i * d / n, d - i * d / n, 'accent', th=0.04, a=0.95 - i * 0.07) for i in range(n)] + \
               [R(cx - 0.02, cy - d / 2, 0.04, d / 2, 'accent', r=a, a=0.7) for a in (-60, -30, 0, 30, 60)]
    cover = [full('ink')] + fan(W / 2, H - 0.6, 5.6, 7) + [R(0.6, 0.6, W - 1.2, H - 1.2, 'ink', l=['accent', 1.25]), R(0.8, 0.8, W - 1.6, H - 1.6, 'ink', l=['accent', 0.5])]
    content = [full('bg')] + fan(12.45, 1.4, 1.6, 4)
    section = [full('ink')] + fan(W / 2, H - 0.6, 5.0, 6) + [R(0.6, 0.6, W - 1.2, H - 1.2, 'ink', l=['accent', 1.25])]
    return _d(cover, content, section, DARK((1.4, 0.9, 10.5, 2.2), (1.4, 3.15, 10.5, 0.6), (1.4, 3.75, 10.5, 0.45), al='center', pic=(5.27, 2.1, 2.8, 1.0), pr=0.1),
              SG((1.4, 1.0, 10.5, 2.0), (1.4, 3.1, 10.5, 0.8), al='center'))


def s_moon(p):
    fog = [O(-2.0, 5.2, 9.5, 1.6, 'white', a=0.07), O(4.0, 5.7, 11.0, 1.8, 'white', a=0.06), O(1.0, 6.4, 13.0, 1.6, 'white', a=0.08)]
    branch = [R(9.0, 4.5, 4.6, 0.12, 'deep', r=-28), R(10.8, 3.7, 2.2, 0.07, 'deep', r=-58), R(11.6, 4.3, 1.6, 0.06, 'deep', r=-8),
              R(12.4, 3.55, 1.2, 0.05, 'deep', r=-40)]
    cover = [full('ink', g=['deep', 'ink', 90]), circ(10.6, 2.2, 2.8, 'soft', a=0.18), circ(10.6, 2.2, 2.2, 'soft'), circ(11.25, 1.85, 2.0, 'ink')] + fog + branch
    content = [full('bg'), circ(12.35, 0.8, 0.85, 'ink2'), circ(12.6, 0.65, 0.72, 'bg')]
    section = [full('ink'), circ(2.6, 2.2, 2.4, 'soft', a=0.2), circ(2.6, 2.2, 1.8, 'soft'), circ(3.15, 1.9, 1.6, 'ink')] + fog
    return _d(cover, content, section, DARK(*LEFT(6.6, 1.6), pic=(8.6, 4.4, 3.2, 2.4), pr=0.2), SG(*SG_R))


def s_thunder(p):
    def bolt(x, y, k, f):
        return [Pa(x, y, 1.0 * k, 1.9 * k, f, r=18), R(x - 0.75 * k, y + 1.62 * k, 1.75 * k, 0.42 * k, f, r=-14),
                Pa(x - 0.75 * k, y + 1.85 * k, 0.9 * k, 1.9 * k, f, r=22), Tri(x - 0.95 * k, y + 3.55 * k, 0.75 * k, 0.9 * k, f, r=200)]
    clouds = [O(7.2, -0.9, 4.6, 2.4, 'deep'), O(9.4, -1.1, 4.8, 2.6, 'ink2'), O(11.2, -0.6, 3.4, 2.0, 'deep'), O(8.4, 0.2, 3.6, 1.6, 'ink2', a=0.9)]
    cover = [full('ink')] + clouds + bolt(10.3, 1.4, 1.3, 'accent') + [R(0, 0, W, H, 'soft', a=0.04)]
    content = [full('bg')] + bolt(12.2, 0.3, 0.32, 'accent') + [O(11.2, 0.1, 1.3, 0.55, 'ink2', a=0.8)]
    section = [full('ink')] + [dict(c, x=c['x'] - 6.8) for c in clouds] + bolt(3.4, 1.5, 1.3, 'accent')
    return _d(cover, content, section, DARK(*LEFT(6.6, 1.8), pic=(8.6, 4.4, 3.6, 2.6), pr=0.15), SG(*SG_R))


def s_ember(p):
    import random
    rnd = random.Random(11)
    def embers(x0, x1, y0, y1, n, k=1.0):
        return [circ(rnd.uniform(x0, x1), rnd.uniform(y0, y1), rnd.choice([0.06, 0.09, 0.13, 0.18]) * k, rnd.choice(['accent', 'accent2', 'soft']),
                     a=rnd.choice([0.55, 0.75, 0.95])) for _ in range(n)]
    cover = [full('ink'), R(0, 5.0, W, 2.5, 'accent2', g=['accent2', 'ink', 90], a=0.55), O(5.0, 6.0, 9.0, 3.0, 'accent', a=0.35)] + embers(7.6, 13.2, 0.4, 6.8, 40) + embers(0.3, 7.4, 5.3, 7.2, 10)
    content = [full('bg')] + embers(11.3, 13.0, 0.25, 1.35, 10, 0.8)
    section = [full('ink'), R(0, 4.6, W, 2.9, 'accent2', g=['accent2', 'ink', 90], a=0.5)] + embers(0.2, 13.2, 4.4, 7.2, 46)
    return _d(cover, content, section, DARK(*LEFT(6.0, 0.8), pic=(8.6, 0.9, 3.6, 2.6), pr=0.2), SG((0.9, 0.9, 11.0, 2.0), (0.9, 2.95, 11.0, 0.8)))


# ======================================================================== the fifty
# (key, name, palette, design, title font, gallery group)
NEW61 = [
    # subjects
    ('blueprint', 'Blueprint', P('0B3D6B', '38BDF8', 'BAE6FD'), s_blueprint, 'light', 'Subjects'),
    ('chalk', 'Chalkboard', P('1F3B2D', 'F5D76E', 'E8F0EA', bg='F7FAF6', ink2='2F5D46'), s_chalk, 'serif', 'Subjects'),
    ('graph', 'Graph Paper', P('1E3A8A', 'EF4444', 'DBEAFE'), s_graph, 'light', 'Subjects'),
    ('news', 'Newsprint', P('111111', 'B91C1C', 'E7E5E4', bg='FAFAF7'), s_news, 'serif', 'Subjects'),
    ('shelf', 'Bookshelf', P('3F2A1D', 'C2410C', 'FED7AA', bg='FFFBF5', accent2='0F766E'), s_shelf, 'serif', 'Subjects'),
    ('grad', 'Graduation', P('0B1F44', 'D4A017', 'FEF3C7', ink2='22407A'), s_grad, 'serif', 'Subjects'),
    ('helix', 'Helix', P('134E4A', '2DD4BF', 'CCFBF1'), s_helix, 'light', 'Subjects'),
    ('circuit', 'Circuit Board', P('0F172A', '22C55E', 'BBF7D0'), s_circuit, 'grotesk', 'Subjects'),
    ('atom', 'Atom', P('1E1B4B', '22D3EE', 'CFFAFE'), s_atom, 'light', 'Subjects'),
    ('clinic', 'Clinic', P('0F4C5C', '14B8A6', 'CCFBF1', bg='FBFEFE'), s_clinic, 'light', 'Subjects'),
    ('pulse', 'Heartbeat', P('0B2545', 'F43F5E', 'FFE4E6'), s_pulse, 'light', 'Subjects'),
    ('cell', 'Cell Study', P('581C87', 'E879F9', 'FAE8FF'), s_cell, 'light', 'Subjects'),
    ('maths', 'Maths', P('1E3A8A', 'F59E0B', 'DBEAFE', ink2='3B5BDB'), s_maths, 'grotesk', 'Subjects'),
    ('flask', 'Lab Flask', P('0C4A6E', '06B6D4', 'CFFAFE'), s_flask, 'light', 'Subjects'),
    ('contour', 'Topography', P('14532D', '65A30D', 'ECFCCB', bg='FBFDF7'), s_contour, 'light', 'Subjects'),
    ('book', 'Open Book', P('1E3A5F', 'B45309', 'FDE68A', accent2='DC2626'), s_book, 'serif', 'Subjects'),
    ('pitch', 'Sports Field', P('14532D', '84CC16', 'D9F99D'), s_pitch, 'grotesk', 'Subjects'),
    ('parchment', 'Parchment', P('4A2C17', '9A3412', 'F5E6C8', bg='FBF5E9'), s_parchment, 'serif', 'Subjects'),
    ('equalizer', 'Sound Bars', P('1F2937', '6366F1', 'E0E7FF', accent2='EC4899'), s_equalizer, 'grotesk', 'Subjects'),
    # calm
    ('hairline', 'Hairline', P('111827', '2563EB', 'E5E7EB'), s_hairline, 'light', 'Calm'),
    ('clouds', 'Soft Clouds', P('1E3A5F', '93C5FD', 'E0F2FE', bg='F8FBFF'), s_clouds, 'light', 'Calm'),
    ('papercut', 'Paper Cut', P('064E3B', '10B981', 'D1FAE5', accent2='6EE7B7'), s_papercut, 'light', 'Calm'),
    ('origami', 'Origami', P('312E81', '818CF8', 'E0E7FF', accent2='F472B6'), s_origami, 'light', 'Calm'),
    ('dunes', 'Desert Dunes', P('7C2D12', 'F59E0B', 'FDE68A', bg='FFFBEB'), s_dunes, 'serif', 'Calm'),
    ('cafe', 'Café', P('3E2723', 'A16207', 'F5E6D3', bg='FFFCF7'), s_cafe, 'serif', 'Calm'),
    ('scales', 'Wave Scales', P('0C4A6E', '0EA5E9', 'E0F2FE'), s_scales, 'light', 'Calm'),
    ('lattice', 'Lattice', P('3F3F46', 'A855F7', 'F3E8FF'), s_lattice, 'light', 'Calm'),
    ('frosted', 'Frosted Glass', P('1E293B', '3B82F6', 'DBEAFE', accent2='F472B6'), s_frosted, 'light', 'Calm'),
    ('brush', 'Brush Stroke', P('0F172A', 'F43F5E', 'FFE4E6'), s_brush, 'serif', 'Calm'),
    ('liquid', 'Liquid Colour', P('3B0764', '8B5CF6', 'EDE9FE', accent2='06B6D4'), s_liquid, 'light', 'Calm'),
    # bold
    ('woven', 'Woven Bands', P('3B2A12', 'E0A100', 'FDE68A', accent2='15803D'), s_woven, 'grotesk', 'Bold'),
    ('dusk', 'Dusk Horizon', P('3B0A45', 'F97316', 'FED7AA', accent2='DB2777', bg='FFF8F1'), s_dusk, 'light', 'Bold'),
    ('swiss', 'Swiss Grid', P('111111', 'E11D48', 'E5E7EB'), s_swiss, 'grotesk', 'Bold'),
    ('construct', 'Constructivist', P('1C1917', 'DC2626', 'FEE2E2', accent2='F5F5F4'), s_construct, 'grotesk', 'Bold'),
    ('confetti', 'Confetti', P('1E3A8A', 'F59E0B', 'FDE68A', accent2='EC4899', ink2='6366F1'), s_confetti, 'grotesk', 'Bold'),
    ('skyline', 'City Skyline', P('1E293B', 'F59E0B', 'CBD5E1', bg='F8FAFC'), s_skyline, 'grotesk', 'Bold'),
    ('bunting', 'Bunting', P('7C2D12', 'F97316', 'FED7AA', accent2='0EA5E9'), s_bunting, 'grotesk', 'Bold'),
    ('prism', 'Prism', P('0F172A', 'F59E0B', 'E2E8F0', ink2='334155'), s_prism, 'light', 'Bold'),
    ('bullseye', 'Bullseye', P('7F1D1D', 'DC2626', 'FEE2E2'), s_bullseye, 'grotesk', 'Bold'),
    ('speech', 'Discussion', P('1E40AF', '60A5FA', 'DBEAFE', accent2='F59E0B'), s_speech, 'light', 'Bold'),
    ('pixels', 'Pixels', P('1E1B4B', '22C55E', 'DCFCE7', accent2='F59E0B'), s_pixels, 'grotesk', 'Bold'),
    ('retro', 'Retro Rainbow', P('431407', 'EA580C', 'FED7AA', accent2='F59E0B', bg='FFF7ED'), s_retro, 'serif', 'Bold'),
    # dark and dramatic
    ('starry', 'Starry Night', P('0B1026', 'FACC15', 'E0E7FF', ink2='1E2A5E'), s_starry, 'serif', 'Dark and dramatic'),
    ('aurora', 'Aurora', P('052E2B', '34D399', 'A7F3D0', accent2='8B5CF6'), s_aurora, 'light', 'Dark and dramatic'),
    ('neon', 'Neon Night', P('0A0A1F', 'F0ABFC', 'A5F3FC', accent2='FDE047', emph='86198F'), s_neon, 'grotesk', 'Dark and dramatic'),
    ('deco', 'Art Deco', P('111111', 'D4AF37', 'F3E5AB', ink2='2A2A2A'), s_deco, 'serif', 'Dark and dramatic'),
    ('moon', 'Midnight Moon', P('0B0F2A', 'C7D2FE', 'E5E7EB', emph='3F4A9E'), s_moon, 'serif', 'Dark and dramatic'),
    ('thunder', 'Thunderstorm', P('1F2937', 'FACC15', 'E5E7EB', ink2='374151'), s_thunder, 'grotesk', 'Dark and dramatic'),
    ('ember', 'Ember', P('1C0A00', 'F97316', 'FDBA74', accent2='DC2626'), s_ember, 'serif', 'Dark and dramatic'),
    ('glass', 'Stained Glass', P('1E293B', '0EA5E9', 'E0F2FE', accent2='F59E0B'), None, 'serif', 'Dark and dramatic'),
]


def s_glass(p):
    jewels = ['accent', 'accent2', 'B91C1C', '15803D', '7C3AED', 'soft']
    def window(x0, y0, cols, rows, s):
        out = []
        for r in range(rows):
            for c in range(cols):
                k = Dm if (r + c) % 2 == 0 else Hx
                out.append(k(x0 + c * s * 0.82, y0 + r * s * 0.72 + (s * 0.36 if c % 2 else 0), s, s * 0.86, jewels[(r * 2 + c) % len(jewels)], l=['deep', 2.25], a=0.95))
        return out
    cover = [full('ink'), R(8.0, 0, 5.333, H, 'deep')] + window(8.1, 0.2, 6, 10, 0.95) + [R(8.0, 0, 0.12, H, 'accent2')]
    content = [full('bg')] + window(11.5, 0.25, 3, 2, 0.5)
    section = [full('ink'), R(0, 0, 5.4, H, 'deep')] + window(0.1, 0.2, 6, 10, 0.95) + [R(5.4, 0, 0.12, H, 'accent2')]
    return _d(cover, content, section, DARK(*LEFT(6.4), pic=(1.0, 4.6, 2.4, 2.4), pr=0.1), SG(*SG_R))


NEW61[-1] = NEW61[-1][:3] + (s_glass,) + NEW61[-1][4:]
GROUPS61 = ['Subjects', 'Calm', 'Bold', 'Dark and dramatic']
TF = {'light': 'Calibri Light', 'serif': 'Georgia', 'grotesk': 'Trebuchet MS'}


def scenes61():
    """{scene key: {decor, geo, cg, sg}} and [(template id, group, name, palette, title font, scene key, gallery group)]."""
    S, L = {}, []
    for key, name, p, motif, tf, gg in NEW61:
        m = motif(p)
        sk = 'scene-' + key
        S[sk] = {'decor': {k: resolve(v, p) for k, v in m['decor'].items()}, 'geo': _colours(m['geo'], p), 'cg': _colours(m['cg'], p), 'sg': _colours(m['sg'], p)}
        for group in ('student', 'tutor'):
            L.append(('%s-%s' % (group, sk), group, name, p, tf, sk, gg))
    return S, L


def build():
    """Every new design as a template, in the server's form (hex without #)."""
    S, L = scenes61()
    T = {}
    for id_, group, name, p, tf, sk, gg in L:
        t = _base(group, name, p, tf, sk)
        t.update(S[sk])
        t['gallery'] = gg
        T[id_] = t
    return T
