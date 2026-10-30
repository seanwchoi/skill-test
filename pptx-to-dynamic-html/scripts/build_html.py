#!/usr/bin/env python3
"""Build one self-contained, scroll-style HTML page from slides.json.

Usage:
    python build_html.py WORK_DIR OUT.html [--title "Deck title"] [--include-hidden]

WORK_DIR is the folder produced by extract_pptx.py (+ render_reference.py).
Images are embedded as data URIs so OUT.html can be shared as a single file.

Behaviour baked into the page (see SKILL.md for the reasoning):
  * slides stacked vertically, one per screen, gentle scroll-snap, keyboard paging
  * entrance animations replay the PPT order/effects once when a slide scrolls in
  * callout shapes are hidden; hovering the area the tail points at reveals them
  * lines/connectors with arrowheads get a light "flow" running along them,
    block arrows get a soft sheen moving in their direction
  * light 2-layer parallax following the mouse (background vs content)
  * prefers-reduced-motion and touch devices get a calm static version
URL flags for checking: ?static (final state, no motion) ?callouts (show all callouts)
"""
import argparse
import base64
import html
import json
import math
import mimetypes
import os

CALLOUT_TARGET_MAX_DIST = 72  # pt


# ------------------------------------------------------------------ helpers
def esc(s):
    return html.escape(s or "", quote=True)


def css_color(c):
    if not c:
        return "transparent"
    c = c.lstrip("#")
    if len(c) == 8:
        r, g, b, a = (int(c[i:i + 2], 16) for i in (0, 2, 4, 6))
        return "rgba(%d,%d,%d,%.3f)" % (r, g, b, a / 255)
    return "#" + c


def rgb(c):
    c = (c or "#000000").lstrip("#")[:6]
    return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))


def luminance(c):
    r, g, b = (v / 255 for v in rgb(c))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def mix(c, other, t):
    a, b = rgb(c), rgb(other)
    return "#%02X%02X%02X" % tuple(round(x + (y - x) * t) for x, y in zip(a, b))


def flow_color(c):
    return mix(c, "#FFFFFF", 0.6) if luminance(c) < 0.6 else mix(c, "#000000", 0.35)


def f(v):
    return ("%.2f" % v).rstrip("0").rstrip(".")


# ------------------------------------------------------------------ geometry
def rr_path(w, h, r, corners=(1, 1, 1, 1)):
    """Rounded rect, corners = (tl, tr, br, bl) flags."""
    r = max(0, min(r, w / 2, h / 2))
    tl, tr, br, bl = (r if c else 0 for c in corners)
    return ("M{tl} 0 H{a} A{tr} {tr} 0 0 1 {w} {tr} V{b} A{br} {br} 0 0 1 {c} {h} "
            "H{bl} A{bl} {bl} 0 0 1 0 {d} V{tl} A{tl} {tl} 0 0 1 {tl} 0 Z").format(
        tl=f(tl), tr=f(tr), br=f(br), bl=f(bl), w=f(w), h=f(h),
        a=f(w - tr), b=f(h - br), c=f(w - br), d=f(h - bl))


def poly(pts):
    return "M" + " L".join("%s %s" % (f(x), f(y)) for x, y in pts) + " Z"


def callout_rect_path(w, h, tip, r=0):
    """Rect (optionally rounded) with a wedge tail pointing at tip (local coords)."""
    tx, ty = tip
    cx, cy = w / 2, h / 2
    dx, dy = (tx - cx) / max(w, 1), (ty - cy) / max(h, 1)
    inside = 0 <= tx <= w and 0 <= ty <= h
    if inside:
        return rr_path(w, h, r)
    r = max(0, min(r, w / 2, h / 2))
    if abs(dy) >= abs(dx):  # tail on top or bottom
        bw = w / 6
        bc = w * (8 / 12) if tx > cx else w * (4 / 12)
        b1, b2 = bc - bw / 2, bc + bw / 2
        if dy > 0:  # bottom
            return ("M{r} 0 H{a} A{r} {r} 0 0 1 {w} {r} V{b} A{r} {r} 0 0 1 {a} {h} "
                    "H{b2} L{tx} {ty} L{b1} {h} H{r} A{r} {r} 0 0 1 0 {b} V{r} A{r} {r} 0 0 1 {r} 0 Z").format(
                r=f(r), a=f(w - r), b=f(h - r), w=f(w), h=f(h), b1=f(b1), b2=f(b2), tx=f(tx), ty=f(ty))
        return ("M{r} 0 H{b1} L{tx} {ty} L{b2} 0 H{a} A{r} {r} 0 0 1 {w} {r} V{b} A{r} {r} 0 0 1 {a} {h} "
                "H{r} A{r} {r} 0 0 1 0 {b} V{r} A{r} {r} 0 0 1 {r} 0 Z").format(
            r=f(r), a=f(w - r), b=f(h - r), w=f(w), h=f(h), b1=f(b1), b2=f(b2), tx=f(tx), ty=f(ty))
    bh = h / 6
    bc = h * (8 / 12) if ty > cy else h * (4 / 12)
    b1, b2 = bc - bh / 2, bc + bh / 2
    if dx > 0:  # right
        return ("M{r} 0 H{a} A{r} {r} 0 0 1 {w} {r} V{b1} L{tx} {ty} L{w} {b2} V{b} A{r} {r} 0 0 1 {a} {h} "
                "H{r} A{r} {r} 0 0 1 0 {b} V{r} A{r} {r} 0 0 1 {r} 0 Z").format(
            r=f(r), a=f(w - r), b=f(h - r), w=f(w), h=f(h), b1=f(b1), b2=f(b2), tx=f(tx), ty=f(ty))
    return ("M{r} 0 H{a} A{r} {r} 0 0 1 {w} {r} V{b} A{r} {r} 0 0 1 {a} {h} H{r} A{r} {r} 0 0 1 0 {b} "
            "V{b2} L{tx} {ty} L0 {b1} V{r} A{r} {r} 0 0 1 {r} 0 Z").format(
        r=f(r), a=f(w - r), b=f(h - r), w=f(w), h=f(h), b1=f(b1), b2=f(b2), tx=f(tx), ty=f(ty))


def callout_ellipse_path(w, h, tip):
    rx, ry = w / 2, h / 2
    ang = math.atan2((tip[1] - ry) / max(ry, 1), (tip[0] - rx) / max(rx, 1))
    a1, a2 = ang - 0.2, ang + 0.2
    p1 = (rx + rx * math.cos(a1), ry + ry * math.sin(a1))
    p2 = (rx + rx * math.cos(a2), ry + ry * math.sin(a2))
    return "M{} {} A{} {} 0 1 0 {} {} L{} {} Z".format(
        f(p2[0]), f(p2[1]), f(rx), f(ry), f(p1[0]), f(p1[1]), f(tip[0]), f(tip[1]))


def shape_svg_body(e):
    """Returns (svg inner markup for the outline, path d used for clip/sheen)."""
    w, h = max(e["w"], 0.01), max(e["h"], 0.01)
    p = e.get("prst") or "rect"
    a = e.get("adj", {})
    ss = min(w, h)

    def adj(name, default):
        return a.get(name, default) / 100000.0

    tip = None
    if e.get("tail"):
        tip = (e["tail"][0] - e["x"], e["tail"][1] - e["y"])
        if e.get("flipH"):
            tip = (w - tip[0], tip[1])
        if e.get("flipV"):
            tip = (tip[0], h - tip[1])

    if p == "custom":
        ds = []
        for path in e["paths"]:
            sx = w / path["w"] if path.get("w") else 1 / 12700
            sy = h / path["h"] if path.get("h") else 1 / 12700
            parts = []
            for c in path["d"]:
                if c[0] == "Z":
                    parts.append("Z")
                    continue
                nums = c[1:]
                parts.append(c[0] + " ".join(f(v * (sx if i % 2 == 0 else sy)) for i, v in enumerate(nums)))
            ds.append((" ".join(parts), path["fill"], path["stroke"]))
        return ds
    if p in ("ellipse", "flowChartConnector"):
        d = "M0 {ry} A{rx} {ry} 0 1 0 {w} {ry} A{rx} {ry} 0 1 0 0 {ry} Z".format(rx=f(w / 2), ry=f(h / 2), w=f(w))
    elif p == "roundRect":
        d = rr_path(w, h, ss * adj("adj", 16667))
    elif p == "flowChartAlternateProcess":
        d = rr_path(w, h, ss * 0.1667)
    elif p == "flowChartTerminator":
        d = rr_path(w, h, h / 2)
    elif p == "round1Rect":
        d = rr_path(w, h, ss * adj("adj", 16667), (0, 1, 0, 0))
    elif p == "round2SameRect":
        d = rr_path(w, h, ss * adj("adj1", 16667), (1, 1, 0, 0))
    elif p == "snip1Rect":
        s = ss * adj("adj", 16667)
        d = poly([(0, 0), (w - s, 0), (w, s), (w, h), (0, h)])
    elif p == "snip2SameRect":
        s = ss * adj("adj1", 16667)
        d = poly([(s, 0), (w - s, 0), (w, s), (w, h), (0, h), (0, s)])
    elif p == "triangle":
        d = poly([(w * adj("adj", 50000), 0), (w, h), (0, h)])
    elif p == "rtTriangle":
        d = poly([(0, 0), (w, h), (0, h)])
    elif p in ("diamond", "flowChartDecision"):
        d = poly([(w / 2, 0), (w, h / 2), (w / 2, h), (0, h / 2)])
    elif p == "parallelogram":
        o = ss * adj("adj", 25000)
        d = poly([(o, 0), (w, 0), (w - o, h), (0, h)])
    elif p == "trapezoid":
        o = ss * adj("adj", 25000)
        d = poly([(o, 0), (w - o, 0), (w, h), (0, h)])
    elif p == "pentagon" or p == "homePlate":
        o = w - ss * adj("adj", 50000)
        d = poly([(0, 0), (o, 0), (w, h / 2), (o, h), (0, h)])
    elif p == "chevron":
        o = ss * adj("adj", 50000)
        d = poly([(0, 0), (w - o, 0), (w, h / 2), (w - o, h), (0, h), (o, h / 2)])
    elif p == "hexagon":
        o = ss * adj("adj", 25000)
        d = poly([(o, 0), (w - o, 0), (w, h / 2), (w - o, h), (o, h), (0, h / 2)])
    elif p == "octagon":
        o = ss * adj("adj", 29289)
        d = poly([(o, 0), (w - o, 0), (w, o), (w, h - o), (w - o, h), (o, h), (0, h - o), (0, o)])
    elif p == "plus":
        o = ss * adj("adj", 25000)
        d = poly([(o, 0), (w - o, 0), (w - o, o), (w, o), (w, h - o), (w - o, h - o), (w - o, h),
                  (o, h), (o, h - o), (0, h - o), (0, o), (o, o)])
    elif p == "donut":
        o = ss * adj("adj", 25000)
        d = ("M0 {ry} A{rx} {ry} 0 1 0 {w} {ry} A{rx} {ry} 0 1 0 0 {ry} Z "
             "M{o} {ry} A{irx} {iry} 0 1 1 {wo} {ry} A{irx} {iry} 0 1 1 {o} {ry} Z").format(
            rx=f(w / 2), ry=f(h / 2), w=f(w), o=f(o), wo=f(w - o), irx=f(w / 2 - o), iry=f(h / 2 - o))
    elif p in ("rightArrow", "notchedRightArrow", "stripedRightArrow", "leftArrow", "leftRightArrow"):
        t = h * adj("adj1", 50000) / 2
        hl = ss * adj("adj2", 50000)
        y1, y2 = h / 2 - t, h / 2 + t
        if p == "leftArrow":
            d = poly([(w, y1), (hl, y1), (hl, 0), (0, h / 2), (hl, h), (hl, y2), (w, y2)])
        elif p == "leftRightArrow":
            d = poly([(0, h / 2), (hl, 0), (hl, y1), (w - hl, y1), (w - hl, 0), (w, h / 2),
                      (w - hl, h), (w - hl, y2), (hl, y2), (hl, h)])
        elif p == "notchedRightArrow":
            n = (h / 2 - y1) * hl / max(h / 2, 1)
            d = poly([(0, y1), (w - hl, y1), (w - hl, 0), (w, h / 2), (w - hl, h), (w - hl, y2), (0, y2), (n, h / 2)])
        else:
            d = poly([(0, y1), (w - hl, y1), (w - hl, 0), (w, h / 2), (w - hl, h), (w - hl, y2), (0, y2)])
    elif p in ("upArrow", "downArrow", "upDownArrow"):
        t = w * adj("adj1", 50000) / 2
        hl = ss * adj("adj2", 50000)
        x1, x2 = w / 2 - t, w / 2 + t
        if p == "upArrow":
            d = poly([(x1, h), (x1, hl), (0, hl), (w / 2, 0), (w, hl), (x2, hl), (x2, h)])
        elif p == "downArrow":
            d = poly([(x1, 0), (x2, 0), (x2, h - hl), (w, h - hl), (w / 2, h), (0, h - hl), (x1, h - hl)])
        else:
            d = poly([(w / 2, 0), (w, hl), (x2, hl), (x2, h - hl), (w, h - hl), (w / 2, h),
                      (0, h - hl), (x1, h - hl), (x1, hl), (0, hl)])
    elif p in ("wedgeRectCallout", "wedgeRoundRectCallout"):
        r = ss * 0.16667 if p == "wedgeRoundRectCallout" else 0
        d = callout_rect_path(w, h, tip or (w * 0.3, h * 1.6), r)
    elif p == "wedgeEllipseCallout":
        d = callout_ellipse_path(w, h, tip or (w * 0.3, h * 1.6))
    elif p == "cloudCallout":
        # ellipse body + scalloped edge (small arcs) + trailing bubbles toward the tip
        rx, ry, n = w / 2, h / 2, 14
        pts = [(rx + rx * math.cos(2 * math.pi * i / n), ry + ry * math.sin(2 * math.pi * i / n)) for i in range(n)]
        d = "M%s %s " % (f(pts[0][0]), f(pts[0][1]))
        for i in range(1, n + 1):
            q = pts[i % n]
            d += "A%s %s 0 0 1 %s %s " % (f(ss * 0.16), f(ss * 0.16), f(q[0]), f(q[1]))
        d += "Z"
        if tip:
            for k, s in ((0.55, 0.09), (0.8, 0.06), (1.0, 0.035)):
                bx, by = rx + (tip[0] - rx) * k, ry + (tip[1] - ry) * k
                rr = ss * s
                d += " M%s %s A%s %s 0 1 0 %s %s A%s %s 0 1 0 %s %s Z" % (
                    f(bx - rr), f(by), f(rr), f(rr), f(bx + rr), f(by), f(rr), f(rr), f(bx - rr), f(by))
    else:  # rect, flowChartProcess and anything simple
        d = "M0 0 H{w} V{h} H0 Z".format(w=f(w), h=f(h))
    return [(d, True, True)]


# ------------------------------------------------------------------ builder
class Builder:
    def __init__(self, work, include_hidden=False):
        self.work = work
        self.data = json.load(open(os.path.join(work, "slides.json"), encoding="utf-8"))
        self.W, self.H = self.data["width"], self.data["height"]
        self.include_hidden = include_hidden
        self.uid = 0
        self.uri_cache = {}

    def cq(self, pt):
        return "%scqw" % f(pt * 100.0 / self.W)

    def pct(self, v, total):
        return "%s%%" % ("%.3f" % (v * 100.0 / total)).rstrip("0").rstrip(".")

    def data_uri(self, rel):
        if rel in self.uri_cache:
            return self.uri_cache[rel]
        path = os.path.join(self.work, rel)
        mime = mimetypes.guess_type(path)[0] or "image/png"
        with open(path, "rb") as fh:
            uri = "data:%s;base64,%s" % (mime, base64.b64encode(fh.read()).decode())
        self.uri_cache[rel] = uri
        return uri

    def next_id(self, p):
        self.uid += 1
        return "%s%d" % (p, self.uid)

    def paint(self, fill, defs):
        if not fill or fill["type"] == "none":
            return "none"
        if fill["type"] == "solid":
            return css_color(fill["color"])
        if fill["type"] == "gradient":
            gid = self.next_id("g")
            stops = "".join('<stop offset="%s%%" stop-color="%s"/>' % (f(s["pos"]), css_color(s["color"]))
                            for s in fill["stops"])
            if fill.get("shape") == "radial":
                defs.append('<radialGradient id="%s">%s</radialGradient>' % (gid, stops))
            else:
                ang = math.radians(fill.get("angle", 0))
                x1, y1 = 0.5 - math.cos(ang) / 2, 0.5 - math.sin(ang) / 2
                x2, y2 = 0.5 + math.cos(ang) / 2, 0.5 + math.sin(ang) / 2
                defs.append('<linearGradient id="%s" x1="%s" y1="%s" x2="%s" y2="%s">%s</linearGradient>' % (
                    gid, f(x1), f(y1), f(x2), f(y2), stops))
            return "url(#%s)" % gid
        if fill["type"] == "image" and fill.get("src"):
            pid = self.next_id("p")
            defs.append('<pattern id="%s" patternContentUnits="objectBoundingBox" width="1" height="1">'
                        '<image href="%s" width="1" height="1" preserveAspectRatio="xMidYMid slice"/></pattern>'
                        % (pid, self.data_uri(fill["src"])))
            return "url(#%s)" % pid
        return "none"

    def css_background(self, bg):
        if bg["type"] == "solid":
            return "background:%s" % css_color(bg["color"])
        if bg["type"] == "gradient":
            stops = ", ".join("%s %s%%" % (css_color(s["color"]), f(s["pos"])) for s in bg["stops"])
            if bg.get("shape") == "radial":
                return "background:radial-gradient(circle, %s)" % stops
            return "background:linear-gradient(%sdeg, %s)" % (f(bg.get("angle", 0) + 90), stops)
        if bg["type"] == "image":
            return "background:url('%s') center/cover" % self.data_uri(bg["src"])
        return "background:#fff"

    # ------------------------------------------------------------ text
    def text_html(self, t):
        if not t:
            return ""
        just = {"t": "flex-start", "ctr": "center", "b": "flex-end", "just": "flex-start",
                "dist": "flex-start"}.get(t["anchor"], "flex-start")
        ins = t["inset"]
        style = "justify-content:%s;padding:%s %s %s %s;" % (
            just, self.cq(ins["t"]), self.cq(ins["r"]), self.cq(ins["b"]), self.cq(ins["l"]))
        if not t.get("wrap", True):
            style += "white-space:nowrap;"
        if t.get("vert") in ("eaVert", "vert", "wordArtVert"):
            style += "writing-mode:vertical-rl;"
        paras = []
        num = 0
        for para in t["paragraphs"]:
            align = {"l": "left", "ctr": "center", "r": "right", "just": "justify", "dist": "justify"}.get(
                para["align"], "left")
            sizes = [r["size"] for r in para["runs"] if not r.get("br")]
            base = max(sizes) if sizes else (para.get("emptySize") or 12)
            lh = 1.2 * (para["line"] or 1.0)
            pst = "text-align:%s;font-size:%s;line-height:%s;" % (align, self.cq(base), f(lh))
            if para.get("spaceBefore"):
                pst += "margin-top:%s;" % self.cq(para["spaceBefore"])
            if para.get("marL"):
                pst += "padding-left:%s;" % self.cq(para["marL"])
            inner = []
            if para.get("bullet") and para["runs"]:
                num += 1
                ch = "%d." % num if para["bullet"] == "#" else para["bullet"]
                first = next((r for r in para["runs"] if not r.get("br")), None)
                inner.append('<span class="bu" style="color:%s">%s</span>' % (
                    css_color(first["color"]) if first else "inherit", esc(ch)))
            for r in para["runs"]:
                if r.get("br"):
                    inner.append("<br>")
                    continue
                fam = ", ".join('"%s"' % x for x in (r.get("font"), r.get("font_ea")) if x)
                rs = "font-size:%s;color:%s;font-family:%s, var(--fallback);" % (
                    self.cq(r["size"]), css_color(r["color"]), fam or "var(--fallback)")
                if r["bold"]:
                    rs += "font-weight:700;"
                if r["italic"]:
                    rs += "font-style:italic;"
                if r["underline"]:
                    rs += "text-decoration:underline;"
                inner.append('<span style="%s">%s</span>' % (rs, esc(r["text"]).replace("\v", "<br>").replace("\n", "<br>")))
            if not para["runs"]:
                inner.append("&nbsp;")
            paras.append('<p style="%s">%s</p>' % (pst, "".join(inner)))
        return '<div class="tx" style="%s">%s</div>' % (style, "".join(paras))

    # ------------------------------------------------------------ elements
    def box_style(self, e, extra=""):
        st = "left:%s;top:%s;width:%s;height:%s;" % (
            self.pct(e["x"], self.W), self.pct(e["y"], self.H),
            self.pct(max(e["w"], 0.5), self.W), self.pct(max(e["h"], 0.5), self.H))
        if e.get("rot"):
            st += "--rot:%sdeg;" % f(e["rot"])
        return st + extra

    def shape(self, e, layer):
        w, h = max(e["w"], 0.01), max(e["h"], 0.01)
        defs = []
        fill = self.paint(e.get("fill"), defs)
        ln = e.get("line") or {}
        stroke = self.paint(ln.get("fill"), defs) if ln.get("fill") else "none"
        sw = ln.get("width", 0.75)
        dash = ' stroke-dasharray="%s %s"' % (f(sw * 3), f(sw * 2)) if ln.get("dash") not in (None, "solid") else ""
        bodies = shape_svg_body(e)
        paths = []
        for d, do_fill, do_stroke in bodies:
            paths.append('<path d="%s" fill="%s" stroke="%s" stroke-width="%s"%s fill-rule="evenodd"/>' % (
                d, fill if do_fill else "none", stroke if do_stroke else "none", f(sw), dash))
        sheen = ""
        if e.get("arrow") and fill != "none":
            cid = self.next_id("c")
            p = e["prst"]
            axis = "y" if p in ("upArrow", "downArrow", "upDownArrow") else "x"
            rev = p in ("leftArrow", "upArrow")
            span = w if axis == "x" else h
            gid = self.next_id("s")
            if axis == "x":
                grad = '<linearGradient id="%s" x1="0" y1="0" x2="1" y2="0">' % gid
                rect = '<rect class="sheen sheen-x%s" x="%s" y="0" width="%s" height="%s" fill="url(#%s)" style="--span:%s"/>' % (
                    " rev" if rev else "", f(-span * 0.45), f(span * 0.45), f(h), gid, f(span * 1.45))
            else:
                grad = '<linearGradient id="%s" x1="0" y1="0" x2="0" y2="1">' % gid
                rect = '<rect class="sheen sheen-y%s" x="0" y="%s" width="%s" height="%s" fill="url(#%s)" style="--span:%s"/>' % (
                    " rev" if rev else "", f(-span * 0.45), f(w), f(span * 0.45), gid, f(span * 1.45))
            col = flow_color(e["fill"]["color"]) if e["fill"]["type"] == "solid" else "#FFFFFF"
            defs.append(grad + '<stop offset="0" stop-color="%s" stop-opacity="0"/>'
                               '<stop offset=".5" stop-color="%s" stop-opacity=".55"/>'
                               '<stop offset="1" stop-color="%s" stop-opacity="0"/></linearGradient>' % (col, col, col))
            defs.append('<clipPath id="%s"><path d="%s"/></clipPath>' % (cid, bodies[0][0]))
            sheen = '<g clip-path="url(#%s)">%s</g>' % (cid, rect)
        flip = ""
        if e.get("flipH") or e.get("flipV"):
            flip = ' style="transform:scale(%d,%d)"' % (-1 if e.get("flipH") else 1, -1 if e.get("flipV") else 1)
        cls = "shadow" if e.get("shadow") else ""
        svg = '<svg class="geo %s" viewBox="0 0 %s %s" preserveAspectRatio="none"%s><defs>%s</defs>%s%s</svg>' % (
            cls, f(w), f(h), flip, "".join(defs), "".join(paths), sheen)
        return svg + self.text_html(e.get("text"))

    def picture(self, e):
        crop = e.get("crop") or {}
        l, t, r, b = (crop.get(k, 0) for k in ("l", "t", "r", "b"))
        vw, vh = max(1 - l - r, 0.01), max(1 - t - b, 0.01)
        img = '<img alt="" src="%s" style="width:%s%%;height:%s%%;left:%s%%;top:%s%%">' % (
            self.data_uri(e["src"]), f(100 / vw), f(100 / vh), f(-l / vw * 100), f(-t / vh * 100))
        radius = ""
        if e.get("prst") in ("ellipse",):
            radius = "border-radius:50%;"
        elif e.get("prst") == "roundRect":
            radius = "border-radius:%s;" % self.cq(min(e["w"], e["h"]) * 0.16667)
        ln = e.get("line") or {}
        border = ""
        if ln.get("fill") and ln["fill"].get("type") == "solid":
            border = "outline:%s solid %s;outline-offset:-%s;" % (
                self.cq(ln.get("width", 0.75)), css_color(ln["fill"]["color"]), self.cq(ln.get("width", 0.75)))
        return '<div class="pic" style="%s%s">%s</div>' % (radius, border, img)

    def line_points(self, e):
        """Path in slide coords (pt) and its approximate length."""
        x, y, w, h = e["x"], e["y"], e["w"], e["h"]
        sx, sy = (x + w, x) if e.get("flipH") else (x, x + w)
        ty0, ty1 = (y + h, y) if e.get("flipV") else (y, y + h)
        p = e.get("prst", "line")
        a = e.get("adj", {})
        if p.startswith("bentConnector"):
            if p == "bentConnector2":
                pts = [(sx, ty0), (sy, ty0), (sy, ty1)]
            else:
                mx = sx + (sy - sx) * a.get("adj1", 50000) / 100000.0
                pts = [(sx, ty0), (mx, ty0), (mx, ty1), (sy, ty1)]
            curve = False
        elif p.startswith("curvedConnector"):
            mx = sx + (sy - sx) * a.get("adj1", 50000) / 100000.0
            pts = [(sx, ty0), (mx, ty0), (mx, ty1), (sy, ty1)]
            curve = True
        else:
            pts = [(sx, ty0), (sy, ty1)]
            curve = False
        if e.get("rot"):
            cx, cy = x + w / 2, y + h / 2
            ang = math.radians(e["rot"])
            pts = [(cx + (px - cx) * math.cos(ang) - (py - cy) * math.sin(ang),
                    cy + (px - cx) * math.sin(ang) + (py - cy) * math.cos(ang)) for px, py in pts]
        length = sum(math.dist(pts[i], pts[i + 1]) for i in range(len(pts) - 1))
        if curve:
            d = "M%s %s C%s %s %s %s %s %s" % tuple(f(v) for pt in pts for v in pt)
        else:
            d = "M" + " L".join("%s %s" % (f(px), f(py)) for px, py in pts)
        return d, length, pts

    def marker(self, end, color, defs, start):
        mid = self.next_id("m")
        size = {"sm": 2.2, "med": 3.2, "lg": 4.6}
        mw, ml = size.get(end.get("w"), 3.2), size.get(end.get("len"), 3.2)
        t = end.get("type")
        orient = "auto-start-reverse" if start else "auto"
        if t in ("oval",):
            shape = '<ellipse cx="%s" cy="%s" rx="%s" ry="%s" fill="%s"/>' % (f(ml / 2), f(mw / 2), f(ml / 2), f(mw / 2), color)
            ref = ml / 2
        elif t == "diamond":
            shape = '<path d="M0 %s L%s 0 L%s %s L%s %s Z" fill="%s"/>' % (
                f(mw / 2), f(ml / 2), f(ml), f(mw / 2), f(ml / 2), f(mw), color)
            ref = ml / 2
        elif t == "arrow":
            shape = '<path d="M0 0 L%s %s L0 %s" fill="none" stroke="%s" stroke-width=".7" stroke-linejoin="round"/>' % (
                f(ml), f(mw / 2), f(mw), color)
            ref = ml - 0.4
        elif t == "stealth":
            shape = '<path d="M0 0 L%s %s L0 %s L%s %s Z" fill="%s"/>' % (f(ml), f(mw / 2), f(mw), f(ml * 0.3), f(mw / 2), color)
            ref = ml * 0.8
        else:  # triangle
            shape = '<path d="M0 0 L%s %s L0 %s Z" fill="%s"/>' % (f(ml), f(mw / 2), f(mw), color)
            ref = ml * 0.8
        defs.append('<marker id="%s" viewBox="-0.5 -0.5 %s %s" markerWidth="%s" markerHeight="%s" refX="%s" refY="%s" '
                    'orient="%s" markerUnits="strokeWidth" overflow="visible">%s</marker>' % (
                        mid, f(ml + 1), f(mw + 1), f(ml + 1), f(mw + 1), f(ref), f(mw / 2), orient, shape))
        return "url(#%s)" % mid

    def line(self, e):
        ln = e.get("line") or {}
        fill = ln.get("fill") or {}
        if fill.get("type") != "solid":
            return None
        color = css_color(fill["color"])
        sw = max(ln.get("width", 0.75), 0.5)
        d, length, pts = self.line_points(e)
        defs = []
        attrs = ""
        head, tail = ln.get("headEnd"), ln.get("tailEnd")
        if head:
            attrs += ' marker-start="%s"' % self.marker(head, color, defs, True)
        if tail:
            attrs += ' marker-end="%s"' % self.marker(tail, color, defs, False)
        dash = ' stroke-dasharray="%s %s"' % (f(sw * 3), f(sw * 2)) if ln.get("dash") not in (None, "solid") else ""
        base = '<path class="ln" d="%s" fill="none" stroke="%s" stroke-width="%s" stroke-linejoin="round"%s%s/>' % (
            d, color, f(sw), dash, attrs)
        flow = ""
        if head or tail:
            seg = max(8.0, sw * 6)
            gap = max(28.0, min(70.0, length / 3))
            period = seg + gap
            dur = period / 55.0  # ~55pt per second, calm
            direction = "rev" if head and not tail else ""
            flow = ('<path class="flow %s" d="%s" fill="none" stroke="%s" stroke-width="%s" stroke-linecap="round" '
                    'stroke-dasharray="%s %s" style="--period:%s;--fdur:%ss"/>') % (
                direction, d, flow_color(fill["color"]), f(sw * 0.9), f(seg), f(gap), f(period), f(dur))
        # the element box is the line's bounding box (+room for arrowheads), so a
        # wipe entrance reveals the line itself, from its start point to its end
        pad = max(6.0, sw * 6)
        xs, ys = [p[0] for p in pts], [p[1] for p in pts]
        bx, by = min(xs) - pad, min(ys) - pad
        bw, bh = max(xs) - min(xs) + 2 * pad, max(ys) - min(ys) + 2 * pad
        (x0, y0), (x1, y1) = pts[0], pts[-1]
        if head and not tail:
            (x0, y0), (x1, y1) = (x1, y1), (x0, y0)
        if abs(x1 - x0) >= abs(y1 - y0):
            wdir = "left" if x1 >= x0 else "right"
        else:
            wdir = "top" if y1 >= y0 else "bottom"
        svg = ('<svg class="wire" viewBox="%s %s %s %s" preserveAspectRatio="none"><defs>%s</defs>%s%s</svg>'
               % (f(bx), f(by), f(bw), f(bh), "".join(defs), base, flow))
        return svg, (bx, by, bw, bh), wdir

    # ------------------------------------------------------------ callout targets
    def callout_targets(self, elements):
        cands = [e for e in elements if not e.get("callout") and e["kind"] in ("shape", "picture", "raster")
                 and e["w"] * e["h"] < 0.6 * self.W * self.H]
        res = {}
        for c in [e for e in elements if e.get("callout")]:
            target = None
            if c.get("tail"):
                tx, ty = c["tail"]
                inside = [e for e in cands if e["x"] - 2 <= tx <= e["x"] + e["w"] + 2 and e["y"] - 2 <= ty <= e["y"] + e["h"] + 2]
                if inside:
                    target = min(inside, key=lambda e: e["w"] * e["h"])
                else:
                    def dist(e):
                        dx = max(e["x"] - tx, 0, tx - e["x"] - e["w"])
                        dy = max(e["y"] - ty, 0, ty - e["y"] - e["h"])
                        return math.hypot(dx, dy)
                    near = sorted(cands, key=dist)
                    if near and dist(near[0]) <= CALLOUT_TARGET_MAX_DIST:
                        target = near[0]
                if target is None:
                    s = 44
                    target = {"id": "hot-%s" % c["id"], "x": tx - s / 2, "y": ty - s / 2, "w": s, "h": s, "virtual": True}
            else:
                def bdist(e):
                    dx = max(e["x"] - (c["x"] + c["w"]), 0, c["x"] - (e["x"] + e["w"]))
                    dy = max(e["y"] - (c["y"] + c["h"]), 0, c["y"] - (e["y"] + e["h"]))
                    return math.hypot(dx, dy)
                pool = [e for e in cands if not (e["x"] <= c["x"] and e["y"] <= c["y"] and
                                                  e["x"] + e["w"] >= c["x"] + c["w"] and e["y"] + e["h"] >= c["y"] + c["h"])]
                if pool:
                    target = min(pool, key=lambda e: (round(bdist(e)), e["w"] * e["h"]))
            if target is not None:
                res.setdefault(target["id"], {"target": target, "callouts": []})["callouts"].append(c)
        return res

    # ------------------------------------------------------------ slide
    def slide(self, s):
        els = s["elements"]
        area = self.W * self.H
        anims = {}
        for a in s["animations"]:
            if a["class"] in ("entr", "emph"):
                anims.setdefault(a["spid"], []).append(a)
        targets = self.callout_targets(els)
        target_ids = set(targets)
        parts = []
        timeline = []
        for z, e in enumerate(els):
            if e["kind"] == "raster" and not e.get("src"):
                continue
            is_bg = e["origin"] != "slide" or (e["w"] * e["h"] > 0.55 * area and e["kind"] in ("shape", "picture"))
            layer = "bg" if is_bg else "fg"
            ea = anims.get(e["id"]) if e["origin"] == "slide" else None
            entr = next((a for a in (ea or []) if a["class"] == "entr"), None)
            emph = next((a for a in (ea or []) if a["class"] == "emph"), None)
            elid = "s%d-%s" % (s["index"], e["id"]) if e["origin"] == "slide" else "s%d-%s-%s" % (s["index"], e["origin"], e["id"])
            classes = ["el", layer]
            data = ""
            if e.get("callout"):
                classes.append("callout")
                layer = "co"
                classes[1] = "co"
            if e["id"] in target_ids and e["origin"] == "slide":
                classes.append("target")
                data += ' data-co="%s"' % " ".join("s%d-%s" % (s["index"], c["id"]) for c in targets[e["id"]]["callouts"])
            wdir = None
            if e["kind"] == "line":
                r = self.line(e)
                if not r:
                    continue
                inner, (bx, by, bw, bh), wdir = r
                style = self.box_style({"x": bx, "y": by, "w": bw, "h": bh})
                classes.append("is-line")
            else:
                style = self.box_style(e)
                if e["kind"] == "shape":
                    inner = self.shape(e, layer)
                elif e["kind"] == "picture":
                    inner = self.picture(e)
                else:
                    inner = '<img class="raster" alt="" src="%s">' % self.data_uri(e["src"])
            if entr and not e.get("callout"):
                eff, edir = entr["effect"], entr["direction"]
                if e["kind"] == "line" and eff != "appear":
                    eff, edir = "wipe", wdir  # lines draw themselves from start to end
                classes.append("pre")
                data += ' data-fx="%s" data-dir="%s"' % (eff, edir)
                timeline.append({"id": elid, "kind": "in", "trigger": entr["trigger"], "order": entr["order"],
                                 "delay": entr["delay"], "dur": entr["dur"]})
            if emph and not e.get("callout"):
                timeline.append({"id": elid, "kind": "pulse", "trigger": emph["trigger"], "order": emph["order"],
                                 "delay": emph["delay"], "dur": emph["dur"]})
            if e.get("flipH") and e["kind"] != "shape" and e["kind"] != "line":
                style += "--fx:-1;"
            parts.append('<div id="%s" class="%s" style="%s"%s><div class="anim">%s</div></div>' % (
                elid, " ".join(classes), style, data, inner))
        # hover hotspots on top (so overlapping text boxes never block the hover)
        hots = []
        for tid, info in targets.items():
            t = info["target"]
            co = " ".join("s%d-%s" % (s["index"], c["id"]) for c in info["callouts"])
            tgt = "" if t.get("virtual") else "s%d-%s" % (s["index"], tid)
            accent = next((c["fill"]["color"] for c in info["callouts"]
                           if c.get("fill", {}).get("type") == "solid"), "#3B82F6")
            hots.append('<div class="hot" style="%s--accent:%s" data-co="%s" data-target="%s"><i class="hint"></i></div>' % (
                self.box_style(dict(t, rot=0)), css_color(accent)[:7] if accent.startswith("#") else css_color(accent), co, tgt))
        timeline.sort(key=lambda a: a["order"])
        return ('<section class="slide" id="slide-%d" data-index="%d"><div class="stage" style="%s" '
                'data-timeline=\'%s\'>%s%s</div></section>') % (
            s["index"], s["index"], self.css_background(s["background"]),
            esc(json.dumps(timeline)), "".join(parts), "".join(hots))

    def build(self, title):
        slides = [s for s in self.data["slides"] if self.include_hidden or not s.get("hidden")]
        body = "".join(self.slide(s) for s in slides)
        tpl = open(os.path.join(os.path.dirname(__file__), "..", "assets", "template.html"), encoding="utf-8").read()
        return (tpl.replace("{{TITLE}}", esc(title))
                .replace("{{RATIO}}", "%s / %s" % (f(self.W), f(self.H)))
                .replace("{{SLIDES}}", body)
                .replace("{{COUNT}}", str(len(slides))))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("work")
    ap.add_argument("out")
    ap.add_argument("--title", default="Presentation")
    ap.add_argument("--include-hidden", action="store_true")
    a = ap.parse_args()
    html_out = Builder(a.work, a.include_hidden).build(a.title)
    with open(a.out, "w", encoding="utf-8") as fh:
        fh.write(html_out)
    print("wrote %s (%.1f KB)" % (a.out, len(html_out.encode()) / 1024))
