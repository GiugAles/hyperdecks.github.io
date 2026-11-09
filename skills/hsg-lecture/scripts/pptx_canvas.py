#!/usr/bin/env python3
"""Copy a slide of a .pptx into a deck.json as a 'canvas' slide: the shapes, arrows, text, pictures and tables
of the PowerPoint slide with their positions, and its click animations as steps.

  python pptx_canvas.py deck.json --pptx lecture.pptx --slide 60            # append as the last slide
  python pptx_canvas.py deck.json --pptx lecture.pptx --slide 60 --at 12    # replace slide 12 of the deck
  python pptx_canvas.py deck.json --pptx lecture.pptx --slide 60 --insert 12  # insert before slide 12 of the deck

The title comes from the title placeholder of the slide (edit it in the deck afterwards if you like). Pictures are
saved in img/ next to the deck as s<slide>-<n>.<ext>. Everything is measured in U, 1% of the slide width, as
render_html.py does, so a slide is placed exactly where it was in PowerPoint. The canvas has no 'sizes of its own':
type sizes are those of the PowerPoint slide (18.7pt body on the HSG template).

Item format (see render_html.py, Html.cv_item):
  {"k":"shape", "g":"rect|roundRect|ellipse|line|...", "x","y","w","h" (U), "fill":"#rrggbb"|null, "fa":alpha,
   "ln":"#rrggbb"|null, "lw":U, "dash":..., "rot":degrees, "fx":bool, "fy":bool, "adj":[...], "d":"svg path in U",
   "head":"triangle", "tail":"triangle", "tx":{text}, "s":first step, "e":step where it disappears}
  {"k":"pic", "src":"img/..", "alt":"..", "x","y","w","h", "crop":[l,t,r,b], "shadow":bool, "s":..}
  {"k":"table", "x","y","w","h", "cols":[U], "rows":[{"h":U,"cells":[{"tx":..,"fill":..}]}], "s":..}
"""
import argparse, colorsys, hashlib, json, math, re, sys, zipfile
from pathlib import Path
from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"a": A, "p": P, "r": R}
DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram"
PT = 1 / 9.6          # 1pt in U (the template slide is 13.33in wide, 1U = 9.6pt)
SKIP_PH = {"title", "ctrTitle", "sldNum", "ftr", "dt"}
warnings = []


def warn(msg):
    if msg not in warnings:
        warnings.append(msg)


class Ctx:
    def __init__(self, prs, slide):
        self.k = 100 / prs.slide_width          # U per EMU
        self.pt = self.k * 12700                # U per pt: 1/9.6 on a 13.33in slide, 1/7.2 on a 10in slide
        self.defsz = {}                         # default type size of text boxes per level (presentation defaultTextStyle)
        dts = prs.part._element.find("p:defaultTextStyle", NS)
        for e in (dts if dts is not None else []):
            m = re.match(r"lvl(\d)pPr", etree.QName(e).localname)
            if m and e.find("a:defRPr", NS) is not None and e.find("a:defRPr", NS).get("sz"):
                self.defsz[int(m.group(1)) - 1] = int(e.find("a:defRPr", NS).get("sz")) / 100
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT
        theme = etree.fromstring(slide.slide_layout.slide_master.part.part_related_by(RT.THEME).blob)
        self.theme = {}
        for c in theme.find(".//a:clrScheme", NS):
            name = etree.QName(c).localname
            e = c[0]
            self.theme[name] = e.get("val") if etree.QName(e).localname == "srgbClr" else e.get("lastClr", "000000")
        self.theme.update(tx1=self.theme["dk1"], bg1=self.theme["lt1"], tx2=self.theme["dk2"], bg2=self.theme["lt2"])


def clr(ctx, parent):
    """Colour of the first colour child of parent: (hex, alpha) or None."""
    if parent is None:
        return None
    for e in parent:
        n = etree.QName(e).localname
        if n not in ("srgbClr", "schemeClr", "prstClr", "sysClr"):
            continue
        if n == "srgbClr":
            h = e.get("val")
        elif n == "schemeClr":
            h = ctx.theme.get(e.get("val"), "000000")
        elif n == "sysClr":
            h = e.get("lastClr", "000000")
        else:
            h = {"black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "008000", "blue": "0000FF", "yellow": "FFFF00"}.get(e.get("val"), "808080")
        r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
        hh, l, s = colorsys.rgb_to_hls(r, g, b)
        alpha = 1.0
        for m in e:
            mn, v = etree.QName(m).localname, int(m.get("val", "0")) / 100000
            if mn == "lumMod":
                l *= v
            elif mn == "lumOff":
                l += v
            elif mn == "satMod":
                s = min(1, s * v)
            elif mn == "alpha":
                alpha = v
            elif mn in ("tint", "shade"):                   # PowerPoint mixes with white or black in linear RGB
                lin = [(c / 12.92 if c <= .04045 else ((c + .055) / 1.055) ** 2.4) for c in colorsys.hls_to_rgb(hh, max(0, min(1, l)), s)]
                lin = [c * v + (1 - v) if mn == "tint" else c * v for c in lin]
                hh, l, s = colorsys.rgb_to_hls(*[(c * 12.92 if c <= .0031308 else 1.055 * c ** (1 / 2.4) - .055) for c in lin])
        r, g, b = colorsys.hls_to_rgb(hh, max(0, min(1, l)), s)
        return "#%02x%02x%02x" % (round(r * 255), round(g * 255), round(b * 255)), alpha
    return None


def style_ref(el, name):
    st = el.find("p:style", NS)
    return None if st is None else st.find("a:" + name, NS)


def fill_of(ctx, el):
    spPr = el.find("p:spPr", NS)
    if spPr is not None:
        if spPr.find("a:noFill", NS) is not None:
            return None
        sf = spPr.find("a:solidFill", NS)
        if sf is not None:
            return clr(ctx, sf)
        gf = spPr.find("a:gradFill/a:gsLst/a:gs", NS)
        if gf is not None:
            return clr(ctx, gf)
    fr = style_ref(el, "fillRef")
    if fr is not None and fr.get("idx") not in ("0", None):
        return clr(ctx, fr)
    return None


def line_of(ctx, el):
    """(colour, width U, dash, head, tail) or None."""
    spPr = el.find("p:spPr", NS)
    ln = spPr.find("a:ln", NS) if spPr is not None else None
    lr = style_ref(el, "lnRef")
    if ln is not None and ln.find("a:noFill", NS) is not None:
        return None
    col = None
    if ln is not None and ln.find("a:solidFill", NS) is not None:
        col = clr(ctx, ln.find("a:solidFill", NS))
    elif lr is not None and lr.get("idx") != "0":
        col = clr(ctx, lr)
    if col is None:
        return None
    w = int(ln.get("w")) if ln is not None and ln.get("w") else 12700
    dash = ln.find("a:prstDash", NS).get("val") if ln is not None and ln.find("a:prstDash", NS) is not None else None
    he = ln.find("a:headEnd", NS) if ln is not None else None
    te = ln.find("a:tailEnd", NS) if ln is not None else None
    ends = [e.get("type") if e is not None and e.get("type") not in (None, "none") else None for e in (he, te)]
    return col[0], w * ctx.k, dash, ends[0], ends[1]


# ---- text

def default_color(ctx, el):
    fr = style_ref(el, "fontRef")
    c = clr(ctx, fr) if fr is not None else None
    return c[0] if c else None


def text_of(ctx, el, tx, placeholder=False, default_col=None, inh=()):
    """The 'tx' object of a shape from its txBody element, or None when there is no text."""
    if tx is None:
        return None
    body = tx.find("a:bodyPr", NS)
    bodies = [b for b in [body] + [b for b, _ in inh] if b is not None]   # a placeholder inherits from layout and master

    def battr(k, d=None):
        return next((b.get(k) for b in bodies if b.get(k) is not None), d)
    scale = 1.0
    if body is not None and body.find("a:normAutofit", NS) is not None:
        scale = int(body.find("a:normAutofit", NS).get("fontScale", "100000")) / 100000
    lst, lal, lls = {}, {}, {}
    for ls in [tx.find("a:lstStyle", NS)] + [l for _, l in inh]:
        if ls is None:
            continue
        for e in ls:
            m = re.match(r"lvl(\d)pPr", etree.QName(e).localname)
            if not m:
                continue
            lv = int(m.group(1)) - 1
            if lv not in lst and e.find("a:defRPr", NS) is not None and e.find("a:defRPr", NS).get("sz"):
                lst[lv] = int(e.find("a:defRPr", NS).get("sz")) / 100
            if lv not in lal and e.get("algn"):
                lal[lv] = e.get("algn")
            if lv not in lls and e.find("a:lnSpc/a:spcPct", NS) is not None:
                lls[lv] = int(e.find("a:lnSpc/a:spcPct", NS).get("val")) / 100000
    paras, any_text = [], False
    for p in tx.findall("a:p", NS):
        pPr = p.find("a:pPr", NS)
        lvl = int(pPr.get("lvl", "0")) if pPr is not None else 0
        al = {"ctr": "c", "r": "r", "just": "j"}.get(pPr.get("algn") if pPr is not None and pPr.get("algn") else lal.get(lvl), "l")
        # placeholder defaults are those of the 13.33in HSG template, kept in U on any slide size
        base = lst.get(lvl) or ((18.7, 14, 12, 12, 12)[min(lvl, 4)] * PT / ctx.pt if placeholder else ctx.defsz.get(lvl, 18))
        runs = []
        for r in p:
            n = etree.QName(r).localname
            if n == "br":
                runs.append({"br": 1})
                continue
            if n not in ("r", "fld"):
                continue
            t = r.findtext("a:t", default="", namespaces=NS)
            if n == "fld" and r.get("type", "").startswith("slidenum"):
                continue
            rPr = r.find("a:rPr", NS)
            sz = (int(rPr.get("sz")) / 100 if rPr is not None and rPr.get("sz") else base) * scale
            run = {"t": t, "sz": round(sz * ctx.pt, 3)}
            if rPr is not None:
                if rPr.get("b") == "1":
                    run["b"] = 1
                if rPr.get("i") == "1":
                    run["i"] = 1
                if rPr.get("u") not in (None, "none"):
                    run["u"] = 1
                if rPr.get("strike") not in (None, "noStrike"):
                    run["st"] = 1
                if rPr.get("baseline") and int(rPr.get("baseline")) != 0:
                    run["bl"] = 1 if int(rPr.get("baseline")) > 0 else -1
                c = clr(ctx, rPr.find("a:solidFill", NS))
                if c:
                    run["c"] = c[0]
                hl = rPr.find("a:hlinkClick", NS)
                rel = getattr(ctx, "part", None) and hl is not None and ctx.part.rels.get(hl.get("{%s}id" % R))
                if rel and rel.is_external:                  # a link, coloured and underlined as PowerPoint shows it
                    run["href"] = rel.target_ref
                    run["c"] = "#" + ctx.theme.get("hlink", "0563C1")   # PowerPoint ignores the run colour on a link
                    run["u"] = 1
                lat = rPr.find("a:latin", NS)
                if lat is not None:
                    face = lat.get("typeface", "")
                    if re.search(r"Courier|Consolas|Mono|Menlo", face):
                        run["f"] = "mono"
                    elif re.search(r"Palatino|Times|Georgia|Cambria", face):
                        run["f"] = "serif"
                    elif face and not face.startswith("+") and not face.startswith("Gill Sans"):
                        run["ff"] = face                 # another typeface (Arial, Roboto Flex...): used where installed
            if "c" not in run and default_col:
                run["c"] = default_col
            if t.strip():
                any_text = True
            runs.append(run)
        any_text_p = any(r.get("t", "").strip() for r in runs)
        if not runs:                                  # an empty paragraph keeps its height
            e = p.find("a:endParaRPr", NS)
            sz = (int(e.get("sz")) / 100 if e is not None and e.get("sz") else base) * scale
            runs.append({"t": "", "sz": round(sz * ctx.pt, 3)})
        para = {"al": al, "r": runs}
        # bullets: the template marks sub-points with a dash; level 1 is plain text
        bu = None
        if pPr is not None and pPr.find("a:buNone", NS) is not None:
            bu = None
        elif not any_text_p:
            bu = None
        elif pPr is not None and pPr.find("a:buChar", NS) is not None:
            bu = pPr.find("a:buChar", NS).get("char")
        elif pPr is not None and pPr.find("a:buAutoNum", NS) is not None:
            bu = "1."
        elif placeholder and lvl >= 1 and any_text_p:
            bu = "-"
        if pPr is not None and pPr.get("marL"):
            para["pl"] = round(int(pPr.get("marL")) * ctx.k, 3)
        elif placeholder and lvl >= 1:
            para["pl"] = round(2.2 + 1.8 * (lvl - 1), 2)
        if pPr is not None and pPr.get("indent"):
            para["in"] = round(int(pPr.get("indent")) * ctx.k, 3)
        elif bu:
            para["in"] = -1.8
        if bu:
            para["bu"] = bu
            para.setdefault("pl", 2.2)
        if lls.get(lvl):
            para["ls"] = lls[lvl]
        if pPr is not None:
            sp = pPr.find("a:lnSpc/a:spcPct", NS)
            if sp is not None:
                para["ls"] = int(sp.get("val")) / 100000
            for key, tag in (("sb", "spcBef"), ("sa", "spcAft")):
                e = pPr.find(f"a:{tag}/a:spcPts", NS)
                if e is not None:
                    para[key] = round(int(e.get("val")) / 100 * ctx.pt * scale, 3)
                e = pPr.find(f"a:{tag}/a:spcPct", NS)
                if e is not None:
                    para[key] = round(int(e.get("val")) / 100000 * next((r["sz"] for r in runs if "sz" in r), 1.9) * 1.2, 3)
        paras.append(para)
    if not any_text:
        return None
    out = {"p": paras}
    if bodies:
        out["a"] = {"t": "t", "ctr": "m", "b": "b"}.get(battr("anchor", "t"), "t")
        ins = [int(battr(k, d)) * ctx.k for k, d in (("lIns", 91440), ("tIns", 45720), ("rIns", 91440), ("bIns", 45720))]
        out["i"] = [round(v, 3) for v in ins]
        if battr("wrap") == "none":
            out["nw"] = 1
        if battr("vert") in ("vert", "vert270"):
            out["v"] = battr("vert")
    else:
        out["a"] = "t"
    return out


# ---- geometry

def custom_path(ctx, geom, w, h):
    """SVG path (in U) of an a:custGeom."""
    paths = []
    for path in geom.findall("a:pathLst/a:path", NS):
        pw, ph = int(path.get("w", "0")), int(path.get("h", "0"))
        sx = w / pw if pw else 1
        sy = h / ph if ph else 1
        if not pw:
            sx = ctx.k
        if not ph:
            sy = ctx.k
        d, cur = [], (0, 0)
        def pt(e):
            return int(e.get("x")) * sx, int(e.get("y")) * sy
        for c in path:
            n = etree.QName(c).localname
            if n == "moveTo":
                cur = pt(c[0]); d.append("M%.3f %.3f" % cur)
            elif n == "lnTo":
                cur = pt(c[0]); d.append("L%.3f %.3f" % cur)
            elif n == "cubicBezTo":
                ps = [pt(e) for e in c]; cur = ps[2]
                d.append("C" + " ".join("%.3f %.3f" % q for q in ps))
            elif n == "quadBezTo":
                ps = [pt(e) for e in c]; cur = ps[1]
                d.append("Q" + " ".join("%.3f %.3f" % q for q in ps))
            elif n == "arcTo":
                wr, hr = int(c.get("wR")) * sx, int(c.get("hR")) * sy
                st, sw = int(c.get("stAng")) / 60000, int(c.get("swAng")) / 60000
                cx, cy = cur[0] - wr * math.cos(math.radians(st)), cur[1] - hr * math.sin(math.radians(st))
                en = math.radians(st + sw)
                cur = (cx + wr * math.cos(en), cy + hr * math.sin(en))
                d.append("A%.3f %.3f 0 %d %d %.3f %.3f" % (wr, hr, abs(sw) > 180, sw > 0, *cur))
            elif n == "close":
                d.append("Z")
        paths.append({"d": " ".join(d), "fill": path.get("fill") != "none", "stroke": path.get("stroke") != "0"})
    return paths


KNOWN = {"rect", "roundRect", "ellipse", "line", "straightConnector1", "rightArrow", "leftArrow", "downArrow", "upArrow",
         "leftRightArrow", "wedgeEllipseCallout", "cloudCallout", "wedgeRectCallout", "wedgeRoundRectCallout",
         "flowChartAlternateProcess", "foldedCorner", "snip1Rect", "can", "plus", "triangle", "cloud", "hexagon",
         "leftBrace", "mathEqual", "curvedConnector2", "curvedConnector3", "bentConnector2", "bentConnector3",
         "arc", "curvedRightArrow", "flowChartProcess", "flowChartTerminator", "diamond", "octagon", "parallelogram"}


def adjusts(el):
    av = el.find("p:spPr/a:prstGeom/a:avLst", NS)
    out = {}
    if av is not None:
        for g in av:
            m = re.match(r"val (-?\d+)", g.get("fmla", ""))
            if m:
                out[g.get("name")] = int(m.group(1))
    return out


# ---- animations: a click is a step; effects before the first click are visible from the start

def _delay(ctn):
    c = ctn.find("p:stCondLst/p:cond", NS) if ctn is not None else None
    v = c.get("delay", "0") if c is not None else "0"
    return int(v) if v.lstrip("-").isdigit() else 0


def _path_points(path, hu):
    """Points (in U, relative to the shape) of a motion path such as 'M 0 0 C 0.1 0.2 0.3 0.2 0.4 0 E', whose
    coordinates are fractions of the slide width and height. Curves are sampled, so the renderer only moves along lines."""
    toks = re.findall(r"[A-Za-z]|-?\d*\.?\d+(?:[eE]-?\d+)?", path)
    pts, cur, start, i, cmd = [], (0.0, 0.0), (0.0, 0.0), 0, "M"
    while i < len(toks):
        t = toks[i]
        if t.isalpha():
            cmd = t.upper()
            i += 1
            if cmd == "Z":
                cur = start
                pts.append(cur)
            continue
        if cmd in ("M", "L"):
            cur = (float(toks[i]), float(toks[i + 1]))
            i += 2
            if cmd == "M":
                start = cur
            pts.append(cur)
        elif cmd == "C":
            p1, p2, p3 = [(float(toks[i + j]), float(toks[i + j + 1])) for j in (0, 2, 4)]
            i += 6
            for n in range(1, 17):
                u = n / 16
                pts.append(tuple((1 - u) ** 3 * a + 3 * (1 - u) ** 2 * u * b + 3 * (1 - u) * u * u * c + u ** 3 * d
                                 for a, b, c, d in zip(cur, p1, p2, p3)))
            cur = p3
        else:
            i += 1
    return [[round(x * 100, 3), round(y * hu, 3)] for x, y in pts] or [[0, 0]]


def animation_steps(slide, hu=56.25):
    """(step of each shape, step where it disappears, {shape: {paragraph index: step}}, {shape: [effects]}); a click
    is a step. A key is the shape id, or (shape id, modelId) for one element of a SmartArt diagram. Each effect is
    {"k": "in"|"out"|"path"|"rot"|"cmd", "s": step, "d": ms from the click, "dur": ms, ...}; hu is the slide height in U."""
    steps, exits, paras, anims = {}, {}, {}, {}
    timing = slide._element.find("p:timing", NS)
    if timing is None:
        return steps, exits, paras, anims
    seq = None
    for c in timing.iter("{%s}cTn" % P):
        if c.get("nodeType") == "mainSeq":
            seq = c
            break
    if seq is None:
        return steps, exits, paras, anims
    step = 0
    for par in seq.findall("p:childTnLst/p:par", NS):
        ctn = par.find("p:cTn", NS)
        cond = ctn.find("p:stCondLst/p:cond", NS)
        if cond is not None and cond.get("delay") == "indefinite":
            step += 1
        for eff in ctn.iter("{%s}cTn" % P):
            cls = eff.get("presetClass")
            if cls is None:
                continue
            # start of the effect: the delays of its group (with / after previous) and of the effect itself
            d, up = _delay(eff) + _delay(ctn), eff.getparent()
            while up is not None and up is not ctn:
                if etree.QName(up).localname == "cTn":
                    d += _delay(up)
                up = up.getparent()
            tgt0 = next(eff.iter("{%s}spTgt" % P), None)
            if tgt0 is not None and tgt0.find("p:txEl", NS) is None:
                dg = tgt0.find("p:graphicEl/a:dgm", NS)
                key = (int(tgt0.get("spid")), dg.get("id")) if dg is not None else int(tgt0.get("spid"))
                bh = [(b, b.find("p:cBhvr/p:cTn", NS)) for b in eff.find("p:childTnLst", NS) if b.find("p:cBhvr/p:cTn", NS) is not None]
                dur = max([_delay(c) + (int(c.get("dur")) if (c.get("dur") or "").isdigit() else 0) for _, c in bh] + [1])
                rec = {"s": step, "d": d, "dur": dur}
                if cls in ("entr", "exit"):
                    fx = [b.get("filter") for b, _ in bh if etree.QName(b).localname == "animEffect"]
                    rec.update(k="in" if cls == "entr" else "out", fx=fx[0] if fx else "appear")
                    if fx and fx[0] != "fade":
                        warn(f"the {fx[0]} effect is played as a fade")
                    if any(etree.QName(b).localname in ("anim", "animMotion", "animScale", "animRot") for b, _ in bh):
                        warn(f"an entrance or exit effect (preset {eff.get('presetID')}) that moves or scales is played as a fade")
                elif cls == "path":
                    m = next((b for b, _ in bh if etree.QName(b).localname == "animMotion"), None)
                    rec.update(k="path", pts=_path_points(m.get("path", ""), hu))
                elif cls == "emph" and any(etree.QName(b).localname == "animRot" for b, _ in bh):
                    b = next(b for b, _ in bh if etree.QName(b).localname == "animRot")
                    rec.update(k="rot", by=round(int(b.get("by", "0")) / 60000, 2))
                elif cls == "mediacall":
                    c = next((b.get("cmd") for b in eff.iter("{%s}cmd" % P)), "")
                    rec.update(k="cmd", c="pause" if "pause" in c.lower() and "toggle" not in c.lower() else "toggle" if "toggle" in c.lower() else "play")
                else:
                    warn(f"a {cls} effect (preset {eff.get('presetID')}) is not copied")
                    rec = None
                if rec:
                    anims.setdefault(key, []).append(rec)
            if cls not in ("entr", "exit"):
                continue
            for tgt in eff.iter("{%s}spTgt" % P):
                sid = int(tgt.get("spid"))
                rg = tgt.find("p:txEl/p:pRg", NS)
                dgm = tgt.find("p:graphicEl/a:dgm", NS)
                if dgm is not None:                              # one element of a SmartArt diagram
                    (steps if cls == "entr" else exits).setdefault((sid, dgm.get("id")), step)
                elif rg is not None and cls == "entr":
                    for k in range(int(rg.get("st")), int(rg.get("end")) + 1):
                        paras.setdefault(sid, {}).setdefault(k, step)
                elif cls == "entr":
                    steps.setdefault(sid, step)
                else:
                    exits.setdefault(sid, step)
    return steps, exits, paras, anims


# ---- the walk

def box_of(ctx, sh, tf, s, e):
    """Position of a shape in U (tf: the offset and scale of the groups it is in), with its steps."""
    ox, oy, sx, sy = tf
    b = {"x": round((ox + sh.left * sx) * ctx.k, 3), "y": round((oy + sh.top * sy) * ctx.k, 3),
         "w": round(sh.width * sx * ctx.k, 3), "h": round(sh.height * sy * ctx.k, 3)}
    if s:
        b["s"] = s
    if e is not None:
        b["e"] = e
    return b


def title_marked(p):
    """The text of a title paragraph with **bold** marks for bold runs."""
    out = ""
    for e in p._p:
        n = etree.QName(e).localname
        if n == "br":
            out += "\n"
        elif n in ("r", "fld"):
            t = e.findtext("a:t", default="", namespaces=NS)
            rPr = e.find("a:rPr", NS)
            out += f"**{t}**" if rPr is not None and rPr.get("b") == "1" and t.strip() else t
    return out.replace("****", "").replace("** **", " ").strip()


class Walker:
    def __init__(self, ctx, slide, num, imgdir, imgrel, steps, exits, paras, anims=None):
        self.ctx, self.slide, self.num = ctx, slide, num
        self.imgdir, self.imgrel = imgdir, imgrel
        self.steps, self.exits, self.paras, self.anims = steps, exits, paras, anims or {}
        self.items, self.title, self.npics, self.title_obj = [], "", 0, False

    def walk(self, shapes, tf=(0, 0, 1, 1), inh=(None, None, None)):
        for sh in shapes:
            el = sh._element
            if el.xpath("(.//p:cNvPr)[1]/@hidden") == ["1"]:       # hidden in PowerPoint
                continue
            sid = int(el.xpath(".//p:cNvPr/@id")[0])
            s = self.steps.get(sid, inh[0])
            e = self.exits.get(sid, inh[1])
            an = self.anims.get(sid, inh[2])     # ponytail: a group's spin turns each child about its own centre
            ph = el.find(".//p:nvPr/p:ph", NS)
            if ph is not None and ph.get("type") in SKIP_PH:
                if ph.get("type") in ("title", "ctrTitle") and sh.has_text_frame:
                    self.title = "\n".join(title_marked(p) for p in sh.text_frame.paragraphs).strip()
                    if el.find(".//a:r/a:rPr[@sz]", NS) is not None:     # mixed sizes: keep it as drawn, no headline
                        self.title_obj = True
                        self.shape(sh, el, box_of(self.ctx, sh, tf, s, e), True)
                continue
            if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
                x = el.find("p:grpSpPr/a:xfrm", NS)
                if x is not None and x.get("rot") not in (None, "0"):
                    warn("a rotated group was drawn unrotated")
                off, ext = x.find("a:off", NS), x.find("a:ext", NS)
                co, ce = x.find("a:chOff", NS), x.find("a:chExt", NS)
                kx = int(ext.get("cx")) / int(ce.get("cx")) if int(ce.get("cx")) else 1
                ky = int(ext.get("cy")) / int(ce.get("cy")) if int(ce.get("cy")) else 1
                ox, oy, sx, sy = tf
                ntf = (ox + (int(off.get("x")) - int(co.get("x")) * kx) * sx, oy + (int(off.get("y")) - int(co.get("y")) * ky) * sy, kx * sx, ky * sy)
                self.walk(sh.shapes, ntf, (s, e, an))
                continue
            ox, oy, sx, sy = tf
            x, y, w, h = ox + sh.left * sx, oy + sh.top * sy, sh.width * sx, sh.height * sy
            box = {"x": round(x * self.ctx.k, 3), "y": round(y * self.ctx.k, 3), "w": round(w * self.ctx.k, 3), "h": round(h * self.ctx.k, 3)}
            xf = el.find(".//a:xfrm", NS)
            if xf is not None:
                if xf.get("rot"):
                    box["rot"] = round(int(xf.get("rot")) / 60000, 2)
                if xf.get("flipH") == "1":
                    box["fx"] = True
                if xf.get("flipV") == "1":
                    box["fy"] = True
            if s:
                box["s"] = s
            if e is not None:
                box["e"] = e
            if an:
                box["an"] = an
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE or etree.QName(el).localname == "pic":
                self.picture(sh, el, box)
            elif getattr(sh, "has_table", False) and sh.has_table:
                self.table(sh, box)
            elif etree.QName(el).localname in ("sp", "cxnSp"):
                self.shape(sh, el, box, ph is not None)
            elif el.find(".//dgm:relIds", {"dgm": DGM}) is not None:
                self.diagram(sh, el, box, sid, s, e, an)
            else:
                warn(f"an object of type {etree.QName(el).localname} was skipped")

    def shape(self, sh, el, box, placeholder):
        ctx = self.ctx
        it = {"k": "shape", **box}
        spPr = el.find("p:spPr", NS)
        prst = spPr.find("a:prstGeom", NS) if spPr is not None else None
        cust = spPr.find("a:custGeom", NS) if spPr is not None else None
        if etree.QName(el).localname == "cxnSp" and prst is None:
            prst = None
        g = prst.get("prst") if prst is not None else None
        if cust is not None:
            it["g"] = "path"
            it["paths"] = custom_path(ctx, cust, box["w"], box["h"])
        elif g:
            it["g"] = g
            if g not in KNOWN:
                warn(f"shape type {g} was drawn as a rectangle")
            adj = adjusts(el)
            if adj:
                it["adj"] = adj
        elif placeholder or el.find("p:txBody", NS) is not None:
            it["g"] = "rect"
        else:
            return
        f = fill_of(ctx, el)
        if f:
            it["fill"] = f[0]
            if f[1] < 1:
                it["fa"] = round(f[1], 2)
        ln = line_of(ctx, el)
        if ln:
            it["ln"], it["lw"] = ln[0], round(ln[1], 3)
            for key, v in (("dash", ln[2]), ("head", ln[3]), ("tail", ln[4])):
                if v:
                    it[key] = v
        col = default_color(ctx, el)
        if placeholder and not col:
            col = self.layout_color(el)
        tx = text_of(ctx, el, el.find("p:txBody", NS), placeholder, col, self.inherit(el) if placeholder else ())
        if tx:
            it["tx"] = tx
            sid = int(el.xpath(".//p:cNvPr/@id")[0])
            for k, st in self.paras.get(sid, {}).items():       # a text built paragraph by paragraph
                if k < len(tx["p"]) and st:
                    tx["p"][k]["s"] = st
        if "fill" not in it and "ln" not in it and "tx" not in it:
            return
        self.items.append(it)

    def diagram(self, sh, el, box, sid, s, e, an=None):
        """A SmartArt diagram, from the drawing PowerPoint keeps of it (ppt/diagrams/drawingN.xml): its shapes are
        ordinary shapes placed relative to the frame. An animation of one diagram element names it by the modelId of
        its drawn shape or of the data node the shape presents."""
        ids = el.find(".//dgm:relIds", {"dgm": DGM})
        data = etree.fromstring(sh.part.related_part(ids.get("{%s}dm" % R)).blob)
        # PowerPoint colours diagram text by the colour definition of the shape's style label, not by the fontRef the
        # saved drawing carries (which can be stale, e.g. white text on a white box)
        try:
            colors = etree.fromstring(sh.part.related_part(ids.get("{%s}cs" % R)).blob)
        except Exception:
            colors = None
        txfill = {}
        if colors is not None:
            for sl in colors.iter("{%s}styleLbl" % DGM):
                c = sl.find("{%s}txFillClrLst" % DGM)
                if c is not None and len(c):
                    txfill[sl.get("name")] = c[0]
        lbl = {pt.get("modelId"): pt.find("{%s}prSet" % DGM).get("presStyleLbl")
               for pt in data.iter("{%s}pt" % DGM) if pt.find("{%s}prSet" % DGM) is not None}
        ext = data.find(".//{%s}dataModelExt" % DSP)
        if ext is None:
            warn("a SmartArt diagram without a saved drawing was skipped (open and save the file in PowerPoint)")
            return
        drawing = etree.fromstring(sh.part.related_part(ext.get("relId")).blob)
        assoc = {pt.get("modelId"): pt.find("{%s}prSet" % DGM).get("presAssocID")
                 for pt in data.iter("{%s}pt" % DGM) if pt.get("type") == "pres" and pt.find("{%s}prSet" % DGM) is not None}
        k = self.ctx.k
        for sp in drawing.iter("{%s}sp" % DSP):
            sp = etree.fromstring(etree.tostring(sp))
            for x in sp.iter():
                if isinstance(x.tag, str) and x.tag.startswith("{%s}" % DSP):
                    x.tag = "{%s}%s" % (P, etree.QName(x).localname)
            mid = sp.get("modelId")
            fr = sp.find("p:style/a:fontRef", NS)
            if fr is not None and lbl.get(mid) in txfill:
                for c in list(fr):
                    fr.remove(c)
                fr.append(etree.fromstring(etree.tostring(txfill[lbl[mid]])))
            st = self.steps.get((sid, mid), self.steps.get((sid, assoc.get(mid)), s))
            ex = self.exits.get((sid, mid), self.exits.get((sid, assoc.get(mid)), e))
            a = self.anims.get((sid, mid), self.anims.get((sid, assoc.get(mid)), an))

            def place(xf):
                off, ext_ = xf.find("a:off", NS), xf.find("a:ext", NS)
                b = {"x": round(box["x"] + int(off.get("x")) * k, 3), "y": round(box["y"] + int(off.get("y")) * k, 3),
                     "w": round(int(ext_.get("cx")) * k, 3), "h": round(int(ext_.get("cy")) * k, 3)}
                if xf.get("rot"):
                    b["rot"] = round(int(xf.get("rot")) / 60000, 2)
                if xf.get("flipH") == "1":
                    b["fx"] = True
                if xf.get("flipV") == "1":
                    b["fy"] = True
                if st:
                    b["s"] = st
                if ex is not None:
                    b["e"] = ex
                if a:
                    b["an"] = a
                return b
            xf = sp.find("p:spPr/a:xfrm", NS)
            if xf is None:
                continue
            txx = sp.find("p:txXfrm", NS)
            body = sp.find("p:txBody", NS)
            if txx is not None and body is not None:            # the text has its own box inside the shape
                sp.remove(body)
                self.shape(None, sp, place(xf), False)
                tx = text_of(self.ctx, sp, body, False, default_color(self.ctx, sp))
                if tx:
                    self.items.append({"k": "shape", "g": "rect", **place(txx), "tx": tx})
            else:
                self.shape(None, sp, place(xf), False)

    def inherit(self, el):
        """(bodyPr, lstStyle) of what a placeholder inherits, nearest first: the layout's placeholder (same idx, or a
        title for a title), the master's placeholder of the same type, and the master's title or body text style."""
        ph = el.find(".//p:nvPr/p:ph", NS)
        typ, idx = ph.get("type", "obj"), ph.get("idx")
        title = typ in ("title", "ctrTitle")
        mtyp = "title" if title else "body" if typ in ("obj", "body", "subTitle") else typ

        def find(root, master):
            for sp in root.iter("{%s}sp" % P):
                q = sp.find("p:nvSpPr/p:nvPr/p:ph", NS)
                if q is None:
                    continue
                qt = q.get("type", "obj")
                if master and (qt == mtyp or (title and qt in ("title", "ctrTitle"))):
                    return sp
                if not master and ((title and qt in ("title", "ctrTitle")) or (not title and idx is not None and q.get("idx") == idx)):
                    return sp
            return None
        out = []
        for part, master in ((self.slide.slide_layout, False), (self.slide.slide_layout.slide_master, True)):
            sp = find(part._element, master)
            if sp is not None:
                out.append((sp.find("p:txBody/a:bodyPr", NS), sp.find("p:txBody/a:lstStyle", NS)))
        ts = self.slide.slide_layout.slide_master._element.find("p:txStyles", NS)
        if ts is not None:
            out.append((None, ts.find("p:titleStyle" if title else "p:bodyStyle" if mtyp == "body" else "p:otherStyle", NS)))
        return out

    def layout_color(self, el):
        """Text colour that the layout sets for a placeholder, or None."""
        ph = el.find(".//p:nvPr/p:ph", NS)
        idx = ph.get("idx")
        for lp in self.slide.slide_layout.placeholders:
            if str(lp.placeholder_format.idx) == str(idx or 0) or (idx is None and lp.placeholder_format.type == ph.get("type")):
                c = lp._element.find("p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr/a:solidFill", NS)
                col = clr(self.ctx, c.getparent().find("a:solidFill", NS).getparent()) if False else None
                sf = lp._element.find("p:txBody/a:lstStyle/a:lvl1pPr/a:defRPr", NS)
                if sf is not None and sf.find("a:solidFill", NS) is not None:
                    col = clr(self.ctx, sf.find("a:solidFill", NS))
                    return col[0] if col else None
        return None

    def picture(self, sh, el, box):
        has_video = el.find(".//a:videoFile", NS) is not None
        svg = el.find(".//a:blip//{http://schemas.microsoft.com/office/drawing/2016/SVG/main}svgBlip", NS)
        try:
            if svg is not None:                              # PowerPoint draws the SVG; the bitmap is only its fallback
                part = sh.part.related_part(svg.get("{%s}embed" % R))
                img = type("Svg", (), {"blob": part.blob, "ext": "svg"})()
            else:                                            # a video is a Movie in python-pptx, without .image: read its poster
                part = sh.part.related_part(el.find("p:blipFill/a:blip", NS).get("{%s}embed" % R))
                img = type("Img", (), {"blob": part.blob, "ext": part.partname.rsplit(".", 1)[-1]})()
        except Exception:
            if not has_video:
                warn("a picture without image data was skipped")
                return
            img = None                                       # a video without a poster: show a black frame
        import hashlib
        if img is None:
            from PIL import Image
            (self.imgdir / "poster-black.png").exists() or Image.new("RGB", (16, 9), "black").save(self.imgdir / "poster-black.png")
            self.npics += 1
            it = {"k": "pic", "src": f"{self.imgrel}/poster-black.png", "alt": "Video", **box}
            self.attach_video(sh, el, it)
            self.items.append(it)
            return
        ext = img.ext.lower()
        self.npics += 1
        name = f"{hashlib.sha1(img.blob).hexdigest()[:10]}.{ext}"       # same picture, same file
        if ext not in ("png", "jpg", "jpeg", "gif", "svg", "webp"):
            png = self.convert(img.blob, ext, name)
            if not png:
                warn(f"picture {name} is a {ext} file, which a browser cannot show; it was skipped (convert it to PNG)")
                return
            name = png
        elif not (self.imgdir / name).exists():
            (self.imgdir / name).write_bytes(img.blob)
        it = {"k": "pic", "src": f"{self.imgrel}/{name}", "alt": el.xpath(".//p:cNvPr/@descr")[0].strip() if el.xpath(".//p:cNvPr/@descr") else "", **box}
        sr = el.find("p:blipFill/a:srcRect", NS)
        if sr is not None and any(sr.get(k) for k in "ltrb"):
            it["crop"] = [int(sr.get(k, "0")) / 100000 for k in "ltrb"]
        if el.find("p:spPr/a:effectLst/a:outerShdw", NS) is not None:
            it["shadow"] = True
        self.attach_video(sh, el, it)
        self.items.append(it)

    def attach_video(self, sh, el, it):
        vf = el.find(".//a:videoFile", NS)
        if vf is None:
            return
        rel = sh.part.rels.get(vf.get("{%s}link" % R))
        if rel is not None and rel.is_external and rel.target_ref.startswith("http"):
            it["embed"] = rel.target_ref                     # an online video (YouTube): played in place on click
            return
        try:
            part = sh.part.related_part(vf.get("{%s}link" % R))
            vext = part.partname.rsplit(".", 1)[-1].lower()
            vname = f"{hashlib.sha1(part.blob).hexdigest()[:10]}.{vext}"
            mdir = self.imgdir.parent / "media"
            mdir.mkdir(exist_ok=True)
            if not (mdir / vname).exists():
                (mdir / vname).write_bytes(part.blob)
            it["video"] = f"media/{vname}"
        except Exception:
            warn("a video could not be extracted; only its poster is shown")

    def convert(self, blob, ext, name):
        """Try to make a PNG of formats a browser cannot show (JPEG XR .wdp, .emf) with Pillow."""
        try:
            from PIL import Image
            import io
            im = Image.open(io.BytesIO(blob))
            out = name.rsplit(".", 1)[0] + ".png"
            im.save(self.imgdir / out)
            return out
        except Exception:
            return None

    def table(self, sh, box):
        ctx, t = self.ctx, sh.table
        tbl = sh._element.find(".//a:tbl", NS)
        pr = tbl.find("a:tblPr", NS)
        first_row = pr is not None and pr.get("firstRow") == "1"
        rows = []
        for ri, row in enumerate(tbl.findall("a:tr", NS)):
            cells = []
            for tc in row.findall("a:tc", NS):
                c = {}
                tcPr = tc.find("a:tcPr", NS)
                f = clr(ctx, tcPr.find("a:solidFill", NS)) if tcPr is not None else None
                if f:
                    c["fill"] = f[0]
                elif tcPr is not None and tcPr.find("a:noFill", NS) is not None:
                    pass
                else:
                    c["fill"] = "head" if first_row and ri == 0 else "a" if ri % 2 else "b"
                c["tx"] = text_of(ctx, tc, tc.find("a:txBody", NS), False, None)
                if tc.get("gridSpan"):
                    c["cs"] = int(tc.get("gridSpan"))
                if tc.get("hMerge") == "1":
                    c["hm"] = 1
                cells.append(c)
            rows.append({"h": round(int(row.get("h")) * ctx.k, 3), "cells": cells})
        self.items.append({"k": "table", **box, "cols": [round(c.width * ctx.k, 3) for c in t.columns], "rows": rows})


def collapse(items):
    return items


def dumps(deck):
    """JSON with one canvas item per line, so a deck stays readable and diffs stay small."""
    big = []

    def mark(o):
        if isinstance(o, dict) and o.get("layout") == "canvas" and "items" in o:
            lines = ",\n".join("    " + json.dumps(i, ensure_ascii=False, separators=(", ", ": ")) for i in o["items"])
            big.append(lines)
            o = {**o, "items": f"@@CANVAS{len(big) - 1}@@"}
        return o
    d = {**deck, "slides": [mark(s) for s in deck["slides"]]}
    out = json.dumps(d, ensure_ascii=False, indent=1)
    for i, lines in enumerate(big):
        out = out.replace(f'"@@CANVAS{i}@@"', "[\n" + lines + "\n   ]")
    return out + "\n"


def convert_slide(pptx_path, num, imgdir, imgrel):
    prs = Presentation(pptx_path)
    slide = prs.slides[num - 1]
    ctx = Ctx(prs, slide)
    ctx.part = slide.part
    steps, exits, paras, anims = animation_steps(slide, prs.slide_height * ctx.k)
    w = Walker(ctx, slide, num, imgdir, imgrel, {}, {}, {}, {})
    if slide._element.get("showMasterSp") != "0":       # texts the layout draws on every slide, e.g. a date
        w.walk([sh for sh in slide.slide_layout.shapes if not sh.is_placeholder and sh.has_text_frame and sh.text_frame.text.strip()])
    w.steps, w.exits, w.paras, w.anims = steps, exits, paras, anims
    w.walk(slide.shapes)
    # every step must change something: renumber the clicks that show, hide, move, turn or play at least one thing
    # (a click with only a sound effect would be a dead key press)
    used = sorted({i[k] for i in w.items for k in ("s", "e") if i.get(k)} | {p["s"] for i in w.items for p in (i.get("tx") or {}).get("p", []) if p.get("s")}
                  | {a["s"] for i in w.items for a in i.get("an", []) if a["s"]})
    remap = {v: n for n, v in enumerate(used, 1)}
    for i in w.items:
        for k in ("s", "e"):
            if i.get(k):
                i[k] = remap[i[k]]
        for p in (i.get("tx") or {}).get("p", []):
            if p.get("s"):
                p["s"] = remap[p["s"]]
        if i.get("an"):
            i["an"] = [{**a, "s": remap.get(a["s"], 0)} for a in i["an"]]
    s = {"layout": "canvas", "title": "" if w.title_obj else w.title, "items": w.items}
    mx = max(remap.values(), default=0)
    if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
        s["notes"] = slide.notes_slide.notes_text_frame.text.strip()
    tr = transition_of(slide)
    if tr:
        s["transition"] = tr
    if title_anchor(slide) == "t":
        s["ta"] = "t"                                    # the headline hangs from the top of its box
    for el in (slide._element, slide.slide_layout._element, slide.slide_layout.slide_master._element):
        bg = el.find("p:cSld/p:bg", NS)                  # the first background set: slide, layout or master
        if bg is not None:
            c = clr(ctx, bg.find("p:bgPr/a:solidFill", NS)) or clr(ctx, bg.find("p:bgRef", NS))
            if c and c[0] != "#ffffff":
                s["bg"] = c[0]
            break
    return s, mx


def title_anchor(slide):
    """Vertical anchor of the title placeholder (t, ctr, b), from the slide, its layout or its master."""
    for src in (slide, slide.slide_layout, slide.slide_layout.slide_master):
        for sh in src.placeholders:
            ph = sh._element.find(".//p:nvPr/p:ph", NS)
            b = sh._element.find("p:txBody/a:bodyPr", NS)
            if ph is not None and ph.get("type") in ("title", "ctrTitle") and b is not None and b.get("anchor"):
                return b.get("anchor")
    return None


def transition_of(slide):
    """The transition into a slide: {"k": "fade"|"morph", "dur": ms} or None. PowerPoint writes newer transitions
    (Morph, durations in ms) in an mc:Choice before an older fallback; the first p:transition is the one it plays."""
    t = next(slide._element.iter("{%s}transition" % P), None)
    if t is None:
        return None
    kind = next((etree.QName(c).localname for c in t if isinstance(c.tag, str) and etree.QName(c).localname != "extLst"), None)
    if kind is None:
        return None
    dur = t.get("{http://schemas.microsoft.com/office/powerpoint/2010/main}dur")
    dur = int(dur) if dur else {"fast": 500, "med": 750, "slow": 1000}.get(t.get("spd", "fast"), 500)
    if kind not in ("fade", "morph"):
        warn(f"the {kind} transition is played as a fade")
        kind = "fade"
    return {"k": kind, "dur": dur}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--slide", type=int, required=True, help="slide number in the .pptx (1-based)")
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--at", type=int, help="replace this slide of the deck (1-based)")
    g.add_argument("--insert", type=int, help="insert before this slide of the deck (1-based)")
    ap.add_argument("--title", help="set the headline instead of the one of the .pptx")
    a = ap.parse_args()
    path = Path(a.deck).resolve()
    deck = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {
        "title": "Canvas test", "lang": "en", "slides": []}
    imgdir = path.parent / "img"
    imgdir.mkdir(exist_ok=True)
    s, mx = convert_slide(a.pptx, a.slide, imgdir, "img")
    if a.title is not None:
        s["title"] = a.title
    if a.at:
        deck["slides"][a.at - 1] = s
    elif a.insert:
        deck["slides"].insert(a.insert - 1, s)
    else:
        deck["slides"].append(s)
    path.write_text(dumps(deck), encoding="utf-8")
    pics = [i for i in s["items"] if i["k"] == "pic" and not i["alt"]]
    print(f"slide {a.slide} -> {len(s['items'])} items, {mx} click steps, title: {s['title']!r}")
    if pics:
        print(f"  {len(pics)} pictures have no alt text yet: " + ", ".join(i["src"] for i in pics))
    for m in warnings:
        print("  warning: " + m)


if __name__ == "__main__":
    main()
