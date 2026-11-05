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
import argparse, colorsys, json, math, re, sys, zipfile
from pathlib import Path
from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"a": A, "p": P, "r": R}
PT = 1 / 9.6          # 1pt in U (the template slide is 13.33in wide, 1U = 9.6pt)
SKIP_PH = {"title", "ctrTitle", "sldNum", "ftr", "dt"}
warnings = []


def warn(msg):
    if msg not in warnings:
        warnings.append(msg)


class Ctx:
    def __init__(self, prs, slide):
        self.k = 100 / prs.slide_width          # U per EMU
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
            elif mn == "tint":
                l = l + (1 - l) * (1 - v)
            elif mn == "shade":
                l *= v
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


def text_of(ctx, el, tx, placeholder=False, default_col=None):
    """The 'tx' object of a shape from its txBody element, or None when there is no text."""
    if tx is None:
        return None
    body = tx.find("a:bodyPr", NS)
    scale = 1.0
    if body is not None and body.find("a:normAutofit", NS) is not None:
        scale = int(body.find("a:normAutofit", NS).get("fontScale", "100000")) / 100000
    lst = {}
    ls = tx.find("a:lstStyle", NS)
    if ls is not None:
        for e in ls:
            m = re.match(r"lvl(\d)pPr", etree.QName(e).localname)
            if m and e.find("a:defRPr", NS) is not None and e.find("a:defRPr", NS).get("sz"):
                lst[int(m.group(1)) - 1] = int(e.find("a:defRPr", NS).get("sz")) / 100
    paras, any_text = [], False
    for p in tx.findall("a:p", NS):
        pPr = p.find("a:pPr", NS)
        lvl = int(pPr.get("lvl", "0")) if pPr is not None else 0
        al = {"ctr": "c", "r": "r", "just": "j"}.get(pPr.get("algn") if pPr is not None else None, "l")
        base = lst.get(lvl) or ((18.7, 14, 12, 12, 12)[min(lvl, 4)] if placeholder else 18)
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
            run = {"t": t, "sz": round(sz * PT, 3)}
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
                lat = rPr.find("a:latin", NS)
                if lat is not None:
                    face = lat.get("typeface", "")
                    if re.search(r"Courier|Consolas|Mono|Menlo", face):
                        run["f"] = "mono"
                    elif re.search(r"Palatino|Times|Georgia|Cambria", face):
                        run["f"] = "serif"
            if "c" not in run and default_col:
                run["c"] = default_col
            if t.strip():
                any_text = True
            runs.append(run)
        any_text_p = any(r.get("t", "").strip() for r in runs)
        if not runs:                                  # an empty paragraph keeps its height
            e = p.find("a:endParaRPr", NS)
            sz = (int(e.get("sz")) / 100 if e is not None and e.get("sz") else base) * scale
            runs.append({"t": "", "sz": round(sz * PT, 3)})
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
        if pPr is not None:
            sp = pPr.find("a:lnSpc/a:spcPct", NS)
            if sp is not None:
                para["ls"] = int(sp.get("val")) / 100000
            for key, tag in (("sb", "spcBef"), ("sa", "spcAft")):
                e = pPr.find(f"a:{tag}/a:spcPts", NS)
                if e is not None:
                    para[key] = round(int(e.get("val")) / 100 * PT * scale, 3)
                e = pPr.find(f"a:{tag}/a:spcPct", NS)
                if e is not None:
                    para[key] = round(int(e.get("val")) / 100000 * runs[0]["sz"] * 1.2, 3)
        paras.append(para)
    if not any_text:
        return None
    out = {"p": paras}
    if body is not None:
        out["a"] = {"t": "t", "ctr": "m", "b": "b"}.get(body.get("anchor", "t"), "t")
        ins = [int(body.get(k, d)) * ctx.k for k, d in (("lIns", 91440), ("tIns", 45720), ("rIns", 91440), ("bIns", 45720))]
        out["i"] = [round(v, 3) for v in ins]
        if body.get("wrap") == "none":
            out["nw"] = 1
        if body.get("vert") in ("vert", "vert270"):
            out["v"] = body.get("vert")
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

def animation_steps(slide):
    steps, exits = {}, {}
    timing = slide._element.find("p:timing", NS)
    if timing is None:
        return steps, exits
    seq = None
    for c in timing.iter("{%s}cTn" % P):
        if c.get("nodeType") == "mainSeq":
            seq = c
            break
    if seq is None:
        return steps, exits
    step = 0
    for par in seq.findall("p:childTnLst/p:par", NS):
        ctn = par.find("p:cTn", NS)
        cond = ctn.find("p:stCondLst/p:cond", NS)
        if cond is not None and cond.get("delay") == "indefinite":
            step += 1
        for eff in ctn.iter("{%s}cTn" % P):
            cls = eff.get("presetClass")
            if cls not in ("entr", "exit"):
                continue
            for tgt in eff.iter("{%s}spTgt" % P):
                sid = int(tgt.get("spid"))
                if cls == "entr":
                    steps.setdefault(sid, step)
                else:
                    exits.setdefault(sid, step)
    return steps, exits


# ---- the walk

class Walker:
    def __init__(self, ctx, slide, num, imgdir, imgrel, steps, exits):
        self.ctx, self.slide, self.num = ctx, slide, num
        self.imgdir, self.imgrel = imgdir, imgrel
        self.steps, self.exits = steps, exits
        self.items, self.title, self.npics = [], "", 0

    def walk(self, shapes, tf=(0, 0, 1, 1), inh=(None, None)):
        for sh in shapes:
            el = sh._element
            if el.xpath("(.//p:cNvPr)[1]/@hidden") == ["1"]:       # hidden in PowerPoint
                continue
            sid = int(el.xpath(".//p:cNvPr/@id")[0])
            s = self.steps.get(sid, inh[0])
            e = self.exits.get(sid, inh[1])
            ph = el.find(".//p:nvPr/p:ph", NS)
            if ph is not None and ph.get("type") in SKIP_PH:
                if ph.get("type") in ("title", "ctrTitle") and sh.has_text_frame:
                    self.title = "\n".join(p.text.strip() for p in sh.text_frame.paragraphs).strip()
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
                self.walk(sh.shapes, ntf, (s, e))
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
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE or etree.QName(el).localname == "pic":
                self.picture(sh, el, box)
            elif getattr(sh, "has_table", False) and sh.has_table:
                self.table(sh, box)
            elif etree.QName(el).localname in ("sp", "cxnSp"):
                self.shape(sh, el, box, ph is not None)
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
        tx = text_of(ctx, el, el.find("p:txBody", NS), placeholder, default_color(ctx, el))
        if tx:
            it["tx"] = tx
        if "fill" not in it and "ln" not in it and "tx" not in it:
            return
        self.items.append(it)

    def picture(self, sh, el, box):
        try:
            img = sh.image
        except Exception:
            warn("a picture without image data was skipped")
            return
        ext = img.ext.lower()
        self.npics += 1
        name = f"s{self.num}-{self.npics}.{ext}"
        if ext not in ("png", "jpg", "jpeg", "gif", "svg", "webp"):
            png = self.convert(img.blob, ext, name)
            if not png:
                warn(f"picture {name} is a {ext} file, which a browser cannot show; it was skipped (convert it to PNG)")
                return
            name = png
        else:
            (self.imgdir / name).write_bytes(img.blob)
        it = {"k": "pic", "src": f"{self.imgrel}/{name}", "alt": el.xpath(".//p:cNvPr/@descr")[0].strip() if el.xpath(".//p:cNvPr/@descr") else "", **box}
        sr = el.find("p:blipFill/a:srcRect", NS)
        if sr is not None and any(sr.get(k) for k in "ltrb"):
            it["crop"] = [int(sr.get(k, "0")) / 100000 for k in "ltrb"]
        if el.find("p:spPr/a:effectLst/a:outerShdw", NS) is not None:
            it["shadow"] = True
        self.items.append(it)

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
    steps, exits = animation_steps(slide)
    w = Walker(ctx, slide, num, imgdir, imgrel, steps, exits)
    w.walk(slide.shapes)
    s = {"layout": "canvas", "title": w.title, "items": w.items}
    mx = max([i.get("s", 0) for i in w.items] + [0])
    if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
        s["notes"] = slide.notes_slide.notes_text_frame.text.strip()
    return s, mx


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
