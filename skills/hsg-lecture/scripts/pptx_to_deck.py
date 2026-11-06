#!/usr/bin/env python3
"""Convert a whole lecture .pptx into a deck.json: title, syllabus and questions slides and slides of plain text
get their normal layouts, every other slide becomes a 'canvas' slide (see pptx_canvas.py). Hidden slides are skipped.

  python pptx_to_deck.py lecture.pptx deck.json --syllabus-from other/deck.json --current 8

Then: write the alt text of every picture (the check lists them), run hsg_deck.py, render with --shots and compare
with the .pdf of the lecture.
"""
import argparse, json, re, sys
from pathlib import Path
from lxml import etree
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parent))
import pptx_canvas as pc

NS = pc.NS
FOOTER = "Advanced Software and Systems Engineering | Institute of Computer Science | University of St.Gallen"
PRESENTERS = ["Simon Mayer", "Alessandro Giugno", "Luka Bekavac"]
META = ["University of St.Gallen", "Master in Computer Science", "", "Advanced Software and Systems Engineering (7,856)"]


def runs_text(p):
    out = ""
    for r in p.findall("a:r", NS):
        t = r.findtext("a:t", default="", namespaces=NS)
        rPr = r.find("a:rPr", NS)
        if t.strip() and rPr is not None and rPr.get("b") == "1" and "**" not in out[-2:]:
            t = f"**{t.strip()}** " if t.endswith(" ") else f"**{t}**"
        out += t
    return out.replace("****", "").strip()


def plain_points(slide):
    """The points of a slide that only has a title and text placeholders, else None."""
    pts, other = [], False
    for sh in slide.shapes:
        el = sh._element
        ph = el.find(".//p:nvPr/p:ph", NS)
        if ph is not None and ph.get("type") in pc.SKIP_PH:
            continue
        if ph is None or etree.QName(el).localname != "sp" or not sh.has_text_frame:
            return None
        if el.find(".//p:txBody//a:buNone", NS) is None and False:
            pass
        for p in el.findall(".//p:txBody/a:p", NS):
            t = runs_text(p)
            if not t:
                continue
            lvl = int(p.find("a:pPr", NS).get("lvl", "0")) if p.find("a:pPr", NS) is not None else 0
            if lvl == 0 or not pts:
                pts.append({"text": t, "sub": []})
            else:
                pts[-1]["sub"].append(t)
    if not pts or len(pts) > 6 or any(len(p["sub"]) > 3 for p in pts):
        return None
    return [p["text"] if not p["sub"] else p for p in [{"text": p["text"], "sub": p["sub"]} for p in pts]] and \
        [(p["text"] if not p["sub"] else {"text": p["text"], "sub": p["sub"]}) for p in pts]


def convert(pptx, out_deck, syllabus_from, current, deck_title=None, kicker=None):
    prs = Presentation(pptx)
    out_deck = Path(out_deck)
    imgdir = out_deck.parent / "img"
    imgdir.mkdir(parents=True, exist_ok=True)
    syl = None
    if syllabus_from:
        syl = next(s for s in json.loads(Path(syllabus_from).read_text(encoding="utf-8"))["slides"] if s["layout"] == "syllabus")
        syl = {**syl, "current": current}
    slides, hidden = [], 0
    for i, sl in enumerate(prs.slides, 1):
        if sl._element.get("show") == "0":
            hidden += 1
            continue
        title = ""
        t = sl.shapes.title
        if t is not None:
            title = "\n".join(p.text.strip() for p in t.text_frame.paragraphs).strip()
        low = title.lower().replace("\n", " ")
        notes = sl.notes_slide.notes_text_frame.text.strip() if sl.has_notes_slide else ""
        if not slides:
            s = {"layout": "title", "title": title.replace(" - ", ": ") or deck_title or "Lecture", "presenters": PRESENTERS, "meta": META}
            if kicker:
                s["kicker"] = kicker
        elif low == "syllabus" and syl:
            s = dict(syl)
        elif low.startswith("any questions"):
            s = {"layout": "questions", "title": title}
        else:
            pts = None if 'presetClass="entr"' in sl._element.xml else plain_points(sl)      # animated slides need the canvas
            if pts and title:
                s = {"layout": "text", "title": title.replace("\n", " "), "points": pts}
            else:
                s, _ = pc.convert_slide(pptx, i, imgdir, "img")
        if notes and "notes" not in s:
            s["notes"] = notes
        slides.append(s)
    deck = {"title": deck_title or slides[0]["title"].replace("\n", " "), "lang": "en", "footer": FOOTER, "slides": slides}
    out_deck.write_text(pc.dumps(deck), encoding="utf-8")
    return deck, hidden


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx")
    ap.add_argument("deck")
    ap.add_argument("--syllabus-from", help="a deck.json whose syllabus slide is copied")
    ap.add_argument("--current", type=int, help="number of this lecture in the syllabus (1 to 12)")
    ap.add_argument("--title")
    ap.add_argument("--kicker")
    a = ap.parse_args()
    deck, hidden = convert(a.pptx, a.deck, a.syllabus_from, a.current, a.title, a.kicker)
    n = len(deck["slides"])
    kinds = {}
    for s in deck["slides"]:
        kinds[s["layout"]] = kinds.get(s["layout"], 0) + 1
    print(f"{n} slides ({hidden} hidden slides skipped): {kinds}")
    for m in pc.warnings:
        print("  warning: " + m)


if __name__ == "__main__":
    main()
