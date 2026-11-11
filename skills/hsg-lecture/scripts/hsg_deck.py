#!/usr/bin/env python3
"""Shared parts of the HSG lecture renderer: the deck format, its checks and the colours.

A deck is one JSON file. render_html.py reads it through load().

deck.json
{
  "title": "Computing on Bits",              # name of the deck (browser tab)
  "lang": "en",
  "footer": "Computersysteme | ...",         # optional quiet line at the foot of every content slide
  "slides": [ {"layout": "...", ...}, ... ]
}

Every slide: "layout", "title", and optionally "lead" (one line under the headline), "callout" (a box beside
the headline), "source" (small credit line), "notes" (speaker notes), "time" (pacing cue shown in the overview).
Speaker notes can also live in speaker-notes.md next to the deck (see load_notes), which overrides "notes".
The fields of each layout are listed in LAYOUTS below and explained in SKILL.md.

  python hsg_deck.py deck.json            # check a deck without rendering
"""
import json, re, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "assets"

# Colours of the HSG slide template (theme "HSG" of the Computersysteme decks)
GREEN, DARK_GREEN = "#00802F", "#0A5F2D"
SALMON, TEAL, SAND, YELLOW = "#EB6969", "#73A5AF", "#E1D7C3", "#FFF04B"
ROW_A, ROW_B = "#CBD8CD", "#E7EDE8"   # table rows: green at 20% and 10% on white
BOX_STYLES = ("green", "outline", "teal")

COMMON = {"layout", "title", "lead", "callout", "source", "notes", "time", "id", "refs", "refs_at", "decor",
          "transition"}   # transition: copied from PowerPoint
NO_LEAD = {"title", "section", "task", "quote"}
# layout: (required fields, optional fields)
LAYOUTS = {
    "title":     ({"title"}, {"kicker", "subtitle", "presenters", "affiliation", "meta", "date", "venue", "picture", "style"}),
    "section":   ({"title"}, {"kicker"}),
    "menu":      ({"title", "items"}, {"current"}),
    "text":      ({"title"}, {"paragraphs", "points", "defs", "groups", "tiles", "after", "aside"}),
    "study":     ({"title", "groups"}, {"aside"}),
    "questions": ({"title"}, {"heading", "points", "after"}),
    "table":     ({"title", "columns", "rows"}, {"after"}),
    "code":      ({"title", "code"}, {"paragraphs", "points", "style", "caption"}),
    "quiz":      ({"title", "items"}, {"aside"}),
    "task":      ({"title", "stamp"}, {"paragraphs", "url", "label"}),
    "image":     ({"title", "image", "alt"}, {"caption", "shadow"}),
    "gallery":   ({"title", "items"}, {"after", "style"}),
    "people":    ({"title", "items"}, set()),
    "quote":     ({"title"}, {"attribution", "picture", "points"}),
    "syllabus":  ({"title", "items"}, {"current", "break_after", "assignments"}),
    "canvas":    ({"title"}, {"items", "bg", "ta"}),       # a slide copied from PowerPoint by pptx_canvas.py; the title may be ""
}
ITEM_KEYS = {
    "gallery":  ({"image", "alt"}, {"caption", "shadow"}),
    "people":   ({"name"}, {"photo", "alt", "role"}),
    "syllabus": ({"date", "title"}, {"kind", "href"}),
}
ITEM_COUNT = {"gallery": (2, 4), "people": (1, 9), "quiz": (2, 5), "menu": (2, 8), "syllabus": (2, 14)}
PICTURE_KEYS = ({"image", "alt"}, {"caption", "shadow"})
ASIDE_KEYS = {"image", "alt", "caption", "shadow", "box", "code"}
BOX_KEYS = ({"text"}, {"style", "heading"})
WORD_CAP = {"text": 90, "study": 110, "questions": 80, "table": 140, "code": 120, "syllabus": 140,
            "people": 60, "title": 50, "quiz": 70, "gallery": 60}


class DeckError(Exception):
    pass


REF_KINDS = ("recap", "next", "see")
INLINE_REF = re.compile(r"\[([^\]]+)\]\(([\w.-]+#[\w-]+)\)")      # [text](08#slide-id), see SKILL.md


def deck_folders(base):
    """The decks next to this one: folder name -> deck.json path (a deck folder holds deck.json)."""
    return {p.parent.name: p for p in sorted(base.parent.glob("*/deck.json"))}


_cache = {}


def load_deck(path):
    path = Path(path)
    if path not in _cache or _cache[path][0] != path.stat().st_mtime:
        _cache[path] = (path.stat().st_mtime, json.loads(path.read_text(encoding="utf-8")))
    return _cache[path][1]


def resolve(to, base):
    """Resolve "07b#uniform-interface" (deck folder name or its prefix, # slide id) for a deck in folder base.
    Returns (folder name, slide id, target slide, target deck) or raises DeckError."""
    m = re.fullmatch(r"([\w.-]+)#([\w-]+)", to or "")
    if not m:
        raise DeckError(f"'{to}' is not a reference like 08-interaction-styles#slide-id")
    key, sid = m.groups()
    folders = deck_folders(base)
    hits = [n for n in folders if n == key or n.startswith(key)]
    if len(hits) != 1:
        raise DeckError(f"'{to}': deck '{key}' not found next to this deck" if not hits else f"'{to}': deck '{key}' is ambiguous ({', '.join(hits)})")
    deck = load_deck(folders[hits[0]])
    for s in deck["slides"]:
        if s.get("id") == sid:
            return hits[0], sid, s, deck
    raise DeckError(f"'{to}': deck {hits[0]} has no slide with id '{sid}'")


def _keys(obj, required, optional, where, errors):
    if not isinstance(obj, dict):
        errors.append(f"{where}: must be an object")
        return False
    for k in sorted(required - obj.keys()):
        errors.append(f"{where}: '{k}' is missing")
    for k in sorted(obj.keys() - required - optional):
        errors.append(f"{where}: unknown field '{k}' (the design has no free fields; see SKILL.md)")
    return True


def words(s):
    """All visible text of a slide part as one string, for counting. Code is not counted."""
    if isinstance(s, str):
        return s
    if isinstance(s, list):
        return " ".join(words(x) for x in s)
    if isinstance(s, dict):
        return " ".join(words(v) for k, v in s.items()
                        if k not in ("layout", "notes", "image", "photo", "alt", "url", "time", "style", "code",
                                     "shadow", "kind", "current", "break_after", "from", "to", "id", "refs", "refs_at", "href",
                                     "decor", "icon"))
    return ""


def points_of(pts):
    """Points are strings or {"text", "sub": [strings]}; returns (top-level count, list of sub lists)."""
    return len(pts or []), [p.get("sub", []) for p in pts or [] if isinstance(p, dict)]


def check(deck, base):
    """Return (errors, warnings). Errors stop the render; warnings are design and structure tells to fix or justify."""
    errors, warnings = [], []
    if not _keys(deck, {"title", "slides"}, {"lang", "footer", "theme", "ref_label", "autoadvance", "pdf", "pdf_version"}, "deck", errors):
        return errors, warnings
    if "autoadvance" in deck and not (isinstance(deck["autoadvance"], (int, float)) and not isinstance(deck["autoadvance"], bool)
                                      and deck["autoadvance"] > 0):
        errors.append("deck: 'autoadvance' must be a number of seconds greater than 0")
    if "pdf" in deck:
        if not (isinstance(deck["pdf"], str) and deck["pdf"].lower().endswith(".pdf")):
            errors.append("deck: 'pdf' must be the path of a .pdf file, relative to html/index.html")
        elif not (base / "html" / deck["pdf"]).is_file():
            warnings.append(f"deck: the PDF {deck['pdf']} does not exist yet (print html/index.html to make it)")
    if "pdf_version" in deck and not ("pdf" in deck and isinstance(deck["pdf_version"], str) and deck["pdf_version"].strip()):
        errors.append("deck: 'pdf_version' must be a text (e.g. a date) and needs 'pdf'")
    if "theme" in deck and not (isinstance(deck["theme"], str) and (base / deck["theme"]).is_file()):
        errors.append(f"deck: theme file not found: {deck['theme']}")
    slides = deck.get("slides") or []
    if not isinstance(slides, list) or not slides:
        errors.append("deck: 'slides' must be a list with at least one slide")
        return errors, warnings

    def need_file(p, where):
        if not isinstance(p, str) or not (base / p).is_file():
            errors.append(f"{where}: file not found: {p}")

    def need_picture(pic, where):
        if _keys(pic, *PICTURE_KEYS, where, errors):
            need_file(pic["image"], where)
            if "alt" in pic and not str(pic["alt"]).strip():
                errors.append(f"{where}: the picture needs 'alt' (what is visible in it)")

    def need_box(b, where):
        if _keys(b, *BOX_KEYS, where, errors) and b.get("style", "green") not in BOX_STYLES:
            errors.append(f"{where}: 'style' must be one of {', '.join(BOX_STYLES)}")

    def need_points(pts, where, most=6):
        if pts is None:
            return
        if not isinstance(pts, list):
            errors.append(f"{where}: 'points' must be a list")
            return
        for i, p in enumerate(pts, 1):
            if isinstance(p, dict):
                _keys(p, {"text"}, {"sub"}, f"{where}, point {i}", errors)
                if len(p.get("sub", [])) > 3:
                    warnings.append(f"{where}, point {i}: more than 3 sub-points")
            elif not isinstance(p, str):
                errors.append(f"{where}, point {i}: a point is a string or an object with 'text' and 'sub'")
        if len(pts) > most:
            warnings.append(f"{where}: {len(pts)} points; {most} is the limit a room can follow, split the slide")

    def need_code(code, where):
        if not isinstance(code, str) or not code.strip():
            errors.append(f"{where}: 'code' must be a non-empty string (use \\n for line breaks)")
            return
        lines = code.split("\n")
        if len(lines) > 18:
            warnings.append(f"{where}: {len(lines)} lines of code; show at most 18 and put the rest in the course materials")
        if max(len(l) for l in lines) > 64:
            warnings.append(f"{where}: a code line has more than 64 characters; it will be set smaller or wrap")

    run = None
    ids = {}

    def need_ref(to, where):
        try:
            resolve(to, base)
        except DeckError as e:
            errors.append(f"{where}: {e}")

    for n, s in enumerate(slides, 1):
        w = f"slide {n}"
        lay = s.get("layout") if isinstance(s, dict) else None
        if lay not in LAYOUTS:
            errors.append(f"{w}: 'layout' must be one of {', '.join(LAYOUTS)}")
            continue
        req, opt = LAYOUTS[lay]
        _keys(s, req | {"layout"}, opt | COMMON, w, errors)
        num = lambda v: isinstance(v, (int, float)) and not isinstance(v, bool)
        for i, it in enumerate(s.get("decor", []) if isinstance(s.get("decor", []), list) else [None], 1):
            if _keys(it, {"x", "y", "size"}, {"emoji", "rot"}, f"{w}, decor {i}", errors) and not all(
                    num(it[k]) for k in ("x", "y", "size") if k in it):
                errors.append(f"{w}, decor {i}: 'x', 'y' (percent of the slide) and 'size' (U) are numbers")
        if lay == "gallery" and s.get("style", "photos") not in ("photos", "qr"):
            errors.append(f"{w}: gallery 'style' must be photos or qr")
        if lay in NO_LEAD and "lead" in s:
            errors.append(f"{w}: '{lay}' has no 'lead'")
        if lay in ("title", "section") and "callout" in s:
            errors.append(f"{w}: '{lay}' has no 'callout'")
        title = s.get("title", "")
        if "id" in s:
            if not re.fullmatch(r"[a-z0-9][a-z0-9-]*", str(s["id"])):
                errors.append(f"{w}: 'id' must be lowercase letters, digits and hyphens")
            elif s["id"] in ids:
                errors.append(f"{w}: 'id' {s['id']} is already used by slide {ids[s['id']]}")
            ids[s["id"]] = n
        if "refs" in s:
            refs = s["refs"]
            if not isinstance(refs, list) or len(refs) > 2:
                errors.append(f"{w}: 'refs' is a list of at most 2 references (more do not fit above the slide)")
            else:
                for i, r in enumerate(refs, 1):
                    if _keys(r, {"to"}, {"kind", "label"}, f"{w}, ref {i}", errors):
                        if r.get("kind", "see") not in REF_KINDS:
                            errors.append(f"{w}, ref {i}: 'kind' must be one of {', '.join(REF_KINDS)}")
                        try:
                            _, _, tslide, _ = resolve(r["to"], base)
                            if not r.get("label") and not str(tslide.get("title", "")).strip():
                                errors.append(f"{w}, ref {i}: the target slide has no title; give the reference a 'label'")
                        except DeckError as e:
                            errors.append(f"{w}, ref {i}: {e}")
        for m_ in INLINE_REF.finditer(words(s)):
            need_ref(m_.group(2), f"{w}, link [{m_.group(1)}]")
        if lay == "syllabus":
            for i, it in enumerate(s.get("items", []), 1):
                if isinstance(it, dict) and it.get("href"):
                    need_ref(it["href"], f"{w}, item {i}")
        if lay == "canvas":
            for i, it in enumerate(s.get("items", []), 1):
                for p in (it.get("tx") or {}).get("p", []) if isinstance(it, dict) else []:
                    for r in p.get("r", []):
                        h = r.get("href")
                        if h and not str(h).startswith(("http://", "https://")):
                            need_ref(h, f"{w}, item {i}, link '{r.get('t', '')[:30]}'")
        if lay == "canvas":
            for i, it in enumerate(s.get("items", []), 1):
                if not isinstance(it, dict) or it.get("k") not in ("shape", "pic", "table"):
                    errors.append(f"{w}, item {i}: 'k' must be shape, pic or table")
                elif it["k"] == "pic":
                    need_file(it.get("src"), f"{w}, item {i}")
                    if not str(it.get("alt", "")).strip():
                        warnings.append(f"{w}, item {i}: the picture has no alt text (describe what is visible, in the deck)")
            if not isinstance(title, str):
                errors.append(f"{w}: 'title' must be a string")
            run = lay
            continue
        if not isinstance(title, str) or not title.strip():
            errors.append(f"{w}: 'title' is empty")
            continue
        title = title.replace("**", "")
        letters = [c for c in title if c.isalpha()]
        if len(letters) > 6 and all(c.isupper() for c in letters):
            errors.append(f"{w}: the title is in capitals; write it in sentence or title case")
        limit = 140 if lay == "quote" else 80
        if len(title) > limit:
            warnings.append(f"{w}: headline has {len(title)} characters (aim for under {limit})")
        if " — " in words(s) or " – " in words(s):
            warnings.append(f"{w}: text contains a dash used as punctuation; rewrite the sentence")
        if s.get("callout") is not None:
            c = s["callout"]
            if isinstance(c, dict) and "wide" in c:
                if not isinstance(c["wide"], bool):
                    errors.append(f"{w}, callout: 'wide' must be true or false")
                c = {k: v for k, v in c.items() if k != "wide"}
            need_box(c, f"{w}, callout")

        items = s.get("items")
        if lay in ITEM_COUNT:
            if not isinstance(items, list):
                errors.append(f"{w}: 'items' must be a list")
                continue
            lo, hi = ITEM_COUNT[lay]
            if not lo <= len(items) <= hi:
                errors.append(f"{w}: '{lay}' takes {lo} to {hi} items, got {len(items)}; split the slide or choose another layout")
        if lay in ITEM_KEYS:
            for i, it in enumerate(items, 1):
                wi = f"{w}, item {i}"
                if not _keys(it, *ITEM_KEYS[lay], wi, errors):
                    continue
                if lay == "gallery":
                    need_picture(it, wi)
                if lay == "people" and it.get("photo"):
                    need_file(it["photo"], wi)
                if lay == "syllabus" and it.get("kind", "lecture") not in ("lecture", "exercise"):
                    errors.append(f"{wi}: 'kind' must be lecture or exercise")
        if lay in ("menu", "quiz"):
            for i, it in enumerate(items, 1):
                if not isinstance(it, str):
                    errors.append(f"{w}, item {i}: must be a string")

        if lay in ("menu", "syllabus") and s.get("current") is not None:
            cur = s["current"] if isinstance(s["current"], list) else [s["current"]]
            if not all(isinstance(c, int) and 1 <= c <= len(items) for c in cur):
                errors.append(f"{w}: 'current' must be the number of an item (1 to {len(items)}), or a list of them")
        if lay == "syllabus":
            ba = s.get("break_after")
            if ba is not None and not (isinstance(ba, int) and 1 <= ba < len(items)):
                errors.append(f"{w}: 'break_after' must be the number of an item before the break")
            for i, a in enumerate(s.get("assignments", []), 1):
                if _keys(a, {"title", "from", "to"}, set(), f"{w}, assignment {i}", errors):
                    if not (isinstance(a["from"], int) and isinstance(a["to"], int) and 1 <= a["from"] <= a["to"] <= len(items)):
                        errors.append(f"{w}, assignment {i}: 'from' and 'to' are item numbers, from <= to")
            if len(s.get("assignments", [])) > 4:
                warnings.append(f"{w}: more than 4 assignments do not fit beside the plan")

        if lay in ("text", "study"):
            if lay == "text" and not any(s.get(k) for k in ("paragraphs", "points", "defs", "groups", "tiles")):
                errors.append(f"{w}: 'text' needs 'paragraphs', 'points', 'defs', 'groups' or 'tiles'")
            tiles = s.get("tiles", [])
            if not isinstance(tiles, list) or len(tiles) not in (0, 3):
                errors.append(f"{w}: 'tiles' takes exactly 3 tiles")
            else:
                for i, t in enumerate(tiles, 1):
                    _keys(t, {"icon", "text"}, set(), f"{w}, tile {i}", errors)
            need_points(s.get("points"), w)
            for i, g in enumerate(s.get("groups", []), 1):
                if _keys(g, {"heading"}, {"points", "paragraphs"}, f"{w}, group {i}", errors):
                    need_points(g.get("points"), f"{w}, group {i}", most=4)
            if len(s.get("groups", [])) > 3:
                warnings.append(f"{w}: more than 3 groups; split the slide")
            for i, d in enumerate(s.get("defs", []), 1):
                if not (isinstance(d, list) and len(d) == 2 and all(isinstance(x, str) for x in d)):
                    errors.append(f"{w}, def {i}: must be a pair [term, detail]")
            if len(s.get("defs", [])) > 8:
                warnings.append(f"{w}: more than 8 definitions; split the slide or use a table")
            a = s.get("aside")
            if a is not None and _keys(a, set(), ASIDE_KEYS, f"{w}, aside", errors):
                kinds = [k for k in ("image", "box", "code") if a.get(k)]
                if len(kinds) != 1:
                    errors.append(f"{w}, aside: give exactly one of 'image', 'box' or 'code'")
                if a.get("image"):
                    need_picture({k: v for k, v in a.items() if k in ("image", "alt", "caption", "shadow")}, f"{w}, aside")
                if a.get("box"):
                    need_box(a["box"], f"{w}, aside box")
                if a.get("code"):
                    need_code(a["code"], f"{w}, aside")
        if lay == "quiz" and s.get("aside") is not None:
            need_picture(s["aside"], f"{w}, aside")
        if lay == "questions":
            need_points(s.get("points"), w, most=5)
        if lay == "title" and s.get("style", "gear") not in ("gear", "photo", "side", "slanted"):
            errors.append(f"{w}: 'style' must be gear, photo, side or slanted")
        if lay == "code":
            need_code(s.get("code"), w)
            need_points(s.get("points"), w, most=5)
            if s.get("style", "outline") not in ("outline", "teal"):
                errors.append(f"{w}: 'style' must be outline or teal")
        if lay == "table":
            cols, rows = s.get("columns"), s.get("rows")
            if not (isinstance(cols, list) and 2 <= len(cols) <= 6):
                errors.append(f"{w}: 'columns' must be a list of 2 to 6 headings")
            elif not (isinstance(rows, list) and rows and all(isinstance(r, list) and len(r) == len(cols) for r in rows)):
                errors.append(f"{w}: 'rows' must be a list of rows with {len(cols)} cells each")
            elif len(rows) > 10:
                warnings.append(f"{w}: {len(rows)} rows; at most 10 fit, split the table")
        if lay == "task" and s.get("url") and not str(s["url"]).startswith(("http://", "https://")):
            errors.append(f"{w}: 'url' must start with http:// or https://")
        if lay == "image":
            need_picture({k: s[k] for k in ("image", "alt", "caption", "shadow") if k in s}, w)
        if lay in ("title", "quote") and s.get("picture"):
            need_picture(s["picture"], f"{w}, picture")
        if lay == "quote":
            need_points(s.get("points"), w, most=4)

        count = len(words(s).split())
        cap = WORD_CAP.get(lay, 80)
        if count > cap:
            warnings.append(f"{w}: {count} words on the slide (aim for under {cap}); move detail to the notes or split")
        # neighbouring slides should not look alike; slides of headline and running text may repeat
        if run == lay and lay not in ("text", "image", "code", "table"):
            warnings.append(f"{w}: same '{lay}' layout as the slide before, so the two look alike; "
                            "give one of them another layout or move a slide between them")
        run = lay
    return errors, warnings


def load_notes(deck, md):
    """Speaker notes from speaker-notes.md next to the deck: the lines under a heading "### N Title" become the
    notes of slide N and replace its "notes" in deck.json. Any other heading ends a slide's notes."""
    if not md.is_file():
        return []
    warnings, n, body, found = [], None, [], {}
    for line in md.read_text(encoding="utf-8").splitlines() + ["#"]:
        if line.startswith("#"):
            if n is not None:
                found[n] = "\n".join(body).strip()
            m = re.match(r"###\s+(\d+)\b", line)
            n, body = (int(m.group(1)) if m else None), []
        elif n is not None:
            body.append(line)
    for n, text in found.items():
        if 1 <= n <= len(deck["slides"]):
            deck["slides"][n - 1]["notes"] = text
        else:
            warnings.append(f"{md.name}: there is no slide {n}")
    return warnings


def load(path):
    """Read and check a deck. Returns (deck, base folder, warnings); exits with the list of errors."""
    path = Path(path).resolve()
    try:
        deck = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as e:
        sys.exit(f"Could not read {path}: {e}")
    notes_warnings = load_notes(deck, path.parent / "speaker-notes.md")
    errors, warnings = check(deck, path.parent)
    warnings += notes_warnings
    if errors:
        sys.exit("The deck has errors:\n  " + "\n  ".join(errors))
    return deck, path.parent, warnings


def host(url):
    return re.sub(r"^https?://(www\.)?", "", url).split("/")[0]


def report(warnings):
    if warnings:
        print("Check these before you report (fix them, or say why they stay):")
        for w in warnings:
            print("  - " + w)


if __name__ == "__main__":
    if len(sys.argv) == 2:
        d, _, warn = load(sys.argv[1])
        print(f"{len(d['slides'])} slides: no errors")
        report(warn)
    else:
        sys.exit(__doc__)
