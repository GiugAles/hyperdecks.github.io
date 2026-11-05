#!/usr/bin/env python3
"""Shared parts of the FuSo presentation renderers: the deck format, its checks, the colours and the icons.

A deck is one JSON file. render_html.py and render_pptx.py both read it through load().

deck.json
{
  "title": "AI + Social Media",          # name of the deck (browser tab, file properties)
  "mode": "light",                       # light | dark, for the whole deck
  "lang": "en",
  "slides": [ {"layout": "...", ...}, ... ]
}

Every slide: "layout", "title", and optionally "highlight" (a phrase of the title, shown in green),
"bg" (plain | grey | green), "lead" (one sentence under the headline), "source" (small credit line),
"notes" (speaker notes), "time" (pacing cue shown in the overview).
The fields of each layout are listed in LAYOUTS below and explained in SKILL.md.

  python fuso_deck.py deck.json            # check a deck without rendering
  python fuso_deck.py --find-symbol lock   # search the Phosphor icon set
"""
import base64, json, os, re, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "assets"

BLACK, GRAY, GREEN = "#111111", "#f4f4f4", "#00802F"
GREEN_ON_DARK = "#1fa64f"   # FuSo green lightened for text on the dark background (4.5:1 on #121212)
ON_GREEN_LINE = "#59ac78"   # white at 35% on FuSo green, as a solid colour for PowerPoint

# Colours per mode and slide background. Light follows the LinkedIn figures, dark the dark mode of
# futuresociety.ch. A deck never shows both: no dark panel in a light deck and no light panel in a dark one.
# "inv" is the second panel of a comparison (green), "band" the base under columns.
# hl_bg None = the highlight is green text; otherwise it is a marker block (green is too weak as text there).
_LIGHT = dict(fg=BLACK, body="#333333", quiet="#666666", divider="#dddddd", line=BLACK, hl_fg=GREEN, hl_bg=None,
              panel_fg=BLACK, panel_body="#333333", panel_quiet="#666666", panel_line=None, icon="#666666",
              inv=GREEN, inv_fg="#ffffff", inv_line="transparent", band=GREEN, band_fg="#ffffff", foot_fg=BLACK, wm="rgba(0,128,47,.18)")
_DARK = dict(fg="#ededed", body="#dadada", quiet="#bdbdbd", divider="#2e2e2e", line="#ededed", hl_fg=GREEN_ON_DARK, hl_bg=None,
             panel_fg="#ededed", panel_body="#dadada", panel_quiet="#bdbdbd", panel_line=None, icon="#ededed",
             inv=GREEN, inv_fg="#ffffff", inv_line="transparent", band=GREEN, band_fg="#ffffff", foot_fg="#ededed", wm="rgba(0,128,47,.55)")
_ON_GREEN = dict(fg="#ffffff", body="#ffffff", quiet="#ffffff", divider=ON_GREEN_LINE, line="#ffffff",
                 hl_fg=GREEN, hl_bg="#ffffff", wm="rgba(255,255,255,.24)",
                 inv="transparent", inv_fg="#ffffff", inv_line="#ffffff")   # on green the second panel is an outline
MODES = {
    "light": {
        "plain": dict(_LIGHT, bg="#ffffff", panel=GRAY, disc=GREEN, icon="#ffffff", foot=GRAY),
        "grey":  dict(_LIGHT, bg=GRAY, panel="#ffffff", disc=GREEN, icon="#ffffff", foot="#ffffff"),
        "green": dict(_LIGHT, **_ON_GREEN, bg=GREEN, panel="#ffffff", disc="#ffffff", icon=GREEN,
                      band="#ffffff", band_fg=GREEN, foot=GRAY),
    },
    "dark": {
        "plain": dict(_DARK, bg="#121212", panel="#1a1a1a", disc=GREEN, icon="#ffffff", foot="#1a1a1a"),
        "grey":  dict(_DARK, bg="#1a1a1a", panel="#121212", disc=GREEN, icon="#ffffff", foot="#121212"),
        "green": dict(_DARK, **_ON_GREEN, bg=GREEN, panel="#121212", panel_line=None, disc="#ffffff", icon=GREEN,
                      band="#121212", band_fg="#ededed", foot="#121212"),
    },
}

# Type scale. One unit (u) is 1% of the slide width: 9.6 pt in PowerPoint, 1vw on a 16:9 screen.
SIZES = dict(h1=6.0, hs=3.6, h2=3.1, lead=1.35, h3=1.6, text=1.3, body=1.15, small=0.95, tiny=0.8)

COMMON = {"layout", "title", "highlight", "bg", "callout", "lead", "source", "notes", "time"}
# layout: (required fields, optional fields)
LAYOUTS = {
    "title":     ({"title"}, {"subtitle", "presenters", "meta", "qr", "banner", "mark", "kicker"}),   # "mark": a picture behind the end of the first title line; "kicker": a pill above the title
    "section":   ({"title"}, {"text"}),
    "statement": ({"title"}, {"text", "attribution"}),
    "text":      ({"title"}, {"paragraphs", "points", "after", "contacts", "aside"}),   # "after": one quiet line under the points
    "columns":   ({"title", "items"}, {"base"}),
    "steps":     ({"title", "items"}, {"numbered"}),
    "compare":   ({"title", "left", "right"}, {"conclusion"}),
    "fuse":      ({"title", "left", "right"}, {"conclusion"}),   # two fields as overlapping circles, FuSo in the overlap
    "agenda":    ({"title", "items"}, set()),
    "people":    ({"title", "items"}, {"questions"}),
    "links":     ({"title", "items"}, set()),
    "flow":      ({"title", "start", "end", "items"}, set()),
    "matrix":    ({"title", "columns", "rows"}, set()),
    "logos":     ({"title", "items"}, set()),
    "image":     ({"title", "image", "alt"}, {"caption"}),
    "closing":   ({"title"}, {"text", "contacts", "qr"}),
}
MARKS = {"brown-bag": "brown_bag.png"}   # the paper bag of the Brown Bag Lunch Talk graphics
NO_LEAD = {"title", "section", "statement", "closing"}
RINGS = NO_LEAD - {"title"}            # layouts that carry the watermark rings of the LinkedIn figures
ITEM_KEYS = {
    "columns": ({"heading"}, {"icon", "text"}),
    "steps":   ({"heading"}, {"text", "icon"}),
    "agenda":  ({"title"}, {"when", "who", "sub"}),
    "people":  ({"name"}, {"photo", "alt", "role", "points", "tag"}),
    "links":   ({"heading", "url"}, {"source", "date", "text", "label"}),
}
ITEM_COUNT = {"columns": (2, 6), "steps": (2, 5), "people": (1, 12), "links": (2, 3), "agenda": (2, 8)}
ASIDE_KEYS = {"side", "title", "photo", "image", "alt", "caption", "icon", "heading", "text", "facts", "qr", "logos", "badge", "label", "note"}
LOGO_KEYS = ({"src", "alt"}, {"src_dark", "name"})
SIDE_KEYS = ({"heading"}, {"points", "text", "icon"})


class DeckError(Exception):
    pass


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
    """All text of a slide part as one string, for counting."""
    if isinstance(s, str):
        return s
    if isinstance(s, list):
        return " ".join(words(x) for x in s)
    if isinstance(s, dict):
        return " ".join(words(v) for k, v in s.items()
                        if k not in ("layout", "bg", "notes", "icon", "image", "photo", "alt", "url", "time", "qr", "highlight"))
    return ""


NUMBER_LABEL = re.compile(r"^\s*(\d{1,2}[\s./:)\-]|[A-Za-z]\)|step\s+\d)", re.I)


def check(deck, base):
    """Return (errors, warnings). Errors stop the render; warnings are design and structure tells to fix or justify."""
    errors, warnings = [], []
    if not _keys(deck, {"title", "mode", "slides"}, {"lang", "partners"}, "deck", errors):
        return errors, warnings
    if deck.get("mode") not in MODES:
        errors.append("deck: 'mode' must be light or dark (one mode for the whole deck)")
    slides = deck.get("slides") or []
    if not isinstance(slides, list) or not slides:
        errors.append("deck: 'slides' must be a list with at least one slide")
        return errors, warnings

    def need_file(p, where):
        if not isinstance(p, str) or not (base / p).is_file():
            errors.append(f"{where}: file not found: {p}")

    def need_qr(q, where):
        if _keys(q, {"url"}, {"label", "heading", "text"}, where, errors) and not str(q.get("url", "")).startswith(("http://", "https://")):
            errors.append(f"{where}: 'url' must start with http:// or https://")

    def need_logos(lst, where):
        for i, lg in enumerate(lst or [], 1):
            if _keys(lg, *LOGO_KEYS, f"{where}, logo {i}", errors):
                need_file(lg["src"], f"{where}, logo {i}")
                if lg.get("src_dark"):
                    need_file(lg["src_dark"], f"{where}, logo {i}")

    need_logos(deck.get("partners"), "deck partners")
    run, highlights = [], 0
    for n, s in enumerate(slides, 1):
        w = f"slide {n}"
        lay = s.get("layout") if isinstance(s, dict) else None
        if lay not in LAYOUTS:
            errors.append(f"{w}: 'layout' must be one of {', '.join(LAYOUTS)}")
            continue
        req, opt = LAYOUTS[lay]
        _keys(s, req | {"layout"}, opt | COMMON, w, errors)
        if lay in NO_LEAD and "lead" in s:
            errors.append(f"{w}: '{lay}' has no 'lead'; use its 'text' or 'subtitle'")
        if s.get("bg", "plain") not in ("plain", "grey", "green"):
            errors.append(f"{w}: 'bg' must be plain, grey or green")
        title = s.get("title", "")
        if not isinstance(title, str) or not title.strip():
            errors.append(f"{w}: 'title' is empty")
            continue
        hl = s.get("highlight")
        if hl:
            highlights += 1
            if hl not in title:
                errors.append(f"{w}: 'highlight' must be a phrase of the title, written the same way")
            elif len(hl.split()) > max(4, len(title.split()) // 2):
                warnings.append(f"{w}: the highlight covers much of the headline; mark only the phrase that carries the message")
        letters = [c for c in title if c.isalpha()]
        if len(letters) > 6 and all(c.isupper() for c in letters):
            errors.append(f"{w}: the title is in capitals; write it as a sentence")
        if lay not in ("title", "statement", "closing") and title.rstrip().endswith("."):
            warnings.append(f"{w}: headline ends with a full stop")
        limit = 110 if lay == "statement" else 75
        if len(title) > limit:
            warnings.append(f"{w}: headline has {len(title)} characters (aim for under {limit})")
        if " — " in words(s) or " – " in words(s):
            warnings.append(f"{w}: text contains a dash used as punctuation; rewrite the sentence")

        items = s.get("items")
        if lay in ITEM_KEYS:
            if not isinstance(items, list):
                errors.append(f"{w}: 'items' must be a list")
                continue
            lo, hi = ITEM_COUNT[lay]
            if not lo <= len(items) <= hi:
                errors.append(f"{w}: '{lay}' takes {lo} to {hi} items, got {len(items)}; split the slide or choose another layout")
            for i, it in enumerate(items, 1):
                wi = f"{w}, item {i}"
                if not _keys(it, *ITEM_KEYS[lay], wi, errors):
                    continue
                head = it.get("heading") or it.get("title") or ""
                if lay != "agenda" and NUMBER_LABEL.match(head):
                    errors.append(f"{wi}: heading starts with a number or step label; only 'steps' with \"numbered\": true "
                                  "shows numbers, and only for a real order")
                letters = [c for c in head if c.isalpha()]
                if len(letters) > 4 and all(c.isupper() for c in letters):
                    errors.append(f"{wi}: heading is in capitals")
                if lay == "people" and it.get("photo"):
                    need_file(it["photo"], wi)
                if lay == "people" and len(it.get("points", [])) > 3:
                    warnings.append(f"{wi}: more than 3 points per person")
                if lay == "links" and not str(it.get("url", "")).startswith(("http://", "https://")):
                    errors.append(f"{wi}: 'url' must start with http:// or https://")
                if lay == "agenda":
                    for j, sub in enumerate(it.get("sub", []), 1):
                        _keys(sub, {"title"}, {"who"}, f"{wi}, sub {j}", errors)
            if lay in ("columns", "steps") and len({bool(it.get("icon")) for it in items if isinstance(it, dict)}) > 1:
                errors.append(f"{w}: give every card an icon or none")
            if lay == "columns" and "base" in s:
                _keys(s["base"], {"text"}, {"label"}, f"{w}, base", errors)
        if lay in ("compare", "fuse"):
            for side in ("left", "right"):
                if _keys(s.get(side), *SIDE_KEYS, f"{w}, {side}", errors) and len(s[side].get("points", [])) > 4:
                    warnings.append(f"{w}, {side}: more than 4 points")
        if lay == "text":
            if not (s.get("paragraphs") or s.get("points") or s.get("contacts")):
                errors.append(f"{w}: 'text' needs 'paragraphs', 'points' or 'contacts'")
            if len(s.get("points", [])) > 5:
                warnings.append(f"{w}: {len(s['points'])} points; five is the limit a room can follow")
            a = s.get("aside")
            if a is not None and _keys(a, set(), ASIDE_KEYS, f"{w}, aside", errors):
                if a.get("image"):
                    need_file(a["image"], f"{w}, aside")
                    if not a.get("alt"):
                        errors.append(f"{w}, aside: the image needs 'alt' (what is visible in it)")
                if a.get("qr"):
                    need_qr(a["qr"], f"{w}, aside qr")
                if a.get("photo") and _keys(a["photo"], {"image", "alt"}, set(), f"{w}, aside photo", errors):
                    need_file(a["photo"]["image"], f"{w}, aside photo")
                need_logos(a.get("logos"), f"{w}, aside")
                if a.get("side", "right") not in ("left", "right"):
                    errors.append(f"{w}, aside: 'side' must be left or right")
                if a.get("qr") and not any(a.get(k) for k in ("text", "facts", "label", "image")):
                    warnings.append(f"{w}, aside: a panel with a QR code is green and the most visible block of the slide; "
                                    "use it only for something that warrants this (an event, a project with its own site), not for a bare link")
                if not any(a.get(k) for k in ("image", "icon", "title", "heading", "text", "facts", "logos")):
                    errors.append(f"{w}, aside: empty")
        if lay == "image":
            need_file(s.get("image"), w)
        if lay == "title" and s.get("banner") and _keys(s["banner"], {"image", "alt"}, {"url"}, f"{w}, banner", errors):
            need_file(s["banner"]["image"], f"{w}, banner")
        if lay == "title" and s.get("mark") not in (None, *MARKS):
            errors.append(f"{w}: 'mark' must be one of {', '.join(MARKS)}")
        if lay == "title" and s.get("qr"):
            warnings.append(f"{w}: QR code on the title slide; keep it only if the user asked for it (the closing slide carries the link)")
        if lay in ("title", "closing") and s.get("qr"):
            need_qr(s["qr"], f"{w}, qr")
        if lay in ("closing", "text"):
            for i, c in enumerate(s.get("contacts", []), 1):
                _keys(c, {"label"}, {"url", "icon"}, f"{w}, contact {i}", errors)

        count = len(words(s).split())
        cap = {"agenda": 130, "people": 120, "links": 90, "title": 60}.get(lay, 70)
        if count > cap:
            warnings.append(f"{w}: {count} words on the slide (aim for under {cap}); move detail to the notes or split")
        # neighbouring slides must not look alike; only slides of headline and running text may repeat
        plain = lay in ("statement", "section") or (lay == "text" and not s.get("aside")
                                                    and not any(isinstance(x, dict) for x in s.get("points", [])))
        if run and run[-1] == lay and not plain:
            warnings.append(f"{w}: same '{lay}' layout as the slide before, so the two look alike; "
                            "give one of them another layout or move a slide between them")
        run = [lay]
    content = [s for s in slides if isinstance(s, dict) and s.get("layout") not in ("title", "section", "closing")]
    if len(content) >= 4 and highlights > len(slides) * 0.6:
        warnings.append(f"deck: {highlights} of {len(slides)} headlines have a green phrase; keep it for the slides "
                        "where one phrase carries the message")
    greens = sum(1 for s in slides if isinstance(s, dict) and s.get("bg") == "green")
    if greens > max(2, len(slides) // 4):
        warnings.append(f"deck: {greens} green slides; green is for the few slides that mark a turn in the talk")
    return errors, warnings


def load(path):
    """Read and check a deck. Returns (deck, base folder, warnings); exits with the list of errors."""
    path = Path(path).resolve()
    try:
        deck = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as e:
        sys.exit(f"Could not read {path}: {e}")
    errors, warnings = check(deck, path.parent)
    if errors:
        sys.exit("The deck has errors:\n  " + "\n  ".join(errors))
    for s in deck["slides"]:
        s.setdefault("bg", "plain")
    return deck, path.parent, warnings


def split_highlight(title, hl):
    """(before, highlight, after); highlight is '' if there is none."""
    if hl and hl in title:
        a, b = title.split(hl, 1)
        return a, hl, b
    return title, "", ""


def host(url):
    return re.sub(r"^https?://(www\.)?", "", url).split("/")[0]


# ---------------------------------------------------------------- icons (Phosphor, MIT)

PHOSPHOR_URL = "https://raw.githubusercontent.com/phosphor-icons/core/main/assets/fill/{name}-fill.svg"
ICON_CACHE = Path(os.environ.get("FUSO_ICON_CACHE", Path.home() / ".cache" / "fuso-icons"))


def fetch_symbol(name):
    """Download one Phosphor icon from GitHub on demand and cache it locally."""
    cached = ICON_CACHE / f"{name}-fill.svg"
    if cached.exists():
        return cached.read_text()
    import urllib.request, urllib.error
    url = PHOSPHOR_URL.format(name=name)
    try:
        with urllib.request.urlopen(url, timeout=20) as r:
            svg = r.read().decode()
    except urllib.error.HTTPError as e:
        sys.exit(f"Icon '{name}' not found ({e.code}). Search with: fuso_deck.py --find-symbol <word>")
    except Exception as e:
        sys.exit(f"Could not download icon '{name}' from {url} ({e}). raw.githubusercontent.com must be reachable.")
    try:
        ICON_CACHE.mkdir(parents=True, exist_ok=True)
        cached.write_text(svg)
    except OSError:
        pass
    return svg


def find_symbols(query):
    cat = json.loads((ASSETS / "phosphor" / "catalog.json").read_text())
    query = [w.lower() for w in query]
    cat = [i for i in cat if not i["name"].endswith("-logo")]  # no brand logos
    hits = [i for i in cat if all(any(w in t for t in [i["name"]] + i["tags"]) for w in query)]
    return [(i["name"], ", ".join(i["tags"][:6])) for i in hits]


def qr_matrix(url):
    import segno
    return segno.make(url, error="m")


def data_uri(p):
    mime = {".svg": "image/svg+xml", ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
            ".webp": "image/webp", ".woff2": "font/woff2"}[Path(p).suffix.lower()]
    return f"data:{mime};base64,{base64.b64encode(Path(p).read_bytes()).decode()}"


def report(warnings):
    if warnings:
        print("Check these before you report (fix them, or say why they stay):")
        for w in warnings:
            print("  - " + w)


if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--find-symbol":
        for n, t in find_symbols(sys.argv[2:]):
            print(f"{n:32s} {t}")
    elif len(sys.argv) == 2:
        d, _, warn = load(sys.argv[1])
        print(f"{len(d['slides'])} slides, mode {d['mode']}: no errors")
        report(warn)
    else:
        sys.exit(__doc__)
