#!/usr/bin/env python3
"""Apply a link plan (links.json) to the decks of a course: stable slide ids, reference pills, inline links, and the
automatic links of syllabus, learning goals and navigators. Run it after every (re-)conversion, then render.

  python apply_links.py asse/links.json          # edits the deck.json files named in the plan, in place
  python apply_links.py asse/links.json --check  # only report slides and texts that could not be found

links.json (paths relative to its own folder; decks are named by the start of their folder name):
{
 "labels":   {"07b": "Lecture 7b"},                          # shown in the pills: "Recap  Lecture 7b: <slide title>"
 "ids":      [["07b", "Uniform Interface#1", "uniform-interface"]],
 "refs":     [["08", "Copied from last week", [["07b#client-server", "recap"]]]],
 "inline":   [["08", "Remote Data Access#1", "Client-Server: Three High-level Options", "07b#client-server"]],
 "syllabus": {"Interaction Styles": "08#start"},             # syllabus bars -> first slide of a deck
 "goals":    {"1": "07b#start", "2": "08#start"},            # 'Lecture 1: ...' lines of the learning goals slides
 "navigators": [["07b", ["Client-Server", "client-server"]]] # boxes with exactly this text, narrower than 20U
}
A finder is a title start ("Uniform Interface#2" = the second slide with that title), "=Title" = exactly this title,
or "~text" = a slide that contains this text; add "#n" for the n-th. Quotes and apostrophes are compared in their plain form.
"""
import argparse, json, re, sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import pptx_canvas as pc

PLAIN = str.maketrans({"’": "'", "‘": "'", "„": '"', "“": '"', "”": '"', " ": " "})
norm = lambda s: re.sub(r"\s+", " ", s.translate(PLAIN)).strip().lower()


def items_of(s):
    """The canvas items of a slide (other layouts, e.g. quiz, have plain strings as items)."""
    return [it for it in s.get("items", []) if isinstance(it, dict)]


def runs_text(s):
    """All (item, paragraph) text of a slide as plain strings."""
    out = []
    for it in items_of(s):
        for p in (it.get("tx") or {}).get("p", []):
            out.append("".join(r.get("t", "") for r in p["r"]))
        if it["k"] == "table" if "k" in it else False:
            for row in it["rows"]:
                for c in row["cells"]:
                    for p in (c.get("tx") or {}).get("p", []):
                        out.append("".join(r.get("t", "") for r in p["r"]))
    for key in ("title", "lead", "after"):
        if isinstance(s.get(key), str):
            out.append(s[key])
    for key in ("points", "paragraphs"):
        for p in s.get(key, []) or []:
            out.append(p if isinstance(p, str) else p.get("text", ""))
    return out


def find(slides, spec):
    m = re.search(r"#(\d+)$", spec)
    nth = int(m.group(1)) if m else 1
    key = norm(spec[:m.start()] if m else spec)
    hits = []
    for s in slides:
        if key.startswith("~"):
            ok = any(key[1:].strip() in norm(t) for t in runs_text(s))
        elif key.startswith("="):
            ok = norm(s.get("title", "").replace("**", "")) == key[1:].strip()
        else:
            ok = norm(s.get("title", "").replace("**", "")).startswith(key)
        if ok:
            hits.append(s)
    return hits[nth - 1] if len(hits) >= nth else None


def split_run(slide, text, href):
    """Make `text` a link: split the paragraph run that contains it (formatting is kept). Returns True when found."""
    key = norm(text)
    for it in slide.get("items", []):
        for p in (it.get("tx") or {}).get("p", []):
            for i, r in enumerate(p["r"]):
                t = r.get("t", "")
                low = t.translate(PLAIN).lower()
                pos = low.find(key)
                if pos < 0:
                    continue
                if r.get("href") == href:
                    return True
                parts = [(t[:pos], None), (t[pos:pos + len(key)], href), (t[pos + len(key):], None)]
                new = []
                for seg, h in parts:
                    if seg:
                        nr = {k: v for k, v in r.items() if k != "href"}
                        nr["t"] = seg
                        if h:
                            nr["href"] = h
                        new.append(nr)
                p["r"][i:i + 1] = new
                return True
    return False


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("plan")
    ap.add_argument("--check", action="store_true", help="do not write, only report")
    a = ap.parse_args()
    plan_path = Path(a.plan).resolve()
    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    root = plan_path.parent
    folders = {p.parent.name: p for p in sorted(root.glob("*/deck.json"))}

    def deck_of(key):
        hits = [n for n in folders if n.startswith(key)]
        if len(hits) != 1:
            sys.exit(f"deck '{key}' not found or ambiguous: {hits}")
        return hits[0]

    decks = {n: json.loads(p.read_text(encoding="utf-8")) for n, p in folders.items()}
    problems = []

    def need(deck, spec):
        s = find(decks[deck_of(deck)]["slides"], spec)
        if s is None:
            problems.append(f"{deck}: no slide for '{spec}'")
        return s

    # ids first: every other link needs them
    for deck, spec, sid in plan.get("ids", []):
        s = need(deck, spec)
        if s is not None:
            s["id"] = sid
    for key, label in plan.get("labels", {}).items():
        decks[deck_of(key)]["ref_label"] = label

    # clean the slate for refs and automatic links, so the plan is the single source
    for d in decks.values():
        for s in d["slides"]:
            s.pop("refs", None)
            for it in items_of(s):
                for p in (it.get("tx") or {}).get("p", []):
                    for r in p["r"]:
                        r.pop("href", None)
            for it in items_of(s):          # merge runs that were split only for a link
                for p in (it.get("tx") or {}).get("p", []):
                    merged = []
                    for r in p["r"]:
                        if merged and "t" in r and "t" in merged[-1] and {k: v for k, v in r.items() if k != "t"} == {k: v for k, v in merged[-1].items() if k != "t"}:
                            merged[-1]["t"] += r["t"]
                        else:
                            merged.append(r)
                    p["r"] = merged
            if s["layout"] == "syllabus":
                for it in s["items"]:
                    it.pop("href", None)

    for deck, spec, refs in plan.get("refs", []):
        s = need(deck, spec)
        if s is not None:
            s["refs"] = [{"to": to, "kind": kind} if len(r) == 2 else {"to": to, "kind": kind, "label": r[2]} for r in refs for to, kind in [r[:2]]]
            if len(s["refs"]) > 2:
                problems.append(f"{deck}: '{spec}' has more than 2 references")
    for deck, spec, text, to in plan.get("inline", []):
        s = need(deck, spec)
        if s is not None and not split_run(s, text, to):
            problems.append(f"{deck}: text '{text}' not found in '{spec}' (inside one formatted run)")

    for title, to in plan.get("syllabus", {}).items():
        for n, d in decks.items():
            for s in d["slides"]:
                if s["layout"] == "syllabus":
                    for it in s["items"]:
                        if norm(it["title"]) == norm(title) and not to.split("#")[0] == n[:len(to.split("#")[0])]:
                            it["href"] = to
    for num, to in plan.get("goals", {}).items():
        for n, d in decks.items():
            if n.startswith(to.split("#")[0]):
                continue
            for s in d["slides"]:
                if s["layout"] == "canvas" and norm(s.get("title", "")).startswith("lecture learning goals"):
                    for it in s["items"]:
                        for p in (it.get("tx") or {}).get("p", []):
                            t = "".join(r.get("t", "") for r in p["r"])
                            m = re.match(rf"Lecture {num}:", t)
                            if m:
                                split_run({"items": [it]}, m.group(0), to)
    for deck, (label, sid) in [(x[0], x[1]) for x in plan.get("navigators", [])]:
        n = deck_of(deck)
        to = f"{deck}#{sid}"
        for s in decks[n]["slides"]:
            if s["layout"] != "canvas" or s.get("id") == sid:
                continue
            for it in s["items"]:
                if it.get("k") == "shape" and it["w"] < 20 and (it.get("tx") or {}).get("p"):
                    for p in it["tx"]["p"]:
                        if norm("".join(r.get("t", "") for r in p["r"])).startswith(norm(label)) and len(p["r"]) <= 2:
                            for r in p["r"]:
                                if r.get("t", "").strip():
                                    r["href"] = to

    if problems:
        print("Could not apply:\n  - " + "\n  - ".join(problems))
    if not a.check:
        for n, p in folders.items():
            p.write_text(pc.dumps(decks[n]), encoding="utf-8")
        nr = sum(len(s.get("refs", [])) for d in decks.values() for s in d["slides"])
        ni = sum(1 for d in decks.values() for s in d["slides"] if s.get("id"))
        nl = sum(1 for d in decks.values() for s in d["slides"] for it in items_of(s) for p in (it.get("tx") or {}).get("p", []) for r in p["r"] if r.get("href"))
        print(f"{ni} slide ids, {nr} reference pills, {nl} linked text runs")
    sys.exit(1 if problems else 0)


if __name__ == "__main__":
    main()
