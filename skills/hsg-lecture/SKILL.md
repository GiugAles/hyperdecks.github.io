---
name: hsg-lecture
description: Creates lecture slides in the University of St.Gallen (HSG) design of the Computersysteme lectures from a deck.json, rendered as an online presentation (one HTML page that also works on a phone and prints to PDF). Use this whenever the user wants lecture slides, a lecture, a teaching session, exercise or homework slides, or course material in the HSG / Computersysteme style, or wants to change or re-render such a deck.
---

# HSG lecture slides

Output: `html/index.html` next to the deck, with its fonts, logo and images in `html/assets/`. In a window of 1000px or wider it shows one 16:9 slide; in a narrower window (phone) the slide fills the window in one column. Printing from the browser gives a PDF with one slide per page. There is no PowerPoint renderer; if the user asks for a .pptx, say so.

Keys while presenting: arrows or space to move, O overview, F full screen, N speaker notes.

Style: the HSG slide template as used in the lectures 01a Introduction, 01b Information Processing and 02 Circuits of Computersysteme (the three .pptx files in `../fuso-presentation/`). White slides, Gill Sans Nova as on the official HSG template (Regular for headlines and text, Light for the large titles of the title, section and quote slides), headline top left at 32pt, points marked with a dash, the slide number in the green HSG tab at the bottom right, HSG green (#00802F) for boxes with questions to the room and table headers, teal (#73A5AF) for remarks, salmon (#EB6969) for tasks and assignments, the title slide with the HSG logo, the gear word cloud and the green wedge "From insight to impact.". Layout, fonts, sizes and colours are fixed in `scripts/render_html.py` and `scripts/hsg_deck.py`. A deck only chooses layouts and fills their fields; there are no free fields and no per-deck sizes (the exception is `canvas`, below, which keeps the sizes and positions of a PowerPoint slide).

## Design rules

* Sizes are those of the template. Never scale text up to fill a slide or down to fit more on it. If a slide is too full, split it; if it is thin, give it the picture, example or question that supports its headline.
* One message per slide, written for students who see the topic for the first time. A build-up (a circuit growing over three slides, the same table with one more column) is several slides with the same layout, which is fine.
* Green boxes are questions or tasks for the room, not decoration: a `callout` beside the headline, the boxes of `quiz`. Outline boxes hold an example or a note, teal boxes a remark. At most one box beside the headline.
* The salmon label of `task` is for homework and preparation ("Solve the first five levels of the NAND Game!"), at most twice per lecture.
* Pictures carry the lecture as much as text does: a photo of the real device, a screenshot of the demo, the cover of the book. Every picture has its `alt`, and a source line when it is not the lecturer's own.
* A lecture usually starts with `title`, the `syllabus` with today's session marked, and `menu`; it usually ends with `questions` (take-homes), `study` (how to study this material) and `task` (preparing for the next lecture or exercise). Keep this frame unless the user asks otherwise.

## Deck format

One JSON file, `deck.json`. Image paths are relative to the folder of the deck.

```json
{"title": "Computing on Bits", "lang": "en",
 "footer": "Computersysteme | Institute of Computer Science | University of St.Gallen",
 "slides": [{"layout": "title", "title": "..."}, {"layout": "text", "title": "...", "points": ["..."]}]}
```

`footer` is optional: a quiet line at the foot of every content slide (hidden on slides with a `source`).

`theme` is optional: a CSS file (path relative to the deck) set on top of the template, for an event with its own look, e.g. `legalhackathons/theme/legalhackathons.css`. It may use the unit `U` and link fonts with `url()`; those files are copied into `html/assets/`. A theme changes colours and fonts only, never sizes; only use one when the user asks for another look.

Every slide has `layout` and `title`, and optionally:

| Field | Meaning |
|---|---|
| `lead` | one line under the headline, at body size (not on `title`, `section`, `task`, `quote`) |
| `callout` | a box beside the headline: `text`, optional `style` (`green` default, `outline`, `teal`), `heading`, `wide` (true lets the box grow to 48U instead of 34U, for lines that should not wrap, e.g. names with affiliations) |
| `source` | small credit line at the foot of the slide, links allowed |
| `notes` | speaker notes (shown with N) |
| `time` | pacing cue shown in the overview |

A line break in a title is written as `\n`, and `**words**` in a title is set bold (use it sparingly, e.g. for the session name). In every other text field: `**words**` bold, `*words*` italic, `` `code` `` monospace, `==words==` yellow marker, `10_{2}` subscript, `2^{w-1}` superscript, `[words](https://...)` a link.

### Layouts

| Layout | For | Fields besides the common ones |
|---|---|---|
| `title` | first slide | `kicker` (line above the title: the course or "Welcome and Introduction"), `presenters` (list of names), `meta` (list of lines, e.g. "University of St.Gallen", "Bachelor in Computer Science", "", "Computersysteme (1,904)"; an empty string is a gap), `picture` (`image`, `alt`; replaces the gear word cloud) |
| `section` | start of a part of the lecture | `kicker` |
| `syllabus` | the course plan | `items` (`date`, `title`, `kind`: `lecture` default or `exercise`), `current` (number of today's session, or a list), `break_after` (number of the session before a break), `assignments` (`title`, `from`, `to`: session numbers the assignment runs over) |
| `menu` | today's topics in a green frame | `items` (2 to 8 strings), `current` (number of the topic now starting, shown green) |
| `text` | running text or points, with an optional picture, box or code on the right | `paragraphs`, `points`, `defs` (pairs `[term, detail]`, set as two columns), `groups` (`heading`, `points`, `paragraphs`), `after` (one line under the content), `aside` (see below). Needs `paragraphs`, `points`, `defs` or `groups`. |
| `table` | a table with a green header row | `columns` (2 to 6 headings), `rows` (lists of cells; the first column is bold), `after` |
| `code` | a program and what to look at in it | `code` (string, `\n` between lines), `paragraphs`, `points`, `style` (`outline` default, `teal`), `caption` (e.g. a link to godbolt.org) |
| `quiz` | 2 to 5 questions for the room, in green boxes | `items` (strings), `aside` (a picture: `image`, `alt`, `caption`) |
| `image` | one large picture | `image`, `alt`, `caption`, `shadow` (false for drawings on white) |
| `gallery` | 2 to 4 pictures side by side | `items` (`image`, `alt`, `caption`, `shadow`), `after` |
| `people` | 1 to 6 people, e.g. the course team | `items` (`name`, `role`, `photo`, `alt`) |
| `quote` | a quote as the headline, right aligned | `title` is the quote with its quotation marks, `attribution`, `picture` (`image`, `alt`), `points` (up to 4, about the person) |
| `questions` | "Any Questions / Comments / Doubts / Concerns?" with the figure on the question mark | `heading` (default "Take-Home"), `points` (the take-homes, up to 5), `after` (one line under them). Without `points` the slide is the headline and the figure only. |
| `study` | "How to Study this Material?" with the light bulb | `groups` (e.g. "Basic Depth" and "Advanced Depth", each with `points`), `aside` |
| `canvas` | a slide copied from a PowerPoint slide: diagrams, arrows, speech bubbles, free arrangements, click builds | `items` (made by `scripts/pptx_canvas.py`, see below); `title` may be empty |
| `task` | homework or preparation, in the tilted salmon label | `stamp` (the task), `paragraphs` (left of the label), `url` and `label` (the link under it) |

Details that are easy to miss:

* `points` are strings, or objects `{"text": "...", "sub": ["...", "..."]}` for sub-points (up to 3).
* `aside` (on `text` and `study`) is exactly one of: a picture (`image`, `alt`, optional `caption`, `shadow`), a box (`box`: `text`, `style`, `heading`) or code (`code`).
* Pictures get the soft drop shadow of the template. Set `"shadow": false` for drawings, diagrams and logos on white.
* Running text is set at body size (18.7pt). When a slide has more than about 55 words, the renderer sets it one step smaller (16pt); that is the smallest size. More text means another slide.
* Formulas and circuit diagrams are pictures (`image` or an `aside` picture); there is no formula renderer. Short ones fit in text with `_{}` and `^{}`.

### The `canvas` layout (slides that only work as drawn)

Use it only when a PowerPoint slide cannot be said with the other layouts: a diagram drawn from shapes and arrows, a message flow, a build-up that appears click by click, a free arrangement of pictures, boxes and notes. It is the one exception to "no free fields", so keep it for those slides (in 07b about a third) and use the normal layouts for everything else.

```
python scripts/pptx_canvas.py deck.json --pptx lecture.pptx --slide 60            # append
python scripts/pptx_canvas.py deck.json --pptx lecture.pptx --slide 60 --at 12     # replace slide 12 of the deck
python scripts/pptx_canvas.py deck.json --pptx lecture.pptx --slide 60 --insert 12 # insert before slide 12
```

* To convert a whole lecture at once: `python scripts/pptx_to_deck.py lecture.pptx deck.json --syllabus-from other/deck.json --current 8` (hidden slides are skipped; title, syllabus, questions and plain text slides get their normal layouts, everything else becomes `canvas`). Pictures are saved once per image (file name from its hash), so decks can share them.
* Text is set in the template's Gill Sans Nova Regular, which is a little wider than the Light of many source decks, so text can wrap onto more lines than in PowerPoint. The headline of a `canvas` slide may use the full width; text that runs out of its box shrinks until it fits, by at most 14% (as PowerPoint's "shrink text on overflow"). If that is not enough, shorten the text or widen the box in `deck.json`.
* The converter reads the .pptx and writes a `canvas` slide: every object with its position (`x`, `y`, `w`, `h` in U, 1% of the slide width), rectangles, rounded boxes, ellipses, block arrows, speech bubbles, lines and connectors with arrowheads, freeform paths, tables and pictures (saved in `img/`). Text stays text, at its PowerPoint size. Colours come from the theme of the slide master.
* **Click animations become steps**: the arrow keys (and the buttons) reveal the next part of the slide before they move on; going back hides it again. Overview, print and the phone show the whole slide. Entrance and exit effects are copied; emphasis and paragraph-by-paragraph builds are not (the shape appears whole).
* The title comes from the title placeholder. The footer, slide number and title placeholders are not copied (the frame draws them).
* After converting: **write the `alt` of every picture** in the deck (the check lists those without), compare the render with the .pdf of the lecture, and read the converter's warnings (shape types it only approximated, rotated groups, picture formats a browser cannot show).
* Not supported: SmartArt, charts, video, rotated groups, gradient fills (first colour is used), shadows of shapes, text effects. Say so in the report; for such a slide a picture of it is the last resort.
* `items` are one per line in `deck.json`; they can be edited by hand (move a box, change a colour, delete an item), but re-running the converter for the slide replaces them.

### Limits the check applies

Headline under 80 characters (140 on `quote`), not in capitals. Words per slide: `text` 90, `study` 110, `table` 140, `code` 120 (code itself is not counted), `syllabus` 140, `questions` 80, `quiz` 70, `people` and `gallery` 60, `title` 50. At most 6 points (4 per group, 5 take-homes), 3 sub-points per point, 8 `defs`, 10 table rows, 18 lines of code of at most 64 characters. Neighbouring slides should not use the same layout, except `text`, `image`, `code` and `table`. No dashes as punctuation.

## Where decks go

* A lecture: `lectures/<course>/<NN-topic>/deck.json`, e.g. `lectures/computersysteme/02-circuits/deck.json`, its pictures in `img/` beside it.
* Check screenshots go outside the repo.

## Steps

1. **Collect the content.** Topics, dates, rooms, course number, names and links come from the user, their existing slides or the course materials. Ask for what is missing. Ask which lecture of the course it is (for `syllabus` and `menu`) and how long the session is if the request does not say. To reuse content from a .pptx, read it with the pptx skill.
2. **Outline the deck** as a list of headlines, one message per slide, in the order of the lecture. Choose the layout that shows the message, following the design rules above. About 1.5 to 2 minutes per slide, more for slides with a question to the room.
3. **Write the texts and run the writing check (mandatory, every time).** Read `references/writing.md` in this skill folder and apply it to every headline and text of the deck, also when you only change a text and render again. Do not replace it with another skill.
4. **Write `deck.json`** and check it without rendering: `python scripts/hsg_deck.py deck.json`. Errors stop the render. Fix the warnings, or keep one only with a reason you report.
5. **Render** with screenshots: `python scripts/render_html.py deck.json --shots <folder outside the repo>`. On an overflow error, shorten the text or split the slide; never change sizes. A thin slide needs its picture, example or question, or a merge with a neighbour (`menu`, `task` and the bare `questions` slide are thin on purpose).
6. **Look at every slide** (`slide-NN.png`) and at the phone view (`contact-phone-NN.png`). Nothing may be cut off, overlap or sit oddly, and no box may be louder than its content deserves. Fix and render again.
7. **Alt text.** Every picture in the deck has its `alt` in the deck, describing only what is visible.
8. **Report** the path of `deck.json` and `html/index.html`, the slides with their headlines, what the writing check changed, and any warning that stays and why. If something in the design looks wrong, say what and why and propose a change; do not edit the scripts on your own.

## Setup (once)

```
pip install playwright pillow
```

The renderer uses Microsoft Edge if installed, otherwise Playwright's Chromium (`playwright install chromium`; ask first, it downloads a browser). The Playwright version must match the browser in `~/.cache/ms-playwright`; if it does not, install the matching version (`pip install "playwright==1.62.*"` for the Chromium build 1234) instead of downloading a new browser.

Fonts: the page uses Gill Sans Nova (Light, Regular, Bold), else Gill Sans MT Pro, and Palatino Linotype where they are installed; these fonts are licensed and are never copied into `html/assets/`. Everywhere else it falls back to the bundled Gillius ADF No2 (GPL with font exception, `assets/fonts/GILLIUS-LICENSE.txt`) and TeX Gyre Pagella (GUST Font License, `assets/fonts/PAGELLA-LICENSE.txt`), which are close but a little wider and heavier; the screenshots of the check use the fallback, so text that fits there fits with Gill Sans.

Pictures bundled in `assets/` come from the Computersysteme decks: the HSG logo, the gear word cloud of the title slide, the figure on the question mark and the light bulb.
