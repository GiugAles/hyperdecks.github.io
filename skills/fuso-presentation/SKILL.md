---
name: fuso-presentation
description: Creates a slide deck in the design of the Future Society Hub (FuSo, University of St.Gallen) from a deck.json, rendered as an online presentation (one HTML page that also works on a phone and prints to PDF). Use this whenever the user wants a presentation, slides, a deck or a pitch for FuSo, wants to change or re-render an existing FuSo deck, or wants the slides shown at a Brown Bag Lunch Talk (introduce the speaker and topic, preview the next talk).
---

# FuSo presentation

Output: `html/index.html` next to the deck, with its fonts, logos and images in `html/assets/`. In a window of 1000px or wider it shows one 16:9 slide; in a narrower window (phone) the slide fills the window in one column. Printing from the browser gives a PDF with one slide per page. There is no PowerPoint renderer yet; if the user asks for a .pptx, say so.

Keys while presenting: arrows or space to move, O overview, F full screen, D dark and light, N speaker notes.

Style: the FuSo lunch talk design combined with futuresociety.ch. Inter, white or grey (#f4f4f4) slides, FuSo green (#00802F) for one phrase of a headline, the icon discs and the few green slides, a circuit line (thick line ending in a dot), faint watermark rings, a footer with the FuSo logo and the slide number in the HSG shape. Layout, fonts, sizes and colours are fixed in `scripts/render_html.py` and `scripts/fuso_deck.py`. A deck only chooses layouts and fills their fields; there are no free fields and no per-deck sizes.

## Design rules

* Sizes are moderate on purpose, close to the official University of St.Gallen template. Never scale text or elements up to fill a slide. If a slide is thin, give it the picture, example or fact that supports its headline, or merge it with a neighbour.
* Few slides, plain words. A deck must be understandable to someone outside FuSo. No slides of project details or paper titles unless the user asks for them. Where the message is visual, use a layout that draws it (`fuse`, `flow`, `matrix`, `steps`).
* Green is the loudest colour of the deck. A green slide marks a turn in the talk (at most 2, or a quarter of the slides). The green panel (an `aside` with a QR code) and the green band (`base` under columns) are only for content that warrants it, such as an event or a project with its own site, never for a bare link or a closing sentence.
* No QR code on the title slide unless the user asks for one. The closing slide carries the link.
* Highlight a phrase only in the headlines where one phrase carries the message, not in every headline.
* Neighbouring slides must not use the same layout, except slides of headline and running text.
* One mode per deck, `light` or `dark`. Use `light` unless the user asks for dark.

## Deck format

One JSON file, `deck.json`. Image paths are relative to the folder of the deck.

```json
{"title": "Name of the deck", "mode": "light", "lang": "en",
 "partners": [{"src": "img/partner.svg", "alt": "Partner name"}],
 "slides": [{"layout": "title", "title": "..."}, {"layout": "text", "title": "...", "points": ["..."]}]}
```

`partners` is optional: logos shown beside the HSG logo on the title slide (`src`, `alt`, optional `src_dark`, `name`).

Every slide has `layout` and `title`, and optionally:

| Field | Meaning |
|---|---|
| `highlight` | a phrase of the title, written the same way, shown in green |
| `bg` | `plain` (default), `grey` or `green` |
| `lead` | one sentence under the headline (not on `title`, `section`, `statement`, `closing`) |
| `callout` | a small panel beside the headline: `heading`, optional `text`, `icon` |
| `source` | small credit line at the foot of the slide |
| `notes` | speaker notes (shown with N) |
| `time` | pacing cue shown in the overview |

A line break in a title is written as `\n`. In `paragraphs`, `points`, `after`, `source` and the `text` and `facts` of an aside, `**words**` are bold, `*words*` green and `[words](https://...)` a link.

### Layouts

| Layout | For | Fields besides the common ones |
|---|---|---|
| `title` | first slide | `subtitle`, `presenters` (list of names), `meta` (quiet line, e.g. the room), `qr` (`url`, `label`), `banner` (`image`, `alt`, `url`), `mark` (`brown-bag`: the paper bag of the talk graphics behind the end of the first title line), `kicker` (a word or two in a green pill above the title, e.g. "Welcome!" on a slide people see when they enter the room) |
| `section` | start of a part | `text` |
| `statement` | one sentence that stands alone, or a quote | `text`, `attribution` |
| `text` | running text or points, with an optional picture or panel on the right | `paragraphs`, `points`, `after` (one quiet line under the points), `contacts` (links with an icon each, as on `closing`), `aside` (see below). Needs `paragraphs`, `points` or `contacts`. |
| `columns` | 2 to 6 parallel items as cards | `items` (`heading`, `text`, `icon`), `base` (`text`, `label`: a band under the cards) |
| `steps` | 2 to 5 items in an order, wired by a line | `items` (`heading`, `text`, `icon`), `numbered` (true shows numbers; only for a real order) |
| `compare` | two sides | `left`, `right` (`heading`, `points`, `text`, `icon`), `conclusion` |
| `fuse` | two fields as overlapping circles with FuSo in the overlap | `left`, `right` (`heading`, `text`, `icon`), `conclusion` |
| `flow` | a path from a start to an end | `start`, `end`, `items` (list of short strings) |
| `matrix` | which offer is for whom | `columns` (list of names), `rows` (`label`, `for`: the column names that apply) |
| `agenda` | programme of 2 to 8 entries | `items` (`title`, `when`, `who`, `sub`: list of `title`, `who`) |
| `people` | 1 to 12 people | `items` (`name`, `photo`, `alt`, `role`, `tag`, `points`), `questions` (list; each person's `points` are then the numbered answers) |
| `links` | 2 or 3 linked cards | `items` (`heading`, `url`, `source`, `date`, `text`, `label`) |
| `logos` | a wall of logos | `items` (`src`, `alt`, `src_dark`) |
| `image` | one large picture | `image`, `alt`, `caption` |
| `closing` | last slide | `text`, `contacts` (`label`, `url`, `icon`), `qr` (`url`, `label`) |

Details that are easy to miss:

* `people` with one person is the speaker slide: the `lead` becomes a label above the headline (e.g. "Today's talk"), the headline is the talk title, and the person stands under a circuit line with a large photo.
* `text` points are a list of strings, or all objects with `icon` and `text` (one block per point). Up to 4 short string points (64 characters or fewer each) are set as large statements.
* `aside`: only `image` with `alt` (and `caption`) gives a picture; only `icon` gives a large icon disc; `heading`, `label`, `text`, `facts` (list of `[term, detail]` pairs), `note` and `logos` give a panel, with the `image` on top if there is one (`badge` puts a label on the picture). A `qr` (`url`, `heading`, `text`, `label`) turns the panel green. `photo` (`image`, `alt`) is a round portrait in the top right corner of the panel. `title` is a heading in headline size, for the title of a talk or event the panel announces. `"side": "left"` puts the aside on the left of the copy.
* `columns` and `steps`: give every card an icon or none.
* Icons are names from the Phosphor icon set (MIT), never drawn. Search with `python scripts/fuso_deck.py --find-symbol <word>`. The icon is downloaded from raw.githubusercontent.com at render time and cached.

### Limits the check applies

Headline under 75 characters (110 on `statement`), written as a sentence, no capitals, no final full stop except on `title`, `statement` and `closing`. Under 70 words per slide (`agenda` 130, `people` 120, `links` 90, `title` 60). At most 5 points, 4 per side of a comparison, 3 per person. No dashes as punctuation.

## Where decks go

* A general deck: `presentations/<deck-name>/deck.json`, its pictures in `presentations/<deck-name>/img/`.
* The slides for a Brown Bag Lunch Talk: `talks/brown-bag-talks/YYYY-MM-DD-<FirstName><LastName>/talk/presentation/deck.json`. The speaker photo stays in `talk/` and is referenced as `../<photo file>`.
* Check screenshots go outside the repo.

## Brown Bag Lunch Talk deck

The slides shown in the room before and after the talk. Three slides:

1. `title`: "Brown Bag\nLunch Talk" in black (no `highlight`) with `mark` "brown-bag" and `kicker` "Welcome!", since participants see this slide when they enter the room, `subtitle` the date and time, `meta` the room.
2. `people` with the speaker alone: `lead` "Today's talk", `title` the talk title, the item with `photo`, `alt`, `name` and `role` (the affiliation).
3. `text`, the next talk of the series in a box and, apart from it, where to follow FuSo. `title` "Next Brown Bag Lunch Talk on <Month Day>" with the date as `highlight`. The next talk is the `aside` panel on the left (`"side": "left"`) and nothing of it stands outside the panel: `title` the talk title in quotation marks (“...”), `label` the speaker, `photo` the speaker's picture if the talk folder has one, `text` a note such as a co-operation if the events page names one (as `text`, so it stands before the talk details), `facts` for date, time, room, lunch and registration, and `qr` https://futuresociety.ch/#events with `heading` "All talks" and `label` "futuresociety.ch" (the panel is green; an event warrants it). The copy on the right is only about FuSo: one paragraph "**Get updates from the Future Society Hub**" and `contacts`, which end on the bottom edge of the box. Each label is only the account name or address, since the icon names the network: LinkedIn (`linkedin-logo`), Bluesky (`butterfly`), SoundCloud (`soundcloud-logo`), the website (`globe`) and the email address (`envelope-simple`), with the links from futuresociety.ch.

Take today's talk from `config.json` in its `talk/` folder. Take the next talk from https://futuresociety.ch/#events and from its talk folder; leave out what neither gives (affiliation, photo) and do not invent it. If there is no next talk yet, ask the user what the third slide should show.

## Steps

1. **Collect the content.** Names, dates, rooms, affiliations, numbers and links come from the user, the source text or https://futuresociety.ch. Ask for what is missing. Ask who the audience is and how long the talk is if the request does not say.
2. **Outline the deck** as a list of headlines, one message per slide, each headline a statement the slide supports. Choose the layout that shows the message, following the design rules above.
3. **Write the texts and run the writing check (mandatory, every time).** Read `references/writing-tropes.md` in this skill folder and apply it to every headline and text of the deck, also when you only change a text and render again. Do not replace it with another skill.
4. **Write `deck.json`** and check it without rendering: `python scripts/fuso_deck.py deck.json`. Errors stop the render. Fix the warnings, or keep one only with a reason you report.
5. **Render** with screenshots: `python scripts/render_html.py deck.json --shots <folder outside the repo>`. On an overflow error, shorten the text or split the slide; never change sizes. On a thin slide, see the design rules.
6. **Look at every slide** (`slide-NN.png`) and at the phone view (`contact-phone-NN.png`). Nothing may be cut off, overlap or sit oddly. Then compare the visual weight of the blocks on each slide: no block may be louder than its content deserves. Fix and render again.
7. **Alt text.** Every picture in the deck has its `alt` in the deck, describing only what is visible.
8. **Report** the path of `deck.json` and `html/index.html`, the slides with their texts, what the writing check changed, and any warning that stays and why. If something in the design looks wrong, say what and why and propose a change; do not edit the scripts on your own.

## Setup (once)

```
pip install playwright pillow segno
```

`segno` is only needed for QR codes. The renderer uses Microsoft Edge if installed, otherwise Playwright's Chromium (`playwright install chromium`; ask first, it downloads a browser).

Fonts (Inter 400/700/900) and logos are bundled in `assets/`; only the icons are downloaded.

Icons: Phosphor Icons, MIT license, https://phosphoricons.com (license in `assets/phosphor/LICENSE`).
