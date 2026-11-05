# Writing check for lecture slides

Run this on every headline, point, box and caption of the deck before rendering, also when you only change one text. This file is the check; it does not depend on any other skill being installed. The slides must read as if the lecturer typed them: short, plain, direct, addressed to students in the room. Anything below that does not fit is rewritten or deleted.

## The register of the Computersysteme decks

* Headlines name the topic or ask the question the slide answers: "Bits: Binary Digits", "Why does this happen?", "Let's build a Calculator!". Title case or sentence case, as the lecturer writes them; keep one style per deck.
* Questions to the room are a teaching tool and belong on the slides ("What number range can be represented using n bits?"). Put them in a green box (`callout`, `quiz`) so they stand out from facts.
* Points are fragments or short sentences, not paragraphs. One fact or one step per point. Sub-points carry the detail.
* An exclamation mark or a ";-)" is fine where the lecturer would say it out loud (a task, a surprise), at most once or twice per deck. Do not add jokes or emoticons the source does not contain.
* Terms are used exactly: name the data type, the unit, the base (`10_{2}` for binary, `0x1F` for hexadecimal), the source of a number.

## How to check

1. Read the text sentence by sentence. For each sentence ask: does it state a fact, a step, a definition or a question for the students, or is it only there for effect or to connect two other sentences? Delete the ones that do none of these.
2. Go through the lists below and rewrite every match in plain words.
3. Check every claim, number and date against the source material or the user. Do not add claims, numbers or conclusions the source does not contain, and do not state them more strongly than the source does. Leave out what you cannot check, and say so.
4. Read the headline and the points aloud once. If a sentence sounds like a slogan, a summary line or a commentator, rewrite it as a plain statement.
5. Check the lengths again after rewriting (`python scripts/hsg_deck.py deck.json`).

## Sentence patterns to remove

* Bridge and commentary sentences that state no fact: "And that matters.", "This is where it gets interesting.", "This changes everything."
* Hooks that dramatise the topic: "Ever wondered ...", "Imagine a world ...", "What if I told you ...". (A real question to the students is not a hook.)
* Contrast frames: "not X, but Y", "It's not just X, it's Y", "more than just X".
* Lists of three and parallel pairs that are only there for rhythm.
* Colon or dash reveals: "The result: ...", "The catch? ...".
* Sweeping claims: "the first", "never", "always", "everyone", "revolutionary", unless the source states them.
* A closing moral or takeaway that the slides before do not support.
* Announcing what comes next instead of showing it: "Let's take a look at ...", "Here's what this means". (Naming the next step of a build-up, "Now we build a larger adder", is fine.)

## Words and phrases to remove

* Inflated verbs for "is" or "shows": unpacks, sheds light on, dives into, delves into, explores, tackles, navigates, underscores, highlights.
* Stiff filler: it is worth noting, in this context, when it comes to, plays a role in, at its core.
* Hype: exciting, fascinating, groundbreaking, game changer, powerful, seamless, cutting-edge, robust, leverage, unlock, landscape, realm, crucial, pivotal, key (as an adjective).

## Formatting

* No dashes as punctuation in running text ("–", "—"); the template's own dash marks the points.
* No emojis other than the lecturer's ";-)" and ":-)", no hashtags.
* Links only where students need them (a tool, a source, the course page), written as `[label](https://...)`.
