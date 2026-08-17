# The Assistant — reveal animation spec

> Because this is a Microsoft add-in, it gets an insufferable office assistant. This document is
> **specification**, not flavour text. Transcribed and tightened from
> [`../origins.md`](../origins.md).

---

## ⚠️ Naming is UNRESOLVED

Wilson has said **Forky**. `origins.md` argues against it — too generic, and strongly associated
with a certain Pixar utensil, which is a real consideration for a public repo carrying a
studio-adjacent name. The alternatives raised there, in rough order of preference:

**Tiney** · **Rippy** · **Sal** · **Dork** · Tines · Forklift · Ripley

**Settle this before the name reaches code, asset filenames, settings copy, or store metadata.**
Renaming a mascot after release is worse than choosing carefully now. This doc says *"the
Assistant"* throughout so nothing has to be rewritten when it's decided.

---

## Character

A **salad fork**. Not a paperclip — a paperclip reads as *Microsoft parody*, while a salad fork with
googly eyes reads as *built by the correct kind of maniacs*. It is also considerably safer legally,
and funnier, which is the same argument twice.

**Design:**
- small silver salad fork, **four tines**
- a bend in the "neck" so he can emote and lean
- **googly eyes, slightly uneven** — the unevenness is the whole face
- smug half-smile / stupid smirk; a hint of arched eyebrow if it can be done
- subtle bounce or float when idle

**Personality:** far too pleased with himself. Unhelpful in an aggressively helpful way. Behaves as
though this were a triumphant intervention. **Not evil — obnoxiously delighted.**

---

## The animation beat

The joke is entirely in the timing. Implement the beats in order.

| # | Beat |
| --- | --- |
| 1 | tape covers the answer |
| 2 | user triggers reveal |
| 3 | fork **rockets in** from the far edge of the screen |
| 4 | slight **overshoot** |
| 5 | hooks the tape with one tine |
| 6 | **YANK — `RRRIP`** |
| 7 | turns to face the user |
| 8 | **smug blink** |
| 9 | micro-pause |
| 10 | tiny cartoon explosion |
| 11 | fork gone; answer visible |

### Beat 8 is load-bearing

> *"That blink-before-explosion is crucial. That's where the whole joke lives."*

If the animation gets shortened for performance, **the blink and the micro-pause are the last things
to go, not the first.** Cutting them leaves a fork that rips tape and vanishes, which is merely an
animation. The comedy is that he stops, looks at you, is visibly pleased with himself, and *only
then* detonates.

The explosion is **small and cheerful** — not violent. The register is:

> *"I have fulfilled my purpose and now must die."*

### Sound design (optional, off by default)

| Beat | Sound |
| --- | --- |
| entrance | tiny whoosh |
| rip | exaggerated adhesive peel |
| blink | small sparkle / click |
| explosion | comically small **poof** |

---

## Settings — he must be escapable

**Non-negotiable: off by default, and disableable forever in one click.** Someone cramming for
boards at 2am must be able to make him stop. A joke you can't get out of is a bug, and it is the
single fastest way to turn a delightful feature into a one-star review.

**Reveal Style**
- `Instant` ← **default**
- `Animated` (tape peels; no character)
- `[the Assistant]` (the full bit)

**General**
- Enable reveal animation
- Use assistant mascot for reveal
- Explosion effects
- **Reduced chaos mode**

**Tooltip copy** (already perfect, don't rewrite):
> *"[Assistant] removes the tape for you. May blink. May explode. Not actually helpful."*

---

## Unlocks

He is better as a **discovery** than as a default. Suggested: unlock at **50 reveals** — by then the
user is actually studying, and the first appearance lands as a reward rather than a distraction.

**Achievements:** *Unwanted Assistance* · *Fork Found* · *Study Utensil* · *Fine Tines*

---

## Branding

**He is not the brand.** RecallTape is a study tool that happens to contain an idiot; it is not "the
fork app." Keep him out of the primary logo, the extension icon, and the top of the README.

Where he belongs: a GIF partway down the README (he is *extremely* good marketing once someone is
already reading), the settings screen, and the moment of unlock.
