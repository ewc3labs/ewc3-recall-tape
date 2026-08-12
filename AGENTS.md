# AGENTS.md — RecallTape

> How to work in this repo. Read this and [`docs/origins.md`](docs/origins.md) before you write code.
> For how to work with Wilson generally, see the workspace-level `AGENTS.md` one directory up.

---

## Read the origin doc first. Seriously.

[`docs/origins.md`](docs/origins.md) is the transcribed conversation that created this project — Paul's
question, the platform research, the naming, the architecture sketch, the mascot. It is not sentimental
filler. **Most of the hard decisions are already made in there, with citations.** If you find yourself
about to research "can a OneNote add-in overlay the canvas," stop: that's answered.

Things that are **decided**, not open:

| Decision | Answer |
| --- | --- |
| Platform | **OneNote Windows desktop**, not OneNote for Web |
| Technology | **C#**, against the OneNote **COM/desktop API** |
| Why not Office.js | Microsoft documents the JS add-in API for **OneNote on the web only** |
| Foothold | Desktop API returns page **XML including the current selection**, and takes it back via `UpdatePageContent` |
| Reference implementation | **[OneMore](https://github.com/stevencohn/OneMore)** — open-source C#, already does ribbon + hotkeys + page XML manipulation. Use it as skeleton/reference. Do not build add-in plumbing from scratch. |
| Distribution | **GitHub Releases + signed installer.** COM add-ins are a local-install model; there is no useful desktop OneNote marketplace. |
| Later, maybe | A second **Office.js/web edition** submitted to Microsoft Marketplace. Different product, same Core. |

If you think one of these is wrong, say so with evidence — but know you're arguing against research that
already happened.

---

## What this actually is

**RecallTape — active-recall masking for Microsoft OneNote.**

You lay opaque tape over an answer or a term in your notes, then test whether you can produce it from
memory before revealing it. Goodnotes and Notability have shipped this for years. OneNote has not.

The pitch, in Wilson's preferred order:

> **Tape for OneNote. Finally.**
> Hide answers. Reveal them later. Study directly in your notes.

**The real design insight from the origin doc:** clicking individual tape strips is *not* the target
UX. Hotkey-driven sequential reveal is **better for active recall** — you cruise the page one-handed,
`Space → answer → Space → answer`. Don't reflexively chase the Goodnotes interaction; the keyboard
version is the superior study loop and it's also the achievable one.

---

## Architecture — keep the engine host-agnostic

Sketched in the origin doc and worth honouring, because it's what turns "a janky OneNote hack" into a
reusable engine:

```text
RecallTape.Core       Tape · TapeSet · RevealState · StudySession · Persistence
RecallTape.OneNote    PageAdapter · SelectionAdapter · OneNoteXml · OverlayRenderer
RecallTape.UI         Ribbon · Hotkeys · StudyController
                      ── later ──
RecallTape.Anki · RecallTape.PDF · RecallTape.Web
```

**Core must not know what OneNote is.** No COM types, no page XML, no Office references above the
`.OneNote` boundary. That boundary is what makes the Anki/PDF/Web hosts possible later, and it is
exactly the A.B.S. principle from the workspace AGENTS.md. If Core starts importing OneNote types,
that's the bug — fix it then, not "later."

---

## MVP scope

1. Select text/object → `Ctrl+Alt+T` → **Tape**
2. Overlay/replace the selection with an opaque mask
3. **Preserve the original content and metadata** — taping is never destructive. This is the one
   requirement that must not be compromised: these are somebody's medical school notes.
4. Study loop: hotkey → reveal → next press → retape and advance
5. **Reveal All** / **Retape All** / **Remove Tape**

**Phase 2** (the part that makes it a learning tool rather than a feature clone): mask handwriting,
mask regions of diagrams, mask labels on anatomy images, randomized reveal order, study mode, hit/miss
statistics, spaced-repetition metadata, export marked items to Anki.

**Content types matter.** This is a Surface. **Ink is the primary input, not an edge case** — typed,
handwritten, and dictated all have to work eventually. Typed-only is a legitimate v0.1; pretending ink
is optional is not. Bulk operations like *"tape all handwriting on this page"* are real features,
because hand-taping fifty terms defeats the purpose.

---

## The mascot is spec, not decoration

Because this is a Microsoft add-in, it gets an insufferable assistant. He is a **salad fork** — four
tines, a bend in the neck so he can emote, googly eyes (slightly uneven), and a smug half-smirk. Not
evil. *Obnoxiously delighted.* Legally distinct from a certain paperclip, and funnier for it.

**The animation beat, which is the entire joke:**

1. tape covers the answer
2. user hits reveal
3. fork rockets in from the far edge of the screen
4. slight overshoot
5. hooks the tape with a tine
6. **YANK / RRRIP**
7. turns to face the user
8. smug blink
9. micro-pause
10. tiny cartoon explosion — *"I have fulfilled my purpose and now must die"*
11. fork gone, answer visible

**Step 8 is load-bearing.** The blink before the explosion is where the comedy lives. If you cut it for
timing, you have removed the feature.

He must be **off by default and fully disableable** — see Settings copy in the origin doc
(*Reveal Style: Instant / Animated / [mascot]*, plus **Reduced chaos mode**). Someone cramming for
boards at 2am must be able to turn him off forever in one click. A joke you can't escape is a bug.

**Open naming question:** Wilson has said **Forky**; the origin doc flags that as too generic and
already strongly associated with a certain Pixar utensil, and prefers **Tiney**, **Rippy**, **Sal**, or
**Dork**. This is unresolved — ask before committing the name to code, assets, or docs. Given the repo
is public and the association is with a litigious studio, "already associated elsewhere" is worth taking
seriously.

---

## Repo standards

This is not a toy repo. Wilson's other extension has **5,350+ installs**, which means naming,
packaging, signed binaries, screenshots/GIFs, SEO, releases and upgrade mechanics matter **from day
one** — not "we'll clean that up later."

- **README sells before it documents.** Tagline, then the problem, then a GIF of tape being ripped off,
  then install, *then* reference material.
- **SEO is a feature.** People will search `onenote tape`, `goodnotes tape onenote`, `hide answers in
  onenote`, `active recall onenote`, `onenote flashcard plugin`. Name for humans, metadata for engines.
- **Never break someone's notes.** Non-destructive masking, reversible, and if a page write can fail
  half-way, that path needs belts and suspenders. Losing a medical student's lecture notes is the one
  unrecoverable failure this project has.
- **Privacy:** notes are deeply personal. No telemetry without an explicit, documented, opt-in decision
  from Wilson. Don't add analytics because it seemed helpful.
- **Credit generously.** OneMore, Goodnotes-as-inspiration, and anything else borrowed gets named and
  linked, the way `excel-power-query-editor` credits `excel-datamashup` and `EditExcelPQM`.

---

## Conventions inherited from the org

- `docs/analysis/` evidence · `docs/design/` architecture and *why* · `docs/RAG_sessions/` how a hard
  problem got solved
- `LICENSE`, `CHANGELOG.md`, issue templates, and real CI — see `excel-power-query-editor` for the
  reference shape (`.github/workflows/{ci,release}.yml`)
- Commit messages carry the **why**; the diff already has the what
- **Don't commit without Wilson reviewing the diff**, and don't push unless asked
