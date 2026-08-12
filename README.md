<div align="center">

<img src="images/EWC3LabsLogo-blue-128x128.png" alt="EWC3 Labs" width="96" height="96">

# RecallTape

### Tape for OneNote. Finally.

**Hide answers. Reveal them later. Study directly in your notes.**

Active-recall masking for Microsoft OneNote — the Goodnotes-style tape tool
that OneNote never shipped.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Status: pre-alpha](https://img.shields.io/badge/status-pre--alpha-orange.svg)](#status)
[![Made by EWC3 Labs](https://img.shields.io/badge/EWC3-Labs-1f6feb.svg)](https://github.com/ewc3labs)

</div>

---

## The problem

> *"Goodnotes has Tape. I use OneNote. Why the hell can't OneNote do this?"*

If you study from your own notes, you already know the technique: cover the answer, try to recall it,
then check. Goodnotes and Notability have shipped a tape tool for years. OneNote — the note app with
the best pen input on Windows, the one running on every Surface in every lecture hall — has never had
one.

The usual workarounds are all bad. Collapse an outline and lose your layout. Scribble over the answer
with a highlighter and then undo it. Keep your notes in OneNote and your recall in a second app, and
do everything twice.

**RecallTape puts the tape in OneNote.**

## What it does

Select an answer, a term, a label on a diagram. Press a hotkey. It's covered.

Then study one-handed:

```
Space → answer → Space → answer → Space
```

Reveal, check yourself, retape, advance. Straight down your own lecture notes, no flashcards to build
and no second app to sync.

> **Why hotkeys instead of tapping each strip?** Because it's *better*. Clicking tiny pieces of tape
> is fiddly on a tablet and slow on a laptop. A single key that reveals-then-advances turns a page of
> notes into a study loop you can run without looking at the keyboard.

### Planned

| | |
| --- | --- |
| **Tape** | `Ctrl+Alt+T` on any selection |
| **Study loop** | reveal → retape → advance, on one key |
| **Reveal All / Retape All** | reset the page in one go |
| **Remove Tape** | non-destructive, always |
| **Ink support** | handwriting is the *point* on a Surface, not an afterthought |
| **Tape all handwriting** | bulk-tape a page instead of hand-picking fifty terms |
| **Study stats** | hits, misses, and eventually spaced repetition |
| **Anki export** | take your taped items with you |

## Status

**Pre-alpha. Nothing to install yet.** This repo currently holds the design, the platform research,
and the origin story. Watch or star if you want to know when there's a build.

## How it's built

Windows **desktop** OneNote, C#, against the OneNote COM API — which hands you the current page as XML
*including the selection*, and takes your modified version back.

That's a deliberate choice, not a default. Microsoft's modern Office.js add-in API targets **OneNote on
the web**, which is not where lecture notes get taken with a pen. The desktop API is older and weirder
and it's the one that can actually reach your ink.

We are standing on [**OneMore**](https://github.com/stevencohn/OneMore) by
[Steven Cohn](https://github.com/stevencohn) — an excellent open-source OneNote add-in that already
solved ribbon integration, hotkeys, and page XML manipulation. Building that plumbing again from
scratch would have been foolish.

The masking engine is kept deliberately host-agnostic, so OneNote is its *first* host rather than its
only one:

```text
RecallTape.Core       the tape model — knows nothing about OneNote
RecallTape.OneNote    page + selection adapters, XML, overlay rendering
RecallTape.UI         ribbon, hotkeys, study controller
```

## Origins

Wilson's son Paul started medical school and asked:

> **"yo how can i tape in onenote? is there an add-on?"**

There wasn't. So we're building one. The full conversation — including the platform research, the
architecture argument, and the part where a salad fork became a required feature — is preserved in
[`docs/origins.md`](docs/origins.md).

## Contributing

Not yet, but soon — the structure has to land first. When it opens up, see
[`CONTRIBUTING.md`](CONTRIBUTING.md). Issues and ideas are welcome any time.

## License

[MIT](LICENSE) — © EWC3 Labs.

Not affiliated with or endorsed by Microsoft, Goodnotes, or Notability. OneNote is a Microsoft
trademark; we just wish it had tape.

---

<div align="center">

**[EWC3 Labs](https://github.com/ewc3labs)** — *tools that make Microsoft's power-user products behave
like somebody actually uses them.*

</div>
