# How to use RecallTape

Cover something up. Look at it later and try to remember what was under it. Peek when you give up.

That's the whole idea. This page covers everything the buttons actually do today, and — just as
usefully — what they don't do yet.

---

## The thirty-second version

You get a **RecallTape** tab in the OneNote ribbon with three buttons:

| Button | What it does |
|---|---|
| **Tape** | Covers whatever you selected. Selected nothing? You get a box you can drag anywhere. |
| **Remove Tape** | Removes just the tape you're pointing at. |
| **Remove All** | Removes every tape on the page. Counts them, asks first, defaults to *No*. |

And one gesture that isn't a button: **Ctrl+click a tape to peek under it.** Ctrl+click again to
cover it back up.

Not installed yet? See [Install](README.md#install) in the README. It takes about a minute.

---

## Taping things

Tape works three different ways depending on what you've got selected. You don't choose the mode —
it picks based on what you gave it.

### Text

Select the words you want hidden, press **Tape**. You get a black bar exactly as wide as what you
selected — not the whole line, not the whole paragraph. Select three words in the middle of a
sentence and you'll cover three words.

The text is still there. It's still searchable, still copyable, still real text. It's just black
on black.

### Ink, images, and diagrams

Select them the normal OneNote way — lasso your ink, click an image, or drag a box around several
things — then press **Tape**. You get **one** box covering the whole selection, not one box per item.

Handy for anatomy: lasso every label on a diagram at once and they vanish together.

### Anything you can't select

Lecture printouts, page backgrounds, PDF slides pasted as page images — OneNote won't let you select
those at all. Some of them aren't really objects, they're the page's wallpaper.

So: **select nothing and press Tape.** You get a free-floating box in the top-left. Drag it over
whatever you want hidden, resize it from the corners. Press Tape again for another one, and each new
box appears slightly offset so they don't stack invisibly.

This is the one to reach for when nothing else works.

---

## Peeking

**Ctrl+click any tape.** Ctrl+click is OneNote's own convention for following a link, and that's
literally what's happening under the hood.

The content appears. The tape doesn't vanish though — it leaves a **grey outline** exactly where it
was. That's deliberate: with twenty tapes on a page, an invisible peeked tape is a tape you'll never
find again. The outline says *"something lives here, and it's currently showing."*

**Ctrl+click the outline** to cover it back up.

---

## Removing tape

### Just the one

**Put your cursor inside the taped text**, or **click a tape box**, then press **Remove Tape**.

You don't have to select anything or highlight anything — just click into it. Which is the point,
because selecting black text on a black background is a game nobody wants to play.

One caveat worth knowing: if you've taped several separate words *in the same paragraph*, clicking
into one of them removes all of them. That's a deliberate trade — the alternative risks stripping
half a tape and leaving the other half sitting there.

### All of it

**Remove All** clears the entire page. It tells you how many it found, asks you to confirm, and the
dialog defaults to **No** — you have to actively choose *Yes*.

There's no undo that reaches back through a OneNote sync, so this one is meant to be slightly
annoying.

---

## Studying with it

The workflow it's built for:

1. Tape the answers on a page — labels, definitions, drug names, the right-hand column of a table
2. Come back later and read the page
3. Try to recall what's under each bar **before** you peek
4. Ctrl+click to check
5. Ctrl+click again to re-cover, so the page is ready for the next pass

The re-cover step is what makes it repeatable. A page stays useful for weeks.

---

## Your notes are safe, and here's the actual reason

Not a promise — a checkable fact. **OneNote backs your notebooks up on its own**, automatically,
independently of RecallTape:

```
%LOCALAPPDATA%\Microsoft\OneNote\16.0\Backup
```

Paste that into File Explorer. You'll find dated copies of your sections sitting there right now,
made before RecallTape ever touched them. If anything ever goes wrong, that folder is your way back,
and it exists whether or not you use this add-in.

Two more things that are true:

- **Tape never deletes your content.** Text tape wraps words in black styling; box tape lays a
  picture on top. The original is underneath, intact, the whole time.
- **Uninstalling doesn't strip your pages.** Tape stays as ordinary OneNote content. If you want a
  page clean, run **Remove All** on it *before* you uninstall.

---

## When something goes wrong

**No RecallTape tab after installing.**
Close OneNote **completely** and reopen it. OneNote lingers in the background after its window
closes, so if the tab is still missing, end any `ONENOTE.EXE` in Task Manager first, then reopen.

**"Missing a file" when installing.**
Extract the *whole* zip before running anything. Running the installer from inside the zip viewer
leaves the rest of the files behind.

**Where did it install?**
`C:\Program Files\EWC3 Labs\RecallTape`. The folder you extracted is only a delivery vehicle — once
the installer has run you can delete it. The uninstaller lives in the installed folder.

**The installer won't run / no "Run as administrator".**
Use `Install RecallTape.cmd` — double-click it, and it asks for admin itself. You don't need to find
a right-click option. It also handles PowerShell's execution-policy block for you.

**Windows says the publisher is unrecognised.**
It isn't code-signed yet. *More info* → *Run anyway*. That's a real gap, honestly stated, and signing
is on the list.

**A tape won't come off.**
Ctrl+click it first to confirm it's ours — if content appears, it's RecallTape's. If nothing happens,
it may be ordinary OneNote highlighting, which we deliberately leave alone.

**A Copilot icon keeps appearing when you click.**
That's Microsoft's, not ours, and it shows up on every OneNote page whether RecallTape is installed
or not. Ignore it.

**Something else.**
There's a log at `%LOCALAPPDATA%\EWC3 Labs\RecallTape\recalltape.log`. It records what each button did and any errors,
with no page content in it. Attach it to an issue.

---

## What it can't do yet

Being straight about the edges, so you don't waste time hunting for these:

- **No keyboard shortcuts.** Ribbon buttons only, for now.
- **Tape doesn't follow content.** Reflow a page and boxes stay where you put them. Anchoring tape to
  an object is designed but not built.
- **No study statistics.** Nothing tracks what you peeked or how often.
- **Boxes resize from corners only**, and OneNote keeps the aspect ratio. That's OneNote's image
  handling, not a setting we can reach. Workaround: make a new box at the shape you want.
- **Sync conflicts are invisible to us.** If OneNote flags a page as conflicted, the add-in can't
  tell — the API exposes no way to ask. Resolve conflicts in OneNote as normal.

---

## Telling us it's broken

[Open an issue](https://github.com/ewc3labs/ewc3-recall-tape/issues) — what you were doing, what
happened, and the log if you have it.

This is pre-alpha software being built alongside somebody actually using it to study. Bug reports
are the whole development process right now, not an interruption to it.
