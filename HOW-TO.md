# How to use RecallTape

Cover something up. Look at it later and try to remember what was under it. Peek when you give up.

That's the whole idea. This page covers everything the buttons actually do today, and — just as
usefully — what they don't do yet.

---

## The thirty-second version

You get a **RecallTape** tab in the OneNote ribbon with three buttons:

| Button | What it does |
|---|---|
| **Apply Tape** | Covers whatever you selected. Selected nothing? You get a box you can drag anywhere. |
| **Remove Selected Tape** | Removes just the tape you're pointing at. |
| **Remove All Tape** | Removes every tape on the page. Counts them, asks first, defaults to *No*. |

And one thing that isn't a button: **peeking**. How you peek depends on which kind of tape it is —
see [Peeking](#peeking), because the two kinds genuinely behave differently.

Not installed yet? See [Install](README.md#install) in the README. It takes about a minute.

---

## Taping things

Tape works three different ways depending on what you've got selected. You don't choose the mode —
it picks based on what you gave it.

### Text

Select the words you want hidden, press **Apply Tape**. You get a black bar exactly as wide as what you
selected — not the whole line, not the whole paragraph. Select three words in the middle of a
sentence and you'll cover three words.

The text is still there. It's still searchable, still copyable, still real text. It's just black
on black.

### Ink, images, and diagrams

Select them the normal OneNote way — lasso your ink, click an image, or drag a box around several
things — then press **Apply Tape**. You get **one** box covering the whole selection, not one box per item.

Handy for anatomy: lasso every label on a diagram at once and they vanish together.

### Anything you can't select

Lecture printouts, page backgrounds, PDF slides pasted as page images — OneNote won't let you select
those at all. Some of them aren't really objects, they're the page's wallpaper.

So: **select nothing and press Apply Tape.** You get a free-floating box in the top-left. Drag it over
whatever you want hidden, resize it from the corners. Press Apply Tape again for another one, and each new
box appears slightly offset so they don't stack invisibly.

This is the one to reach for when nothing else works.

---

## Peeking

There are two kinds of tape and **they peek differently.** This is the one place RecallTape isn't
uniform yet, so it's worth thirty seconds of your attention.

### Tape boxes — Ctrl+click

Anything covered by a **box** — ink, images, page backgrounds, free boxes you dragged into place —
is a real object with a link on it. **Ctrl+click it** and the content appears.

The tape doesn't vanish; it leaves a **grey outline** exactly where it was. That's deliberate: with
twenty tapes on a page, an invisible peeked tape is one you'll never find again. The outline says
*"something lives here, and it's showing right now."* **Ctrl+click the outline** to cover it back up.

This reveal is **sticky**. It stays revealed until you put it back, including after you navigate
away and return.

### Taped text — select it

Taped **text** is not an object, it's your own words coloured black on black. There's nothing to
click, and Ctrl+click does nothing at all.

Instead, **drag-select across it** — the selection highlight sits behind the letters and makes them
legible. Or press **Ctrl+A** to select everything and reveal every taped run on the page at once.

This reveal is **momentary**. Click anywhere else and it hides itself again, with nothing to put
back and no state left behind.

### Which is better?

Honestly, the momentary one — for studying. Select to check yourself, click away, and the page is
already reset for the next pass. Boxes make you undo the reveal by hand.

They should be the same gesture, and making taped text clickable is on the list. Until then this is
the real behaviour rather than the behaviour we'd like to describe.

---

## Removing tape

### Just the one

**Put your cursor inside the taped text**, or **click a tape box**, then press **Remove Selected Tape**.

You don't have to select anything or highlight anything — just click into it. Which is the point,
because selecting black text on a black background is a game nobody wants to play.

One caveat worth knowing: if you've taped several separate words *in the same paragraph*, clicking
into one of them removes all of them. That's a deliberate trade — the alternative risks stripping
half a tape and leaving the other half sitting there.

### All of it

**Remove All Tape** clears the entire page. It tells you how many it found, asks you to confirm, and the
dialog defaults to **No** — you have to actively choose *Yes*.

There's no undo that reaches back through a OneNote sync, so this one is meant to be slightly
annoying.

---

## Settings

The **Settings** button on the left of the tab opens a small menu. Nothing in here is required
reading — the defaults are the right ones for studying.

### Show developer tools

Off by default. Turning it on adds a **Developer** group with:

- **Dump Page XML** — writes the current page's raw XML to the log folder. If you report a bug, this
  is the single most useful thing you can attach. **Read it first**: page XML can contain your file
  paths, your name, and the full text of anything on the page, including OCR'd text from images.
- **Survey Notebooks** — counts what kinds of content your notebooks contain. Useful to us, slow on
  a big notebook, and it does not change anything.
- **Icon Browser** — shows Office icons drawn by OneNote itself. It exists because which icons work
  varies per application, and this is the only reliable way to find out.

### Logging

RecallTape keeps a plain-text log of what each button did — no page content, ever, only actions and
errors. It is the first thing to look at when something misbehaves.

- **Write a log file** — on by default. Turn it off and nothing is written at all.
- **Log size limit** — 1 MB, 5 MB or 20 MB. When the log passes the limit it rolls over: the current
  file becomes `recalltape.log.1` and a fresh one starts. **Only one old file is kept**, so the most
  RecallTape will ever use is twice the limit. At the 1 MB default that is 2 MB, forever, no matter
  how long you use it.
- **Open log folder** — opens `%LOCALAPPDATA%\EWC3 Labs\RecallTape` in File Explorer.

Page dumps from **Dump Page XML** land in that same folder and are *not* rolled or deleted — they are
written only when you ask, so cleaning them up is yours to do.

## Studying with it

The workflow it's built for:

1. Tape the answers on a page — labels, definitions, drug names, the right-hand column of a table
2. Come back later and read the page
3. Try to recall what's under each bar **before** you peek
4. Reveal to check yourself — Ctrl+click a box, or drag-select over taped text
5. Put boxes back with another Ctrl+click; taped text re-hides on its own when you click away

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
  page clean, run **Remove All Tape** on it *before* you uninstall.

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
If it's a box, Ctrl+click it — if content appears, it's ours. If it's text, drag-select it: our tape
hides text that is still perfectly selectable and copyable. If neither does anything, it may be
ordinary OneNote highlighting, which we deliberately leave alone.

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
