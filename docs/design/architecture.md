# Architecture

> Why RecallTape is built as an engine with a host, rather than as a OneNote add-in that happens to
> mask things. Background and citations: [`../origins.md`](../origins.md).

---

## The decision that shapes everything

**`RecallTape.Core` must not know what OneNote is.**

```text
RecallTape.Core       Tape · TapeSet · RevealState · StudySession · Persistence
                      ↑ no COM. no page XML. no Office references. ever.
RecallTape.OneNote    PageAdapter · SelectionAdapter · OneNoteXml · OverlayRenderer
RecallTape.UI         Ribbon · Hotkeys · StudyController

                      ── later, and only possible because of the line above ──
RecallTape.Anki   ·   RecallTape.PDF   ·   RecallTape.Web
```

Masking-for-recall is a general idea: *a region of a document, hidden, revealable, with study state
attached.* OneNote is the first place we need it, not the only place it makes sense. Keeping the model
host-free is what lets an Anki exporter, a PDF host, or an Office.js web edition reuse the engine
instead of reimplementing it.

**This is the project's A.B.S. line** (see `AGENTS.md`). When `Core` starts importing OneNote types —
and something will always want to, usually for a "quick" reason — that is the bug. Fix it then, not
"later," because later is when there are three hosts and the leak is load-bearing.

---

## Why desktop C# and not Office.js

Decided in `origins.md` with citations. Short version:

| | OneNote for Web | **OneNote Windows desktop** |
| --- | --- | --- |
| API | Office.js | COM / desktop API |
| Microsoft's docs target | this one, explicitly | still documented, older |
| Marketplace | yes | no useful one |
| Reaches your **ink** | not meaningfully | **yes** |
| Where lecture notes actually happen | no | **yes** |

The modern, better-documented, marketplace-friendly API is the one that cannot serve the actual user:
a medical student taking pen notes on a Surface in a lecture hall. So we take the older API and the
harder distribution story, because that's where the notes are.

**The foothold:** the desktop API returns the current page as XML *with the selection identified*, and
accepts a modified page back via `UpdatePageContent`. That single capability is what makes the whole
product possible.

## Standing on OneMore

[OneMore](https://github.com/stevencohn/OneMore) (Steven Cohn, open-source C#) already solved add-in
registration, ribbon integration, hotkeys, and page XML manipulation for desktop OneNote. Rebuilding
that plumbing to prove a point would be foolish.

**Unresolved:** fork it, depend on it, or use it as a reference and write a thinner host? Blocked on a
licence-compatibility check. Whatever we choose, it gets credited prominently — see `AGENTS.md`.

---

## The tape model (sketch)

```
Tape          one masked region: anchor into host content, original payload, reveal state
TapeSet       all tapes on a page/section, ordered for the study loop
RevealState   hidden · revealed · permanently removed
StudySession  a pass over a TapeSet — order, position, hits, misses
Persistence   where tape state lives across sessions
```

**The hard part is `Tape.anchor`.** It has to survive the user editing around it, OneNote syncing the
page, and the page being reopened later. A brittle anchor produces orphaned tape, which looks like data
loss even when it isn't. Expect to iterate here — and per house method, expect to tear the first
version out once the failure modes are visible.

**`Persistence` is an open question**, deliberately unresolved in `PLAN.md`: tape state inside the page
XML travels with the content and survives sync, but is fussier to write; state stored beside the page
is easy and risks orphaning. Decide with a spike, not with an argument.

## Non-destructive, always

The original content and its metadata are preserved on tape and restored on removal. Not "usually" —
always, including when a write fails half-way. Belts and suspenders on anything that rewrites a page:
these are somebody's notes, and there is no undo button that reaches back to last Tuesday.

## The study loop is keyboard-first

Not a compromise — a design conclusion from `origins.md`. Clicking individual tape strips is fiddly on
a tablet and slow everywhere. One key that reveals, then retapes and advances, turns a page of notes
into a loop you can run one-handed without looking down:

```
Space → answer → Space → answer → Space
```

Direct manipulation of individual strips, if it ever becomes achievable on the desktop canvas, is an
*addition* to this — never the thing it gets replaced by.
