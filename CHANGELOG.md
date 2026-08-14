# Changelog

All notable changes to RecallTape are documented here.

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/);
versioning follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.0.3] - 2026-08-13

A day of using v0.0.2 on real lecture notes. Everything below was found by pressing a
button, not by reading the code — including three separate bugs that all traced back to
one wrong assumption about a colour.

### Fixed

- **Tape would not come off.** Tape identity was the string `#1F1F1E`, and OneNote
  rewrites inline CSS on save — dropping `color:` and snapping that value to the named
  colour `black`. Freshly applied tape matched and came off; the same tape after a save
  did not, so you selected a solid black bar and were told there was no tape there. Worse,
  `#1F1F1E` was never distinctive: it is OneNote's own default body colour and appears on
  ordinary untaped paragraphs. Tape is now recognised by the *shape* of what survives — a
  span setting both `background` and `mso-highlight` dark — which also still finds tape
  from older versions.
- **Putting the cursor in taped text did nothing.** A caret selects no text, so a removal
  path that required a selection could never fire from a click — and you cannot see what
  you are selecting when it is black on black. OneNote marks a caret as a *zero-length*
  text run, which is now what we look for. It also splits the run at that point, so one
  tape comes back as two; removal covers the whole paragraph rather than stripping half a
  tape.
- **Coloured text showed straight through tape.** OneNote flattens nested spans into one
  run, and where our wrapper and the author's own colour disagreed, the author's won. The
  tape is now applied to the innermost text, so it wins instead — and the author's
  formatting is left wrapped around it, untouched, and restored exactly on removal.
- **Remove All needed two presses on a mixed page.** Every write bumps the page timestamp,
  and we passed one captured up front, so the second write of any batch failed the
  concurrency check. Counting was right, which is why it looked like the two kinds of tape
  were fighting. Each write now re-reads first, and one failure no longer abandons the
  rest of the page.
- **Taping a whole text box taped it twice** — a box over the container *and* every run
  inside it, leaving black text behind when the box came off. A whole container now means
  cover the box; text inside means cover those words.
- **About left OneNote unusable.** The dialog had no owner, so it opened behind OneNote in
  a background process where it could not be seen or reached — every click just beeped.
  Closing OneNote did not help, because the dialog kept our add-in alive and holding its
  own files. Dialogs now belong to OneNote's window.

### Added

- **Settings menu.** Developer tools are off by default — diagnostics are no use while
  studying. Logging can be turned off, and the log now **rolls at a size you choose**
  (1/5/20 MB, keeping one previous file), so it can never quietly grow without bound.
- **Ribbon laid out for reach.** The RecallTape tab sits at the far right, so the pointer
  arrives at the right-hand end — and **Apply Tape** is now the rightmost button rather
  than the leftmost, with **Remove All Tape** furthest from it.
- **HOW-TO.md**, shipped in the zip as well as on the repo page.
- **Icon Browser** and **Tape Lab** in the developer tools: Office icons and tape recipes
  rendered by OneNote itself, because what actually works varies by application and
  guessing has been expensive.

### Changed

- **Installs to `C:\Program Files\EWC3 Labs\RecallTape`** rather than wherever the zip was
  extracted. A folder created off `C:\` inherits *Authenticated Users: Modify*, meaning any
  signed-in user could replace the add-in OneNote loads. Program Files is read-only to
  non-admins. Extract the zip anywhere now and delete it afterwards.
- **The installer offers to close OneNote for you**, waits for it to actually exit, and
  releases the background process that holds our files — rather than telling you to sort
  it out.
- Logs and page dumps moved to `%LOCALAPPDATA%\EWC3 Labs\RecallTape`, out of the temp
  folder Windows empties.

### Known limitations

- **Hyperlinks cannot be taped as text.** OneNote reserves link colour and repaints it for
  contrast against whatever is behind it, so a black bar gets white words drawn on top.
  Eight different recipes were tested; it wins all of them. Cover links with a tape box
  instead — one box covers a whole list and peeks with Ctrl+click like any other.

## [0.0.2] - 2026-08-13

Everything here came from two people using v0.0.1 on real notes within hours of it
shipping. None of it was on the roadmap that morning.

### Added

- **Free tape box.** Press **Tape** with nothing selected and you get a draggable,
  resizable box, placed on the page and scrolled into view. This is the answer to
  content OneNote refuses to hand over: a page background image is
  `backgroundImage="true"` and is **never** marked as selected, no matter where you
  click, so selection-driven taping can never reach it. Rather than fight for a
  selection that will not come, place an object and get out of the way — our tape is a
  `one:Image`, so OneNote already gives it move and resize handles for free.
- **Remove Tape** now removes only what is selected — click a tape strip, or click
  inside taped text, then press it.
- **Remove All** is separate, counts first, and asks before clearing a page. It defaults
  to *No*, because a stray Enter should not undo a study session.
- Tooltips on both removal buttons, since "Remove Tape" and "Remove All" differ in a way
  worth stating.

### Changed

- **A peeked tape keeps a visible grey outline.** Previously it went fully transparent,
  which made it invisible *and* unfindable — with twenty tapes on a page, revealing one
  and scrolling away meant sweeping the page hunting for the hyperlink cursor. The
  outline is generated at the box's real pixel size at the moment you peek it, because a
  border baked into the 8×8 source PNG would stretch to 30px on a 240×80 box.
- Both removal scripts now route every registry deletion through a guard that refuses any
  key not naming RecallTape. Cheap insurance on the one operation that cannot be undone.

### Notes

- Alpha. Installable, unsigned, and actively changing.
- Corner resize handles keep their aspect ratio, and that is not something an add-in can
  change — `one:Size` exposes only width, height and `isSetByUser`, with no aspect
  property anywhere in the schema. For an arbitrary rectangle today: scribble the shape
  in ink, select it, and tape the ink.

## [0.0.1] - 2026-08-13

First public release. Alpha: it works, it is not finished, and it has not met a real exam week.

### Added

- Repository scaffolding: README, AGENTS.md, license, CI, issue templates, docs structure.
- `docs/origins.md` — the conversation that created the project, preserved verbatim.
- `docs/PLAN.md` — MVP scope and phased roadmap.
- `docs/design/architecture.md` — host-agnostic Core with OneNote as first host.
- `docs/design/mascot.md` — the reveal animation specification.
- `src/RecallTape.OneNote/` — first add-in build (net48): hand-declared COM interop and a spike
  command that dumps page XML. Not yet functional; see Known Issues.
- `tools/register.ps1` · `tools/unregister.ps1` — reproducible per-user dev registration.
- `docs/analysis/onemore-onenote-interaction.md` — how OneMore reaches OneNote, with diagrams.
- `docs/RAG_Sessions/` — durable shared session memory, plus `README_RAG_SYSTEM.md`.

- **Tape and peek.** `Tape` covers the selection, `Remove All Tape` clears the page, and Ctrl+clicking an
  overlay toggles it between covered and revealed (~45ms). Text and positioned content use different
  mechanisms, because the XML has two anchor worlds.
- `RecallTape.ProtocolHandler.exe` plus a named-pipe command service — the `recalltape://` chain that
  makes tape clickable, since OneNote will not serve an external process.
- `SurveyNotebooks` command — walks every open notebook and reports content-type totals, structural
  skeletons, and `recognizedText` samples.
- `docs/analysis/onenote-page-xml-shapes.md` — what OneNote page XML actually contains, measured.
- `docs/project/` — roadmap and check-in punchlist (PMO-light, per the `ewc3labs-project-roadmap` skill).
- `.editorconfig`, and `.gitattributes` replaced with the EWC3 Labs LF-everywhere policy.
- README Support section — Buy Me a Coffee, plus the free ways to help. (The sidebar Sponsor button already comes from the org-level `ewc3labs/.github` funding file; no repo-level override needed.)
- GitHub About filled in: 14 discovery topics and an SEO-leading description.

### Changed

- Target framework fixed at **`net48`**, not `net8.0`. It is what the reference implementation uses and
  what is already installed on every Windows 10/11 machine — no runtime install for end users. CI's
  `dotnet-version: 8.0.x` is the build *driver*, not the target.
- `docs/design/architecture.md` — documents the `dllhost` COM surrogate hosting model, the build/runtime
  decisions, and an explicit rejection of `UpdatePageContent(..., force: true)` on concurrency grounds.

- **Reversed the hand-declared-interop decision.** The add-in now references the real `Extensibility`
  and `Microsoft.Office.Interop.OneNote` PIAs. Hand-declared COM interfaces — correct IID, dual layout,
  typelib method order, explicit DispIds, `[TypeIdentifier]` — are constructed by OneNote and then
  never called. `dynamic` fails separately, on `IDispatch::GetTypeInfo`.
- Application object is acquired per operation via `new Application()` and released immediately, rather
  than caching the one passed to `OnConnection` (which prevents OneNote shutting down).
- `docs/PLAN.md` flagged: phase order is inverted relative to real notes (892 typed elements vs 11,685
  ink and 1,131 images). Re-sequencing tracked as `RT-10`.

### Proven

- **The add-in reads real page XML.** Full lifecycle (ctor, `OnConnection`, `GetCustomUI`,
  `OnStartupComplete`, `OnDisconnection`) plus `GetPageContent` and `GetHierarchy`, verified against a
  synthetic page and 79 pages of a real medical-school notebook.
- **The platform bet holds.** `CoCreateInstance` returns a dead object to an external process and a live
  one from inside a OneNote-spawned surrogate — the gate is on the caller, not the API.

### Resolved

- **OneMore is MPL-2.0**, RecallTape is MIT. Closes the fork-vs-depend-vs-reference question in
  `docs/PLAN.md`: **reference implementation only**, since MPL-2.0 file-level copyleft would follow any
  forked file permanently.

### Known Issues

- Corner resize handles keep their aspect ratio; OneNote owns that behaviour and the schema has no aspect property.
- **Zero `InkWord` in 11,685 real ink elements**, so `recognizedText` is unavailable and handwriting
  masking must cluster `InkDrawing` bounding boxes. Unclear whether that is the build, a setting, or an
  artifact of the notebook having been copied.
- Long operations give no progress feedback and look frozen (`RT-08`).
- The COM surrogate holds the built DLL open; OneNote must be closed before a rebuild.

### Notes

- Alpha. Installable, unsigned, and actively changing.
