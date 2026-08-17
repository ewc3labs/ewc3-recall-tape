# RecallTape — Plan

Status: **pre-alpha, Phase 0.** Nothing installable yet.

Source of truth for *why* any of this is shaped the way it is: [`origins.md`](origins.md).
Architecture: [`design/architecture.md`](design/architecture.md).

---

## The one-line goal

Let somebody study from their own OneNote notes without leaving OneNote or building flashcards.

## Non-negotiables

These constrain every decision below.

1. **Never damage someone's notes.** Taping is reversible and non-destructive; the original content
   and its metadata survive. These are medical school lecture notes. Losing them is the only truly
   unrecoverable failure this project has.
2. **Ink is not an edge case.** The target user is on a Surface with a pen. Typed-only is a
   legitimate *first* milestone, never the finished product.
3. **The joke is optional; the tool is not.** Mascot and animation ship **off by default** and are
   disableable forever in one click. A study aid that can't be made quiet at 2am is broken.
4. **No telemetry** without an explicit, documented, opt-in decision. Notes are personal.

---

> **Phase order under revision (2026-08-12).** A survey of 79 pages of real notes found **892** typed
> text elements against **11,685** ink and **1,131** images. Phase 1 is typed-text taping and Phase 2 is
> ink and regions -- backwards for the actual user, who said so independently: *"I'm often covering
> pictures and printouts."* Evidence:
> [`analysis/onenote-page-xml-shapes.md`][analysis-onenote]. Re-sequencing is slice
> `RT-10` in [`project/RecallTape_Development_Roadmap.md`][project-recalltape].

## Phase 0 — decide and prove (we are here)

The platform research is **done** — see `origins.md`. What remains is proving it with running code
before designing on top of it.

- [ ] **Spike:** attach to desktop OneNote, fetch current page XML, confirm the selection is
      identifiable in it, write a trivially modified page back via `UpdatePageContent`.
- [ ] **Spike:** determine how a taped region is best represented — modify the content's styling in
      place, or place a cover object over it? Decide for typed text first.
- [ ] **Spike:** find out what selected **ink** looks like in the page XML, and whether it can be
      masked by the same mechanism. This is the highest-risk unknown in the project.
- [ ] Decide: fork [OneMore][onemore], take it as a dependency, or use it as a reference
      implementation and write a thin host of our own. *(Licence compatibility check required before
      forking.)*
- [ ] Confirm the install/trust story for a COM add-in on a normal user's machine — signing, and
      what SmartScreen does to an unsigned installer.

**Exit criterion:** a hotkey in real OneNote that visibly covers a real selection and uncovers it
again, without eating the content. Ugly is fine. That's the whole bar.

## Phase 1 — MVP

- [ ] `Ctrl+Alt+T` tapes the current selection
- [ ] Original content + metadata preserved and restorable
- [ ] Study loop: reveal → next press retapes and advances to the next tape
- [ ] **Reveal All** / **Retape All** / **Remove Tape**
- [ ] Ribbon buttons for everything that has a hotkey
- [ ] Tape state survives closing and reopening the page
- [ ] Settings: hotkey bindings, reveal style
- [ ] Signed installer published via GitHub Releases

## Phase 2 — the part that makes it a learning tool

Cloning Goodnotes' tape gets parity. This is where it becomes something Goodnotes doesn't have.

- [ ] Mask **handwriting**
- [ ] Mask **regions of diagrams** and labels on anatomy images
- [ ] **Tape all handwriting on this page** (bulk operations — hand-taping fifty terms defeats the
      point)
- [ ] Randomised reveal order
- [ ] **Study mode** — a focused review pass over a page or section
- [ ] Hit/miss statistics
- [ ] Spaced-repetition metadata
- [ ] **Export taped items to Anki**

## Phase 3 — other hosts

Only meaningful because `RecallTape.Core` never learned what OneNote is.

- [ ] `RecallTape.Anki`
- [ ] `RecallTape.PDF`
- [ ] `RecallTape.Web` — an Office.js edition for OneNote on the web, publishable to Microsoft
      Marketplace. Different host, different distribution, same Core.

---

## Deliberately out of scope (for now)

- **OneNote for Mac / iPad.** Different extensibility story entirely.
- **Real-time collaborative taping.** No.
- **Cloud sync of tape state.** Tape lives with the page; if that proves impossible, revisit — but
  don't start by building a service.
- **Clicking individual tape strips as the primary interaction.** See `origins.md`: keyboard-driven
  sequential reveal is a *better* study loop and a *cheaper* build. If direct manipulation becomes
  achievable later, it's an addition, not the target.

## Open questions

- **Mascot name.** Wilson has said *Forky*; `origins.md` prefers **Tiney** / **Rippy** / **Sal** /
  **Dork** on the grounds that Forky is generic and strongly associated with a Pixar character.
  Unresolved — worth settling before it reaches code, assets, or store metadata.
- **OneMore: fork, depend, or reference?** Blocked on the licence check above.
- **Where does tape state actually live** — inside the page XML, or beside it? Page XML keeps it
  with the content and surviving sync; beside it is easier but risks orphaning.

[analysis-onenote]: analysis/onenote-page-xml-shapes.md
[onemore]: https://github.com/stevencohn/OneMore
[project-recalltape]: project/RecallTape_Development_Roadmap.md
