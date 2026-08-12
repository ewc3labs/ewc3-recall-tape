# Contributing to RecallTape

Pre-alpha — the structure is still landing, so code contributions aren't open yet. **Issues and ideas
are very welcome now**, especially from people who actually study from OneNote.

## The one rule that outranks everything

**Never damage someone's notes.** Taping is non-destructive and reversible, including when a page write
fails half-way. These are lecture notes people cannot recreate. If a change makes note loss possible,
it doesn't ship, however good the feature is.

## Before you build something

Read [`docs/origins.md`](docs/origins.md) and [`docs/PLAN.md`](docs/PLAN.md). A lot of decisions are
already made with reasoning attached — the platform choice, the keyboard-first study loop, the
host-agnostic engine. Arguing against them is fine; re-deriving them by accident wastes your evening.

## Ground rules

- **`RecallTape.Core` never learns what OneNote is.** No COM types or page XML above the `.OneNote`
  boundary. That line is what makes Anki/PDF/Web hosts possible later.
- **Commit messages carry the *why*.** The diff already shows what changed. Explain the reasoning and
  the alternative you rejected.
- **Tests must fail against the bug.** A regression test that passes before your fix is a false alibi.
  Check both directions.
- **Ink is not an edge case.** Typed-only is a fine first step; assuming it's the finish line is not.
- **No telemetry.** Notes are personal. Any analytics needs an explicit, documented, opt-in decision.
- **Credit prior art.** [OneMore](https://github.com/stevencohn/OneMore) is doing a lot of heavy lifting
  here; anything else borrowed gets named and linked too.

## Reporting bugs

Use the issue templates. **Say up front whether any note content was lost or altered** — that jumps the
queue past everything else.

Include the content type (typed / ink / dictated / image). Ink and typed text travel completely
different code paths, and "it didn't work" without that detail is usually unactionable.

## Licence

By contributing you agree your work is licensed under the [MIT Licence](LICENSE).
