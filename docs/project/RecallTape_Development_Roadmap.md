<div align="center">
<table>
<tr>
<td style="padding: 0 40px; vertical-align: middle;">
<img src="../../images/EWC3LabsLogo-blue-128x128.png" alt="EWC3 Labs" width="72" height="72">
</td>
<td style="vertical-align: middle;">
<h1 style="margin: 0;">RecallTape</h1>
<h3 style="margin: 5px 0;"><strong>Development Roadmap — where we are, and what is next</strong></h3>
</td>
</tr>
</table>
</div>

---

AsOf: 2026-08-13 (end of session)

## Current Focus

**Productization, then `v0.1.0`.** Creature features wait; the goal is software a stranger can
install, use and uninstall without being told anything.

1. `RT-36` `RT-37` `RT-38` — smoke the settings menu, log rolling and new ribbon on a real page
2. `RT-21` + `RT-22` — signing and a real installer; the zip-and-script story is the weakest part left
3. `RT-08` — progress feedback, so a long Survey does not look frozen
4. Then bump to `v0.1.0` and go back to the product: `RT-15` + `RT-16`

Parked until after 0.1.0, deliberately: `RT-07` → `RT-20` (anchoring), `RT-26` → `RT-25` (reporter).

## Last Numbers

Read this **before** minting an ID. It sits above the tables because it is an input to writing one,
not a summary of them.

| Series | Last Num | Series Description |
| --- | --- | --- |
| RT | RT-44 | RecallTape feature slices and fixes |

`Last Num` is a cache over the Delivery Index, not a second source of truth — the IDs in the tables are
authoritative. When minting, take the next number **and** confirm it is unused across every sub-table,
then bump this cell in the same edit.

## Delivery Index

**Rows are one line.** `Doc` pins a filename; anything wanting a paragraph wants a slice doc.

States: `⬜ planned` · `🟨 coded` — built and deployed, nobody has used it yet · `💨 proven` — someone
used it and it worked. `🟨` is not a lesser `💨`; it is an admission that we do not know yet.

### Proven

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| RT-01 | 💨 proven | Reach the OneNote API at all | M | — | proven 2026-08-12 |
| RT-02 | 💨 proven | Add-in loads and reads page XML | M | — | proven 2026-08-12 |
| RT-03 | 💨 proven | Cover a selection and uncover it | M | — | proven 2026-08-12 |
| RT-11 | 💨 proven | Click-to-peek on overlay tape | M | — | proven 2026-08-12 — ~45ms |
| RT-13 | 💨 proven | Release pipeline to a draft GitHub Release | S | — | shipped 2026-08-13 — v0.0.1 public |
| RT-14 | 💨 proven | Install from the release zip, start to finish | S | — | proven 2026-08-13 — Paul, not our machine |
| RT-24 | 💨 proven | Free tape box, dragged onto anything | M | — | proven 2026-08-13 — no selection needed |
| RT-27 | 💨 proven | Remove one tape; Remove All asks first | S | — | proven 2026-08-13 |
| RT-28 | 💨 proven | Peeked tape keeps a visible outline | S | — | proven 2026-08-13 — drawn at real size |
| RT-30 | 💨 proven | An installer a non-developer can actually run | M | — | proven 2026-08-13 — double-click .cmd |
| RT-31 | 💨 proven | Guard every registry deletion | S | — | proven 2026-08-13 — refuses non-RecallTape keys |
| RT-23 | 💨 proven | Verify the add-in loads on ARM64 | S | — | proven 2026-08-13 — Snapdragon X Elite |
| RT-33 | 💨 proven | Identify tape by style shape, not a colour string | S | — | proven 2026-08-13 — OneNote rewrites our CSS |
| RT-34 | 💨 proven | Remove tape from the caret, not only a selection | S | — | proven 2026-08-13 — caret is a zero-length T |
| RT-35 | 💨 proven | Install per-machine into Program Files | M | — | proven 2026-08-13 — Users:(RX), nobody can swap the DLL |
| RT-39 | 💨 proven | Validate ribbon XML against customui14.xsd | S | — | proven 2026-08-13 — caught the bug that deleted the tab |
| RT-40 | 💨 proven | A how-to a student can follow without us | S | HOW-TO.md | shipped 2026-08-13 — linked twice, ships in the zip |

### The product — none of this exists yet

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| RT-42 | ⬜ planned | Hyperlink taped TEXT so clicking it works too | M | — | unlocks text peek AND uniform click; gates RT-44 |
| RT-44 | ⬜ planned | Study Mode: click tape to reveal, click away to carry on | L | — | on hyperlinks, NOT a mouse hook — no viewport exists |
| RT-15 | ⬜ planned | Hotkeys, so tape is not ribbon-only | M | — | `Ctrl+Alt+T`; needs a keyboard hook |
| RT-16 | ⬜ planned | Study loop: reveal, retape, advance on one key | L | — | the whole point; needs RT-12 first |
| RT-17 | ⬜ planned | Reveal All / Retape All | S | — | Retape All needs remembered tape state |
| RT-18 | ⬜ planned | Confirm tape survives close and reopen | S | — | likely free — tape is page content |
| RT-19 | ⬜ planned | Settings: hotkey bindings, reveal style | M | — | pending RT-15 |

### Content types

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| RT-04 | 🟦 coded | Overlay a whole selected image or ink group | M | — | works; sub-region is RT-06 |
| RT-05 | ⬜ planned | Cluster InkDrawing boxes to tape handwriting | L | — | zero InkWord in 11,708 strokes — we segment |
| RT-06 | ⬜ planned | Occlude part of an image using OCRToken boxes | L | — | unblocked — printout slides carry OCRData |

### Reporting and privacy

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| RT-26 | ⬜ planned | Scrub page XML of identity and content | M | — | blocks RT-25 — paths, names, CIDs, OCR text |
| RT-25 | ⬜ planned | Report an issue from inside OneNote | L | — | needs RT-26; consent per attachment |

### Robustness

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| RT-12 | 💨 proven | Re-read the page timestamp before every write | S | — | proven 2026-08-13 — 5 tapes, one pass |
| RT-32 | ⬜ planned | Survive taping a page that is in sync conflict | M | — | API cannot see conflict state at all |
| RT-20 | ⬜ planned | Anchor tapes to an object, move them when it moves | M | — | blocked in practice by RT-07 |
| RT-07 | ⬜ planned | Does objectID survive a sync round-trip | S | — | now testable — 2 machines share a notebook |

| RT-08 | ⬜ planned | Progress feedback for long operations | S | — | Survey looks frozen on a big notebook |
| RT-37 | 🟨 coded | Bounded logs: roll at a size, keep one previous | S | — | 1/5/20 MB, default 1; logging can be off |

### Shipping

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| RT-21 | ⬜ planned | Code signing, so SmartScreen stops warning | M | — | needs a certificate; costs money |
| RT-22 | ⬜ planned | MSI installer to replace zip + script | L | — | after RT-21; WiX headless, per-machine like OneMore |
| RT-36 | 🟨 coded | Settings menu; developer tools off by default | M | — | ribbon menu, not a dialog; HKCU-backed |
| RT-38 | 🟨 coded | Ribbon laid out for reach, with icons that render | S | — | Tape large and rightmost; icon browser found them |
| RT-41 | ⬜ planned | Let users pick their own tape icon | S | — | the browser makes this small; whimsy belongs to them |
| RT-43 | ⏸ retired | Remove arms, then a click removes the tape | S | — | built, worked, backed out 2026-08-13 — its modal cost more than the no-op |
| RT-10 | ⬜ planned | Re-sequence PLAN.md phases to match real notes | S | — | flagged in PLAN.md, not yet done |

### Parked / Retired

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| RT-09 | ⏸️ deferred | The Assistant, as a click-through overlay window | XL | — | deferred — after the tool is good; name unresolved |
| RT-29 | 🟥 cancelled | Free the corner handles from aspect ratio | S | — | not ours — no aspect property exists in the schema |

**States:** ⬜ `planned` · 🟦 `coded` · 💨 `proven` (actually run and observed) · 🟩 `shipped` ·
🟧 `blocked` · ⏸️ `deferred` · 🟥 `cancelled`

**Est:** `S` an hour or two · `M` a session · `L` several sessions · `XL` a project in itself.
Day counts were fake precision — five slices landed in a single day while wearing "1d" and "3d"
labels that meant nothing. Sizes are for spotting what is big, not for scheduling.

`Status`: `pending` · `started YYYY-MM-DD` · `proven YYYY-MM-DD` · `shipped YYYY-MM-DD` ·
`deferred — {condition}` · `blocked on {who/what}` · `seen again YYYY-MM-DD`

## Working Rules

- The roadmap is the single planning surface and owns the status of everything minted into it.
- New work arrives via [`RecallTape_Checkin_Punchlist.md`](RecallTape_Checkin_Punchlist.md) and is
  promoted here once the problem is understood well enough to describe with receipts.
- Backlogging = state `⏸️ deferred`, with the admission condition in `Status`.
- Technical detail lives in `docs/design/` and `docs/analysis/`; rows link, they do not duplicate.
- Load the `ewc3labs-project-roadmap` skill before structural edits.

## Related Documentation

- **[What OneNote page XML actually contains](../analysis/onenote-page-xml-shapes.md)** | [GitHub Link 🔗](https://github.com/ewc3labs/ewc3-recall-tape/blob/main/docs/analysis/onenote-page-xml-shapes.md)
  - *The measured receipts behind most of the Content types section.*
- **[How OneMore talks to OneNote](../analysis/onemore-onenote-interaction.md)** | [GitHub Link 🔗](https://github.com/ewc3labs/ewc3-recall-tape/blob/main/docs/analysis/onemore-onenote-interaction.md)
  - *The reference implementation, and what we take from it.*
- **[PLAN.md](../PLAN.md)** | [GitHub Link 🔗](https://github.com/ewc3labs/ewc3-recall-tape/blob/main/docs/PLAN.md)
  - *Phases and non-negotiables. RT-10 re-sequences it.*
