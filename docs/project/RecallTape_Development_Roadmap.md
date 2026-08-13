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

AsOf: 2026-08-12

## Current Focus

1. `RT-14` — next; blocks publishing v0.0.1 to Paul
2. `RT-12` — then; a study loop cannot sit on writes that fail without retrying
3. `RT-15` + `RT-16` — the actual product: a hotkey and a loop, not a ribbon button
4. `RT-23` — blocked on Paul; only his Surface can prove the ARM64 work

## Last Numbers

Read this **before** minting an ID. It sits above the tables because it is an input to writing one,
not a summary of them.

| Series | Last Num | Series Description |
| --- | --- | --- |
| RT | RT-23 | RecallTape feature slices and fixes |

`Last Num` is a cache over the Delivery Index, not a second source of truth — the IDs in the tables are
authoritative. When minting, take the next number **and** confirm it is unused across every sub-table,
then bump this cell in the same edit.

## Delivery Index

**Rows are one line.** `Doc` pins a filename; anything wanting a paragraph wants a slice doc.

### Proven

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| RT-01 | 💨 proven | Reach the OneNote API at all | M | — | proven 2026-08-12 |
| RT-02 | 💨 proven | Add-in loads and reads page XML | M | — | proven 2026-08-12 |
| RT-03 | 💨 proven | Cover a selection and uncover it | M | — | proven 2026-08-12 |
| RT-11 | 💨 proven | Click-to-peek on overlay tape | M | — | proven 2026-08-12 — ~45ms |
| RT-13 | 💨 proven | Release pipeline to a draft GitHub Release | S | — | proven 2026-08-12 — draft, unpublished |

### The product — none of this exists yet

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
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
| RT-06 | ⬜ planned | Occlude part of an image using OCRToken boxes | L | — | 92,558 tokens available; Paul's core case |

### Robustness

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| RT-12 | ⬜ planned | Re-read and retry on stale-timestamp writes | S | — | 0x80042010 seen live |
| RT-20 | ⬜ planned | Reconcile overlays whose anchor content moved | M | — | drag the content, the tape stays put |
| RT-07 | 🟧 blocked | Does objectID survive a sync round-trip | S | — | blocked on Paul — needs 2 machines |
| RT-08 | ⬜ planned | Progress feedback for long operations | S | — | Survey looks frozen on a big notebook |

### Shipping

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| RT-14 | 🟦 coded | Install from the release zip, start to finish | S | — | written, never run — blocks publishing |
| RT-23 | 🟧 blocked | Verify the add-in loads on ARM64 | S | — | blocked on Paul's Surface |
| RT-21 | ⬜ planned | Code signing, so SmartScreen stops warning | M | — | needs a certificate; costs money |
| RT-22 | ⬜ planned | MSI installer to replace zip + script | L | — | after RT-21; WiX builds headless |
| RT-10 | ⬜ planned | Re-sequence PLAN.md phases to match real notes | S | — | flagged in PLAN.md, not yet done |

### Parked / Retired

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| RT-09 | ⏸️ deferred | The Assistant, as a click-through overlay window | XL | — | deferred — after the tool is good; name unresolved |

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
