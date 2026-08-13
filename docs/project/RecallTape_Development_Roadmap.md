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

1. `RT-12` — next; writes fail on a stale timestamp with no retry, and the study loop will hit it hard
2. `RT-04` — region overlay over images and printouts, Paul's primary case
3. `RT-07` — blocked on Paul; needs two machines and one shared notebook, both of which now exist

## Delivery Index

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --: | --- | --- |
| RT-01 | 💨 proven | Reach the OneNote API at all | 1d | — | proven 2026-08-12 |
| RT-02 | 💨 proven | Add-in loads, reads page XML | 1d | — | proven 2026-08-12 |
| RT-03 | 💨 proven | Cover a selection and uncover it | 2d | — | proven 2026-08-12 — Phase 0 exit criterion met |
| RT-04 | ⬜ planned | Region overlay over images and printouts | 3d | — | Paul's primary use case |
| RT-05 | ⬜ planned | Cluster InkDrawing boxes to tape handwriting | 3d | — | confirmed 2026-08-12 — zero InkWord in 11,708 strokes |
| RT-06 | ⬜ planned | Image occlusion driven by OCRToken boxes | 3d | — | pending RT-04 |
| RT-07 | 🟧 blocked | Does objectID survive a sync round-trip | 1d | — | blocked on Paul — needs 2 machines, 1 notebook |
| RT-08 | ⬜ planned | Progress feedback for long operations | 1d | — | Survey looks frozen on a big notebook |
| RT-09 | ⏸️ deferred | The Assistant, as a click-through overlay window | ? | — | deferred — after the tool works; name unresolved |
| RT-10 | ⬜ planned | Re-sequence PLAN.md phases to match real notes | 0.5d | — | typed text is 892 of 13,708 elements |
| RT-11 | 💨 proven | Click-to-peek on overlay tape via Image hyperlink | 2d | — | proven 2026-08-12 — ~45ms per toggle |
| RT-12 | ⬜ planned | Re-read and retry on stale-timestamp writes | 1d | — | 0x80042010 seen live; study loop writes fast |

**States:** ⬜ `planned` · 🟦 `coded` · 💨 `proven` (actually run, not just built) · 🟩 `shipped` ·
🟧 `blocked` · ⏸️ `deferred` · 🟥 `cancelled`

`Status`: `pending` · `started YYYY-MM-DD` · `shipped YYYY-MM-DD` · `deferred — {condition}` ·
`blocked on {who/what}` · `seen again YYYY-MM-DD`

## Last Numbers

| Series | Last Num | Series Description |
| --- | --- | --- |
| RT | RT-12 | RecallTape feature slices and fixes |

## Working Rules

- The roadmap is the single planning surface and owns the status of everything minted into it.
- New work arrives via [`RecallTape_Checkin_Punchlist.md`](RecallTape_Checkin_Punchlist.md) and is
  promoted here once the problem is understood well enough to describe with receipts.
- Backlogging = state `⏸️ deferred`, with the admission condition in `Status`.
- Technical detail lives in `docs/design/` and `docs/analysis/`; rows link, they do not duplicate.
- Load the `ewc3labs-project-roadmap` skill before structural edits.

## Related Documentation

- **[What OneNote page XML actually contains](../analysis/onenote-page-xml-shapes.md)** | [GitHub Link 🔗](https://github.com/ewc3labs/ewc3-recall-tape/blob/main/docs/analysis/onenote-page-xml-shapes.md)
  - *The measured receipts behind RT-03 through RT-06.*
- **[How OneMore talks to OneNote](../analysis/onemore-onenote-interaction.md)** | [GitHub Link 🔗](https://github.com/ewc3labs/ewc3-recall-tape/blob/main/docs/analysis/onemore-onenote-interaction.md)
  - *The reference implementation, and what we take from it.*
- **[PLAN.md](../PLAN.md)** | [GitHub Link 🔗](https://github.com/ewc3labs/ewc3-recall-tape/blob/main/docs/PLAN.md)
  - *Phases and non-negotiables. RT-10 re-sequences it.*
