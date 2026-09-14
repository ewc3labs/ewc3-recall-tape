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

AsOf: 2026-09-14 (states re-judged against the new legend)

## Current Focus

**Productization, then `v0.1.0`.** Creature features wait; the goal is software a stranger can
install, use and uninstall without being told anything.

1. `RT-37` — watch the log actually roll at its size limit; the only unobserved part of the settings
   work
2. `RT-21` + `RT-22` — signing and a real installer; the zip-and-script story is the weakest part
   left
3. `RT-08` — progress feedback, so a long Survey does not look frozen
4. Then bump to `v0.1.0` and go back to the product: `RT-15` + `RT-16`

Parked until after 0.1.0, deliberately: `RT-07` → `RT-20` (anchoring), `RT-26` → `RT-25` (reporter).

## ID Prefixes

**Read this before minting an ID.** It sits above the tables because it is an input to writing one,
not a summary of them.

| Prefix | Scope | Owner | Last Used | Series |
| --- | --- | --- | --- | --- |
| RT | global | ewc3-recall-tape | <!--ewc3:lastRT-->RT-45<!--/ewc3:lastRT--> | RecallTape feature slices and fixes |
| FIX | repo-local | ewc3-recall-tape | <!--ewc3:lastFIX-->FIX-0<!--/ewc3:lastFIX--> | small corrections not worth a slice |

**Last Used is derived** from the ID tables below by `ewc3-docs values`, and CI fails if it is
stale. That matters more than it sounds: this cell is the INPUT to minting an ID, not a summary of
one, so when it drifts the next person takes a number that is already taken and nothing complains.
EPQE's read `PQ-31` while `PQ-32` through `PQ-34` existed below it, which is what prompted
automating it. **Max, not a count** — counting rows agrees with the highest ID only while a series
is contiguous.

**A global prefix belongs to exactly one roadmap** across all of EWC3 Labs, so a reference points
somewhere unambiguous. See the [prefix registry][prefix-registry] in HQ.

**`FIX` is repo-local, and that is canon.** Every roadmap owns its own `FIX` series, so `FIX-3` here
is a different thing from `FIX-3` elsewhere — safe, because a fix is never referenced from outside
the repository it fixes.

`ewc3-docs series` enforces both rules: it fails if this roadmap uses a prefix it has not declared,
or if another roadmap claims the same global one.

## States

Read this **before** choosing a row's state. It sits above the tables for the same reason ID
Prefixes does: it is an input to writing a row, not a summary of them.

| | State | Means |
| --- | --- | --- |
| ⬜ | `planned` | minted, not started |
| 🟨 | `coded` | built, and never run against anything |
| 🟦 | `tested` | unit or smoke tested — probably OK, not proven |
| 🟩 | `proven` | met a **real environment** — `Status` must name *where* |
| ✅ | `done` | complete, and **not software**, so the ladder above does not apply |
| ⛔ | `blocked` | off the ladder — waiting on someone or something |
| ⏸️ | `deferred` | off the ladder — not now, **may return** |
| 🟥 | `cancelled` | off the ladder — decided against; `Status` says `reverted`, `refuted` or `retired`, and why |

**🟩 must name *where*.** `proven 2026-09-14 — work PC, P:\ share`, never a bare `proven`. A proof
is a claim *and the environment it held in*, so a later break somewhere else reads as a gap in scope
rather than a lie. **The where has to be the real environment for that kind of work:** CI is the
real environment for CI tooling, and only a test bench for a product feature.

**✅ is for work that could not be tested at all** — a reply sent, a pattern retired, a guide
written. If a row is software it belongs on the ladder, and ✅ on it is a mislabel.

**Never round up.** If the record does not show which rung a row reached, it is 🟨, with the doubt
in `Status`. **Keep 🟥 rows** — a refuted finding is the record of *why not*, and the next person to
report the same thing needs to find it.

**Est:** `S` an hour or two · `M` a session · `L` several sessions · `XL` a project in itself. Day
counts are fake precision on a project nobody is scheduling. Sizes are for spotting what is big.

`Status`: `pending` · `started YYYY-MM-DD` · `proven YYYY-MM-DD — {where}` ·
`done YYYY-MM-DD — {what done means}` · `deferred — {condition}` · `blocked on {who/what}` ·
`cancelled — {reverted|refuted|retired}: {why}` · `seen again YYYY-MM-DD`

## Delivery Index

**Rows are one line.** `Doc` pins a filename; anything wanting a paragraph wants a slice doc.

### Proven

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| RT-01 | 🟩 proven | Reach the OneNote API at all | M | — | proven 2026-08-12 — Wilson's PC, OneNote 16.0.20228 |
| RT-02 | 🟩 proven | Add-in loads and reads page XML | M | — | proven 2026-08-12 — Wilson's PC, OneNote, 79 pages of Paul's notebook |
| RT-03 | 🟩 proven | Cover a selection and uncover it | M | — | proven 2026-08-12 — Wilson's PC, OneNote, both anchor worlds |
| RT-11 | 🟩 proven | Click-to-peek on overlay tape | M | — | proven 2026-08-12 — Wilson's PC, OneNote, ~45ms |
| RT-13 | 🟩 proven | Release pipeline to a draft GitHub Release | S | — | proven 2026-08-12 — GitHub Actions; has cut v0.0.1–v0.0.3 since |
| RT-14 | 🟩 proven | Install from the release zip, start to finish | S | — | proven 2026-08-13 — Paul's Surface, a machine that did not build it |
| RT-24 | 🟩 proven | Free tape box, dragged onto anything | M | — | proven 2026-08-13 — Wilson's PC, OneNote; no selection needed |
| RT-27 | 🟩 proven | Remove one tape; Remove All asks first | S | — | proven 2026-08-13 — Wilson's PC, OneNote; Remove All on a 4-tape page |
| RT-28 | 🟩 proven | Peeked tape keeps a visible outline | S | — | proven 2026-08-13 — Wilson's PC, OneNote; drawn at real size |
| RT-30 | 🟩 proven | An installer a non-developer can actually run | M | — | proven 2026-08-13 — Paul's Surface, via RT-14's install; v0.0.1 zip has the .cmd |
| RT-23 | 🟩 proven | Verify the add-in loads on ARM64 | S | — | proven 2026-08-13 — Paul's Surface, Snapdragon X Elite |
| RT-33 | 🟩 proven | Identify tape by style shape, not a colour string | S | — | proven 2026-08-13 — Wilson's PC, OneNote, via RT-12's 5-tape pass |
| RT-39 | 🟩 proven | Validate ribbon XML against customui14.xsd | S | — | proven 2026-08-13 — Wilson's PC, real ribbon; xsd gone from default path by 2026-09-14 |

### The product — none of this exists yet

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| RT-44 | ⬜ planned | Study Mode: click tape to reveal, click away to carry on | L | — | BOXES ONLY — text cannot be clickable and hidden at once |
| RT-45 | ⬜ planned | Tape text that is already a hyperlink | M | — | box workaround documented; only worth it if users ask |
| RT-15 | ⬜ planned | Hotkeys, so tape is not ribbon-only | M | — | `Ctrl+Alt+T`; needs a keyboard hook |
| RT-16 | ⬜ planned | Study loop: reveal, retape, advance on one key | L | — | the whole point; needs RT-12 first |
| RT-17 | ⬜ planned | Reveal All / Retape All | S | — | Retape All needs remembered tape state |
| RT-18 | ⬜ planned | Confirm tape survives close and reopen | S | — | likely free — tape is page content |
| RT-19 | ⬜ planned | Settings: hotkey bindings, reveal style | M | — | pending RT-15 |

### Content types

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| RT-04 | 🟩 proven | Overlay a whole selected image or ink group | M | — | proven 2026-08-12 — Wilson's PC, OneNote; ink again 2026-08-13; sub-region is RT-06 |
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
| RT-12 | 🟩 proven | Re-read the page timestamp before every write | S | — | proven 2026-08-13 — Wilson's PC, OneNote, 5 tapes in one pass |
| RT-34 | 🟦 tested | Remove tape from the caret, not only a selection | S | — | tested 2026-08-13 — against a captured page dump; live use never recorded |
| RT-32 | ⬜ planned | Survive taping a page that is in sync conflict | M | — | API cannot see conflict state at all |
| RT-20 | ⬜ planned | Anchor tapes to an object, move them when it moves | M | — | needs RT-07's answer first |
| RT-07 | ⬜ planned | Does objectID survive a sync round-trip | S | — | now testable — 2 machines share a notebook |
| RT-08 | ⬜ planned | Progress feedback for long operations | S | — | Survey looks frozen on a big notebook |
| RT-37 | 🟨 coded | Bounded logs: roll at a size, keep one previous | S | — | only the size picker was smoke-tested 2026-08-13; the roll itself has never run |

### Shipping

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| RT-21 | ⬜ planned | Code signing, so SmartScreen stops warning | M | — | **cheaper than assumed.** Azure Trusted Signing is ~$10/month with no hardware token — verified against ScreenToGif's installer, whose cert is Microsoft-issued with a **3-day** lifetime, timestamped so it stays valid. Price is no longer the blocker; eligibility is. Also covers the add-in itself, which an org enforcing *Require Application Add-ins to be signed* will otherwise refuse to load at all — a harder failure than a SmartScreen prompt |
| RT-22 | ⬜ planned | MSI installer to replace zip + script | L | — | after RT-21; WiX headless, per-machine like OneMore |
| RT-35 | 🟦 tested | Install per-machine into Program Files | M | — | tested 2026-08-13 — Wilson's PC, ACLs and registry; not yet on a machine that did not build it |
| RT-31 | 🟦 tested | Guard every registry deletion | S | — | tested 2026-08-13 — refuses the 2 bad paths, allows the 5 good; no real uninstall recorded |
| RT-40 | ✅ done | A how-to a student can follow without us | S | HOW-TO.md | done 2026-08-13 — linked twice, ships in the zip; no student seen using it yet |
| RT-36 | 🟦 tested | Settings menu; developer tools off by default | M | — | tested 2026-08-13 — smoke in OneNote: dev tools toggle is instant; HKCU-backed |
| RT-38 | 🟦 tested | Ribbon laid out for reach, with icons that render | S | — | tested 2026-08-13 — seen in OneNote; icon browser found them; size='large' is ignored |
| RT-41 | ⬜ planned | Let users pick their own tape icon | S | — | the browser makes this small; whimsy belongs to them |
| RT-10 | ⬜ planned | Re-sequence PLAN.md phases to match real notes | S | — | flagged in PLAN.md, not yet done |

### Parked / Retired

| ID | State | Slice | Est | Doc | Status |
| --- | --- | --- | --- | --- | --- |
| RT-09 | ⏸️ deferred | The Assistant, as a click-through overlay window | XL | — | deferred — after the tool is good; name unresolved |
| RT-29 | 🟥 cancelled | Free the corner handles from aspect ratio | S | — | cancelled — refuted: no aspect property exists in the schema |
| RT-42 | 🟥 cancelled | Hyperlink taped TEXT so clicking it works too | M | — | cancelled — refuted 2026-08-13: OneNote strips link text colour, 8 recipes |
| RT-43 | 🟥 cancelled | Remove arms, then a click removes the tape | S | — | cancelled — reverted 2026-08-13: built, worked; its modal cost more than the no-op |

## Working Rules

- The roadmap is the single planning surface and owns the status of everything minted into it.
- New work arrives via [`RecallTape_Checkin_Punchlist.md`][recalltape-checkin] and is promoted here
  once the problem is understood well enough to describe with receipts.
- Backlogging = state `⏸️ deferred`, with the admission condition in `Status`.
- Technical detail lives in `docs/design/` and `docs/analysis/`; rows link, they do not duplicate.
- Load the `ewc3labs-project-roadmap` skill before structural edits.

## Related Documentation

- **[What OneNote page XML actually contains][what-onenote-page]** | [GitHub Link 🔗][github-link]
  - *The measured receipts behind most of the Content types section.*
- **[How OneMore talks to OneNote][how-onemore-talks-to]** | [GitHub Link 🔗][github-link-2]
  - *The reference implementation, and what we take from it.*
- **[PLAN.md](../PLAN.md)** | [GitHub Link 🔗][github-link-3]
  - *Phases and non-negotiables. RT-10 re-sequences it.*

[github-link]: https://github.com/ewc3labs/ewc3-recall-tape/blob/main/docs/analysis/onenote-page-xml-shapes.md
[github-link-2]: https://github.com/ewc3labs/ewc3-recall-tape/blob/main/docs/analysis/onemore-onenote-interaction.md
[github-link-3]: https://github.com/ewc3labs/ewc3-recall-tape/blob/main/docs/PLAN.md
[how-onemore-talks-to]: ../analysis/onemore-onenote-interaction.md
[prefix-registry]: https://github.com/ewc3labs/ewc3labs-hq
[recalltape-checkin]: RecallTape_Checkin_Punchlist.md
[what-onenote-page]: ../analysis/onenote-page-xml-shapes.md
