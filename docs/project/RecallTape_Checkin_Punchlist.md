<div align="center">
<table>
<tr>
<td style="padding: 0 40px; vertical-align: middle;">
<img src="../../images/EWC3LabsLogo-blue-128x128.png" alt="EWC3 Labs" width="72" height="72">
</td>
<td style="vertical-align: middle;">
<h1 style="margin: 0;">RecallTape</h1>
<h3 style="margin: 5px 0;"><strong>Check-in Punchlist — stuff we noticed, and what is still open</strong></h3>
</td>
</tr>
</table>
</div>

---

Started: 2026-08-12

## How this works

Things we saw and need to address **sometime**. No evidence required to write one down — that is the
point. An item can sit here indefinitely without guilt.

An item is promoted to a roadmap slice **once someone has done enough analysis to document the problem
with receipts**. Until then it stays an unchecked box.

**Unchecked items are "what's still open."** There is no separate status file.

| Mark | Means |
| --- | --- |
| `- [ ]` | open — noticed, not yet understood |
| `- [~]` | being looked at |
| `- [x]` | minted into the roadmap, or delivered |

Checked items name their slice: `[RT-nn] minted in roadmap to address`. Before minting anything new,
search the roadmap **including Parked / Retired** — the same thing gets noticed months apart.

---

## Check-in — 2026-08-12 (Paul's requirements, over Telegram)

> Context: Wilson shared the repo mid-build. Paul — the actual user, the one who asked the original
> question — reacted with requirements before we had anything to show him.

**Decisions (Wilson + Paul, 2026-08-12):**

- Region/shape overlay is the primary interaction, not a Phase 2 nicety. Confirmed by the notebook
  survey: 892 typed elements against 11,685 ink and 1,131 images.
- Paul copied his entire notebook into a shared one for testing, so we develop against real notes with
  no risk to the originals.
- `InkWord` is not the Ink-to-Text feature Paul declines to use — but his notebook has zero of either,
  so we design for `InkDrawing`.

**Items:**

- [x] *"Also I was thinking about pen input, idk if that's a wrench in the gears"* — [RT-05] minted in roadmap to address
- [x] *"Is it like a shape overlay or a text select?"* — answered: it can be both, and they are different anchor types. [RT-03] + [RT-04] minted in roadmap to address
- [x] *"I think some sort of shape overlay would be ideal, I'm often covering pictures and printouts in OneNote"* — [RT-04] minted in roadmap to address; also drove [RT-10]
- [ ] Paul: *"There's a way to write and it turns into typed text. Hate it, why would anyone ever do that."* Wilson: *"BECAUSE PEOPLE WILL DO IT."* — Ink-to-Text output is a plain `one:T`, so the text path covers it. Confirm once RT-03 lands.

---

## Check-in — 2026-08-12 (Phase 0 spike)

> Context: first working session. Went to run the Phase 0 spike from PLAN.md, discovered the platform
> assumption was untested, and spent the day finding out whether the project was buildable at all.

**Decisions (Wilson, 2026-08-12):**

- Build the spike as a real COM add-in rather than a PowerShell harness — external automation is dead,
  so there is no scripting fallback to design around.
- Reference implementation only for OneMore (MPL-2.0 vs our MIT); read it, do not fork it.
- Target `net48`, so a student never installs a runtime.

**Items:**

- [x] OneNote returns E_FAIL to every external automation client on 16.0.20228 — [RT-01] minted in roadmap to address
- [x] Add-in must use the real PIA types; hand-declared COM interfaces silently fail — [RT-02] minted in roadmap to address
- [x] Survey Notebooks gives no feedback and looks frozen on a large notebook — [RT-08] minted in roadmap to address
- [x] Gemini's suggestion of a click-through overlay window is wrong for tape, right for the mascot — [RT-09] minted in roadmap to address
- [x] PLAN.md phase order is inverted relative to real notes — [RT-10] minted in roadmap to address
- [x] Zero `InkWord` in 79 real pages — **answered 2026-08-12.** Paul added a page written specifically
      to force one; the survey went 79 -> 80 pages, +23 `InkDrawing`, +2 `T`, and **still zero
      `InkWord`**. It is the build, not the copy: modern OneNote does not emit the element. Ink-to-Text
      lands in `one:T`, which the text path already covers. Folded into [RT-05], which is now a
      confirmed cost rather than a maybe.
- [x] Reveal is asymmetric: taped text peeks when you select over it, overlay images do nothing —
      **fixed 2026-08-12** by [RT-11]. Ctrl+click an overlay now toggles it, ~45ms.
- [ ] Reveal gestures are still *different*: text peeks on select, overlays need Ctrl+click. Better than
      one working and one not, but a user still has to learn two things. Worth a design pass.
- [x] Writes fail with 0x80042010 "the last modified date does not match" when the page moves between
      our read and our write — the concurrency check doing its job, but with no retry the operation just
      fails. [RT-12] minted in roadmap to address.
- [ ] Select-to-peek on text is free, and also a leak — a stray drag or Ctrl+A reveals every answer on
      the page. Feature, bug, or both? Needs a design call before the study loop is built.
- [x] The build silently depended on OneMore being installed — the OneNote PIA was only in the GAC
      because OneMore's installer put it there (GAC entry dated 2026-08-01, its release date). Both
      interop assemblies are now vendored in `lib/`, so the build is hermetic. Paul does **not** need
      OneMore to run RecallTape, and never did.
- [x] Paul is on a Snapdragon X Elite Surface (ARM64). Two things this forces, both now done:
      `PlatformTarget=AnyCPU` (an x64-only assembly cannot load into an ARM64 host at all), and the
      DCOM `LaunchPermission` SDDL (ARM64 does not grant local-launch to BUILTIN\Users, so COM refuses
      to start dllhost and the add-in never loads). Neither is verified on real ARM64 hardware yet.
- [ ] Verify on Paul's Surface: does the add-in load, and does click-to-peek work, on real ARM64?
- [ ] `Ctrl+Alt+T` on non-US layouts — still unchecked, and now more urgent since the target user has
      a Surface, not the keyboard we test on.
- [ ] Do `OCRToken` coordinates share the page coordinate space or are they image-relative? Must be
      settled before any overlay is placed from them.
- [ ] `Ctrl+Alt+T` collides with `AltGr` on Italian/UK keyboard layouts — check before it reaches a README.
- [ ] OneMore upstream bug: `HashtagProvider.Upgrade4to5` throws `duplicate column name: included`,
      leaving its hashtag catalog stuck at v4. Worth reporting from the fork.
- [ ] The COM surrogate holds the add-in DLL open — OneNote must be closed before every rebuild. Fine
      for now; annoying enough to automate if the loop gets tighter.
- [ ] Fusion (CLR assembly bind) logging was enabled machine-wide during debugging. Turn it off.
