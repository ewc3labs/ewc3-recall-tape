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

An item is promoted to a roadmap slice **once someone has done enough analysis to document the
problem with receipts**. Until then it stays an unchecked box.

**Unchecked items are "what's still open."** There is no separate status file.

| Mark | Means |
| --- | --- |
| `- [ ]` | open — noticed, not yet understood |
| `- [~]` | being looked at |
| `- [x]` | minted into the roadmap, or delivered |

Checked items name their slice: `[RT-nn] minted in roadmap to address`. Before minting anything new,
search the roadmap **including Parked / Retired** — the same thing gets noticed months apart.

---

## Check-in — 2026-08-13 (Paul installs it, and immediately wants more)

> Context: v0.0.1 published. Paul installed it himself from the release zip, on his Snapdragon X Elite
> Surface, with no help beyond the README — which closes both open shipping slices at once, on
> hardware nobody here owns.

**Decisions (Wilson, 2026-08-13):**

- Feature requests from Paul go to GitHub Issues with screenshots, not to this punchlist. He is a
  user now, not a bystander, and users file issues.

**Items:**

- [x] Paul's shared notebook arrived with sync conflicts ALREADY on it — first open, his own history
      across his own devices, not ours. The lesson is not blame, it is that conflicted pages are a
      normal condition of real notebooks. And the API cannot see conflict state at all — no
      attribute exists anywhere in the schema — so we tape them blind. [RT-32].
- [ ] Unknown and worth an hour before the study loop: on a conflicted page, does `GetPageContent`
      return the side the user is looking at? Does a tape survive manual resolution, or vanish with
      the losing side and leave markers pointing at nothing?
- [x] Version history: NOT in the API (29 methods, none touch it). But
      `GetSpecialLocation(slBackUpFolder)`
      + `OpenHierarchy(path, cftNone)` + `GetPageContent` chains into a real read-the-old-version
        path, and `SectionGroup/@isRecycleBin` enumerates deleted pages. OneNote had ALREADY backed
        the shared notebook up locally, dated before we touched it.
- [ ] Tell users the backup folder exists. It is the honest answer to "what if this eats my notes",
      it costs one line of UI, and it is true today with no work from us.
- [ ] Does `UpdatePageContent` accept a page fragment — just our element and its objectID — instead
      of the whole page? Documented, but OneMore always sends whole pages, so it is untrodden. A
      smaller write is a better position against a sync layer we cannot inspect.
- [ ] A study loop that writes on every reveal multiplies that footprint. Decide before [RT-16]
      whether reveal state must persist at all, or can live in memory for the session.
- [ ] Microsoft now pops a Copilot affordance wherever you click on a page. Two notes: it may
      collide with Ctrl+click-to-peek, worth watching; and it makes the [RT-09] mascot a parody of
      something *currently shipping* rather than a nostalgia joke. Arguably better.
- [x] Free tape box when nothing is selectable, dragged onto the target — Wilson's design, [RT-24],
      proven same day. Works because our tape IS a one:Image, so OneNote already gives it move and
      resize handles; we place an object and get out of the way.
- [x] Remove one tape vs Remove All, with a confirm on the destructive one — [RT-27].
- [x] A peeked tape was invisible and unfindable; it now keeps a gray outline — [RT-28].
- [x] Can corner handles resize asymmetrically? — **no, and not our call.** `Size` exposes only
      width, height and isSetByUser; there is no aspect property anywhere in the schema. Corner
      behavior belongs to OneNote's manipulation UI. [RT-29] canceled. Arbitrary rectangles are
      already reachable by scribbling the shape in ink and taping the ink.
- [ ] Anchor mode: explicit Anchor/FreeForm toggle, or implied by whether the page has an anchor?
      Leaning implied — a mode is a thing the user has to keep track of. Undecided.
- [x] *"Okay it works!"* — install from the release zip succeeded on a machine that did not build
      it, on ARM64. Closes [RT-14] and [RT-23], both of which were reasoned but untested until now.
- [x] *"Now can it be done over top of a background image? Like something not selectable"* — [RT-24]
      minted in roadmap to address. A page background is a `one:Image` with
      `backgroundImage="true"`; OneNote never lets the user select it, so nothing ever arrives
      marked `selected="all"` and selection-driven taping cannot reach it. Needs a way to specify a
      region without selecting an object.
- [x] Workaround available today, confirmed on a real page: scribble over the area with the pen,
      select the ink, hit **Tape**. Ink is selectable and positioned, so its bounding box becomes
      the tape.
- [x] Do background images carry `OCRData`? — **yes**, and it is better than hoped. A PowerPoint
      printout becomes 39 background images, every one OCR'd, 449 tokens across the deck. [RT-06]
      unblocked: occluding a term on a lecture slide needs no drawing at all.
- [x] *"we need an easy button to open a GH issue from OneNote"* — [RT-25] minted in roadmap to
      address, gated behind [RT-26]. One page of Paul's notes contains his username, his school and
      course folders, his full name, four Microsoft account CIDs, the page title, and the OCR'd text
      of all 39 lecture slides. The last of those is also somebody else's copyrighted material, and
      a public issue is publication. "Attach the XML" cannot be one click.
- [ ] Does a background image carry `OCRData`? If so, RT-06 reaches it too and Paul's printouts
      become tapeable by word rather than by rectangle.

---

## Check-in — 2026-08-12 (Paul's requirements, over Telegram)

> Context: Wilson shared the repo mid-build. Paul — the actual user, the one who asked the original
> question — reacted with requirements before we had anything to show him.

**Decisions (Wilson + Paul, 2026-08-12):**

- Region/shape overlay is the primary interaction, not a Phase 2 nicety. Confirmed by the notebook
  survey: 892 typed elements against 11,685 ink and 1,131 images.
- Paul copied his entire notebook into a shared one for testing, so we develop against real notes
  with no risk to the originals.
- `InkWord` is not the Ink-to-Text feature Paul declines to use — but his notebook has zero of
  either, so we design for `InkDrawing`.

**Items:**

- [x] *"Also I was thinking about pen input, idk if that's a wrench in the gears"* — [RT-05] minted
      in roadmap to address
- [x] *"Is it like a shape overlay or a text select?"* — answered: it can be both, and they are
      different anchor types. [RT-03] + [RT-04] minted in roadmap to address
- [x] *"I think some sort of shape overlay would be ideal, I'm often covering pictures and printouts
      in OneNote"* — [RT-04] minted in roadmap to address; also drove [RT-10]
- [ ] Paul: *"There's a way to write and it turns into typed text. Hate it, why would anyone ever do
      that."* Wilson: *"BECAUSE PEOPLE WILL DO IT."* — Ink-to-Text output is a plain `one:T`, so the
      text path covers it. Confirm once RT-03 lands.

---

## Check-in — 2026-08-12 (Phase 0 spike)

> Context: first working session. Went to run the Phase 0 spike from PLAN.md, discovered the platform
> assumption was untested, and spent the day finding out whether the project was buildable at all.

**Decisions (Wilson, 2026-08-12):**

- Build the spike as a real COM add-in rather than a PowerShell harness — external automation is
  dead, so there is no scripting fallback to design around.
- Reference implementation only for OneMore (MPL-2.0 vs our MIT); read it, do not fork it.
- Target `net48`, so a student never installs a runtime.

**Items:**

- [x] OneNote returns E_FAIL to every external automation client on 16.0.20228 — [RT-01] minted in
      roadmap to address
- [x] Add-in must use the real PIA types; hand-declared COM interfaces silently fail — [RT-02]
      minted in roadmap to address
- [x] Survey Notebooks gives no feedback and looks frozen on a large notebook — [RT-08] minted in
      roadmap to address
- [x] Gemini's suggestion of a click-through overlay window is wrong for tape, right for the mascot
      — [RT-09] minted in roadmap to address
- [x] PLAN.md phase order is inverted relative to real notes — [RT-10] minted in roadmap to address
- [x] Zero `InkWord` in 79 real pages — **answered 2026-08-12.** Paul added a page written
      specifically to force one; the survey went 79 -> 80 pages, +23 `InkDrawing`, +2 `T`, and
      **still zero `InkWord`**. It is the build, not the copy: modern OneNote does not emit the
      element. Ink-to-Text lands in `one:T`, which the text path already covers. Folded into
      [RT-05], which is now a confirmed cost rather than a maybe.
- [x] Reveal is asymmetric: taped text peeks when you select over it, overlay images do nothing —
      **fixed 2026-08-12** by [RT-11]. Ctrl+click an overlay now toggles it, ~45ms.
- [ ] Reveal gestures are still *different*: text peeks on select, overlays need Ctrl+click. Better
      than one working and one not, but a user still has to learn two things. Worth a design pass.
- [x] Writes fail with 0x80042010 "the last modified date does not match" when the page moves
      between our read and our write — the concurrency check doing its job, but with no retry the
      operation just fails. [RT-12] minted in roadmap to address.
- [ ] Select-to-peek on text is free, and also a leak — a stray drag or Ctrl+A reveals every answer
      on the page. Feature, bug, or both? Needs a design call before the study loop is built.
- [x] The build silently depended on OneMore being installed — the OneNote PIA was only in the GAC
      because OneMore's installer put it there (GAC entry dated 2026-08-01, its release date). Both
      interop assemblies are now vendored in `lib/`, so the build is hermetic. Paul does **not**
      need OneMore to run RecallTape, and never did.
- [x] Paul is on a Snapdragon X Elite Surface (ARM64). Two things this forces, both now done:
      `PlatformTarget=AnyCPU` (an x64-only assembly cannot load into an ARM64 host at all), and the
      DCOM `LaunchPermission` SDDL (ARM64 does not grant local-launch to BUILTIN\Users, so COM
      refuses to start dllhost and the add-in never loads). This bites even a local administrator:
      with UAC on, a normal process runs with a filtered token where `BUILTIN\Administrators` is
      deny-only, so the `AU` ace is what actually grants it. Neither change is verified on real
      ARM64 hardware yet.
- [ ] Verify on Paul's Surface: does the add-in load, and does click-to-peek work, on real ARM64?
- [ ] `Ctrl+Alt+T` on non-US layouts — still unchecked, and now more urgent since the target user
      has a Surface, not the keyboard we test on.
- [x] Could not disable RecallTape from OneNote's COM Add-ins dialog without running OneNote
      elevated — **fixed 2026-08-13.** Cause was a stray `HKLM` add-in registration left behind by a
      debugging script during the OnConnection hunt; our installer only ever writes `HKCU`. Office
      will not let a non-elevated user untick an add-in registered machine-wide, because clearing
      the box writes `LoadBehavior` into that hive. Stray key removed, and
      `uninstall.ps1`/`unregister.ps1` now sweep `HKLM` defensively even though we never create it.
- [ ] Diagnostic scripts changed machine state that outlived the diagnostic — this stray key and the
      machine-wide fusion logging were both found days later by accident. Worth a habit: scratch
      scripts that touch the registry should print what they changed and how to undo it.
- [x] Paul could not install it at all — **fixed 2026-08-13**, and every cause was ours. The release
      was still a DRAFT so he could not see it (*"There are no releases on that project"*), found
      the tag, and downloaded GitHub's "Source code (zip)", which contains no built program. Our
      error then told him to "extract the whole zip and run it from there" — which he had done, from
      the wrong zip. Now detected explicitly and named.
- [x] A `.ps1` has no "Run as administrator" on its context menu, and scripts out of a downloaded
      zip are blocked by execution policy — two obstacles between a med student and a study tool.
      Added `Install RecallTape.cmd` / `Uninstall RecallTape.cmd`, which self-elevate, unblock, and
      run with a per-process policy override. Double-click, accept UAC, done.
- [ ] Nobody has installed RecallTape from a release zip successfully yet. [RT-14] stays open until
      someone has, on a machine that is not the one it was built on.
- [ ] Do `OCRToken` coordinates share the page coordinate space or are they image-relative? Must be
      settled before any overlay is placed from them.
- [ ] `Ctrl+Alt+T` collides with `AltGr` on Italian/UK keyboard layouts — check before it reaches a
      README.
- [ ] OneMore upstream bug: `HashtagProvider.Upgrade4to5` throws `duplicate column name: included`,
      leaving its hashtag catalog stuck at v4. Worth reporting from the fork.
- [ ] The COM surrogate holds the add-in DLL open — OneNote must be closed before every rebuild.
      Fine for now; annoying enough to automate if the loop gets tighter.
- [ ] Fusion (CLR assembly bind) logging was enabled machine-wide during debugging. Turn it off.
