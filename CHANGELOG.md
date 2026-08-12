# Changelog

All notable changes to RecallTape are documented here.

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/);
versioning follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

### Changed
- Target framework fixed at **`net48`**, not `net8.0`. It is what the reference implementation uses and
  what is already installed on every Windows 10/11 machine — no runtime install for end users. CI's
  `dotnet-version: 8.0.x` is the build *driver*, not the target.
- `docs/design/architecture.md` — documents the `dllhost` COM surrogate hosting model, the build/runtime
  decisions, and an explicit rejection of `UpdatePageContent(..., force: true)` on concurrency grounds.

### Resolved
- **OneMore is MPL-2.0**, RecallTape is MIT. Closes the fork-vs-depend-vs-reference question in
  `docs/PLAN.md`: **reference implementation only**, since MPL-2.0 file-level copyleft would follow any
  forked file permanently.

### Proven
- **The add-in reads real page XML.** Full lifecycle (ctor, `OnConnection`, `GetCustomUI`,
  `OnStartupComplete`, `OnDisconnection`) plus `GetPageContent` and `GetHierarchy`, verified against a
  synthetic page and 79 pages of a real medical-school notebook.
- **The platform bet holds.** `CoCreateInstance` returns a dead object to an external process and a live
  one from inside a OneNote-spawned surrogate — the gate is on the caller, not the API.

### Added
- `SurveyNotebooks` command — walks every open notebook and reports content-type totals, structural
  skeletons, and `recognizedText` samples.
- `docs/analysis/onenote-page-xml-shapes.md` — what OneNote page XML actually contains, measured.
- `docs/project/` — roadmap and check-in punchlist (PMO-light, per the `ewc3labs-project-roadmap` skill).
- `.editorconfig`, and `.gitattributes` replaced with the EWC3 Labs LF-everywhere policy.
- README Support section — Buy Me a Coffee, plus the free ways to help. (The sidebar Sponsor button already comes from the org-level `ewc3labs/.github` funding file; no repo-level override needed.)
- GitHub About filled in: 14 discovery topics and an SEO-leading description.

### Changed
- **Reversed the hand-declared-interop decision.** The add-in now references the real `Extensibility`
  and `Microsoft.Office.Interop.OneNote` PIAs. Hand-declared COM interfaces — correct IID, dual layout,
  typelib method order, explicit DispIds, `[TypeIdentifier]` — are constructed by OneNote and then
  never called. `dynamic` fails separately, on `IDispatch::GetTypeInfo`.
- Application object is acquired per operation via `new Application()` and released immediately, rather
  than caching the one passed to `OnConnection` (which prevents OneNote shutting down).
- `docs/PLAN.md` flagged: phase order is inverted relative to real notes (892 typed elements vs 11,685
  ink and 1,131 images). Re-sequencing tracked as `RT-10`.

### Known Issues
- `UpdatePageContent` is not yet exercised by our own code — nothing here can write to a page (`RT-03`).
- **Zero `InkWord` in 11,685 real ink elements**, so `recognizedText` is unavailable and handwriting
  masking must cluster `InkDrawing` bounding boxes. Unclear whether that is the build, a setting, or an
  artifact of the notebook having been copied.
- Long operations give no progress feedback and look frozen (`RT-08`).
- The COM surrogate holds the built DLL open; OneNote must be closed before a rebuild.

### Notes
- Pre-alpha. No installable build yet.
