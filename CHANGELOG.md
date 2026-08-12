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

### Known Issues
- **OneNote no longer serves external COM automation clients** on Office 16.0.20228 (Current Channel).
  The add-in path is the only way in; there is no scripting fallback.
- The spike add-in is constructed by OneNote inside the surrogate, but `OnConnection` is never called.
- A `GetPageContent` round-trip from inside the surrogate is **not yet proven** on this build.

### Notes
- Pre-alpha. No installable build yet.
