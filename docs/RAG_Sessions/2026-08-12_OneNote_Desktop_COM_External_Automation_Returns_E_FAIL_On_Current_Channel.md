# OneNote desktop COM: external automation returns E_FAIL on Current Channel 16.0.20228

- **Date**: 2026-08-12
- **Repository**: ewc3-recall-tape
- **Status**: active reference

---

## 🎯 Context

First working session on RecallTape. `docs/PLAN.md` Phase 0 gates the entire project on one spike:
attach to desktop OneNote, fetch the current page XML, confirm the selection is identifiable in it, and
write a trivially modified page back via `UpdatePageContent`.

`docs/origins.md` had already decided the platform with citations — desktop OneNote, C#, COM API,
because that is the only path that reaches ink. The research was done. What had never happened was
**running a single line of code against it**.

The session went straight at that, on a clean machine (`EWC3WS2`) where OneNote had never been
installed. The answer was not the one the plan assumed.

---

## ❓ Problem Statement

Can a program on this machine reach OneNote's documented COM automation API — and if not, is the
failure specific to this box, to this Office build, or to the whole platform bet?

This is existential rather than academic: if `GetPageContent` / `UpdatePageContent` are unreachable,
RecallTape as designed cannot exist.

---

## 🧱 Current State

**Environment as measured:**

| Fact | Value |
| --- | --- |
| OneNote | `C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE`, "Microsoft OneNote" |
| Office build | **16.0.20228.20158**, x64, Click-to-Run, **Current Channel** |
| Product IDs | `O365HomePremRetail, OneNoteFreeRetail` |
| OS | Windows 11 Pro 26H2, build 26300 |
| Store/UWP OneNote | **none** — only `Microsoft.Office.OneNoteVirtualPrinter` (a print driver) |
| COM registration | textbook: `LocalServer32 → ONENOTE.EXE`, `Application2 Class`, TypeLib `{0EA692EE-…}` v1.1, no `AppID`, no `RunAs` |

**External automation is dead on this build:**

| Probe | Result |
| --- | --- |
| `CoCreateInstance("OneNote.Application")` | succeeds, returns an object |
| `IDispatch::GetTypeInfo` on it | succeeds — full `IApplication` surface: `GetPageContent`, `UpdatePageContent`, `GetHierarchy`, `FindPages`, `Windows`, … |
| `GetSpecialLocation(0 or 2)` | **`0x80004005 E_FAIL`** |
| `GetHierarchy(…)`, all 5 scopes | **`0x80004005 E_FAIL`** |
| `Application.Windows` (parameterless property) | **null** |
| `Marshal::GetActiveObject("OneNote.Application")` | **`0x800401E3 MK_E_UNAVAILABLE`** — never registers in the ROT |
| Attached / cold-start / headless / 61s patient retry | identical `E_FAIL` every time |

**Control group, same host, same script style, same minute:** Excel and Word both bind via the ROT and
answer property calls normally (`.Version=16.0`, `Workbooks.Count`, `Documents.Count`). The harness is
not the problem.

**The interface is present and correctly typed. The server answers metadata requests and fails every
real call.** `GetSpecialLocation` merely returns a folder path — it touches no notebook, needs no sync,
requires no sign-in — and it fails identically. That single data point kills every "not initialised
yet" explanation.

---

## ⚖️ Options Evaluated

### Option 1 — PowerShell harness driving OneNote over COM
- Pros: zero toolchain (machine had .NET runtimes but **no SDK**), instant edit-rerun loop, disposable by design — exactly the prototype the house method wants torn out.
- Cons: depends entirely on external automation.
- Outcome: **rejected — proven impossible.** Cost about an hour and was still the cheapest way to learn it. Four hours into a C# add-in we would have hit the same `E_FAIL` through a debugger, wondering what *we* had done wrong.

### Option 2 — Install OneMore as an independent oracle
- Pros: known-good, actively-maintained third-party consumer of this exact API; verdict doesn't depend on a line of our code; closes the `docs/PLAN.md` licence-check item for free.
- Cons: installs software on the dev box; unsigned MSI.
- Outcome: **accepted.** Decisive, and it answered several unrelated questions at once (target framework, registration shape, distribution reality).

### Option 3 — Bridge add-in exposing `IApplication` over a named pipe
- Pros: restores the fast REPL loop that Option 1 lost, while getting real API access.
- Cons: a whole IPC surface to build and debug, larger than the thing being tested.
- Outcome: **rejected** (architect's call). A file drop gives identical feedback with no plumbing: the add-in writes page XML to `%TEMP%`, the agent reads the file.

### Option 4 — In-process managed add-in (plain RegAsm registration)
- Pros: simplest possible registration.
- Cons: —
- Outcome: **rejected — measured failure.** OneNote demotes `LoadBehavior` 3 → 2 and never constructs the class. Tried first deliberately, so that adopting the surrogate would be an observation rather than cargo-culting.

### Option 5 — COM surrogate (`AppID` + empty `DllSurrogate`)
- Pros: matches the reference implementation; isolates a crash to a `dllhost` instead of taking OneNote down.
- Cons: everything crosses a process boundary and must marshal.
- Outcome: **accepted.** With it, OneNote *does* construct our class inside `dllhost.exe`.

### Option 6 — Getting the COM interface definitions
- **`COMReference` / `tlbimp`**: rejected — `TlbImp.exe` ships with the Windows SDK, absent on a box with the .NET SDK and no Visual Studio. Verified absent before committing to the approach.
- **Redistribute the PIAs**: rejected for a spike — extra binaries and a licence question.
- **Hand-declare the interfaces**: **accepted.** Three small interfaces, every IID read off *this machine* rather than recalled.

### Option 7 — OneMore: fork / depend / reference
- Outcome: **reference only.** OneMore is **MPL-2.0**; RecallTape is MIT and HQ calls that non-negotiable. MPL-2.0 is file-level copyleft — forked files stay MPL permanently. Reading and reimplementing keeps RecallTape MIT; pasting does not.

---

## ✅ Decision

1. **The platform bet is not disproven, but it is no longer assumed.** External automation is dead on
   16.0.20228. The add-in extensibility path still constructs our class. Whether the API answers *from
   inside that surrogate* is the open question in Follow-Up.
2. **RecallTape ships as a COM add-in registered into a `dllhost` surrogate.** There is no external
   automation fallback to design around, so nothing may depend on one.
3. **Target `net48`**, not `net8.0`. Evidence: OneMore's `supportedRuntime` is
   `.NETFramework,Version=v4.8`, with no `runtimeconfig.json` and no `.comhost.dll`. Independently, .NET
   Framework 4.8 is present on every Windows 10/11 machine — **zero runtime install for a student who
   just wants tape.** CI's `dotnet-version: 8.0.x` remains fine as a *build driver*; the project targets
   net48, and that distinction needs a comment before someone "fixes" it.
4. **Hand-declare the COM interfaces**; no Windows SDK on the build box, no redistributed PIAs.
5. **Use OneMore as reference, never as a source of files.** Credit it prominently.

---

## 🛠️ Implementation Notes

**Registration recipe that gets our class constructed** (`tools/register.ps1`):

- `RegAsm.exe /codebase` → `InprocServer32 = mscoree.dll` plus `Class` / `Assembly` / `CodeBase` / `RuntimeVersion`
- `HKLM\SOFTWARE\Classes\CLSID\{clsid}` → `AppID = {clsid}`
- `HKLM\SOFTWARE\Classes\AppID\{clsid}` → `DllSurrogate = ""`
- `HKCU` **and** `HKLM` `…\Microsoft\Office\OneNote\AddIns\RecallTape.AddIn` → `LoadBehavior = 3`

**`LoadBehavior` is the diagnostic.** OneNote silently rewrites 3 → 2 when a load fails. It is the only
signal it volunteers, and it looks identical whether the class was never constructed or was constructed
and then failed. Re-arm it to 3 before every test or the next run is meaningless.

**IIDs, all read from this machine rather than memory:**

| Interface | IID | Source |
| --- | --- | --- |
| `IDTExtensibility2` | `{B65AD801-ABAF-11D0-BB8B-00A0C90F2744}` | byte-searched in `C:\Program Files\Common Files\DESIGNER\MSADDNDR.OLB` |
| `IRibbonExtensibility` | `{000C0396-0000-0000-C000-000000000046}` | `HKLM\SOFTWARE\Classes\Interface` |
| `IRibbonControl` | `{000C0395-0000-0000-C000-000000000046}` | same |

`IDTExtensibility2` must be **dual** with methods in typelib order (`OnConnection` … `OnBeginShutdown`,
DispIds 1–5). Confirmed authoritatively by reflecting over the *embedded* `Extensibility.IDTExtensibility2`
inside `River.OneMoreAddIn.dll`: `ComImport` + `Guid` + `TypeIdentifier`, and **no `InterfaceType`
attribute at all**, which defaults to dual. Declaring it `InterfaceIsIDispatch` gives the CCW no vtable
past `IDispatch`, and the universal typelib marshaler (`ProxyStubClsid32 = {00020424-…}`) drives the
interface through its vtable.

**The class must carry only `[ComVisible(true)]`, `[Guid]`, `[ProgId]`.** Specifying
`[ClassInterface(ClassInterfaceType.None)]` is wrong; the reference implementation carries no
`ClassInterface` attribute at all.

**Diagnostics that actually worked, in order of usefulness:**

1. **Log from the constructor, not just `OnConnection`.** Without it, "never instantiated" and
   "instantiated but never called" are indistinguishable — both present as `LoadBehavior=2` and an empty
   log directory. Adding one constructor log line converted a total unknown into a narrow one.
2. **CLR fusion binding log** (`HKLM\SOFTWARE\Microsoft\Fusion`, `EnableLog`/`ForceLog`/`LogFailures`,
   `LogPath`). Shows whether the surrogate ever loaded the assembly. **Turn it off afterwards** — it
   slows every .NET process on the machine.
3. **Diffing against the working implementation** — its registry footprint, then its assembly metadata
   by reflection, then its source.

**Beware a false lead:** the fusion log shows a *failed* default-context bind probing
`C:\Windows\system32\` alongside a *successful* CodeBase bind. That failure is normal — **OneMore
produces the identical pair and works.**

---

## 💡 Lessons Learned

- **Documentation is not a running system.** `origins.md` cited Microsoft's own docs saying the desktop
  API works. It is still documented. It no longer answers external callers. The citations were correct
  and the conclusion was stale, and only running code could tell the difference.
- **A version string is not proof of a working API.** OneMore's startup log prints
  `OneNote 15.0 (16.0.20228.20158 x64)` and it was briefly taken as evidence the API worked. It is not:
  the version comes from `HKCR\OneNote.Application\CurVer` and the build from `ONENOTE.EXE`'s file
  header. Both local reads. **Check what a log line is actually made of before resting a decision on it.**
- **Bring a control group.** Excel and Word answering normally in the same script, same minute, is what
  turned "my harness is broken" into "OneNote is broken" in one step.
- **Find the call with the fewest dependencies.** `GetSpecialLocation` returns a folder path — no
  notebooks, no sync, no sign-in. Its failure eliminated a whole family of hypotheses at once.
- **A pattern that reproduces across four contexts is the system, not the setup.** Attached, cold-start,
  headless, and 61 seconds of patient retry all gave byte-identical `E_FAIL`.
- **Try the simple thing first even when you expect the complex one.** Registering in-process first made
  "use the surrogate" a measured result rather than an imitation.
- **Reverse-engineering a working implementation through its registry footprint is the wrong tool when
  its source is open.** Three failed attempts went by before reading `AddIn.cs`. Cost: about an hour.
- **The most valuable thing in a mature codebase is often a comment.** OneMore's commented-out
  `GC.SuppressFinalize` — *"DO NOT call this otherwise OneNote will not shutdown properly"* — is
  nine years of pain compressed into one line, and it looks exactly like a bug without it.

---

## 🔜 Follow-Up Work

- [x] ~~**BLOCKING: prove a page round-trip works from inside the surrogate.**~~ **Passed, same day.**
      OneMore's `ConvertMarkdownCommand` — a read *and* a write — ran clean on this build, and our own
      add-in then read real XML off 79 pages. The asymmetry is the finding: **`CoCreateInstance` returns
      a dead object to an external process and the identical call works from inside a OneNote-spawned
      surrogate.** `UpdatePageContent` is still only proven by OneMore's behaviour, not our own code;
      `RT-03` closes that.
- [x] ~~Fix the spike bug: class constructed but `OnConnection` never called.~~ **Root cause:
      hand-declared COM interfaces do not satisfy Office's add-in contract.** Correct IID
      (byte-verified against the typelib), dual layout, typelib method order, explicit DispIds, even
      `[TypeIdentifier]` for COM type-equivalence — every check passed and the whole thing still failed.
      Only the real embedded `Extensibility` PIA works. Two further traps followed, each failing
      differently and none loudly:
      `dynamic` throws `E_FAIL` inside `GetITypeInfoFromIDispatch` because the C# COM binder wants
      `IDispatch::GetTypeInfo` and OneNote refuses it; and casting to `IApplication` yields a *lazy*
      `E_NOINTERFACE` — the cast appears to succeed, then throws on first member access — because
      `IApplication` is not a `ComImport` type. The working forms are `ON.Application`, or better
      `new ON.Application()`, which is what OneMore does exclusively.
- [x] ~~Finish the original Phase 0 spike — page XML for typed text, ink, and an image.~~ Done; written
      up with receipts in [`../analysis/onenote-page-xml-shapes.md`](../analysis/onenote-page-xml-shapes.md).
      Headline: **two anchor worlds** (flow text vs positioned ink/images), tape is a positioned
      `one:Image` at `z=max+1`, and **zero `InkWord` in 11,685 real ink elements**.
- [ ] Adopt the ARM64 `LaunchPermission` SDDL — Surface devices are literally the target hardware.
- [ ] Pre-register an EventLog source at install time, or the first error masks the real one.
- [ ] **Do not** inherit OneMore's `UpdatePageContent(xml, DateTime.MinValue, xs2013, force: true)`.
      That disables optimistic concurrency. See `docs/analysis/onemore-onenote-interaction.md` §9.
- [ ] Verify `Ctrl+Alt+T` against non-US keyboard layouts — on Italian/UK, `AltGr` *is* `Ctrl+Alt`.
- [ ] Turn off fusion logging when debugging concludes.
- [ ] Re-test on Paul's machine — real target hardware, real workload, and a different Office channel.

---

## 📂 Related Files

- `docs/analysis/onemore-onenote-interaction.md` — how OneMore does all of this, with diagrams
- `docs/PLAN.md` — Phase 0 checklist this session was working against
- `docs/design/architecture.md` — updated with the hosting shape this session established
- `docs/origins.md` — the platform decision this session stress-tested
- `src/RecallTape.OneNote/Interop.cs` — hand-declared COM interfaces, IIDs sourced in comments
- `src/RecallTape.OneNote/AddIn.cs` — the spike add-in
- `tools/register.ps1` · `tools/unregister.ps1` — the registration recipe above

---

## 🗂️ Update README_RAG_SYSTEM.md

Indexed as a canonical active session in `docs/RAG_Sessions/README_RAG_SYSTEM.md`.
