# lib/ — vendored Microsoft interop assemblies

Two Microsoft assemblies, committed deliberately so the build is **hermetic**: any machine with the
.NET SDK can build RecallTape, with no Office, no Visual Studio, and no other add-in installed.

| File | Version | Provenance |
| --- | --- | --- |
| `Microsoft.Office.Interop.OneNote.dll` | 15.0.0.0 | strong name `71e9bce111e9429c`; Authenticode **Valid**, `CN=Microsoft Corporation` |
| `Extensibility.dll` | 7.0.3300.0 | strong name `b03f5f7f11d50a3a` (Microsoft). Not Authenticode-signed — it is the .NET 1.0-era assembly, predating ubiquitous signing |

```text
sha256  257b042bbab38406cb720fb9b2275828b003c6be15933227ceac68e08b846412  Microsoft.Office.Interop.OneNote.dll
sha256  b342088c8a4403092703bf40062041265e12edd204aff4f6532226478a65cbb2  Extensibility.dll
```

## Why vendored rather than resolved from the machine

The build previously resolved both from the GAC. That worked on exactly one computer — and for a bad
reason: **the OneNote PIA was in that GAC because OneMore's installer put it there.** Its GAC entry is
dated `2026-08-01`, OneMore 7.3.0's release date. A clean machine or a CI runner has neither assembly,
so the build silently depended on an unrelated add-in being installed first.

Rejected alternatives:

- **A Visual Studio runner.** OneMore does this (`gha-setup-vsdevenv`), because VS ships the Office
  PIAs. It works, and it means needing Visual Studio to compile a small add-in.
- **The `Interop.Microsoft.Office.Interop.OneNote` NuGet package.** Not published by Microsoft — a
  third-party repack. Trusting a stranger's copy of a COM interop assembly that runs inside OneNote,
  to avoid committing 60 KB, is a bad trade.

## What actually ships

Only `Microsoft.Office.Interop.OneNote.dll` — it is referenced with `EmbedInteropTypes=False` and copied
beside the add-in, the same way OneMore redistributes it.

`Extensibility.dll` is **build-time only**: it is referenced with `EmbedInteropTypes=True`, so the
interface definitions are compiled into our assembly and nothing is redistributed.

## Updating these

Don't, unless there is a reason. If you must, take them from a machine with Office and Visual Studio,
verify the strong name and Authenticode status as above, and update the hashes in this file.
