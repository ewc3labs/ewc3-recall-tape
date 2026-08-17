<div align="center">

<img src="images/EWC3LabsLogo-blue-128x128.png" alt="EWC3 Labs" width="96" height="96">

# RecallTape

### Tape for OneNote. Finally.

**Hide answers. Reveal them later. Study directly in your notes.**

Active-recall masking for Microsoft OneNote — the Goodnotes-style tape tool that OneNote never
shipped.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE) [![Status:
alpha](https://img.shields.io/badge/status-alpha-orange.svg)](#status) [![Made by EWC3
Labs](https://img.shields.io/badge/EWC3-Labs-1f6feb.svg)](https://github.com/ewc3labs) [![Buy Me a
Coffee](https://img.shields.io/badge/Buy%20Me%20a%20Coffee-ffdd00?logo=buymeacoffee&logoColor=black)](https://www.buymeacoffee.com/ewc3labs)

</div>

---

## The problem

> *"Goodnotes has Tape. I use OneNote. Why the hell can't OneNote do this?"*

If you study from your own notes, you already know the technique: cover the answer, try to recall
it, then check. Goodnotes and Notability have shipped a tape tool for years. OneNote — the note app
with the best pen input on Windows, the one running on every Surface in every lecture hall — has
never had one.

The usual workarounds are all bad. Collapse an outline and lose your layout. Scribble over the
answer with a highlighter and then undo it. Keep your notes in OneNote and your recall in a second
app, and do everything twice.

**RecallTape puts the tape in OneNote.**

## What it does

Select an answer, a term, a label on a diagram. Press a hotkey. It's covered.

*(Already installed? The [how-to](HOW-TO.md) has the buttons that exist today — the hotkeys below
are still on the way.)*

Then study one-handed:

```
Space → answer → Space → answer → Space
```

Reveal, check yourself, retape, advance. Straight down your own lecture notes, no flashcards to
build and no second app to sync.

> **Why hotkeys instead of tapping each strip?** Because it's *better*. Clicking tiny pieces of tape
> is fiddly on a tablet and slow on a laptop. A single key that reveals-then-advances turns a page of
> notes into a study loop you can run without looking at the keyboard.

### Planned

| | |
| --- | --- |
| **Tape** | `Ctrl+Alt+T` on any selection |
| **Study loop** | reveal → retape → advance, on one key |
| **Reveal All / Retape All** | reset the page in one go |
| **Remove Tape** | non-destructive, always |
| **Ink support** | handwriting is the *point* on a Surface, not an afterthought |
| **Tape all handwriting** | bulk-tape a page instead of hand-picking fifty terms |
| **Study stats** | hits, misses, and eventually spaced repetition |
| **Anki export** | take your taped items with you |

## Status

**Alpha.** Taping and peeking work on real notes, and you can install it. It is not signed, the
installer is a script rather than an MSI, and it has not been through anyone's exam week yet. Things
will break. [Tell us when they do][tell-us-when-they-do] — that is what an alpha is for.

## Install

**Requires:** Windows, OneNote **desktop** (the Office one, not the Store app), and .NET Framework
4.8 — which is already on every Windows 10/11 machine.

### ⬇ [Download RecallTape][download-recalltape]

1. **Extract the zip anywhere** — Downloads is fine
2. Double-click **`Install RecallTape.cmd`** and say yes to the admin prompt
3. Restart OneNote
4. Delete the folder you extracted, whenever you like

RecallTape installs itself to `C:\Program Files\EWC3 Labs\RecallTape` and runs from there, so the
folder you extracted stops mattering the moment step 2 finishes.

*(That link always fetches the newest version. If you would rather browse, [the releases
page][the-releases-page] has it under Assets — grab `RecallTape-install.zip`, not "Source code
(zip)".)*

You'll find a **RecallTape** tab in the ribbon.

### 📖 [How to use it →](HOW-TO.md)

Three buttons and one gesture. Covers taping text, ink and unselectable page backgrounds, peeking
(which differs between boxes and text), removing, and what to do when something misbehaves.

Administrator is needed once — to copy into Program Files and register a COM class machine-wide.
There is no per-user install: the DCOM permissions that make RecallTape work on ARM64 are machine
configuration. OneMore, the add-in most OneNote users already know, is per-machine for the same
reason.

Program Files is deliberate rather than tidy: it is the one location a standard user cannot write
to, so nobody can swap the DLL that OneNote loads. Logs and page dumps live under
`%LOCALAPPDATA%\EWC3 Labs\RecallTape`, which is per-user and needs no privilege.

> **Windows will warn you.** The download isn't code-signed yet, so SmartScreen flags it as unrecognized
> — click *More info* → *Run anyway*. That's a real warning about a real gap, not a formality; a signing
> certificate is on the list before this is something we ask strangers to trust.

**To uninstall:** run `Uninstall RecallTape.cmd` from `C:\Program Files\EWC3 Labs\RecallTape`. It
removes every registry key and the installed folder, and leaves nothing behind. Your notes are
untouched — but tape already on a page stays there as ordinary content, so use **Remove All Tape**
on any page you want cleared *before* uninstalling.

## How it's built

Windows **desktop** OneNote, C#, against the OneNote COM API — which hands you the current page as
XML *including the selection*, and takes your modified version back.

That's a deliberate choice, not a default. Microsoft's modern Office.js add-in API targets **OneNote
on the web**, which is not where lecture notes get taken with a pen. The desktop API is older and
weirder and it's the one that can actually reach your ink.

It's also narrower than the documentation suggests. On current Office builds OneNote no longer
answers *external* programs at all — scripts and standalone tools get an object that fails every
call. The API is reachable **only from a registered add-in**, which is why RecallTape ships as one:
a signed installer and a ribbon, not a script you run. That single constraint shapes most of the
engineering, and it was found by measurement rather than by reading docs — the [receipts are in the
repo](docs/RAG_Sessions/2026-08-12_OneNote_Desktop_COM_External_Automation_Returns_E_FAIL_On_Current_Channel.md),
alongside [how and why we keep them][how-and-why-we-keep].

The add-in itself runs **outside** OneNote, in a COM surrogate process. If RecallTape ever crashes,
it takes down a `dllhost.exe` and OneNote keeps running with your notes intact. Given what this tool
touches, that isn't a detail.

We are standing on [**OneMore**][onemore] by [Steven Cohn][steven-cohn] — an excellent open-source
OneNote add-in that already solved ribbon integration, hotkeys, and page XML manipulation, and has
been maintained against a shifting undocumented target for nine years. It is the best documentation
this API has, because it's the one that actually runs. We use it as a **reference implementation**
rather than a base: OneMore is MPL-2.0 and RecallTape is MIT, so we read it, learn from it, and
write our own. Our read of how it works is public too, in [`docs/analysis/`][docs-analysis].

The masking engine is kept deliberately host-agnostic, so OneNote is its *first* host rather than
its only one:

```text
RecallTape.Core       the tape model — knows nothing about OneNote
RecallTape.OneNote    page + selection adapters, XML, overlay rendering
RecallTape.UI         ribbon, hotkeys, study controller
```

Built on **.NET Framework 4.8** — already present on every Windows 10/11 machine, so installing
RecallTape never means installing a runtime first.

## Origins

Wilson's son Paul started medical school and asked:

> **"yo how can i tape in onenote? is there an add-on?"**

There wasn't. So we're building one. The full conversation — including the platform research, the
architecture argument, and the part where a salad fork became a required feature — is preserved in
[`docs/origins.md`](docs/origins.md).

## Contributing

Not yet, but soon — the structure has to land first. When it opens up, see
[`CONTRIBUTING.md`](CONTRIBUTING.md). Issues and ideas are welcome any time.

## Support this project

RecallTape is free, open source, and always will be — no paywalls, no premium tier, no "unlock ink
masking for $4.99."

If it saves you time before an exam, you can [**buy me a coffee**][buy-me-a-coffee]. Entirely
optional, and genuinely appreciated.

☕ **[buymeacoffee.com/ewc3labs][buy-me-a-coffee]**

Free ways to help that matter just as much:

- **Star the repo** — it's most of how anyone finds a tool like this
- **File an issue** when something breaks, especially with a page that reproduces it
- **Tell someone who takes notes on a Surface**

## License

[MIT](LICENSE) — © EWC3 Labs.

Not affiliated with or endorsed by Microsoft, Goodnotes, or Notability. OneNote is a Microsoft
trademark; we just wish it had tape.

---

<div align="center">

**[EWC3 Labs][ewc3-labs]** — *tools that make Microsoft's power-user products behave like somebody
actually uses them.*

</div>

[buy-me-a-coffee]: https://www.buymeacoffee.com/ewc3labs
[docs-analysis]: docs/analysis/onemore-onenote-interaction.md
[download-recalltape]: https://github.com/ewc3labs/ewc3-recall-tape/releases/latest/download/RecallTape-install.zip
[ewc3-labs]: https://github.com/ewc3labs
[how-and-why-we-keep]: docs/RAG_Sessions/README_RAG_SYSTEM.md
[onemore]: https://github.com/stevencohn/OneMore
[steven-cohn]: https://github.com/stevencohn
[tell-us-when-they-do]: https://github.com/ewc3labs/ewc3-recall-tape/issues
[the-releases-page]: https://github.com/ewc3labs/ewc3-recall-tape/releases/latest
