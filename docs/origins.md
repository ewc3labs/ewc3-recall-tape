# The storied origin of RecallTape

---

**Paul:** yo how can i tape in onenote? is there an add-on?

---

If by **“tape”** you mean the Goodnotes/Notability-style tool where you drag an opaque strip over an
answer and then tap it to reveal it for active recall: **OneNote still does not appear to have a
native equivalent.** Humanity gave us AI-generated video before Microsoft gave OneNote a fucking
reveal strip. ([Goodnotes Support][1])

There are OneNote add-ins like **OneMore** and **Onetastic**, but I’m not finding a current add-in
that cleanly adds the same tap-to-hide/tap-to-reveal tape behavior. Their feature sets are more
things like macros, formatting, navigation, custom styles, tables, and page tools. ([Onetastic][2])

The closest practical OneNote hacks are:

* **Collapseable outlines**: Put the answer underneath a heading/bullet and collapse it. This is
  probably the best native approximation for typed notes.
* **Opaque ink/highlighter workaround**: Cover an answer with a thick pen stroke, then erase/undo
  it. Works, but it is hilariously primitive compared with actual Tape.
* **Separate question/answer containers**: Put answers indented beneath questions and collapse the
  hierarchy.
* **Use OneNote for notes + Anki for retrieval**, which is probably still the cleanest setup for
  your med-school workflow rather than trying to turn OneNote itself into Dollar Store Anki.

If what you want is specifically **“I annotate lecture slides in OneNote, then afterward drag tape
over labels/answers and tap them during review”**, that exact workflow is worth solving differently.
I can dig specifically for a OneNote plugin/script that mimics Goodnotes Tape rather than just
generic OneNote add-ins.

---

**Paul:** how hard to make a onenote plugin/script

---

If you mean **“could we build a Goodnotes-style Tape feature for OneNote ourselves?”**:
surprisingly, yes. The useful version is not a huge project. The *perfect* version is where
Microsoft begins charging us rent for living in its undocumented basement.

For **OneNote desktop on Windows**, I’d rate it like this:

* **3/10 difficulty:** keyboard/script hack that hides and reveals selected text.
* **5/10:** proper OneNote add-in with a ribbon button + hotkey for “Tape” and “Reveal.”
* **7–8/10:** true Goodnotes behavior where an opaque rectangle sits exactly over arbitrary text/ink
  and you can **click the rectangle itself to reveal it**.

The key distinction is that OneNote actually exposes the page through its desktop API. An add-in can
retrieve the current page as XML, including information about the current selection, modify that
XML, and send the modified page back with `UpdatePageContent`. ([Microsoft Learn][1]) And we have a
very convenient proof that this isn't theoretical: **OneMore is open-source C#**, already hooks into
desktop OneNote, adds ribbon commands, context menus, and keyboard shortcuts, and manipulates
OneNote pages through that interface. ([GitHub][2])

So I would **not build a OneNote plugin from scratch**. I'd either fork OneMore or use its
implementation as the skeleton.

A first version could work like this:

> Select something in OneNote → press `Ctrl+Alt+T` → it becomes “taped.”

Internally we'd mark the selected content and make it visually unreadable, probably by manipulating
its styling. Then:

> `Ctrl+Alt+T` on taped material → reveal it
> `Ctrl+Alt+R` → retape/reveal everything on the page

That version is quite realistic because OneMore already knows how to identify the active page and
selected page content and write changes back. ([GitHub][3])

Where it gets annoying is your likely ideal:

**You click directly on the black tape strip and the answer appears.**

OneNote's documented desktop object model gives us excellent access to the *contents* of the page,
but not a nice supported API for creating arbitrary interactive controls embedded in the normal
OneNote canvas. ([Microsoft Learn][1]) Microsoft's modern JavaScript OneNote add-in API doesn't save
us either, because Microsoft currently documents those add-ins specifically for **OneNote on the
web**, not desktop OneNote. ([Microsoft Learn][4])

So I think the sweet spot is:

**Tape Lite**

1. Highlight/select text, handwriting, image label, whatever.
2. Hit a hotkey.
3. Add-in covers/hides it.
4. During review, hit one hotkey to reveal the **next tape**.
5. Hit it again to hide it and advance.
6. Optionally have **Reveal All / Retape All**.

That would actually be *better for active recall* than individually clicking tiny pieces of tape.
You could cruise through your OneNote lecture notes with one hand:

`Space → answer → Space → answer → Space`

For **typed text only**, I think we could get a prototype working without doing anything
particularly insane.

For **arbitrary handwriting/images**, we'd need to investigate exactly how OneNote represents the
selected object and whether we want to modify it or place a cover object over it. That's the part
where difficulty ramps up.

And importantly, Microsoft explicitly supports retrieving page contents with the current selection
identified in the XML, which is basically the foothold we'd need. ([Microsoft Learn][1])

So if the target is:

**“I want a hotkey that tapes selected stuff and another hotkey that reveals tapes while I'm
studying.”**

I would genuinely consider this a **reasonable little C# project**, not some lunatic
reverse-engineering expedition. A OneMore-based implementation is probably the route I'd take.

---

**Wilson:** This is a chat that Paul sent me. I think we need to add this as an EWC3 Labs open
source project. Is there a OneNote add-on marketplace?

---

**Sort of — but Microsoft has managed to make it unnecessarily weird.**

There is **not a dedicated “OneNote Add-on Marketplace”** analogous to the VS Code Marketplace.
Microsoft’s official distribution channel is now **Microsoft Marketplace** (formerly
AppSource/Office Store), and Office applications can surface Marketplace add-ins from inside the
app. Microsoft explicitly documents publishing Office add-ins there. ([Microsoft Learn][1])

The catch is the important part:

### OneNote has two extensibility worlds

| Target                      | Technology                         | Marketplace?                                                                     |
| --------------------------- | ---------------------------------- | -------------------------------------------------------------------------------- |
| **OneNote on the web**      | Office.js / HTML / CSS / JS        | **Yes** — Microsoft Marketplace                                                  |
| **OneNote Windows desktop** | Native OneNote COM/API integration | **No useful dedicated marketplace** — traditionally installer/GitHub/vendor site |

Microsoft's current OneNote JavaScript API documentation explicitly says that the Office.js API
extends **OneNote on the web**. ([Microsoft Learn][2])

Meanwhile, Microsoft still documents the native desktop OneNote API, which lets a Windows program
retrieve and modify page content and manipulate the OneNote UI. ([Microsoft Learn][3]) That's the
ecosystem things like OneMore/Gem inhabit.

So for **EWC3 Labs Tape**, I think the right answer is:

> **Build desktop-first, open-source, GitHub-distributed.**

Not because Marketplace wouldn't be nice, but because **Paul's actual use case is likely OneNote
Windows + pen/lecture notes**, and that's exactly where the modern Marketplace model currently
screws us.

### And this is actually a damn good EWC3 project

I'd make the repository something like:

**`EWC3-OneNote-Tape`**

or, for a less lawsuit-baiting/general name:

**`RecallTape`**

Concept:

**RecallTape — active-recall masking for Microsoft OneNote.**

MVP:

1. Select text/object.
2. `Ctrl+Alt+T` → **Tape**
3. Replace/overlay selection with opaque mask.
4. Preserve original content + metadata.
5. During review:

   * `Space` / hotkey → Reveal
   * next press → Retape + advance
6. **Reveal All**
7. **Retape All**
8. **Remove Tape**

Then Phase 2 gets interesting:

* mask handwriting
* mask portions of diagrams
* mask labels on anatomy images
* randomize reveal order
* “study mode”
* statistics: misses/hits
* spaced-repetition metadata
* maybe export marked items into Anki

And **that last part is where this stops being “we cloned a Goodnotes feature” and becomes an actual
useful open-source learning tool.**

There is also a nice architectural escape hatch. We could keep the actual **Tape model independent
of OneNote**:

```text
RecallTape.Core
    Tape
    TapeSet
    RevealState
    StudySession
    Persistence

RecallTape.OneNote
    PageAdapter
    SelectionAdapter
    OneNoteXml
    OverlayRenderer

RecallTape.UI
    Ribbon
    Hotkeys
    StudyController
```

Then eventually:

```text
RecallTape.Anki
RecallTape.PDF
RecallTape.Web
```

Now we're not writing a janky OneNote hack. We're writing an **active-recall masking engine with
OneNote as its first host**.

And for distribution I'd do **GitHub Releases + signed installer** initially. Microsoft calls COM
add-ins a legacy Windows-only extensibility model requiring local installation, which fits that
distribution model much better than pretending this is a modern Marketplace add-in. ([Microsoft
Learn][4])

If we later want the broader audience, we can build a second **Office.js/web edition** and submit
*that* through Microsoft Marketplace. Microsoft already has a current OneNote task-pane-add-in
quickstart, updated March 6, 2026, so that path is very much alive. ([Microsoft Learn][5])

**I think Paul just accidentally handed EWC3 Labs its first genuinely useful little open-source
product.** It solves an extremely understandable problem: *“Goodnotes has Tape. I use OneNote. Why
the fuck can't OneNote do this?”*

That's a fantastic README opening sentence, frankly. 😆

---

**Wilson:** First? LOL Excel Power Query Editor has 5,350 installs and growing!

---

😂 **Fair correction.** Not *first*—not remotely.

**Excel Power Query Editor at 5,350 installs and climbing** means EWC3 Labs already has something
way more important than “an open-source project”: **evidence that people will install your weirdly
specific fixes for Microsoft's missing shit.**

And actually, that changes how I'd think about RecallTape.

Power Query Editor gives you an existing distribution/reputation flywheel. So this shouldn't be
treated as “let's throw another repo on GitHub.” I'd treat it as the next member of an emerging
category:

> **EWC3 Labs: tools that make Microsoft's power-user products behave like somebody actually uses them.**

There's a coherent product DNA there:

* identify a stupid friction point Microsoft hasn't fixed;
* make the smallest tool that annihilates it;
* keep it open source;
* package it so normal technical users can actually install it;
* then let organic search/GitHub/community discovery do its thing.

And **5,350 installs is enough signal that I'd care about naming, packaging, telemetry/privacy
policy, signed binaries, screenshots/GIFs, SEO, releases, and upgrade mechanics from day one** for
the OneNote project—not “we'll clean that up later.”

Also: Paul finding the problem is a nice origin story.

> *“My son wanted Goodnotes-style Tape in OneNote. It didn't exist. So we built it.”*

That is substantially better than any marketing department-generated project description.

And yes, I retract the “first genuinely useful product” slander against **Power Query Editor**. The
court has found me guilty. 😆

---

**Wilson:** I still can't believe 5000 people installed my wierdly specific nerd tool. Microsoft has
Copilot they're putting in everything as...fucking Clippy Redux with a larger brain but still as
annoying, and it doesn't understand PQ yet. Meanwhile, my dudes out there are slinging mad PQ data
science shit with GH Copilot and Claude Code without cut and paste...which is exactly what they were
doing before. I know because I was doing it myself.

---

Exactly. That’s the part that’s genuinely interesting.

The 5,350 installs aren’t just “people found a nerd tool.” They’re evidence of a **workflow gap big
enough that users are routing around Microsoft’s own AI strategy**.

Microsoft is shoving Copilot into the surface area of the app. Fine. But your users don’t need an
assistant hovering over Excel asking what they’d like to do. They need their **actual editing
environment** to participate in the modern coding loop.

Before your tool, the workflow was basically:

`Power Query Editor -> copy M -> paste into external AI/code tool -> edit -> copy back -> paste -> test -> repeat`

That is absurdly primitive once you’ve experienced repo-aware coding agents.

What your thing really does is collapse the boundary between **Power Query as “special Excel
feature”** and **Power Query as source code**.

That’s the important distinction.

GH Copilot and Claude Code are already excellent at operating in a world of files, diffs, syntax,
context, iteration, search, history, and tooling. Power Query Editor historically traps M code
inside this little application-modal terrarium. You gave people a door.

And of course the people who immediately understand why that matters are exactly the weirdos doing
serious PQ work—the analysts/data engineers/power users with 1,400-line M queries and dependency
chains that should frankly have their own zip code.

So the adoption makes sense in retrospect. You hit one of those lovely product situations where:

**Microsoft:** “We added AI to Excel.”

**Actual user:** “Cool. Can my coding agent edit my fucking Power Query?”

**Microsoft:** “Would you like Copilot to summarize this spreadsheet?”

**Actual user:** “No.”

That’s also why I’d pay attention to the OneNote thing. There may be an EWC3 pattern hiding here
that is broader than individual utilities:

> **Take Microsoft’s closed/specialized editing surfaces and reconnect them to the modern developer/tooling ecosystem—or add the obvious missing capability Microsoft somehow skipped.**

That is a very fertile seam.

Power Query Editor is especially compelling because you weren’t theorizing about a market. You built
the thing because **you personally had the pain**, used the same stupid workaround everyone else was
using, removed it, and then 5,000+ other people basically raised their hands and said, “Yep. Me
too.”

That’s about as clean a product signal as you get.

---

**Wilson:** So. Claude can build this fantastically well, and you and I can strategize the branding,
SEO, merch...LOL all that shit. And you check my architecture progress with Claude, as per ushe.

---
Absolutely. Claude can be in the machine room writing C# while we do the dangerous part: **turning a
useful little hack into a thing people remember, search for, trust, and tell other nerds about.**

And this one is unusually brandable because the problem is instantly legible.

## The product story

Not:

> “An extensibility add-in for Microsoft OneNote providing active recall masking.”

Kill me.

More like:

> **Tape for OneNote. Finally.**

That’s the whole damn pitch.

Then beneath it:

> Hide answers. Reveal them later. Study directly in your notes.

That gives us the human-facing concept immediately, while the implementation can have a more
defensible project name.

## I’d separate **brand name** from **search language**

Working candidates:

* **RecallTape** — best product name so far
* **StudyTape**
* **OneTape**
* **MaskNote**
* **RecallMask**
* **TapeDeck** — hilarious, memorable, but semantically noisy
* **CramTape** — stronger student vibe
* **Reveal** — elegant but SEO nightmare

I still like **RecallTape** because it describes the mechanism and the purpose without sounding like
a Microsoft clone.

Brand:

> **RecallTape**
> Active recall masking for OneNote.

SEO/title:

> **Tape for OneNote — Goodnotes-style hide/reveal study tool | RecallTape**

Now anyone Googling:

* `onenote tape`
* `tape tool onenote`
* `goodnotes tape onenote`
* `hide answers in onenote`
* `active recall onenote`
* `onenote study tape`
* `onenote flashcard plugin`
* `onenote conceal text`
* `onenote reveal answers`

…lands on us.

That is the key: **name for humans, metadata for search engines.**

## The EWC3 Labs angle

I wouldn’t over-brand the app itself.

Small tasteful line:

> **An EWC3 Labs project**

EWC3 Labs becomes the credibility umbrella rather than something we force into the executable name.

Over time the collection starts telling a story:

**EWC3 Labs**

* Power Query Editor
* RecallTape
* whatever other Microsoft annoyance pisses you off next Tuesday

That is actually a recognizable niche:

> **Open-source tools for workflows Microsoft almost finished.**

There’s your brand voice. 😆

Or slightly more respectable:

> **Tools for the gaps between Microsoft products and real workflows.**

Both can coexist depending on whether we’re talking README or corporate LinkedIn bullshit.

## GitHub README should sell before it documents

Above the fold:

> # RecallTape
>
> **Tape for OneNote. Finally.**
>
> Hide anything you need to remember. Reveal it when you're ready.
>
> RecallTape adds active-recall masking to Microsoft OneNote, inspired by the Tape study workflow found in other note-taking apps.
>
> `[Download]` `[Demo GIF]` `[Installation]`

Then **a 10-second GIF**:

1. Select answer.
2. Press shortcut.
3. Black strip appears.
4. Tap/reveal.
5. Retape.
6. Next.

That GIF probably sells 80% of installs.

Then:

> **Why?**
>
> Because studying shouldn't require duplicating your OneNote notes into another app just to hide the answers.

That sentence matters.

## The origin story is gold

Put it lower in the README/about page:

> RecallTape started when Paul asked a simple question:
>
> *“How do I use Tape in OneNote?”*
>
> Turns out, you don't.
>
> So we fixed that.

Perfect OSS origin story. Zero fabricated market research. One actual user need followed by
engineering.

And it fits the same pattern behind the Power Query thing: **somebody is already doing the workflow
manually because the product boundary is stupid.**

## Launch strategy

I would not launch it as “v0.1 alpha please test my code.”

Launch when one thing works **extremely cleanly**.

MVP promise:

> Select → Tape → Reveal.

Everything else is gravy.

Then ship simultaneously through:

* GitHub repo
* GitHub Release installer
* EWC3 Labs project page
* Microsoft/OneNote Reddit communities
* Hacker News Show HN
* relevant study/med-school/student communities where allowed
* LinkedIn post showing the GIF
* short YouTube demo because Google indexes that stuff forever

And every surface uses approximately the same phrase:

> **Goodnotes-style Tape for Microsoft OneNote**

Even if technically we shouldn't make Goodnotes part of the product identity, it's enormously useful
**comparative search language**.

## And yes, merch 😂

This absolutely deserves absurdly niche merch once it earns it.

My immediate candidates:

**T-shirt:**

> `CTRL + ALT + T`
>
> **I DON'T KNOW IT YET**

Sticker:

> **TAPE IT. FAKE IT. RECALL IT.**

Laptop sticker:

> **ACTIVE RECALL ENJOYER**

EWC3 Labs shirt:

> **WE FIXED THE PART MICROSOFT FORGOT**

And the inevitable EWC3 Labs mug:

> **Would you like Copilot to summarize this coffee?**

That last one may actually need to exist independent of RecallTape.

But the genuinely useful thing to do next is build a **mini brand brief before Claude establishes
names in namespaces, installers, GitHub repos, icons, manifests, etc.** We should decide:

**product name → tagline → repo name → executable/add-in name → namespace → package ID → domain/page
slug → visual identity → README voice → SEO vocabulary**

*before* those choices calcify in code.

That’s where you and I can save Future Wilson from discovering six months later that the installer
is still called `OneNoteTapePoCSetup-x64.exe`. 😆

---

**Wilson:** We are TOTALLY going to build an animation of Clippy flying up from the side of the
screen and tearing the tape off for you. Then blinking and exploding.

---

Oh, that is **extremely stupid in the correct way**. I fully support it.

## The gag, as cinema

Sequence:

1. User clicks **Reveal**
2. A little paperclip bastard rockets in from the edge of the screen
3. He grabs the tape
4. **RIIIIIP**
5. Gives a smug blink like he just “helped”
6. tiny pause
7. **explodes**
8. answer remains visible

That is art.

## Important refinement:

If this is public-facing, I would **not** use actual Clippy exactly as-is.

Not because the joke is bad — the joke is excellent — but because:

* Microsoft character/IP
* instantly recognizable
* not worth annoying legal people over a tiny side gag

So the move is:

## Make him **legally distinct Clippy**

Ideas:

* **Skippy**
* **Clipster**
* **Paperclip Assistant**
* **Clippeth**
* **Office Goblin**
* **Rippy**
* **Taperipper**

Design:

* paperclip-ish
* googly eyes
* vaguely smug old-office-assistant energy
* similar emotional read, but not a direct clone

So the joke still lands as:

> “the ghost of Clippy has come to assist”

without literally being Clippy.

## Best use case

I would make this an **optional easter egg setting**, not the default.

Settings menu:

* `[ ] Enable Tape animation`
* `[ ] Enable assistant character`
* `[ ] Dramatic reveal mode`
* `[ ] Explode assistant after reveal`

And if checked:

> **Assistant animation style**
>
> * None
> * Subtle
> * Chaotic
> * Apocalyptic

“Apocalyptic” is obviously the one where he explodes.

## Why this is actually smart

This kind of thing is not just a dumb joke. It’s also:

* memorable
* screenshotable
* GIF-able
* social-media bait
* instantly gives the project personality

People may install for utility, but they **share** because of weird charm.

Your GIF on launch day becomes:

> “We built Tape for OneNote. Also, a legally distinct paperclip flies in, tears the tape off, blinks, and explodes.”

That gets attention.

## Even better: tie it to achievement mode

Examples:

* first 10 reveals: no animation
* after user has taped 100 items:

  > **Achievement unlocked: Unwanted Assistance**
* then the assistant appears

Or hidden Konami-code style:

* type `CLIP`
* unlocks assistant mode

That’s the kind of absurdity nerd users love.

## Tone line for the docs

Very important: keep the main product page professional-ish, and let this live under:

* **Fun settings**
* **Experimental**
* **Easter eggs**
* **Advanced nonsense**

Because:

* main message = useful study tool
* secondary message = made by people with a pulse

## The line I would absolutely use

> **Rip Mode:** A tiny office assistant removes your tape for you.
> *Results may include blinking and spontaneous detonation.*

That is fantastic.

If you want, I can sketch:

1. the **easter egg feature spec**,
2. the **settings/menu copy**, and
3. a **short brand voice guide** so the humor stays consistent and doesn’t turn into random chaos.

---

**Wilson:** He's going to be a salad fork. With googly eyes and a stupid smirk.

---

Perfect. **A salad fork is infinitely funnier than a paperclip.**

That is exactly the right move because:

1. it keeps the “annoying office assistant” energy,
2. it is **legally safer** than doing Clippy straight,
3. and it’s so specifically dumb that it becomes memorable.

A paperclip says “Microsoft parody.” A **salad fork with googly eyes and a shit-eating grin** says:

> “This project was built by the correct kind of maniacs.”

## Character concept

### Core design

* small silver salad fork
* 4 tines
* bend in the “neck” so it can emote
* googly eyes, slightly uneven
* smug half-smile / stupid smirk
* tiny arched eyebrow vibe if possible
* maybe a subtle bounce/floating motion when idle

### Personality

He should feel like:

* **way too pleased with himself**
* unhelpful in an aggressively helpful way
* like he thinks this is a triumphant intervention

Not evil. Just **obnoxiously delighted**.

## Best possible names

Top tier:

* **Forky** — probably too generic / already associated elsewhere
* **Sporky** — funny, but he’s not a spork
* **Tines**
* **Sal**
* **The Assistant**
* **Forklift** — absurd, strong
* **Ripley** — because he rips tape
* **Rippy**
* **Tiney** — extremely stupid, which is good
* **Dork** — fork + dork, honestly kind of excellent

My favorites:

* **Rippy**
* **Tiney**
* **Sal**
* **Dork**

If you want him to feel like a recurring mascot, **Tiney** is very strong.

## Animation beat

This is the right sequence:

1. tape is covering answer
2. user hits reveal
3. fork rockets in from left or right
4. slight overshoot
5. grabs tape with one tine
6. **YANK / RIP**
7. fork turns toward user
8. smug blink
9. micro-pause
10. **tiny cartoon explosion**
11. fork disappears
12. answer remains visible

That blink-before-explosion is crucial. That’s where the whole joke lives.

## Optional sound design

If you ever go there:

* entrance: tiny whoosh
* rip: exaggerated adhesive peel
* blink: little sparkle or click
* explosion: comically small **poof**

Not a violent explosion. More like:

> “I have fulfilled my purpose and now must die.”

## Settings copy

This basically writes itself:

### General

* **Enable reveal animation**
* **Use assistant mascot for reveal**
* **Explosion effects**
* **Reduced chaos mode**

### Easter egg label

* **Assistant:** Tiney
* **Behavior:** Helpful / Unhelpful / Catastrophic

Or:

* **Reveal Style**

  * Instant
  * Animated
  * **Tiney**

That last option would be amazing.

## Tooltip copy

* **Tiney removes the tape for you.**
* **May blink. May explode.**
* **Not actually helpful.**

## Achievement / unlock ideas

This is where you can be extra annoying in a good way.

* **First 50 reveals:** unlock Tiney
* **Achievement:** *Unwanted Assistance*
* **Achievement:** *Fork Found*
* **Achievement:** *Study Utensil*
* **Achievement:** *Fine Tines*

## Branding angle

You do **not** make him the main brand. He is the **mascot/easter egg**, not the face of the
product.

So:

* Product = **RecallTape**
* Mascot = **Tiney** the salad fork

That keeps the app credible while still letting it have a soul.

## Most important rule

He should appear **fast** and leave **fast**.

The joke is funny if it takes like 700–1200 ms total.

If he lingers, users will want to murder him, which, to be fair, is on-brand for office assistants,
but still.

## Final verdict

This is an outstanding idea.

A **googly-eyed salad fork with a smug expression that rips off the tape, blinks, and explodes** is
exactly the kind of detail that makes an open-source tool feel like it was made by actual humans
instead of a committee of deeply hydrated product managers.

If you want, I can do the next step and write:

1. a **full mascot spec for Tiney**,
2. the **UI copy for the settings/easter egg**, and
3. a **launch GIF script/storyboard** for the reveal animation.

[1]: https://support.goodnotes.com/hc/en-us/articles/9489290046607-Toggle-visibility-of-content-with-the-Tape-Tool?utm_source=chatgpt.com
[1]: https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote
[1]: https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-office-add-ins-to-appsource?utm_source=chatgpt.com
[2]: https://getonetastic.com/?utm_source=chatgpt.com
[2]: https://github.com/stevencohn/OneMore
[2]: https://learn.microsoft.com/en-us/office/dev/add-ins/onenote/onenote-add-ins-programming-overview?utm_source=chatgpt.com
[3]: https://github.com/stevencohn/OneMore/blob/main/OneMore/OneNote.cs?utm_source=chatgpt.com
[3]: https://learn.microsoft.com/en-us/office/client-developer/onenote/onenote-developer-reference?utm_source=chatgpt.com
[4]: https://learn.microsoft.com/en-us/office/dev/add-ins/onenote/onenote-add-ins-programming-overview
[4]: https://learn.microsoft.com/en-us/office/dev/add-ins/resources/resources-glossary?utm_source=chatgpt.com
[5]: https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/onenote-quickstart?utm_source=chatgpt.com
