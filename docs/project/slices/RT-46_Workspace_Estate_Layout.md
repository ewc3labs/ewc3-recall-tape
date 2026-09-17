# RT-46 — Say which clone layout the workspace assumes

> **Keep this doc small.** One slice or one fix. If it grows past roughly a page, either the slice is
> really several slices, or it belongs in `modules/`.

| | |
| --- | --- |
| **State** | ⬜ planned |
| **Est** | S |
| **Roadmap** | [`RecallTape_Development_Roadmap.md`][recalltape] |

## What and why

`ewc3-recall-tape.code-workspace` opens HQ at `../../ewc3labs-hq`. That is right for the EWC3 Labs
layout — `C:\DEV\ewc3labs-hq\` beside `C:\DEV\ewc3labs\`, which holds the project repos — but the
layout is written down only in HQ's `ewc3labs-new-machine` skill, not in this repository.

Codex read the same file on PR #2 and concluded that the paths ascend one directory too far. It was
wrong, and it was wrong for the reason a newcomer would be: nothing here says where HQ lives. In a
public repository, anyone who clones only RecallTape opens a workspace with two missing folders and
no explanation.

## Steps

- [ ] Say in one place a contributor sees — README "Development" or CONTRIBUTING — that the
      workspace is for the EWC3 Labs estate layout, and name that layout
- [ ] Say that cloning only this repo works fine: open the folder, not the workspace

## Watch out for

- Do not "fix" the paths to `../ewc3labs-hq`. That breaks the real layout to satisfy a guessed one.
- Do not restate the whole new-machine setup here; HQ owns it. One sentence and the layout is
  enough.

## Done when

- A reader of this repository alone can tell why HQ's folders are missing, and what to open instead.

## Links

- **[PR #2 Codex finding][pr-2-codex-finding]** — *the misread that prompted this slice*

[pr-2-codex-finding]: https://github.com/ewc3labs/ewc3-recall-tape/pull/2
[recalltape]: ../RecallTape_Development_Roadmap.md
