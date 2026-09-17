# RecallTape — Slices

**One slice or one fix per file.** These are the tiny execution docs the roadmap's Delivery Index
points at, so that table rows can stay one line instead of swelling into paragraphs.

## Naming

```text
<SLICE_ID>_<Short_Title>.md      e.g.  RT-04_Ink_Selection_Probe.md
```

The filename is the interface. Someone scanning this folder — or the roadmap's `Doc` column, which
pins these filenames — should know what a doc is without opening it. Never `RT-04.md`, never
`notes.md`.

## What goes in one

Execution detail: what to do, in what order, what will bite you, and how you will know it is done.
Template: `template_project_slice.md`. **Keep it to about a page.**

## What does not

| Content | Home |
| --- | --- |
| architecture, and *why*, options rejected | `docs/design/` |
| current-state investigation and evidence | `docs/analysis/` |
| how a hard problem was actually solved | `docs/RAG_Sessions/` |
| sequencing and priority | the roadmap's Current Focus |
| several slices in one subsystem | `modules/` |

## Lifecycle

A slice doc closes when the slice ships. Delete it or leave it as history — but do not let a
finished slice doc sit in the folder pretending to be active work. The roadmap row's state is the
truth.
