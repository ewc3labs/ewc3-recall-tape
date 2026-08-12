# RecallTape RAG Documentation System

- **Purpose**: *"Be kind to our future selves"* — durable, shared knowledge management for humans **and** agents
- **Pattern**: version-controlled repo memory under `/docs/RAG_Sessions/`, separate from any one agent's private memory

---

## 🎯 What Belongs Here?

Knowledge that must survive session resets, laptop wipes, agent vendor switches, and new contributors.

Belongs here:

- hard-won problem → solution narratives, especially non-obvious root causes
- architectural reasoning, **including the options that were rejected and why**
- discovered platform behaviour that contradicts the documentation
- workarounds whose *reason* is invisible in the code
- compressed session memory worth preserving in-repo

Does **not** belong here:

- scratch reasoning or transient task logs
- anything already obvious from the code, the diff, or `git log`
- raw chat transcripts — convert them into maintainable knowledge

### Shared vs private memory

`/docs/RAG_Sessions/` is the **shared** surface. Agent-native memory (Claude auto-memory, Copilot
workspace memory) is a **private, per-user, per-tool cache** — right for preferences and working habits,
and for **pointers** to these docs.

**Pointer, not copy.** These docs are the single source of truth for durable technical reasoning. A copy
drifting inside invisible per-user memory is undetectable and will eventually contradict the repo. If a
durable insight is found living only in agent memory, promote it here and reduce the memory to a pointer.

---

## 🗂️ Session Index

### Canonical Active Sessions (Read First)

1. **OneNote desktop COM: external automation returns E_FAIL on Current Channel 16.0.20228**
   **File**: `2026-08-12_OneNote_Desktop_COM_External_Automation_Returns_E_FAIL_On_Current_Channel.md`
   **Use for**: why RecallTape must be a COM add-in in a `dllhost` surrogate and cannot be driven by an external script; the working registration recipe; the diagnostic techniques that found it.
   **Key Decisions**:
   - External OneNote automation is dead on this build — no ROT registration, every method returns `E_FAIL`
   - RecallTape ships as a COM add-in registered into a surrogate; **no external-automation fallback exists to design around**
   - Target `net48`, not `net8.0` — already present on every Windows box, so no runtime install for a student
   - **Use the real Office PIAs.** Hand-declared COM interfaces are constructed by OneNote and then never called; `dynamic` fails on `IDispatch::GetTypeInfo`. This reverses the original decision, and the reversal was expensive
   - OneMore is **MPL-2.0** vs our MIT — reference implementation only, never a source of files
   - **Resolved same day**: the round-trip works from inside the surrogate; only the *caller* is gated

### Superseded Sessions (Historical Traceability)

*(none yet)*

### Outdated Sessions (`_ARCHIVE/`)

*(none yet)*

---

## 📝 Naming Convention

**Format**: `YYYY-MM-DD_Descriptive_Title.md`

The filename is the human interface. The architect browses this folder by eye, and the name alone must
convey *what happened and when* without opening anything. Be generous — prefer
`2026-08-12_OneNote_Desktop_COM_External_Automation_Returns_E_FAIL_On_Current_Channel.md` over
`2026-08-12_com_notes.md`. Never trade scannability for brevity.

> **Convention resolved 2026-08-12:** `docs/RAG_Sessions/` with `YYYY-MM-DD_Descriptive_Title.md`,
> per the `ewc3labs-agent-rag-system` skill. HQ `AGENTS.md` was updated to match, so there is one
> convention across EWC3 Labs.

---

## 🧭 How to Use These Docs

### For AI agents starting a new session

1. Read `README.md`, then `AGENTS.md`, then `docs/origins.md` — the platform decisions live there.
2. Read this index and load only the sessions relevant to the current problem.
3. Treat these as durable context, **not** as a substitute for checking live code and live systems.
   This corpus exists precisely because documentation went stale once already.

### For humans

1. Use the index to find prior decisions fast.
2. Read a session for reasoning and trade-offs, not just the conclusion.
3. Use it for onboarding and memory refresh.

---

## 📚 Primary References

**Always current:**

- [`README.md`](../../README.md) — what RecallTape is
- [`AGENTS.md`](../../AGENTS.md) — repo contract and decided-vs-open questions
- [`docs/PLAN.md`](../PLAN.md) — phased roadmap and non-negotiables
- [`docs/design/`](../design/) — architecture, and **why**
- [`docs/analysis/`](../analysis/) — investigation and current-state evidence, notably
  [`onenote-page-xml-shapes.md`](../analysis/onenote-page-xml-shapes.md) (what the API actually returns)
  and [`onemore-onenote-interaction.md`](../analysis/onemore-onenote-interaction.md)
- [`docs/project/`](../project/) — roadmap and punchlist: where we are and what is next

**Durable historical context:**

- `docs/RAG_Sessions/YYYY-MM-DD_*.md` — timestamped reasoning

Repo contracts are the current authority. RAG sessions preserve the *why*; they do not override
`AGENTS.md`.

---

## ✍️ Contributing New Sessions

1. Create `YYYY-MM-DD_Descriptive_Title.md` in `/docs/RAG_Sessions/`.
2. Start from the canonical session template (`templates/docs/template_RAG_session.md` in `ewc3labs-hq`).
3. Capture the problem, the options considered, the decision, and the lessons — reasoning, not just conclusions.
4. Update the index above.
5. If you keep agent-native memory, store a **pointer** to the doc, never a copy.

**When to write one:** a non-obvious root cause, a multi-hour debug, a design decision with rejected
alternatives, or a workaround whose reason will be invisible in the code. Routine fixes do not qualify —
doc spam erodes the corpus.

---

## 🔁 Superseded and Archive Standards

- **Canonical** — current starting points; read first.
- **Superseded** — still explains how the design evolved; stays in `/docs/RAG_Sessions/` with a pointer to its replacement.
- **Outdated** — design fully replaced; move to `/docs/RAG_Sessions/_ARCHIVE/` and drop from the active lists.

```markdown
## Superseded Notice

This session is superseded for active guidance.

- Superseded by: `/docs/RAG_Sessions/YYYY-MM-DD_New_Canonical_Session.md`
- Superseded on: YYYY-MM-DD
- Reason: {one-line reason this older memory is no longer canonical}

Keep this file only for historical traceability.
Use the superseding session for current decisions and implementation guidance.
```

Always name the replacement. Never leave ambiguous supersession.

---

## 🧱 Related Documentation Conventions

| Path | Holds |
| --- | --- |
| `docs/` root | front-page docs — `PLAN.md`, `origins.md` |
| `docs/design/` | architecture — what the design is and **why**, options rejected |
| `docs/analysis/` | investigation — current-state evidence, deep dives |
| `docs/RAG_Sessions/` | reasoning memory — how hard problems were actually solved |
| agent-native memory | private per-user cache — preferences and **pointers** into the above |

---

**Last Updated**: 2026-08-12
