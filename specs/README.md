# Specifications — Office LLM Harness

This directory contains speckit feature specifications for the Office LLM Harness project, a VSTO-based Windows desktop Office add-in that exposes controlled document interaction tools to Open WebUI through a local MCP server.

## Phases

| #   | Spec                                  | Branch               | Scope                                                                                                                   |
| --- | ------------------------------------- | -------------------- | ----------------------------------------------------------------------------------------------------------------------- |
| 0   | [Spike](001-spike/)                   | `001-spike`          | Minimal MCP server + PowerPoint add-in + Open WebUI integration validation                                              |
| 1   | [PowerPoint MVP](002-powerpoint-mvp/) | `002-powerpoint-mvp` | Deck outline, slide read, shape text update, speaker notes, audit log, task pane                                        |
| 2   | [Word MVP](003-word-mvp/)             | `003-word-mvp`       | Outline, paragraphs, rewrite selection (tracked changes), review comments, shared context abstraction                   |
| 3   | [Excel MVP](004-excel-mvp/)           | `004-excel-mvp`      | Workbook map, read/write range, write formula, create table, range limits, formula validation                           |
| 4   | [Outlook MVP](005-outlook-mvp/)       | `005-outlook-mvp`    | Current item read, thread summary, draft reply (never auto-send), category apply, policy filter, send confirmation gate |

## Progression

Each phase is **independently completable** and builds on the shared infrastructure established by earlier phases:

```
Phase 0 (Spike) → Phase 1 (PPT) → Phase 2 (Word) → Phase 3 (Excel) → Phase 4 (Outlook)
     │                │               │                │                 │
     └── MCP server  ├── PPT tools   ├── Word tools   ├── Excel tools   ├── Outlook tools
         + IPC        + audit log     + shared ctx     + range limits   + policy filter
```

## Workflow

For each phase, follow the speckit workflow:

1. `/speckit.specify` — Refine the specification (already done)
2. `/speckit.plan` — Create implementation plan
3. `/speckit.tasks` — Generate actionable tasks
4. `/speckit.implement` — Execute implementation

## Constitution

All phases must comply with the project constitution at [`.specify/memory/constitution.md`](../.specify/memory/constitution.md).

Core principles:

- **Safety-First**: No arbitrary COM/shell/macro execution as model-callable tools
- **Local-Only**: MCP binds to `127.0.0.1` only
- **User Control**: Model proposes, user approves — all mutations require confirmation
- **Minimal Surface Area**: Narrow, intentional tools with typed inputs
- **Phased Delivery**: Each host is an independent MVP phase
