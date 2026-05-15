# Specifications — Office LLM Harness

This directory contains speckit feature specifications for the Office LLM Harness project, a Windows desktop Office add-in harness that exposes controlled document interaction tools to Open WebUI through a local MCP server.

## Architecture

```
Open WebUI                    MCP Server (port 3000)              Office Add-ins
    │                                │                               │
    │── tools/call ─────────────────►│                               │
    │   { tool, { instanceId?, ... } │                               │
    │                                │── route to instance ─────────►│
    │◄── result ─────────────────────│                               │
    │                                │◄── result ────────────────────│
    │                                │◄── /instances/register ──────│  (on load)
    │                                │◄── /instances/:id/heartbeat ─│  (every 10s)
    │                                │►── /instances/:id/commands ──│  (poll every 2s)
    │                                │◄── /instances/:id/result ────│  (after execution)
```

**Key design decisions:**
- **Central hub**: Single MCP server process, one port (3000)
- **Add-ins connect TO server**: Office JS Add-ins register on load, poll for commands
- **Fixed tool list**: Tools have optional `instanceId` parameter (defaults to most recent instance)
- **Multiple instances**: Each Office session gets its own registered instance (e.g., `powerpoint_1`, `word_2`)
- **Transport**: Streamable HTTP + stdio (MCPo compatible)

## Phases

| # | Spec | Branch | Scope |
|---|------|--------|-------|
| 0 | [Spike](001-spike/) | `001-spike` | MCP server hub + Office JS Add-in registration + Open WebUI integration validation |
| 1 | [PowerPoint MVP](002-powerpoint-mvp/) | `002-powerpoint-mvp` | Deck outline, slide read, shape update, speaker notes, audit log, task pane |
| 2 | [Word MVP](003-word-mvp/) | `003-word-mvp` | Outline, paragraphs, rewrite selection (tracked changes), review comments, shared context abstraction |
| 3 | [Excel MVP](004-excel-mvp/) | `004-excel-mvp` | Workbook map, read/write range, write formula, create table, range limits, formula validation |
| 4 | [Outlook MVP](005-outlook-mvp/) | `005-outlook-mvp` | Current item read, thread summary, draft reply (never auto-send), category apply, policy filter, send confirmation gate |

## Progression

Each phase is **independently completable** and builds on the shared infrastructure established by earlier phases:

```
Phase 0 (Spike) → Phase 1 (PPT) → Phase 2 (Word) → Phase 3 (Excel) → Phase 4 (Outlook)
     │                │               │                │                 │
     └── MCP hub     ├── PPT tools   ├── Word tools   ├── Excel tools   ├── Outlook tools
         + registry    + audit log     + shared ctx     + range limits   + policy filter
```

## Tool Design

Tools use a **fixed list** with optional `instanceId` parameter to target specific registered instances:

| Tool | Parameters | Default Instance |
|------|-----------|------------------|
| `office_get_active_app` | *(none)* | N/A — lists all instances |
| `powerpoint_get_deck_outline` | `instanceId?`, `includeSpeakerNotes?`, `includeHiddenSlides?` | Most recent |
| `powerpoint_get_slide` | `instanceId?`, `slideIndex` (required) | Most recent |
| `powerpoint_update_shape_text` | `instanceId?`, `slideIndex`, `shapeId`, `text` | Most recent |
| `powerpoint_update_speaker_notes` | `instanceId?`, `slideIndex`, `notes` | Most recent |
| `word_get_outline` | `instanceId?`, `maxDepth?` | Most recent |
| `word_rewrite_selection` | `instanceId?`, `tone` | Most recent |
| `excel_get_workbook_map` | `instanceId?` | Most recent |
| `excel_read_range` | `instanceId?`, `sheetName`, `address` | Most recent |
| `excel_write_range` | `instanceId?`, `range`, `values` | Most recent |
| `outlook_get_current_item` | `instanceId?` | Most recent |
| `outlook_summarize_thread` | `instanceId?` | Most recent |
| `outlook_draft_reply` | `instanceId?`, `tone`, `keyPoints` | Most recent |

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
