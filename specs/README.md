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
- **Unified add-in**: One Office JS Add-in, one manifest — auto-detects host via `Office.onReady(info.host)` and presents host-relevant tools
- **Add-ins connect TO server**: Office JS Add-ins register on load, poll for commands
- **Host-aware tool routing**: Instance registration includes host type (`PowerPoint`, `Word`, `Excel`, `Outlook`); MCP server filters available tools by host
- **Multiple instances**: Each Office session gets its own registered instance (e.g., `powerpoint_1`, `word_2`)
- **Transport**: Streamable HTTP + stdio (MCPo compatible) + OpenAPI REST bridge

## Phases

| #   | Spec                                  | Branch               | Scope                                                                                                                   |
| --- | ------------------------------------- | -------------------- | ----------------------------------------------------------------------------------------------------------------------- |
| 0   | [Spike](001-spike/)                   | `001-spike`          | MCP server hub + unified Office JS Add-in (auto-detects host) + Open WebUI integration validation                       |
| 1   | [PowerPoint MVP](002-powerpoint-mvp/) | `002-powerpoint-mvp` | Deck outline, slide read, shape update, speaker notes, audit log, task pane                                             |
| 2   | [Word MVP](003-word-mvp/)             | `003-word-mvp`       | Outline, paragraphs, rewrite selection (tracked changes), review comments, shared context abstraction                   |
| 3   | [Excel MVP](004-excel-mvp/)           | `004-excel-mvp`      | Workbook map, read/write range, write formula, create table, range limits, formula validation                           |
| 4   | [Outlook MVP](005-outlook-mvp/)       | `005-outlook-mvp`    | Current item read, thread summary, draft reply (never auto-send), category apply, policy filter, send confirmation gate |
| 5   | [PowerPoint v2](006-powerpoint-v2/)   | `006-powerpoint-v2`  | Full shape properties, image export, table reading, selection context, direct write operations, shape CRUD, slide management |

## Progression

Each phase is **independently completable** and builds on the shared infrastructure established by earlier phases:

```
Phase 0 (Spike) → Phase 1 (PPT) → Phase 2 (Word) → Phase 3 (Excel) → Phase 4 (Outlook)
     │                │               │                │                 │
     └── MCP hub     ├── PPT tools   ├── Word tools   ├── Excel tools   ├── Outlook tools
         + unified      + audit log     + shared ctx     + range limits    + policy filter
           add-in        + task pane
           (auto-detect
            host)
```

## Tool Design

Tools are organized by host. The MCP server registers a **host-aware tool list** — only tools relevant to the connected instance's host type are exposed. All tools accept an optional `instanceId` parameter to target specific registered instances.

### Host Detection

When the add-in loads, `Office.onReady((info) => ...)` provides `info.host`:

- `HostType.PowerPoint` → registers as `powerpoint_N`, exposes PowerPoint tools
- `HostType.Word` → registers as `word_N`, exposes Word tools
- `HostType.Excel` → registers as `excel_N`, exposes Excel tools
- `HostType.Outlook` → registers as `outlook_N`, exposes Outlook tools

The instance registration (`POST /instances/register`) includes the `appName` field, and the MCP server uses this to filter which tools are available for each instance.

### Unified vs Host-Specific Tools

| Tool                              | Host       | Parameters                                                    |
| --------------------------------- | ---------- | ------------------------------------------------------------- |
| `office_get_active_app`           | **All**    | _(none)_ — lists all instances                                |
| `powerpoint_get_deck_outline`     | PowerPoint | `instanceId?`, `includeSpeakerNotes?`, `includeHiddenSlides?` |
| `powerpoint_get_slide`            | PowerPoint | `instanceId?`, `slideIndex` (required)                        |
| `powerpoint_update_shape_text`    | PowerPoint | `instanceId?`, `slideIndex`, `shapeId`, `text`                |
| `powerpoint_update_speaker_notes` | PowerPoint | `instanceId?`, `slideIndex`, `notes`                          |
| `word_get_outline`                | Word       | `instanceId?`, `maxDepth?`                                    |
| `word_rewrite_selection`          | Word       | `instanceId?`, `tone`                                         |
| `excel_get_workbook_map`          | Excel      | `instanceId?`                                                 |
| `excel_read_range`                | Excel      | `instanceId?`, `sheetName`, `address`                         |
| `excel_write_range`               | Excel      | `instanceId?`, `range`, `values`                              |
| `outlook_get_current_item`        | Outlook    | `instanceId?`                                                 |
| `outlook_summarize_thread`        | Outlook    | `instanceId?`                                                 |
| `outlook_draft_reply`             | Outlook    | `instanceId?`, `tone`, `keyPoints`                            |

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
