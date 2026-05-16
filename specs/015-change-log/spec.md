# Feature Specification: Phase 15 — Operation Audit & Change Log

**Feature Branch**: `015-change-log`

**Created**: 2026-05-16

**Status**: Draft

**Depends on**: All MVP phases complete (audit log already exists)

**Input**: User description: "Expose the existing audit log as a tool so the LLM can review what it has done, undo operations, and provide change summaries to the user. Turns the audit trail into a first-class feature."

## Architecture

The backend already maintains an `AuditLog` that records every tool call. This phase exposes it via new tools: `office_get_change_log` to query the log, and `office_undo_last` to undo the most recent operation.

```
LLM Client                     MCP Server (port 3000)              Office Add-in
    │                                │                               │
    │── tools/call ─────────────────►│                               │
    │   { office_get_change_log      │                               │
    │     { lastN: 10 }}             │── read AuditLog ──────────────►│ (no add-in call needed)
    │◄── { entries: [...] } ─────────│                               │
    │                                │                               │
    │── tools/call ─────────────────►│                               │── push "undo" command ───────►│
    │   { office_undo_last           │                               │   (triggers Ctrl+Z in add-in)
    │     { instanceId }}            │                               │
    │                                │◄── { undone: true } ──────────│
    │◄── { undone: true, tool: ... } │                               │
```

## User Scenarios & Testing

### User Story 1 — Review what was done (Priority: P0)

A user asks "What did you just do to my document?" The LLM calls `office_get_change_log` and receives a chronological list of all operations with before/after values.

**Why this priority**: Trust is the #1 concern with LLM editing. Users need visibility into what changed. The audit data already exists — we just need to expose it.

**Independent Test**: Perform several operations, call `office_get_change_log`, verify the response lists all operations in order.

**Acceptance Scenarios**:

1. **Given** the LLM has performed 5 write operations, **When** developer calls `office_get_change_log` with `lastN: 10`, **Then** the response lists all 5 operations in chronological order with tool name, instance ID, timestamp, inputs, and outcome
2. **Given** the LLM replaced text 3 times and added a comment, **When** developer calls `office_get_change_log` with `filter: "write"`, **Then** only the 3 replacements and the comment appear (read operations excluded)
3. **Given** operations on two different documents (word_1 and word_2), **When** developer calls `office_get_change_log` with `instanceId: "word_1"`, **Then** only operations on word_1 appear

---

### User Story 2 — Undo the last operation (Priority: P0)

A user says "Wait, undo that last change." The LLM calls `office_undo_last` and the most recent write operation is reversed via `Ctrl+Z` in the add-in.

**Why this priority**: Error recovery is essential. Without undo, mistakes are permanent. With it, the user (or LLM) can experiment confidently.

**Independent Test**: Perform a write operation, call `office_undo_last`, verify the change is reversed.

**Acceptance Scenarios**:

1. **Given** the LLM just replaced text in a Word document with tracked changes, **When** developer calls `office_undo_last` with `instanceId: "word_1"`, **Then** the last replacement is undone and a reverse tracked change appears
2. **Given** the LLM wrote values to an Excel range, **When** developer calls `office_undo_last`, **Then** the values are reverted to their previous state
3. **Given** no write operations have been performed, **When** developer calls `office_undo_last`, **Then** the tool returns `{ undone: false, reason: "No operations to undo" }`

---

### User Story 3 — Change summary for the user (Priority: P1)

A user asks "Summarize all the changes you made." The LLM calls `office_get_change_log` and produces a human-readable summary grouped by type.

**Why this priority**: Change summaries close the communication loop between LLM and user. The audit data is already there — the LLM just needs structured access.

**Independent Test**: Perform mixed operations, call `office_get_change_log`, verify the LLM can generate a meaningful summary.

**Acceptance Scenarios**:

1. **Given** 10 operations (3 reads, 5 writes, 2 searches), **When** developer calls `office_get_change_log` with `filter: "write"`, **Then** only the 5 write operations are returned
2. **Given** a change log with timestamps, **When** the response is returned, **Then** each entry includes `timestamp` in ISO 8601 format

---

### User Story 4 — Undo by operation ID (Priority: P2)

A user says "Undo the table insertion but keep the text changes." The LLM calls `office_undo_last` with a specific `operationId` to undo just that one operation.

**Why this priority**: Selective undo gives fine-grained control. Lower priority because it requires undo-queue management in the add-in.

**Independent Test**: Perform 3 operations, undo the middle one by ID, verify only that operation is reversed.

**Acceptance Scenarios**:

1. **Given** 3 write operations with IDs op-1, op-2, op-3, **When** developer calls `office_undo_last` with `operationId: "op-2"`, **Then** only op-2 is undone and op-1 and op-3 remain
2. **Given** a stale operation ID, **When** developer calls `office_undo_last`, **Then** the tool returns `{ undone: false, reason: "Operation already undone or not found" }`

---

### Edge Cases

- **Undo in Word with tracked changes**: Undo rejects the tracked change — the original text is restored. This is the desired behavior.
- **Undo in PowerPoint**: PowerPoint groups all operations within a single `PowerPoint.run()` batch. Undo reverses the entire batch, not individual operations. Document this limitation.
- **Undo in Excel**: Excel supports per-cell undo. The LLM can undo the last batch write.
- **Concurrent operations**: If two tool calls overlap, undo targets the most recent completed operation.
- **Audit log size**: Cap at 1000 entries. Older entries are pruned on a FIFO basis.

## Requirements

### Functional Requirements

- **FR-001**: The MCP server MUST expose `office_get_change_log` accepting optional `instanceId`, `filter` ("read"|"write"|"all"), `lastN`, and `fromTimestamp`.
- **FR-002**: The MCP server MUST expose `office_undo_last` accepting `instanceId` and optional `operationId`.
- **FR-003**: Each audit entry MUST include `operationId`, `timestamp`, `toolName`, `instanceId`, `inputs` (summarized), `outcome` ("success"|"error"|"timeout"), and `duration`.
- **FR-004**: `office_undo_last` MUST trigger `Ctrl+Z` / undo in the target add-in instance.
- **FR-005**: The audit log MUST be scoped per server process (not persisted across restarts).
- **FR-006**: The audit log MUST retain at most 1000 entries, pruning oldest entries first.

### Key Entities

- **Audit Entry**: A single recorded operation with metadata (tool, instance, inputs, outcome, timing).
- **Change Log**: A filtered, ordered view of audit entries, queryable by instance, type, and time range.
- **Undo Operation**: Triggers the host application's native undo (Ctrl+Z) to reverse the last batch.

## Success Criteria

### Measurable Outcomes

- **SC-001**: `office_get_change_log` returns results in under 100ms (server-side data, no add-in round-trip).
- **SC-002**: `office_undo_last` reverses the last operation within 2 seconds.
- **SC-003**: Audit entries are accurate — every tool call produces exactly one entry.
- **SC-004**: All existing tests continue to pass.

## Assumptions

- The existing `AuditLog` class in `Models/AuditLog.cs` can be extended without breaking changes.
- The add-in can trigger programmatic undo via `document.undo()` (Word/Excel) or the native undo stack (PowerPoint).
- Undo granularity: Word = per-tracked-change, Excel = per-batch-write, PowerPoint = per-`PowerPoint.run()` batch.
