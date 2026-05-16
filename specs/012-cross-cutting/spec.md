# Feature Specification: Phase 12 — Cross-Cutting Improvements

**Feature Branch**: `012-cross-cutting`

**Created**: 2026-05-16

**Status**: Draft

**Depends on**: All MVP phases complete

**Input**: User description: "Unified document context across all Office hosts, batch read operations for parallel tool invocation, tool suggestions based on active host, and consistent error handling with stable error codes across all tools."

## Architecture

Cross-cutting tools work across all Office hosts. They leverage the existing instance registration system to provide unified context and improved LLM experience. The `office_get_document_context` tool auto-detects the active host and dispatches a host-specific context-gathering command. Batch operations allow the MCP server to fan out multiple tool calls to one or more add-in instances in parallel, reducing round-trip latency.

```
LLM Client                     MCP Server (port 3000)              Office Add-ins
    │                                │                               │
    │── tools/call ─────────────────►│                               │
    │   { office_get_document_context│                               │
    │     { instanceId } }           │                               │
    │                                │── detect host type ──────────►│
    │                                │── dispatch context command ──►│
    │                                │◄── host-specific context ─────│
    │◄── unified context ───────────│                               │
    │                                │                               │
    │── tools/call ─────────────────►│                               │
    │   { office_batch_call          │                               │
    │     { calls: [...] } }         │                               │
    │                                │── fan out N commands ─────────►│
    │                                │◄── collect N results ─────────│
    │◄── combined results ──────────│                               │
    │                                │◄── /instances/register ──────│  (on load)
    │                                │◄── /instances/:id/heartbeat ─│  (every 10s)
    │                                │►── /instances/:id/commands ──│  (poll every 2s)
    │                                │◄── /instances/:id/result ────│  (after execution)
```

## User Scenarios & Testing

### User Story 1 — Unified document context (Priority: P0)

A user starts a conversation with "What am I looking at?" The LLM calls `office_get_document_context` and receives a rich summary: which app is active, which document is open, document metadata (title, page count for Word, sheet count for Excel, slide count for PowerPoint), active selection, and recent edits. One call gives the LLM everything it needs to understand the user's environment without needing host-specific tool knowledge.

**Why this priority**: Without a unified context call, the LLM must first call `office_get_active_app` to determine the host, then call host-specific tools to gather context. This wastes two or more round-trips before any useful work begins. A single unified context call dramatically improves the first-interaction experience.

**Independent Test**: Call `office_get_document_context` from each host (Word, Excel, PowerPoint, Outlook) and verify the response contains host-appropriate metadata.

**Acceptance Scenarios**:

1. **Given** a Word document with 15 pages and 3 headings is open, **When** developer calls `office_get_document_context`, **Then** the response includes document name, word count, page count, top heading outline, selection text, and change tracking status
2. **Given** an Excel workbook with 4 sheets is open, **When** developer calls `office_get_document_context`, **Then** the response includes workbook name, sheet names with used range sizes, active sheet name, and active cell address
3. **Given** a PowerPoint presentation with 12 slides is open, **When** developer calls `office_get_document_context`, **Then** the response includes presentation name, slide count, current slide index and title, and selection type
4. **Given** an Outlook inbox is open with an email selected, **When** developer calls `office_get_document_context`, **Then** the response includes current folder name, selected item subject/sender/date, mailbox name, and delegate status
5. **Given** multiple instances of the same host (two Word windows), **When** developer calls `office_get_document_context`, **Then** the response includes all instances with an indication of which is active

---

### User Story 2 — Batch read operations (Priority: P1)

A user asks "Show me slides 1, 3, and 5." The LLM calls `office_batch_call` with multiple tool invocations. The server dispatches all commands in parallel to the add-in, collects the results, and returns them as a combined response. This avoids N sequential round-trips that would otherwise be needed.

**Why this priority**: Multi-item queries are a common LLM interaction pattern. Without batch support, the LLM must make N sequential calls, each incurring a full request-response cycle. Batch operations reduce latency from O(N _ roundTrip) to O(1 _ roundTrip + max(executionTime)).

**Independent Test**: Call `office_batch_call` with 3 read operations and verify all results are returned in a single response.

**Acceptance Scenarios**:

1. **Given** a presentation with 10 slides, **When** developer calls `office_batch_call` with `powerpoint_get_slide` for slides 1, 3, and 5, **Then** the response includes all three slide shapes in a single result
2. **Given** a batch with 5 read operations, **When** one operation fails, **Then** the response includes successful results for the other 4 and an error entry for the failed one
3. **Given** a batch exceeding the maximum (10 operations), **When** developer calls `office_batch_call`, **Then** the response returns an error with `errorCode: "INVALID_PARAMETER"`
4. **Given** mixed host operations (Word + Excel), **When** developer calls `office_batch_call`, **Then** the server routes each operation to the correct instance

---

### User Story 3 — Tool suggestions based on context (Priority: P2)

A user asks "What can you do with this spreadsheet?" The LLM calls `office_suggest_tools` and receives a list of relevant tools for the current host, each with a brief description and example invocation. This helps the LLM discover capabilities without relying solely on the tool definitions.

**Why this priority**: Tool discovery improves the LLM's ability to self-serve. While tool definitions are available via `tools/list`, they lack usage context. Suggested tools with examples help the LLM choose the right tool and format the correct parameters on the first try.

**Independent Test**: Call `office_suggest_tools` from each host and verify the response lists host-appropriate tools with descriptions and examples.

**Acceptance Scenarios**:

1. **Given** Excel is the active host, **When** developer calls `office_suggest_tools`, **Then** the response lists all Excel tools with descriptions and at least one example invocation per tool
2. **Given** Word is the active host, **When** developer calls `office_suggest_tools`, **Then** the response lists all Word tools including tracked-change tools
3. **Given** PowerPoint is the active host, **When** developer calls `office_suggest_tools`, **Then** the response lists all PowerPoint tools organized by category (Read, Write)
4. **Given** a `category` filter of "Read", **When** developer calls `office_suggest_tools`, **Then** the response includes only read-category tools for the active host

---

### User Story 4 — Consistent error handling (Priority: P0)

All tools return errors in a consistent format: `{ error: string, errorCode: string, details?: object }`. Error codes are documented and stable so the LLM can learn to recover from common failure modes.

**Why this priority**: Inconsistent error formats force the LLM to parse free-form error messages, leading to unreliable recovery strategies. Stable error codes enable the LLM to build a reliable mental model of failure modes and appropriate recovery actions.

**Independent Test**: Trigger each error code and verify the response matches the documented format.

**Acceptance Scenarios**:

1. **Given** an unregistered `instanceId`, **When** any tool is called, **Then** the response includes `{ error: "...", errorCode: "INSTANCE_NOT_FOUND", details: { instanceId } }`
2. **Given** a missing required parameter, **When** any tool is called, **Then** the response includes `{ error: "...", errorCode: "MISSING_PARAMETER", details: { parameter: "slideIndex" } }`
3. **Given** a parameter value out of range (e.g., slide index 999), **When** any tool is called, **Then** the response includes `{ error: "...", errorCode: "INVALID_PARAMETER", details: { parameter: "slideIndex", value: 999, validRange: "1-12" } }`
4. **Given** the add-in does not respond within the timeout, **When** any tool is called, **Then** the response includes `{ error: "...", errorCode: "TIMEOUT", details: { timeoutMs: 30000 } }`
5. **Given** a write operation requiring confirmation, **When** the tool is called without a confirmation token, **Then** the response includes `{ error: "...", errorCode: "CONFIRMATION_REQUIRED", details: { preview: "..." } }`

---

### User Story 5 — Document metadata across hosts (Priority: P1)

A user asks "How long is this document?" The LLM calls `office_get_document_stats` and receives host-specific quantifiable metrics: Word (word count, page count, paragraph count), Excel (sheet count, cell count, file size), PowerPoint (slide count, shape count), Outlook (email count in current view).

**Why this priority**: Document stats are a common first question. Unlike `office_get_document_context` which returns rich narrative context, `office_get_document_stats` returns focused, quantifiable metrics suitable for comparison and tracking.

**Independent Test**: Call `office_get_document_stats` from each host and verify the response contains the documented metrics.

**Acceptance Scenarios**:

1. **Given** a Word document, **When** developer calls `office_get_document_stats`, **Then** the response includes word count, page count, paragraph count, and character count
2. **Given** an Excel workbook, **When** developer calls `office_get_document_stats`, **Then** the response includes sheet count, total used cell count, and file size
3. **Given** a PowerPoint presentation, **When** developer calls `office_get_document_stats`, **Then** the response includes slide count, total shape count, and total text character count
4. **Given** an Outlook mailbox, **When** developer calls `office_get_document_stats`, **Then** the response includes current folder item count and total unread count

---

### Edge Cases

- **Multiple instances of same host** (two Word windows): `office_get_document_context` returns context for all instances and indicates which is active (most recently registered or last heartbeat).
- **Document not yet saved** (Untitled): Return "Untitled" as the document name.
- **Very large document** (1000+ pages): Context gathering must complete within 3 seconds. Limit heading outline to top 10 entries. For stats, approximate where exact counts are expensive.
- **Offline/disconnected state**: If the add-in is unreachable, return the last known context with a `stale: true` flag and the original capture timestamp.
- **Batch with mixed success/failure**: Partial results must be returned for successful operations; failed operations must include the error code and message alongside the successful results.
- **Batch targeting non-existent instances**: Return error per-operation without failing the entire batch.

## Requirements

### Functional Requirements

- **FR-001**: `office_get_document_context` MUST auto-detect the active host and return host-appropriate context.
- **FR-002**: For Word: `office_get_document_context` MUST return document name, word count, page count, heading outline (top 5), selection text, and change tracking status.
- **FR-003**: For Excel: `office_get_document_context` MUST return workbook name, sheet names with used range sizes, active sheet, and active cell.
- **FR-004**: For PowerPoint: `office_get_document_context` MUST return presentation name, slide count, current slide index and title, and selection type.
- **FR-005**: For Outlook: `office_get_document_context` MUST return current folder, selected item subject/sender/date, mailbox name, and delegate status.
- **FR-006**: `office_get_document_stats` MUST return quantifiable metrics per host as documented in User Story 5.
- **FR-007**: All error responses MUST include `error` (string) and `errorCode` (stable string enum) fields.
- **FR-008**: Error code enum MUST include: `INSTANCE_NOT_FOUND`, `MISSING_PARAMETER`, `INVALID_PARAMETER`, `TIMEOUT`, `HOST_NOT_AVAILABLE`, `PERMISSION_DENIED`, `RANGE_TOO_LARGE`, `INVALID_FORMULA`, `CONFIRMATION_REQUIRED`.
- **FR-009**: `office_get_document_context` MUST complete within 3 seconds for any host.
- **FR-010**: Context response MUST include a `timestamp` field indicating when the context was captured.
- **FR-011**: `office_batch_call` MUST accept an array of up to 10 tool invocations and execute them in parallel.
- **FR-012**: `office_batch_call` MUST return per-operation results, preserving the order of the input array.
- **FR-013**: `office_batch_call` MUST support routing each operation to the correct add-in instance when multiple instances are registered.
- **FR-014**: `office_suggest_tools` MUST return host-appropriate tool listings with descriptions and example invocations.
- **FR-015**: Error responses MAY include a `details` object with additional context (parameter name, valid range, timeout value, etc.).

### Error Code Reference

| Code                    | Meaning                                    | Recovery                                 |
| ----------------------- | ------------------------------------------ | ---------------------------------------- |
| `INSTANCE_NOT_FOUND`    | instanceId not registered or timed out     | Call `office_get_active_apps`            |
| `MISSING_PARAMETER`     | Required parameter missing                 | Check tool schema                        |
| `INVALID_PARAMETER`     | Parameter value out of range or wrong type | Check bounds/type                        |
| `TIMEOUT`               | Add-in did not respond in time             | Retry once, check connection             |
| `HOST_NOT_AVAILABLE`    | Office host crashed or add-in unloaded     | Re-register instance                     |
| `PERMISSION_DENIED`     | Operation not allowed (policy/permissions) | Inform user, suggest manual action       |
| `RANGE_TOO_LARGE`       | Range exceeds size limit                   | Reduce range, paginate                   |
| `INVALID_FORMULA`       | Formula syntax error                       | Fix formula syntax                       |
| `CONFIRMATION_REQUIRED` | Write operation needs human approval       | Obtain confirmation token from task pane |

### Key Entities

- **Document Context**: A unified snapshot of the current Office environment including host type, document metadata, selection state, and host-specific details.
- **Document Stats**: Quantifiable metrics about the current document, varying by host type.
- **Batch Request**: An ordered array of tool invocations dispatched in parallel, with ordered results.
- **Error Response**: A structured error with stable `errorCode` string, human-readable `error` message, and optional `details` object.
- **Tool Suggestion**: A tool listing entry with name, category, description, and example invocation.

## Success Criteria

### Measurable Outcomes

- **SC-001**: `office_get_document_context` returns within 3 seconds for any host (Word, Excel, PowerPoint, Outlook).
- **SC-002**: Error responses are consistent across all 39+ tools — every error includes `error` and `errorCode` fields.
- **SC-003**: Context response gives the LLM enough information to be useful without additional tool calls in at least 80% of common scenarios.
- **SC-004**: All existing tests continue to pass (no regressions from error format standardization).
- **SC-005**: `office_batch_call` with 5 read operations returns in less time than 5 sequential calls (measurable latency reduction).
- **SC-006**: Every documented error code is triggered by at least one test case.

## Assumptions

- The unified Office JS Add-in is deployed and functional in Word, Excel, PowerPoint, and Outlook.
- Instance registration via `POST /instances/register` and heartbeat via `POST /instances/:id/heartbeat` are already implemented and reliable.
- The add-in can detect its host via `Office.onReady(info.host)` and dispatch host-specific context-gathering logic.
- Each host's JavaScript API provides the necessary metadata (word count, page count, slide count, etc.) within the 3-second budget.
- Batch operations target a single add-in instance per operation; cross-instance batching routes based on `instanceId` per operation.
- Error code standardization applies to both new cross-cutting tools and existing host-specific tools (retrofitted where needed).
