# Feature Specification: Phase 8 — SignalR Transport

**Feature Branch**: `008-signalr-transport`

**Created**: 2026-05-16

**Status**: Draft

**Depends on**: Phase 0 (Spike) — infrastructure must exist

**Input**: Upgrade MCP server transport from HTTP polling to SignalR (WebSocket) for real-time command delivery to Office add-ins, plus fix instance ID naming bug.

## Architecture

The MCP server gains ASP.NET Core SignalR alongside the existing HTTP endpoints. The add-in connects to a SignalR hub on startup instead of relying solely on HTTP polling. The server pushes commands instantly via `Hub.Clients.Group(instanceId)`. Results are sent back via hub invocation. HTTP polling remains as a fallback for environments where WebSocket connections fail.

```
Open WebUI                    MCP Server (port 3000)              Office Add-ins
    │                                │                               │
    │── tools/call ─────────────────►│                               │
    │   { tool, { instanceId?, ... } │                               │
    │                                │── SignalR "ExecuteCommand" ──►│  (instant push)
    │                                │◄── hub.invoke("ReportResult") │  (instant return)
    │◄── result ─────────────────────│                               │
    │                                │                               │
    │                                │◄── /instances/register ──────│  (on load)
    │                                │  ── SignalR hub.connect() ──►│  (on load)
    │                                │◄── /instances/:id/heartbeat ─│  (every 10s)
    │                                │                               │
    │   [Fallback path when WebSocket unavailable]                  │
    │                                │►── /instances/:id/commands ──│  (poll every 2s)
    │                                │◄── /instances/:id/result ────│  (after execution)
```

Current latency: 500–2500ms overhead from polling. Target: <20ms overhead with SignalR.

## User Scenarios & Testing

### User Story 1 — LLM gets faster responses (Priority: P0)

An LLM calls a tool and currently waits 1–3 seconds for the add-in to poll. With SignalR, the command is pushed instantly and the result returns in <20ms overhead. This makes the system feel interactive.

**Why this priority**: The entire value proposition of an LLM-driven Office tool depends on latency. Multi-second pauses between tool calls make conversations feel broken. Sub-200ms round-trips enable interactive workflows.

**Independent Test**: Call any tool (e.g., `powerpoint_get_slides`) and measure the time from MCP request to response. With SignalR, it should complete in <200ms vs the current 1.5–3s.

**Acceptance Scenarios**:

1. **Given** an add-in connected via SignalR, **When** the MCP server receives a tool call, **Then** the command is pushed to the add-in via `ExecuteCommand` within 10ms
2. **Given** the add-in completes command execution, **When** it invokes `ReportResult` on the hub, **Then** the MCP server resolves the waiting `TaskCompletionSource` and returns the result to the caller
3. **Given** multiple tool calls in sequence, **When** each returns via SignalR, **Then** total latency per call is <200ms end-to-end

---

### User Story 2 — Instance IDs reflect actual host (Priority: P0)

When a Word instance registers, it should get `word_1` not `powerpoint_1`. Same for Excel (`excel_1`) and Outlook (`outlook_1`). Currently `InstanceRegistry.cs` line 51 always generates `powerpoint_{N}` regardless of `appName`.

**Why this priority**: Correct instance IDs are foundational — tool routing, logging, and multi-host support all depend on knowing which host an instance represents. This is a bug fix that blocks multi-host correctness.

**Independent Test**: Register an instance with `appName: "Word"` and verify the returned instance ID starts with `word_`, not `powerpoint_`.

**Acceptance Scenarios**:

1. **Given** a Word add-in registers with `appName: "Word"`, **When** the server calls `RegisterInstance("Word", docName)`, **Then** the instance ID is `word_1` (or next available `word_N`)
2. **Given** an Excel add-in registers, **When** the server processes registration, **Then** the instance ID starts with `excel_`
3. **Given** a PowerPoint add-in registers, **When** the server processes registration, **Then** the instance ID starts with `powerpoint_`
4. **Given** multiple Word instances register, **When** each calls `RegisterInstance`, **Then** each gets a unique sequential ID (`word_1`, `word_2`, etc.)

---

### User Story 3 — Graceful fallback when WebSocket unavailable (Priority: P1)

In some corporate environments, WebSocket connections may be blocked by firewalls or proxy servers. The add-in should detect SignalR connection failure and fall back to HTTP polling without user intervention.

**Why this priority**: Enterprise environments have unpredictable network policies. The system must work everywhere, not just in WebSocket-friendly networks. Fallback ensures reliability without sacrificing performance where WebSocket is available.

**Independent Test**: Block WebSocket connections (e.g., via browser dev tools or proxy rule), verify the add-in detects the failure and switches to HTTP polling automatically.

**Acceptance Scenarios**:

1. **Given** WebSocket connections are blocked, **When** the add-in starts and SignalR fails to connect, **Then** the add-in falls back to HTTP polling within 5 seconds
2. **Given** the add-in is running on HTTP polling fallback, **When** the user checks the task pane, **Then** the connection state shows "HTTP Polling" (not "Connected")
3. **Given** WebSocket becomes available while on fallback, **When** the add-in's reconnect timer fires, **Then** SignalR reconnects and the task pane updates to "Connected"

---

### User Story 4 — Multiple concurrent commands (Priority: P1)

Currently commands are processed sequentially (the `isProcessingCommands` guard in the add-in). With SignalR, support concurrent command execution so the LLM can issue multiple independent reads in parallel.

**Why this priority**: An LLM often needs to read multiple slides or paragraphs to build context. Sequential processing multiplies latency by the number of reads. Concurrent execution keeps latency constant regardless of batch size.

**Independent Test**: Issue two independent `powerpoint_get_shape_text` calls simultaneously. Verify both execute concurrently and return independently.

**Acceptance Scenarios**:

1. **Given** two independent read commands arrive, **When** the add-in receives both via SignalR, **Then** both execute concurrently without blocking each other
2. **Given** a write command is in progress, **When** a second write command arrives for the same instance, **Then** the second command waits for the first to complete (write serialization)
3. **Given** concurrent read commands, **When** each completes independently, **Then** results are delivered via `ReportResult` with matching command IDs

---

### Edge Cases

- What if SignalR disconnects mid-command? The command should time out normally via `CommandStore.WaitForResult`. If the add-in reconnects and completes the command, the result is delivered via HTTP as a retry.
- What if the add-in opens before the server starts? SignalR automatic reconnect with exponential backoff handles this — the add-in keeps trying until the server becomes available.
- What about multiple add-in windows (same host)? Each gets a unique instanceId and joins its own SignalR group. Commands are routed to the correct group.
- What if the server restarts? The add-in's SignalR connection drops, auto-reconnect kicks in, and the add-in re-registers via HTTP on reconnect.

## Requirements

### Functional Requirements

- **FR-001**: The MCP server MUST expose a SignalR hub endpoint at `/hubs/commands`
- **FR-002**: The add-in MUST connect to the SignalR hub on startup (in addition to existing HTTP registration)
- **FR-003**: The server MUST push commands to the add-in via `Hub.Clients.Group(instanceId).SendAsync("ExecuteCommand", commandId, commandName, args)`
- **FR-004**: The add-in MUST invoke `hub.invoke("ReportResult", commandId, success, error, payload)` to return results
- **FR-005**: `CommandStore.WaitForResult` MUST use `TaskCompletionSource` instead of polling
- **FR-006**: HTTP polling endpoints MUST remain functional as fallback
- **FR-007**: The add-in MUST auto-detect SignalR connection failure and fall back to HTTP polling
- **FR-008**: `InstanceRegistry` MUST generate instance IDs with correct host prefix (`word_`, `excel_`, `outlook_`, `powerpoint_`)
- **FR-009**: The SignalR connection MUST support automatic reconnect with exponential backoff
- **FR-010**: The server MUST clean up SignalR groups when instances disconnect (OnDisconnectedAsync)

### Non-Functional Requirements

- **NFR-001**: Command delivery latency MUST be <50ms (vs current 500–2000ms)
- **NFR-002**: Must not break existing 82 C# tests and 97 TypeScript tests
- **NFR-003**: SignalR MUST use WebSocket transport primarily, Server-Sent Events as fallback
- **NFR-004**: Connection state MUST be visible in the task pane UI

### Key Entities

- **SignalR Hub (`CommandHub`)**: Manages groups by instanceId. Handles command/result relay. Inherits `Hub`. Methods: `ReportResult(commandId, success, error, payload)`, `JoinGroup(instanceId)`. Overrides: `OnDisconnectedAsync` for cleanup.
- **Client Methods**: `ExecuteCommand(commandId, commandName, args)` — pushed from server to add-in via group broadcast.
- **Connection State**: Enum — `Connected` (SignalR active), `Reconnecting` (SignalR retrying), `FallenBack` (HTTP polling active). Displayed in task pane UI.
- **CommandStore (updated)**: Uses `TaskCompletionSource<CommandResult>` per command ID instead of polling loop. Resolved when `ReportResult` arrives via SignalR hub method or HTTP endpoint.

## Success Criteria

### Measurable Outcomes

- **SC-001**: Tool call latency reduced from 1.5–3s to <200ms end-to-end
- **SC-002**: Instance IDs correctly reflect host type (word*, excel*, outlook*, powerpoint*)
- **SC-003**: HTTP polling fallback works when WebSocket is blocked
- **SC-004**: All existing tests (82 C# + 97 TypeScript) continue to pass
- **SC-005**: Task pane shows connection state (WebSocket vs HTTP polling)

## Assumptions

- ASP.NET Core SignalR is compatible with the project's .NET 8 target.
- The Office JS Add-in runtime supports WebSocket connections (confirmed: modern Office on web and desktop supports WebSockets).
- HTTP polling endpoints remain unchanged — they are the compatibility layer.
- The `isProcessingCommands` guard in the add-in can be safely removed or made per-command for concurrent execution.
- SignalR client library (`@microsoft/signalr`) is available for the TypeScript add-in.
