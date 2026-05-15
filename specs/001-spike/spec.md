# Feature Specification: Phase 0 — Spike

**Feature Branch**: `001-spike`

**Created**: 2026-05-14

**Status**: Draft

**Input**: User description: "Create a minimal local MCP server as central hub, one PowerPoint Office JS Add-in that registers with the MCP server, connect Open WebUI to the MCP endpoint, validate end-to-end flow."

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
- **Add-ins connect TO server**: Office JS Add-in registers on load, polls for commands
- **Fixed tool list**: Tools have optional `instanceId` parameter (defaults to most recent instance)
- **Multiple instances**: Each PowerPoint session gets its own registered instance
- **Transport**: Streamable HTTP + stdio (MCPo compatible)

## User Scenarios & Testing

### User Story 1 — Developer validates end-to-end flow (Priority: P1)

A developer starts the MCP server, opens PowerPoint with the Office JS Add-in loaded, and calls `office_get_active_app` from Open WebUI. The add-in auto-registers with the MCP server, and the tool returns information about all registered instances.

**Why this priority**: This is the foundational validation — if this doesn't work, no other phase can proceed. It proves the entire architecture: MCP Server → Add-in registration → Open WebUI integration.

**Independent Test**: Developer runs `office_get_active_app` from Open WebUI chat and receives a valid JSON response listing registered instances.

**Acceptance Scenarios**:

1. **Given** PowerPoint is open with the add-in loaded, **When** developer calls `office_get_active_app` via Open WebUI, **Then** the tool returns a list of registered instances with app name, document name, and instance ID
2. **Given** no Office host is running, **When** developer calls `office_get_active_app`, **Then** the tool returns an error message indicating no instances are registered
3. **Given** Open WebUI runs in Docker on the same workstation, **When** the MCP endpoint is registered as `http://host.docker.internal:3000/mcp`, **Then** the connection succeeds without network errors

---

### User Story 2 — Add-in registers and stays alive (Priority: P1)

A developer loads the PowerPoint Office JS Add-in. The add-in automatically registers itself with the MCP server, receives a unique instance ID (e.g., `powerpoint_1`), and begins sending heartbeats every 10 seconds to stay alive. If the add-in stops responding for 30 seconds, the MCP server marks it as timed out.

**Why this priority**: Validates the registration and heartbeat mechanism that enables multiple concurrent Office sessions.

**Independent Test**: Load the add-in, check MCP server logs for registration, verify heartbeat endpoint receives periodic requests, verify timeout cleanup works.

**Acceptance Scenarios**:

1. **Given** PowerPoint with a presentation open, **When** the add-in loads, **Then** the MCP server registers an instance with ID `powerpoint_1` and documents the presentation name
2. **Given** the add-in is running, **When** 10 seconds pass, **Then** the MCP server receives a heartbeat and updates the instance's last-seen timestamp
3. **Given** the add-in closes without unregistering, **When** 30 seconds pass, **Then** the MCP server marks the instance as timed out and removes it from active instances
4. **Given** two PowerPoint presentations are open, **When** both add-ins load, **Then** the MCP server registers two separate instances (`powerpoint_1` and `powerpoint_2`)

---

### User Story 3 — Open WebUI discovers and calls tools (Priority: P1)

A developer configures Open WebUI's External Tools with the local MCP endpoint at `http://127.0.0.1:3000/mcp`. The tool list appears in Open WebUI, including `office_get_active_app` and placeholder tools for PowerPoint operations. The developer invokes a tool and receives a response routed through the MCP server to the appropriate add-in instance.

**Why this priority**: Validates the Open WebUI → MCP integration path, which is the primary user-facing contract.

**Independent Test**: Register the MCP endpoint in Open WebUI, verify all tools appear in the chat UI, invoke a tool and verify response.

**Acceptance Scenarios**:

1. **Given** MCP server is running, **When** developer registers the endpoint in Open WebUI External Tools, **Then** the tool list refreshes and all tools appear (`office_get_active_app`, `powerpoint_get_deck_outline`, etc.)
2. **Given** an add-in is registered, **When** developer calls `powerpoint_get_deck_outline` without specifying `instanceId`, **Then** the tool is routed to the most recently registered instance
3. **Given** multiple instances are registered, **When** developer calls a tool with `{"instanceId": "powerpoint_2", ...}`, **Then** the tool is routed specifically to that instance
4. **Given** the MCP server is stopped, **When** developer calls any tool, **Then** Open WebUI shows a connection error (not a silent failure)

---

### Edge Cases

- What happens when multiple Office hosts are running simultaneously (e.g., PowerPoint and Word)? The spike only targets PowerPoint; other hosts will register with different prefixes (e.g., `word_1`).
- What happens when the MCP server port is already in use? The server should report an error and exit gracefully.
- How does the system behave when Open WebUI is on a different machine? Out of scope for spike — documented as a known limitation.
- What happens when an add-in crashes without unregistering? The heartbeat timeout (30s) handles this automatically.

## Requirements

### Functional Requirements

- **FR-001**: The MCP server MUST run on a single configurable port (default: 3000) and expose both Streamable HTTP (`/mcp`) and stdio transports.
- **FR-002**: The MCP server MUST expose `office_get_active_app` as a tool that returns a list of all registered Office instances.
- **FR-003**: The MCP server MUST expose `powerpoint_get_deck_outline` with optional `instanceId`, `includeSpeakerNotes`, and `includeHiddenSlides` parameters.
- **FR-004**: The Office JS Add-in MUST register itself with the MCP server on load via POST `/instances/register`.
- **FR-005**: The Office JS Add-in MUST send heartbeats every 10 seconds via POST `/instances/{id}/heartbeat`.
- **FR-006**: The Office JS Add-in MUST poll for pending commands every 2 seconds via GET `/instances/{id}/commands`.
- **FR-007**: The Office JS Add-in MUST report command results via POST `/instances/{id}/result`.
- **FR-008**: The MCP server MUST auto-cleanup timed-out instances (no heartbeat for 30 seconds).
- **FR-009**: Tool calls with no `instanceId` MUST default to the most recently registered active instance.
- **FR-010**: The MCP server MUST support MCPo integration via stdio JSON-RPC transport.

### Key Entities

- **MCP Server**: Central hub process owning MCP transport, instance registry, command queue, and tool execution logic.
- **Office JS Add-in**: PowerPoint add-in (not VSTO) running in the task pane iframe. Registers with MCP server, polls for commands, reports results.
- **Instance Registry**: Thread-safe dictionary mapping instance IDs to Office application metadata.
- **Command Store**: Queue for pending commands dispatched to specific instances; supports claiming, result waiting, and cleanup.
- **Tool Envelope**: Standard JSON response structure: `{content: [{type, text}], isError}`.

## Success Criteria

### Measurable Outcomes

- **SC-001**: A developer can call `office_get_active_app` from Open WebUI and receive a valid response listing registered instances within 3 seconds.
- **SC-002**: The Office JS Add-in registers with the MCP server automatically on load and receives an instance ID.
- **SC-003**: Docker-hosted Open WebUI can reach the MCP endpoint via `http://host.docker.internal:3000/mcp` without port mapping workarounds.
- **SC-004**: The MCP server auto-removes instances that stop sending heartbeats within 30 seconds.
- **SC-005**: Two concurrent PowerPoint presentations register as separate instances and can be targeted independently via `instanceId` parameter.

## Assumptions

- Windows desktop Office with Office JS Add-in support (PowerPoint 2019+ or Microsoft 365) is available on the developer machine.
- Open WebUI version supports native MCP Streamable HTTP external tools (v0.6.31+).
- The developer has .NET 8 SDK for building the MCP server and Node.js/npm for building the add-in.
- No authentication is needed for localhost-only spike testing.
- The MCP server handles multiple concurrent add-in instances without performance degradation.
