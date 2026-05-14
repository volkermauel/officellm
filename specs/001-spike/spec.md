# Feature Specification: Phase 0 — Spike

**Feature Branch**: `001-spike`

**Created**: 2026-05-14

**Status**: Draft

**Input**: User description: "Create a minimal local MCP Streamable HTTP server exposing office_get_active_app, one PowerPoint VSTO add-in that reports active presentation and selected slide, connect Open WebUI to the local MCP endpoint, validate Docker networking with host.docker.internal if Open WebUI runs in Docker."

## User Scenarios & Testing

### User Story 1 — Developer validates end-to-end flow (Priority: P1)

A developer installs the spike artifacts, launches PowerPoint, and sees that Open WebUI can discover and call the `office_get_active_app` tool through the local MCP server. The tool returns the active host name, document name, and selection metadata.

**Why this priority**: This is the foundational validation — if this doesn't work, no other phase can proceed. It proves the entire architecture: VSTO → IPC → MCP → Open WebUI.

**Independent Test**: Developer runs `office_get_active_app` from Open WebUI chat and receives a valid JSON response with host, document, and selection fields.

**Acceptance Scenarios**:

1. **Given** PowerPoint is open with a presentation, **When** developer calls `office_get_active_app` via Open WebUI, **Then** the tool returns `{"ok": true, "app": "PowerPoint", "documentName": "...", "selection": {...}}`
2. **Given** no Office host is running, **When** developer calls `office_get_active_app`, **Then** the tool returns `{"ok": false, "errorCode": "NO_ACTIVE_OFFICE"}`
3. **Given** Open WebUI runs in Docker on the same workstation, **When** the MCP endpoint is registered as `http://host.docker.internal:<port>/mcp`, **Then** the connection succeeds without network errors

---

### User Story 2 — Spike add-in reports PowerPoint state (Priority: P1)

A developer creates and deploys a minimal PowerPoint VSTO add-in that, when loaded, can report the active presentation name and currently selected slide index through the local MCP server.

**Why this priority**: Validates the VSTO → IPC → MCP pipeline with real Office COM interaction. Proves that the helper process model works.

**Independent Test**: Load the add-in in PowerPoint, call the tool from Open WebUI, verify the reported presentation name and slide index match reality.

**Acceptance Scenarios**:

1. **Given** PowerPoint with a multi-slide presentation, **When** developer selects slide 3, **Then** `office_get_active_app` reports `selectedSlideIndex: 2` (zero-based)
2. **Given** a new blank presentation, **When** developer calls the tool, **Then** the document name is the default PowerPoint name (e.g., "Presentation1")
3. **Given** the MCP helper process is not running, **When** the VSTO add-in loads, **Then** it starts the helper process automatically

---

### User Story 3 — Open WebUI discovers and registers the MCP tool (Priority: P2)

A developer configures Open WebUI's External Tools with the local MCP endpoint. The tool list from `office_get_active_app` appears in the Open WebUI chat interface as an available tool.

**Why this priority**: Validates the Open WebUI → MCP integration path, which is the primary user-facing contract.

**Independent Test**: Register the MCP endpoint in Open WebUI, verify the tool appears in the chat UI and can be invoked.

**Acceptance Scenarios**:

1. **Given** MCP server is running on `127.0.0.1:<port>`, **When** developer registers the endpoint in Open WebUI External Tools, **Then** the tool list refreshes and `office_get_active_app` appears
2. **Given** the MCP server is stopped, **When** developer calls any tool, **Then** Open WebUI shows a connection error (not a silent failure)
3. **Given** Docker networking, **When** endpoint uses `host.docker.internal`, **Then** tools are discoverable from the Docker-hosted Open WebUI instance

---

### Edge Cases

- What happens when multiple Office hosts are running simultaneously (e.g., PowerPoint and Word)? The spike only targets PowerPoint; other hosts should be ignored.
- What happens when the MCP server port is already in use? The server should report an error and suggest an alternative port.
- How does the system behave when Open WebUI is on a different machine? Out of scope for spike — documented as a known limitation.

## Requirements

### Functional Requirements

- **FR-001**: The MCP server MUST expose `office_get_active_app` as a Streamable HTTP tool at `/mcp`.
- **FR-002**: `office_get_active_app` MUST return the active Office host name, document name, and selection metadata in a JSON envelope with `ok`, `app`, `documentId`, and `result` fields.
- **FR-003**: The VSTO add-in MUST connect to the MCP helper process over localhost (HTTP or named pipe).
- **FR-004**: The VSTO add-in MUST start the MCP helper process if it is not already running.
- **FR-005**: The MCP server MUST support Streamable HTTP protocol as used by Open WebUI external tools.
- **FR-006**: The system MUST work when Open WebUI runs in Docker on the same workstation using `host.docker.internal`.
- **FR-007**: The MCP server MUST return error envelopes with `ok: false`, `errorCode`, and `message` fields when no Office host is active.

### Key Entities

- **MCP Helper Process**: Lightweight server process owning MCP transport, tool registry, and request routing.
- **VSTO Add-in**: PowerPoint-only add-in providing COM access, selection state, and confirmation UI scaffolding.
- **IPC Channel**: Communication link between VSTO add-in and MCP helper (localhost HTTP for spike).
- **Tool Envelope**: Standard JSON response structure: `{ok, app, documentId, result, warnings, requiresConfirmation, auditId}`.

## Success Criteria

### Measurable Outcomes

- **SC-001**: A developer can call `office_get_active_app` from Open WebUI and receive a valid response within 3 seconds.
- **SC-002**: The VSTO add-in starts the MCP helper process automatically on first tool invocation.
- **SC-003**: Docker-hosted Open WebUI can reach the MCP endpoint via `host.docker.internal` without port mapping workarounds.
- **SC-004**: The tool response envelope matches the documented schema (all required fields present).

## Assumptions

- Windows desktop Office with VSTO support is available on the developer machine.
- Open WebUI version supports native MCP Streamable HTTP external tools.
- The developer has .NET 8 SDK and Visual Studio (or VS Build Tools) installed for VSTO development.
- No authentication is needed for localhost-only spike testing.
- One MCP helper process per user session is acceptable for the spike.
