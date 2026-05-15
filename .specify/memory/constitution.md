# Office LLM Harness — Constitution

## Core Principles

### I. Safety-First (NON-NEGOTIABLE)

No arbitrary COM execution, shell execution, or macro execution may be exposed as model-callable tools. Destructive operations MUST return previews and require explicit user confirmation before applying changes. All tool outputs are bounded in size to prevent excessive context leakage.

### II. Local-Only Execution (NON-NEGOTIABLE)

The MCP server binds exclusively to `127.0.0.1` for MVP. No remote network exposure. Office automation runs only within the user's active session. The model never sends entire documents by default — context is always selection-scoped and bounded.

### III. User Control & Transparency (NON-NEGOTIABLE)

The model proposes; the user approves. All document mutations go through a confirmation gate with before/after diffs. An audit log records every tool call with timestamp, correlation fields, document handle, mutation status, and confirmation ID. Undo support is mandatory for all mutation tools.

### IV. Minimal Surface Area

Tools expose narrow, intentional capabilities with typed inputs and bounded outputs. Never expose broader functionality than a tool requires. Tool naming uses stable prefixes (`office_`, `word_`, `excel_`, `powerpoint_`, `outlook_`). Output formats are structured (JSON envelopes) for deterministic parsing.

### V. Phased Delivery

Each Office host is implemented as an independent MVP phase. Each phase delivers a working, testable subset of tools. No phase depends on another host being complete. The architecture is designed so server-side MCP logic can be reused across hosts.

## Technology Constraints

- **Platform**: Windows desktop Office only (VSTO + COM). No Web, Mac, or mobile in MVP.
- **MCP Transport**: Streamable HTTP over localhost. Named pipes preferred for production IPC.
- **Language**: C# / .NET 8 for VSTO add-ins and MCP server.
- **LLM Frontend**: Open WebUI for chat, model routing, and MCP external tool registration.
- **Audit Storage**: Local JSONL or SQLite for MVP; centralized logging later.

## Development Workflow

- All specs follow speckit workflow: `/speckit.specify` → `/speckit.plan` → `/speckit.tasks` → `/speckit.implement`.
- Each phase is an independent feature branch (`###-phase-name`).
- Testing gates: tool schema validation, safety gate verification, Office host smoke tests before any phase can be marked complete.
- Prompt injection defense: system prompts MUST instruct the model not to follow embedded instructions from document content.

## Governance

This constitution supersedes all other development practices. Amendments require documenting the change, its rationale, and migration impact. All PRs and reviews must verify compliance with Safety-First, Local-Only Execution, and User Control principles.

**Version**: 1.0.0 | **Ratified**: 2026-05-14 | **Last Amended**: 2026-05-14
