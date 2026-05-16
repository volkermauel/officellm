# Office LLM Harness — Development Rules

## ⚠️ MANDATORY: Tests Must Pass Before Commit

**Before any commit, ALL automated tests must pass.** This is non-negotiable.

```bash
# Run all tests (C# unit + integration tests)
dotnet test tests/mcp-server.Tests/

# Build everything first
dotnet build src/mcp-server/
cd src/powerpoint-addin && npx webpack --mode production
```

If tests fail — **do not commit**. Fix the root cause, re-run, and only commit when green.

## Implementation Workflow (TDD)

When implementing new features or fixing bugs, follow this order:

1. **Architecture** — Understand the codebase, identify affected modules, plan the approach
2. **Tests first** — Write tests that express the expected behavior _before_ implementation
3. **Implementation** — Write the code to make those tests pass
4. **Verification** — Run the full test suite to confirm nothing is broken

```bash
# 1. Understand existing tests
rg "test|Test" tests/ --include='*.cs'

# 2. Add new tests (red)
#    Write failing tests first, verify they fail

# 3. Implement (green)
#    Write code to make tests pass

# 4. Verify full suite
dotnet test tests/mcp-server.Tests/ --verbosity normal
```

## Project Structure

```
src/
├── mcp-server/           # .NET 8 MCP server (hub)
│   ├── AppBuilder.cs     # HTTP app factory (testable)
│   ├── Program.cs        # Entry point + stdio transport
│   ├── Models/
│   │   ├── InstanceRegistry.cs  # Thread-safe instance tracking
│   │   ├── CommandStore.cs      # Command queue with wait/timeout
│   │   └── McpResponse.cs       # Standard MCP response envelope
│   └── Tools/
│       ├── OfficeTools.cs     # Tool implementations
│       └── McpToolEngine.cs   # Tool dispatch + definitions
├── powerpoint-addin/     # Office JS PowerPoint Add-in
│   ├── src/
│   │   ├── app.ts            # Main entry point
│   │   └── communication.ts  # MCP client (register, poll, heartbeat)
│   └── package.json          # webpack build
tests/
└── mcp-server.Tests/         # xUnit test suite
    ├── InstanceRegistryTests.cs
    ├── CommandStoreTests.cs
    ├── McpResponseTests.cs
    ├── McpToolEngineTests.cs
    └── HttpEndpointTests.cs      # WebApplicationFactory integration tests
specs/                        # Speckit specifications
```

## Available Endpoints

| Transport       | Endpoint                        | Purpose                                           |
| --------------- | ------------------------------- | ------------------------------------------------- |
| Streamable HTTP | `POST /mcp`                     | MCP protocol (initialize, tools/list, tools/call) |
| REST Bridge     | `GET /openapi.json`             | OpenAPI 3.0 spec for non-MCP clients              |
| REST Bridge     | `POST /api/{toolName}`          | Direct REST calls to any MCP tool                 |
| Swagger UI      | `GET /docs`                     | Interactive API documentation                     |
| Instance mgmt   | `POST /instances/register`      | Add-in registration                               |
| Instance mgmt   | `POST /instances/:id/heartbeat` | Keep-alive (every 10s)                            |
| Instance mgmt   | `GET /instances/:id/commands`   | Poll for pending commands                         |
| Instance mgmt   | `POST /instances/:id/result`    | Report command results                            |
| Health          | `GET /health`                   | Server health check                               |
| Stdio           | stdin/stdout                    | MCPo-compatible JSON-RPC transport                |

## Build Commands

```bash
# Build MCP server
dotnet build src/mcp-server/

# Build PowerPoint add-in
cd src/powerpoint-addin && npx webpack --mode production

# Run tests
dotnet test tests/mcp-server.Tests/

# Full pipeline (build + test)
dotnet build src/mcp-server/ && cd src/powerpoint-addin && npx webpack --mode production && cd ../../ && dotnet test tests/mcp-server.Tests/
```

## Running the Server

```bash
# Default port 3000
dotnet run --project src/mcp-server/

# Custom port
dotnet run --project src/mcp-server/ 8080
```

## Word JS API Rules

When working with the Word JavaScript API in the add-in:

- **`Word.run()` is lowercase** — `Word.run(async (context) => { ... })`, not `Word.Run()`.
- **`context.sync()` is required before reading properties** — After loading objects with `.load()`, you MUST call `await context.sync()` before accessing loaded properties. Failing to sync results in "property not loaded" errors.
- **changeTrackingMode pattern for mutations** — Word tracked changes use `document.changeTrackingMode`:
  ```typescript
  // Standard mutation pattern:
  // 1. Save current mode
  // 2. Set TrackMineOnly
  // 3. Perform mutation
  // 4. Restore original mode
  // 5. Return { tracked: true }
  const originalMode = context.document.changeTrackingMode;
  context.document.changeTrackingMode = Word.ChangeTrackingMode.trackMineOnly;
  // ... perform mutations ...
  context.document.changeTrackingMode = originalMode;
  await context.sync();
  ```
- **`changeTrackingMode` values**: `"Off"` | `"TrackAll"` | `"TrackMineOnly"` (requires WordApi 1.4+)
- **`getTextFrameOrNullObject` equivalent** — Word has no direct equivalent. Use `range.getHtml()` or `range.getText()` to read content. For null-safe patterns, check `range.isNullObject` after `context.sync()`.
- **Range-based operations** — Word operates on `Range` objects. Get the current selection via `context.document.getSelection()`, then manipulate it as a range.

### Tool Inventory (29 tools)

| # | Tool Name | Host | Category |
|---|-----------|------|----------|
| 1 | `powerpoint_get_slides` | PowerPoint | Read |
| 2 | `powerpoint_get_slide` | PowerPoint | Read |
| 3 | `powerpoint_get_shapes` | PowerPoint | Read |
| 4 | `powerpoint_get_shape` | PowerPoint | Read |
| 5 | `powerpoint_get_shape_text` | PowerPoint | Read |
| 6 | `powerpoint_set_shape_text` | PowerPoint | Write |
| 7 | `powerpoint_add_shape` | PowerPoint | Write |
| 8 | `powerpoint_delete_shape` | PowerPoint | Write |
| 9 | `powerpoint_set_shape_properties` | PowerPoint | Write |
| 10 | `powerpoint_add_slide` | PowerPoint | Write |
| 11 | `powerpoint_delete_slide` | PowerPoint | Write |
| 12 | `powerpoint_move_slide` | PowerPoint | Write |
| 13 | `powerpoint_duplicate_slide` | PowerPoint | Write |
| 14 | `powerpoint_set_slide_layout` | PowerPoint | Write |
| 15 | `powerpoint_get_notes` | PowerPoint | Read |
| 16 | `powerpoint_set_notes` | PowerPoint | Write |
| 17 | `powerpoint_get_selection` | PowerPoint | Read |
| 18 | `powerpoint_set_selection` | PowerPoint | Write |
| 19 | `office_get_active_app` | Shared | Read |
| 20 | `word_get_outline` | Word | Read |
| 21 | `word_get_paragraphs` | Word | Read |
| 22 | `word_replace_text` | Word | Write (tracked) |
| 23 | `word_insert_after_heading` | Word | Write (tracked) |
| 24 | `word_add_review_comments` | Word | Write |
| 25 | `word_get_selection` | Word | Read |
| 26 | `word_get_tracked_changes` | Word | Read |
| 27 | `word_accept_all_changes` | Word | Write |
| 28 | `word_reject_all_changes` | Word | Write |
| 29 | `office_get_document_context` | Shared | Read |

**Mutation modes by host**:
- **Word**: Tracked changes (`changeTrackingMode: "TrackMineOnly"`) — user accepts/rejects via Word Review ribbon or tracked change tools. No confirmation gate needed.
- **PowerPoint**: Direct write with undo group per `PowerPoint.run()` batch — no tracked changes API exists. Undo (Ctrl+Z) reverses the entire batch.

## Speckit Workflow

For feature development, follow the speckit workflow:

1. `/speckit.specify` — Refine the specification
2. `/speckit.plan` — Create implementation plan
3. `/speckit.tasks` — Generate actionable tasks
4. `/speckit.implement` — Execute implementation

See `specs/README.md` for phase details.

## Word JS API Quick Reference

```
changeTrackingMode: "Off" | "TrackAll" | "TrackMineOnly"  (WordApi 1.4+)

Mutation pattern:
  save mode → set TrackMineOnly → mutate → restore mode → return { tracked: true }

PowerPoint: NO tracked changes API. Direct-write with undo grouped per PowerPoint.run() batch.
```
