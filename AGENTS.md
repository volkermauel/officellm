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

## Speckit Workflow

For feature development, follow the speckit workflow:

1. `/speckit.specify` — Refine the specification
2. `/speckit.plan` — Create implementation plan
3. `/speckit.tasks` — Generate actionable tasks
4. `/speckit.implement` — Execute implementation

See `specs/README.md` for phase details.
