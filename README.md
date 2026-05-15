# Office LLM Harness

A Windows desktop Office add-in harness that exposes controlled document interaction tools to Open WebUI through a local MCP server, enabling LLM-assisted workflows in Word, Excel, PowerPoint and Outlook.

## Architecture

```
+--------------------+        +-------------------------+
| Word / Excel /     |        | Open WebUI              |
| PowerPoint /       |        | - Chat UI               |
| Outlook            |        | - Model routing         |
+---------+----------+        | - MCP external tools    |
          | Office JS API                  |
          v                               | MCP Streamable HTTP
+---------+----------+                    |
| Office JS Add-in   |                    |
| - Task pane UI     |                    |
| - Office API       |                    |
| - Confirmation UI  |                    |
+---------+----------+                    |
          | localhost HTTP                  |
          v                               |
+---------+-------------------------------+-------------+
| Local Office MCP Server (.NET 8)                      |
| - MCP protocol endpoint (port 3000)                   |
| - Tool registry & command dispatch                    |
| - Instance registry with heartbeat tracking           |
| - Two-phase confirmation for mutations                |
| - Audit log (JSONL)                                   |
+------------------------------------------------------+
```

### Data Flow

1. **LLM calls tool** → `POST /mcp` with `tools/call` method
2. **MCP server queues command** → stored in `CommandStore`, add-in polls `GET /instances/{id}/commands`
3. **Add-in executes** → `PowerPoint.run()` calls Office JS API
4. **Add-in reports result** → `POST /instances/{id}/result`
5. **MCP server returns** → result passed back to LLM

For **mutation tools** (`update_shape_text`, `update_speaker_notes`), step 2–4 is replaced by a two-phase confirmation flow with diff preview.

## Components

| Component             | Language        | Description                                                                                                                               |
| --------------------- | --------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| **MCP Server**           | C# (.NET 8)     | Self-contained executable exposing MCP tools over Streamable HTTP. Command dispatch, instance registry, confirmation flow, audit logging. |
| **Unified Office Add-in** | TypeScript/HTML | Single Office JS Add-in that auto-detects host (Word/Excel/PowerPoint/Outlook) via `Office.onReady()`. One manifest for all hosts. |
| **Express Server**       | Node.js         | Serves static add-in files + dynamic `manifest.xml` (URLs from Host header). Docker/K8s deployment.                                       |

## Project Structure

```
src/
├── mcp-server/           # .NET 8 MCP server
│   ├── OfficeMcpServer.csproj
│   ├── Program.cs        # Entry point, MCP endpoints, bridge server
│   ├── Models/
│   │   └── McpResponse.cs
│   └── Tools/
│       └── OfficeTools.cs
├── powerpoint-addin/     # Unified Office JS Add-in (all hosts)
│   ├── manifest.xml      # Unified manifest (Presentation + Document + Workbook + Mailbox)
│   ├── package.json
│   ├── webpack.config.js
│   ├── tsconfig.json
│   └── src/
│       ├── index.html    # Task pane UI
│       ├── app.ts        # Main entry point
│       └── communication.ts # Office API wrappers + HTTP client
scripts/
├── build.sh              # Build script (all/mcp/addin/dev)
└── dev.sh                # Development server
specs/                    # Speckit specifications
├── README.md
├── 001-spike/
├── 002-powerpoint-mvp/
├── 003-word-mvp/
├── 004-excel-mvp/
└── 005-outlook-mvp/
```

## Quick Start

### Prerequisites

- **Windows** with Office desktop (PowerPoint 2019+ or Microsoft 365)
- **.NET 8 SDK** (for building the MCP server)
- **Node.js 18+** and **npm** (for building the add-in)

### Build

```bash
# Build everything (production)
./scripts/build.sh

# Build MCP server only
./scripts/build.sh mcp

# Build PowerPoint add-in only
./scripts/build.sh addin

# Development mode (add-in with hot reload)
./scripts/build.sh dev
```

### Run

```bash
# 1. Start the MCP server
dotnet run --project src/mcp-server/

# Or use the published executable:
# ./src/mcp-server/publish/win-x64/office-mcp-server.exe

# 2. Sideload the PowerPoint add-in
# See "Add-in Sideloading" section below

# 3. Open PowerPoint and load the add-in
# File → Options → Trust Center → Trust Center Settings → Trusted Add-ins
```

### Add-in Sideloading

For development, sideload the add-in by pointing PowerPoint to the manifest:

1. Open PowerPoint
2. Go to **File → Options → Trust Center → Trust Center Settings**
3. Select **Trusted Add-in Publishers**
4. Click **Add** and browse to `src/powerpoint-addin/manifest.xml`

Or use the [Office Add-in CLI](https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-add-in-for-word) for network sideloading:

```bash
npx office-addin-debugging start src/powerpoint-addin/manifest.xml ppt
```

## Development

### Office JS Add-in

```bash
cd src/powerpoint-addin
npm install
npm run dev    # Starts webpack dev server on port 3000
npm run build  # Production build
```

### MCP Server

```bash
cd src/mcp-server
dotnet restore
dotnet run     # Development
dotnet publish -c Release -r win-x64 --self-contained true  # Production executable
```

## Development Notes

These are hard-won lessons from building this project. Follow them to avoid known pitfalls.

### Office JS API

- **`PowerPoint.run()` is lowercase** — not `Run()`, `Excel.Run()`, etc. The Office JS API uses camelCase (`run`, not `Run`). A typo here silently fails with "PowerPoint.run() not available" because `PowerPoint.Run` is `undefined`.
- **`context.load(collection, ["items"])`** — always load `items` on collections. Do NOT use `"notCoveredByParallelization"` or other VSTO-era properties; those don't exist in Office JS.
- **`context.load(obj, ["id", "name"])`** — pass property names as an array of strings, not a comma-separated string.
- **`@types/office-js` is incomplete** — many newer PowerPoint context types lack type definitions. Use `any` casts and `PowerPoint.run(async (context: any) => ...)`.
- **`Office.context.document.url`** — gives the document URL/path (if available). Use this for the real document name in `getOfficeState()`.

### MCP Protocol

- **`params.arguments`** — tool call arguments come under `params.arguments`, NOT `params.input`. The MCP spec uses `arguments`.
- **`inputSchema`** — tool definitions use `inputSchema` (not `parameters` or `schema`). The `properties` and `required` fields follow JSON Schema.
- **`@default` in C#** — JSON Schema uses `default` as the keyword. In C#, use `@default` (with `@` prefix since `default` is a reserved word).

### Add-in Architecture

- **Command polling must process results** — `startCommandPolling()` must call `processPendingCommands()`, not just `pollForCommands()`. The poll returns commands that must be dispatched to `processCommand()`.
- **`processCommand()` calls `reportResult()` internally** — don't double-report. The handler already POSTs the result back to the MCP server.
- **`Office.onReady()` can fire multiple times** — use an `isInitialized` guard to prevent double-registration.
- **No duplicate `<script>` tags** — `HtmlWebpackPlugin` already injects `bundle.js`. Adding a static `<script src="bundle.js">` causes double initialization.

### ASP.NET / MCP Server

- **Suppress noisy request logging** — polling endpoints fire every 2s. Set `Microsoft.AspNetCore` logging to `Warning` level:
  ```csharp
  builder.Logging.AddFilter("Microsoft.AspNetCore", LogLevel.Warning);
  ```
- **`CleanupTimedOut()` must remove instances** — don't just set `IsAlive = false` or the same dead instance logs "timed out" every 30 seconds. Actually remove it from the dictionary.
- **`Results.Json()` serializes anonymous types correctly** — `Dictionary<string, object>` with anonymous type values works fine with `System.Text.Json`.

### Deployment

- **Manifest uses `{{BASE_URL}}` placeholders** — replaced at request time by Express server. No hardcoded URLs.
- **`<AppDomain>` not `<Domain>`** — Office manifest XML uses `<AppDomain>`, not `<Domain>`.
- **`<Host Name="Presentation"/>`** — the `Name` attribute is capitalized.
- **No `<RequestedWidth>`** on TaskPaneApp — only valid for Content app types.
- **Traefik ingress** — use `traefik` ingress class, not `nginx`.
- **`{{ }}` in Helm** — no spaces inside braces. Some editors auto-format and insert `{ { } }` which breaks templates.

## Testing

### MCP Server (C#)

58 xUnit tests covering models, command routing, and HTTP endpoints:

```bash
dotnet test tests/mcp-server.Tests/
```

- Uses `WebApplicationFactory<Program>` for in-process HTTP integration tests.
- `McpToolEngine.ResetForTesting()` clears static state before each test class.
- Pre-commit hook (`.git/hooks/pre-commit`) enforces all tests pass.

### Express Server (Node.js)

6 tests for dynamic manifest generation:

```bash
cd server && npm test
```

### Add-in (TypeScript)

The add-in must be tested in a real Windows Office environment:

1. Build both components
2. Start the MCP server
3. Sideload the PowerPoint add-in
4. Verify task pane loads in PowerPoint
5. Test each tool via Open WebUI's MCP integration

## CI/CD

| Workflow      | Trigger             | Jobs                                                                                               |
| ------------- | ------------------- | -------------------------------------------------------------------------------------------------- |
| `ci.yml`      | Push to master/main | Build+Test → Docker push (`ghcr.io/volkermauel/officellm-static:latest`) → Windows `.exe` artifact |
| `release.yml` | Tag `v*`            | GitHub Release with `.exe` + Docker semver tag                                                     |

Download latest exe:

```bash
gh run download --name office-mcp-server-win-x64
```

## Specifications

See [`specs/`](specs/) for detailed feature specifications organized by implementation phase:

- [Phase 0: Spike](specs/001-spike/) - Minimal MCP server + PowerPoint add-in
- [Phase 1: PowerPoint MVP](specs/002-powerpoint-mvp/) - Full PowerPoint tool set
- [Phase 2: Word MVP](specs/003-word-mvp/) - Word tools + shared context
- [Phase 3: Excel MVP](specs/004-excel-mvp/) - Excel read/write tools
- [Phase 4: Outlook MVP](specs/005-outlook-mvp/) - Email tools + policy filter

## License

Private / Internal Use Only
