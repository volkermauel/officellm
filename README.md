# Office LLM Harness

A Windows desktop Office add-in harness that exposes controlled document interaction tools to Open WebUI through a local MCP server, enabling LLM-assisted workflows in Word, Excel, PowerPoint and Outlook.

## Architecture

```
+--------------------+        +-------------------------+
| Word / Excel /     |        | Open WebUI              |
| PowerPoint /       |        | - Chat UI               |
| Outlook            |        | - Model routing         |
+---------+----------+        | - MCP external tools    |
          | VSTO / COM                    |
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
| - Tool registry                                      |
| - Bridge server (port 8765)                           |
| - Audit log                                          |
+------------------------------------------------------+
```

## Components

| Component             | Language        | Description                                                                                                           |
| --------------------- | --------------- | --------------------------------------------------------------------------------------------------------------------- |
| **MCP Server**        | C# (.NET 8)     | Self-contained executable exposing MCP tools over Streamable HTTP. Includes a bridge server for add-in communication. |
| **PowerPoint Add-in** | TypeScript/HTML | Office JS Add-in running as a task pane in PowerPoint. Provides the Office API interaction layer.                     |
| **Bridge Server**     | .NET (embedded) | Lightweight HTTP server on port 8765 that queues commands from the MCP server for the add-in to process.              |

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
├── powerpoint-addin/     # Office JS PowerPoint Add-in
│   ├── manifest.xml      # Add-in manifest
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

## Testing

The add-in must be tested in a real Windows Office environment:

1. Build both components
2. Start the MCP server
3. Sideload the PowerPoint add-in
4. Verify task pane loads in PowerPoint
5. Test each tool via Open WebUI's MCP integration

## Specifications

See [`specs/`](specs/) for detailed feature specifications organized by implementation phase:

- [Phase 0: Spike](specs/001-spike/) - Minimal MCP server + PowerPoint add-in
- [Phase 1: PowerPoint MVP](specs/002-powerpoint-mvp/) - Full PowerPoint tool set
- [Phase 2: Word MVP](specs/003-word-mvp/) - Word tools + shared context
- [Phase 3: Excel MVP](specs/004-excel-mvp/) - Excel read/write tools
- [Phase 4: Outlook MVP](specs/005-outlook-mvp/) - Email tools + policy filter

## License

Private / Internal Use Only
