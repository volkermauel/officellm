# Office LLM Harness — Deployment Guide

## Overview

The Office LLM Harness consists of two components:

1. **MCP Server** — Runs **locally on each user's PC** (Windows executable). Handles Office JS add-in communication, command routing, and tool execution.
2. **Static File Server** — Centrally hosted via Helm chart. Serves PowerPoint/Word/Excel/Outlook add-in task pane UI files (JS, HTML, CSS).

### Architecture

```
┌─────────────────┐     HTTPS      ┌──────────────────┐
│  Internal Ingress│◄─────────────►│  Kubernetes /     │
│  (Nginx/Apache) │               │  Helm Chart       │
│                 │               │  (nginx)          │
└─────────────────┘               └───────────────────┘
                                    ▲
                                    │ serves static files
                                    │
                              ┌─────┴──────────────────┐
                              │  ghcr.io/...-static    │
                              │  (PowerPoint add-in UI)│
                              └────────────────────────┘

┌─────────────────┐     HTTP      ┌──────────────────┐
│  User's PC      │◄─────────────►│  Office Add-ins  │
│  (Windows)      │  Port 3000   │  (PowerPoint,    │
│                 │              │   Word, Excel)   │
│  MCP Server     │              │  Task pane loads │
│  (office-mcp-   │              │  from Helm host  │
│   server.exe)   │              │                  │
└─────────────────┘              └──────────────────┘
```

## 1. Deploy Static File Server (Helm)

### Prerequisites

- Kubernetes cluster with nginx ingress controller
- Access to `ghcr.io/YOUR_ORG/office-llm-harness-static`

### Install via Helm

```bash
# Add your GHCR image repo (replace YOUR_ORG)
helm repo add office-llm-harness https://YOUR_ORG.github.io/office-llm-harness/

# Create values.override.yaml
cat > values.override.yaml <<EOF
image:
  repository: ghcr.io/YOUR_ORG/office-llm-harness-static
  tag: v1.0.0

ingress:
  hosts:
    - host: addins.yourcompany.internal
      paths:
        - path: /
          pathType: Prefix
  tls:
    - secretName: office-llm-harness-tls
      hosts:
        - addins.yourcompany.internal
EOF

# Install
helm install office-llm-harness ./helm/office-llm-harness -f values.override.yaml

# Verify
kubectl get pods -l app.kubernetes.io/name=office-llm-harness
kubectl get ingress office-llm-harness
```

### TLS Secret

```bash
kubectl create secret tls office-llm-harness-tls \
  --cert=tls.crt \
  --key=tls.key \
  -n default
```

## 2. Deploy MCP Server (Per-User)

The MCP server runs on each user's PC. Download the latest release:

### Windows

```powershell
# Download latest release
Invoke-WebRequest -Uri "https://github.com/YOUR_ORG/office-llm-harness/releases/download/v1.0.0/office-mcp-server-win-x64.zip" -OutFile "office-mcp-server.zip"
Expand-Archive -Path "office-mcp-server.zip" -DestinationPath ".\office-mcp-server\"

# Run (default port 3000)
.\office-mcp-server\office-mcp-server.exe

# Custom port
.\office-mcp-server\office-mcp-server.exe 8080
```

### Linux (development/testing)

```bash
curl -L -o office-mcp-server.zip https://github.com/YOUR_ORG/office-llm-harness/releases/download/v1.0.0/office-mcp-server-linux-x64.zip
unzip office-mcp-server.zip
chmod +x office-mcp-server
./office-mcp-server 3000
```

## 3. Configure Office Add-in

Update the manifest (`manifest.xml`) to point to your Helm-hosted static files:

```xml
<DefaultSettings>
  <SourceLocation DefaultValue="https://addins.yourcompany.internal/index.html"/>
</DefaultSettings>
```

Then load the add-in in PowerPoint:

1. **File** → **Options** → **Trust Center** → **Trust Center Settings**
2. **Trusted Add-in Paths** → **Add** → path to `manifest.xml`
3. **Restart PowerPoint**

## Environment Variables (Static File Server)

No environment variables needed — nginx serves static files directly.

## Health Checks

| Endpoint     | Method | Description                     |
| ------------ | ------ | ------------------------------- |
| `/health`    | GET    | Returns `{"status":"ok"}`       |
| `/`          | GET    | Serves `index.html` (task pane) |
| `/bundle.js` | GET    | Serves minified JS bundle       |

### Kubernetes Health Check

The deployment includes liveness, readiness, and startup probes configured in `helm/office-llm-harness/templates/deployment.yaml`.

## Scaling

The static file server is **stateless** — no persistent state between requests. Horizontal scaling is straightforward:

```yaml
# helm/office-llm-harness/values.yaml
replicaCount: 3
```

## GitHub Container Registry (GHCR)

Images are automatically pushed to GHCR on every push to `master`/`main` and on every tag release.

```bash
# Pull latest
docker pull ghcr.io/YOUR_ORG/office-llm-harness-static:latest

# Pull specific version
docker pull ghcr.io/YOUR_ORG/office-llm-harness-static:v1.0.0
```

### Authentication

```bash
echo ${{ secrets.GITHUB_TOKEN }} | docker login ghcr.io -u ${{ github.actor }} --password-stdin
```

## Releases

Releases are created automatically when you push a git tag:

```bash
git tag v1.0.0
git push origin v1.0.0
```

This triggers:

1. Build of Windows (`win-x64`) and Linux (`linux-x64`) MCP server executables
2. Creation of a GitHub Release with `.zip` assets
3. Docker image build and push to GHCR (static file server)

### Downloading Releases

```powershell
# Windows
Invoke-WebRequest -Uri "https://github.com/YOUR_ORG/office-llm-harness/releases/download/v1.0.0/office-mcp-server-win-x64.zip" -OutFile "office-mcp-server.zip"
```

```bash
# Linux
curl -L -o office-mcp-server.zip https://github.com/YOUR_ORG/office-llm-harness/releases/download/v1.0.0/office-mcp-server-linux-x64.zip
```
