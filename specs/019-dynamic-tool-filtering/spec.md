# Feature Specification: Phase 19 — Dynamic Tool Filtering by Active Host

**Feature Branch**: `019-dynamic-tool-filtering`

**Created**: 2026-05-16

**Status**: Draft

**Depends on**: All MVP phases, InstanceRegistry with `GetHostPrefix()` (Phase 8)

**Motivation**: With 107 tools, the `tools/list` response is ~40KB of JSON. An LLM client that only has PowerPoint connected still receives 30 Word tools, 19 Excel tools, and 12 Outlook tools — wasting context window tokens. This phase filters `tools/list` to only expose tools for hosts that have active registered instances.

## Architecture

```
LLM Client                     MCP Server (port 3000)
    │                                │
    │── POST /mcp ──────────────────►│
    │   { method: "tools/list" }     │
    │                                │── McpToolEngine.GetToolDefinitions(activeHosts)
    │                                │   activeHosts = { "powerpoint", "word" }
    │                                │   → filter: keep office_* + powerpoint_* + word_*
    │◄── { tools: [...62 tools] } ───│   (instead of 107)
```

**Key insight**: `GetToolDefinitions()` currently returns a static array. It needs to accept a set of active host prefixes and filter accordingly. Shared tools (`office_*`) are always included.

## User Scenarios & Testing

### User Story 1 — Only PowerPoint connected (Priority: P0)

A user has only PowerPoint running. The LLM calls `tools/list` and receives only PowerPoint tools + shared tools (≈35 tools instead of 107). This saves ~70 tools × ~400 bytes ≈ 28KB of context per request.

**Acceptance Scenarios**:

1. **Given** only a PowerPoint instance is registered, **When** the LLM calls `tools/list`, **Then** the response includes all `office_*` and `powerpoint_*` tools, and excludes all `word_*`, `excel_*`, and `outlook_*` tools
2. **Given** no instances are registered, **When** the LLM calls `tools/list`, **Then** the response includes only `office_get_active_apps` (the discovery tool) — no host-specific tools

---

### User Story 2 — Multiple hosts connected (Priority: P0)

A user has PowerPoint and Excel open. The LLM calls `tools/list` and receives PowerPoint + Excel + shared tools (≈55 tools).

**Acceptance Scenarios**:

1. **Given** instances for PowerPoint and Excel are registered, **When** the LLM calls `tools/list`, **Then** the response includes `office_*`, `powerpoint_*`, and `excel_*` tools, but not `word_*` or `outlook_*`
2. **Given** all 4 hosts are registered, **When** the LLM calls `tools/list`, **Then** all 107 tools are returned (same as current behavior)

---

### User Story 3 — Host appears mid-session (Priority: P1)

A user opens Word while the LLM is already working. On the next `tools/list` call, Word tools appear.

**Acceptance Scenarios**:

1. **Given** only PowerPoint is registered, **When** the LLM calls `tools/list`, **Then** Word tools are absent; **When** a Word instance registers, **And** the LLM calls `tools/list` again, **Then** Word tools are now included

---

### User Story 4 — REST API bridge also filters (Priority: P1)

The OpenAPI spec and `/api/{toolName}` endpoints also reflect only active hosts.

**Acceptance Scenarios**:

1. **Given** only PowerPoint is registered, **When** the LLM fetches `/openapi.json`, **Then** only PowerPoint and shared tool endpoints are listed

---

### Edge Cases

- **Timed-out instances**: Should we count them? **No** — only active (non-timed-out) instances count. Use `GetActiveInstances()`.
- **`tools/call` for a filtered-out tool**: If the LLM caches an old tool list and calls a tool that's now filtered out, it should still work if the instance exists. The filter is advisory (reduces context), not a permission gate. **Do not reject calls to filtered-out tools.**
- **Static `GetToolDefinitions()` callers**: Some code paths call `GetToolDefinitions()` without host context (e.g., `HandleSuggestTools`, `HandleBatchCall`). These should continue to see all tools so they can suggest/call any registered host's tools.

## Requirements

### Functional Requirements

- **FR-001**: `McpToolEngine.GetToolDefinitions()` MUST accept an optional `HashSet<string> activeHosts` parameter. When null or empty, it returns all tools (backward-compatible).
- **FR-002**: When `activeHosts` is provided, the method MUST return only tools whose name matches one of: `office_*` (always), or `{host}_*` where `{host}` is in `activeHosts`.
- **FR-003**: The `tools/list` MCP endpoint in `AppBuilder.cs` MUST query `InstanceRegistry.GetActiveInstances()` to determine active hosts and pass them to `GetToolDefinitions()`.
- **FR-004**: The host detection MUST use the existing `GetHostType()` logic to map instance app names to host prefixes.
- **FR-005**: The `tools/call` handler MUST NOT reject calls based on filtering. Filtering is advisory only.
- **FR-006**: The REST API bridge (`/openapi.json`, `/api/{toolName}`) SHOULD also filter endpoints by active hosts (stretch goal — can be deferred).
- **FR-007**: The `office_get_active_apps` tool MUST always be included regardless of filtering.

### Test Adjustments

Existing tests call `GetToolDefinitions()` with no arguments (the static all-tools path). The following changes are needed:

| Test File | Test Method | Change |
|-----------|-------------|--------|
| `McpToolEngineTests.cs` | `GetToolDefinitions_Returns107Tools` | Keep as-is (no-arg = all tools). Rename to `GetToolDefinitions_NoFilter_ReturnsAllTools` |
| `McpToolEngineTests.cs` | `GetToolDefinitions_ContainsAllToolNames` | Keep as-is (no-arg = all tools) |
| `McpToolEngineTests.cs` | `GetToolDefinitions_AllToolsHaveRequiredFields` | Keep as-is |
| `McpToolEngineTests.cs` | `GetToolDefinitions_PowerPointToolsRequireInstanceId` | Keep as-is |
| `McpToolEngineTests.cs` | *(new)* | `GetToolDefinitions_WithPowerPointHost_ReturnsOnlyOfficeAndPowerPointTools` |
| `McpToolEngineTests.cs` | *(new)* | `GetToolDefinitions_WithNoHosts_ReturnsOnlyOfficeTools` |
| `McpToolEngineTests.cs` | *(new)* | `GetToolDefinitions_WithMultipleHosts_ReturnsCorrectSubset` |
| `HttpEndpointTests.cs` | `Mcp_ToolsList_Returns107Tools` | Adjust: register no instances → should get only `office_get_active_apps`. Rename to `Mcp_ToolsList_NoInstances_ReturnsMinimalTools` |
| `HttpEndpointTests.cs` | *(new)* | `Mcp_ToolsList_WithPowerPointInstance_ReturnsPowerPointAndOfficeTools` |
| `HttpEndpointTests.cs` | *(new)* | `Mcp_ToolsList_WithAllHosts_ReturnsAllTools` |
| `HttpEndpointTests.cs` | `Mcp_ToolsList_AllNewToolsHaveDescriptions` | Register all hosts first, then check |
| `HttpEndpointTests.cs` | `OpenApi_ContainsAllToolEndpoints` | Register all hosts first |

### Key Design Decisions

1. **Filter in `GetToolDefinitions()`, not in a wrapper**: The method signature changes from `object[] GetToolDefinitions()` to `object[] GetToolDefinitions(HashSet<string>? activeHosts = null)`. The default `null` means "all tools" — backward-compatible.

2. **Prefix-based matching**: A tool like `word_get_bookmarks` has prefix `word`. The filter checks if the tool name starts with `office_` OR starts with `{host}_` for any host in `activeHosts`.

3. **Host prefix derivation**: Use the existing `GetHostType()` which returns `"powerpoint"`, `"word"`, `"excel"`, `"outlook"`, or `"unknown"`. Instances with `"unknown"` don't contribute any host-specific tools.

4. **No filter on `tools/call`**: The LLM might cache a tool list from when Word was connected, then Word disconnects, and the LLM tries a Word tool. The call will fail with `INSTANCE_NOT_FOUND` — which is the correct behavior. No additional gating needed.

## Implementation Plan

### Step 1: Modify `McpToolEngine.GetToolDefinitions()` signature

```csharp
// Before:
public static object[] GetToolDefinitions() => [...];

// After:
public static object[] GetToolDefinitions(HashSet<string>? activeHosts = null)
{
    var allTools = new object[] { ... }; // same array literal
    if (activeHosts == null || activeHosts.Count == 0)
        return allTools;

    return allTools.Where(t =>
    {
        var json = JsonSerializer.Serialize(t);
        var doc = JsonDocument.Parse(json);
        var name = doc.RootElement.GetProperty("name").GetString()!;
        if (name.StartsWith("office_")) return true;
        foreach (var host in activeHosts)
        {
            if (name.StartsWith(host + "_")) return true;
        }
        return false;
    }).ToArray();
}
```

**Optimization**: Pre-compute the tool name → host prefix mapping once at startup instead of serializing every call. Add a private helper:

```csharp
private static readonly (object Tool, string Name)[] _toolIndex = [...];
private static string GetToolHost(string toolName) { ... }
```

### Step 2: Update `AppBuilder.cs` `tools/list` handler

```csharp
case "tools/list":
    var activeHosts = GetActiveHostPrefixes();  // from registry
    var filteredTools = McpToolEngine.GetToolDefinitions(activeHosts);
    return Results.Json(new { jsonrpc = "2.0", id = ..., result = new { tools = filteredTools } });
```

Add a helper to extract unique host prefixes from active instances:

```csharp
private HashSet<string> GetActiveHostPrefixes()
{
    var instances = McpToolEngine.GetRegistry().GetActiveInstances();
    return instances
        .Select(i => GetHostType(i.AppName))
        .Where(h => h != "unknown")
        .ToHashSet();
}
```

### Step 3: Update tests

- All existing tests that call `GetToolDefinitions()` with no args continue to work (backward-compatible)
- Add new tests for filtered scenarios
- HttpEndpointTests that check tool counts need to register instances first

### Step 4: Update `Program.cs` stdio transport (if applicable)

The stdio transport also handles `tools/list` — apply the same filtering.

## Success Criteria

### Measurable Outcomes

- **SC-001**: With 0 instances registered, `tools/list` returns ≤ 8 tools (shared tools only: `office_get_active_apps`, `office_get_document_context`, `office_get_document_stats`, `office_batch_call`, `office_suggest_tools`, `office_export_document`).
- **SC-002**: With 1 host registered, `tools/list` returns only shared + that host's tools (e.g., PowerPoint: ~35 tools).
- **SC-003**: With all 4 hosts registered, `tools/list` returns all 107 tools.
- **SC-004**: Context savings: ~30KB less JSON per `tools/list` call when only 1 host is active.
- **SC-005**: All 91 existing tests continue to pass without modification (backward-compatible no-arg call).
- **SC-006**: New tests cover: 0 hosts, 1 host, 2 hosts, all hosts.
