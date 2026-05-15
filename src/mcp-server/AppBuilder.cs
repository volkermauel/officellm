using System.Text.Json;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.DependencyInjection;
using OfficeMcpServer.Models;
using OfficeMcpServer.Tools;

namespace OfficeMcpServer;

/// <summary>
/// Creates and configures the MCP server WebApplication.
/// Separated from Program.cs to enable WebApplicationFactory-based integration testing.
/// </summary>
public static class AppBuilder
{
    /// <summary>
    /// Creates a pre-configured WebApplication. The caller must call app.Run() to start it.
    /// For testing, pass mcpHost=null and mcpPort=null to use ASP.NET's default URL mechanism.
    /// </summary>
    public static WebApplication Create(string? mcpHost = "127.0.0.1", int? mcpPort = null)
    {
        // args are now parsed in Program.cs before calling Create()

        var builder = WebApplication.CreateBuilder();

        if (mcpPort.HasValue && !string.IsNullOrEmpty(mcpHost))
        {
            builder.WebHost.UseUrls($"http://{mcpHost}:{mcpPort.Value}");
        }

        // Read CORS origins from env var (comma-separated) or use defaults
        string? corsOriginsEnv = Environment.GetEnvironmentVariable("CORS_ORIGINS");
        string[] allowedOrigins;

        if (!string.IsNullOrWhiteSpace(corsOriginsEnv))
        {
            allowedOrigins = corsOriginsEnv.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        }
        else
        {
            allowedOrigins = ["http://127.0.0.1:*", "https://127.0.0.1:*"];
        }

        builder.Services.AddCors(options =>
        {
            options.AddPolicy("AllowedOrigins", policy =>
            {
                policy.WithOrigins(allowedOrigins)
                      .AllowAnyHeader()
                      .AllowAnyMethod();
            });
        });

        var app = builder.Build();
        app.UseCors("AllowedOrigins");

        // --- Static file serving (PowerPoint/Word/Excel/Outlook add-in UI) ---
        string? staticFilesPath = Environment.GetEnvironmentVariable("STATIC_FILES_PATH")
            ?? Environment.GetEnvironmentVariable("STATIC_FILES_DIR");

        if (!string.IsNullOrEmpty(staticFilesPath))
        {
            app.UseStaticFiles(new StaticFileOptions
            {
                FileProvider = new PhysicalFileProvider(staticFilesPath),
                RequestPath = ""
            });
        }
        else
        {
            // Default: look for wwwroot in the current directory
            string wwwroot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot");
            if (Directory.Exists(wwwroot))
            {
                app.UseStaticFiles(new StaticFileOptions
                {
                    FileProvider = new PhysicalFileProvider(wwwroot),
                    RequestPath = ""
                });
            }
        }

        // --- Resolve shared state ---
        var registry = McpToolEngine.GetRegistry();
        var commandStore = McpToolEngine.GetCommandStore();

        // --- Start cleanup timer (runs every 30s) ---
        _ = Task.Run(async () =>
        {
            while (true)
            {
                await Task.Delay(30000);
                registry.CleanupTimedOut();
            }
        });

        // ============================================================
        // INSTANCE MANAGEMENT ENDPOINTS
        // ============================================================

        app.MapPost("/instances/register", async (HttpContext context) =>
        {
            using var reader = new StreamReader(context.Request.Body);
            var json = await reader.ReadToEndAsync();
            var data = System.Text.Json.JsonSerializer.Deserialize<System.Text.Json.JsonElement>(json);
            string appName = data.GetProperty("appName").GetString() ?? "Unknown";
            string documentName = data.TryGetProperty("documentName", out var dn) ? dn.GetString() ?? "" : "";

            string instanceId = registry.RegisterInstance(appName, documentName);
            Console.WriteLine($"Registered instance: {instanceId} ({appName} - {documentName})");
            return Results.Json(new { instanceId, appName, documentName });
        });

        app.MapPost("/instances/{instanceId}/heartbeat", async (HttpContext context, string instanceId) =>
        {
            using var reader = new StreamReader(context.Request.Body);
            var json = await reader.ReadToEndAsync();
            var data = System.Text.Json.JsonSerializer.Deserialize<System.Text.Json.JsonElement>(json);
            string? appName = data.TryGetProperty("appName", out var a) ? a.GetString() : null;
            string? documentName = data.TryGetProperty("documentName", out var d) ? d.GetString() : null;
            registry.UpdateHeartbeat(instanceId, appName, documentName);
            return Results.Json(new { status = "ok" });
        });

        app.MapGet("/instances", () =>
        {
            var instances = registry.GetActiveInstances().Select(i => new
            {
                i.InstanceId, i.AppName, i.DocumentName, i.IsAlive,
                RegisteredAt = i.RegisteredAt.ToString("o"),
            });
            return Results.Json(new { instances });
        });

        app.MapGet("/instances/{instanceId}/commands", (string instanceId) =>
        {
            var commands = commandStore.GetPendingCommands(instanceId);
            if (commands.Any())
                foreach (var cmd in commands)
                    commandStore.MarkClaimed(cmd.Id, instanceId);
            return Results.Json(new { commands });
        });

        app.MapPost("/instances/{instanceId}/confirm", async (HttpContext context, string instanceId) =>
        {
            using var reader = new StreamReader(context.Request.Body);
            var json = await reader.ReadToEndAsync();
            var data = System.Text.Json.JsonSerializer.Deserialize<System.Text.Json.JsonElement>(json);
            string commandId = data.GetProperty("commandId").GetString() ?? "";
            string confirmationToken = data.GetProperty("confirmationToken").GetString() ?? "";
            bool approved = data.TryGetProperty("approved", out var a) && a.GetBoolean();

            var confirmationStore = McpToolEngine.GetConfirmationStore();
            var commandStore = McpToolEngine.GetCommandStore();

            if (!approved)
            {
                // Reject: log and return error result
                commandStore.CompleteCommand(commandId, false, "", "User rejected the change");
                return Results.Json(new { status = "rejected" });
            }

            // Approve: re-execute the tool with the confirmation token
            // The tool handler will validate the token and apply the change
            if (!confirmationStore.ValidateToken(confirmationToken))
            {
                return Results.BadRequest(new { error = "Invalid or expired confirmation token" });
            }

            // Get the pending command to re-execute with token
            var pendingCmd = commandStore.GetPendingCommands(instanceId).FirstOrDefault(c => c.Id == commandId);
            if (pendingCmd == null)
            {
                return Results.NotFound(new { error = "Command not found" });
            }

            // Re-execute the tool with confirmation token
            var toolName = pendingCmd.Command;
            var argsWithToken = new Dictionary<string, object>
            {
                ["instanceId"] = instanceId,
                ["confirmationToken"] = confirmationToken
            };

            if (pendingCmd.Args is System.Text.Json.JsonElement argsEl)
            {
                foreach (var prop in argsEl.EnumerateObject())
                {
                    if (prop.Name != "confirmationToken")
                        argsWithToken[prop.Name] = prop.Value;
                }
            }

            var result = await McpToolEngine.ExecuteTool(toolName, System.Text.Json.JsonSerializer.SerializeToElement(argsWithToken));
            commandStore.CompleteCommand(commandId, true, "", result);

            return Results.Json(new { status = "approved", result });
        });

        app.MapPost("/instances/{instanceId}/result", async (HttpContext context, string instanceId) =>
        {
            using var reader = new StreamReader(context.Request.Body);
            var json = await reader.ReadToEndAsync();
            var data = System.Text.Json.JsonSerializer.Deserialize<System.Text.Json.JsonElement>(json);
            string commandId = data.GetProperty("commandId").GetString() ?? "";
            bool success = data.TryGetProperty("success", out var s) && s.GetBoolean();
            string? error = data.TryGetProperty("error", out var e) ? e.GetString() : null;
            object? payload = data.TryGetProperty("payload", out var p) ? JsonSerializer.Deserialize<object>(p) : null;
            Console.WriteLine($"Instance {instanceId} completed command {commandId}: success={success}");
            commandStore.CompleteCommand(commandId, success, error ?? "", payload);
            return Results.Json(new { status = "ok" });
        });

        // ============================================================
        // MCP PROTOCOL ENDPOINTS (Streamable HTTP)
        // ============================================================

        app.MapPost("/mcp", async (HttpContext context) =>
        {
            using var reader = new StreamReader(context.Request.Body);
            var json = await reader.ReadToEndAsync();
            if (string.IsNullOrWhiteSpace(json))
                return Results.BadRequest("Empty request body");

            try
            {
                var message = System.Text.Json.JsonSerializer.Deserialize<System.Text.Json.JsonElement>(json);
                string method = message.GetProperty("method").GetString() ?? "";

                switch (method)
                {
                    case "initialize":
                        return Results.Json(new
                        {
                            jsonrpc = "2.0", id = message.GetProperty("id"),
                            result = new
                            {
                                protocolVersion = "2024-11-05",
                                serverInfo = new { name = "OfficeMcpServer", version = "0.1.0-spike" },
                                capabilities = new { tools = new { listChanged = true } }
                            }
                        }, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });

                    case "notifications/initialized":
                        return Results.Json(new { jsonrpc = "2.0", result = (object?)null });

                    case "tools/list":
                        var tools = McpToolEngine.GetToolDefinitions();
                        return Results.Json(new
                        { jsonrpc = "2.0", id = message.GetProperty("id"), result = new { tools } },
                        new System.Text.Json.JsonSerializerOptions { WriteIndented = true });

                    case "tools/call":
                        var toolName = message.GetProperty("params").GetProperty("name").GetString() ?? "";
                        System.Text.Json.JsonElement? toolArgs = null;
                        if (message.GetProperty("params").TryGetProperty("input", out var input))
                            toolArgs = input;
                        var result = await McpToolEngine.ExecuteTool(toolName, toolArgs);
                        return Results.Json(new
                        { jsonrpc = "2.0", id = message.GetProperty("id"), result },
                        new System.Text.Json.JsonSerializerOptions { WriteIndented = true });

                    default:
                        return Results.Problem(detail: $"Method not found: {method}", statusCode: StatusCodes.Status404NotFound);
                }
            }
            catch (Exception ex)
            {
                return Results.Problem(detail: $"Parse error: {ex.Message}", statusCode: StatusCodes.Status400BadRequest);
            }
        });

        // ============================================================
        // HEALTH ENDPOINT
        // ============================================================

        app.MapGet("/health", () =>
        {
            var activeCount = registry.GetActiveInstances().Count;
            return Results.Json(new { status = "ok", activeInstances = activeCount });
        });

        // ============================================================
        // OPENAPI REST BRIDGE
        // ============================================================

        app.MapGet("/openapi.json", (HttpRequest request) =>
        {
            var tools = McpToolEngine.GetToolDefinitions();
            var baseUrl = $"{request.Scheme}://{request.Host}";

            var paths = new Dictionary<string, object>();

            foreach (var toolObj in tools)
            {
                var tool = (dynamic)toolObj;
                string toolName = tool.name;
                string description = tool.description;
                var inputSchema = tool.inputSchema;

                var propDict = new Dictionary<string, object>();
                var requiredList = new List<string>();

                if (inputSchema is Dictionary<string, object> inputProps
                    && inputProps.TryGetValue("properties", out var propsObj)
                    && propsObj is Dictionary<string, object> props)
                {
                    foreach (var kvp in props)
                    {
                        var propDef = (dynamic)kvp.Value;
                        var openApiProp = new Dictionary<string, object>
                        {
                            ["type"] = propDef.type ?? "string",
                            ["description"] = propDef.description ?? ""
                        };
                        propDict[kvp.Key] = openApiProp;
                    }
                }

                if (inputSchema is Dictionary<string, object> inputReq
                    && inputReq.TryGetValue("required", out var reqObj)
                    && reqObj is string[] reqArr)
                {
                    requiredList = reqArr.ToList();
                }

                var requestSchema = new
                {
                    type = "object",
                    properties = propDict,
                    required = requiredList
                };

                paths[$"/api/{toolName}"] = new
                {
                    post = new
                    {
                        summary = description,
                        operationId = toolName,
                        tags = new[] { "Office Tools" },
                        requestBody = new
                        {
                            required = true,
                            content = new Dictionary<string, object>
                            {
                                ["application/json"] = new { schema = requestSchema }
                            }
                        },
                        responses = new Dictionary<string, object>
                        {
                            ["200"] = new
                            {
                                description = "Tool execution result",
                                content = new Dictionary<string, object>
                                {
                                    ["application/json"] = new
                                    {
                                        schema = new
                                        {
                                            type = "object",
                                            properties = new Dictionary<string, object>
                                            {
                                                ["content"] = new { type = "array", description = "Array of content blocks" },
                                                ["isError"] = new { type = "boolean", description = "Whether the tool call resulted in an error" }
                                            }
                                        }
                                    }
                                }
                            },
                            ["400"] = new { description = "Invalid request" },
                            ["404"] = new { description = "Tool not found" }
                        }
                    }
                };
            }

            var spec = new
            {
                openapi = "3.0.0",
                info = new
                {
                    title = "Office LLM Harness API",
                    version = "0.1.0-spike",
                    description = "REST API bridge for Office MCP tools. Each MCP tool is exposed as a POST endpoint."
                },
                servers = new[] { new { url = baseUrl, description = "Local MCP server" } },
                paths,
                components = new
                {
                    schemas = new Dictionary<string, object>
                    {
                        ["ToolResult"] = new
                        {
                            type = "object",
                            properties = new Dictionary<string, object>
                            {
                                ["content"] = new { type = "array", items = new { type = "object" } },
                                ["isError"] = new { type = "boolean" }
                            }
                        }
                    }
                }
            };

            return Results.Json(spec, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
        });

        app.MapPost("/api/{toolName}", async (HttpContext context, string toolName) =>
        {
            using var reader = new StreamReader(context.Request.Body);
            var json = await reader.ReadToEndAsync();

            System.Text.Json.JsonElement? args = null;
            if (!string.IsNullOrWhiteSpace(json))
            {
                try { args = System.Text.Json.JsonSerializer.Deserialize<System.Text.Json.JsonElement>(json); }
                catch { return Results.BadRequest(new { error = "Invalid JSON body" }); }
            }

            var result = await McpToolEngine.ExecuteTool(toolName, args);
            return Results.Json(result, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
        });

        app.MapGet("/docs", () => Results.Text(
            "<!DOCTYPE html><html><head><title>Office LLM Harness API</title>" +
            "<meta charset=\"utf-8\"/><meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">" +
            "<link rel=\"stylesheet\" type=\"text/css\" href=\"https://unpkg.com/swagger-ui-dist@5/swagger-ui.css\">" +
            "</head><body><div id=\"swagger-ui\"></div>" +
            "<script src=\"https://unpkg.com/swagger-ui-dist@5/swagger-ui-bundle.js\"></script>" +
            "<script>SwaggerUIBundle({url:'/openapi.json',dom_id:'#swagger-ui'," +
            "presets:[SwaggerUIBundle.presets.apis,SwaggerUIBundle.SwaggerUIStandalonePreset]," +
            "layout:'BaseLayout'})</script></body></html>",
            "text/html", System.Text.Encoding.UTF8));

        return app;
    }
}
