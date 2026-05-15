/**
 * Communication module for the Office JS Add-in.
 * Handles HTTP communication with the local MCP server.
 *
 * For the spike, the add-in serves as a simple HTTP endpoint that
 * the MCP server calls to get Office state information.
 */

const MCP_SERVER_URL = "http://127.0.0.1:3000";

// --- MCP Server Communication ---

export async function callMcpTool(
  toolName: string,
  args: Record<string, unknown> = {}
): Promise<unknown> {
  const messageId = Date.now();
  const payload = JSON.stringify({
    jsonrpc: "2.0",
    id: messageId,
    method: "tools/call",
    params: { name: toolName, input: args },
  });

  try {
    const response = await fetch(`${MCP_SERVER_URL}/mcp`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: payload,
    });

    if (!response.ok) {
      throw new Error(`MCP server returned HTTP ${response.status}`);
    }

    const data = await response.json();
    return data.result;
  } catch (error) {
    console.error("MCP tool call failed:", error);
    throw error;
  }
}

// --- Office State Retrieval ---
// These functions interact with the Office object model
// via async callbacks (Office JS pattern).

export interface OfficeState {
  app: string;
  documentName: string;
  slideCount: number;
  currentSlideIndex: number;
}

/**
 * Gets basic Office state information.
 * Uses Office.onReady() to ensure the API is available.
 */
export function getOfficeState(): Promise<OfficeState> {
  return new Promise((resolve, reject) => {
    Office.onReady((info) => {
      if (info.host === Office.HostType.PowerPoint || info.host === Office.HostType.Presentation) {
        // PowerPoint is active - gather basic state
        resolve({
          app: "PowerPoint",
          documentName: "(loading...)",
          slideCount: 0,
          currentSlideIndex: 0,
        });
      } else {
        resolve({
          app: info.host || "Unknown",
          documentName: "(no document)",
          slideCount: 0,
          currentSlideIndex: 0,
        });
      }
    });
  });
}

/**
 * Sends the current Office context to Open WebUI via MCP.
 */
export async function sendContextToLLM(): Promise<void> {
  await callMcpTool("office_get_active_app", {});
}
