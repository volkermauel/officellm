/**
 * Main entry point for the Office LLM Harness PowerPoint Add-in.
 * Initializes Office JS, sets up the task pane UI, and registers
 * the HTTP endpoint for MCP server communication.
 */

/// <reference types="@types/office-js" />

import {
  getActiveApp,
  getDeckOutline,
  getSlide,
  updateShapeText,
  updateSpeakerNotes,
} from "./communication";

// --- State ---
let officeReady = false;
let mcpServerReachable = false;

// --- Initialization ---

Office.onReady((info) => {
  console.log(`Office ready: ${info.host} (${info.state})`);

  if (info.state === "Ready") {
    officeReady = true;
    updateOfficeStatus(true);
    refreshContext();
  } else {
    // Wait a bit more for Office to fully initialize
    setTimeout(() => {
      if (info.state !== "Ready") {
        console.warn("Office not ready after timeout");
        updateOfficeStatus(false);
      }
    }, 3000);
  }
});

// --- MCP Server Health Check ---

async function checkMcpServer(): Promise<void> {
  try {
    const response = await fetch("http://127.0.0.1:3000/health", {
      method: "GET",
      signal: AbortSignal.timeout(3000),
    });
    mcpServerReachable = response.ok;
    updateMcpStatus(mcpServerReachable);
  } catch {
    mcpServerReachable = false;
    updateMcpStatus(false);
  }
}

// Run health check periodically
checkMcpServer();
setInterval(checkMcpServer, 10000);

// --- UI Updates ---

function updateMcpStatus(connected: boolean): void {
  const el = document.getElementById("mcpStatus");
  if (!el) return;
  el.className = connected ? "badge connected" : "badge disconnected";
  el.textContent = connected ? "Connected" : "Disconnected";
}

function updateOfficeStatus(ready: boolean): void {
  const el = document.getElementById("officeStatus");
  if (!el) return;
  el.className = ready ? "badge connected" : "badge disconnected";
  el.textContent = ready ? "Ready" : "Not Ready";
}

function updateContextDisplay(
  docName: string,
  slideCount: number,
  currentSlide: number
): void {
  const section = document.getElementById("contextSection");
  const nameEl = document.getElementById("docName");
  const slideEl = document.getElementById("slideInfo");
  const countEl = document.getElementById("slideCount");

  if (section) section.style.display = "block";
  if (nameEl) nameEl.textContent = docName || "Untitled";
  if (slideEl)
    slideEl.textContent = currentSlide >= 0 ? `Slide ${currentSlide + 1}` : "—";
  if (countEl) countEl.textContent = `${slideCount} slide${slideCount !== 1 ? "s" : ""}`;
}

function addLogEntry(message: string): void {
  const log = document.getElementById("activityLog");
  if (!log) return;

  // Remove loading indicator if present
  const loading = log.querySelector(".loading");
  if (loading) loading.remove();

  const entry = document.createElement("div");
  entry.className = "log-entry";
  const time = new Date().toLocaleTimeString();
  entry.textContent = `[${time}] ${message}`;

  // Prepend (newest first)
  log.insertBefore(entry, log.firstChild);

  // Keep only last 20 entries
  while (log.children.length > 20) {
    log.removeChild(log.lastChild!);
  }
}

// --- Actions ---

async function refreshContext(): Promise<void> {
  addLogEntry("Refreshing context...");

  try {
    const result = await getActiveApp();
    if (result.ok && result.result) {
      const data = result.result as {
        app: string;
        documentName: string;
        slideCount: number;
      };
      updateContextDisplay(data.documentName, data.slideCount, 0);
      addLogEntry(`Context refreshed: ${data.documentName} (${data.slideCount} slides)`);
    } else {
      addLogEntry(`Error: ${result.error || "Unknown error"}`);
    }
  } catch (error) {
    addLogEntry(`Connection error: ${error instanceof Error ? error.message : String(error)}`);
  }
}

async function sendToLLM(): Promise<void> {
  addLogEntry("Sending context to LLM via MCP...");

  try {
    // Call the MCP server's office_get_active_app tool
    const response = await fetch("http://127.0.0.1:3000/mcp", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        jsonrpc: "2.0",
        id: Date.now(),
        method: "tools/call",
        params: { name: "office_get_active_app", input: {} },
      }),
    });

    if (!response.ok) {
      throw new Error(`MCP server returned HTTP ${response.status}`);
    }

    const data = await response.json();
    addLogEntry("Context sent to LLM successfully");
  } catch (error) {
    addLogEntry(`Failed to send: ${error instanceof Error ? error.message : String(error)}`);
  }
}

// --- Expose functions for HTML onclick handlers ---
declare global {
  interface Window {
    refreshContext: () => Promise<void>;
    sendToLLM: () => Promise<void>;
  }
}

window.refreshContext = refreshContext;
window.sendToLLM = sendToLLM;

// --- Local HTTP Server for MCP Communication ---
// The MCP server calls this endpoint on port 8765.
// Since we can't run a real server in the Office iframe,
// we use a workaround: the MCP server sends requests via
// postMessage to a small proxy page, or we use a shared
// localStorage/IndexedDB message bus.

// For the spike, we implement a simple approach:
// The add-in exposes commands through a global registry
// that the MCP server can reach via a small local proxy.
// The proxy is a minimal HTML page served from the same
// webpack dev server that acts as a bridge.

// Bridge endpoint: handled by webpack dev server's proxy
// or by a dedicated Express/Koa server in production.

console.log("Office LLM Harness PowerPoint Add-in loaded");
