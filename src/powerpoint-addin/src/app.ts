/**
 * Main entry point for the Office LLM Harness PowerPoint Add-in.
 * Registers with the MCP server, polls for commands, and manages the task pane UI.
 */

/// <reference types="@types/office-js" />

import {
  registerWithMcp,
  startHeartbeat,
  startCommandPolling,
  pollForCommands,
  getOfficeState,
  sendContextToLLM,
} from "./communication";

// --- State ---
let instanceId: string | null = null;

// --- Initialization ---

Office.onReady((info) => {
  console.log(`Office ready: ${info.host}`);
  updateOfficeStatus(true);

  // Register with MCP server
  initWithMcp();
});

async function initWithMcp(): Promise<void> {
  try {
    const state = await getOfficeState();
    instanceId = await registerWithMcp(state.app, state.documentName);
    updateContextDisplay(state.documentName, state.slideCount, state.currentSlideIndex);

    // Start heartbeat and command polling
    startHeartbeat(10000);
    startCommandPolling(2000);

    console.log(`Add-in registered as: ${instanceId}`);
    addLogEntry(`Registered as ${instanceId}`);

    // Check for pending commands
    await processPendingCommands();
  } catch (error) {
    console.error("Registration failed:", error);
    addLogEntry(`Registration failed: ${error instanceof Error ? error.message : String(error)}`);
  }
}

// --- UI Updates ---

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

  const loading = log.querySelector(".loading");
  if (loading) loading.remove();

  const entry = document.createElement("div");
  entry.className = "log-entry";
  const time = new Date().toLocaleTimeString();
  entry.textContent = `[${time}] ${message}`;

  log.insertBefore(entry, log.firstChild);

  // Keep only last 20 entries
  while (log.children.length > 20) {
    log.removeChild(log.lastChild!);
  }
}

// --- Command Processing ---

async function processPendingCommands(): Promise<void> {
  if (!instanceId) return;

  try {
    const commands = await pollForCommands();
    for (const cmd of commands) {
      addLogEntry(`Received command: ${cmd.command}`);
      // Commands will be executed when the Office JS API is available
      // For now, just acknowledge receipt
      addLogEntry(`Command ${cmd.id} queued for execution`);
    }
  } catch (error) {
    addLogEntry(`Command poll error: ${error instanceof Error ? error.message : String(error)}`);
  }
}

// --- Actions ---

async function refreshContext(): Promise<void> {
  addLogEntry("Refreshing context...");

  try {
    const state = await getOfficeState();
    updateContextDisplay(state.documentName, state.slideCount, state.currentSlideIndex);
    addLogEntry(`Context refreshed: ${state.documentName}`);

    // Re-register with updated info
    if (instanceId) {
      instanceId = await registerWithMcp(state.app, state.documentName);
      addLogEntry(`Re-registered as ${instanceId}`);
    }
  } catch (error) {
    addLogEntry(`Error: ${error instanceof Error ? error.message : String(error)}`);
  }
}

async function sendToLLM(): Promise<void> {
  addLogEntry("Sending context to LLM via MCP...");

  try {
    await sendContextToLLM();
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

// --- Start ---

console.log("Office LLM Harness PowerPoint Add-in loaded");
