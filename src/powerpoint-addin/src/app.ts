/**
 * Main entry point for the Office LLM Harness PowerPoint Add-in.
 * Registers with the MCP server, polls for commands, and manages the task pane UI.
 */

/// <reference types="@types/office-js" />

import {
	registerWithMcp,
	startHeartbeat,
	pollForCommands,
	getOfficeState,
	connectSignalR,
	setCommandHandler,
	onConnectionStateChange,
} from "./communication";
import { processCommand as processPptCommand } from "./powerpoint-commands";
import { processCommand as processWordCommand } from "./word-commands";
import { processCommand as processExcelCommand } from "./excel-commands";
import { processCommand as processOutlookCommand } from "./outlook-commands";

// --- Host-aware command dispatch ---
const HOST_DISPATCH: Record<string, typeof processPptCommand> = {
	powerpoint_: processPptCommand,
	word_: processWordCommand,
	excel_: processExcelCommand,
	outlook_: processOutlookCommand,
};

function processCommand(
	commandId: string,
	commandName: string,
	args: unknown,
): Promise<unknown> {
	for (const [prefix, handler] of Object.entries(HOST_DISPATCH)) {
		if (commandName.startsWith(prefix)) {
			return handler(commandId, commandName, args);
		}
	}
	return Promise.resolve({ error: `Unknown host for command: ${commandName}` });
}

// --- State ---
let instanceId: string | null = null;

// --- Initialization ---
let isInitialized = false;

Office.onReady((info) => {
	console.log(`Office ready: ${info.host}`);
	updateOfficeStatus(true);

	// Register with MCP server (only once)
	if (!isInitialized) {
		isInitialized = true;
		initWithMcp();
	}
});

function updateMcpStatus(connected: boolean, text?: string): void {
	const el = document.getElementById("mcpStatus");
	if (!el) return;
	el.className = connected ? "badge connected" : "badge disconnected";
	el.textContent = text ?? (connected ? "Connected" : "Disconnected");
}

async function initWithMcp(): Promise<void> {
	updateMcpStatus(false, "Connecting...");
	try {
		const state = await getOfficeState();
		instanceId = await registerWithMcp(state.app, state.documentName);
		updateMcpStatus(true);
		updateContextDisplay(
			state.documentName,
			state.slideCount,
			state.currentSlideIndex,
		);

		// Start heartbeat
		startHeartbeat(10000);

		// Set up command handler for SignalR
		setCommandHandler(async (commandId, commandName, args) => {
			addLogEntry(`[SignalR] Executing: ${commandName}`);
			const result = await processCommand(commandId, commandName, args);
			addLogEntry(`[SignalR] Command ${commandId} completed`);
			return result;
		});

		// Monitor connection state
		onConnectionStateChange((state) => {
			if (state === "connected") {
				updateMcpStatus(true, "Connected (SignalR)");
			} else if (state === "reconnecting") {
				updateMcpStatus(false, "Reconnecting...");
			} else {
				updateMcpStatus(true, "Connected (HTTP polling)");
			}
		});

		// Connect SignalR (falls back to HTTP polling automatically)
		await connectSignalR();

		// Always start HTTP polling as fallback (commands may arrive while reconnecting)
		setInterval(() => processPendingCommands(), 2000);

		console.log(`Add-in registered as: ${instanceId}`);
		addLogEntry(`Registered as ${instanceId}`);

		// Process any already-queued commands
		await processPendingCommands();
	} catch (error) {
		console.error("Registration failed:", error);
		updateMcpStatus(false, "Failed");
		addLogEntry(
			`Registration failed: ${error instanceof Error ? error.message : String(error)}`,
		);
	}
}

// --- Diff Preview & Confirmation (removed) ---

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
	currentSlide: number,
): void {
	const section = document.getElementById("contextSection");
	const nameEl = document.getElementById("docName");
	const slideEl = document.getElementById("slideInfo");
	const countEl = document.getElementById("slideCount");

	if (section) section.style.display = "block";
	if (nameEl) nameEl.textContent = docName || "Untitled";
	if (slideEl)
		slideEl.textContent = currentSlide >= 0 ? `Slide ${currentSlide + 1}` : "—";
	if (countEl)
		countEl.textContent = `${slideCount} slide${slideCount !== 1 ? "s" : ""}`;
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

let isProcessingCommands = false;

async function processPendingCommands(): Promise<void> {
	if (!instanceId || isProcessingCommands) return;

	isProcessingCommands = true;
	try {
		const commands = await pollForCommands();
		for (const cmd of commands) {
			addLogEntry(`Executing: ${cmd.command}`);
			await processCommand(cmd.id, cmd.command, cmd.args);
			addLogEntry(`Command ${cmd.id} completed`);
		}
	} catch (error) {
		addLogEntry(
			`Command poll error: ${error instanceof Error ? error.message : String(error)}`,
		);
	} finally {
		isProcessingCommands = false;
	}
}

// --- Actions ---

async function refreshContext(): Promise<void> {
	addLogEntry("Refreshing context...");

	try {
		const state = await getOfficeState();
		updateContextDisplay(
			state.documentName,
			state.slideCount,
			state.currentSlideIndex,
		);
		addLogEntry(`Context refreshed: ${state.documentName}`);

		// Re-register with updated info
		if (instanceId) {
			instanceId = await registerWithMcp(state.app, state.documentName);
			addLogEntry(`Re-registered as ${instanceId}`);
		}
	} catch (error) {
		addLogEntry(
			`Error: ${error instanceof Error ? error.message : String(error)}`,
		);
	}
}

// --- Actions ---

// --- Expose functions for HTML onclick handlers ---
declare global {
	interface Window {
		refreshContext: () => Promise<void>;
	}
}

window.refreshContext = refreshContext;

// --- Start ---

console.log("Office LLM Harness PowerPoint Add-in loaded");
