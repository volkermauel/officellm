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
import { processCommand } from "./powerpoint-commands";

// --- State ---
let instanceId: string | null = null;
let pendingConfirmation: {
	commandId: string;
	toolName: string;
	confirmationToken: string;
	diffPreview: unknown;
} | null = null;

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
		updateContextDisplay(
			state.documentName,
			state.slideCount,
			state.currentSlideIndex,
		);

		// Start heartbeat and command polling
		startHeartbeat(10000);
		startCommandPolling(2000);

		console.log(`Add-in registered as: ${instanceId}`);
		addLogEntry(`Registered as ${instanceId}`);

		// Check for pending commands
		await processPendingCommands();
	} catch (error) {
		console.error("Registration failed:", error);
		addLogEntry(
			`Registration failed: ${error instanceof Error ? error.message : String(error)}`,
		);
	}
}

// --- Diff Preview & Confirmation ---

interface DiffHunk {
	oldStart: number;
	oldLines: number;
	newStart: number;
	newLines: number;
	lines: string[];
}

interface DiffPreview {
	toolName: string;
	diff: {
		paths: Array<{
			path: string;
			hunks: DiffHunk[];
		}>;
	};
}

function showDiffPreview(diffData: unknown): void {
	const diffSection = document.getElementById("diffSection");
	const beforeContent = document.getElementById("diffBeforeContent");
	const afterContent = document.getElementById("diffAfterContent");

	if (!diffSection || !beforeContent || !afterContent) return;

	// Parse diff data
	const diff = diffData as DiffPreview;
	let beforeText = "";
	let afterText = "";

	if (diff.diff && diff.diff.paths) {
		for (const pathEntry of diff.diff.paths) {
			for (const hunk of pathEntry.hunks) {
				for (const line of hunk.lines) {
					if (line.startsWith("-")) {
						beforeText += line + "\n";
					} else if (line.startsWith("+")) {
						afterText += line + "\n";
					} else if (!line.startsWith("@@")) {
						beforeText += line + "\n";
						afterText += line + "\n";
					}
				}
			}
		}
	}

	beforeContent.textContent = beforeText || "(no changes)";
	afterContent.textContent = afterText || "(no changes)";

	diffSection.classList.add("active");
	addLogEntry(`Diff preview shown for ${diff.toolName}`);
}

function hideDiffPreview(): void {
	const diffSection = document.getElementById("diffSection");
	if (diffSection) diffSection.classList.remove("active");
}

function showConfirmationPending(): void {
	const el = document.getElementById("confirmationPending");
	if (el) el.classList.add("active");
}

function hideConfirmationPending(): void {
	const el = document.getElementById("confirmationPending");
	if (el) el.classList.remove("active");
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

async function processPendingCommands(): Promise<void> {
	if (!instanceId) return;

	try {
		const commands = await pollForCommands();
		for (const cmd of commands) {
			addLogEntry(`Received command: ${cmd.command}`);
			// Execute the command via Office JS API
			const result = await processCommand(cmd.id, cmd.command, cmd.args);

			// Check if confirmation is required
			if (
				result &&
				typeof result === "object" &&
				"requiresConfirmation" in result
			) {
				const confirmationData = result as {
					requiresConfirmation: boolean;
					confirmationToken?: string;
					diffPreview?: unknown;
					toolName?: string;
				};
				if (
					confirmationData.requiresConfirmation &&
					confirmationData.confirmationToken
				) {
					pendingConfirmation = {
						commandId: cmd.id,
						toolName: confirmationData.toolName || cmd.command,
						confirmationToken: confirmationData.confirmationToken,
						diffPreview: confirmationData.diffPreview,
					};
					showDiffPreview(confirmationData.diffPreview);
					showConfirmationPending();
					addLogEntry(`Confirmation required for ${cmd.command}`);
				} else {
					addLogEntry(`Command ${cmd.id} executed successfully`);
				}
			} else {
				addLogEntry(`Command ${cmd.id} executed successfully`);
			}
		}
	} catch (error) {
		addLogEntry(
			`Command poll error: ${error instanceof Error ? error.message : String(error)}`,
		);
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

// --- Confirmation Actions ---

async function approveChange(): Promise<void> {
	if (!pendingConfirmation || !instanceId) return;

	try {
		await fetch(`http://127.0.0.1:3000/instances/${instanceId}/confirm`, {
			method: "POST",
			headers: { "Content-Type": "application/json" },
			body: JSON.stringify({
				commandId: pendingConfirmation.commandId,
				confirmationToken: pendingConfirmation.confirmationToken,
				approved: true,
			}),
		});
		addLogEntry(`Change approved: ${pendingConfirmation.toolName}`);
	} catch (error) {
		addLogEntry(
			`Approval failed: ${error instanceof Error ? error.message : String(error)}`,
		);
	} finally {
		hideDiffPreview();
		hideConfirmationPending();
		pendingConfirmation = null;
	}
}

async function rejectChange(): Promise<void> {
	if (!pendingConfirmation || !instanceId) return;

	try {
		await fetch(`http://127.0.0.1:3000/instances/${instanceId}/confirm`, {
			method: "POST",
			headers: { "Content-Type": "application/json" },
			body: JSON.stringify({
				commandId: pendingConfirmation.commandId,
				confirmationToken: pendingConfirmation.confirmationToken,
				approved: false,
			}),
		});
		addLogEntry(`Change rejected: ${pendingConfirmation.toolName}`);
	} catch (error) {
		addLogEntry(
			`Rejection failed: ${error instanceof Error ? error.message : String(error)}`,
		);
	} finally {
		hideDiffPreview();
		hideConfirmationPending();
		pendingConfirmation = null;
	}
}

async function sendToLLM(): Promise<void> {
	addLogEntry("Sending context to LLM via MCP...");

	try {
		await sendContextToLLM();
		addLogEntry("Context sent to LLM successfully");
	} catch (error) {
		addLogEntry(
			`Failed to send: ${error instanceof Error ? error.message : String(error)}`,
		);
	}
}

// --- Expose functions for HTML onclick handlers ---
declare global {
	interface Window {
		refreshContext: () => Promise<void>;
		sendToLLM: () => Promise<void>;
		approveChange: () => Promise<void>;
		rejectChange: () => Promise<void>;
	}
}

window.refreshContext = refreshContext;
window.sendToLLM = sendToLLM;
window.approveChange = approveChange;
window.rejectChange = rejectChange;

// --- Start ---

console.log("Office LLM Harness PowerPoint Add-in loaded");
