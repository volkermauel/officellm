/**
 * Communication module for the Office JS Add-in.
 * Connects TO the MCP server — registers as an instance,
 * polls for commands, and reports results.
 */

const MCP_SERVER_URL = "http://127.0.0.1:3000";

// --- State ---
let instanceId: string | null = null;
let officeReady = false;

// ============================================================
// INSTANCE REGISTRATION & HEARTBEAT
// ============================================================

/**
 * Registers this add-in instance with the MCP server.
 * Returns the assigned instance ID.
 */
export async function registerWithMcp(
	appName: string,
	documentName: string,
): Promise<string> {
	const response = await fetch(`${MCP_SERVER_URL}/instances/register`, {
		method: "POST",
		headers: { "Content-Type": "application/json" },
		body: JSON.stringify({ appName, documentName }),
	});

	if (!response.ok) {
		throw new Error(`Registration failed: HTTP ${response.status}`);
	}

	const data = await response.json();
	instanceId = data.instanceId;
	console.log(`Registered with MCP server: ${instanceId}`);
	return instanceId;
}

/**
 * Sends a heartbeat to keep this instance alive.
 * Should be called periodically (e.g., every 10 seconds).
 */
export async function sendHeartbeat(): Promise<void> {
	if (!instanceId) return;

	try {
		await fetch(`${MCP_SERVER_URL}/instances/${instanceId}/heartbeat`, {
			method: "POST",
			headers: { "Content-Type": "application/json" },
			body: JSON.stringify({ appName: "PowerPoint", documentName: "(active)" }),
		});
	} catch (error) {
		console.warn("Heartbeat failed:", error);
	}
}

/**
 * Starts periodic heartbeat polling.
 */
export function startHeartbeat(intervalMs = 10000): void {
	sendHeartbeat(); // Send immediately
	setInterval(sendHeartbeat, intervalMs);
}

// ============================================================
// COMMAND POLLING
// ============================================================

interface PendingCommand {
	id: string;
	command: string;
	args?: unknown;
}

/**
 * Polls the MCP server for pending commands.
 * Returns an array of unclaimed commands.
 */
export async function pollForCommands(): Promise<PendingCommand[]> {
	if (!instanceId) return [];

	try {
		const response = await fetch(
			`${MCP_SERVER_URL}/instances/${instanceId}/commands`,
		);
		if (!response.ok) return [];

		const data = await response.json();
		return data.commands || [];
	} catch (error) {
		console.warn("Command poll failed:", error);
		return [];
	}
}

/**
 * Reports a command result back to the MCP server.
 */
export async function reportResult(
	commandId: string,
	success: boolean,
	error?: string,
	payload?: unknown,
): Promise<void> {
	if (!instanceId) return;

	const body: Record<string, unknown> = {
		commandId,
		success,
	};
	if (error) body.error = error;
	if (payload) body.payload = payload;

	try {
		await fetch(`${MCP_SERVER_URL}/instances/${instanceId}/result`, {
			method: "POST",
			headers: { "Content-Type": "application/json" },
			body: JSON.stringify(body),
		});
	} catch (error) {
		console.warn("Result report failed:", error);
	}
}

/**
 * Starts periodic command polling.
 */
export function startCommandPolling(intervalMs = 2000): void {
	setInterval(pollForCommands, intervalMs);
}

// ============================================================
// OFFICE STATE RETRIEVAL
// ============================================================

export interface OfficeState {
	app: string;
	documentName: string;
	slideCount: number;
	currentSlideIndex: number;
}

/**
 * Gets basic Office state information.
 */
export function getOfficeState(): Promise<OfficeState> {
	return new Promise((resolve) => {
		Office.onReady((info) => {
			officeReady = true;
			resolve({
				app: info.host || "Unknown",
				documentName: "(loading...)",
				slideCount: 0,
				currentSlideIndex: 0,
			});
		});
	});
}

/**
 * Sends the current Office context to Open WebUI via MCP.
 */
export async function sendContextToLLM(): Promise<void> {
	if (!instanceId) {
		console.warn("Not registered with MCP server");
		return;
	}

	try {
		const response = await fetch(`${MCP_SERVER_URL}/mcp`, {
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
		console.log("Context sent to LLM:", data);
	} catch (error) {
		console.error("Failed to send context:", error);
	}
}
