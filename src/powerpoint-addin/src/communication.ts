/**
 * Communication module for the Office JS Add-in.
 * Connects TO the MCP server — registers as an instance,
 * polls for commands, and reports results.
 */

export const MCP_SERVER_URL = "http://127.0.0.1:3000";

// --- State ---
let instanceId: string | null = null;

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
	instanceId = data.instanceId ?? "";
	console.log(`Registered with MCP server: ${instanceId}`);
	return instanceId!;
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

	const url = `${MCP_SERVER_URL}/instances/${instanceId}/result`;
	const MAX_RETRIES = 3;

	for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
		try {
			const response = await fetch(url, {
				method: "POST",
				headers: { "Content-Type": "application/json" },
				body: JSON.stringify(body),
			});
			if (response.ok) return;
			if (attempt === MAX_RETRIES) {
				console.warn(`Result report failed: HTTP ${response.status} after ${MAX_RETRIES} attempts`);
			}
		} catch (err) {
			if (attempt === MAX_RETRIES) {
				console.warn(`Result report failed after ${MAX_RETRIES} attempts:`, err);
			} else {
				await new Promise((r) => setTimeout(r, 500 * attempt));
			}
		}
	}
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
			// Try to get the real document name from Office context
			let documentName = "Untitled";
			try {
				// eslint-disable-next-line @typescript-eslint/no-explicit-any
				const doc: any = (window as any).Office?.context?.document;
				if (doc?.url) {
					// Extract filename from URL path
					try {
						documentName = decodeURIComponent(
							doc.url.split("/").pop() || "Untitled",
						);
					} catch {
						documentName = doc.url;
					}
				}
			} catch {
				// Office context not available yet
			}

			resolve({
				app: (info.host as unknown as string) || "Unknown",
				documentName,
				slideCount: 0,
				currentSlideIndex: 0,
			});
		});
	});
}
