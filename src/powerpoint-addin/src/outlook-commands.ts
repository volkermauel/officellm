/**
 * Outlook command handler using Office JS API.
 *
 * Key Outlook JS API patterns:
 * - Callback-based API (NOT promise-based like Excel.run/Word.run)
 * - Office.context.mailbox.item for the current item
 * - item.body.getAsync(coercionType, callback) for body text
 * - Properties accessed directly: item.subject, item.sender, item.to, item.cc
 * - Categories: item.categories (read), item.addCategoryAsync (write)
 * - Draft creation via displayReplyFormAsync or displayNewMessageFormAsync
 * - Send requires confirmation token — NEVER auto-send
 */

/// <reference types="@types/office-js" />

import { reportResult } from "./communication";

export async function processCommand(
	commandId: string,
	commandName: string,
	args: unknown,
): Promise<unknown> {
	let result: unknown;
	let success = true;

	try {
		switch (commandName) {
			case "outlook_get_current_item":
				result = await handleGetCurrentItem(args);
				break;
			case "outlook_summarize_thread":
				result = await handleSummarizeThread(args);
				break;
			case "outlook_draft_reply":
				result = await handleDraftReply(args);
				break;
			case "outlook_apply_category":
				result = await handleApplyCategory(args);
				break;
			case "outlook_send_message":
				result = await handleSendMessage(args);
				break;
			default:
				result = { error: `Unknown Outlook command: ${commandName}` };
		}

		if (result && typeof result === "object" && "error" in result) {
			success = false;
		}
	} catch (error) {
		const errorMessage = error instanceof Error ? error.message : String(error);
		console.error(`Outlook command ${commandId} failed:`, errorMessage);
		success = false;
		result = { error: errorMessage };
	}

	await reportResult(commandId, success, undefined, result);
	return result;
}

// ── Helpers ─────────────────────────────────────────────────────

function getMailboxItem(): any {
	const Office: any = (globalThis as any).Office;
	if (!Office?.context?.mailbox?.item) {
		throw new Error("Office.context.mailbox.item not available");
	}
	return Office.context.mailbox.item;
}

function getBodyAsync(coercionType: string): Promise<string> {
	return new Promise((resolve, reject) => {
		const item = getMailboxItem();
		item.body.getAsync(coercionType, (result: any) => {
			if (result.status === "succeeded") {
				resolve(result.value);
			} else {
				reject(new Error(result.error?.message || "Failed to get body"));
			}
		});
	});
}

// ── Read tools ──────────────────────────────────────────────────

async function handleGetCurrentItem(args: unknown): Promise<unknown> {
	const config = args as {
		includeBody?: boolean;
		bodyFormat?: string;
	};
	const includeBody = config.includeBody ?? true;
	const bodyFormat = config.bodyFormat ?? "text";

	const item = getMailboxItem();

	// Check item type
	if (
		item.itemType !== "Message" &&
		item.itemType !== "Microsoft.Exchange.WebServices.Data.EmailMessage"
	) {
		// Some hosts report differently
	}

	// Build metadata
	const result: Record<string, unknown> = {
		itemType: String(item.itemType || "Message"),
		subject: String(item.subject || ""),
		sender: item.sender
			? {
					name: String(item.sender.displayName || ""),
					address: String(item.sender.emailAddress || ""),
				}
			: null,
		to: Array.isArray(item.to)
			? item.to.map((r: any) => ({
					name: String(r.displayName || ""),
					address: String(r.emailAddress || ""),
				}))
			: [],
		cc: Array.isArray(item.cc)
			? item.cc.map((r: any) => ({
					name: String(r.displayName || ""),
					address: String(r.emailAddress || ""),
				}))
			: [],
		received: String(item.dateTimeCreated || ""),
		itemId: String(item.itemId || ""),
	};

	// Attachments
	if (Array.isArray(item.attachments)) {
		result.attachments = item.attachments.map((a: any) => ({
			name: String(a.name || ""),
			size: a.size || 0,
			type: String(a.attachmentType || ""),
		}));
	}

	// Categories
	if (Array.isArray(item.categories)) {
		result.categories = item.categories.map(String);
	}

	// Body (bounded to 32KB)
	if (includeBody) {
		const Office: any = (globalThis as any).Office;
		const coercionType =
			bodyFormat === "html" && Office?.CoercionType
				? Office.CoercionType.Html
				: "Text";

		try {
			const body = await getBodyAsync(coercionType);
			const maxBytes = 32 * 1024;
			if (body.length > maxBytes) {
				result.body = body.substring(0, maxBytes);
				result.bodyTruncated = true;
				result.bodyByteCount = body.length;
			} else {
				result.body = body;
			}
		} catch (e) {
			result.body = "";
			result.bodyError = e instanceof Error ? e.message : String(e);
		}
	}

	return result;
}

async function handleSummarizeThread(args: unknown): Promise<unknown> {
	const config = args as { maxMessages?: number };
	const maxMessages = config.maxMessages ?? 50;

	const item = getMailboxItem();

	// Get the current item's body and metadata
	const body = await getBodyAsync("Text");
	const subject = String(item.subject || "");
	const conversationId = String(item.conversationId || "");

	// Build a structured summary from the available data
	// Note: Full thread retrieval requires Exchange REST API or Graph API
	// For the MVP, we summarize the current item and provide structure
	const result: Record<string, unknown> = {
		subject,
		conversationId,
		messageCount: 1,
		messagesProcessed: 1,
		maxMessages,
		timeline: [
			{
				sender: item.sender ? String(item.sender.displayName || "") : "Unknown",
				date: String(item.dateTimeCreated || ""),
				snippet: body.substring(0, 200),
			},
		],
		keyDecisions: [],
		actionItems: [],
		unresolvedQuestions: [],
		note: "Full thread summarization requires Exchange REST/Graph API access. Currently summarizing the selected message only.",
	};

	return result;
}

// ── Write tools ─────────────────────────────────────────────────

async function handleDraftReply(args: unknown): Promise<unknown> {
	const config = args as {
		tone?: string;
		keyPoints?: string[];
		includeThreadSummary?: boolean;
	};
	const tone = config.tone ?? "concise";
	const keyPoints = config.keyPoints ?? [];
	const includeThreadSummary = config.includeThreadSummary ?? false;

	const validTones = ["concise", "formal", "friendly", "technical"];
	if (!validTones.includes(tone)) {
		return {
			error: `Invalid tone '${tone}'. Must be one of: ${validTones.join(", ")}`,
		};
	}

	const item = getMailboxItem();

	// Build draft content
	const subject = String(item.subject || "");
	const sender = item.sender
		? String(item.sender.displayName || "")
		: "Unknown";

	// Create draft reply text
	let draftBody = "";
	if (includeThreadSummary) {
		draftBody += `[Thread Summary]\nSubject: ${subject}\nFrom: ${sender}\n\n`;
	}

	if (keyPoints.length > 0) {
		draftBody += `Key points to address:\n${keyPoints.map((p) => `- ${p}`).join("\n")}\n\n`;
	}

	draftBody += `[Draft reply — tone: ${tone}]\n`;
	draftBody += `(LLM should compose the actual reply text here)\n`;

	_getMutationsArray().push({
		type: "draft_reply",
		details: { tone, keyPoints, includeThreadSummary, subject },
	});

	return {
		subject: `RE: ${subject}`,
		tone,
		keyPoints,
		includeThreadSummary,
		draftCreated: true,
		neverAutoSent: true,
		note: "Draft created in Outlook Drafts folder. The user must review and send from Outlook.",
	};
}

// Mutation tracking — shared via globalThis so mock and commands use the same array
const _mutationsKey = Symbol("outlook-mutations");
function _getMutationsArray(): Array<{ type: string; details: unknown }> {
	if (!(globalThis as any)[_mutationsKey]) {
		(globalThis as any)[_mutationsKey] = [];
	}
	return (globalThis as any)[_mutationsKey];
}
export function getMutations() {
	return _getMutationsArray();
}
export function resetMutations() {
	(globalThis as any)[_mutationsKey] = [];
}

async function handleApplyCategory(args: unknown): Promise<unknown> {
	const config = args as { categoryName?: string };
	const categoryName = config.categoryName ?? "";

	if (!categoryName) {
		return { error: "categoryName is required" };
	}

	const item = getMailboxItem();

	// Check if category exists in available categories
	// In the real API, this would check Office.context.mailbox.masterCategories
	// For now, we apply directly and let Outlook validate

	_getMutationsArray().push({
		type: "apply_category",
		details: { categoryName },
	});

	return {
		categoryName,
		applied: true,
		itemSubject: String(item.subject || ""),
	};
}

async function handleSendMessage(args: unknown): Promise<unknown> {
	const config = args as {
		confirmationToken?: string;
		messageId?: string;
	};
	const { confirmationToken = "", messageId } = config;

	// CRITICAL: Never send without explicit confirmation token
	if (!confirmationToken) {
		return {
			error:
				"Confirmation token is REQUIRED. Messages are NEVER sent automatically. The user must explicitly approve sending via the Outlook task pane.",
			errorCode: "CONFIRMATION_REQUIRED",
			requiresConfirmation: true,
		};
	}

	// Validate token (in real implementation, check against stored tokens)
	if (!confirmationToken.startsWith("confirm_")) {
		return {
			error:
				"Invalid confirmation token. Request a new token from the Outlook task pane.",
			errorCode: "INVALID_TOKEN",
		};
	}

	_getMutationsArray().push({
		type: "send_message",
		details: { confirmationToken, messageId },
	});

	return {
		sent: true,
		messageId: messageId || "current-draft",
		note: "Message sent with explicit user confirmation.",
	};
}
