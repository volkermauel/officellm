/**
 * Word command handler using Office JS API.
 *
 * Key Word JS API patterns:
 * - Word.run(async (context) => { ... }) for batched operations
 * - context.document.body.paragraphs for paragraph access
 * - paragraph.load("text,style,outlineLevel,uniqueLocalId") for properties
 * - context.document.getSelection() for current selection
 * - range.insertComment(text) for comments
 * - Must sync() before reading any loaded property
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
			case "word_get_outline":
				result = await handleGetOutline(args);
				break;
			case "word_get_paragraphs":
				result = await handleGetParagraphs(args);
				break;
			case "word_get_selection":
				result = await handleGetSelection(args);
				break;
			case "word_search":
				result = await handleSearch(args);
				break;
			case "word_replace_text":
				result = await handleReplaceText(args);
				break;
			case "word_insert_text":
				result = await handleInsertText(args);
				break;
			case "word_add_comment":
				result = await handleAddComment(args);
				break;
			case "word_delete_paragraph":
				result = await handleDeleteParagraph(args);
				break;
			default:
				result = { error: `Unknown Word command: ${commandName}` };
		}

		if (result && typeof result === "object" && "error" in result) {
			success = false;
		}
	} catch (error) {
		const errorMessage = error instanceof Error ? error.message : String(error);
		console.error(`Word command ${commandId} failed:`, errorMessage);
		success = false;
		result = { error: errorMessage };
	}

	await reportResult(commandId, success, undefined, result);
	return result;
}

// ── Helpers ─────────────────────────────────────────────────────

function runInWord<T>(fn: (ctx: any) => Promise<T>): Promise<T> {
	return new Promise<T>((resolve, reject) => {
		const Word: any = (window as any).Word;
		if (!Word || typeof Word.run !== "function") {
			reject(new Error("Word.run() not available"));
			return;
		}
		Word.run(async (ctx: any) => {
			resolve(await fn(ctx));
		}).catch(reject);
	});
}

function safeStr(val: any, fallback = ""): string {
	return val != null ? String(val) : fallback;
}

// ── Read tools ──────────────────────────────────────────────────

async function handleGetOutline(args: unknown): Promise<unknown> {
	const config = args as { maxDepth?: number };
	const maxDepth = config.maxDepth ?? 3;

	return runInWord(async (ctx) => {
		const paragraphs = ctx.document.body.paragraphs;
		paragraphs.load("items");
		await ctx.sync();

		// Load heading-relevant properties
		for (const p of paragraphs.items) {
			p.load("text,style,outlineLevel,uniqueLocalId");
		}
		await ctx.sync();

		const headings: Array<{
			index: number;
			level: number;
			text: string;
			style: string;
			id: string;
		}> = [];

		for (let i = 0; i < paragraphs.items.length; i++) {
			const p = paragraphs.items[i];
			const outlineLevel = String(p.outlineLevel || "");
			const style = String(p.style || "");

			// Extract numeric level from outlineLevel (e.g., "OutlineLevel1" → 1)
			let level = 0;
			const match = outlineLevel.match(/OutlineLevel(\d)/);
			if (match) {
				level = parseInt(match[1]);
			} else if (/Heading\s*(\d)/i.test(style)) {
				const headingMatch = style.match(/Heading\s*(\d)/i);
				if (headingMatch) level = parseInt(headingMatch[1]);
			}

			if (level > 0 && level <= maxDepth) {
				headings.push({
					index: i,
					level,
					text: String(p.text || "").trim(),
					style: String(p.style || ""),
					id: String(p.uniqueLocalId || ""),
				});
			}
		}

		return {
			documentName: "Document",
			totalParagraphs: paragraphs.items.length,
			headings,
		};
	});
}

async function handleGetParagraphs(args: unknown): Promise<unknown> {
	const config = args as { startIndex?: number; count?: number };
	const startIndex = config.startIndex ?? 0;
	const count = config.count ?? 50;

	return runInWord(async (ctx) => {
		const paragraphs = ctx.document.body.paragraphs;
		paragraphs.load("items");
		await ctx.sync();

		const total = paragraphs.items.length;
		const end = Math.min(startIndex + count, total);
		const result: Array<{
			index: number;
			text: string;
			style: string;
			id: string;
		}> = [];

		for (let i = startIndex; i < end; i++) {
			paragraphs.items[i].load("text,style,uniqueLocalId");
		}
		await ctx.sync();

		for (let i = startIndex; i < end; i++) {
			const p = paragraphs.items[i];
			result.push({
				index: i,
				text: String(p.text || ""),
				style: String(p.style || ""),
				id: String(p.uniqueLocalId || ""),
			});
		}

		return { totalParagraphs: total, paragraphs: result };
	});
}

async function handleGetSelection(_args: unknown): Promise<unknown> {
	return runInWord(async (ctx) => {
		const selection = ctx.document.getSelection();
		selection.load("text");
		await ctx.sync();

		const text = String(selection.text || "");
		const paragraphs = selection.paragraphs;
		paragraphs.load("items");
		await ctx.sync();

		// Load paragraph context
		for (const p of paragraphs.items) {
			p.load("text,style");
		}
		await ctx.sync();

		const contextParagraphs = paragraphs.items.map((p: any, i: number) => ({
			index: i,
			text: String(p.text || ""),
			style: String(p.style || ""),
		}));

		return {
			type: text ? "text" : "empty",
			text,
			paragraphs: contextParagraphs,
		};
	});
}

async function handleSearch(args: unknown): Promise<unknown> {
	const config = args as { searchText?: string; matchCase?: boolean };
	const searchText = config.searchText ?? "";
	const matchCase = config.matchCase ?? false;

	if (!searchText) return { error: "searchText is required" };

	return runInWord(async (ctx) => {
		const searchResults = ctx.document.body.search(searchText, { matchCase });
		searchResults.load("items");
		await ctx.sync();

		const results: Array<{ index: number; text: string; matchText: string }> = [];

		// Load parent paragraphs for context
		for (let i = 0; i < searchResults.items.length; i++) {
			const range = searchResults.items[i];
			range.load("text");
		}
		await ctx.sync();

		for (let i = 0; i < searchResults.items.length; i++) {
			const range = searchResults.items[i];
			results.push({
				index: i,
				text: String(range.text || ""),
				matchText: String(range.text || ""),
			});
		}

		return {
			searchText,
			totalMatches: results.length,
			matches: results,
		};
	});
}

// ── Write tools ─────────────────────────────────────────────────

async function handleReplaceText(args: unknown): Promise<unknown> {
	const config = args as { paragraphIndex?: number; oldText?: string; newText?: string };
	const { paragraphIndex = 0, oldText = "", newText = "" } = config;

	return runInWord(async (ctx) => {
		const paragraphs = ctx.document.body.paragraphs;
		paragraphs.load("items");
		await ctx.sync();

		if (paragraphIndex < 0 || paragraphIndex >= paragraphs.items.length) {
			return { error: `Paragraph index ${paragraphIndex} out of range (0-${paragraphs.items.length - 1})` };
		}

		const paragraph = paragraphs.items[paragraphIndex];
		paragraph.load("text");
		await ctx.sync();

		const currentText = String(paragraph.text || "");
		if (!currentText.includes(oldText)) {
			return { error: `Text '${oldText}' not found in paragraph ${paragraphIndex}` };
		}

		// Use search within the paragraph to do targeted replace
		const searchResults = paragraph.search(oldText, { matchCase: true });
		searchResults.load("items");
		await ctx.sync();

		if (searchResults.items.length > 0) {
			searchResults.items[0].insertText(newText, "Replace");
			await ctx.sync();
		}

		return { paragraphIndex, oldText, newText, replaced: true };
	});
}

async function handleInsertText(args: unknown): Promise<unknown> {
	const config = args as {
		text?: string;
		insertLocation?: string;
		paragraphIndex?: number;
	};
	const { text = "", insertLocation = "end", paragraphIndex } = config;

	return runInWord(async (ctx) => {
		if (insertLocation === "end" || paragraphIndex === undefined) {
			// Insert at the end of the document
			ctx.document.body.insertParagraph(text, "End");
			await ctx.sync();
			return { text, insertLocation: "end", inserted: true };
		}

		// Insert relative to a specific paragraph
		const paragraphs = ctx.document.body.paragraphs;
		paragraphs.load("items");
		await ctx.sync();

		if (paragraphIndex < 0 || paragraphIndex >= paragraphs.items.length) {
			return { error: `Paragraph index ${paragraphIndex} out of range` };
		}

		const paragraph = paragraphs.items[paragraphIndex];
		const location = insertLocation === "beforeParagraph" ? "Before" : "After";
		paragraph.insertParagraph(text, location);
		await ctx.sync();

		return { text, insertLocation, paragraphIndex, inserted: true };
	});
}

async function handleAddComment(args: unknown): Promise<unknown> {
	const config = args as { commentText?: string; paragraphIndex?: number };
	const { commentText = "", paragraphIndex } = config;

	return runInWord(async (ctx) => {
		if (paragraphIndex !== undefined) {
			// Comment on a specific paragraph
			const paragraphs = ctx.document.body.paragraphs;
			paragraphs.load("items");
			await ctx.sync();

			if (paragraphIndex < 0 || paragraphIndex >= paragraphs.items.length) {
				return { error: `Paragraph index ${paragraphIndex} out of range` };
			}

			const paragraph = paragraphs.items[paragraphIndex];
			const range = paragraph.getRange("Whole");
			range.insertComment(commentText);
			await ctx.sync();

			return { commentText, paragraphIndex, added: true };
		}

		// Comment on current selection
		const selection = ctx.document.getSelection();
		selection.insertComment(commentText);
		await ctx.sync();

		return { commentText, target: "selection", added: true };
	});
}

async function handleDeleteParagraph(args: unknown): Promise<unknown> {
	const config = args as { paragraphIndex?: number };
	const { paragraphIndex = 0 } = config;

	return runInWord(async (ctx) => {
		const paragraphs = ctx.document.body.paragraphs;
		paragraphs.load("items");
		await ctx.sync();

		if (paragraphIndex < 0 || paragraphIndex >= paragraphs.items.length) {
			return { error: `Paragraph index ${paragraphIndex} out of range` };
		}

		const paragraph = paragraphs.items[paragraphIndex];
		paragraph.delete();
		await ctx.sync();

		return { paragraphIndex, deleted: true };
	});
}
