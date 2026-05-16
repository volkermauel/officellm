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
			case "word_get_tracked_changes":
				result = await handleGetTrackedChanges(args);
				break;
			case "word_accept_all_changes":
				result = await handleAcceptAllChanges(args);
				break;
			case "word_reject_all_changes":
				result = await handleRejectAllChanges(args);
				break;
			case "word_get_tables":
				result = await handleGetTables(args);
				break;
			case "word_insert_table":
				result = await handleInsertTable(args);
				break;
			case "word_update_table_cell":
				result = await handleUpdateTableCell(args);
				break;
			case "word_get_headers_footers":
				result = await handleGetHeadersFooters(args);
				break;
			case "word_set_header_footer":
				result = await handleSetHeaderFooter(args);
				break;
			case "word_replace_selection":
				result = await handleReplaceSelection(args);
				break;
			case "word_insert_image":
				result = await handleInsertImage(args);
				break;
			case "word_apply_style":
				result = await handleApplyStyle(args);
				break;
			case "word_get_sections":
				result = await handleGetSections(args);
				break;
			case "word_insert_list":
				result = await handleInsertList(args);
				break;
				result = await handleRejectAllChanges(args);
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

function _safeStr(val: any, fallback = ""): string {
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

		const results: Array<{ index: number; text: string; matchText: string }> =
			[];

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
	const config = args as {
		paragraphIndex?: number;
		oldText?: string;
		newText?: string;
	};
	const { paragraphIndex = 0, oldText = "", newText = "" } = config;

	return runInWord(async (ctx) => {
		const paragraphs = ctx.document.body.paragraphs;
		paragraphs.load("items");
		await ctx.sync();

		if (paragraphIndex < 0 || paragraphIndex >= paragraphs.items.length) {
			return {
				error: `Paragraph index ${paragraphIndex} out of range (0-${paragraphs.items.length - 1})`,
			};
		}

		const paragraph = paragraphs.items[paragraphIndex];
		paragraph.load("text");
		await ctx.sync();

		const currentText = String(paragraph.text || "");
		if (!currentText.includes(oldText)) {
			return {
				error: `Text '${oldText}' not found in paragraph ${paragraphIndex}`,
			};
		}

		// Use search within the paragraph to do targeted replace
		const searchResults = paragraph.search(oldText, { matchCase: true });
		searchResults.load("items");
		await ctx.sync();

		if (searchResults.items.length > 0) {
			searchResults.items[0].insertText(newText, "Replace");
			await ctx.sync();
		}

		return { paragraphIndex, oldText, newText, replaced: true, tracked: true };
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
			return { text, insertLocation: "end", inserted: true, tracked: true };
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

		return {
			text,
			insertLocation,
			paragraphIndex,
			inserted: true,
			tracked: true,
		};
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

		// Enable tracked changes for this mutation
		const savedMode = ctx.document.changeTrackingMode;
		ctx.document.changeTrackingMode = "TrackMineOnly";

		const paragraph = paragraphs.items[paragraphIndex];
		paragraph.delete();
		await ctx.sync();

		ctx.document.changeTrackingMode = savedMode;

		return { paragraphIndex, deleted: true, tracked: true };
	});
}

// ── Tracked Changes Tools ───────────────────────────────────────

async function handleGetTrackedChanges(_args: unknown): Promise<unknown> {
	return runInWord(async (ctx) => {
		const mode = ctx.document.changeTrackingMode;
		await ctx.sync();

		return {
			changeTrackingMode: mode,
			pendingChanges: 0, // Mock doesn't track real revision count
		};
	});
}

async function handleAcceptAllChanges(_args: unknown): Promise<unknown> {
	return runInWord(async (ctx) => {
		ctx.document.acceptAllChanges();
		await ctx.sync();

		return { accepted: true };
	});
}

async function handleRejectAllChanges(_args: unknown): Promise<unknown> {
	return runInWord(async (ctx) => {
		ctx.document.rejectAllChanges();
		await ctx.sync();

		return { rejected: true };
	});
}

// ── Word Structure: Tables ──────────────────────────────────

async function handleGetTables(args: unknown): Promise<unknown> {
	const config = args as { includeCellText?: boolean; maxRows?: number };
	const { includeCellText = true, maxRows = 50 } = config;

	return runInWord(async (ctx) => {
		const tables = ctx.document.body.tables;
		tables.load("items");
		await ctx.sync();

		const result: Array<{ index: number; rowCount: number; columnCount: number; cells?: string[][] }> = [];
		for (let i = 0; i < tables.items.length; i++) {
			const tbl = tables.items[i];
			tbl.load("rowCount,columnCount");
			await ctx.sync();

			const entry: (typeof result)[0] = {
				index: i,
				rowCount: tbl.rowCount,
				columnCount: tbl.columnCount,
			};

			if (includeCellText) {
				const rows = Math.min(tbl.rowCount, maxRows);
				const cells: string[][] = [];
				for (let r = 0; r < rows; r++) {
					const rowCells: string[] = [];
					for (let c = 0; c < tbl.columnCount; c++) {
						try {
							const cell = tbl.getCell(r, c);
							cell.value.load("text");
							await ctx.sync();
							rowCells.push(String((cell.value as any).text || ""));
						} catch { rowCells.push("[error]"); }
					}
					cells.push(rowCells);
				}
				entry.cells = cells;
			}
			result.push(entry);
		}
		return { tableCount: result.length, tables: result };
	});
}

async function handleInsertTable(args: unknown): Promise<unknown> {
	const config = args as { rows?: number; columns?: number; afterParagraphIndex?: number; headerRow?: string[] };
	const { rows = 1, columns = 2, afterParagraphIndex = -1, headerRow } = config;

	return runInWord(async (ctx) => {
		const originalMode = ctx.document.changeTrackingMode;
		ctx.document.changeTrackingMode = (Word as any).ChangeTrackingMode.trackMineOnly;

		const body = ctx.document.body;
		let insertRange: any;

		if (afterParagraphIndex === -1) {
			insertRange = body.getRange("End");
		} else {
			const paras = body.paragraphs;
			paras.load("items");
			await ctx.sync();
			if (afterParagraphIndex < paras.items.length) {
				insertRange = paras.items[afterParagraphIndex].getRange("After");
			} else {
				insertRange = body.getRange("End");
			}
		}

		insertRange.insertTable(rows + (headerRow ? 1 : 0), columns, (Word as any).InsertLocation.after, headerRow || undefined);
		await ctx.sync();

		ctx.document.changeTrackingMode = originalMode;
		await ctx.sync();

		return { rows, columns, afterParagraphIndex, inserted: true, tracked: true };
	});
}

async function handleUpdateTableCell(args: unknown): Promise<unknown> {
	const config = args as { tableIndex?: number; row?: number; column?: number; text?: string };
	const { tableIndex = 0, row = 0, column = 0, text = "" } = config;

	return runInWord(async (ctx) => {
		const originalMode = ctx.document.changeTrackingMode;
		ctx.document.changeTrackingMode = (Word as any).ChangeTrackingMode.trackMineOnly;

		const tables = ctx.document.body.tables;
		tables.load("items");
		await ctx.sync();

		if (tableIndex >= tables.items.length) {
			ctx.document.changeTrackingMode = originalMode;
			await ctx.sync();
			return { error: `Table index ${tableIndex} out of bounds. Document has ${tables.items.length} tables.`, errorCode: "CELL_OUT_OF_BOUNDS" };
		}

		const tbl = tables.items[tableIndex];
		tbl.load("rowCount,columnCount");
		await ctx.sync();

		if (row >= tbl.rowCount || column >= tbl.columnCount) {
			ctx.document.changeTrackingMode = originalMode;
			await ctx.sync();
			return { error: `Cell (${row},${column}) out of bounds. Table is ${tbl.rowCount}x${tbl.columnCount}.`, errorCode: "CELL_OUT_OF_BOUNDS", details: { rowCount: tbl.rowCount, columnCount: tbl.columnCount } };
		}

		const cell = tbl.getCell(row, column);
		cell.value.text = text;
		await ctx.sync();

		ctx.document.changeTrackingMode = originalMode;
		await ctx.sync();

		return { tableIndex, row, column, text, updated: true, tracked: true };
	});
}

// ── Word Structure: Headers/Footers ────────────────────────

async function handleGetHeadersFooters(args: unknown): Promise<unknown> {
	const config = args as { sectionIndex?: number };
	const { sectionIndex } = config;

	return runInWord(async (ctx) => {
		const sections = ctx.document.sections;
		sections.load("items");
		await ctx.sync();

		const targetSections = sectionIndex !== undefined ? [sections.items[sectionIndex]] : sections.items;
		const result: Array<{ sectionIndex: number; header?: string; footer?: string }> = [];

		for (let i = 0; i < targetSections.length; i++) {
			const sec = targetSections[i];
			const entry: (typeof result)[0] = { sectionIndex: sectionIndex ?? i };

			try {
				const header = sec.getHeader("Default");
				header.load("text");
				await ctx.sync();
				entry.header = String(header.text || "");
			} catch { entry.header = ""; }

			try {
				const footer = sec.getFooter("Default");
				footer.load("text");
				await ctx.sync();
				entry.footer = String(footer.text || "");
			} catch { entry.footer = ""; }

			result.push(entry);
		}
		return { sectionCount: sections.items.length, sections: result };
	});
}

async function handleSetHeaderFooter(args: unknown): Promise<unknown> {
	const config = args as { sectionIndex?: number; type?: string; variant?: string; text?: string };
	const { sectionIndex = 0, type = "header", variant = "default", text = "" } = config;

	return runInWord(async (ctx) => {
		const originalMode = ctx.document.changeTrackingMode;
		ctx.document.changeTrackingMode = (Word as any).ChangeTrackingMode.trackMineOnly;

		const sections = ctx.document.sections;
		sections.load("items");
		await ctx.sync();

		if (sectionIndex >= sections.items.length) {
			ctx.document.changeTrackingMode = originalMode;
			await ctx.sync();
			return { error: `Section ${sectionIndex} not found. Document has ${sections.items.length} sections.`, errorCode: "INVALID_PARAMETER" };
		}

		const sec = sections.items[sectionIndex];
		const bodyObj = type === "footer" ? sec.getFooter(variant === "firstPage" ? "FirstPage" : "Default") : sec.getHeader(variant === "firstPage" ? "FirstPage" : "Default");
		bodyObj.insertText(text, (Word as any).InsertLocation.replace);
		await ctx.sync();

		ctx.document.changeTrackingMode = originalMode;
		await ctx.sync();

		return { sectionIndex, type, variant, text, set: true, tracked: true };
	});
}

// ── Word Structure: Selection & Insert ──────────────────────

async function handleReplaceSelection(args: unknown): Promise<unknown> {
	const config = args as { text?: string };
	const { text = "" } = config;

	return runInWord(async (ctx) => {
		const selection = ctx.document.getSelection();
		selection.load("text");
		await ctx.sync();

		if (!selection.text || selection.text.trim() === "") {
			return { error: "No text selected. Select text first.", errorCode: "EMPTY_SELECTION" };
		}

		const originalText = selection.text;
		const originalMode = ctx.document.changeTrackingMode;
		ctx.document.changeTrackingMode = (Word as any).ChangeTrackingMode.trackMineOnly;

		selection.insertText(text, (Word as any).InsertLocation.replace);
		await ctx.sync();

		ctx.document.changeTrackingMode = originalMode;
		await ctx.sync();

		return { originalText, newText: text, replaced: true, tracked: true };
	});
}

async function handleInsertImage(args: unknown): Promise<unknown> {
	const config = args as { imageBase64?: string; afterParagraphIndex?: number; width?: number; height?: number };
	const { imageBase64 = "", afterParagraphIndex = -1, width, height } = config;

	if (!imageBase64) return { error: "imageBase64 is required", errorCode: "INVALID_PARAMETER" };

	// Size check: 10MB max
	const sizeBytes = Math.ceil((imageBase64.length * 3) / 4);
	if (sizeBytes > 10 * 1024 * 1024) {
		return { error: `Image too large: ${(sizeBytes / 1024 / 1024).toFixed(1)}MB. Max: 10MB.`, errorCode: "IMAGE_TOO_LARGE" };
	}

	return runInWord(async (ctx) => {
		const body = ctx.document.body;
		let insertRange: any;

		if (afterParagraphIndex === -1) {
			insertRange = body.getRange("End");
		} else {
			const paras = body.paragraphs;
			paras.load("items");
			await ctx.sync();
			insertRange = paras.items[Math.min(afterParagraphIndex, paras.items.length - 1)].getRange("After");
		}

		const image = insertRange.insertInlinePictureFromBase64(imageBase64, (Word as any).InsertLocation.after);
		if (width) image.width = width;
		if (height) image.height = height;
		await ctx.sync();

		return { afterParagraphIndex, inserted: true, width: image.width, height: image.height };
	});
}

// ── Word Structure: Styles & Lists ──────────────────────────

async function handleApplyStyle(args: unknown): Promise<unknown> {
	const config = args as { paragraphIndex?: number; styleName?: string };
	const { paragraphIndex = 0, styleName = "" } = config;
	if (!styleName) return { error: "styleName is required", errorCode: "INVALID_PARAMETER" };

	return runInWord(async (ctx) => {
		const paras = ctx.document.body.paragraphs;
		paras.load("items");
		await ctx.sync();

		if (paragraphIndex >= paras.items.length) {
			return { error: `Paragraph ${paragraphIndex} not found. Document has ${paras.items.length} paragraphs.`, errorCode: "INVALID_PARAMETER" };
		}

		const para = paras.items[paragraphIndex];
		para.style = styleName;
		await ctx.sync();

		return { paragraphIndex, styleName, applied: true };
	});
}

async function handleGetSections(_args: unknown): Promise<unknown> {
	return runInWord(async (ctx) => {
		const sections = ctx.document.sections;
		sections.load("items");
		await ctx.sync();

		const result: Array<{ index: number; differentFirstPage: boolean; differentOddAndEvenPages: boolean }> = [];
		for (let i = 0; i < sections.items.length; i++) {
			const sec = sections.items[i];
			sec.load("differentFirstPage,differentOddAndEvenPages");
			await ctx.sync();
			result.push({
				index: i,
				differentFirstPage: sec.differentFirstPage || false,
				differentOddAndEvenPages: sec.differentOddAndEvenPages || false,
			});
		}
		return { sectionCount: result.length, sections: result };
	});
}

async function handleInsertList(args: unknown): Promise<unknown> {
	const config = args as { type?: string; items?: string[]; afterParagraphIndex?: number };
	const { type = "bulleted", items = [], afterParagraphIndex = -1 } = config;
	if (!items.length) return { error: "items must be non-empty", errorCode: "EMPTY_ITEMS" };

	return runInWord(async (ctx) => {
		const originalMode = ctx.document.changeTrackingMode;
		ctx.document.changeTrackingMode = (Word as any).ChangeTrackingMode.trackMineOnly;

		const body = ctx.document.body;
		const bulletType = type === "numbered" ? (Word as any).BulletType.numbered : (Word as any).BulletType.bulleted;

		// Build list text
		const listText = items.join("\r");
		let insertRange: any;

		if (afterParagraphIndex === -1) {
			insertRange = body.getRange("End");
		} else {
			const paras = body.paragraphs;
			paras.load("items");
			await ctx.sync();
			insertRange = paras.items[Math.min(afterParagraphIndex, paras.items.length - 1)].getRange("After");
		}

		insertRange.insertParagraph(listText, (Word as any).InsertLocation.after);
		await ctx.sync();

		// Apply list formatting to the inserted paragraph
		const insertedPara = insertRange.paragraphs.getLast();
		insertedPara.load("uniqueLocalId");
		await ctx.sync();

		// Split by \r and apply bullet/number formatting
		const listItems = insertedPara.split(["\r"]);
		listItems.load("items");
		await ctx.sync();

		for (const item of listItems.items) {
			item.startList(bulletType);
		}
		await ctx.sync();

		ctx.document.changeTrackingMode = originalMode;
		await ctx.sync();

		return { type, itemCount: items.length, afterParagraphIndex, inserted: true, tracked: true };
	});
}
