/**
 * PowerPoint command handler using Office JS API.
 *
 * Critical API rules (verified by testing):
 * - getTextFrameOrNullObject() creates a NEW object each call — MUST store references
 * - ctx.load(tf, "isNullObject,textRange/text") — load on the context
 * - isNullObject must be explicitly loaded — not available by default
 * - slide.load("shapes/items/$none") loads collection items without properties
 * - shape.load("id,name,type,left,top,width,height,rotation") — comma-separated for direct props
 * - MUST sync() before reading ANY loaded property
 * - PowerPoint.run() rejects on failure — .catch() on the returned promise
 * - shape.getImageAsBase64() and slide.getImageAsBase64() return ClientResult<string> — sync() first
 * - shape.fill and textRange.font are nested objects — load via slash paths
 */

/// <reference types="@types/office-js" />

import { reportResult } from "./communication";

// ── Command dispatch ────────────────────────────────────────────

export async function processCommand(
	commandId: string,
	commandName: string,
	args: unknown,
): Promise<unknown> {
	let result: unknown;
	let success = true;

	try {
		switch (commandName) {
			// Read tools
			case "powerpoint_get_deck_outline":
				result = await handleGetDeckOutline(args);
				break;
			case "powerpoint_get_slide":
				result = await handleGetSlide(args);
				break;
			case "powerpoint_get_slide_image":
				result = await handleGetSlideImage(args);
				break;
			case "powerpoint_get_shape_image":
				result = await handleGetShapeImage(args);
				break;
			case "powerpoint_get_table":
				result = await handleGetTable(args);
				break;
			case "powerpoint_get_selection":
				result = await handleGetSelection(args);
				break;
			case "powerpoint_get_speaker_notes":
				result = await handleGetSpeakerNotes(args);
				break;

			// Write tools
			case "powerpoint_update_shape_text":
				result = await handleUpdateShapeText(args);
				break;
			case "powerpoint_update_shape_properties":
				result = await handleUpdateShapeProperties(args);
				break;
			case "powerpoint_update_speaker_notes":
				result = await handleUpdateSpeakerNotes(args);
				break;

			// Shape CRUD
			case "powerpoint_add_textbox":
				result = await handleAddTextbox(args);
				break;
			case "powerpoint_add_image":
				result = await handleAddImage(args);
				break;
			case "powerpoint_add_table":
				result = await handleAddTable(args);
				break;
			case "powerpoint_delete_shape":
				result = await handleDeleteShape(args);
				break;

			// Slide management
			case "powerpoint_add_slide":
				result = await handleAddSlide(args);
				break;
			case "powerpoint_delete_slide":
				result = await handleDeleteSlide(args);
				break;
			case "powerpoint_move_slide":
				result = await handleMoveSlide(args);
				break;

			default:
				result = { error: `Unknown command: ${commandName}` };
		}

		if (result && typeof result === "object" && "error" in result) {
			success = false;
		}
	} catch (error) {
		const errorMessage = error instanceof Error ? error.message : String(error);
		console.error(`Command ${commandId} failed:`, errorMessage);
		success = false;
		result = { error: errorMessage };
	}

	await reportResult(commandId, success, undefined, result);
	return result;
}

// ── Helpers ─────────────────────────────────────────────────────

const TITLE_SKIP_PATTERNS = [
	/slide\s*number/i,
	/footer/i,
	/header/i,
	/date/i,
	/background/i,
];

function isContentShape(name: string): boolean {
	return !TITLE_SKIP_PATTERNS.some((p) => p.test(name));
}

/**
 * Wraps PowerPoint.run() in a Promise that properly rejects on errors.
 */
function runInPowerPoint<T>(fn: (ctx: any) => Promise<T>): Promise<T> {
	return new Promise<T>((resolve, reject) => {
		const PowerPoint: any = (window as any).PowerPoint;
		if (!PowerPoint || typeof PowerPoint.run !== "function") {
			reject(new Error("PowerPoint.run() not available"));
			return;
		}

		PowerPoint.run(async (ctx: any) => {
			resolve(await fn(ctx));
		}).catch(reject);
	});
}

function safeStr(val: any, fallback = ""): string {
	return val != null ? String(val) : fallback;
}

function safeNum(val: any, fallback = 0): number {
	return typeof val === "number" ? val : fallback;
}

// ── Read tools ──────────────────────────────────────────────────

async function handleGetDeckOutline(args: unknown): Promise<unknown> {
	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		const totalSlides = pres.slides.items.length;

		// Load shape items
		for (const sl of pres.slides.items) {
			sl.load("shapes/items/$none");
		}
		await ctx.sync();

		// Load direct shape properties: id, name, type, left, top, width, height
		for (const sl of pres.slides.items) {
			for (const s of sl.shapes.items) {
				s.load("id,name,type,left,top,width,height");
			}
		}
		await ctx.sync();

		// Build slide list
		const slideList = [];
		for (let i = 0; i < totalSlides; i++) {
			const shapes = pres.slides.items[i].shapes.items;

			// Load text for title detection
			const tfs: any[] = [];
			for (const s of shapes) {
				const tf = s.getTextFrameOrNullObject();
				ctx.load(tf, "isNullObject,textRange/text");
				tfs.push(tf);
			}
			await ctx.sync();

			let title = "";
			const shapeData = [];
			for (let j = 0; j < shapes.length; j++) {
				const s = shapes[j];
				const tf = tfs[j];
				let text = "";
				if (!tf.isNullObject && tf.textRange?.text) {
					text = String(tf.textRange.text).trim();
				}

				const shapeName = safeStr(s.name);
				if (!title && text && isContentShape(shapeName)) {
					title = text;
				}

				shapeData.push({
					id: safeStr(s.id),
					name: shapeName,
					type: safeStr(s.type),
					left: safeNum(s.left),
					top: safeNum(s.top),
					width: safeNum(s.width),
					height: safeNum(s.height),
					text,
				});
			}

			slideList.push({
				index: i,
				title: title || `Slide ${i + 1}`,
				shapes: shapeData,
			});
		}

		return { documentName: "Presentation", totalSlides, slides: slideList };
	});
}

async function handleGetSlide(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number };
	const slideIndex = config.slideIndex ?? 0;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range (0-${pres.slides.items.length - 1})` };
		}

		const slide = pres.slides.items[slideIndex];

		// Step 1: Load shape items
		slide.load("shapes/items/$none");
		await ctx.sync();

		// Step 2: Load direct properties
		for (const s of slide.shapes.items) {
			s.load("id,name,type,left,top,width,height,rotation");
		}
		await ctx.sync();

		// Step 3: Load text frames + font + fill
		const textFrameData: any[] = [];
		const fillData: any[] = [];
		for (const s of slide.shapes.items) {
			// Text frame
			const tf = s.getTextFrameOrNullObject();
			ctx.load(tf, "isNullObject,textRange/text,textRange/font/name,textRange/font/size,textRange/font/bold,textRange/font/italic,textRange/font/color");
			textFrameData.push(tf);

			// Fill
			ctx.load(s.fill, "foregroundColor,transparency");
			fillData.push(s.fill);
		}
		await ctx.sync();

		// Step 4: Build result
		let slideTitle = "";
		const shapeList = [];

		for (let i = 0; i < slide.shapes.items.length; i++) {
			const s = slide.shapes.items[i];
			const tf = textFrameData[i];
			const fill = fillData[i];

			let text = "";
			let font: Record<string, unknown> | undefined;
			if (!tf.isNullObject) {
				text = safeStr(tf.textRange?.text);
				if (tf.textRange?.font) {
					font = {
						name: safeStr(tf.textRange.font.name),
						size: safeNum(tf.textRange.font.size),
						bold: !!tf.textRange.font.bold,
						italic: !!tf.textRange.font.italic,
						color: safeStr(tf.textRange.font.color),
					};
				}
			}

			const shapeName = safeStr(s.name);
			if (!slideTitle && text.trim() && isContentShape(shapeName)) {
				slideTitle = text.trim();
			}

			shapeList.push({
				id: safeStr(s.id),
				name: shapeName,
				type: safeStr(s.type),
				left: safeNum(s.left),
				top: safeNum(s.top),
				width: safeNum(s.width),
				height: safeNum(s.height),
				rotation: safeNum(s.rotation),
				text,
				font,
				fillColor: safeStr(fill.foregroundColor),
				fillTransparency: safeNum(fill.transparency),
			});
		}

		return {
			slideIndex,
			title: slideTitle || `Slide ${slideIndex + 1}`,
			shapes: shapeList,
		};
	});
}

async function handleGetSlideImage(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; width?: number; height?: number };
	const slideIndex = config.slideIndex ?? 0;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];
		const imageOptions: any = {};
		if (config.width) imageOptions.width = config.width;
		if (config.height) imageOptions.height = config.height;

		const imageResult = slide.getImageAsBase64(imageOptions);
		await ctx.sync();

		return {
			slideIndex,
			image: `data:image/png;base64,${imageResult.value}`,
		};
	});
}

async function handleGetShapeImage(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; shapeId?: string; width?: number; height?: number };
	const slideIndex = config.slideIndex ?? 0;
	const shapeId = config.shapeId ?? "";

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];
		slide.load("shapes/items/$none");
		await ctx.sync();

		// Find shape by ID or name
		let targetShape: any = null;
		for (const s of slide.shapes.items) {
			s.load("id,name");
		}
		await ctx.sync();

		for (const s of slide.shapes.items) {
			if (safeStr(s.id) === shapeId || safeStr(s.name) === shapeId) {
				targetShape = s;
				break;
			}
		}

		if (!targetShape) {
			return { error: `Shape '${shapeId}' not found on slide ${slideIndex}` };
		}

		const imageOptions: any = {};
		if (config.width) imageOptions.width = config.width;
		if (config.height) imageOptions.height = config.height;

		const imageResult = targetShape.getImageAsBase64(imageOptions);
		await ctx.sync();

		return {
			slideIndex,
			shapeId,
			image: `data:image/png;base64,${imageResult.value}`,
		};
	});
}

async function handleGetTable(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; shapeId?: string };
	const slideIndex = config.slideIndex ?? 0;
	const shapeId = config.shapeId ?? "";

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];
		slide.load("shapes/items/$none");
		await ctx.sync();

		// Find the shape
		let targetShape: any = null;
		for (const s of slide.shapes.items) {
			s.load("id,name,type");
		}
		await ctx.sync();

		for (const s of slide.shapes.items) {
			if (safeStr(s.id) === shapeId || safeStr(s.name) === shapeId) {
				targetShape = s;
				break;
			}
		}

		if (!targetShape) {
			return { error: `Shape '${shapeId}' not found on slide ${slideIndex}` };
		}

		// Get the table
		const table = targetShape.getTable();
		table.load("rowCount,columnCount");
		await ctx.sync();

		const rows = table.rowCount;
		const cols = table.columnCount;

		// Read all cells
		const cells: string[][] = [];
		for (let r = 0; r < rows; r++) {
			const row: string[] = [];
			for (let c = 0; c < cols; c++) {
				const cell = table.getCell(r, c);
				const tf = cell.textFrame;
				ctx.load(tf, "textRange/text");
				row.push(""); // placeholder
				// Store for later reading
				(cell as any).__row = r;
				(cell as any).__col = c;
			}
			cells.push(row);
		}

		// Alternative approach: load cell texts individually
		const cellTexts: any[] = [];
		for (let r = 0; r < rows; r++) {
			for (let c = 0; c < cols; c++) {
				const cell = table.getCell(r, c);
				const tf = cell.textFrame;
				ctx.load(tf, "textRange/text");
				cellTexts.push(tf);
			}
		}
		await ctx.sync();

		// Fill in the cells array
		let idx = 0;
		for (let r = 0; r < rows; r++) {
			for (let c = 0; c < cols; c++) {
				const tf = cellTexts[idx++];
				cells[r][c] = safeStr(tf.textRange?.text);
			}
		}

		return {
			slideIndex,
			shapeId,
			rowCount: rows,
			columnCount: cols,
			cells,
		};
	});
}

async function handleGetSelection(_args: unknown): Promise<unknown> {
	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;

		// Try to get selected text range
		try {
			const textRange = pres.getSelectedTextRange();
			textRange.load("text,font/name,font/size,font/bold,font/italic,font/color");

			const parentTf = textRange.getParentTextFrame();
			const parentShape = parentTf.getParentShape();
			const parentSlide = parentShape.getParentSlide();
			parentShape.load("id,name");
			parentSlide.load("id");
			await ctx.sync();

			return {
				type: "text",
				text: safeStr(textRange.text),
				font: {
					name: safeStr(textRange.font.name),
					size: safeNum(textRange.font.size),
					bold: !!textRange.font.bold,
					italic: !!textRange.font.italic,
					color: safeStr(textRange.font.color),
				},
				shapeId: safeStr(parentShape.id),
				shapeName: safeStr(parentShape.name),
				slideId: safeStr(parentSlide.id),
			};
		} catch {
			// No text selected — try shapes
		}

		// Try to get selected shapes
		try {
			const selectedShapes = pres.getSelectedShapes();
			selectedShapes.load("items");
			await ctx.sync();

			if (selectedShapes.items.length > 0) {
				for (const s of selectedShapes.items) {
					s.load("id,name,type");
				}
				await ctx.sync();

				return {
					type: "shapes",
					shapes: selectedShapes.items.map((s: any) => ({
						id: safeStr(s.id),
						name: safeStr(s.name),
						type: safeStr(s.type),
					})),
				};
			}
		} catch {
			// No shapes selected
		}

		return { type: "none" };
	});
}

async function handleGetSpeakerNotes(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; slideRange?: string };
	const slideIndex = config.slideIndex;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		// Determine which slides to read
		let indices: number[] = [];
		if (slideIndex != null) {
			indices = [slideIndex];
		} else if (config.slideRange) {
			const match = config.slideRange.match(/^(\d+)-(\d+)$/);
			if (match) {
				const from = parseInt(match[1]);
				const to = parseInt(match[2]);
				for (let i = from; i <= to && i < pres.slides.items.length; i++) {
					indices.push(i);
				}
			} else {
				return { error: `Invalid slideRange format: '${config.slideRange}'. Use '2-5'.` };
			}
		} else {
			// Default: all slides
			for (let i = 0; i < pres.slides.items.length; i++) {
				indices.push(i);
			}
		}

		// Read notes for each slide
		const notes: Array<{ slideIndex: number; notes: string }> = [];
		const notesTfs: any[] = [];

		for (const idx of indices) {
			if (idx < 0 || idx >= pres.slides.items.length) {
				notes.push({ slideIndex: idx, notes: "" });
				notesTfs.push(null);
				continue;
			}

			const slide = pres.slides.items[idx];
			try {
				const notesSlide = slide.getNotesSlideOrNullObject();
				const tf = notesSlide.textFrame;
				ctx.load(tf, "isNullObject,textRange/text");
				notesTfs.push(tf);
			} catch {
				notesTfs.push(null);
			}
		}
		await ctx.sync();

		for (let i = 0; i < indices.length; i++) {
			const tf = notesTfs[i];
			let noteText = "";
			if (tf && !tf.isNullObject && tf.textRange?.text) {
				noteText = String(tf.textRange.text).trim();
			}
			notes.push({ slideIndex: indices[i], notes: noteText });
		}

		return { notes };
	});
}

// ── Write tools ─────────────────────────────────────────────────

async function handleUpdateShapeText(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; shapeId?: string; text?: string };
	const { slideIndex = 0, shapeId = "", text = "" } = config;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];
		slide.load("shapes/items/$none");
		await ctx.sync();

		for (const s of slide.shapes.items) {
			s.load("id,name");
		}
		await ctx.sync();

		// Find shape
		const shape = slide.shapes.items.find(
			(s: any) => safeStr(s.id) === shapeId || safeStr(s.name) === shapeId,
		);
		if (!shape) {
			return { error: `Shape '${shapeId}' not found on slide ${slideIndex}` };
		}

		// Get text frame — check if it's a text-bearing shape
		const tf = shape.getTextFrameOrNullObject();
		ctx.load(tf, "isNullObject");
		await ctx.sync();

		if (tf.isNullObject) {
			return { error: `Shape '${shapeId}' does not support text (type: image/table/etc)` };
		}

		// Set the text
		tf.textRange.text = text;
		await ctx.sync();

		return { slideIndex, shapeId, newText: text };
	});
}

async function handleUpdateShapeProperties(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex?: number;
		shapeId?: string;
		left?: number;
		top?: number;
		width?: number;
		height?: number;
		rotation?: number;
		fontName?: string;
		fontSize?: number;
		bold?: boolean;
		italic?: boolean;
		color?: string;
	};
	const { slideIndex = 0, shapeId = "" } = config;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];
		slide.load("shapes/items/$none");
		await ctx.sync();

		for (const s of slide.shapes.items) {
			s.load("id,name");
		}
		await ctx.sync();

		const shape = slide.shapes.items.find(
			(s: any) => safeStr(s.id) === shapeId || safeStr(s.name) === shapeId,
		);
		if (!shape) {
			return { error: `Shape '${shapeId}' not found on slide ${slideIndex}` };
		}

		const updated: string[] = [];

		// Position and size
		if (config.left !== undefined) { shape.left = config.left; updated.push("left"); }
		if (config.top !== undefined) { shape.top = config.top; updated.push("top"); }
		if (config.width !== undefined) { shape.width = config.width; updated.push("width"); }
		if (config.height !== undefined) { shape.height = config.height; updated.push("height"); }
		if (config.rotation !== undefined) { shape.rotation = config.rotation; updated.push("rotation"); }

		// Font properties
		if (config.fontName !== undefined || config.fontSize !== undefined ||
			config.bold !== undefined || config.italic !== undefined || config.color !== undefined) {
			const tf = shape.getTextFrameOrNullObject();
			ctx.load(tf, "isNullObject");
			await ctx.sync();

			if (!tf.isNullObject) {
				const font = tf.textRange.font;
				if (config.fontName !== undefined) { font.name = config.fontName; updated.push("fontName"); }
				if (config.fontSize !== undefined) { font.size = config.fontSize; updated.push("fontSize"); }
				if (config.bold !== undefined) { font.bold = config.bold; updated.push("bold"); }
				if (config.italic !== undefined) { font.italic = config.italic; updated.push("italic"); }
				if (config.color !== undefined) { font.color = config.color; updated.push("color"); }
			}
		}

		await ctx.sync();

		return { slideIndex, shapeId, updated };
	});
}

async function handleUpdateSpeakerNotes(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; notes?: string };
	const { slideIndex = 0, notes = "" } = config;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];
		const notesSlide = slide.getNotesSlide();
		notesSlide.load("textFrame");
		await ctx.sync();

		notesSlide.textFrame.textRange.text = notes;
		await ctx.sync();

		return { slideIndex, newNotes: notes };
	});
}

// ── Shape CRUD ──────────────────────────────────────────────────

async function handleAddTextbox(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex?: number;
		text?: string;
		left?: number;
		top?: number;
		width?: number;
		height?: number;
	};
	const { slideIndex = 0, text = "", left = 100, top = 100, width = 300, height = 100 } = config;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];
		const textbox = slide.shapes.addTextBox(text, { left, top, width, height });
		textbox.load("id,name");
		await ctx.sync();

		return {
			slideIndex,
			shapeId: safeStr(textbox.id),
			name: safeStr(textbox.name),
		};
	});
}

async function handleAddImage(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex?: number;
		imageBase64?: string;
		left?: number;
		top?: number;
		width?: number;
		height?: number;
	};
	const { slideIndex = 0, imageBase64 = "", left = 100, top = 100 } = config;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];

		// Strip data URI prefix if present
		const base64Data = imageBase64.replace(/^data:image\/[^;]+;base64,/, "");

		const options: any = { left, top };
		if (config.width !== undefined) options.width = config.width;
		if (config.height !== undefined) options.height = config.height;

		const picture = slide.shapes.addPicture(base64Data, options);
		picture.load("id,name");
		await ctx.sync();

		return {
			slideIndex,
			shapeId: safeStr(picture.id),
			name: safeStr(picture.name),
		};
	});
}

async function handleAddTable(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex?: number;
		rows?: number;
		columns?: number;
		left?: number;
		top?: number;
		width?: number;
		height?: number;
	};
	const { slideIndex = 0, rows = 2, columns = 2, left = 100, top = 100 } = config;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];

		const options: any = { left, top };
		if (config.width !== undefined) options.width = config.width;
		if (config.height !== undefined) options.height = config.height;

		const table = slide.shapes.addTable(rows, columns, options);
		table.load("id,name");
		await ctx.sync();

		return {
			slideIndex,
			shapeId: safeStr(table.id),
			name: safeStr(table.name),
		};
	});
}

async function handleDeleteShape(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; shapeId?: string };
	const { slideIndex = 0, shapeId = "" } = config;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];
		slide.load("shapes/items/$none");
		await ctx.sync();

		for (const s of slide.shapes.items) {
			s.load("id,name");
		}
		await ctx.sync();

		const shape = slide.shapes.items.find(
			(s: any) => safeStr(s.id) === shapeId || safeStr(s.name) === shapeId,
		);
		if (!shape) {
			return { error: `Shape '${shapeId}' not found on slide ${slideIndex}` };
		}

		shape.delete();
		await ctx.sync();

		return { slideIndex, shapeId, deleted: true };
	});
}

// ── Slide management ────────────────────────────────────────────

async function handleAddSlide(args: unknown): Promise<unknown> {
	const config = args as { atIndex?: number };

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		const options: any = {};
		if (config.atIndex !== undefined) {
			options.index = config.atIndex;
		}

		pres.slides.add(options);

		// Reload to get the new slide
		pres.load("slides");
		await ctx.sync();

		// The new slide is at the specified index (or end)
		const newIndex = config.atIndex !== undefined
			? Math.min(config.atIndex, pres.slides.items.length - 1)
			: pres.slides.items.length - 1;

		const newSlide = pres.slides.items[newIndex];
		newSlide.load("id");
		await ctx.sync();

		return {
			slideIndex: newIndex,
			slideId: safeStr(newSlide.id),
		};
	});
}

async function handleDeleteSlide(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number };
	const { slideIndex = 0 } = config;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];
		slide.delete();
		await ctx.sync();

		return { slideIndex, deleted: true };
	});
}

async function handleMoveSlide(args: unknown): Promise<unknown> {
	const config = args as { fromIndex?: number; toIndex?: number };
	const { fromIndex = 0, toIndex = 0 } = config;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (fromIndex < 0 || fromIndex >= pres.slides.items.length) {
			return { error: `fromIndex ${fromIndex} out of range` };
		}

		const slide = pres.slides.items[fromIndex];
		slide.load("id");
		await ctx.sync();

		const slideId = safeStr(slide.id);
		slide.moveTo(toIndex);
		await ctx.sync();

		return { fromIndex, toIndex, slideId };
	});
}
