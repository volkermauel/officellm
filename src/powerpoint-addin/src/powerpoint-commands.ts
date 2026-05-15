/**
 * PowerPoint command handler using Office JS API.
 *
 * Critical API rules (verified by testing):
 * - getTextFrameOrNullObject() creates a NEW object each call — MUST store references
 * - ctx.load(tf, "isNullObject,textRange/text") — load on the context
 * - isNullObject must be explicitly loaded — not available by default
 * - slide.load("shapes/items/$none") loads collection items without properties
 * - shape.load("id,name") — comma-separated works on direct properties
 * - MUST sync() before reading ANY loaded property
 * - PowerPoint.run() rejects on failure — use .catch() on the returned promise
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
			case "powerpoint_get_deck_outline":
				result = await handleGetDeckOutline(args);
				break;
			case "powerpoint_get_slide":
				result = await handleGetSlide(args);
				break;
			case "powerpoint_update_shape_text":
				result = await handleUpdateShapeText(args);
				break;
			case "powerpoint_update_speaker_notes":
				result = await handleUpdateSpeakerNotes(args);
				break;
			default:
				result = { error: `Unknown command: ${commandName}` };
		}

		// If result contains an error field, report as failure
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

// Shapes to skip when looking for slide titles (non-content shapes)
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
 * PowerPoint.run() returns its own promise — we attach .catch() so that
 * both callback-thrown errors AND run-level errors propagate as rejections.
 */
function runInPowerPoint<T>(fn: (ctx: any) => Promise<T>): Promise<T> {
	return new Promise<T>((resolve, reject) => {
		// eslint-disable-next-line @typescript-eslint/no-explicit-any
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

async function handleGetDeckOutline(_args: unknown): Promise<unknown> {
	return runInPowerPoint(async (ctx: any) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		const slides = pres.slides;
		const totalSlides = slides.items.length;

		// Load shapes items for all slides
		for (const sl of slides.items) {
			sl.load("shapes/items/$none");
		}
		await ctx.sync();

		// Load textFrame references (MUST store — getTextFrameOrNullObject creates new objects)
		const slideTextFrames: any[][] = [];
		for (let i = 0; i < totalSlides; i++) {
			const sl = slides.items[i];
			const tfs: any[] = [];
			for (const s of sl.shapes.items) {
				const tf = s.getTextFrameOrNullObject();
				ctx.load(tf, "isNullObject,textRange/text");
				tfs.push(tf);
			}
			slideTextFrames.push(tfs);
		}
		await ctx.sync();

		// Extract titles (first content shape with non-empty text)
		const slideList: Array<{ index: number; title: string }> = [];
		for (let i = 0; i < totalSlides; i++) {
			const shapes = slides.items[i].shapes.items;
			const tfs = slideTextFrames[i];
			let title = "";

			for (let j = 0; j < shapes.length; j++) {
				const tf = tfs[j];
				if (tf.isNullObject) continue;
				const shapeName = String(shapes[j].name || "");
				if (!isContentShape(shapeName)) continue;

				const text = String(tf.textRange?.text || "").trim();
				if (text && !title) {
					title = text;
				}
			}

			slideList.push({ index: i, title: title || `Slide ${i + 1}` });
		}

		return { documentName: "Presentation", totalSlides, slides: slideList };
	});
}

async function handleGetSlide(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number };
	const slideIndex = config.slideIndex ?? 0;

	return runInPowerPoint(async (ctx: any) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range (0-${pres.slides.items.length - 1})` };
		}

		const slide = pres.slides.items[slideIndex];

		// Step 1: Load shapes items + id/name
		slide.load("shapes/items/$none");
		await ctx.sync();

		for (const s of slide.shapes.items) {
			s.load("id,name");
		}
		await ctx.sync();

		// Step 2: Load textFrame references
		const tfs: any[] = [];
		for (const s of slide.shapes.items) {
			const tf = s.getTextFrameOrNullObject();
			ctx.load(tf, "isNullObject,textRange/text");
			tfs.push(tf);
		}
		await ctx.sync();

		// Step 3: Build result
		const shapeList: Array<{ id: string; name: string; shapeType: string; text: string }> = [];
		let slideTitle = "";

		for (let i = 0; i < slide.shapes.items.length; i++) {
			const s = slide.shapes.items[i];
			const tf = tfs[i];
			let text = "";

			if (!tf.isNullObject && tf.textRange?.text) {
				text = String(tf.textRange.text);
			}

			shapeList.push({
				id: String(s.id || ""),
				name: String(s.name || ""),
				shapeType: "unknown",
				text,
			});

			if (!slideTitle && text.trim() && isContentShape(String(s.name || ""))) {
				slideTitle = text.trim();
			}
		}

		return {
			slideIndex,
			title: slideTitle || `Slide ${slideIndex + 1}`,
			shapes: shapeList,
			speakerNotes: "",
			isHidden: false,
		};
	});
}

async function handleUpdateShapeText(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; shapeId?: string; text?: string; confirmationToken?: string };
	const { slideIndex = 0, shapeId = "", text = "" } = config;
	return { slideIndex, shapeId, newText: text, message: "Shape text update simulated (not yet implemented)" };
}

async function handleUpdateSpeakerNotes(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; notes?: string; confirmationToken?: string };
	const { slideIndex = 0, notes = "" } = config;
	return { slideIndex, newNotes: notes, message: "Speaker notes update simulated (not yet implemented)" };
}
