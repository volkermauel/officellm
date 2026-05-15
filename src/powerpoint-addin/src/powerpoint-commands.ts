/**
 * PowerPoint command handler using Office JS API.
 *
 * Critical API rules (verified by testing):
 * - getTextFrameOrNullObject() creates a NEW object each call — MUST store references
 * - ctx.load(tf, "isNullObject,textRange/text") — load on the context, not tf.load()
 * - isNullObject must be explicitly loaded — not available by default
 * - slide.load("shapes/items/$none") loads collection items without properties
 * - shape.load("id,name") — comma-separated works on direct properties
 * - MUST sync() before reading ANY loaded property
 */

/// <reference types="@types/office-js" />

import { reportResult } from "./communication";

export async function processCommand(
	commandId: string,
	commandName: string,
	args: unknown,
): Promise<unknown> {
	let result: unknown;

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

		await reportResult(commandId, true, undefined, result);
	} catch (error) {
		const errorMessage = error instanceof Error ? error.message : String(error);
		console.error(`Command ${commandId} failed:`, errorMessage);
		await reportResult(commandId, false, errorMessage);
		result = { error: errorMessage };
	}

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
	return !TITLE_SKIP_PATTERNS.some(p => p.test(name));
}

async function handleGetDeckOutline(_args: unknown): Promise<unknown> {
	// eslint-disable-next-line @typescript-eslint/no-explicit-any
	const PowerPoint: any = (window as any).PowerPoint;
	if (!PowerPoint || typeof PowerPoint.run !== "function") {
		return { documentName: "(unknown)", totalSlides: 0, slides: [], error: "PowerPoint.run() not available" };
	}

	return new Promise((resolve) => {
		PowerPoint.run(async (ctx: any) => {
			try {
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

				resolve({ documentName: "Presentation", totalSlides, slides: slideList });
			} catch (error) {
				resolve({ error: `PowerPoint.run failed: ${error instanceof Error ? error.message : String(error)}` });
			}
		});
	});
}

async function handleGetSlide(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number };
	const slideIndex = config.slideIndex ?? 0;

	// eslint-disable-next-line @typescript-eslint/no-explicit-any
	const PowerPoint: any = (window as any).PowerPoint;
	if (!PowerPoint || typeof PowerPoint.run !== "function") {
		return { slideIndex, title: "(not available)", shapes: [], speakerNotes: "", isHidden: false, error: "PowerPoint.run() not available" };
	}

	return new Promise((resolve) => {
		PowerPoint.run(async (ctx: any) => {
			try {
				const pres = ctx.presentation;
				pres.load("slides");
				await ctx.sync();

				if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
					resolve({ error: `Slide index ${slideIndex} out of range (0-${pres.slides.items.length - 1})` });
					return;
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

				resolve({
					slideIndex,
					title: slideTitle || `Slide ${slideIndex + 1}`,
					shapes: shapeList,
					speakerNotes: "",
					isHidden: false,
				});
			} catch (error) {
				resolve({ error: `PowerPoint.run failed: ${error instanceof Error ? error.message : String(error)}` });
			}
		});
	});
}

async function handleUpdateShapeText(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; shapeId?: string; text?: string; confirmationToken?: string };
	const { slideIndex = 0, shapeId = "", text = "", confirmationToken: _ct } = config;
	return { slideIndex, shapeId, newText: text, message: "Shape text update simulated (not yet implemented)" };
}

async function handleUpdateSpeakerNotes(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; notes?: string; confirmationToken?: string };
	const { slideIndex = 0, notes = "", confirmationToken: _ct } = config;
	return { slideIndex, newNotes: notes, message: "Speaker notes update simulated (not yet implemented)" };
}
