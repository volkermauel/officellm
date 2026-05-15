/**
 * PowerPoint command handler using Office JS API.
 *
 * Critical API rules:
 * - shape.textFrame THROWS InvalidArgument → use getTextFrameOrNullObject()
 * - getTextFrameOrNullObject() returns object with isNullObject=true if no text
 * - shape.load("id") loads ONE property, shape.load("id,name") loads multiple
 * - slide.load("shapes/items/$none") loads collection items without properties
 * - slash paths like "textFrame/textRange/text" load nested props
 * - MUST sync() before reading ANY loaded property
 * - load() errors fire at sync() time, not at load() time
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

async function handleGetDeckOutline(_args: unknown): Promise<unknown> {
	// eslint-disable-next-line @typescript-eslint/no-explicit-any
	const PowerPoint: any = (window as any).PowerPoint;
	if (!PowerPoint || typeof PowerPoint.run !== "function") {
		return { documentName: "(unknown)", totalSlides: 0, slides: [], error: "PowerPoint.run() not available" };
	}

	return new Promise((resolve) => {
		PowerPoint.run(async (context: any) => {
			try {
				const presentation = context.presentation;
				presentation.load("slides");
				await context.sync();

				const slides = presentation.slides;
				const totalSlides = slides.items.length;
				const slideList: Array<{ index: number; title: string }> = [];

				// Load shapes items for all slides
				for (const slide of slides.items) {
					slide.load("shapes/items/$none");
				}
				await context.sync();

				// For each shape, get text via getTextFrameOrNullObject
				for (const slide of slides.items) {
					for (const shape of slide.shapes.items) {
						const tf = shape.getTextFrameOrNullObject();
						tf.load("textRange/text");
					}
				}
				await context.sync();

				// Extract titles
				for (let i = 0; i < totalSlides; i++) {
					const shapes = slides.items[i].shapes.items;
					let title = "";

					for (const shape of shapes) {
						try {
							const tf = shape.getTextFrameOrNullObject();
							if (!tf.isNullObject && tf.textRange?.text) {
								const text = String(tf.textRange.text).trim();
								if (text && !title) {
									title = text;
								}
							}
						} catch {
							// skip
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
		PowerPoint.run(async (context: any) => {
			try {
				const presentation = context.presentation;
				presentation.load("slides");
				await context.sync();

				if (slideIndex < 0 || slideIndex >= presentation.slides.items.length) {
					resolve({ error: `Slide index ${slideIndex} out of range (0-${presentation.slides.items.length - 1})` });
					return;
				}

				const slide = presentation.slides.items[slideIndex];

				// Step 1: Load shapes items
				slide.load("shapes/items/$none");
				await context.sync();

				// Step 2: Load id + name on each shape, and text via getTextFrameOrNullObject
				for (const shape of slide.shapes.items) {
					shape.load("id");
					shape.load("name");
					const tf = shape.getTextFrameOrNullObject();
					tf.load("textRange/text");
				}
				await context.sync();

				// Step 3: Build result
				const shapeList: Array<{ id: string; name: string; shapeType: string; text: string }> = [];
				let slideTitle = "";

				for (const shape of slide.shapes.items) {
					let text = "";
					try {
						const tf = shape.getTextFrameOrNullObject();
						if (!tf.isNullObject && tf.textRange?.text) {
							text = String(tf.textRange.text);
						}
					} catch {
						// no text
					}

					shapeList.push({
						id: String(shape.id || ""),
						name: String(shape.name || ""),
						shapeType: "unknown",
						text,
					});

					if (!slideTitle && text.trim()) {
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
