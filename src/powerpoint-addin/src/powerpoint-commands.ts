/**
 * PowerPoint command handler using Office JS API.
 *
 * Correct API patterns (from official docs + snippets):
 * - presentation.load("slides") → sync → slides.items available
 * - slide.load("shapes/items/$none") → sync → shapes.items available (no props)
 * - shape.load("textFrame/textRange/text") → sync → text available
 * - shape.load("id") → sync → shape.id available
 * - Slash paths support ONE property at a time (not comma-separated)
 * - shape.textFrame THROWS if shape has no text frame → wrap in try/catch
 * - presentation.name requires separate load — just skip it
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
		return {
			documentName: "(unknown)",
			totalSlides: 0,
			slides: [],
			error: "PowerPoint.run() not available",
		};
	}

	return new Promise((resolve) => {
		PowerPoint.run(async (context: any) => {
			try {
				const presentation = context.presentation;

				// Step 1: Load slides collection
				presentation.load("slides");
				await context.sync();

				const slides = presentation.slides;
				const totalSlides = slides.items.length;
				const slideList: Array<{ index: number; title: string }> = [];

				// Step 2: Load shapes items for all slides (no properties yet)
				for (const slide of slides.items) {
					slide.load("shapes/items/$none");
				}
				await context.sync();

				// Step 3: Load text on each shape
				for (const slide of slides.items) {
					for (const shape of slide.shapes.items) {
						shape.load("textFrame/textRange/text");
					}
				}
				await context.sync();

				// Step 4: Extract titles (first shape with non-empty text)
				for (let i = 0; i < totalSlides; i++) {
					const shapes = slides.items[i].shapes.items;
					let title = "";

					for (const shape of shapes) {
						try {
							const text = String(shape.textFrame?.textRange?.text || "").trim();
							if (text && !title) {
								title = text;
							}
						} catch {
							// shape has no textFrame
						}
					}

					slideList.push({ index: i, title: title || `Slide ${i + 1}` });
				}

				resolve({
					documentName: "Presentation",
					totalSlides,
					slides: slideList,
				});
			} catch (error) {
				resolve({
					error: `PowerPoint.run failed: ${error instanceof Error ? error.message : String(error)}`,
				});
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
		return {
			slideIndex,
			title: "(not available)",
			shapes: [],
			speakerNotes: "",
			isHidden: false,
			error: "PowerPoint.run() not available",
		};
	}

	return new Promise((resolve) => {
		PowerPoint.run(async (context: any) => {
			try {
				const presentation = context.presentation;

				// Step 1: Load slides
				presentation.load("slides");
				await context.sync();

				if (slideIndex < 0 || slideIndex >= presentation.slides.items.length) {
					resolve({ error: `Slide index ${slideIndex} out of range (0-${presentation.slides.items.length - 1})` });
					return;
				}

				const slide = presentation.slides.items[slideIndex];

				// Step 2: Load shapes items (no properties yet)
				slide.load("shapes/items/$none");
				await context.sync();

				// Step 3: Load id, name, and text on each shape individually
				for (const shape of slide.shapes.items) {
					shape.load("id");
					shape.load("name");
					shape.load("textFrame/textRange/text");
				}
				await context.sync();

				// Step 4: Build result
				const shapeList: Array<{
					id: string;
					name: string;
					shapeType: string;
					text: string;
				}> = [];

				let slideTitle = "";
				for (const shape of slide.shapes.items) {
					let text = "";
					try {
						text = String(shape.textFrame?.textRange?.text || "");
					} catch {
						// no text frame
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
				resolve({
					error: `PowerPoint.run failed: ${error instanceof Error ? error.message : String(error)}`,
				});
			}
		});
	});
}

async function handleUpdateShapeText(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex?: number;
		shapeId?: string;
		text?: string;
		confirmationToken?: string;
	};

	const {
		slideIndex = 0,
		shapeId = "",
		text = "",
		confirmationToken: _confirmationToken,
	} = config;

	return {
		slideIndex,
		shapeId,
		newText: text,
		message: "Shape text update simulated (not yet implemented)",
	};
}

async function handleUpdateSpeakerNotes(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex?: number;
		notes?: string;
		confirmationToken?: string;
	};

	const {
		slideIndex = 0,
		notes = "",
		confirmationToken: _confirmationToken,
	} = config;

	return {
		slideIndex,
		newNotes: notes,
		message: "Speaker notes update simulated (not yet implemented)",
	};
}
