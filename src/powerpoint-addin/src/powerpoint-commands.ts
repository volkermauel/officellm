/**
 * PowerPoint command handler using Office JS API.
 * Executes PowerPoint-specific commands received from the MCP server.
 *
 * Key Office JS patterns:
 * - context.load(obj, ["prop1", "prop2"]) queues a load
 * - await context.sync() actually fetches the data
 * - MUST sync BEFORE reading any loaded property
 * - slide.name is internal ("Slide 1"), NOT the title text
 *   — must read shapes to find the title
 */

/// <reference types="@types/office-js" />

import { reportResult } from "./communication";

// --- Command Processing ---

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

// --- Command Handlers ---

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
				const slides = presentation.slides;

				// Step 1: Load slides collection
				context.load(slides, ["items"]);
				await context.sync();

				const totalSlides = slides.items.length;
				const slideList: Array<{ index: number; title: string }> = [];

				// Step 2: For each slide, load its shapes
				for (let i = 0; i < totalSlides; i++) {
					const slide = slides.items[i];
					const shapes = slide.shapes;
					context.load(shapes, ["items"]);
				}
				await context.sync();

				// Step 3: For each shape, load text properties
				for (let i = 0; i < totalSlides; i++) {
					const shapes = slides.items[i].shapes.items;

					for (const shape of shapes) {
						context.load(shape, ["name", "textFrame"]);
					}

					slideList.push({ index: i, title: "" });
				}
				await context.sync();

				// Step 4: Read text from shapes (now synced)
				for (let i = 0; i < totalSlides; i++) {
					const shapes = slides.items[i].shapes.items;

					for (const shape of shapes) {
						try {
							if (shape.textFrame) {
								context.load(shape.textFrame, ["textRange"]);
							}
						} catch {
							// shape doesn't have a textFrame
						}
					}
				}
				await context.sync();

				// Step 5: Extract actual text
				for (let i = 0; i < totalSlides; i++) {
					const shapes = slides.items[i].shapes.items;
					let title = "";
					let foundTitle = false;

					for (const shape of shapes) {
						try {
							if (shape.textFrame && shape.textFrame.textRange) {
								const text = shape.textFrame.textRange.text || "";
								if (text.trim()) {
									// First non-empty text is likely the title
									if (!foundTitle) {
										title = text.trim();
										foundTitle = true;
									}
								}
							}
						} catch {
							// skip
						}
					}

					slideList[i].title = title || `Slide ${i + 1}`;
				}

				resolve({
					documentName: presentation.name || "Untitled",
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
				const slides = presentation.slides;

				// Step 1: Load slides
				context.load(slides, ["items"]);
				await context.sync();

				if (slideIndex < 0 || slideIndex >= slides.items.length) {
					resolve({ error: `Slide index ${slideIndex} out of range (0-${slides.items.length - 1})` });
					return;
				}

				const slide = slides.items[slideIndex];

				// Step 2: Load shapes on this slide
				const shapes = slide.shapes;
				context.load(shapes, ["items"]);
				await context.sync();

				// Step 3: Load properties for each shape
				for (const shape of shapes.items) {
					context.load(shape, ["id", "name", "shapeType", "textFrame"]);
				}
				await context.sync();

				// Step 4: Load text ranges
				for (const shape of shapes.items) {
					try {
						if (shape.textFrame) {
							context.load(shape.textFrame, ["textRange"]);
						}
					} catch {
						// no text frame
					}
				}
				await context.sync();

				// Step 5: Load actual text
				for (const shape of shapes.items) {
					try {
						if (shape.textFrame && shape.textFrame.textRange) {
							context.load(shape.textFrame.textRange, ["text"]);
						}
					} catch {
						// no text range
					}
				}
				await context.sync();

				// Step 6: Build shape list with actual text
				const shapeList: Array<{
					id: string;
					name: string;
					shapeType: string;
					text: string;
				}> = [];

				let slideTitle = "";
				for (const shape of shapes.items) {
					let text = "";
					try {
						if (shape.textFrame?.textRange?.text) {
							text = shape.textFrame.textRange.text;
						}
					} catch {
						// no text
					}

					shapeList.push({
						id: String(shape.id || ""),
						name: String(shape.name || ""),
						shapeType: String(shape.shapeType || "unknown"),
						text,
					});

					// First shape with non-empty text is the title
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

	return new Promise((resolve) => {
		resolve({
			slideIndex,
			shapeId,
			newText: text,
			message: "Shape text update simulated (Office JS API not available)",
		});
	});
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

	return new Promise((resolve) => {
		resolve({
			slideIndex,
			newNotes: notes,
			message: "Speaker notes update simulated (Office JS API not available)",
		});
	});
}
