/**
 * PowerPoint command handler using Office JS API.
 * Executes PowerPoint-specific commands received from the MCP server.
 */

/// <reference types="@types/office-js" />

import { reportResult } from "./communication";

// --- Command Processing ---

/**
 * Processes a command received from the MCP server.
 * Returns the result payload (may include requiresConfirmation flag).
 */
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
	return new Promise((resolve) => {
		// Config params: includeSpeakerNotes, includeHiddenSlides (used in full implementation)
		// eslint-disable-next-line @typescript-eslint/no-explicit-any
		const OfficeAny: any = (window as any).Office;
		if (!OfficeAny || typeof OfficeAny.context === "undefined") {
			resolve({ error: "PowerPoint context not available" });
			return;
		}

		// Use PowerPoint.Run() to access the presentation API
		// eslint-disable-next-line @typescript-eslint/no-explicit-any
		const PowerPoint: any = (window as any).PowerPoint;
		if (!PowerPoint || typeof PowerPoint.Run !== "function") {
			resolve({
				documentName: "(unknown)",
				totalSlides: 0,
				slides: [],
				message: "PowerPoint.Run() not available in this environment",
			});
			return;
		}

		PowerPoint.Run(async (context: any) => {
			try {
				const presentation = context.presentation;
				const slides = presentation.slides;
				context.load(slides, "notCoveredByParallelization");
				await context.sync();

				const slideList: Array<{
					index: number;
					title: string;
					hasSpeakerNotes: boolean;
				}> = [];

				for (let i = 0; i < slides.items.length; i++) {
					const slide = slides.items[i];
					context.load(slide, "name,slideLayoutItem");
					slideList.push({
						index: i,
						title: slide.name || `Slide ${i + 1}`,
						hasSpeakerNotes: false,
					});
				}

				await context.sync();

				resolve({
					documentName: presentation.name || "Untitled",
					totalSlides: slideList.length,
					slides: slideList,
				});
			} catch (error) {
				resolve({ error: error instanceof Error ? error.message : String(error) });
			}
		});
	});
}

async function handleGetSlide(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number };
	const slideIndex = config.slideIndex ?? 0;

	return new Promise((resolve) => {
		// eslint-disable-next-line @typescript-eslint/no-explicit-any
		const PowerPoint: any = (window as any).PowerPoint;
		if (!PowerPoint || typeof PowerPoint.Run !== "function") {
			resolve({
				slideIndex,
				title: "(not available)",
				shapes: [],
				speakerNotes: "",
				isHidden: false,
			});
			return;
		}

		PowerPoint.Run(async (context: any) => {
			try {
				const presentation = context.presentation;
				const slides = presentation.slides;
				context.load(slides);
				await context.sync();

				if (slideIndex < 0 || slideIndex >= slides.items.length) {
					resolve({ error: `Slide index ${slideIndex} out of range` });
					return;
				}

				const slide = slides.items[slideIndex];
				const shapes = slide.shapes;
				context.load(shapes, "notCoveredByParallelization");
				await context.sync();

				const shapeList: Array<{
					id: string;
					name: string;
					shapeType: string;
					text: string;
				}> = [];

				for (const shape of shapes.items) {
					context.load(shape, "name,shapeType");
					if (shape.hasTextFrame()) {
						const textFrame = shape.getTextFrame();
						context.load(textFrame, "textRange");
						await context.sync();
						shapeList.push({
							id: shape.id || "",
							name: shape.name || "",
							shapeType: shape.shapeType || "unknown",
							text: textFrame.textRange.text || "",
						});
					} else {
						shapeList.push({
							id: shape.id || "",
							name: shape.name || "",
							shapeType: shape.shapeType || "unknown",
							text: "",
						});
					}
				}

				resolve({
					slideIndex,
					title: slide.name || `Slide ${slideIndex + 1}`,
					shapes: shapeList,
					speakerNotes: "",
					isHidden: false,
				});
			} catch (error) {
				resolve({ error: error instanceof Error ? error.message : String(error) });
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
		// Stub: In a real implementation, this would use PowerPoint JS API
		// to update shape text
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
		// Stub: In a real implementation, this would use PowerPoint JS API
		// to update speaker notes
		resolve({
			slideIndex,
			newNotes: notes,
			message: "Speaker notes update simulated (Office JS API not available)",
		});
	});
}
