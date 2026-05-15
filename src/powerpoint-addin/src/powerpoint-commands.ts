/**
 * PowerPoint command handler using Office JS API.
 * Executes PowerPoint-specific commands received from the MCP server.
 */

/// <reference types="@types/office-js" />

import { reportResult } from "./communication";

// --- Command Processing ---

/**
 * Processes a command received from the MCP server.
 */
export async function processCommand(
	commandId: string,
	commandName: string,
	args: unknown,
): Promise<void> {
	try {
		let result: unknown;

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
	}
}

// --- Command Handlers ---

async function handleGetDeckOutline(_args: unknown): Promise<unknown> {
	// Stub: In a real implementation, this would use PowerPoint JS API:
	// await PowerPoint.run(async (context) => {
	//   const slides = context.presentation.slides;
	//   slides.load("id,position,title");
	//   await context.sync();
	//   // Iterate through slides and extract shape info
	// });

	return {
		documentName: "Presentation1.pptx",
		totalSlides: 0,
		slides: [],
		message: "Office JS PowerPoint API not available in this environment",
	};
}

async function handleGetSlide(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number };
	const slideIndex = config.slideIndex ?? 0;

	return {
		slideIndex,
		title: "Slide Title",
		shapes: [],
		speakerNotes: "",
		isHidden: false,
	};
}

async function handleUpdateShapeText(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex?: number;
		shapeId?: string;
		text?: string;
		confirmationToken?: string;
	};

	const { slideIndex = 0, shapeId = "", text = "" } = config;

	return {
		slideIndex,
		shapeId,
		newText: text,
		message: "Shape text update simulated (Office JS API not available)",
	};
}

async function handleUpdateSpeakerNotes(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex?: number;
		notes?: string;
		confirmationToken?: string;
	};

	const { slideIndex = 0, notes = "" } = config;

	return {
		slideIndex,
		newNotes: notes,
		message: "Speaker notes update simulated (Office JS API not available)",
	};
}
