/**
 * Excel command handler using Office JS API.
 *
 * Key Excel JS API patterns:
 * - Excel.run(async (context) => { ... }) for batched operations
 * - context.workbook.worksheets.getItem(name) for sheet access
 * - sheet.getRange("A1:D10") for range access
 * - range.values — 2D array for reading/writing cell values
 * - range.formulas — 2D array for reading/writing formulas
 * - range.text — 2D array of display strings (read-only)
 * - range.numberFormat — 2D array for number format strings
 * - Must sync() before reading any loaded property
 * - Tables: sheet.tables.add(address, hasHeaders) creates a ListObject
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
			case "excel_get_workbook_map":
				result = await handleGetWorkbookMap(args);
				break;
			case "excel_read_range":
				result = await handleReadRange(args);
				break;
			case "excel_write_range":
				result = await handleWriteRange(args);
				break;
			case "excel_write_formula":
				result = await handleWriteFormula(args);
				break;
			case "excel_create_table":
				result = await handleCreateTable(args);
				break;
			default:
				result = { error: `Unknown Excel command: ${commandName}` };
		}

		if (result && typeof result === "object" && "error" in result) {
			success = false;
		}
	} catch (error) {
		const errorMessage = error instanceof Error ? error.message : String(error);
		console.error(`Excel command ${commandId} failed:`, errorMessage);
		success = false;
		result = { error: errorMessage };
	}

	await reportResult(commandId, success, undefined, result);
	return result;
}

// ── Helpers ─────────────────────────────────────────────────────

function runInExcel<T>(fn: (ctx: any) => Promise<T>): Promise<T> {
	return new Promise<T>((resolve, reject) => {
		const Excel: any = (globalThis as any).Excel;
		if (!Excel || typeof Excel.run !== "function") {
			reject(new Error("Excel.run() not available"));
			return;
		}
		Excel.run(async (ctx: any) => {
			resolve(await fn(ctx));
		}).catch(reject);
	});
}

// ── Read tools ──────────────────────────────────────────────────

async function handleGetWorkbookMap(args: unknown): Promise<unknown> {
	const config = args as {
		includeTables?: boolean;
		includeNamedRanges?: boolean;
	};
	const includeTables = config.includeTables ?? true;
	const includeNamedRanges = config.includeNamedRanges ?? true;

	return runInExcel(async (ctx) => {
		const worksheets = ctx.workbook.worksheets;
		worksheets.load("items");
		await ctx.sync();

		const sheets: Array<{
			name: string;
			position: number;
			visibility: string;
			usedRange: string;
			rowCount: number;
			columnCount: number;
			tables?: Array<{ name: string; range: string; columns: string[] }>;
		}> = [];

		for (let i = 0; i < worksheets.items.length; i++) {
			const ws = worksheets.items[i];
			ws.load("name,position,visibility");
		}
		await ctx.sync();

		for (const ws of worksheets.items) {
			const usedRange = ws.getUsedRange();
			usedRange.load("address,columnCount,rowCount");
			await ctx.sync();

			const sheetInfo: (typeof sheets)[0] = {
				name: ws.name,
				position: ws.position,
				visibility: String(ws.visibility || "Visible"),
				usedRange: usedRange.address
					? String(usedRange.address).split("!").pop() || ""
					: "",
				rowCount: usedRange.rowCount || 0,
				columnCount: usedRange.columnCount || 0,
			};

			if (includeTables) {
				const tables = ws.tables;
				tables.load("items");
				await ctx.sync();

				const tableList: Array<{
					name: string;
					range: string;
					columns: string[];
				}> = [];
				for (const t of tables.items) {
					t.load("name");
					t.range.load("address");
					await ctx.sync();

					const cols = t.columns;
					cols.load("items");
					await ctx.sync();

					const colNames: string[] = [];
					for (const col of cols.items) {
						col.load("name");
					}
					await ctx.sync();
					for (const col of cols.items) {
						colNames.push(String(col.name));
					}

					tableList.push({
						name: String(t.name),
						range: String(t.range.address).split("!").pop() || "",
						columns: colNames,
					});
				}
				sheetInfo.tables = tableList;
			}

			sheets.push(sheetInfo);
		}

		const result: Record<string, unknown> = {
			sheetCount: sheets.length,
			sheets,
		};

		if (includeNamedRanges) {
			const names = ctx.workbook.names;
			names.load("items");
			await ctx.sync();

			const namedRanges: Array<{ name: string; refersTo: string }> = [];
			for (const nr of names.items) {
				nr.load("name,refersTo");
			}
			await ctx.sync();
			for (const nr of names.items) {
				namedRanges.push({
					name: String(nr.name),
					refersTo: String(nr.refersTo),
				});
			}
			(result as any).namedRanges = namedRanges;
		}

		return result;
	});
}

async function handleReadRange(args: unknown): Promise<unknown> {
	const config = args as {
		sheetName?: string;
		address?: string;
		includeFormulas?: boolean;
		includeNumberFormats?: boolean;
	};
	const {
		sheetName = "",
		address = "",
		includeFormulas = true,
		includeNumberFormats = false,
	} = config;

	if (!sheetName) return { error: "sheetName is required" };
	if (!address) return { error: "address is required" };

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		const range = sheet.getRange(address);

		const propsToLoad = [
			"values",
			"text",
			"rowCount",
			"columnCount",
			"address",
		];
		if (includeFormulas) propsToLoad.push("formulas");
		if (includeNumberFormats) propsToLoad.push("numberFormat");

		range.load(propsToLoad.join(","));
		await ctx.sync();

		const rowCount = range.rowCount as number;
		const columnCount = range.columnCount as number;
		const totalCells = rowCount * columnCount;

		// Size check: 256KB max (~250K characters)
		const values = range.values as (string | number | boolean | null)[][];
		const serializedSize = JSON.stringify(values).length;
		const maxBytes = 256 * 1024;

		const result: Record<string, unknown> = {
			sheetName,
			address: String(range.address).split("!").pop() || address,
			rowCount,
			columnCount,
			totalCells,
			values,
			text: range.text,
		};

		if (includeFormulas) {
			result.formulas = range.formulas;
		}

		if (includeNumberFormats) {
			result.numberFormats = range.numberFormat;
		}

		if (serializedSize > maxBytes) {
			result.truncated = true;
			result.byteCount = serializedSize;
		}

		return result;
	});
}

// ── Write tools ─────────────────────────────────────────────────

async function handleWriteRange(args: unknown): Promise<unknown> {
	const config = args as {
		sheetName?: string;
		address?: string;
		values?: (string | number | boolean | null)[][];
	};
	const { sheetName = "", address = "", values = [] } = config;

	if (!sheetName) return { error: "sheetName is required" };
	if (!address) return { error: "address is required" };
	if (!values || values.length === 0)
		return { error: "values must be a non-empty 2D array" };

	// Range size check: reject > 1M cells
	const totalCells = values.reduce((acc, row) => acc + row.length, 0);
	if (totalCells > 1_000_000) {
		return {
			error: `Range too large: ${totalCells} cells. Maximum is 1,000,000.`,
			errorCode: "RANGE_TOO_LARGE",
		};
	}

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		const range = sheet.getRange(address);

		// Read before-values for diff preview
		range.load("values");
		await ctx.sync();

		const beforeValues = range.values as (string | number | boolean | null)[][];

		// Write new values
		range.values = values;
		await ctx.sync();

		// Build diff preview (sample first 5 rows)
		const sampleRows = Math.min(values.length, 5);
		const diff: Array<{
			row: number;
			col: number;
			before: string;
			after: string;
		}> = [];

		for (let r = 0; r < sampleRows; r++) {
			for (let c = 0; c < values[r].length; c++) {
				const before = beforeValues?.[r]?.[c] ?? null;
				const after = values[r][c];
				if (String(before) !== String(after)) {
					diff.push({
						row: r,
						col: c,
						before: String(before ?? ""),
						after: String(after ?? ""),
					});
				}
			}
		}

		return {
			sheetName,
			address,
			rowsWritten: values.length,
			columnsWritten: values[0]?.length || 0,
			totalCells: values.reduce((acc, row) => acc + row.length, 0),
			written: true,
			diffPreview: diff,
			undoable: true,
		};
	});
}

async function handleWriteFormula(args: unknown): Promise<unknown> {
	const config = args as {
		sheetName?: string;
		address?: string;
		formula?: string;
	};
	const { sheetName = "", address = "", formula = "" } = config;

	if (!sheetName) return { error: "sheetName is required" };
	if (!address) return { error: "address is required" };
	if (!formula) return { error: "formula is required" };

	// Formula validation: must start with '='
	if (!formula.startsWith("=")) {
		return {
			error: `Invalid formula: must start with '='. Got: '${formula}'`,
			errorCode: "INVALID_FORMULA",
		};
	}

	// Basic syntax check: check for obvious mismatched parentheses
	let depth = 0;
	for (const ch of formula) {
		if (ch === "(") depth++;
		if (ch === ")") depth--;
		if (depth < 0) {
			return {
				error: `Invalid formula: unmatched closing parenthesis in '${formula}'`,
				errorCode: "INVALID_FORMULA",
			};
		}
	}
	if (depth !== 0) {
		return {
			error: `Invalid formula: unmatched opening parenthesis in '${formula}'`,
			errorCode: "INVALID_FORMULA",
		};
	}

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		const range = sheet.getRange(address);

		// Build the formula array — one formula per row in the range
		range.load("rowCount,columnCount");
		await ctx.sync();

		const rowCount = range.rowCount as number;
		const colCount = range.columnCount as number;

		// If single cell, wrap formula; if range, repeat formula across cells
		const formulaArray: string[][] = [];
		for (let r = 0; r < rowCount; r++) {
			const row: string[] = [];
			for (let c = 0; c < colCount; c++) {
				row.push(formula);
			}
			formulaArray.push(row);
		}

		range.formulas = formulaArray;
		await ctx.sync();

		return {
			sheetName,
			address,
			formula,
			cellsWritten: rowCount * colCount,
			written: true,
			undoable: true,
		};
	});
}

async function handleCreateTable(args: unknown): Promise<unknown> {
	const config = args as {
		sheetName?: string;
		address?: string;
		tableName?: string;
		hasHeaders?: boolean;
	};
	const { sheetName = "", address = "", tableName, hasHeaders = true } = config;

	if (!sheetName) return { error: "sheetName is required" };
	if (!address) return { error: "address is required" };

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		const table = sheet.tables.add(address, hasHeaders);

		table.load("name");
		table.range.load("address");
		await ctx.sync();

		let finalName = String(table.name);

		// Rename if custom name provided
		if (tableName) {
			table.name = tableName;
			await ctx.sync();
			finalName = tableName;
		}

		return {
			sheetName,
			address,
			tableName: finalName,
			hasHeaders,
			created: true,
			undoable: true,
		};
	});
}
