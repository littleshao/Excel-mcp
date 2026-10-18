import { McpServer } from "@modelcontextprotocol/server";
import { z } from "zod";

import { error, trace } from "./logger";
import { ExcelManager } from "./excel-manager";
import { ExcelError } from "./types";

type ZodSchemaShape = Record<string, z.ZodType>;

interface ToolAnnotations {
	readOnlyHint?: boolean;
	destructiveHint?: boolean;
	openWorldHint?: boolean;
}

type ToolCallback = (args: any) => Promise<string>;

const SERVER_INSTRUCTIONS = `Manipulate Excel workbooks (.xlsx, .xls, .xlsm, .csv).

Start with excel_open_workbook (to load an existing file) or excel_create_workbook.
Each returns a workbook id that you must pass to every later call — there is no
implicit "current workbook". Call excel_save_workbook to write changes to disk
when done, and excel_close_workbook to free memory.

Coordinates use Excel-style addresses: A1, B2, AB10. Ranges are A1:C5. Sheet
can be given by name ("Sheet1") or by 0-based index (0).

When reading large ranges, prefer excel_read_range to get a 2D array in one
call rather than reading cell by cell. For structural edits (insert/delete
rows or columns), remember that cell references may shift.`;

export const getAgentVersion = (): string => {
	const json = require("../package.json");
	return json.version;
};

export const createMcpServer = (): McpServer => {
	const server = new McpServer({
		name: "excel-mcp",
		version: getAgentVersion(),
	}, {
		instructions: SERVER_INSTRUCTIONS,
	});

	const manager = new ExcelManager();

	const tool = (
		name: string,
		title: string,
		description: string,
		paramsSchema: ZodSchemaShape,
		annotations: ToolAnnotations,
		cb: ToolCallback,
	) => {
		server.registerTool(name, {
			title,
			description,
			inputSchema: z.object(paramsSchema),
			annotations,
		}, (async (args: any) => {
			try {
				trace(`Invoking ${name} with args: ${JSON.stringify(args)}`);
				const response = await cb(args);
				return { content: [{ type: "text", text: response }] };
			} catch (err: any) {
				const message = err instanceof ExcelError
					? `${err.message}. Please fix the issue and try again.`
					: `Error: ${err.message}`;
				error(`Tool ${name} failed: ${message}`);
				return {
					content: [{ type: "text", text: message }],
					isError: true,
				};
			}
		}) as any);
	};

	// ─── Workbook lifecycle ────────────────────────────────────────────────

	tool(
		"excel_create_workbook",
		"Create Workbook",
		"Create a new empty workbook with one default sheet. Returns a workbook id used by all other tools.",
		{
			filePath: z.string().optional().describe("Optional path to associate with the workbook (used when saving without a path)."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async ({ filePath }) => {
			const info = manager.createWorkbook(filePath);
			return JSON.stringify(info, null, 2);
		},
	);

	tool(
		"excel_open_workbook",
		"Open Workbook",
		"Open an existing Excel workbook from disk. Returns a workbook id used by all other tools.",
		{
			filePath: z.string().describe("Absolute path to the .xlsx, .xls, .xlsm or .csv file to open."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async ({ filePath }) => {
			const info = await manager.openWorkbook(filePath);
			return JSON.stringify(info, null, 2);
		},
	);

	tool(
		"excel_save_workbook",
		"Save Workbook",
		"Save the workbook to disk. If no filePath is given, saves to the original path the workbook was opened from.",
		{
			id: z.string().describe("The workbook id returned by excel_open_workbook or excel_create_workbook."),
			filePath: z.string().optional().describe("Path to save to. If omitted, saves to the workbook's current path."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async ({ id, filePath }) => {
			const saved = await manager.saveWorkbook(id, filePath);
			return `Workbook saved to: ${saved}`;
		},
	);

	tool(
		"excel_close_workbook",
		"Close Workbook",
		"Close a workbook session, removing it from memory. Unsaved changes are lost.",
		{
			id: z.string().describe("The workbook id to close."),
		},
		{ readOnlyHint: false, destructiveHint: true },
		async ({ id }) => {
			manager.closeWorkbook(id);
			return `Workbook ${id} closed.`;
		},
	);

	tool(
		"excel_list_workbooks",
		"List Workbooks",
		"List all currently open workbook sessions.",
		{},
		{ readOnlyHint: true },
		async () => {
			const workbooks = manager.listWorkbooks();
			if (workbooks.length === 0) {
				return "No workbooks are currently open.";
			}

			return JSON.stringify(workbooks, null, 2);
		},
	);

	tool(
		"excel_get_workbook_info",
		"Get Workbook Info",
		"Get summary information about a workbook: file path, sheet count, and sheet names.",
		{
			id: z.string().describe("The workbook id."),
		},
		{ readOnlyHint: true },
		async ({ id }) => {
			const info = manager.getWorkbookInfo(id);
			return JSON.stringify(info, null, 2);
		},
	);

	// ─── Sheet operations ──────────────────────────────────────────────────

	tool(
		"excel_list_sheets",
		"List Sheets",
		"List all sheets in a workbook with their dimensions and visibility.",
		{
			id: z.string().describe("The workbook id."),
		},
		{ readOnlyHint: true },
		async ({ id }) => {
			const sheets = manager.listSheets(id);
			return JSON.stringify(sheets, null, 2);
		},
	);

	tool(
		"excel_add_sheet",
		"Add Sheet",
		"Add a new sheet to the workbook.",
		{
			id: z.string().describe("The workbook id."),
			name: z.string().describe("Name of the new sheet."),
			tabColor: z.string().optional().describe("Optional ARGB hex color for the sheet tab, e.g. \"FF00FF00\"."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async ({ id, name, tabColor }) => {
			const sheet = manager.addSheet(id, name, tabColor);
			return `Sheet "${name}" added.`;
		},
	);

	tool(
		"excel_delete_sheet",
		"Delete Sheet",
		"Delete a sheet from the workbook. The last sheet cannot be deleted.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
		},
		{ readOnlyHint: false, destructiveHint: true },
		async ({ id, sheet }) => {
			manager.deleteSheet(id, sheet);
			return `Sheet "${sheet}" deleted.`;
		},
	);

	tool(
		"excel_rename_sheet",
		"Rename Sheet",
		"Rename a sheet.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Current sheet name or 0-based index."),
			newName: z.string().describe("New name for the sheet."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async ({ id, sheet, newName }) => {
			manager.renameSheet(id, sheet, newName);
			return `Sheet renamed to "${newName}".`;
		},
	);

	// ─── Cell operations ───────────────────────────────────────────────────

	tool(
		"excel_read_cell",
		"Read Cell",
		"Read a single cell's value, type, formula, number format and style.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			address: z.string().describe("Cell address, e.g. \"A1\"."),
		},
		{ readOnlyHint: true },
		async ({ id, sheet, address }) => {
			const cell = manager.readCell(id, sheet, address);
			return JSON.stringify(cell, null, 2);
		},
	);

	tool(
		"excel_write_cell",
		"Write Cell",
		"Write a value to a single cell.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			address: z.string().describe("Cell address, e.g. \"A1\"."),
			value: z.union([z.string(), z.number(), z.boolean(), z.null()]).describe("Value to write (string, number, boolean, or null)."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async ({ id, sheet, address, value }) => {
			const cell = manager.writeCell(id, sheet, address, value);
			return `Wrote to ${address}: ${JSON.stringify(value)}`;
		},
	);

	tool(
		"excel_read_range",
		"Read Range",
		"Read a rectangular range of cells as a 2D array of values.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			range: z.string().describe("Range address, e.g. \"A1:C10\"."),
		},
		{ readOnlyHint: true },
		async ({ id, sheet, range }) => {
			const data = manager.readRange(id, sheet, range);
			return JSON.stringify(data, null, 2);
		},
	);

	tool(
		"excel_write_range",
		"Write Range",
		"Write a 2D array of values to the sheet starting at the given cell address.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			startAddress: z.string().describe("Top-left cell address to start writing, e.g. \"A1\"."),
			data: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))).describe("2D array of values. Each inner array is a row."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async ({ id, sheet, startAddress, data }) => {
			const result = manager.writeRange(id, sheet, startAddress, data);
			return `Wrote ${result.rows} rows x ${result.cols} cols starting at ${startAddress}. Range: ${result.range}`;
		},
	);

	tool(
		"excel_clear_range",
		"Clear Range",
		"Clear the contents of a range. Optionally also clear cell styles.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			range: z.string().describe("Range address, e.g. \"A1:C10\"."),
			clearStyle: z.boolean().optional().describe("If true, also remove cell styles. Defaults to false."),
		},
		{ readOnlyHint: false, destructiveHint: true },
		async ({ id, sheet, range, clearStyle }) => {
			manager.clearRange(id, sheet, range, clearStyle ?? false);
			return `Range ${range} cleared.`;
		},
	);

	tool(
		"excel_set_formula",
		"Set Formula",
		"Set an Excel formula on a cell, e.g. \"=SUM(A1:A10)\". The server does not evaluate formulas: pass `result` to store the cached value that will be read back until Excel recalculates.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			address: z.string().describe("Cell address, e.g. \"B11\"."),
			formula: z.string().describe("The formula including the leading =, e.g. \"=A1+B1\"."),
			result: z.union([z.string(), z.number(), z.boolean(), z.null()]).optional().describe("Cached value of the formula after evaluation."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async ({ id, sheet, address, formula, result }) => {
			const cell = manager.setFormula(id, sheet, address, formula, result);
			return `Formula set on ${address}: ${formula}${result !== undefined ? ` (cached result: ${JSON.stringify(result)})` : ""}`;
		},
	);

	// ─── Row / Column operations ───────────────────────────────────────────

	tool(
		"excel_insert_rows",
		"Insert Rows",
		"Insert empty rows before the given row number.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			beforeRow: z.number().int().positive().describe("Row number before which to insert (1-based)."),
			count: z.number().int().positive().optional().describe("Number of rows to insert. Defaults to 1."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async ({ id, sheet, beforeRow, count }) => {
			manager.insertRows(id, sheet, beforeRow, count ?? 1);
			return `Inserted ${count ?? 1} row(s) before row ${beforeRow}.`;
		},
	);

	tool(
		"excel_insert_columns",
		"Insert Columns",
		"Insert empty columns before the given column number.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			beforeCol: z.number().int().positive().describe("Column number before which to insert (1-based)."),
			count: z.number().int().positive().optional().describe("Number of columns to insert. Defaults to 1."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async ({ id, sheet, beforeCol, count }) => {
			manager.insertColumns(id, sheet, beforeCol, count ?? 1);
			return `Inserted ${count ?? 1} column(s) before column ${beforeCol}.`;
		},
	);

	tool(
		"excel_delete_rows",
		"Delete Rows",
		"Delete rows from the sheet.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			startRow: z.number().int().positive().describe("First row to delete (1-based)."),
			count: z.number().int().positive().optional().describe("Number of rows to delete. Defaults to 1."),
		},
		{ readOnlyHint: false, destructiveHint: true },
		async ({ id, sheet, startRow, count }) => {
			manager.deleteRows(id, sheet, startRow, count ?? 1);
			return `Deleted ${count ?? 1} row(s) starting at row ${startRow}.`;
		},
	);

	tool(
		"excel_delete_columns",
		"Delete Columns",
		"Delete columns from the sheet.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			startCol: z.number().int().positive().describe("First column to delete (1-based)."),
			count: z.number().int().positive().optional().describe("Number of columns to delete. Defaults to 1."),
		},
		{ readOnlyHint: false, destructiveHint: true },
		async ({ id, sheet, startCol, count }) => {
			manager.deleteColumns(id, sheet, startCol, count ?? 1);
			return `Deleted ${count ?? 1} column(s) starting at column ${startCol}.`;
		},
	);

	tool(
		"excel_set_row_height",
		"Set Row Height",
		"Set the height of a row in points.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			row: z.number().int().positive().describe("Row number (1-based)."),
			height: z.number().positive().describe("Row height in points."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async ({ id, sheet, row, height }) => {
			manager.setRowHeight(id, sheet, row, height);
			return `Row ${row} height set to ${height}pt.`;
		},
	);

	tool(
		"excel_set_column_width",
		"Set Column Width",
		"Set the width of a column in characters.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			col: z.number().int().positive().describe("Column number (1-based)."),
			width: z.number().positive().describe("Column width in characters."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async ({ id, sheet, col, width }) => {
			manager.setColumnWidth(id, sheet, col, width);
			return `Column ${col} width set to ${width}.`;
		},
	);

	// ─── Style & Format ────────────────────────────────────────────────────

	tool(
		"excel_apply_style",
		"Apply Style",
		"Apply font, fill, alignment and border styles to a cell or range.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			target: z.string().describe("Cell address or range, e.g. \"A1\" or \"A1:D10\"."),
			fontName: z.string().optional().describe("Font family name, e.g. \"Arial\"."),
			fontSize: z.number().positive().optional().describe("Font size in points."),
			bold: z.boolean().optional().describe("Bold text."),
			italic: z.boolean().optional().describe("Italic text."),
			underline: z.boolean().optional().describe("Underlined text."),
			fontColor: z.string().optional().describe("Font color as ARGB hex, e.g. \"FFFF0000\" for red."),
			fillColor: z.string().optional().describe("Background fill color as ARGB hex, e.g. \"FFFFFF00\" for yellow."),
			horizontalAlign: z.enum(["left", "center", "right", "fill", "justify"]).optional().describe("Horizontal alignment."),
			verticalAlign: z.enum(["top", "middle", "bottom", "justify"]).optional().describe("Vertical alignment."),
			wrapText: z.boolean().optional().describe("Wrap text in cell."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async (args) => {
			const { id, sheet, target, fontName, fontSize, bold, italic, underline, fontColor, fillColor, horizontalAlign, verticalAlign, wrapText } = args;
			const style: any = {};
			if (fontName || fontSize || bold || italic || underline || fontColor) {
				style.font = {};
				if (fontName) style.font.name = fontName;
				if (fontSize) style.font.size = fontSize;
				if (bold !== undefined) style.font.bold = bold;
				if (italic !== undefined) style.font.italic = italic;
				if (underline !== undefined) style.font.underline = underline;
				if (fontColor) style.font.color = fontColor;
			}

			if (fillColor) {
				style.fill = { fgColor: fillColor };
			}

			if (horizontalAlign || verticalAlign || wrapText !== undefined) {
				style.alignment = {};
				if (horizontalAlign) style.alignment.horizontal = horizontalAlign;
				if (verticalAlign) style.alignment.vertical = verticalAlign;
				if (wrapText !== undefined) style.alignment.wrapText = wrapText;
			}

			manager.applyStyle(id, sheet, target, style);
			return `Style applied to ${target}.`;
		},
	);

	tool(
		"excel_set_number_format",
		"Set Number Format",
		"Set the number format (e.g. currency, date, percentage) on a cell or range.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			target: z.string().describe("Cell address or range, e.g. \"A1\" or \"A1:A100\"."),
			format: z.string().describe("Excel number format string, e.g. \"#,##0.00\", \"$#,##0.00\", \"yyyy-mm-dd\", \"0.00%\"."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async ({ id, sheet, target, format }) => {
			manager.setNumberFormat(id, sheet, target, format);
			return `Number format "${format}" applied to ${target}.`;
		},
	);

	tool(
		"excel_merge_cells",
		"Merge Cells",
		"Merge a range of cells into one.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			range: z.string().describe("Range to merge, e.g. \"A1:D1\"."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async ({ id, sheet, range }) => {
			manager.mergeCells(id, sheet, range);
			return `Merged range ${range}.`;
		},
	);

	tool(
		"excel_unmerge_cells",
		"Unmerge Cells",
		"Unmerge a previously merged range of cells.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			range: z.string().describe("Range to unmerge, e.g. \"A1:D1\"."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async ({ id, sheet, range }) => {
			manager.unmergeCells(id, sheet, range);
			return `Unmerged range ${range}.`;
		},
	);

	// ─── Data operations ───────────────────────────────────────────────────

	tool(
		"excel_sort_range",
		"Sort Range",
		"Sort the rows in a range by one or more columns.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			range: z.string().describe("Range to sort, e.g. \"A1:D100\"."),
			sortColumn: z.number().int().positive().describe("Column number to sort by (1-based, absolute)."),
			order: z.enum(["asc", "desc"]).optional().describe("Sort order. Defaults to asc."),
			hasHeader: z.boolean().optional().describe("If true, the first row is treated as a header and not sorted. Defaults to false."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async ({ id, sheet, range, sortColumn, order, hasHeader }) => {
			manager.sortRange(id, sheet, range, [{ column: sortColumn, order: order || "asc" }], hasHeader ?? false);
			return `Range ${range} sorted by column ${sortColumn} ${order || "asc"}.`;
		},
	);

	tool(
		"excel_find_data",
		"Find Data",
		"Search for a value or pattern within a range (or the entire sheet). Returns matching cells with their addresses and values.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			query: z.string().describe("The text or regex pattern to search for."),
			range: z.string().optional().describe("Range to search within, e.g. \"A1:Z100\". Defaults to the whole sheet."),
			matchCase: z.boolean().optional().describe("Case-sensitive search. Defaults to false."),
			useRegex: z.boolean().optional().describe("Treat the query as a regular expression. Defaults to false."),
		},
		{ readOnlyHint: true },
		async ({ id, sheet, query, range, matchCase, useRegex }) => {
			const results = manager.findData(id, sheet, query, range, matchCase ?? false, useRegex ?? false);
			if (results.length === 0) {
				return `No matches found for "${query}".`;
			}

			return JSON.stringify({ count: results.length, matches: results.slice(0, 100), truncated: results.length > 100 }, null, 2);
		},
	);

	tool(
		"excel_get_last_row",
		"Get Last Row",
		"Get the row number of the last non-empty row in the sheet.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
		},
		{ readOnlyHint: true },
		async ({ id, sheet }) => {
			const row = manager.getLastRow(id, sheet);
			return `Last non-empty row: ${row}`;
		},
	);

	tool(
		"excel_get_last_column",
		"Get Last Column",
		"Get the column number of the last non-empty column in the sheet.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
		},
		{ readOnlyHint: true },
		async ({ id, sheet }) => {
			const col = manager.getLastColumn(id, sheet);
			return `Last non-empty column: ${col}`;
		},
	);

	// ─── Charts ────────────────────────────────────────────────────────────

	tool(
		"excel_add_chart",
		"Add Chart",
		"Add a chart to a sheet based on data ranges.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			type: z.enum(["line", "bar", "pie", "doughnut", "area", "scatter"]).describe("Chart type."),
			position: z.string().describe("Cell address where the chart's top-left corner is placed, e.g. \"E2\"."),
			title: z.string().optional().describe("Chart title."),
			xAxisTitle: z.string().optional().describe("X-axis title."),
			yAxisTitle: z.string().optional().describe("Y-axis title."),
			seriesName: z.string().describe("Name of the data series."),
			valuesRange: z.string().describe("Range containing the data values, e.g. \"B2:B10\"."),
			categoriesRange: z.string().optional().describe("Range containing the category labels, e.g. \"A2:A10\"."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async ({ id, sheet, type, position, title, xAxisTitle, yAxisTitle, seriesName, valuesRange, categoriesRange }) => {
			const chartId = manager.addChart(id, sheet, {
				type,
				title,
				xAxisTitle,
				yAxisTitle,
				position,
				series: [{ name: seriesName, valuesRange, categoriesRange }],
			});
			return `Chart added at ${position} (id: ${chartId}).`;
		},
	);

	// ─── Import / Export ───────────────────────────────────────────────────

	tool(
		"excel_export_to_csv",
		"Export to CSV",
		"Export a sheet's contents as a CSV string.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
		},
		{ readOnlyHint: true },
		async ({ id, sheet }) => {
			return manager.exportToCsv(id, sheet);
		},
	);

	tool(
		"excel_import_from_csv",
		"Import from CSV",
		"Import CSV data into a sheet starting at the given cell.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			csv: z.string().describe("CSV data as a string."),
			startAddress: z.string().optional().describe("Top-left cell to start importing at. Defaults to A1."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async ({ id, sheet, csv, startAddress }) => {
			const result = manager.importFromCsv(id, sheet, csv, startAddress || "A1");
			return `Imported ${result.rows} rows x ${result.cols} cols. Range: ${result.range}`;
		},
	);

	tool(
		"excel_copy_range",
		"Copy Range",
		"Copy the values and styles from one range to another location within the same sheet.",
		{
			id: z.string().describe("The workbook id."),
			sheet: z.string().describe("Sheet name or 0-based index."),
			sourceRange: z.string().describe("Source range, e.g. \"A1:B5\"."),
			targetStart: z.string().describe("Top-left cell of the destination, e.g. \"D1\"."),
		},
		{ readOnlyHint: false, destructiveHint: false },
		async ({ id, sheet, sourceRange, targetStart }) => {
			manager.copyRange(id, sheet, sourceRange, targetStart);
			return `Copied ${sourceRange} to ${targetStart}.`;
		},
	);

	return server;
};
