import { Workbook, Worksheet, Cell } from "exceljs";
import { randomUUID } from "node:crypto";
import { existsSync, writeFileSync } from "node:fs";

import {
	CellInfo,
	CellValue,
	CellStyle,
	ChartConfig,
	ExcelError,
	FindResult,
	RangeData,
	SheetInfo,
	SortConfig,
	WorkbookInfo,
} from "./types";
import {
	buildRange,
	cellAddress,
	isValidExcelExtension,
	parseCellAddress,
	parseRange,
} from "./excel-utils";
import { trace } from "./logger";

interface WorkbookSession {
	workbook: Workbook;
	filePath: string;
	dirty: boolean;
}

export class ExcelManager {
	private sessions = new Map<string, WorkbookSession>();

	// ─── Workbook lifecycle ────────────────────────────────────────────────

	/**
	 * Create a new empty workbook and register it as a session.
	 */
	public createWorkbook(filePath?: string): WorkbookInfo {
		const workbook = new Workbook();
		workbook.addWorksheet("Sheet1");
		const id = randomUUID();
		this.sessions.set(id, { workbook, filePath: filePath || "", dirty: true });
		return this.getWorkbookInfo(id);
	}

	/**
	 * Open an existing workbook from disk.
	 */
	public async openWorkbook(filePath: string): Promise<WorkbookInfo> {
		if (!existsSync(filePath)) {
			throw new ExcelError(`File not found: ${filePath}`);
		}

		if (!isValidExcelExtension(filePath)) {
			throw new ExcelError(`Unsupported file type: ${filePath}. Supported: .xlsx, .xls, .xlsm, .csv`);
		}

		const workbook = new Workbook();
		try {
			if (filePath.toLowerCase().endsWith(".csv")) {
				await workbook.csv.readFile(filePath);
			} else {
				await workbook.xlsx.readFile(filePath);
			}
		} catch (err: any) {
			throw new ExcelError(`Failed to open workbook: ${err.message}`);
		}

		const id = randomUUID();
		this.sessions.set(id, { workbook, filePath, dirty: false });
		return this.getWorkbookInfo(id);
	}

	/**
	 * Save a workbook to disk. If no path given, save to the original path.
	 */
	public async saveWorkbook(id: string, filePath?: string): Promise<string> {
		const session = this.getSession(id);
		const target = filePath || session.filePath;
		if (!target) {
			throw new ExcelError("No file path specified and workbook has no original path. Provide a filePath.");
		}

		try {
			if (target.toLowerCase().endsWith(".csv")) {
				// exceljs exposes CSV only on Workbook for reading; export the first
				// sheet through our own CSV serializer for symmetric single-sheet output.
				const firstSheet = session.workbook.worksheets[0];
				if (!firstSheet) {
					throw new ExcelError("Workbook has no sheets to export.");
				}

				writeFileSync(target, this.sheetToCsv(firstSheet), "utf8");
			} else {
				await session.workbook.xlsx.writeFile(target);
			}
		} catch (err: any) {
			if (err instanceof ExcelError) {
				throw err;
			}

			throw new ExcelError(`Failed to save workbook: ${err.message}`);
		}

		session.filePath = target;
		session.dirty = false;
		trace(`Saved workbook ${id} -> ${target}`);
		return target;
	}

	/**
	 * Close a workbook session (removes from memory without saving).
	 */
	public closeWorkbook(id: string): void {
		if (!this.sessions.has(id)) {
			throw new ExcelError(`Workbook not found: ${id}`);
		}

		this.sessions.delete(id);
	}

	/**
	 * List all open workbook sessions.
	 */
	public listWorkbooks(): WorkbookInfo[] {
		const result: WorkbookInfo[] = [];
		for (const id of this.sessions.keys()) {
			result.push(this.getWorkbookInfo(id));
		}

		return result;
	}

	/**
	 * Get summary info for a workbook.
	 */
	public getWorkbookInfo(id: string): WorkbookInfo {
		const session = this.getSession(id);
		return {
			id,
			filePath: session.filePath,
			sheetCount: session.workbook.worksheets.length,
			sheets: session.workbook.worksheets.map(ws => ws.name),
		};
	}

	// ─── Sheet operations ──────────────────────────────────────────────────

	/**
	 * List all sheets in a workbook with their metadata.
	 */
	public listSheets(id: string): SheetInfo[] {
		const session = this.getSession(id);
		return session.workbook.worksheets.map((ws, index) => ({
			name: ws.name,
			index,
			rowCount: ws.rowCount,
			columnCount: ws.columnCount,
			actualRowCount: ws.actualRowCount,
			actualColumnCount: ws.actualColumnCount,
			state: ws.state as SheetInfo["state"],
		}));
	}

	/**
	 * Add a new sheet.
	 */
	public addSheet(id: string, name: string, tabColor?: string): SheetInfo {
		const session = this.getSession(id);
		if (session.workbook.worksheets.some(ws => ws.name === name)) {
			throw new ExcelError(`Sheet "${name}" already exists.`);
		}

		const options = tabColor ? { properties: { tabColor: { argb: tabColor } } } : undefined;
		const ws = session.workbook.addWorksheet(name, options as any);
		session.dirty = true;
		return {
			name: ws.name,
			index: session.workbook.worksheets.indexOf(ws),
			rowCount: ws.rowCount,
			columnCount: ws.columnCount,
			actualRowCount: ws.actualRowCount,
			actualColumnCount: ws.actualColumnCount,
			state: "visible",
		};
	}

	/**
	 * Delete a sheet by name or index.
	 */
	public deleteSheet(id: string, sheet: string | number): void {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		if (session.workbook.worksheets.length <= 1) {
			throw new ExcelError("Cannot delete the last sheet in a workbook.");
		}

		ws.destroy();
		session.dirty = true;
	}

	/**
	 * Rename a sheet.
	 */
	public renameSheet(id: string, sheet: string | number, newName: string): void {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		if (session.workbook.worksheets.some(w => w.name === newName && w !== ws)) {
			throw new ExcelError(`Sheet "${newName}" already exists.`);
		}

		ws.name = newName;
		session.dirty = true;
	}

	// ─── Cell operations ───────────────────────────────────────────────────

	/**
	 * Read a single cell.
	 */
	public readCell(id: string, sheet: string | number, address: string): CellInfo {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		const { row, col } = parseCellAddress(address);
		return this.cellToInfo(ws.getCell(row, col));
	}

	/**
	 * Write a value to a cell.
	 */
	public writeCell(id: string, sheet: string | number, address: string, value: CellValue): CellInfo {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		const { row, col } = parseCellAddress(address);
		const cell = ws.getCell(row, col);
		cell.value = value as any;
		session.dirty = true;
		return this.cellToInfo(cell);
	}

	/**
	 * Read a range of cells as a 2D array.
	 */
	public readRange(id: string, sheet: string | number, range: string): RangeData {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		const { startRow, startCol, endRow, endCol } = parseRange(range);

		const data: CellValue[][] = [];
		for (let r = startRow; r <= endRow; r++) {
			const rowValues: CellValue[] = [];
			for (let c = startCol; c <= endCol; c++) {
				rowValues.push(this.extractValue(ws.getCell(r, c)));
			}

			data.push(rowValues);
		}

		return {
			range: buildRange(startRow, startCol, endRow, endCol),
			rows: endRow - startRow + 1,
			cols: endCol - startCol + 1,
			data,
		};
	}

	/**
	 * Write a 2D array of values starting at a given cell.
	 */
	public writeRange(id: string, sheet: string | number, startAddress: string, data: CellValue[][]): RangeData {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		const { row: startRow, col: startCol } = parseCellAddress(startAddress);

		const rows = data.length;
		const cols = rows > 0 ? Math.max(...data.map(row => row.length)) : 0;

		for (let r = 0; r < rows; r++) {
			for (let c = 0; c < data[r].length; c++) {
				ws.getCell(startRow + r, startCol + c).value = data[r][c] as any;
			}
		}

		session.dirty = true;
		const endRow = startRow + rows - 1;
		const endCol = startCol + cols - 1;
		return {
			range: buildRange(startRow, startCol, endRow, endCol),
			rows,
			cols,
			data,
		};
	}

	/**
	 * Clear the contents (and optionally styles) of a range.
	 */
	public clearRange(id: string, sheet: string | number, range: string, clearStyle = false): void {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		const { startRow, startCol, endRow, endCol } = parseRange(range);

		for (let r = startRow; r <= endRow; r++) {
			for (let c = startCol; c <= endCol; c++) {
				const cell = ws.getCell(r, c);
				cell.value = null;
				if (clearStyle) {
					cell.style = {} as any;
				}
			}
		}

		session.dirty = true;
	}

	/**
	 * Set a formula on a cell. The leading "=" is optional.
	 * exceljs does not evaluate formulas; pass `result` to store the cached
	 * value that other tools (and Excel before recalculation) will read back.
	 */
	public setFormula(id: string, sheet: string | number, address: string, formula: string, result?: CellValue): CellInfo {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		const { row, col } = parseCellAddress(address);
		const cell = ws.getCell(row, col);
		// exceljs expects the formula body without the leading "="
		const body = formula.startsWith("=") ? formula.slice(1) : formula;
		cell.value = (result !== undefined ? { formula: body, result } : { formula: body }) as any;
		session.dirty = true;
		return this.cellToInfo(cell);
	}

	// ─── Row / Column operations ───────────────────────────────────────────

	/**
	 * Insert rows before the given row number.
	 */
	public insertRows(id: string, sheet: string | number, beforeRow: number, count = 1): void {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		const blankRows: any[] = Array.from({ length: count }, () => []);
		ws.insertRows(beforeRow, blankRows);
		session.dirty = true;
	}

	/**
	 * Insert columns before the given column number.
	 */
	public insertColumns(id: string, sheet: string | number, beforeCol: number, count = 1): void {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		// spliceColumns inserts the columns passed as trailing args
		const blankColumns: any[] = Array.from({ length: count }, () => []);
		ws.spliceColumns(beforeCol, 0, ...blankColumns);
		session.dirty = true;
	}

	/**
	 * Delete rows.
	 */
	public deleteRows(id: string, sheet: string | number, startRow: number, count = 1): void {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		ws.spliceRows(startRow, count);
		session.dirty = true;
	}

	/**
	 * Delete columns.
	 */
	public deleteColumns(id: string, sheet: string | number, startCol: number, count = 1): void {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		ws.spliceColumns(startCol, count);
		session.dirty = true;
	}

	/**
	 * Set row height.
	 */
	public setRowHeight(id: string, sheet: string | number, row: number, height: number): void {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		ws.getRow(row).height = height;
		session.dirty = true;
	}

	/**
	 * Set column width.
	 */
	public setColumnWidth(id: string, sheet: string | number, col: number, width: number): void {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		ws.getColumn(col).width = width;
		session.dirty = true;
	}

	// ─── Style & Format ────────────────────────────────────────────────────

	/**
	 * Apply a style to a cell or range.
	 */
	public applyStyle(id: string, sheet: string | number, target: string, style: CellStyle): void {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);

		if (target.includes(":")) {
			const { startRow, startCol, endRow, endCol } = parseRange(target);
			for (let r = startRow; r <= endRow; r++) {
				for (let c = startCol; c <= endCol; c++) {
					this.mergeStyle(ws.getCell(r, c), style);
				}
			}
		} else {
			const { row, col } = parseCellAddress(target);
			this.mergeStyle(ws.getCell(row, col), style);
		}

		session.dirty = true;
	}

	/**
	 * Set number format on a cell or range.
	 */
	public setNumberFormat(id: string, sheet: string | number, target: string, format: string): void {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);

		if (target.includes(":")) {
			const { startRow, startCol, endRow, endCol } = parseRange(target);
			for (let r = startRow; r <= endRow; r++) {
				for (let c = startCol; c <= endCol; c++) {
					ws.getCell(r, c).numFmt = format;
				}
			}
		} else {
			const { row, col } = parseCellAddress(target);
			ws.getCell(row, col).numFmt = format;
		}

		session.dirty = true;
	}

	/**
	 * Merge a range of cells.
	 */
	public mergeCells(id: string, sheet: string | number, range: string): void {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		const { startRow, startCol, endRow, endCol } = parseRange(range);
		ws.mergeCells(startRow, startCol, endRow, endCol);
		session.dirty = true;
	}

	/**
	 * Unmerge a range of cells.
	 */
	public unmergeCells(id: string, sheet: string | number, range: string): void {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		const { startRow, startCol, endRow, endCol } = parseRange(range);
		ws.unMergeCells(startRow, startCol, endRow, endCol);
		session.dirty = true;
	}

	// ─── Data operations ───────────────────────────────────────────────────

	/**
	 * Sort a range by one or more columns.
	 */
	public sortRange(id: string, sheet: string | number, range: string, sorts: SortConfig[], hasHeader = false): void {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		const { startRow, startCol, endRow, endCol } = parseRange(range);

		if (sorts.length === 0) {
			throw new ExcelError("At least one sort column must be specified.");
		}

		const dataStartRow = hasHeader ? startRow + 1 : startRow;

		// Extract rows
		const rows: { values: any[] }[] = [];
		for (let r = dataStartRow; r <= endRow; r++) {
			const values: any[] = [];
			for (let c = startCol; c <= endCol; c++) {
				values.push(ws.getCell(r, c).value);
			}

			rows.push({ values });
		}

		// Sort
		rows.sort((a, b) => {
			for (const sort of sorts) {
				const colIdx = sort.column - startCol;
				const cmp = this.compareValues(a.values[colIdx], b.values[colIdx]);
				if (cmp !== 0) {
					return sort.order === "desc" ? -cmp : cmp;
				}
			}

			return 0;
		});

		// Write back
		for (let i = 0; i < rows.length; i++) {
			const targetRow = dataStartRow + i;
			for (let c = startCol; c <= endCol; c++) {
				ws.getCell(targetRow, c).value = rows[i].values[c - startCol];
			}
		}

		session.dirty = true;
	}

	/**
	 * Find cells matching a value or pattern within a range or sheet.
	 */
	public findData(id: string, sheet: string | number, query: string, range?: string, matchCase = false, useRegex = false): FindResult[] {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);

		let startRow = 1;
		let startCol = 1;
		let endRow = 1;
		let endCol = 1;
		if (range) {
			const parsed = parseRange(range);
			startRow = parsed.startRow;
			startCol = parsed.startCol;
			endRow = parsed.endRow;
			endCol = parsed.endCol;
		} else {
			const bounds = this.sheetBounds(ws);
			endRow = Math.max(bounds.lastRow, 1);
			endCol = Math.max(bounds.lastCol, 1);
		}

		const results: FindResult[] = [];
		const pattern = useRegex ? new RegExp(query, matchCase ? "" : "i") : null;
		const normalizedQuery = matchCase ? query : query.toLowerCase();

		for (let r = startRow; r <= endRow; r++) {
			for (let c = startCol; c <= endCol; c++) {
				const value = this.extractValue(ws.getCell(r, c));
				const strValue = value === null || value === undefined ? "" : String(value);
				const candidate = matchCase ? strValue : strValue.toLowerCase();

				const matched = pattern ? pattern.test(strValue) : candidate.includes(normalizedQuery);
				if (matched) {
					results.push({
						cell: cellAddress(r, c),
						value,
						row: r,
						col: c,
					});
				}
			}
		}

		return results;
	}

	/**
	 * Get the last non-empty row in a sheet (1-based row number).
	 */
	public getLastRow(id: string, sheet: string | number): number {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		return this.sheetBounds(ws).lastRow;
	}

	/**
	 * Get the last non-empty column in a sheet (1-based column number).
	 */
	public getLastColumn(id: string, sheet: string | number): number {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		return this.sheetBounds(ws).lastCol;
	}

	// ─── Charts ────────────────────────────────────────────────────────────

	/**
	 * Add a chart to a sheet based on data ranges.
	 */
	public addChart(id: string, sheet: string | number, config: ChartConfig): string {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);

		const typeMap: Record<string, string> = {
			line: "line",
			bar: "bar",
			pie: "pie",
			doughnut: "doughnut",
			area: "area",
			scatter: "scatter",
		};

		// Approximate pixel position from the anchor cell (col ~64px, row ~20px)
		const pos = parseCellAddress(config.position);
		const x = Math.max(pos.col - 1, 0) * 64;
		const y = Math.max(pos.row - 1, 0) * 20;

		// exceljs supports charts at runtime but its typings omit addChart
		(ws as any).addChart(typeMap[config.type] as any, {
			title: config.title,
			data: config.series.map(s => ({
				name: s.name,
				categories: s.categoriesRange,
				values: s.valuesRange,
			})),
			axes: [
				config.xAxisTitle ? { title: config.xAxisTitle } : {},
				config.yAxisTitle ? { title: config.yAxisTitle } : {},
			],
		} as any, {
			x,
			y,
			w: 480,
			h: 288,
		} as any);

		session.dirty = true;
		return `chart_${randomUUID().slice(0, 8)}`;
	}

	/**
	 * List media objects (charts and images) anchored on a sheet.
	 */
	public listCharts(id: string, sheet: string | number): string[] {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		const media = (ws as any)._media || [];
		return media
			.filter((m: any) => m.type === "chart")
			.map((m: any, i: number) => `chart_${i} (${m.chart?.type ?? "unknown"})`);
	}

	// ─── Import / Export ───────────────────────────────────────────────────

	/**
	 * Export a sheet to CSV string.
	 */
	public exportToCsv(id: string, sheet: string | number): string {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		return this.sheetToCsv(ws);
	}

	/**
	 * True max row/column numbers with values. exceljs's actualRowCount counts
	 * non-empty rows (not the max row number), which breaks trailing ranges,
	 * so scan eachRow (which yields only rows holding values) instead.
	 */
	private sheetBounds(ws: Worksheet): { lastRow: number; lastCol: number } {
		let lastRow = 0;
		let lastCol = 0;
		ws.eachRow({ includeEmpty: false }, row => {
			if (row.number > lastRow) {
				lastRow = row.number;
			}

			// row.values[0] is a null placeholder; array length = max col + 1
			const vals = row.values as any;
			const rowMaxCol = (Array.isArray(vals) ? vals.length : 1) - 1;
			if (rowMaxCol > lastCol) {
				lastCol = rowMaxCol;
			}
		});

		return { lastRow: Math.max(lastRow, 0), lastCol: Math.max(lastCol, 0) };
	}

	private sheetToCsv(ws: Worksheet): string {
		const { lastRow, lastCol } = this.sheetBounds(ws);
		const lines: string[] = [];

		for (let r = 1; r <= lastRow; r++) {
			const cells: string[] = [];
			for (let c = 1; c <= lastCol; c++) {
				const value = this.extractValue(ws.getCell(r, c));
				let str = value === null || value === undefined ? "" : String(value);
				if (str.includes(",") || str.includes("\"") || str.includes("\n")) {
					str = `"${str.replace(/"/g, "\"\"")}"`;
				}

				cells.push(str);
			}

			lines.push(cells.join(","));
		}

		return lines.join("\n");
	}

	/**
	 * Import CSV data into a sheet starting at a given cell.
	 */
	public importFromCsv(id: string, sheet: string | number, csv: string, startAddress = "A1"): RangeData {
		const lines = csv.split(/\r?\n/).filter(l => l.length > 0);
		const data: CellValue[][] = lines.map(line => this.parseCsvLine(line));
		return this.writeRange(id, sheet, startAddress, data);
	}

	// ─── Copy ──────────────────────────────────────────────────────────────

	/**
	 * Copy a range to another location (values, styles and formats).
	 */
	public copyRange(id: string, sheet: string | number, sourceRange: string, targetStart: string): void {
		const session = this.getSession(id);
		const ws = this.resolveSheet(session, sheet);
		const src = parseRange(sourceRange);
		const dest = parseCellAddress(targetStart);

		const height = src.endRow - src.startRow + 1;
		const width = src.endCol - src.startCol + 1;

		for (let r = 0; r < height; r++) {
			for (let c = 0; c < width; c++) {
				const srcCell = ws.getCell(src.startRow + r, src.startCol + c);
				const destCell = ws.getCell(dest.row + r, dest.col + c);
				destCell.value = srcCell.value;
				destCell.style = srcCell.style;
				destCell.numFmt = srcCell.numFmt;
			}
		}

		session.dirty = true;
	}

	// ─── Internal helpers ──────────────────────────────────────────────────

	private getSession(id: string): WorkbookSession {
		const session = this.sessions.get(id);
		if (!session) {
			throw new ExcelError(`Workbook not found: ${id}. Open or create a workbook first.`);
		}

		return session;
	}

	private resolveSheet(session: WorkbookSession, sheet: string | number): Worksheet {
		let ws: Worksheet | undefined;
		if (typeof sheet === "number") {
			ws = session.workbook.worksheets[sheet];
		} else {
			ws = session.workbook.getWorksheet(sheet);
		}

		if (!ws) {
			throw new ExcelError(`Sheet "${sheet}" not found. Available: ${session.workbook.worksheets.map(w => w.name).join(", ")}`);
		}

		return ws;
	}

	private extractValue(cell: Cell): CellValue {
		const value = cell.value;
		if (value === null || value === undefined) {
			return null;
		}

		if (typeof value === "object" && !(value instanceof Date)) {
			// Formula cell: { formula, result }
			if ("formula" in value) {
				const f = value as { formula: string; result?: any };
				return f.result !== undefined && f.result !== null ? this.normalizeValue(f.result) : null;
			}

			// Rich text cell: { richText: [{ text }] }
			if ("richText" in value) {
				const rt = value as { richText: Array<{ text: string }> };
				return rt.richText.map(p => p.text).join("");
			}

			// Hyperlink cell: { text, hyperlink }
			if ("hyperlink" in value) {
				const h = value as { text?: string; hyperlink: string };
				return h.text || h.hyperlink;
			}
		}

		return this.normalizeValue(value);
	}

	private normalizeValue(value: any): CellValue {
		if (value === null || value === undefined) {
			return null;
		}

		if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
			return value;
		}

		if (value instanceof Date) {
			return value;
		}

		if (typeof value === "object" && "text" in value) {
			return String(value.text);
		}

		return String(value);
	}

	private cellToInfo(cell: Cell): CellInfo {
		const info: CellInfo = {
			address: cell.address,
			// exceljs typings type row/col as string; runtime values are numbers
			row: Number(cell.row),
			column: Number(cell.col),
			value: this.extractValue(cell),
			type: String(cell.type),
		};

		const value = cell.value;
		if (value && typeof value === "object" && !(value instanceof Date) && "formula" in value) {
			const f = value as { formula: string; result?: any };
			info.formula = f.formula;
			info.result = f.result !== undefined && f.result !== null ? this.normalizeValue(f.result) : undefined;
		}

		if (cell.numFmt) {
			info.numberFormat = cell.numFmt;
		}

		if (cell.style) {
			info.style = this.serializeStyle(cell.style);
		}

		if (cell.isMerged) {
			info.merged = true;
		}

		return info;
	}

	private serializeStyle(style: any): CellStyle {
		const result: CellStyle = {};
		if (style.font) {
			result.font = {
				name: style.font.name,
				size: style.font.size,
				bold: style.font.bold,
				italic: style.font.italic,
				underline: !!style.font.underline,
				color: style.font.color ? style.font.color.argb : undefined,
			};
		}

		if (style.fill && style.fill.fgColor) {
			result.fill = {
				type: style.fill.type,
				fgColor: style.fill.fgColor.argb,
			};
		}

		if (style.alignment) {
			result.alignment = {
				horizontal: style.alignment.horizontal,
				vertical: style.alignment.vertical,
				wrapText: style.alignment.wrapText,
				textRotation: style.alignment.textRotation,
			};
		}

		if (style.border) {
			result.border = {
				top: style.border.top,
				bottom: style.border.bottom,
				left: style.border.left,
				right: style.border.right,
			};
		}

		return result;
	}

	private mergeStyle(cell: Cell, style: CellStyle): void {
		if (style.font) {
			const current: any = cell.font || {};
			cell.font = {
				...current,
				name: style.font.name ?? current.name,
				size: style.font.size ?? current.size,
				bold: style.font.bold ?? current.bold,
				italic: style.font.italic ?? current.italic,
				underline: style.font.underline ?? current.underline,
				color: style.font.color ? { argb: style.font.color } : current.color,
			} as any;
		}

		if (style.fill && style.fill.fgColor) {
			cell.fill = {
				type: "pattern",
				pattern: "solid",
				fgColor: { argb: style.fill.fgColor },
			} as any;
		}

		if (style.alignment) {
			cell.alignment = { ...cell.alignment, ...style.alignment } as any;
		}

		if (style.border) {
			cell.border = style.border as any;
		}
	}

	private compareValues(a: any, b: any): number {
		if (a === null || a === undefined) return b === null || b === undefined ? 0 : -1;
		if (b === null || b === undefined) return 1;

		if (typeof a === "number" && typeof b === "number") {
			return a - b;
		}

		return String(a).localeCompare(String(b));
	}

	private parseCsvLine(line: string): CellValue[] {
		const result: CellValue[] = [];
		let current = "";
		let inQuotes = false;

		for (let i = 0; i < line.length; i++) {
			const char = line[i];
			if (inQuotes) {
				if (char === "\"") {
					if (line[i + 1] === "\"") {
						current += "\"";
						i++;
					} else {
						inQuotes = false;
					}
				} else {
					current += char;
				}
			} else if (char === "\"") {
				inQuotes = true;
			} else if (char === ",") {
				result.push(this.coerceCsvValue(current));
				current = "";
			} else {
				current += char;
			}
		}

		result.push(this.coerceCsvValue(current));
		return result;
	}

	private coerceCsvValue(raw: string): CellValue {
		const trimmed = raw.trim();
		if (trimmed === "") return "";
		if (!Number.isNaN(Number(trimmed))) {
			return Number(trimmed);
		}

		if (trimmed === "true") return true;
		if (trimmed === "false") return false;
		return trimmed;
	}
}
