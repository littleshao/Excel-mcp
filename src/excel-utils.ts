// Excel utility functions: column letter conversions, range parsing, etc.

/**
 * Convert a column number (1-based) to a column letter (A, B, ..., Z, AA, AB, ...).
 */
export const columnToLetter = (col: number): string => {
	if (col < 1) {
		throw new Error(`Column number must be >= 1, got ${col}`);
	}

	let result = "";
	let n = col;
	while (n > 0) {
		const remainder = (n - 1) % 26;
		result = String.fromCharCode(65 + remainder) + result;
		n = Math.floor((n - 1) / 26);
	}

	return result;
};

/**
 * Convert a column letter to a column number (1-based).
 */
export const letterToColumn = (letter: string): number => {
	const upper = letter.toUpperCase();
	let result = 0;
	for (let i = 0; i < upper.length; i++) {
		const charCode = upper.charCodeAt(i);
		if (charCode < 65 || charCode > 90) {
			throw new Error(`Invalid column letter: "${letter}"`);
		}

		result = result * 26 + (charCode - 64);
	}

	return result;
};

/**
 * Parse a cell address like "A1", "AB12" into { row, column } (1-based).
 */
export const parseCellAddress = (address: string): { row: number; col: number } => {
	const match = /^([A-Za-z]+)(\d+)$/.exec(address.trim());
	if (!match) {
		throw new Error(`Invalid cell address: "${address}"`);
	}

	return {
		col: letterToColumn(match[1]),
		row: Number.parseInt(match[2], 10),
	};
};

/**
 * Build a cell address from row and column (1-based).
 */
export const cellAddress = (row: number, col: number): string => {
	return `${columnToLetter(col)}${row}`;
};

/**
 * Parse a range like "A1:B5" or "Sheet1!A1:B5" into bounds.
 * Returns { sheet, startRow, startCol, endRow, endCol } all 1-based.
 */
export const parseRange = (range: string): {
	sheet?: string;
	startRow: number;
	startCol: number;
	endRow: number;
	endCol: number;
} => {
	const trimmed = range.trim();

	let sheet: string | undefined;
	let rangePart = trimmed;

	const bangIndex = trimmed.indexOf("!");
	if (bangIndex !== -1) {
		sheet = trimmed.substring(0, bangIndex).replace(/^['"]|['"]$/g, "");
		rangePart = trimmed.substring(bangIndex + 1);
	}

	const parts = rangePart.split(":");
	if (parts.length !== 2) {
		// single cell
		const { row, col } = parseCellAddress(rangePart);
		return { sheet, startRow: row, startCol: col, endRow: row, endCol: col };
	}

	const start = parseCellAddress(parts[0]);
	const end = parseCellAddress(parts[1]);

	return {
		sheet,
		startRow: Math.min(start.row, end.row),
		startCol: Math.min(start.col, end.col),
		endRow: Math.max(start.row, end.row),
		endCol: Math.max(start.col, end.col),
	};
};

/**
 * Build a range string from bounds (1-based).
 */
export const buildRange = (startRow: number, startCol: number, endRow: number, endCol: number): string => {
	return `${cellAddress(startRow, startCol)}:${cellAddress(endRow, endCol)}`;
};

/**
 * Check if a file path has a valid Excel extension.
 */
export const isValidExcelExtension = (filePath: string): boolean => {
	const lower = filePath.toLowerCase();
	return lower.endsWith(".xlsx") || lower.endsWith(".xls") || lower.endsWith(".xlsm") || lower.endsWith(".csv");
};

/**
 * Determine if a value is numeric (for number formatting decisions).
 */
export const isNumeric = (value: unknown): boolean => {
	if (typeof value === "number") {
		return Number.isFinite(value);
	}

	if (typeof value === "string" && value.trim() !== "") {
		return !Number.isNaN(Number(value));
	}

	return false;
};

/**
 * Serialize a cell value to a JSON-friendly representation.
 */
export const serializeValue = (value: unknown): unknown => {
	if (value instanceof Date) {
		return { __type: "date", value: value.toISOString() };
	}

	return value;
};
