// Core type definitions for Excel MCP

export type CellValue = string | number | boolean | null | Date;

export interface CellInfo {
	address: string;
	row: number;
	column: number;
	value: CellValue;
	type: string;
	formula?: string;
	result?: CellValue;
	numberFormat?: string;
	style?: CellStyle;
	merged?: boolean;
}

export interface CellStyle {
	font?: {
		name?: string;
		size?: number;
		bold?: boolean;
		italic?: boolean;
		underline?: boolean;
		color?: string;
	};
	fill?: {
		type?: string;
		fgColor?: string;
		bgColor?: string;
	};
	alignment?: {
		horizontal?: "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
		vertical?: "top" | "middle" | "bottom" | "distributed" | "justify";
		wrapText?: boolean;
		textRotation?: number;
	};
	border?: {
		top?: BorderStyle;
		bottom?: BorderStyle;
		left?: BorderStyle;
		right?: BorderStyle;
		diagonal?: BorderStyle;
	};
}

export interface BorderStyle {
	style?: "thin" | "dotted" | "dashDot" | "hair" | "dashDotDot" | "slantDashDot" | "mediumDashed" | "mediumDashDotDot" | "mediumDashDot" | "medium" | "double" | "thick";
	color?: string;
}

export interface SheetInfo {
	name: string;
	index: number;
	rowCount: number;
	columnCount: number;
	actualRowCount: number;
	actualColumnCount: number;
	state: "visible" | "hidden" | "veryHidden";
}

export interface WorkbookInfo {
	id: string;
	filePath: string;
	sheetCount: number;
	sheets: string[];
}

export interface RangeData {
	range: string;
	rows: number;
	cols: number;
	data: CellValue[][];
}

export interface SortConfig {
	column: number;
	order: "asc" | "desc";
}

export interface FindResult {
	cell: string;
	value: CellValue;
	row: number;
	col: number;
}

export interface ChartConfig {
	type: "line" | "bar" | "pie" | "doughnut" | "area" | "scatter";
	title?: string;
	xAxisTitle?: string;
	yAxisTitle?: string;
	series: {
		name: string;
		categoriesRange?: string;
		valuesRange: string;
	}[];
	position: string;
}

export class ExcelError extends Error {
	constructor(message: string) {
		super(message);
		this.name = "ExcelError";
	}
}
