import { appendFileSync } from "node:fs";

const writeLog = (level: string, message: string) => {
	if (process.env.LOG_FILE) {
		const timestamp = new Date().toISOString();
		appendFileSync(process.env.LOG_FILE, `[${timestamp}] ${level} ${message}\n`);
	}

	console.error(message);
};

export const trace = (message: string) => {
	if (process.env.EXCEL_MCP_DEBUG) {
		writeLog("DEBUG", message);
	}
};

export const info = (message: string) => {
	writeLog("INFO", message);
};

export const warn = (message: string) => {
	writeLog("WARN", message);
};

export const error = (message: string) => {
	writeLog("ERROR", message);
};
