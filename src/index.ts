#!/usr/bin/env node
import { createMcpExpressApp } from "@modelcontextprotocol/express";
import { toNodeHandler } from "@modelcontextprotocol/node";
import { createMcpHandler } from "@modelcontextprotocol/server";
import { serveStdio } from "@modelcontextprotocol/server/stdio";
import { Request, Response } from "express";
import { program } from "commander";

import { createMcpServer, getAgentVersion } from "./server";
import { error } from "./logger";

const logMcpError = (err: Error) => {
	error("Error handling MCP request: " + (err.stack ?? String(err)));
};

const startHttpServer = async (host: string, port: number) => {
	const app = createMcpExpressApp({ host });

	const authToken = process.env.EXCEL_MCP_AUTH;
	if (!authToken) {
		error("WARNING: EXCEL_MCP_AUTH is not set. The HTTP server will accept unauthenticated connections.");
	}

	if (authToken) {
		app.use((req, res, next) => {
			if (req.headers.authorization !== `Bearer ${authToken}`) {
				res.status(401).json({ error: "Unauthorized" });
				return;
			}

			next();
		});
	}

	const handler = createMcpHandler(() => createMcpServer(), { onerror: logMcpError });
	const node = toNodeHandler(handler, { onerror: logMcpError });

	app.post("/mcp", (req: Request, res: Response) => {
		void node(req, res, req.body);
	});

	const methodNotAllowed = (_req: Request, res: Response) => {
		res.status(405).json({
			jsonrpc: "2.0",
			error: { code: -32000, message: "Method not allowed." },
			id: null,
		});
	};

	app.get("/mcp", methodNotAllowed);
	app.delete("/mcp", methodNotAllowed);

	app.listen(port, host, () => {
		error(`excel-mcp ${getAgentVersion()} streamable http server listening on http://${host}:${port}/mcp`);
	});
};

const startStdioServer = async () => {
	try {
		serveStdio(() => createMcpServer(), { onerror: logMcpError });

		const shutdown = () => process.exit(0);
		process.on("SIGINT", shutdown);
		process.on("SIGTERM", shutdown);

		error(`excel-mcp ${getAgentVersion()} server running on stdio`);
	} catch (err: any) {
		console.error("Fatal error in main():", err);
		error("Fatal error in main(): " + JSON.stringify(err.stack));
		process.exit(1);
	}
};

const main = async () => {
	program
		.version(getAgentVersion())
		.option("--listen <listen>", "Start Streamable HTTP server on [host:]port")
		.option("--stdio", "Start stdio server (default)")
		.parse(process.argv);

	const options = program.opts();

	if (options.listen) {
		const listen = (options.listen as string).trim();
		const lastColon = listen.lastIndexOf(":");
		let host = "localhost";
		let rawPort: string;

		if (lastColon > 0) {
			host = listen.substring(0, lastColon);
			rawPort = listen.substring(lastColon + 1);
		} else {
			rawPort = listen;
		}

		const port = Number.parseInt(rawPort, 10);
		if (!host || !rawPort || !Number.isInteger(port) || port < 1 || port > 65535) {
			error(`Invalid --listen value "${listen}". Expected [host:]port with port 1-65535.`);
			process.exit(1);
		}

		await startHttpServer(host, port);
	} else {
		await startStdioServer();
	}
};

main().then();
