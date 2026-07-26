#!/usr/bin/env node

import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

import { createPptxGlimpseMcpServer } from "./index.js";

async function main(): Promise<void> {
  const server = createPptxGlimpseMcpServer();
  await server.connect(new StdioServerTransport());
}

void main().catch((error: unknown) => {
  console.error("Failed to start pptx-glimpse MCP server:", error);
  process.exitCode = 1;
});
