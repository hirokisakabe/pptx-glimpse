import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";

import { previewPptx } from "./preview-pptx.js";
import { PACKAGE_VERSION } from "./version.js";

const supportCoverageCountsSchema = z.object({
  inputElements: z.number().int().nonnegative(),
  outputElements: z.number().int().nonnegative(),
  skippedElements: z.number().int().nonnegative(),
  unresolvedElements: z.number().int().nonnegative(),
  fallbackElements: z.number().int().nonnegative(),
  warnings: z.number().int().nonnegative(),
});

export function createPptxGlimpseMcpServer(): McpServer {
  const server = new McpServer({
    name: "pptx-glimpse",
    version: PACKAGE_VERSION,
  });

  server.registerTool(
    "preview_pptx",
    {
      title: "Preview PPTX slides",
      description:
        "Convert all or selected slides from a local PPTX file to PNG image content. Returns slide-to-content mappings, diagnostics, and support coverage as structured content.",
      inputSchema: z.object({
        filePath: z.string().min(1).describe("Path to the local PPTX file"),
        slides: z
          .array(z.number().int().positive())
          .min(1)
          .optional()
          .describe("Optional 1-based slide numbers to preview"),
      }),
      outputSchema: z.object({
        slides: z.array(
          z.object({
            slideNumber: z.number().int().positive(),
            contentIndex: z.number().int().nonnegative(),
            width: z.number().int().positive(),
            height: z.number().int().positive(),
          }),
        ),
        diagnostics: z.object({
          total: z.number().int().nonnegative(),
          info: z.number().int().nonnegative(),
          warnings: z.number().int().nonnegative(),
          errors: z.number().int().nonnegative(),
        }),
        supportCoverage: z.object({
          overall: supportCoverageCountsSchema,
          slides: z.array(
            supportCoverageCountsSchema.extend({
              slideNumber: z.number().int().positive(),
            }),
          ),
        }),
      }),
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    (input) => previewPptx(input),
  );

  return server;
}

export type {
  PreviewPptxDependencies,
  PreviewPptxInput,
  PreviewPptxStructuredContent,
} from "./preview-pptx.js";
export { previewPptx } from "./preview-pptx.js";
