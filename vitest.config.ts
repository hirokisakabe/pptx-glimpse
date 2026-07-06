import { fileURLToPath } from "node:url";

import { defineConfig } from "vitest/config";

export default defineConfig({
  resolve: {
    alias: {
      "@pptx-glimpse/document": fileURLToPath(
        new URL("./packages/document/src/index.ts", import.meta.url),
      ),
      "@pptx-glimpse/editor-core": fileURLToPath(
        new URL("./packages/editor-core/src/index.ts", import.meta.url),
      ),
    },
  },
  test: {
    globals: true,
    include: [
      "packages/*/src/**/*.test.ts",
      "e2e/**/*.test.ts",
      "vrt/editor-validity/editor-validity.test.ts",
    ],
    testTimeout: 30000,
    coverage: {
      provider: "v8",
      include: ["packages/*/src/**/*.ts"],
      exclude: ["packages/*/src/**/*.test.ts"],
      reporter: ["text", "html", "json-summary"],
    },
    benchmark: {
      include: ["bench/**/*.bench.ts"],
    },
  },
});
