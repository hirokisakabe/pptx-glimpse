import { fileURLToPath } from "node:url";

import { defineConfig } from "vitest/config";

export default defineConfig({
  resolve: {
    alias: {
      "@pptx-glimpse/document/experimental": fileURLToPath(new URL(
        "./packages/document/src/experimental.ts",
        import.meta.url,
      )),
    },
  },
  test: {
    globals: true,
    include: ["packages/*/src/**/*.test.ts", "e2e/**/*.test.ts"],
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
