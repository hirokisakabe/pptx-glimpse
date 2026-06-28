import { fileURLToPath } from "node:url";

import { defineConfig } from "vitest/config";

export default defineConfig({
  resolve: {
    alias: {
      "@pptx-glimpse/document": fileURLToPath(
        new URL("./packages/document/src/index.ts", import.meta.url),
      ),
    },
  },
  test: {
    globals: true,
    include: ["vrt/**/*.test.ts"],
    testTimeout: 60000,
  },
});
