import { fileURLToPath } from "node:url";

import { defineConfig } from "vitest/config";

export default defineConfig({
  resolve: {
    alias: {
      "@pptx-glimpse/document/experimental": fileURLToPath(
        new URL("./packages/pptx-glimpse-document/src/experimental.ts", import.meta.url),
      ),
    },
  },
  test: {
    globals: true,
    include: ["vrt/**/*.test.ts"],
    testTimeout: 60000,
  },
});
