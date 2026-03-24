import { resolve } from "path";
import { fileURLToPath } from "url";
import { defineConfig } from "vitest/config";

const __dirname = fileURLToPath(new URL(".", import.meta.url));

export default defineConfig({
  resolve: {
    alias: {
      "pptx-glimpse-renderer": resolve(__dirname, "packages/pptx-glimpse-renderer/src/index.ts"),
    },
  },
  test: {
    globals: true,
    include: ["vrt/**/*.test.ts"],
    testTimeout: 60000,
  },
});
