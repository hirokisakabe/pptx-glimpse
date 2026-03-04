import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    globals: true,
    include: ["vrt/**/*.test.ts"],
    testTimeout: 60000,
  },
});
