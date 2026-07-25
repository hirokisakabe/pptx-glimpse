import { defineConfig } from "@playwright/test";

export default defineConfig({
  testDir: "./e2e",
  testMatch: /.*\.playwright\.ts/,
  globalSetup: "./e2e/global-setup.ts",
  timeout: 60_000,
  use: {
    browserName: "chromium",
    viewport: { width: 1280, height: 720 },
  },
});
