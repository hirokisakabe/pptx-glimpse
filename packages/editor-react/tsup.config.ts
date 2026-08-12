import { copyFile } from "node:fs/promises";

import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.ts"],
  format: ["esm"],
  dts: true,
  clean: true,
  external: ["pptx-glimpse", "react", "react-dom"],
  async onSuccess() {
    await copyFile("src/styles.css", "dist/styles.css");
  },
});
