import { defineConfig } from "tsup";

export default defineConfig({
  entry: { index: "packages/pptx-glimpse/src/index.ts" },
  format: ["cjs", "esm"],
  dts: true,
  clean: true,
  noExternal: ["pptx-glimpse-renderer"],
});
