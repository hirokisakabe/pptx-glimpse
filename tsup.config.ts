import { defineConfig } from "tsup";

export default defineConfig({
  entry: { index: "packages/pptx-glimpse/src/index.ts" },
  format: ["cjs", "esm"],
  dts: { resolve: ["pptx-glimpse-renderer"] },
  clean: true,
  noExternal: ["pptx-glimpse-renderer"],
});
