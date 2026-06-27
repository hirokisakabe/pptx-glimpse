import { defineConfig } from "tsup";

export default defineConfig({
  entry: { index: "packages/core/src/index.ts" },
  format: ["cjs", "esm"],
  dts: { resolve: ["@pptx-glimpse/renderer"] },
  clean: true,
  noExternal: ["@pptx-glimpse/document", "@pptx-glimpse/renderer"],
});
