import { defineConfig } from "tsup";

export default defineConfig({
  entry: { index: "packages/core/src/index.ts" },
  format: ["cjs", "esm"],
  dts: { resolve: ["@pptx-glimpse/renderer"] },
  clean: true,
  external: ["@pptx-glimpse/document", "@pptx-glimpse/editor"],
  noExternal: ["@pptx-glimpse/renderer"],
});
