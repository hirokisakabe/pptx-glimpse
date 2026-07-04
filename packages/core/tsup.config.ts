import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.ts", "src/browser.ts"],
  format: ["cjs", "esm"],
  dts: { resolve: ["@pptx-glimpse/renderer"] },
  clean: true,
  noExternal: ["@pptx-glimpse/document", "@pptx-glimpse/renderer"],
});
