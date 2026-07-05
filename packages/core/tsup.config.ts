import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.ts", "src/browser.ts", "src/cli.ts"],
  format: ["cjs", "esm"],
  dts: {
    resolve: [
      "@pptx-glimpse/editor-core",
      "@pptx-glimpse/renderer",
      "@pptx-glimpse/renderer/png",
      "@pptx-glimpse/renderer/png/browser",
    ],
  },
  clean: true,
  noExternal: ["@pptx-glimpse/document", "@pptx-glimpse/editor-core", "@pptx-glimpse/renderer"],
});
