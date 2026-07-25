import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.ts", "src/browser.ts", "src/cli.ts"],
  format: ["cjs", "esm"],
  dts: {
    resolve: [
      "@pptx-glimpse/renderer",
      "@pptx-glimpse/renderer/png",
      "@pptx-glimpse/renderer/png/browser",
    ],
  },
  clean: true,
  external: ["@pptx-glimpse/document", "@pptx-glimpse/editor"],
  noExternal: ["@pptx-glimpse/renderer"],
});
