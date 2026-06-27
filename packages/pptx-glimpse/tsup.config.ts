import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.ts"],
  format: ["cjs", "esm"],
  dts: true,
  clean: true,
  noExternal: ["@pptx-glimpse/document", "pptx-glimpse-renderer"],
});
