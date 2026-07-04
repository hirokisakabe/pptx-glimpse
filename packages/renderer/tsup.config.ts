import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.ts", "src/node.ts", "src/png.ts"],
  format: ["cjs", "esm"],
  dts: true,
  clean: true,
});
