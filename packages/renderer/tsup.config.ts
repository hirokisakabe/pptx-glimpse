import { defineConfig } from "tsup";

export default defineConfig({
  entry: ["src/index.ts", "src/node.ts"],
  format: ["cjs", "esm"],
  dts: true,
  clean: true,
});
