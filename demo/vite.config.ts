import { defineConfig } from "vite";
import path from "path";

export default defineConfig({
  root: __dirname,
  base: "/pptx-glimpse/",
  resolve: {
    alias: {
      sharp: path.resolve(__dirname, "src/stubs/sharp.ts"),
    },
  },
  define: {
    global: "globalThis",
  },
  build: {
    outDir: "dist",
    target: "es2022",
  },
  optimizeDeps: {
    include: ["buffer", "jszip", "fast-xml-parser"],
  },
});
