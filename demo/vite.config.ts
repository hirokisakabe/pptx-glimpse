import { defineConfig } from "vite";

export default defineConfig({
  root: __dirname,
  base: "/pptx-glimpse/",
  build: {
    outDir: "dist",
    target: "es2022",
  },
  optimizeDeps: {
    include: ["jszip", "fast-xml-parser"],
  },
});
