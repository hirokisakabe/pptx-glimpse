import type { KnipConfig } from "knip";

const config: KnipConfig = {
  ignore: ["demo/**"],
  ignoreBinaries: ["rg"],
  workspaces: {
    ".": {
      entry: [
        "scripts/dev-server-render.ts",
        "scripts/extract-font-metrics.ts",
        "vrt/snapshot/update-snapshots.ts",
        "e2e/dev-server-editor.playwright.ts",
        "e2e/browser-standalone-viewer.playwright.ts",
        "e2e/browser-standalone-editor.playwright.ts",
        "e2e/demo-browser-editor.playwright.ts",
      ],
      ignoreDependencies: ["typedoc-plugin-markdown"],
    },
    "packages/core": {
      ignoreDependencies: ["@resvg/resvg-wasm"],
    },
  },
};

export default config;
