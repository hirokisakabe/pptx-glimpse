import type { KnipConfig } from "knip";

const config: KnipConfig = {
  ignore: ["demo/**"],
  workspaces: {
    ".": {
      entry: [
        "scripts/dev-server-render.ts",
        "scripts/extract-font-metrics.ts",
        "vrt/snapshot/update-snapshots.ts",
        "bench/conversion.bench.ts",
        "e2e/dev-server-editor.playwright.ts",
        "e2e/browser-standalone-viewer.playwright.ts",
      ],
    },
    "packages/core": {
      ignoreDependencies: ["@resvg/resvg-wasm"],
    },
  },
};

export default config;
