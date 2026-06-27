import type { KnipConfig } from "knip";

const config: KnipConfig = {
  ignore: ["demo/**"],
  ignoreDependencies: [
    // core bundles @pptx-glimpse/renderer, whose PNG path loads this runtime dependency.
    "@resvg/resvg-wasm",
  ],
  workspaces: {
    ".": {
      entry: [
        "scripts/dev-server-render.ts",
        "scripts/extract-font-metrics.ts",
        "vrt/snapshot/update-snapshots.ts",
        "bench/conversion.bench.ts",
      ],
    },
  },
};

export default config;
