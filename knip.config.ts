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
      ],
      ignoreDependencies: [
        "pptx-glimpse-renderer",
        "opentype.js",
        "fast-xml-parser",
        "sharp",
      ],
    },
    "packages/pptx-glimpse": {},
    "packages/pptx-glimpse-renderer": {},
  },
};

export default config;
