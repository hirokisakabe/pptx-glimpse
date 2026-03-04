import type { KnipConfig } from "knip";

const config: KnipConfig = {
  ignore: ["demo/**"],
  entry: [
    "scripts/dev-server-render.ts",
    "scripts/extract-font-metrics.ts",
    "vrt/snapshot/update-snapshots.ts",
    "bench/conversion.bench.ts",
  ],
};

export default config;
