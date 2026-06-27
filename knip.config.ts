import type { KnipConfig } from "knip";

const config: KnipConfig = {
  ignore: ["demo/**"],
  ignoreDependencies: [
    // root tsup が packages/pptx-glimpse/src/index.ts を bundle した output が runtime で参照する。
    // root 側のコードからは直接 import されないが、公開パッケージの runtime dep として必要。
    "@resvg/resvg-wasm",
    "fflate",
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
    "packages/pptx-glimpse": {
      ignoreDependencies: [
        // Core currently dogfoods document through workspace source imports.
        // Keep the manifest dependency explicit until package imports migrate.
        "@pptx-glimpse/document",
      ],
    },
  },
};

export default config;
