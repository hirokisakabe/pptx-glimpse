import { resolve } from "node:path";
import type { NextConfig } from "next";
import nextra from "nextra";

const withNextra = nextra({
  contentDirBasePath: "/docs/api",
});

const nextConfig: NextConfig = {
  outputFileTracingRoot: resolve(import.meta.dirname, ".."),
  turbopack: {
    resolveAlias: {
      "next-mdx-import-source-file": "./mdx-components.tsx",
      "pptx-glimpse": "../packages/core/dist/browser.js",
    },
  },
};

export default withNextra(nextConfig);
