import { resolve } from "node:path";
import type { NextConfig } from "next";
import nextra from "nextra";

const withNextra = nextra({
  contentDirBasePath: "/docs/api",
});

const nextConfig: NextConfig = {
  outputFileTracingRoot: resolve(import.meta.dirname, ".."),
  webpack: (config) => {
    config.resolve = config.resolve ?? {};
    config.resolve.alias = {
      ...config.resolve.alias,
      "pptx-glimpse$": resolve(import.meta.dirname, "../packages/core/dist/browser.js"),
    };
    return config;
  },
};

export default withNextra(nextConfig);
