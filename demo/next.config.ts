import { resolve } from "path";
import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  outputFileTracingRoot: resolve(import.meta.dirname, ".."),
  serverExternalPackages: ["pptx-glimpse", "sharp"],
  webpack: (config, { isServer }) => {
    if (isServer) {
      config.externals = [
        ...(Array.isArray(config.externals) ? config.externals : []),
        "sharp",
        "pptx-glimpse",
      ];
    }
    return config;
  },
};

export default nextConfig;
