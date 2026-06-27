import type { ConvertOptions } from "../../packages/core/src/converter.js";

// Local snapshot VRT must not depend on parsing every OS system font. The Docker
// snapshot job and local VRT both use this option so snapshots are generated and
// verified under the same font-loading policy.
export const VRT_RENDER_OPTIONS = {
  skipSystemFonts: true,
} as const satisfies Pick<ConvertOptions, "skipSystemFonts">;
