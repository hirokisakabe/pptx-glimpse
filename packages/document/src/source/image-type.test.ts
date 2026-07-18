import { describe, expect, it } from "vitest";

import { detectSupportedImageType } from "./image-type.js";

describe("detectSupportedImageType", () => {
  it("detects PNG and JPEG magic bytes with their content types and extensions", () => {
    expect(
      detectSupportedImageType(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])),
    ).toEqual({ contentType: "image/png", extension: "png" });
    expect(detectSupportedImageType(new Uint8Array([0xff, 0xd8, 0xff]))).toEqual({
      contentType: "image/jpeg",
      extension: "jpeg",
    });
  });

  it("rejects unsupported and truncated magic bytes", () => {
    expect(detectSupportedImageType(new Uint8Array([0x47, 0x49, 0x46]))).toBeUndefined();
    expect(
      detectSupportedImageType(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a])),
    ).toBeUndefined();
    expect(detectSupportedImageType(new Uint8Array([0xff, 0xd8]))).toBeUndefined();
  });
});
