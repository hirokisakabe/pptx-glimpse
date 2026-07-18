export interface SupportedImageType {
  readonly contentType: "image/png" | "image/jpeg";
  readonly extension: "png" | "jpeg";
}

export function detectSupportedImageType(bytes: Uint8Array): SupportedImageType | undefined {
  if (startsWithBytes(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) {
    return { contentType: "image/png", extension: "png" };
  }
  if (startsWithBytes(bytes, [0xff, 0xd8, 0xff])) {
    return { contentType: "image/jpeg", extension: "jpeg" };
  }
  return undefined;
}

export function startsWithBytes(bytes: Uint8Array, prefix: readonly number[]): boolean {
  if (bytes.length < prefix.length) return false;
  return prefix.every((value, index) => bytes[index] === value);
}
