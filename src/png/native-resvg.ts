export interface ResvgInstance {
  render(): RenderedImage;
  free?(): void;
}

export interface RenderedImage {
  asPng(): Uint8Array;
  width: number;
  height: number;
  free?(): void;
}

export type ResvgConstructor = new (
  svg: string,
  options?: Record<string, unknown>,
) => ResvgInstance;

let nativeResvg: ResvgConstructor | null | undefined; // undefined = not yet tried

export async function tryLoadNativeResvg(): Promise<ResvgConstructor | null> {
  if (nativeResvg !== undefined) return nativeResvg;
  try {
    const mod = await import("@resvg/resvg-js");
    nativeResvg = mod.Resvg as ResvgConstructor;
    return nativeResvg;
  } catch {
    nativeResvg = null;
    return null;
  }
}
