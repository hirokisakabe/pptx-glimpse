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
    // Use a variable to prevent bundlers from statically resolving this optional import
    const specifier = "@resvg/resvg-js";
    // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
    const mod: { Resvg: ResvgConstructor } = await import(/* @vite-ignore */ specifier);
    nativeResvg = mod.Resvg;
    return nativeResvg;
  } catch {
    nativeResvg = null;
    return null;
  }
}
