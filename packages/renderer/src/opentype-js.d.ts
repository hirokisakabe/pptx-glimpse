declare module "opentype.js" {
  export function parse(buffer: ArrayBuffer): unknown;

  export class Font {
    constructor(options: Record<string, unknown>);
    toArrayBuffer(): ArrayBuffer;
  }

  export class Glyph {
    constructor(options: Record<string, unknown>);
  }
}
