declare module "rtf.js/dist/EMFJS.bundle.js" {
  export class Renderer {
    constructor(data: ArrayBuffer);
    render(settings: Record<string, string | number>): unknown;
  }

  export function loggingEnabled(enabled: boolean): void;

  const emfJs: { Renderer: typeof Renderer; loggingEnabled: typeof loggingEnabled };
  export { emfJs as EMFJS };
  export default emfJs;
}

declare module "rtf.js/dist/WMFJS.bundle.js" {
  export class Renderer {
    constructor(data: ArrayBuffer);
    render(settings: Record<string, string | number>): unknown;
  }

  export function loggingEnabled(enabled: boolean): void;

  const wmfJs: { Renderer: typeof Renderer; loggingEnabled: typeof loggingEnabled };
  export { wmfJs as WMFJS };
  export default wmfJs;
}
