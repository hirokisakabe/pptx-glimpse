export interface DocumentExperimentalApi {
  readonly packageName: "@pptx-glimpse/document";
  readonly status: "experimental";
}

export const documentExperimentalApi: DocumentExperimentalApi = {
  packageName: "@pptx-glimpse/document",
  status: "experimental",
};

// CleanDoc source model 型 (experimental)。後続の reader / writer / computed
// view 実装が参照する source of truth の最小 contract。
export * from "./source/index.js";

// CleanDoc source reader (experimental)。PPTX package を読み、package graph と
// presentation metadata を含む CleanDoc source を返す。
export * from "./reader/index.js";
