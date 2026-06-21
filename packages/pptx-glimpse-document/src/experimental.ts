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

// CleanDoc computed view generator (experimental)。source model を mutation せず、
// slide/layout/master/theme cascade と relationship を解決した derived view を返す。
export * from "./computed/index.js";

// CleanDoc source writer (experimental)。まずは no-edit structural round-trip
// 用に、保持済み package material を PPTX ZIP として書き戻す。
export * from "./writer/index.js";
