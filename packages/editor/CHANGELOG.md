# @pptx-glimpse/editor

## 0.3.0

### Minor Changes

- 2f33268: 既存 Chart の系列名、category label、数値を chart XML と embedded workbook へ一貫して反映する typed document API と editor command / convenience API を追加する。
- 93ee164: 既存 Table のセル内 paragraph / run を、通常の shape text と同じ公開 text operation・editor command で編集して保存できるようにする。

### Patch Changes

- Updated dependencies [2f33268]
- Updated dependencies [38d59ad]
- Updated dependencies [93ee164]
  - @pptx-glimpse/document@0.13.0

## 0.2.0

### Minor Changes

- f38b91a: source node を直接渡して既存の editor command を実行できる convenience method を追加
- 8e5b47b: Unify headless editor operation failures under a shared discriminated result and expose
  `PptxEditorError` with stable operation and read/render/write codes from both high-level runtime
  entries.

### Patch Changes

- ffa4e0c: Add package-specific English README guidance and align public package metadata with the current
  rendering, document, and editing responsibilities.
- Updated dependencies [ffa4e0c]
  - @pptx-glimpse/document@0.12.1

## 0.1.0

### Minor Changes

- b5efe3d: UI 非依存の headless 編集 API を `@pptx-glimpse/editor` として公開し、browser editor に command の一括適用と ProseMirror 非依存の text body view を追加する。

> 0.0.14 以前は private package `@pptx-glimpse/editor-core` として管理されていた履歴です。

## 0.0.14

### Patch Changes

- Updated dependencies [c60442c]
  - @pptx-glimpse/document@0.12.0

## 0.0.13

### Patch Changes

- Updated dependencies [4c3cd80]
- Updated dependencies [433da8e]
  - @pptx-glimpse/document@0.11.0

## 0.0.12

### Patch Changes

- Updated dependencies [01568a2]
- Updated dependencies [5d60794]
  - @pptx-glimpse/document@0.10.0

## 0.0.11

### Patch Changes

- Updated dependencies [f16b0f5]
- Updated dependencies [79c3c55]
  - @pptx-glimpse/document@0.9.2

## 0.0.10

### Patch Changes

- Updated dependencies [4e00197]
  - @pptx-glimpse/document@0.9.1

## 0.0.9

### Patch Changes

- Updated dependencies [0034bb8]
- Updated dependencies [593b9d5]
- Updated dependencies [fd81dae]
- Updated dependencies [dd05595]
- Updated dependencies [bfb53c9]
- Updated dependencies [04b883d]
- Updated dependencies [3ad0961]
  - @pptx-glimpse/document@0.9.0

## 0.0.8

### Patch Changes

- Updated dependencies [772b4f9]
- Updated dependencies [897d4fd]
  - @pptx-glimpse/document@0.8.0

## 0.0.7

### Patch Changes

- Updated dependencies [b954ce5]
- Updated dependencies [514df94]
  - @pptx-glimpse/document@0.7.0

## 0.0.6

### Patch Changes

- Updated dependencies [f20f4b4]
- Updated dependencies [1d69ee3]
  - @pptx-glimpse/document@0.6.0

## 0.0.5

### Patch Changes

- Updated dependencies [ef9b64e]
  - @pptx-glimpse/document@0.5.0

## 0.0.4

### Patch Changes

- Updated dependencies [27fb259]
- Updated dependencies [c35eec8]
- Updated dependencies [d131e50]
- Updated dependencies [675d0f0]
- Updated dependencies [10c2b13]
  - @pptx-glimpse/document@0.4.0

## 0.0.3

### Patch Changes

- Updated dependencies [c5f2302]
- Updated dependencies [7f46470]
- Updated dependencies [c50dc1a]
  - @pptx-glimpse/document@0.3.0

## 0.0.2

### Patch Changes

- Updated dependencies [f0136a9]
- Updated dependencies [c57532b]
- Updated dependencies [f627f71]
- Updated dependencies [3d61817]
- Updated dependencies [020f949]
- Updated dependencies [69ae720]
- Updated dependencies [8904a5c]
  - @pptx-glimpse/document@0.2.0

## 0.0.1

### Patch Changes

- Updated dependencies [b32b8a8]
  - @pptx-glimpse/document@0.1.0
