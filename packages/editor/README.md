# @pptx-glimpse/editor

UI framework に依存しない `PptxSourceModel` の編集層です。`@pptx-glimpse/document`
で PPTX を読み書きし、この package で command の検証・適用、selection、undo / redo、
warning の取得を行います。

`@pptx-glimpse/editor` は 0.x です。API は独立して利用できますが、command と結果型は
1.0 までに変更される可能性があります。

## インストール

```sh
npm install @pptx-glimpse/document @pptx-glimpse/editor
```

## 最小例

```ts
import { readPptx, writePptx } from "@pptx-glimpse/document";
import { createEditorSession } from "@pptx-glimpse/editor";

const source = readPptx(input);
const run = source.slides[0]?.shapes
  .find((shape) => shape.kind === "shape")
  ?.textBody?.paragraphs[0]?.runs[0];

if (run?.handle === undefined) {
  throw new Error("editable text run not found");
}

const session = createEditorSession(source);
const result = session.apply({
  kind: "replaceTextRunPlainText",
  handle: run.handle,
  text: "Edited",
});
if (!result.ok) {
  throw new Error(result.message);
}

const output = writePptx(session.document);
```

複数 command は `session.applyAll(commands)` で一つの undo 履歴として適用できます。
`selectShape()` / `deselectShape()`、`undo()` / `redo()`、`canUndo` / `canRedo` も
session から利用できます。

## 対応 command

- text: `replaceTextRunPlainText`, `replaceParagraphPlainText`,
  `setTextRunProperties`, `clearTextRunProperties`, `setParagraphProperties`,
  `clearParagraphProperties`
- shape: `moveShape`, `resizeShape`, `setShapeTransform`, `setShapeFill`,
  `setShapeOutline`, `addTextBox`, `addConnector`, `deleteShape`
- media: `replaceImage`
- slide: `addEmptySlideFromLayout`, `duplicateSlide`, `moveSlide`, `deleteSlide`

## 既知の制約

- source handle がない要素や、document 層が安全に保持できない編集は
  `invalid-command` として拒否されます。
- table、chart、SmartArt の新規編集 command はありません。
- `replaceImage` が共有 media part を更新する場合、参照する複数画像が同時に変わるため
  `shared-media-part` warning を返します。
- DOM、React、ProseMirror、renderer は含みません。描画を伴う browser editor facade は
  `pptx-glimpse` の browser entry を利用してください。
