---
"pptx-glimpse": major
---

非推奨の ProseMirror text-body 互換 API（`PptxEditorShapeInfo.editableTextBody`、`PptxEditorTextBodyInfo`、`PptxEditorSession.applyTextBodyDocJson()`）と `prosemirror-model` 依存を削除する。テキスト編集には `PptxEditorShapeInfo.textBody` と `PptxEditorSession.applyAll()` を使用してください。
