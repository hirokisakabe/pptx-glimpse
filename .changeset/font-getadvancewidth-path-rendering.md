---
"pptx-glimpse": patch
---

パスレンダリング時の幅計算を `font.getAdvanceWidth()` に切り替えた。`OpentypeFullFont` インターフェースに `getAdvanceWidth(text, fontSize)` メソッドを追加し、`renderSegmentAsPath` および `measureLineWidth` でフォントが解決できる場合は `getAdvanceWidth()` を使用するよう変更。これによりパス描画位置と幅測定の一貫性が向上する。
