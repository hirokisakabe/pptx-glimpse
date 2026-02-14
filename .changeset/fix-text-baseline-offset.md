---
"pptx-glimpse": patch
---

fix: テキスト1行目のベースライン計算で ascender のみを使用するように修正

SVG の `<text y="...">` はベースライン位置を指定するが、従来は `lineHeightRatio`（ascender + descender）をオフセットとして使用していたため、descender 分だけテキストが下にずれて表示されていた。`getAscenderRatio()` を追加し、1行目のベースラインオフセット計算で使用するように変更。
