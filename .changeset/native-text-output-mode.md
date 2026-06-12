---
"pptx-glimpse": minor
---

`convertPptxToSvg` に `textOutput` オプションを追加。`"text"` を指定すると、テキストをグリフのアウトライン `<path>` ではなくネイティブ `<text>` 要素 + サブセット化フォントの `@font-face`（data URI）埋め込みで出力する。ブラウザでのインライン表示時にヒンティング等のネイティブテキスト描画が効いて文字が滑らかになり、テキスト選択・コピーも可能になる。デフォルトは従来どおり `"path"` で、`convertPptxToPng` は常にパス出力で変換される。
