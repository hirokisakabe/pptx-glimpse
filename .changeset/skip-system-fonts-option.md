---
"pptx-glimpse": minor
---

`ConvertOptions` に `skipSystemFonts?: boolean` オプションを追加。`true` を指定すると OS のシステムフォントディレクトリをスキャンせず、`fontDirs` で指定したディレクトリのみを使用する。コンテナ・サーバレス環境や起動時間を短縮したい CLI ツールでの利用を想定。
