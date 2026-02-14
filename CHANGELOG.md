# pptx-glimpse

## 0.1.4

### Patch Changes

- fcb6a4e: fix: テーブルスタイル参照時のデフォルトボーダーフォールバックを追加

  tableStyleId が指定されたテーブルでセルにインラインボーダーが定義されていない場合に、デフォルトの黒枠線を適用するようにした。Google Slides や PowerPoint で作成された PPTX ファイルのテーブル枠線が正しくレンダリングされるようになる。

## 0.1.3

### Patch Changes

- 133e7b6: fix: テキスト1行目のベースライン計算で ascender のみを使用するように修正

  SVG の `<text y="...">` はベースライン位置を指定するが、従来は `lineHeightRatio`（ascender + descender）をオフセットとして使用していたため、descender 分だけテキストが下にずれて表示されていた。`getAscenderRatio()` を追加し、1行目のベースラインオフセット計算で使用するように変更。
