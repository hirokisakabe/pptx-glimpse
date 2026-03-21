# pptx-glimpse

## 0.4.1

### Patch Changes

- 31b78b1: roundRect 系プリセットジオメトリの adj 値を OOXML 仕様通り 0〜50000 にクランプするよう修正

## 0.4.0

### Minor Changes

- 40f4dbc: メディアファイルの遅延読み込みを実装。PPTX 内の画像データを初期読み込み時に全展開せず、描画時に必要なファイルだけを解凍するように変更。大規模ファイルのメモリ使用量を削減。

## 0.3.1

### Patch Changes

- 62e5eb8: fix: normAutofit の defaultFontSize 二重スケーリングを修正し、テキスト自動縮小が正しく動作するよう改善

## 0.3.0

### Minor Changes

- 7330871: TTC (TrueType Collection) フォントの fontBuffers 対応を追加。TTC バッファを自動的に個別の TTF/OTF に分割し、opentype.js でパースできるようにした。

## 0.2.2

### Patch Changes

- e600c05: パース/レンダリング失敗時のサイレントスキップ箇所に debug レベルの警告を追加

## 0.2.1

### Patch Changes

- 1060113: CJK 文字に対する BOLD_FACTOR 適用を除外し、太字 CJK テキストの不要な折り返しを修正

## 0.2.0

### Minor Changes

- b29fe5c: eaVert モードで CJK 文字を直立表示するように対応
- a428791: 株価チャート (stockChart)、サーフェスチャート (surfaceChart/surface3DChart)、ofPieチャート (ofPieChart) のパース・レンダリングに対応

### Patch Changes

- 69101bd: ctrTitle/subTitle プレースホルダーがマスターの title/body lstStyle からテキストカラーを正しく継承するよう修正

## 0.1.4

### Patch Changes

- fcb6a4e: fix: テーブルスタイル参照時のデフォルトボーダーフォールバックを追加

  tableStyleId が指定されたテーブルでセルにインラインボーダーが定義されていない場合に、デフォルトの黒枠線を適用するようにした。Google Slides や PowerPoint で作成された PPTX ファイルのテーブル枠線が正しくレンダリングされるようになる。

## 0.1.3

### Patch Changes

- 133e7b6: fix: テキスト1行目のベースライン計算で ascender のみを使用するように修正

  SVG の `<text y="...">` はベースライン位置を指定するが、従来は `lineHeightRatio`（ascender + descender）をオフセットとして使用していたため、descender 分だけテキストが下にずれて表示されていた。`getAscenderRatio()` を追加し、1行目のベースラインオフセット計算で使用するように変更。
