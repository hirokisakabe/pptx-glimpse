# pptx-glimpse

## 0.11.1

### Patch Changes

- 237b32d: パスレンダリング時の幅計算を `font.getAdvanceWidth()` に切り替えた。`OpentypeFullFont` インターフェースに `getAdvanceWidth(text, fontSize)` メソッドを追加し、`renderSegmentAsPath` および `measureLineWidth` でフォントが解決できる場合は `getAdvanceWidth()` を使用するよう変更。これによりパス描画位置と幅測定の一貫性が向上する。
- 95c262a: OOXML Strict 形式のSmartArt（Diagram）URIに対応。Transitional URI（`http://schemas.openxmlformats.org/drawingml/2006/diagram`）のみを直接比較していた判定を、Strict URI（`http://purl.oclc.org/ooxml/drawingml/diagram`）にも対応した allowlist 判定に変更。

## 0.11.0

### Minor Changes

- 85f0cfa: `ConvertOptions` に `skipSystemFonts?: boolean` オプションを追加。`true` を指定すると OS のシステムフォントディレクトリをスキャンせず、`fontDirs` で指定したディレクトリのみを使用する。コンテナ・サーバレス環境や起動時間を短縮したい CLI ツールでの利用を想定。

## 0.10.4

### Patch Changes

- ca117e4: スライド自身の placeholder のうち TextBody の全 run のテキストが空のものを描画対象から除外。テンプレート PPTX で未入力のプレースホルダー枠線・背景が出力 SVG に残ってしまう不具合を修正した。

## 0.10.3

### Patch Changes

- 248e1ba: 異なるフォーマットのテキストラン間でスペースが消える問題を修正。XML パーサーのテキスト値トリムを無効化し、SVG の tspan モードで xml:space="preserve" を設定。

## 0.10.2

### Patch Changes

- 523356e: 単一の段落要素内に複数の段落プロパティが交互配置される非標準PPTXで箇条書きが正しくレンダリングされない問題を修正

## 0.10.1

### Patch Changes

- bbb366e: fix: buFont未指定時に箇条書き記号がテキストランのフォントにフォールバックするよう修正
- efbc301: fix: プレースホルダの箇条書きスタイル（bullet, marginLeft, indent）がスライドマスターのtxStylesから正しく継承されるよう修正

## 0.10.0

### Minor Changes

- e78f34e: feat: フォントオブジェクトキャッシュの導入 - convertPptxToSvg の2回目以降の呼び出しでフォント読み込みをスキップし高速化

## 0.9.0

### Minor Changes

- 61a0fba: CJKフォントが見つからない場合のOS別フォールバックチェーンを追加し、font.notFound警告を出力するようにした

## 0.8.0

### Minor Changes

- d9bb29b: スライドが0枚のPPTXファイルで `presentation.noSlides` 警告を出力するように追加

## 0.7.1

### Patch Changes

- 87a455b: Carlito フォントの CJK advance width 不整合によるテキスト重なりを修正

## 0.7.0

### Minor Changes

- 37b86de: SVG→PNG変換を@resvg/resvg-js（ネイティブアドオン）から@resvg/resvg-wasm（WASM）に移行し、esbuild等でのバンドルを可能にした

## 0.6.1

### Patch Changes

- 90bb744: feat: プレースホルダー図形の位置・サイズ継承に対応

## 0.6.0

### Minor Changes

- 5906f6d: SVG→PNG変換のsharpを@resvg/resvg-jsに置き換え、ネイティブモジュール依存を排除

## 0.5.0

### Minor Changes

- 334a370: チャートの軸ラベル・目盛りを描画するよう改善
  - 棒グラフ・折れ線グラフ・エリアチャート・散布図・バブルチャート・株価チャートにY軸数値ラベルを追加
  - レーダーチャートの頂点ラベルが表示されない問題を修正（multiLvlStrRef 対応）

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
