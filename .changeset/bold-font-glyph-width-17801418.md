---
"pptx-glimpse": patch
---

bold テキストの幅計算でボールド体フォント（`"${fontFamily} Bold"` / `"${fontFamily}-Bold"` の命名規則）がフォントマップに登録されている場合、固定係数 `BOLD_FACTOR (1.05)` ではなく実グリフの `advanceWidth` を使用するよう改善しました。フォントマッピング（例: `Calibri → Carlito`）経由での Bold 解決にも対応します。ボールド体フォントが未登録の場合は従来の `BOLD_FACTOR` フォールバックが引き続き動作します。
