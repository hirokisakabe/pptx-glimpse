/**
 * ネイティブ <text> 出力モード用のフォント使用状況コレクタ。
 * レンダリング中に「フォント名 → 使用文字集合」を収集し、
 * サブセット化フォントの @font-face 埋め込みに利用する。
 */

export interface FontUsage {
  /** resolveFont に渡すフォント名の優先順リスト (tspan の font-family と同順) */
  fonts: (string | null)[];
  /** このフォントで描画される文字の集合 */
  chars: Set<string>;
}

export class FontUsageCollector {
  /** key: tspan の font-family リスト先頭のフォント名 (= @font-face で宣言する名前) */
  private usages = new Map<string, FontUsage>();

  record(fonts: (string | null)[], text: string): void {
    const primary = fonts.find((f) => f !== null && f !== undefined);
    if (!primary || text.length === 0) return;

    let usage = this.usages.get(primary);
    if (!usage) {
      usage = { fonts: [...fonts], chars: new Set() };
      this.usages.set(primary, usage);
    }
    for (const char of text) {
      usage.chars.add(char);
    }
  }

  getUsages(): Map<string, FontUsage> {
    return this.usages;
  }

  reset(): void {
    this.usages.clear();
  }
}

let currentCollector: FontUsageCollector | null = null;

export function setFontUsageCollector(collector: FontUsageCollector): void {
  currentCollector = collector;
}

export function getFontUsageCollector(): FontUsageCollector | null {
  return currentCollector;
}

export function resetFontUsageCollector(): void {
  currentCollector = null;
}
