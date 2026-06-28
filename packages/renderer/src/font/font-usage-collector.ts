/**
 * Font usage collector for native <text> output mode.
 * Collect "font name -> character set used" during rendering,
 * Used for @font-face embedding of subsetted fonts.
 */

export interface FontUsage {
  /** Priority list of font names to be passed to resolveFont (same order as tspan's font-family) */
  fonts: (string | null)[];
  /** the set of characters drawn with this font */
  chars: Set<string>;
}

export class FontUsageCollector {
  /** key: font name at the beginning of tspan's font-family list (= name declared with @font-face) */
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
