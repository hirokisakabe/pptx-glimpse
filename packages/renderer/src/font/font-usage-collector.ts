/**
 * Internal note.
 * Internal note.
 * Internal note.
 */

export interface FontUsage {
  /** Internal note. */
  fonts: (string | null)[];
  /** Internal note. */
  chars: Set<string>;
}

export class FontUsageCollector {
  /** Internal note. */
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
