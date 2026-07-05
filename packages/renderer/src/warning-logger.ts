/**
 * Controls how unsupported-feature warnings are collected and printed.
 *
 * `"off"` disables collection, `"warn"` collects entries and prints summaries,
 * and `"debug"` also prints each warning as it is recorded.
 */
export type LogLevel = "off" | "warn" | "debug";

/**
 * One unsupported or approximated feature encountered during conversion.
 */
export interface WarningEntry {
  /**
   * Stable feature key, for example `"sp.style"` or `"bodyPr@vert"`.
   */
  feature: string;
  /**
   * Human-readable warning message.
   */
  message: string;
  /**
   * Optional location context, for example `"Slide 1"`.
   */
  context?: string;
}

/**
 * Aggregated warnings from the most recent conversion cycle.
 */
export interface WarningSummary {
  /**
   * Total number of warning entries.
   */
  totalCount: number;
  /**
   * Warning counts grouped by feature key.
   */
  features: { feature: string; message: string; count: number }[];
}

export interface WarningLogger {
  warn(feature: string, message: string, context?: string): void;
  debug(feature: string, message: string, context?: string): void;
  getWarningSummary(): WarningSummary;
  flushWarnings(): WarningSummary;
  getWarningEntries(): readonly WarningEntry[];
  getLogLevel(): LogLevel;
}

const PREFIX = "[pptx-glimpse]";

class InMemoryWarningLogger implements WarningLogger {
  private entries: WarningEntry[] = [];
  private readonly featureCounts = new Map<string, { message: string; count: number }>();

  constructor(private readonly level: LogLevel) {}

  warn(feature: string, message: string, context?: string): void {
    if (this.level === "off") return;
    this.record(feature, message, context);

    if (this.level === "debug") {
      const ctx = context ? ` (${context})` : "";
      console.warn(`${PREFIX} SKIP: ${feature} - ${message}${ctx}`);
    }
  }

  debug(feature: string, message: string, context?: string): void {
    if (this.level !== "debug") return;
    this.record(feature, message, context);

    const ctx = context ? ` (${context})` : "";
    console.warn(`${PREFIX} DEBUG: ${feature} - ${message}${ctx}`);
  }

  getWarningSummary(): WarningSummary {
    const features: WarningSummary["features"] = [];
    for (const [feature, { message, count }] of this.featureCounts) {
      features.push({ feature, message, count });
    }
    return { totalCount: this.entries.length, features };
  }

  flushWarnings(): WarningSummary {
    const summary = this.getWarningSummary();

    if (this.level !== "off" && summary.features.length > 0) {
      console.warn(`${PREFIX} Summary: ${summary.features.length} unsupported feature(s) detected`);
      for (const { feature, count } of summary.features) {
        console.warn(`  - ${feature}: ${count} occurrence(s)`);
      }
    }

    this.entries = [];
    this.featureCounts.clear();

    return summary;
  }

  getWarningEntries(): readonly WarningEntry[] {
    return this.entries;
  }

  getLogLevel(): LogLevel {
    return this.level;
  }

  private record(feature: string, message: string, context?: string): void {
    this.entries.push({ feature, message, ...(context !== undefined && { context }) });

    const existing = this.featureCounts.get(feature);
    if (existing) {
      existing.count++;
    } else {
      this.featureCounts.set(feature, { message, count: 1 });
    }
  }
}

let activeLogger: WarningLogger = createWarningLogger("off");

export function createWarningLogger(level: LogLevel): WarningLogger {
  return new InMemoryWarningLogger(level);
}

export function initWarningLogger(level: LogLevel): void {
  activeLogger = createWarningLogger(level);
}

export function getActiveWarningLogger(): WarningLogger {
  return activeLogger;
}

export function warn(feature: string, message: string, context?: string): void {
  activeLogger.warn(feature, message, context);
}

export function debug(feature: string, message: string, context?: string): void {
  activeLogger.debug(feature, message, context);
}

/**
 * Return a grouped summary of currently collected warnings.
 *
 * This reads the active warning logger state. The public conversion APIs flush
 * warnings at the end of a conversion so summaries can be printed; after that
 * flush, this returns an empty summary until new warnings are recorded. It is
 * mainly useful for lower-level renderer workflows, tests, and integrations
 * that manage the warning cycle directly.
 */
export function getWarningSummary(): WarningSummary {
  return activeLogger.getWarningSummary();
}

export function flushWarnings(): WarningSummary {
  return activeLogger.flushWarnings();
}

/**
 * Return individual warning entries collected in the current warning cycle.
 *
 * Entries include optional slide or element context when available. The public
 * conversion APIs flush entries at the end of a conversion, so this is mainly
 * useful for lower-level renderer workflows, tests, and integrations that
 * manage the warning cycle directly.
 */
export function getWarningEntries(): readonly WarningEntry[] {
  return activeLogger.getWarningEntries();
}

export function getLogLevel(): LogLevel {
  return activeLogger.getLogLevel();
}
