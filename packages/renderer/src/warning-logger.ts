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

const PREFIX = "[pptx-glimpse]";

let currentLevel: LogLevel = "off";
let entries: WarningEntry[] = [];
const featureCounts = new Map<string, { message: string; count: number }>();

export function initWarningLogger(level: LogLevel): void {
  currentLevel = level;
  entries = [];
  featureCounts.clear();
}

export function warn(feature: string, message: string, context?: string): void {
  if (currentLevel === "off") return;

  entries.push({ feature, message, ...(context !== undefined && { context }) });

  const existing = featureCounts.get(feature);
  if (existing) {
    existing.count++;
  } else {
    featureCounts.set(feature, { message, count: 1 });
  }

  if (currentLevel === "debug") {
    const ctx = context ? ` (${context})` : "";
    console.warn(`${PREFIX} SKIP: ${feature} - ${message}${ctx}`);
  }
}

export function debug(feature: string, message: string, context?: string): void {
  if (currentLevel !== "debug") return;

  entries.push({ feature, message, ...(context !== undefined && { context }) });

  const existing = featureCounts.get(feature);
  if (existing) {
    existing.count++;
  } else {
    featureCounts.set(feature, { message, count: 1 });
  }

  const ctx = context ? ` (${context})` : "";
  console.warn(`${PREFIX} DEBUG: ${feature} - ${message}${ctx}`);
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
  const features: WarningSummary["features"] = [];
  for (const [feature, { message, count }] of featureCounts) {
    features.push({ feature, message, count });
  }
  return { totalCount: entries.length, features };
}

export function flushWarnings(): WarningSummary {
  const summary = getWarningSummary();

  if (currentLevel !== "off" && summary.features.length > 0) {
    console.warn(`${PREFIX} Summary: ${summary.features.length} unsupported feature(s) detected`);
    for (const { feature, count } of summary.features) {
      console.warn(`  - ${feature}: ${count} occurrence(s)`);
    }
  }

  entries = [];
  featureCounts.clear();

  return summary;
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
  return entries;
}

export function getLogLevel(): LogLevel {
  return currentLevel;
}
