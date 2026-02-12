export type LogLevel = "off" | "warn" | "debug";

export interface WarningEntry {
  /** The feature key, e.g. "sp.style", "bodyPr@vert" */
  feature: string;
  /** Human-readable description */
  message: string;
  /** Location context, e.g. "Slide 1" */
  context?: string;
}

export interface WarningSummary {
  totalCount: number;
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

export function getWarningEntries(): readonly WarningEntry[] {
  return entries;
}

export function getLogLevel(): LogLevel {
  return currentLevel;
}
