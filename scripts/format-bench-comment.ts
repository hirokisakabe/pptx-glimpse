/**
 * vitest bench --outputJson の結果をPRコメント用マークダウンに整形する。
 *
 * Usage:
 *   npx tsx scripts/format-bench-comment.ts <current.json> [baseline.json]
 *
 * - baseline.json が指定された場合: 比較テーブルを生成
 * - baseline.json がない場合: 現在の結果のみ表示
 */
import { readFileSync } from "node:fs";

interface VitestBenchmark {
  name: string;
  hz: number;
  rme: number;
  sampleCount: number;
}

interface VitestGroup {
  fullName: string;
  benchmarks: VitestBenchmark[];
}

interface VitestFile {
  filepath: string;
  groups: VitestGroup[];
}

interface VitestBenchOutput {
  files: VitestFile[];
}

interface FlatBenchmark {
  suite: string;
  name: string;
  hz: number;
  rme: number;
  sampleCount: number;
}

function extractSuiteName(fullName: string): string {
  const parts = fullName.split(" > ");
  return parts.length > 1 ? parts.slice(1).join(" > ") : fullName;
}

function parseBenchOutput(path: string): FlatBenchmark[] {
  const raw = JSON.parse(readFileSync(path, "utf-8")) as VitestBenchOutput;
  const results: FlatBenchmark[] = [];
  for (const file of raw.files) {
    for (const group of file.groups) {
      const suite = extractSuiteName(group.fullName);
      for (const bench of group.benchmarks) {
        results.push({
          suite,
          name: bench.name,
          hz: bench.hz,
          rme: bench.rme,
          sampleCount: bench.sampleCount,
        });
      }
    }
  }
  return results;
}

function formatHz(hz: number): string {
  return hz.toLocaleString("en-US", { maximumFractionDigits: 2 });
}

function changeStatus(pct: number): string {
  if (pct <= -20) return ":rotating_light:";
  if (pct <= -10) return ":warning:";
  if (pct >= 10) return ":rocket:";
  return "";
}

function generateComparisonTable(current: FlatBenchmark[], baseline: FlatBenchmark[]): string {
  const baselineMap = new Map<string, FlatBenchmark>();
  for (const b of baseline) {
    baselineMap.set(`${b.suite} > ${b.name}`, b);
  }

  const lines: string[] = [
    "## Benchmark Results",
    "",
    "| Suite | Benchmark | ops/sec (current) | ops/sec (baseline) | Change | |",
    "|---|---|---:|---:|---:|---|",
  ];

  for (const c of current) {
    const key = `${c.suite} > ${c.name}`;
    const b = baselineMap.get(key);
    if (b) {
      const pct = ((c.hz - b.hz) / b.hz) * 100;
      const sign = pct >= 0 ? "+" : "";
      const status = changeStatus(pct);
      lines.push(
        `| ${c.suite} | ${c.name} | ${formatHz(c.hz)} | ${formatHz(b.hz)} | ${sign}${pct.toFixed(1)}% | ${status} |`,
      );
    } else {
      lines.push(`| ${c.suite} | ${c.name} | ${formatHz(c.hz)} | - | new | |`);
    }
  }

  lines.push(
    "",
    "> **Note**: Benchmark results can vary between CI runs due to shared infrastructure. Changes within ±10% are generally noise.",
  );

  return lines.join("\n");
}

function generateCurrentOnlyTable(current: FlatBenchmark[]): string {
  const lines: string[] = [
    "## Benchmark Results",
    "",
    "> :information_source: No baseline data available. Comparison will be available after the first merge to main.",
    "",
    "| Suite | Benchmark | ops/sec | ±RME | Samples |",
    "|---|---|---:|---|---:|",
  ];

  for (const c of current) {
    lines.push(
      `| ${c.suite} | ${c.name} | ${formatHz(c.hz)} | ±${c.rme.toFixed(2)}% | ${c.sampleCount} |`,
    );
  }

  return lines.join("\n");
}

function main() {
  const currentPath = process.argv[2];
  const baselinePath = process.argv[3];

  if (!currentPath) {
    console.error("Usage: npx tsx scripts/format-bench-comment.ts <current.json> [baseline.json]");
    process.exit(1);
  }

  const current = parseBenchOutput(currentPath);

  if (baselinePath) {
    const baseline = parseBenchOutput(baselinePath);
    console.log(generateComparisonTable(current, baseline));
  } else {
    console.log(generateCurrentOnlyTable(current));
  }
}

main();
