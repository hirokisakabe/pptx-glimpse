/**
 * vitest bench --outputJson の出力を
 * benchmark-action/github-action-benchmark の customBiggerIsBetter 形式に変換する。
 *
 * Usage: npx tsx scripts/transform-bench-json.ts <input.json> <output.json>
 */
import { readFileSync, writeFileSync } from "node:fs";

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

interface BenchmarkEntry {
  name: string;
  unit: string;
  value: number;
  range: string;
  extra: string;
}

function extractSuiteName(fullName: string): string {
  // "bench/conversion.bench.ts > E2E conversion" → "E2E conversion"
  const parts = fullName.split(" > ");
  return parts.length > 1 ? parts.slice(1).join(" > ") : fullName;
}

function main() {
  const inputPath = process.argv[2];
  const outputPath = process.argv[3];

  if (!inputPath || !outputPath) {
    console.error("Usage: npx tsx scripts/transform-bench-json.ts <input.json> <output.json>");
    process.exit(1);
  }

  const raw = JSON.parse(readFileSync(inputPath, "utf-8")) as VitestBenchOutput;
  const results: BenchmarkEntry[] = [];

  for (const file of raw.files) {
    for (const group of file.groups) {
      const suiteName = extractSuiteName(group.fullName);
      for (const bench of group.benchmarks) {
        results.push({
          name: `${suiteName} > ${bench.name}`,
          unit: "ops/sec",
          value: bench.hz,
          range: `\u00b1${bench.rme.toFixed(2)}%`,
          extra: `${bench.sampleCount} samples`,
        });
      }
    }
  }

  writeFileSync(outputPath, JSON.stringify(results, null, 2) + "\n");
  console.log(`Transformed ${results.length} benchmarks → ${outputPath}`);
}

main();
