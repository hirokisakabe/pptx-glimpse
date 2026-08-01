import { dirname } from "node:path";
import { fileURLToPath } from "node:url";

import { ESLint, type Linter } from "eslint";

const repoRoot = dirname(fileURLToPath(new URL("../package.json", import.meta.url)));
const eslint = new ESLint({ cwd: repoRoot });

const representativeFiles = [
  "packages/core/src/index.ts",
  "scripts/test-render.ts",
  "vrt/snapshot/vrt-cases.ts",
  "bench/conversion.bench.ts",
  "e2e/smoke.test.ts",
];

const commonRules = [
  "simple-import-sort/imports",
  "simple-import-sort/exports",
  "import-x/no-restricted-paths",
  "@typescript-eslint/no-unnecessary-type-assertion",
  "@typescript-eslint/no-unsafe-type-assertion",
  "@typescript-eslint/consistent-type-assertions",
];

function ruleSeverity(rule: Linter.RuleEntry | undefined): Linter.RuleSeverity | undefined {
  return Array.isArray(rule) ? rule[0] : rule;
}

function assertErrorRules(
  file: string,
  rules: Linter.RulesRecord | undefined,
  expectedRules: readonly string[],
): void {
  if (rules === undefined) {
    throw new Error(`ESLint did not resolve rules for ${file}`);
  }

  for (const ruleName of expectedRules) {
    if (ruleSeverity(rules[ruleName]) !== 2) {
      throw new Error(`${ruleName} is not enabled as an error for ${file}`);
    }
  }
}

for (const file of representativeFiles) {
  const config = await eslint.calculateConfigForFile(file);
  assertErrorRules(file, config?.rules, commonRules);
}

const packageConfig = await eslint.calculateConfigForFile(representativeFiles[0]);
assertErrorRules(representativeFiles[0], packageConfig?.rules, [
  "import-x/no-relative-packages",
  "import-x/no-extraneous-dependencies",
  "no-restricted-imports",
]);

console.log(
  `Verified the root ESLint config and policy rules for ${representativeFiles.length} lint targets.`,
);
