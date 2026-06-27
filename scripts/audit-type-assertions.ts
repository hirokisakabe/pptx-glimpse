import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";

import ts from "typescript";

const CATEGORIES = [
  "xmlOrExternalInput",
  "testFixture",
  "asConst",
  "brandedConstructor",
  "doubleUnknown",
  "objectLiteral",
  "arrayLiteral",
  "asAny",
  "externalInterop",
  "other",
] as const;

type Category = (typeof CATEGORIES)[number];

interface CategoryBucket {
  readonly description: string;
  count: number;
  readonly examples: string[];
}

const buckets: Record<Category, CategoryBucket> = {
  xmlOrExternalInput: {
    description: "XML / external input boundary",
    count: 0,
    examples: [],
  },
  testFixture: {
    description: "Test fixture / mock construction",
    count: 0,
    examples: [],
  },
  asConst: {
    description: "`as const` literal preservation",
    count: 0,
    examples: [],
  },
  brandedConstructor: {
    description: "Branded unit / handle constructors",
    count: 0,
    examples: [],
  },
  doubleUnknown: {
    description: "Double assertion (`as unknown as X`)",
    count: 0,
    examples: [],
  },
  objectLiteral: {
    description: "Object literal assertion",
    count: 0,
    examples: [],
  },
  arrayLiteral: {
    description: "Array literal assertion",
    count: 0,
    examples: [],
  },
  asAny: {
    description: "`as any`",
    count: 0,
    examples: [],
  },
  externalInterop: {
    description: "External library / platform interop",
    count: 0,
    examples: [],
  },
  other: {
    description: "Other narrow assertions",
    count: 0,
    examples: [],
  },
};

const lintedFileGlobs = [
  "packages/*/src/**/*.ts",
  "packages/*/src/**/*.tsx",
  "vrt/**/*.ts",
  "scripts/**/*.ts",
  "bench/**/*.ts",
  "e2e/**/*.ts",
];

function listSourceFiles(): string[] {
  const globArgs = lintedFileGlobs.flatMap((glob) => ["-g", glob]);
  return execFileSync("rg", ["--files", ...globArgs, "."], { encoding: "utf8" })
    .split("\n")
    .filter((file) => file.length > 0);
}

function classify(file: string, sourceFile: ts.SourceFile, node: AssertionNode): Category {
  const typeText = node.type.getText(sourceFile);
  const expressionText = node.expression.getText(sourceFile);
  const assertionText = node.getText(sourceFile);

  if (typeText === "const") return "asConst";
  if (typeText === "any") return "asAny";
  if (assertionText.includes("as unknown as")) return "doubleUnknown";
  if (ts.isObjectLiteralExpression(node.expression)) return "objectLiteral";
  if (ts.isArrayLiteralExpression(node.expression)) return "arrayLiteral";
  if (isXmlOrExternalInput(typeText)) return "xmlOrExternalInput";
  if (isBrandedConstructor(typeText)) return "brandedConstructor";
  if (isTestOrVrtFile(file)) return "testFixture";
  if (isExternalInterop(typeText, expressionText)) return "externalInterop";
  return "other";
}

function isXmlOrExternalInput(typeText: string): boolean {
  return /\b(XmlNode|XmlOrderedNode|Record<string, unknown>|PackageJson)\b/.test(typeText);
}

function isBrandedConstructor(typeText: string): boolean {
  return /\b(Emu|Pt|HundredthPt|OoxmlPercent|OoxmlAngle|PartPath|RelationshipId|SourceNodeId|RawSidecarId)\b/.test(
    typeText,
  );
}

function isTestOrVrtFile(file: string): boolean {
  return /\.(test|e2e\.test)\.ts$/.test(file) || file.startsWith("vrt/");
}

function isExternalInterop(typeText: string, expressionText: string): boolean {
  return (
    /\b(ArrayBuffer|OpentypeFullFont|SubsettableFont|ReturnType<typeof crypto\.randomUUID>|Error|VitestBenchOutput|SlideSvg)\b/.test(
      typeText,
    ) || expressionText.startsWith("JSON.parse(")
  );
}

function record(
  category: Category,
  file: string,
  sourceFile: ts.SourceFile,
  node: AssertionNode,
): void {
  const bucket = buckets[category];
  const line = sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile)).line + 1;
  bucket.count += 1;
  if (bucket.examples.length < 3) {
    bucket.examples.push(`${file}:${line}: ${node.getText(sourceFile).replace(/\s+/g, " ")}`);
  }
}

type AssertionNode = ts.AsExpression | ts.TypeAssertion;

function isAssertionNode(node: ts.Node): node is AssertionNode {
  return ts.isAsExpression(node) || ts.isTypeAssertionExpression(node);
}

function visit(file: string, sourceFile: ts.SourceFile, node: ts.Node): void {
  if (isAssertionNode(node)) {
    record(classify(file, sourceFile, node), file, sourceFile, node);
  }
  ts.forEachChild(node, (child) => visit(file, sourceFile, child));
}

for (const file of listSourceFiles()) {
  const source = readFileSync(file, "utf8");
  const sourceFile = ts.createSourceFile(file, source, ts.ScriptTarget.Latest, true);
  visit(file, sourceFile, sourceFile);
}

const total = CATEGORIES.reduce((sum, category) => sum + buckets[category].count, 0);

console.log(`# Type assertion audit\n`);
console.log(`Total assertions: ${total}\n`);
console.log("| Category | Count |");
console.log("| --- | ---: |");
for (const category of CATEGORIES) {
  console.log(`| ${buckets[category].description} | ${buckets[category].count} |`);
}
console.log("\n## Examples\n");
for (const category of CATEGORIES) {
  const bucket = buckets[category];
  console.log(`### ${bucket.description}\n`);
  if (bucket.examples.length === 0) {
    console.log("- none\n");
    continue;
  }
  for (const example of bucket.examples) {
    console.log(`- ${example}`);
  }
  console.log("");
}
