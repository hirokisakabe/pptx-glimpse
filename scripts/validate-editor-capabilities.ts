import { readdirSync, readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import {
  type CapabilityStatus,
  CORE_SESSION_NON_CAPABILITY_MEMBERS,
  DOCUMENT_NON_CAPABILITY_EXPORTS,
  EDITOR_CAPABILITIES,
  type EditorCapability,
} from "./editor-capability-manifest.js";

const REPOSITORY_ROOT = fileURLToPath(new URL("../", import.meta.url));

export interface CapabilitySources {
  readonly documentRoot: string;
  readonly editorRoot: string;
  readonly coreSession: string;
  readonly reactSources: string;
  readonly matrixDocument: string;
}

export function validateEditorCapabilities(sources: CapabilitySources): readonly string[] {
  const errors: string[] = [];
  errors.push(...validateEditorCapabilityManifest(EDITOR_CAPABILITIES));
  const documentApis = EDITOR_CAPABILITIES.flatMap((entry) => entry.documentApis);
  const editorCommands = EDITOR_CAPABILITIES.flatMap((entry) => entry.editorCommands);
  const coreMembers = EDITOR_CAPABILITIES.flatMap((entry) => entry.coreMembers);
  const uiCommandKinds = EDITOR_CAPABILITIES.flatMap((entry) => entry.uiCommandKinds);
  const uiCoreMembers = EDITOR_CAPABILITIES.flatMap((entry) => entry.uiCoreMembers);
  const publicDocumentExports = extractNamedValueExports(sources.documentRoot);
  compareSets(
    "document public value exports",
    publicDocumentExports,
    [...documentApis, ...DOCUMENT_NON_CAPABILITY_EXPORTS],
    errors,
  );

  const publicEditorCommands = extractEditorCommandKinds(sources.editorRoot);
  compareSets("EditorCommand kinds", publicEditorCommands, editorCommands, errors);

  const publicCoreMembers = extractCoreSessionMembers(sources.coreSession);
  compareSets(
    "PptxEditorSession public members",
    publicCoreMembers,
    [...coreMembers, ...CORE_SESSION_NON_CAPABILITY_MEMBERS],
    errors,
  );

  const publicUiCommandKinds = extractStringKinds(sources.reactSources).filter((kind) =>
    publicEditorCommands.includes(kind),
  );
  compareSets("editor-react command kinds", publicUiCommandKinds, uiCommandKinds, errors);

  const uiCoreCalls = extractMemberCalls(sources.reactSources).filter((member) =>
    coreMembers.includes(member),
  );
  compareSets("editor-react direct core capability calls", uiCoreCalls, uiCoreMembers, errors);

  const generatedMatrix = renderEditorCapabilityMatrix();
  if (sources.matrixDocument !== generatedMatrix) {
    errors.push(
      "docs/development/editor-capability-matrix.md is stale; regenerate it from EDITOR_CAPABILITIES.",
    );
  }
  return errors;
}

export function validateEditorCapabilityManifest(
  capabilities: readonly EditorCapability[],
): readonly string[] {
  const errors: string[] = [];
  reportDuplicates(
    "capability names",
    capabilities.map((entry) => entry.capability),
    errors,
  );
  for (const field of [
    "documentApis",
    "editorCommands",
    "coreMembers",
    "uiCommandKinds",
    "uiCoreMembers",
  ] as const) {
    reportDuplicates(
      `${field} ownership`,
      capabilities.flatMap((entry) => entry[field]),
      errors,
    );
  }

  for (const entry of capabilities) {
    validateStatus(entry, "document", entry.documentApis.length > 0, errors);
    validateStatus(entry, "editor", entry.editorCommands.length > 0, errors);
    validateStatus(
      entry,
      "core",
      entry.editorCommands.length > 0 || entry.coreMembers.length > 0,
      errors,
    );
    validateStatus(
      entry,
      "ui",
      entry.uiCommandKinds.length > 0 || entry.uiCoreMembers.length > 0,
      errors,
    );
    for (const kind of entry.uiCommandKinds) {
      if (!entry.editorCommands.includes(kind)) {
        errors.push(
          `${entry.capability}: uiCommandKinds member ${kind} is not owned by editorCommands.`,
        );
      }
    }
    for (const member of entry.uiCoreMembers) {
      if (!entry.coreMembers.includes(member)) {
        errors.push(
          `${entry.capability}: uiCoreMembers member ${member} is not owned by coreMembers.`,
        );
      }
    }
  }
  return errors;
}

function validateStatus(
  entry: EditorCapability,
  layer: "document" | "editor" | "core" | "ui",
  hasCapability: boolean,
  errors: string[],
): void {
  const supported = entry[layer].kind === "supported";
  if (supported !== hasCapability) {
    errors.push(
      `${entry.capability}: ${layer} status is ${entry[layer].kind} but row metadata is ${hasCapability ? "present" : "empty"}.`,
    );
  }
}

function reportDuplicates(label: string, values: readonly string[], errors: string[]): void {
  const seen = new Set<string>();
  const duplicates = new Set<string>();
  for (const value of values) {
    if (seen.has(value)) duplicates.add(value);
    seen.add(value);
  }
  if (duplicates.size > 0) {
    errors.push(`${label}: duplicate: ${[...duplicates].sort().join(", ")}`);
  }
}

function renderEditorCapabilityMatrix(): string {
  const rows = EDITOR_CAPABILITIES.map(
    (entry) =>
      `| ${entry.capability} | ${entry.documentApis.map(code).join("<br>")} | ${formatStatus(entry.document)} | ${formatStatus(entry.editor)} | ${formatStatus(entry.core)} | ${formatStatus(entry.ui)} |`,
  ).join("\n");
  return (
    `# Editor Capability Matrix

This matrix tracks how every public \`@pptx-glimpse/document\` root editing/authoring API
reaches the headless \`EditorCommand\`, public \`PptxEditorSession\`, and
\`@pptx-glimpse/editor-react\` layers. It records reachability, not rendering fidelity; see
[document feature support](../../packages/document/docs/feature-support.md) for reader, computed,
writer, edit, and preservation coverage.

## Status legend

- **S — supported**: the layer exposes the capability; the manifest records source evidence.
- **I — intentional boundary**: the capability deliberately stops before this layer for the stated
  package-responsibility or product-scope reason.
- **T — tracked**: the capability is not present at this layer and an existing issue owns the work.

The machine-readable source of truth is
[\`scripts/editor-capability-manifest.ts\`](../../scripts/editor-capability-manifest.ts). The
validation test compares that manifest with document root value exports, every \`EditorCommand\`
kind, public \`PptxEditorSession\` members, React command usage, and this rendered table. Therefore
adding or changing a mutation/capability requires updating the manifest and this document in the
same change.

## Matrix

| Capability | Document root API | Document | EditorCommand | Core session | React UI |
| --- | --- | :---: | :---: | :---: | :---: |
${rows}

## Boundary notes

- ` +
    "`@pptx-glimpse/editor-react` continues to import production APIs only from the public " +
    "`pptx-glimpse` root. This matrix does not authorize lower-layer imports.\n" +
    `- A tracked status names an existing issue; closed implementation issues are evidence for an
  already-supported layer, not a substitute for current source evidence.
- New document authoring/editing exports, editor commands, core session capabilities, or React UI
  operations must be classified as supported, intentional, or tracked. Unclassified additions fail
  \`scripts/validate-editor-capabilities.test.ts\`.
`
  );
}

function formatStatus(status: CapabilityStatus): string {
  if (status.kind === "supported") return `S (${status.evidence.map(code).join(", ")})`;
  if (status.kind === "tracked") {
    return `T ([#${status.issue}](https://github.com/hirokisakabe/pptx-glimpse/issues/${status.issue}): ${status.reason})`;
  }
  return `I (${status.reason})`;
}

function code(value: string): string {
  return `\`${value}\``;
}

export function extractNamedValueExports(source: string): readonly string[] {
  const names: string[] = [];
  for (const match of source.matchAll(/export\s*\{([\s\S]*?)\}\s*from\s*"[^"]+";/gu)) {
    const body = match[1];
    if (body === undefined) continue;
    for (const item of body.split(",")) {
      const normalized = item.trim();
      if (normalized === "" || normalized.startsWith("type ")) continue;
      const name = normalized.split(/\s+as\s+/u).at(-1);
      if (name !== undefined) names.push(name.trim());
    }
  }
  return uniqueSorted(names);
}

export function extractEditorCommandKinds(source: string): readonly string[] {
  const union = source.match(/export type EditorCommand\s*=([\s\S]*?);/u)?.[1];
  if (union === undefined) return [];
  const aliases = new Map<string, string>();
  for (const match of source.matchAll(
    /(?:export\s+)?type\s+([A-Z][A-Za-z0-9]*)\s*=([\s\S]*?);/gu,
  )) {
    const [, name, expression] = match;
    if (name !== undefined && expression !== undefined) aliases.set(name, expression);
  }
  const kinds = new Map<string, string>();
  for (const match of source.matchAll(
    /export interface\s+([A-Z][A-Za-z0-9]*Command)(?:\s+extends[^\{]+)?\s*\{[\s\S]*?readonly kind:\s*"([^"]+)";/gu,
  )) {
    const [, name, kind] = match;
    if (name !== undefined && kind !== undefined) kinds.set(name, kind);
  }

  const resolvedKinds = new Set<string>();
  const visited = new Set<string>();
  const visitExpression = (expression: string): void => {
    for (const match of expression.matchAll(/\b([A-Z][A-Za-z0-9]*Command)\b/gu)) {
      const name = match[1];
      if (name === undefined || visited.has(name)) continue;
      visited.add(name);
      const kind = kinds.get(name);
      if (kind !== undefined) resolvedKinds.add(kind);
      const alias = aliases.get(name);
      if (alias !== undefined) visitExpression(alias);
    }
  };
  visitExpression(union);
  return uniqueSorted([...resolvedKinds]);
}

export function extractCoreSessionMembers(source: string): readonly string[] {
  const classStart = source.indexOf("export class PptxEditorSession");
  const classEnd = source.indexOf("\nfunction buildLayoutCatalog", classStart);
  const classSource = source.slice(classStart, classEnd === -1 ? undefined : classEnd);
  const members = [
    ...classSource.matchAll(/^  (?:(?:public|static|async|get)\s+)*([A-Za-z][A-Za-z0-9]*)\s*\(/gmu),
  ]
    .flatMap((match) => (match[1] === undefined ? [] : [match[1]]))
    .filter((name) => name !== "constructor");
  return uniqueSorted(members);
}

function extractStringKinds(source: string): readonly string[] {
  return uniqueSorted(
    [...source.matchAll(/\bkind:\s*"([^"]+)"/gu)].flatMap((match) =>
      match[1] === undefined ? [] : [match[1]],
    ),
  );
}

function extractMemberCalls(source: string): readonly string[] {
  return uniqueSorted(
    [...source.matchAll(/\b[A-Za-z][A-Za-z0-9]*\.([A-Za-z][A-Za-z0-9]*)\s*\(/gu)].flatMap(
      (match) => (match[1] === undefined ? [] : [match[1]]),
    ),
  );
}

function compareSets(
  label: string,
  actual: readonly string[],
  expected: readonly string[],
  errors: string[],
): void {
  const actualSet = new Set(actual);
  const expectedSet = new Set(expected);
  const unclassified = [...actualSet].filter((value) => !expectedSet.has(value)).sort();
  const stale = [...expectedSet].filter((value) => !actualSet.has(value)).sort();
  if (unclassified.length > 0) errors.push(`${label}: unclassified: ${unclassified.join(", ")}`);
  if (stale.length > 0) errors.push(`${label}: manifest-only: ${stale.join(", ")}`);
}

function uniqueSorted(values: readonly string[]): readonly string[] {
  return [...new Set(values)].sort();
}

export function readRepositoryCapabilitySources(): CapabilitySources {
  const read = (relativePath: string) => readFileSync(`${REPOSITORY_ROOT}${relativePath}`, "utf8");
  const readProductionSources = (relativeDirectory: string): string =>
    readdirSync(`${REPOSITORY_ROOT}${relativeDirectory}`, { withFileTypes: true })
      .filter(
        (entry) =>
          entry.isFile() &&
          /\.tsx?$/u.test(entry.name) &&
          !/\.(?:test|spec)\.tsx?$/u.test(entry.name),
      )
      .map((entry) => read(`${relativeDirectory}/${entry.name}`))
      .join("\n");
  return {
    documentRoot: read("packages/document/src/index.ts"),
    editorRoot: [
      read("packages/editor/src/index.ts"),
      readProductionSources("packages/editor/src/commands"),
    ].join("\n"),
    coreSession: read("packages/core/src/pptx-editor-session.ts"),
    reactSources: readProductionSources("packages/editor-react/src"),
    matrixDocument: read("docs/development/editor-capability-matrix.md"),
  };
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const errors = validateEditorCapabilities(readRepositoryCapabilitySources());
  if (errors.length > 0) {
    for (const error of errors) console.error(error);
    process.exitCode = 1;
  } else {
    console.log("Editor capability manifest is current.");
  }
}
