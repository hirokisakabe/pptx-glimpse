import { posix } from "node:path";

import { fromMarkdown } from "mdast-util-from-markdown";

const API_REFERENCE_PREFIX = "/docs/api/";

interface MarkdownNode {
  readonly type: string;
  readonly url?: string;
  readonly position?: {
    readonly start: { readonly offset?: number };
    readonly end: { readonly offset?: number };
  };
  readonly children?: readonly unknown[];
}

interface Replacement {
  readonly start: number;
  readonly end: number;
  readonly value: string;
}

/**
 * Removes TypeDoc's source extension from generated API-reference links without changing
 * prose, code fences, or links outside the generated content tree.
 */
export function normalizeGeneratedMarkdownLinks(source: string): string {
  const replacements: Replacement[] = [];
  collectLinkReplacements(fromMarkdown(source), source, replacements);

  let result = source;
  for (const replacement of replacements.sort((left, right) => right.start - left.start)) {
    result = result.slice(0, replacement.start) + replacement.value + result.slice(replacement.end);
  }
  return result;
}

export function inlineGeneratedConvertOptions(source: string, properties: string): string {
  const returnsStart = source.indexOf("\n## Returns\n");
  if (returnsStart === -1) {
    throw new Error("Unable to find generated Returns section");
  }
  return `${source.slice(0, returnsStart).trimEnd()}

## ConvertOptions

The options accepted by this function are expanded here from
[\`ConvertOptions\`](/docs/api/node/interfaces/ConvertOptions).

${properties.trim()}
${source.slice(returnsStart)}`;
}

export function formatGeneratedEditorCommand(source: string): string {
  const signatureStart = source.indexOf("\n```ts\n");
  const signatureEnd =
    signatureStart === -1 ? -1 : source.indexOf("\n```\n", signatureStart + "\n```ts\n".length);
  if (signatureStart === -1 || signatureEnd === -1) {
    throw new Error("Unable to find generated EditorCommand union signature");
  }

  let commandCount = 0;
  const withoutSignature =
    source.slice(0, signatureStart) + source.slice(signatureEnd + "\n```\n".length);
  const formatted = withoutSignature
    .replace("## Union Members", "## Commands")
    .replace(/### Type Literal\n\n```ts\n([\s\S]*?)```/g, (_match, declaration: string) => {
      const kindMatch = declaration.match(/^  kind: "([^"]+)";$/m);
      if (kindMatch === null) {
        throw new Error("Unable to find a top-level kind in an EditorCommand union member");
      }
      const [kindLine, kind] = kindMatch;
      const declarationWithoutKind = declaration.replace(`${kindLine}\n`, "");
      const orderedDeclaration = declarationWithoutKind.replace("{\n", `{\n${kindLine}\n`);
      commandCount += 1;
      return `### \`${kind}\`\n\n\`\`\`ts\n${orderedDeclaration}\`\`\``;
    });

  if (commandCount === 0) {
    throw new Error("No EditorCommand union members were formatted");
  }
  return formatted;
}

function collectLinkReplacements(
  node: MarkdownNode,
  source: string,
  replacements: Replacement[],
): void {
  if (node.type === "link" && node.url !== undefined && node.position !== undefined) {
    const start = node.position.start.offset;
    const end = node.position.end.offset;
    if (start !== undefined && end !== undefined) {
      const destinationRange = findInlineDestinationRange(node, source, start, end);
      if (destinationRange !== undefined) {
        const rawDestination = source.slice(destinationRange.start, destinationRange.end);
        const normalized = normalizeRawGeneratedDestination(rawDestination, node.url);
        if (normalized !== undefined) {
          replacements.push({
            start: destinationRange.start,
            end: destinationRange.end,
            value: normalized,
          });
        }
      }
    }
  }

  for (const child of node.children ?? []) {
    if (isMarkdownNode(child)) {
      collectLinkReplacements(child, source, replacements);
    }
  }
}

function findInlineDestinationRange(
  node: MarkdownNode,
  source: string,
  start: number,
  end: number,
): { readonly start: number; readonly end: number } | undefined {
  let labelContentEnd = start + 1;
  for (const child of node.children ?? []) {
    if (isMarkdownNode(child)) {
      labelContentEnd = Math.max(labelContentEnd, child.position?.end.offset ?? labelContentEnd);
    }
  }

  const closingBracket = source.indexOf("]", labelContentEnd);
  if (closingBracket === -1 || closingBracket >= end) {
    return undefined;
  }

  let cursor = closingBracket + 1;
  while (cursor < end && /\s/.test(source[cursor])) cursor += 1;
  if (source[cursor] !== "(") {
    return undefined;
  }
  cursor += 1;
  while (cursor < end && /\s/.test(source[cursor])) cursor += 1;

  if (source[cursor] === "<") {
    const destinationStart = cursor + 1;
    cursor = destinationStart;
    while (cursor < end) {
      if (source[cursor] === "\\" && cursor + 1 < end) {
        cursor += 2;
      } else if (source[cursor] === ">") {
        return { start: destinationStart, end: cursor };
      } else {
        cursor += 1;
      }
    }
    return undefined;
  }

  const destinationStart = cursor;
  let parenthesisDepth = 0;
  while (cursor < end) {
    const character = source[cursor];
    if (character === "\\" && cursor + 1 < end) {
      cursor += 2;
    } else if (character === "(") {
      parenthesisDepth += 1;
      cursor += 1;
    } else if (character === ")") {
      if (parenthesisDepth === 0) {
        return { start: destinationStart, end: cursor };
      }
      parenthesisDepth -= 1;
      cursor += 1;
    } else if (/\s/.test(character) && parenthesisDepth === 0) {
      return { start: destinationStart, end: cursor };
    } else {
      cursor += 1;
    }
  }
  return undefined;
}

function isMarkdownNode(value: unknown): value is MarkdownNode {
  return (
    typeof value === "object" && value !== null && "type" in value && typeof value.type === "string"
  );
}

function normalizeRawGeneratedDestination(
  rawDestination: string,
  decodedDestination: string,
): string | undefined {
  const isGeneratedLink =
    decodedDestination.startsWith(API_REFERENCE_PREFIX) ||
    decodedDestination.startsWith("./") ||
    decodedDestination.startsWith("../");
  if (!isGeneratedLink) {
    return undefined;
  }

  const decodedHashIndex = decodedDestination.indexOf("#");
  const decodedPath =
    decodedHashIndex === -1 ? decodedDestination : decodedDestination.slice(0, decodedHashIndex);
  if (!decodedPath.endsWith(".mdx")) {
    return undefined;
  }

  const rawHashIndex = rawDestination.indexOf("#");
  const rawPath = rawHashIndex === -1 ? rawDestination : rawDestination.slice(0, rawHashIndex);
  const rawHash = rawHashIndex === -1 ? "" : rawDestination.slice(rawHashIndex);
  if (rawPath.endsWith("/index.mdx")) {
    return `${rawPath.slice(0, -"/index.mdx".length)}${rawHash}`;
  }
  if (rawPath.endsWith(".mdx")) {
    return `${rawPath.slice(0, -".mdx".length)}${rawHash}`;
  }
  return undefined;
}

export function formatNavigationTitle(title: string, path?: string): string {
  if (path === "node/index.mdx") {
    return "Node.js entry point";
  }
  if (path === "browser/index.mdx") {
    return "Browser entry point";
  }
  return title;
}

/** Validates a TypeDoc navigation path before it is used for filesystem output. */
export function isValidNavigationPath(path: string): boolean {
  if (
    path === "" ||
    path.includes("\\") ||
    posix.isAbsolute(path) ||
    posix.extname(path) !== ".mdx"
  ) {
    return false;
  }
  const segments = path.split("/");
  return (
    segments.every((segment) => segment !== "" && segment !== "." && segment !== "..") &&
    posix.normalize(path) === path
  );
}

export function navigationPathParts(path: string): {
  readonly directory: string;
  readonly name: string;
} {
  if (!isValidNavigationPath(path)) {
    throw new Error(`Invalid TypeDoc navigation path: ${path}`);
  }
  const withoutExtension = path.slice(0, -posix.extname(path).length);
  return {
    directory: posix.dirname(withoutExtension),
    name: posix.basename(withoutExtension),
  };
}
