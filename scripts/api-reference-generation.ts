import { posix } from "node:path";

import { fromMarkdown } from "mdast-util-from-markdown";

const API_REFERENCE_PREFIX = "/docs/api-reference/";

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

function collectLinkReplacements(
  node: MarkdownNode,
  source: string,
  replacements: Replacement[],
): void {
  if (node.type === "link" && node.url !== undefined && node.position !== undefined) {
    const start = node.position.start.offset;
    const end = node.position.end.offset;
    if (start !== undefined && end !== undefined) {
      const rawLink = source.slice(start, end);
      const destinationMatch = /\]\(\s*(?:<([^>\n]+)>|([^\s)]+))/.exec(rawLink);
      const rawDestination = destinationMatch?.[1] ?? destinationMatch?.[2];
      if (destinationMatch !== null && rawDestination !== undefined) {
        const normalized = normalizeGeneratedDestination(node.url);
        if (normalized !== undefined) {
          const destinationOffset = destinationMatch[0].lastIndexOf(rawDestination);
          const destinationStart = start + destinationMatch.index + destinationOffset;
          replacements.push({
            start: destinationStart,
            end: destinationStart + rawDestination.length,
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

function isMarkdownNode(value: unknown): value is MarkdownNode {
  return (
    typeof value === "object" && value !== null && "type" in value && typeof value.type === "string"
  );
}

function normalizeGeneratedDestination(destination: string): string | undefined {
  const isGeneratedLink =
    destination.startsWith(API_REFERENCE_PREFIX) ||
    destination.startsWith("./") ||
    destination.startsWith("../");
  if (!isGeneratedLink) {
    return undefined;
  }

  const hashIndex = destination.indexOf("#");
  const path = hashIndex === -1 ? destination : destination.slice(0, hashIndex);
  const hash = hashIndex === -1 ? "" : destination.slice(hashIndex);
  if (!path.endsWith(".mdx")) {
    return undefined;
  }

  const withoutExtension = path.endsWith("/index.mdx")
    ? path.slice(0, -"/index.mdx".length)
    : path.slice(0, -".mdx".length);
  return `${withoutExtension}${hash}`;
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
