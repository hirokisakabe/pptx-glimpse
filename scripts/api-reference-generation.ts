import { posix } from "node:path";

const API_REFERENCE_PREFIX = "/docs/api-reference/";

/**
 * Removes TypeDoc's source extension from generated API-reference links without changing
 * prose, code fences, or links outside the generated content tree.
 */
export function normalizeGeneratedMarkdownLinks(source: string): string {
  let fence: { marker: "`" | "~"; length: number } | undefined;
  let inlineCodeLength: number | undefined;

  return source
    .split(/(?<=\n)/)
    .map((line) => {
      const containerContent = stripBlockquoteMarkers(line);
      if (fence !== undefined) {
        const closingFence = /^ {0,3}(`{3,}|~{3,})[ \t]*(?:\r?\n)?$/.exec(containerContent);
        if (
          closingFence !== null &&
          closingFence[1].startsWith(fence.marker) &&
          closingFence[1].length >= fence.length
        ) {
          fence = undefined;
        }
        return line;
      }

      if (inlineCodeLength === undefined) {
        const openingFence = /^ {0,3}(`{3,}|~{3,})([^\r\n]*)/.exec(containerContent);
        if (
          openingFence !== null &&
          (openingFence[1].startsWith("~") || !openingFence[2].includes("`"))
        ) {
          fence = {
            marker: openingFence[1].startsWith("`") ? "`" : "~",
            length: openingFence[1].length,
          };
          return line;
        }
      }

      let result = "";
      let cursor = 0;
      for (const codeRun of line.matchAll(/`+/g)) {
        const index = codeRun.index;
        const segment = line.slice(cursor, index);
        result +=
          inlineCodeLength === undefined ? normalizeGeneratedLinkDestinations(segment) : segment;
        result += codeRun[0];

        if (inlineCodeLength === undefined) {
          inlineCodeLength = codeRun[0].length;
        } else if (inlineCodeLength === codeRun[0].length) {
          inlineCodeLength = undefined;
        }
        cursor = index + codeRun[0].length;
      }

      const remainder = line.slice(cursor);
      result +=
        inlineCodeLength === undefined ? normalizeGeneratedLinkDestinations(remainder) : remainder;
      return result;
    })
    .join("");
}

function stripBlockquoteMarkers(line: string): string {
  let content = line;
  while (true) {
    const marker = /^ {0,3}>[ \t]?/.exec(content);
    if (marker === null) {
      return content;
    }
    content = content.slice(marker[0].length);
  }
}

function normalizeGeneratedLinkDestinations(source: string): string {
  return source.replace(/\]\(([^)\s]+)\)/g, (match, destination: string) => {
    const isGeneratedLink =
      destination.startsWith(API_REFERENCE_PREFIX) ||
      destination.startsWith("./") ||
      destination.startsWith("../");
    if (!isGeneratedLink) {
      return match;
    }

    const hashIndex = destination.indexOf("#");
    const path = hashIndex === -1 ? destination : destination.slice(0, hashIndex);
    const hash = hashIndex === -1 ? "" : destination.slice(hashIndex);
    if (!path.endsWith(".mdx")) {
      return match;
    }

    const withoutExtension = path.endsWith("/index.mdx")
      ? path.slice(0, -"/index.mdx".length)
      : path.slice(0, -".mdx".length);
    return `](${withoutExtension}${hash})`;
  });
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
