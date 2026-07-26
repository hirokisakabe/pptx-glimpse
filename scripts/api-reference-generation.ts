import { posix } from "node:path";

const API_REFERENCE_PREFIX = "/docs/api-reference/";

/**
 * Removes TypeDoc's source extension from generated API-reference links without changing
 * prose, code fences, or links outside the generated content tree.
 */
export function normalizeGeneratedMarkdownLinks(source: string): string {
  let fence: { marker: "`" | "~"; length: number } | undefined;

  return source
    .split(/(?<=\n)/)
    .map((line) => {
      const fenceMatch = /^\s*(`{3,}|~{3,})/.exec(line);
      if (fenceMatch !== null) {
        const marker = fenceMatch[1].startsWith("`") ? "`" : "~";
        if (fence === undefined) {
          fence = { marker, length: fenceMatch[1].length };
        } else if (fence.marker === marker && fenceMatch[1].length >= fence.length) {
          fence = undefined;
        }
        return line;
      }
      if (fence !== undefined) {
        return line;
      }

      return line.replace(/\]\(([^)\s]+)\)/g, (match, destination: string) => {
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
    })
    .join("");
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
