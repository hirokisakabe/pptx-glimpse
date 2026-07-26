import {
  isValidNavigationPath,
  navigationPathParts,
  normalizeGeneratedMarkdownLinks,
} from "./api-reference-generation.js";

describe("normalizeGeneratedMarkdownLinks", () => {
  it("normalizes generated page and index links while preserving anchors", () => {
    expect(
      normalizeGeneratedMarkdownLinks(
        [
          "[function](/docs/api-reference/node/functions/convertPptxToSvg.mdx)",
          "[module](/docs/api-reference/node/index.mdx#functions)",
          "[root](../../index.mdx)",
        ].join("\n"),
      ),
    ).toBe(
      [
        "[function](/docs/api-reference/node/functions/convertPptxToSvg)",
        "[module](/docs/api-reference/node#functions)",
        "[root](../..)",
      ].join("\n"),
    );
  });

  it("does not change prose, external links, or fenced code", () => {
    const source = [
      "Use `README.mdx` as a filename.",
      "[external](https://example.com/guide.mdx)",
      "```md",
      "[example](/docs/api-reference/node/example.mdx)",
      "```",
    ].join("\n");

    expect(normalizeGeneratedMarkdownLinks(source)).toBe(source);
  });
});

describe("TypeDoc navigation paths", () => {
  it("uses POSIX path semantics", () => {
    expect(navigationPathParts("node/interfaces/ConvertOptions.mdx")).toEqual({
      directory: "node/interfaces",
      name: "ConvertOptions",
    });
  });

  it.each(["", "/node/index.mdx", "../outside.mdx", "node/../outside.mdx", "node\\index.mdx"])(
    "rejects unsafe path %j",
    (path) => {
      expect(isValidNavigationPath(path)).toBe(false);
    },
  );
});
