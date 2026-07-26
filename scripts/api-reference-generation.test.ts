import {
  formatGeneratedEditorCommand,
  formatNavigationTitle,
  inlineGeneratedConvertOptions,
  isValidNavigationPath,
  navigationPathParts,
  normalizeGeneratedMarkdownLinks,
} from "./api-reference-generation.js";

describe("generated page post-processing", () => {
  it("expands ConvertOptions before the return value", () => {
    const source = "# Function\n\n## Parameters\n\nInput.\n\n## Returns\n\nA report.\n";
    const result = inlineGeneratedConvertOptions(source, "### width?\n\n`number`");

    expect(result).toContain("## Parameters\n\nInput.\n\n## ConvertOptions");
    expect(result).toContain("### width?\n\n`number`\n\n## Returns\n\nA report.");
  });

  it("renders EditorCommand members by kind and moves kind above the payload", () => {
    const source = `# Type Alias: EditorCommand

\`\`\`ts
type EditorCommand = First | Second;
\`\`\`

## Union Members

### Type Literal

\`\`\`ts
{
  handle: SourceHandle;
  kind: "moveShape";
  offset: {
    kind: "relative";
  };
}
\`\`\`

### Type Literal

\`\`\`ts
{
  kind: "deleteShape";
  handle: SourceHandle;
}
\`\`\`
`;
    const result = formatGeneratedEditorCommand(source);

    expect(result).not.toContain("type EditorCommand =");
    expect(result).toContain("## Commands");
    expect(result).toContain(`### \`moveShape\`

\`\`\`ts
{
  kind: "moveShape";
  handle: SourceHandle;`);
    expect(result).toContain('    kind: "relative";');
    expect(result).toContain("### `deleteShape`");
  });
});

describe("normalizeGeneratedMarkdownLinks", () => {
  it("normalizes generated page and index links while preserving anchors", () => {
    expect(
      normalizeGeneratedMarkdownLinks(
        [
          "[function](/docs/api/node/functions/convertPptxToSvg.mdx)",
          "[module](/docs/api/node/index.mdx#functions)",
          "[root](../../index.mdx)",
        ].join("\n"),
      ),
    ).toBe(
      [
        "[function](/docs/api/node/functions/convertPptxToSvg)",
        "[module](/docs/api/node#functions)",
        "[root](../..)",
      ].join("\n"),
    );
  });

  it("does not change prose, external links, or fenced code", () => {
    const source = [
      "Use `README.mdx` as a filename.",
      "Keep `[example](./page.mdx)` unchanged inside code.",
      "[external](https://example.com/guide.mdx)",
      "```md",
      "[example](/docs/api/node/example.mdx)",
      "```",
    ].join("\n");

    expect(normalizeGeneratedMarkdownLinks(source)).toBe(source);
  });

  it("handles long and blockquoted CommonMark fences", () => {
    const source = [
      "````md",
      "```ts",
      "[long fence](/docs/api/node/example.mdx)",
      "```",
      "````",
      "> ```md",
      "> [quote](/docs/api/node/example.mdx)",
      "> ```",
    ].join("\n");

    expect(normalizeGeneratedMarkdownLinks(source)).toBe(source);
  });

  it("preserves links in list fences and normalizes links after literal backticks", () => {
    const source = [
      "- ~~~md",
      "  [list fence](./inside.mdx)",
      "  ~~~",
      "Escaped \\` and unmatched ` then [outside](./outside.mdx)",
    ].join("\n");
    const expected = [
      "- ~~~md",
      "  [list fence](./inside.mdx)",
      "  ~~~",
      "Escaped \\` and unmatched ` then [outside](./outside)",
    ].join("\n");

    expect(normalizeGeneratedMarkdownLinks(source)).toBe(expected);
  });

  it("normalizes balanced, escaped, and angle-bracket destinations without touching labels", () => {
    const source = [
      "[balanced](./foo_(bar).mdx)",
      "[escaped](./foo\\(bar\\).mdx)",
      "[angle](<./foo bar.mdx>)",
      "[one-sided](./foo\\).mdx)",
      "[angle escaped](<./foo\\>.mdx>)",
      "[`label](./fake.mdx)`](./target.mdx)",
    ].join("\n");
    const expected = [
      "[balanced](./foo_(bar))",
      "[escaped](./foo\\(bar\\))",
      "[angle](<./foo bar>)",
      "[one-sided](./foo\\))",
      "[angle escaped](<./foo\\>>)",
      "[`label](./fake.mdx)`](./target)",
    ].join("\n");

    expect(normalizeGeneratedMarkdownLinks(source)).toBe(expected);
  });
});

describe("TypeDoc navigation paths", () => {
  it.each(["node/index.mdx", "node/interfaces/ConvertOptions.mdx"])(
    "accepts safe path %j",
    (path) => {
      expect(isValidNavigationPath(path)).toBe(true);
    },
  );

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
      expect(() => navigationPathParts(path)).toThrow("Invalid TypeDoc navigation path");
    },
  );
});

describe("formatNavigationTitle", () => {
  it("formats entry points without changing exported identifiers", () => {
    expect(formatNavigationTitle("node", "node/index.mdx")).toBe("Node.js entry point");
    expect(formatNavigationTitle("browser", "browser/index.mdx")).toBe("Browser entry point");
    expect(
      formatNavigationTitle("DEFAULT_FONT_MAPPING", "node/variables/DEFAULT_FONT_MAPPING.mdx"),
    ).toBe("DEFAULT_FONT_MAPPING");
  });
});
