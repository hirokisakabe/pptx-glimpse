import {
  extractCoreSessionMembers,
  extractEditorCommandKinds,
  extractNamedValueExports,
  readRepositoryCapabilitySources,
  validateEditorCapabilities,
} from "./validate-editor-capabilities.js";

describe("editor capability manifest", () => {
  it("matches every tracked public layer and the durable matrix", () => {
    expect(validateEditorCapabilities(readRepositoryCapabilitySources())).toEqual([]);
  });

  it("detects a new document root value export", () => {
    const sources = readRepositoryCapabilitySources();
    const errors = validateEditorCapabilities({
      ...sources,
      documentRoot: `${sources.documentRoot}\nexport { newMutation } from "./new.js";\n`,
    });

    expect(errors).toContain("document public value exports: unclassified: newMutation");
  });

  it("extracts command kinds only from the EditorCommand union", () => {
    const source = `
export interface FirstCommand { readonly kind: "first"; }
export interface UnusedCommand { readonly kind: "unused"; }
export type EditorCommand = FirstCommand;
`;

    expect(extractEditorCommandKinds(source)).toEqual(["first"]);
  });

  it("extracts public root values and public core members", () => {
    expect(
      extractNamedValueExports(
        `export type { A } from "./a.js"; export { value, old as renamed } from "./b.js";`,
      ),
    ).toEqual(["renamed", "value"]);
    expect(
      extractCoreSessionMembers(`export class PptxEditorSession {
  static async create() {}
  get document() {}
  async apply() {}
  async #privateMethod() {}
}`),
    ).toEqual(["apply", "create", "document"]);
  });
});
