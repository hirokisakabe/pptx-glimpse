import { EDITOR_CAPABILITIES } from "./editor-capability-manifest.js";
import {
  extractCoreSessionMembers,
  extractEditorCommandKinds,
  extractNamedValueExports,
  readRepositoryCapabilitySources,
  validateEditorCapabilities,
  validateEditorCapabilityManifest,
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

  it("detects an unclassified EditorCommand", () => {
    const sources = readRepositoryCapabilitySources();
    const editorRoot = sources.editorRoot.replace(
      "export type EditorCommand =\n",
      'export interface NewCommand { readonly kind: "newCommand"; }\nexport type EditorCommand =\n  | NewCommand\n',
    );

    expect(validateEditorCapabilities({ ...sources, editorRoot })).toContain(
      "EditorCommand kinds: unclassified: newCommand",
    );
  });

  it("detects an unclassified public PptxEditorSession member", () => {
    const sources = readRepositoryCapabilitySources();
    const coreSession = sources.coreSession.replace(
      "\n}\n\nfunction buildLayoutCatalog",
      "\n  public async newCapability() {}\n}\n\nfunction buildLayoutCatalog",
    );

    expect(validateEditorCapabilities({ ...sources, coreSession })).toContain(
      "PptxEditorSession public members: unclassified: newCapability",
    );
  });

  it("detects an unclassified editor-react command", () => {
    const sources = readRepositoryCapabilitySources();
    const reactSources = `${sources.reactSources}\nconst command = { kind: "setShapeFill" };`;

    expect(validateEditorCapabilities({ ...sources, reactSources })).toContain(
      "editor-react command kinds: unclassified: setShapeFill",
    );
  });

  it("detects an unclassified editor-react direct core call", () => {
    const sources = readRepositoryCapabilitySources();
    const reactSources = `${sources.reactSources}\nsession.groupShapes([]);`;

    expect(validateEditorCapabilities({ ...sources, reactSources })).toContain(
      "editor-react direct core capability calls: unclassified: groupShapes",
    );
  });

  it("detects a stale rendered matrix", () => {
    const sources = readRepositoryCapabilitySources();
    const errors = validateEditorCapabilities({
      ...sources,
      matrixDocument: `${sources.matrixDocument}\n`,
    });

    expect(errors).toContain(
      "docs/development/editor-capability-matrix.md is stale; regenerate it from EDITOR_CAPABILITIES.",
    );
  });

  it("detects duplicate ownership and status inconsistencies in the manifest", () => {
    const first = EDITOR_CAPABILITIES[0];
    expect(
      validateEditorCapabilityManifest([
        first,
        {
          ...first,
          capability: "Duplicate ownership",
          editor: { kind: "supported", evidence: ["test"] },
        },
      ]),
    ).toEqual(
      expect.arrayContaining([
        "documentApis ownership: duplicate: createPptx, createPptxAuthoringSession",
        "Duplicate ownership: editor status is supported but row metadata is empty.",
      ]),
    );
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
  public async explicitPublic() {}
  async #privateMethod() {}
}`),
    ).toEqual(["apply", "create", "document", "explicitPublic"]);
  });
});
