import { readFile } from "node:fs/promises";
import { resolve } from "node:path";

import { describe, expect, it } from "vitest";

const demoComponentsRoot = resolve(import.meta.dirname, "../demo/src/components");
const packageRoot = resolve(import.meta.dirname, "../packages/editor-react");
const packageSourceRoot = resolve(packageRoot, "src");
const reusableComponentFiles = [
  "PptxEditor.tsx",
  "EditorToolbar.tsx",
  "EditorSlideStrip.tsx",
  "direct-text-editor-lifecycle.ts",
  "pptx-editor-controller.ts",
  "use-pptx-editor-controller.ts",
];

describe("editor React package boundary", () => {
  it("keeps package components independent from Next.js, demo source, and lower packages", async () => {
    for (const fileName of reusableComponentFiles) {
      const source = await readFile(resolve(packageSourceRoot, fileName), "utf8");
      expect(source, fileName).not.toMatch(/from ["']next(?:\/|["'])/);
      expect(source, fileName).not.toContain("DemoEditorShell");
      expect(source, fileName).not.toMatch(
        /from ["']@pptx-glimpse\/(?:document|editor|renderer|mcp)(?:\/|["'])/,
      );
      expect(source, fileName).not.toMatch(/from ["']pptx-glimpse\//);
    }
  });

  it("declares browser peers and an explicit stylesheet export", async () => {
    const packageJson = await readFile(resolve(packageRoot, "package.json"), "utf8");

    expect(packageJson).toContain('"private": true');
    expect(packageJson).toContain('"pptx-glimpse": "*"');
    expect(packageJson).toContain('"react": "^19.2.0"');
    expect(packageJson).toContain('"react-dom": "^19.2.0"');
    expect(packageJson).toContain('"./styles.css": "./dist/styles.css"');
  });

  it("composes file-oriented host policy around the package editor", async () => {
    const workspace = await readFile(resolve(demoComponentsRoot, "EditorWorkspace.tsx"), "utf8");
    const shell = await readFile(resolve(demoComponentsRoot, "DemoEditorShell.tsx"), "utf8");
    const editor = await readFile(resolve(packageSourceRoot, "PptxEditor.tsx"), "utf8");

    expect(workspace).toContain('from "@pptx-glimpse/editor-react"');
    expect(workspace).toContain("<PptxEditor session={editor}>");
    expect(workspace).toContain("<DemoEditorShell");
    for (const hostResponsibility of [
      "Open PPTX",
      "Open sample",
      "Add fonts",
      "Download PPTX",
      "Presentation file name",
      "beforeunload",
    ]) {
      expect(shell).toContain(hostResponsibility);
      expect(editor).not.toContain(hostResponsibility);
    }
  });

  it("scopes every distributed stylesheet selector beneath the package root", async () => {
    const stylesheet = await readFile(resolve(packageSourceRoot, "styles.css"), "utf8");
    const demoGlobals = await readFile(
      resolve(import.meta.dirname, "../demo/src/app/globals.css"),
      "utf8",
    );

    expect(stylesheet).toContain(".pptx-glimpse-editor");
    expect(stylesheet).not.toMatch(/(?:^|\n)(?:body|:root|\.editor-|\.selection-|\.button-row)/);
    expect(demoGlobals).not.toContain(".editor-commandbar");
    expect(demoGlobals).not.toContain(".editor-slide-frame");
    expect(demoGlobals).not.toContain(".selection-box");
  });

  it("resets toolbar-local run selection when the editor session identity changes", async () => {
    const toolbar = await readFile(resolve(packageSourceRoot, "EditorToolbar.tsx"), "utf8");
    const editor = await readFile(resolve(packageSourceRoot, "PptxEditor.tsx"), "utf8");

    expect(toolbar).toContain("readonly selectionScope: object");
    expect(toolbar).toContain('setFontSize("24")');
    expect(toolbar).toContain('setTypeface("")');
    expect(toolbar).toContain('setColor("#2454a6")');
    expect(toolbar).toContain("[selectedShapeKey, selectionScope]");
    expect(toolbar).toContain("[selectionScope]");
    expect(editor).toContain("selectionScope={controller}");
  });

  it("keeps host state and transient gesture lifecycle scoped to the active session", async () => {
    const shell = await readFile(resolve(demoComponentsRoot, "DemoEditorShell.tsx"), "utf8");
    const editor = await readFile(resolve(packageSourceRoot, "PptxEditor.tsx"), "utf8");
    const slideStrip = await readFile(resolve(packageSourceRoot, "EditorSlideStrip.tsx"), "utf8");

    expect(editor).toContain("resetScope: controller");
    expect(shell).toContain("[controls.resetScope, fileName]");
    expect(shell).not.toContain("}, [controls]);");
    expect(shell).toContain("[controls.hasUnsavedChanges]");
    expect(editor).not.toContain("directTextSessionRef");
    expect(editor).toContain("committedInteractionScopeRef.current = controller");
    expect(editor).toContain("interactionScope={controller}");
    expect(slideStrip).toContain("drag.interactionScope !== interactionScope");
  });
});
