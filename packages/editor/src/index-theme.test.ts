import { createPptx, readPptx, type SourceHandle, writePptx } from "@pptx-glimpse/document";
import { describe, expect, it } from "vitest";

import { createEditorSession } from "./index.js";

describe("EditorSession theme scheme commands", () => {
  it("applies a handle-based theme edit and restores it through undo/redo history", () => {
    const source = readPptx(writePptx(createPptx()));
    const handle = requireThemeHandle(source.themes[0]?.handle);
    const session = createEditorSession(source);

    const applied = session.updateThemeScheme(handle, {
      colorScheme: { accent1: "123456" },
      fontScheme: { major: { latin: "Brand Display" } },
    });
    expect(applied).toMatchObject({ ok: true });
    expect(session.document.themes[0]).toMatchObject({
      colorScheme: { colors: { accent1: { kind: "srgb", hex: "123456" } } },
      fontScheme: { majorLatin: "Brand Display" },
    });
    expect(session.undoDepth).toBe(1);

    expect(session.undo()).toMatchObject({ ok: true });
    expect(session.document).toBe(source);
    expect(session.redoDepth).toBe(1);

    expect(session.redo()).toMatchObject({ ok: true });
    expect(session.document.themes[0]).toMatchObject({
      colorScheme: { colors: { accent1: { kind: "srgb", hex: "123456" } } },
      fontScheme: { majorLatin: "Brand Display" },
    });
  });

  it("returns invalid-command without changing document, selection, or history", () => {
    const source = readPptx(writePptx(createPptx()));
    const session = createEditorSession(source);
    const handle = requireThemeHandle(source.themes[0]?.handle);
    const beforeSelection = session.selection;

    const result = session.apply({
      kind: "updateThemeScheme",
      handle,
      // @ts-expect-error Runtime validation covers transported/untyped commands.
      colorScheme: { accent1: "invalid" },
    });
    expect(result).toMatchObject({ ok: false, code: "invalid-command" });
    if (result.ok) throw new Error("invalid theme command unexpectedly succeeded");
    expect(result.message).toContain("6-digit hex color");
    expect(session.document).toBe(source);
    expect(session.selection).toBe(beforeSelection);
    expect(session.undoDepth).toBe(0);
    expect(session.redoDepth).toBe(0);
  });
});

function requireThemeHandle(handle: SourceHandle | undefined): SourceHandle {
  if (handle === undefined) throw new Error("fixture theme handle is missing");
  return handle;
}
