import { asPartPath, asSourceNodeId } from "@pptx-glimpse/document";
import type { SourceHandle } from "pptx-glimpse";
import { describe, expect, it, vi } from "vitest";

import {
  PptxEditorController,
  type PptxEditorControllerSession,
} from "./pptx-editor-controller.js";

const CLEAN_HISTORY = {
  canUndo: false,
  canRedo: false,
  undoDepth: 0,
  redoDepth: 0,
};

const FIRST_HANDLE: SourceHandle = {
  partPath: asPartPath("ppt/slides/slide1.xml"),
  nodeId: asSourceNodeId("shape-1"),
};
const SECOND_HANDLE: SourceHandle = {
  partPath: asPartPath("ppt/slides/slide2.xml"),
  nodeId: asSourceNodeId("shape-2"),
};

class TestSession implements PptxEditorControllerSession {
  readonly slides = [
    { slideNumber: 1, svg: "<svg>one</svg>" },
    { slideNumber: 2, svg: "<svg>two</svg>" },
  ];
  selection: { readonly shapeHandle: SourceHandle } | undefined;
  history = CLEAN_HISTORY;
  saveCalls = 0;
  readonly shapeSlideNumbers: number[] = [];

  shapes(slideNumber: number) {
    this.shapeSlideNumbers.push(slideNumber);
    const handle = slideNumber === 1 ? FIRST_HANDLE : SECOND_HANDLE;
    return [
      {
        id: `shape-${slideNumber.toString()}`,
        kind: "shape" as const,
        handle,
        bounds: { x: 0, y: 0, width: 10, height: 10 },
      },
    ];
  }

  selectShape(handle: SourceHandle) {
    this.selection = { shapeHandle: handle };
  }

  save(): ReturnType<PptxEditorControllerSession["save"]> {
    this.saveCalls += 1;
    return { ok: true, pptx: new Uint8Array([1, 2, 3]), history: this.history };
  }

  markEdited() {
    this.history = {
      canUndo: true,
      canRedo: false,
      undoDepth: 1,
      redoDepth: 0,
    };
  }

  undo() {
    this.history = {
      canUndo: false,
      canRedo: true,
      undoDepth: 0,
      redoDepth: 1,
    };
  }
}

describe("PptxEditorController", () => {
  it("clamps slide selection and refreshes shapes for the active slide", () => {
    const session = new TestSession();
    const controller = new PptxEditorController(session);

    controller.selectSlide(99);

    expect(controller.getSnapshot()).toMatchObject({ currentIndex: 1, slides: session.slides });
    expect(session.shapeSlideNumbers).toEqual([1, 2]);
  });

  it("keeps busy true while work is queued and gates selection mutations", async () => {
    const session = new TestSession();
    const controller = new PptxEditorController(session);
    let finishFirst = () => {};
    const firstGate = new Promise<void>((resolve) => {
      finishFirst = resolve;
    });
    const first = controller.run(() => firstGate, { success: "First" });
    const second = controller.run(() => undefined, { success: "Second" });

    expect(controller.getSnapshot().busy).toBe(true);
    expect(controller.selectSlide(1)).toBe(false);
    expect(controller.selectShape(FIRST_HANDLE)).toBe(false);
    expect(session.selection).toBeUndefined();
    finishFirst();
    await first;
    expect(controller.getSnapshot().busy).toBe(true);
    await second;
    expect(controller.getSnapshot().busy).toBe(false);
    expect(controller.selectShape(FIRST_HANDLE)).toBe(true);
    expect(session.selection).toEqual({ shapeHandle: FIRST_HANDLE });
  });

  it("serializes mutations and publishes a coherent session snapshot", async () => {
    const session = new TestSession();
    const controller = new PptxEditorController(session);
    const events: string[] = [];
    let finishFirst = () => {};
    const firstGate = new Promise<void>((resolve) => {
      finishFirst = resolve;
    });

    const first = controller.run(
      async (activeSession) => {
        events.push("first:start");
        await firstGate;
        activeSession.markEdited();
        events.push("first:end");
      },
      { success: "First complete" },
    );
    const second = controller.run(
      () => {
        events.push("second");
      },
      { success: "Second complete" },
    );

    await vi.waitFor(() => expect(events).toEqual(["first:start"]));
    expect(controller.getSnapshot().busy).toBe(true);
    finishFirst();
    await Promise.all([first, second]);

    expect(events).toEqual(["first:start", "first:end", "second"]);
    expect(controller.getSnapshot()).toMatchObject({
      busy: false,
      dirty: true,
      message: "Second complete",
      history: { undoDepth: 1 },
    });
  });

  it("transports operation failures without rejecting the controller queue", async () => {
    const controller = new PptxEditorController(new TestSession());

    await expect(
      controller.run(
        () => {
          throw new Error("command failed");
        },
        { success: "unused" },
      ),
    ).resolves.toBe(false);
    expect(controller.getSnapshot()).toMatchObject({ error: "command failed", busy: false });
    await expect(controller.run(() => undefined, { success: "Recovered" })).resolves.toBe(true);

    expect(controller.getSnapshot()).toMatchObject({
      busy: false,
      error: "",
      message: "Recovered",
    });
  });

  it("marks the current history depth clean after save", async () => {
    const session = new TestSession();
    const controller = new PptxEditorController(session);
    await controller.run(
      (activeSession) => {
        activeSession.markEdited();
      },
      { success: "Edited" },
    );

    expect(controller.getSnapshot().dirty).toBe(true);
    const saved = await controller.save();
    expect(saved).toMatchObject({ ok: true });
    expect(session.saveCalls).toBe(1);
    expect(controller.getSnapshot().dirty).toBe(true);
    expect(controller.getSnapshot().message).toBe("Edited");
    expect(controller.markSaved(saved!.history, "PPTX downloaded")).toBe(true);
    expect(controller.getSnapshot()).toMatchObject({
      busy: false,
      dirty: false,
      message: "PPTX downloaded",
    });
  });

  it("keeps the document dirty when the host does not acknowledge a serialized save", async () => {
    const session = new TestSession();
    const controller = new PptxEditorController(session);
    await controller.run((activeSession) => activeSession.markEdited(), { success: "Edited" });

    await expect(controller.save()).resolves.toMatchObject({ ok: true });
    controller.setError("download failed");

    expect(controller.getSnapshot()).toMatchObject({
      dirty: true,
      error: "download failed",
      message: "Edited",
    });
  });

  it("uses history revision identity when a new branch returns to the saved depth", async () => {
    const session = new TestSession();
    const controller = new PptxEditorController(session);
    await controller.run((activeSession) => activeSession.markEdited(), { success: "Edited" });
    const saved = await controller.save();
    expect(controller.markSaved(saved!.history, "Saved")).toBe(true);

    await controller.run((activeSession) => activeSession.undo(), {
      success: "Undone",
      historyAction: "undo",
    });
    expect(controller.getSnapshot().dirty).toBe(true);
    await controller.run((activeSession) => activeSession.markEdited(), { success: "Replaced" });

    expect(controller.getSnapshot()).toMatchObject({
      dirty: true,
      history: { undoDepth: 1, redoDepth: 0 },
    });
  });

  it("keeps recovered operation errors while publishing committed session state", async () => {
    const session = new TestSession();
    const controller = new PptxEditorController(session);

    await expect(
      controller.run(
        () => {
          session.markEdited();
          throw new Error("preview refresh failed");
        },
        {
          success: "unused",
          recoverError: () => "Text updated; slide preview could not refresh",
        },
      ),
    ).resolves.toBe(true);

    expect(controller.getSnapshot()).toMatchObject({
      dirty: true,
      error: "preview refresh failed",
      message: "Text updated; slide preview could not refresh",
    });
  });

  it("only exposes session selection when it belongs to the active slide", () => {
    const session = new TestSession();
    const controller = new PptxEditorController(session);
    expect(controller.selectShape(FIRST_HANDLE)).toBe(true);
    expect(controller.getSnapshot().selectedShapeKey).not.toBeNull();

    expect(controller.selectSlide(1)).toBe(true);

    expect(controller.getSnapshot()).toMatchObject({ currentIndex: 1, selectedShapeKey: null });
  });
});
