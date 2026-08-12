import type { SourceHandle } from "pptx-glimpse";
import { describe, expect, it, vi } from "vitest";

import {
  EditorController,
  type EditorControllerSession,
} from "../demo/src/components/use-editor-controller.js";

const CLEAN_HISTORY = {
  canUndo: false,
  canRedo: false,
  undoDepth: 0,
  redoDepth: 0,
};

class TestSession implements EditorControllerSession {
  readonly slides = [
    { slideNumber: 1, svg: "<svg>one</svg>" },
    { slideNumber: 2, svg: "<svg>two</svg>" },
  ];
  readonly selection = undefined;
  history = CLEAN_HISTORY;
  saveCalls = 0;
  readonly shapeSlideNumbers: number[] = [];

  shapes(slideNumber: number) {
    this.shapeSlideNumbers.push(slideNumber);
    return [];
  }

  selectShape(_handle: SourceHandle) {}

  save(): ReturnType<EditorControllerSession["save"]> {
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
}

describe("EditorController", () => {
  it("clamps slide selection and refreshes shapes for the active slide", () => {
    const session = new TestSession();
    const controller = new EditorController(session);

    controller.selectSlide(99);

    expect(controller.getSnapshot()).toMatchObject({ currentIndex: 1, slides: session.slides });
    expect(session.shapeSlideNumbers).toEqual([1, 2]);
  });

  it("serializes mutations and publishes a coherent session snapshot", async () => {
    const session = new TestSession();
    const controller = new EditorController(session);
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
    const controller = new EditorController(new TestSession());

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
    const controller = new EditorController(session);
    await controller.run(
      (activeSession) => {
        activeSession.markEdited();
      },
      { success: "Edited" },
    );

    expect(controller.getSnapshot().dirty).toBe(true);
    await expect(controller.save()).resolves.toMatchObject({ ok: true });
    expect(session.saveCalls).toBe(1);
    expect(controller.getSnapshot()).toMatchObject({
      busy: false,
      dirty: false,
      message: "PPTX downloaded",
    });
  });
});
