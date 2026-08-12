"use client";

import type {
  EditorCommand,
  PptxEditorSession,
  PptxEditorShapeInfo,
  SourceHandle,
} from "pptx-glimpse";
import { type ReactNode, useLayoutEffect, useRef, useState } from "react";

type EditorHistory = PptxEditorSession["history"];
type TextRunProperties = Extract<
  EditorCommand,
  { readonly kind: "setTextRunProperties" }
>["properties"];

export interface EditorTextRunOption {
  readonly label: string;
  readonly text: string;
  readonly handle: SourceHandle;
}

interface EditorToolbarProps {
  readonly busy: boolean;
  readonly error: string;
  readonly selectionScope: object;
  readonly selectedShape: PptxEditorShapeInfo | null;
  readonly selectedShapeKey: string | null;
  readonly slidesCount: number;
  readonly textRuns: readonly EditorTextRunOption[];
  readonly onAddTextBox: () => void;
  readonly onApplyTextProperties: (handle: SourceHandle, properties: TextRunProperties) => void;
  readonly onClearTextProperties: (handle: SourceHandle) => void;
  readonly onDeleteShape: () => void;
  readonly onDeleteSlide: () => void;
  readonly onDuplicateSlide: () => void;
  readonly onReplaceImage: (file: File) => void;
}

interface EditorHistoryToolbarProps {
  readonly busy: boolean;
  readonly children?: ReactNode;
  readonly history: EditorHistory;
  readonly onRedo: () => void;
  readonly onUndo: () => void;
}

export function EditorHistoryToolbar({
  busy,
  children,
  history,
  onRedo,
  onUndo,
}: EditorHistoryToolbarProps) {
  return (
    <div className="editor-commandbar" aria-label="Editing history">
      <button disabled={busy || !history.canUndo} type="button" onClick={onUndo}>
        Undo
      </button>
      <button disabled={busy || !history.canRedo} type="button" onClick={onRedo}>
        Redo
      </button>
      {children}
    </div>
  );
}

/** Editing-only controls. Host file, sample, font, and download actions deliberately live elsewhere. */
export function EditorToolbar({
  busy,
  error,
  selectionScope,
  selectedShape,
  selectedShapeKey,
  slidesCount,
  textRuns,
  onAddTextBox,
  onApplyTextProperties,
  onClearTextProperties,
  onDeleteShape,
  onDeleteSlide,
  onDuplicateSlide,
  onReplaceImage,
}: EditorToolbarProps) {
  const [selectedRunIndex, setSelectedRunIndex] = useState(0);
  const [fontSize, setFontSize] = useState("24");
  const [typeface, setTypeface] = useState("");
  const [color, setColor] = useState("#2454a6");
  const imageInputRef = useRef<HTMLInputElement | null>(null);
  const selectedRun = textRuns[selectedRunIndex];

  useLayoutEffect(() => {
    setSelectedRunIndex(0);
  }, [selectedShapeKey, selectionScope]);

  useLayoutEffect(() => {
    setSelectedRunIndex(0);
    setFontSize("24");
    setTypeface("");
    setColor("#2454a6");
  }, [selectionScope]);

  return (
    <aside className="editor-panel" aria-label="Editing controls" data-editor-component="toolbar">
      <div className="panel-group">
        <div className="panel-title">Slide</div>
        <div className="button-row">
          <button disabled={busy} type="button" onClick={onDuplicateSlide}>
            Duplicate
          </button>
          <button disabled={busy || slidesCount <= 1} type="button" onClick={onDeleteSlide}>
            Delete
          </button>
        </div>
      </div>

      <div className="panel-group">
        <div className="panel-title">Shape</div>
        <button disabled={busy} type="button" onClick={onAddTextBox}>
          Add text box
        </button>
        <button
          disabled={busy || selectedShape?.editableDelete !== true}
          type="button"
          onClick={onDeleteShape}
        >
          Delete shape
        </button>
        <button
          data-testid="replace-image-button"
          disabled={busy || selectedShape?.editableImageReplacement === undefined}
          type="button"
          onClick={() => imageInputRef.current?.click()}
        >
          Replace image
        </button>
        <input
          ref={imageInputRef}
          data-testid="image-replacement-input"
          disabled={busy || selectedShape?.editableImageReplacement === undefined}
          hidden
          type="file"
          accept={selectedShape?.editableImageReplacement?.accept}
          onChange={(event) => {
            const file = event.target.files?.[0];
            event.target.value = "";
            if (file !== undefined) onReplaceImage(file);
          }}
        />
      </div>

      <div className="panel-group">
        <div className="panel-title">Text</div>
        <select
          data-testid="text-run-select"
          disabled={busy || textRuns.length === 0}
          value={Math.min(selectedRunIndex, Math.max(textRuns.length - 1, 0))}
          onChange={(event) => setSelectedRunIndex(Number(event.target.value))}
        >
          {textRuns.length === 0 ? <option>No text run selected</option> : null}
          {textRuns.map((run, index) => (
            <option key={handleKey(run.handle)} value={index}>
              {run.label}
            </option>
          ))}
        </select>
        <p className="panel-note">Double-click the selected text shape or press Enter to edit.</p>
        <div className="format-toolbar" role="group" aria-label="Text style">
          <button
            disabled={busy || selectedRun === undefined}
            type="button"
            onClick={() => selectedRun && onApplyTextProperties(selectedRun.handle, { bold: true })}
          >
            B
          </button>
          <button
            disabled={busy || selectedRun === undefined}
            type="button"
            onClick={() =>
              selectedRun && onApplyTextProperties(selectedRun.handle, { italic: true })
            }
          >
            I
          </button>
          <button
            disabled={busy || selectedRun === undefined}
            type="button"
            onClick={() =>
              selectedRun && onApplyTextProperties(selectedRun.handle, { underline: true })
            }
          >
            U
          </button>
          <input
            aria-label="Text color"
            disabled={busy || selectedRun === undefined}
            type="color"
            value={color}
            onChange={(event) => {
              setColor(event.target.value);
              if (selectedRun !== undefined) {
                onApplyTextProperties(selectedRun.handle, {
                  color: { kind: "srgb", hex: event.target.value.slice(1).toUpperCase() },
                });
              }
            }}
          />
        </div>
        <div className="text-property-row">
          <input
            aria-label="Font size"
            disabled={busy || selectedRun === undefined}
            min={1}
            type="number"
            value={fontSize}
            onChange={(event) => setFontSize(event.target.value)}
          />
          <button
            disabled={busy || selectedRun === undefined || !isPositiveFiniteNumber(fontSize)}
            type="button"
            onClick={() =>
              selectedRun &&
              onApplyTextProperties(selectedRun.handle, { fontSize: pt(Number(fontSize)) })
            }
          >
            Size
          </button>
        </div>
        <div className="text-property-row">
          <input
            aria-label="Typeface"
            disabled={busy || selectedRun === undefined}
            placeholder="Typeface"
            value={typeface}
            onChange={(event) => setTypeface(event.target.value)}
          />
          <button
            disabled={busy || selectedRun === undefined || typeface.trim() === ""}
            type="button"
            onClick={() =>
              selectedRun &&
              onApplyTextProperties(selectedRun.handle, { typeface: typeface.trim() })
            }
          >
            Font
          </button>
        </div>
        <button
          disabled={busy || selectedRun === undefined}
          type="button"
          onClick={() => selectedRun && onClearTextProperties(selectedRun.handle)}
        >
          Clear style
        </button>
      </div>

      {error !== "" ? (
        <div className="error compact-error" data-testid="editor-error">
          {error}
        </div>
      ) : null}
    </aside>
  );
}

function handleKey(handle: SourceHandle): string {
  return [
    handle.partPath ?? "",
    handle.nodeId ?? "",
    handle.relationshipId ?? "",
    handle.orderingSlot === undefined ? "" : String(handle.orderingSlot),
  ].join("\u0000");
}

function pt(value: number): NonNullable<TextRunProperties["fontSize"]> {
  // This is the package-local constructor for the branded public point command field.
  // eslint-disable-next-line @typescript-eslint/no-unsafe-type-assertion
  return value as NonNullable<TextRunProperties["fontSize"]>;
}

function isPositiveFiniteNumber(value: string): boolean {
  const number = Number(value);
  return Number.isFinite(number) && number > 0;
}
