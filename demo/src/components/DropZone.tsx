"use client";

import { useCallback, useRef, useState } from "react";

export function DropZone({
  compact = false,
  fontFiles,
  onFile,
  onFontFiles,
  onSample,
}: {
  compact?: boolean;
  fontFiles: File[];
  onFile: (file: File) => void;
  onFontFiles: (files: File[]) => void;
  onSample: (sample: SamplePptx, mode: SampleOpenMode) => void;
}) {
  const [dragOver, setDragOver] = useState(false);
  const pptxInputRef = useRef<HTMLInputElement>(null);
  const fontInputRef = useRef<HTMLInputElement>(null);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragOver(false);
      const file = e.dataTransfer.files[0];
      if (file && isPptxFile(file)) {
        onFile(file);
      }
    },
    [onFile],
  );

  const handlePptxChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (file) onFile(file);
    },
    [onFile],
  );

  const handleFontChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      onFontFiles(Array.from(e.target.files ?? []).filter(isFontFile));
    },
    [onFontFiles],
  );

  return (
    <div
      className={`drop-zone${dragOver ? " drag-over" : ""}${compact ? " compact" : ""}`}
      data-testid="drop-zone"
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={handleDrop}
    >
      <div className="drop-zone-copy">
        <p className="drop-zone-title">Open a PPTX in this browser</p>
        <p className="drop-zone-note">
          View, edit, and save without sending your file to a server.
        </p>
      </div>
      <div className="file-actions">
        <button
          className="file-label primary"
          type="button"
          onClick={() => pptxInputRef.current?.click()}
        >
          Choose PPTX
        </button>
        <button
          className="file-label edit-sample"
          data-testid="sample-edit-basic-theme"
          type="button"
          onClick={() => onSample(SAMPLE_PPTX_FILES[0], "edit")}
        >
          Edit a sample
        </button>
        <button
          className="file-label secondary"
          type="button"
          onClick={() => fontInputRef.current?.click()}
        >
          Add fonts
        </button>
        <div className="sample-actions" role="group" aria-label="Sample PPTX files">
          {SAMPLE_PPTX_FILES.map((sample) => (
            <button
              className="file-label secondary"
              data-testid={`sample-${sample.id}`}
              key={sample.id}
              type="button"
              onClick={() => onSample(sample, "view")}
            >
              {sample.label}
            </button>
          ))}
        </div>
      </div>
      {fontFiles.length > 0 ? (
        <p className="font-count" data-testid="font-count">
          {fontFiles.length} font file{fontFiles.length === 1 ? "" : "s"} ready
        </p>
      ) : null}
      <input
        ref={pptxInputRef}
        data-testid="pptx-input"
        type="file"
        accept=".pptx"
        hidden
        onChange={handlePptxChange}
      />
      <input
        ref={fontInputRef}
        data-testid="font-input"
        type="file"
        accept=".ttf,.otf,.ttc,font/ttf,font/otf"
        hidden
        multiple
        onChange={handleFontChange}
      />
    </div>
  );
}

export interface SamplePptx {
  readonly id: string;
  readonly label: string;
  readonly filename: string;
  readonly href: string;
}

export type SampleOpenMode = "view" | "edit";

export const SAMPLE_PPTX_FILES: readonly SamplePptx[] = [
  {
    id: "editor-demo",
    label: "View sample",
    filename: "editor-demo.pptx",
    href: "/samples/editor-demo.pptx",
  },
];

function isPptxFile(file: File): boolean {
  return file.name.toLowerCase().endsWith(".pptx");
}

function isFontFile(file: File): boolean {
  return /\.(?:ttf|otf|ttc)$/i.test(file.name);
}
