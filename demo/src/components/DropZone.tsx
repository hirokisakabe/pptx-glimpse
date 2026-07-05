"use client";

import { useCallback, useRef, useState } from "react";

export function DropZone({
  compact = false,
  fontFiles,
  onFile,
  onFontFiles,
}: {
  compact?: boolean;
  fontFiles: File[];
  onFile: (file: File) => void;
  onFontFiles: (files: File[]) => void;
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
        <p className="drop-zone-title">Drop a PPTX file</p>
        <p className="drop-zone-note">Conversion runs locally in your browser.</p>
      </div>
      <div className="file-actions">
        <button className="file-label primary" onClick={() => pptxInputRef.current?.click()}>
          Choose PPTX
        </button>
        <button className="file-label secondary" onClick={() => fontInputRef.current?.click()}>
          Add fonts
        </button>
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

function isPptxFile(file: File): boolean {
  return file.name.toLowerCase().endsWith(".pptx");
}

function isFontFile(file: File): boolean {
  return /\.(?:ttf|otf|ttc)$/i.test(file.name);
}
