export const DEV_EDITOR_STYLES = `    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background: #1a1a2e;
      color: #e0e0e0;
    }
    #header {
      padding: 12px 20px;
      background: #16213e;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }
    #header h1 { font-size: 14px; font-weight: 600; color: #a0a0c0; }
    #status { font-size: 12px; color: #4caf50; }
    #status.rendering { color: #ff9800; }
    #status.error { color: #f44336; }
    #main { display: flex; height: calc(100vh - 48px); }
    #sidebar {
      width: 180px;
      overflow-y: auto;
      background: #16213e;
      padding: 8px;
      border-right: 1px solid #2a2a4a;
    }
    #editor-panel {
      width: 260px;
      background: #111827;
      border-left: 1px solid #2a2a4a;
      padding: 12px;
      display: flex;
      flex-direction: column;
      gap: 10px;
    }
    #editor-panel label {
      display: flex;
      flex-direction: column;
      gap: 4px;
      color: #a0a0c0;
      font-size: 11px;
      font-weight: 600;
    }
    #editor-panel select,
    #editor-panel input {
      width: 100%;
      min-height: 32px;
      border: 1px solid #334155;
      border-radius: 4px;
      background: #0f172a;
      color: #e5e7eb;
      padding: 6px 8px;
      font: inherit;
    }
    #editor-panel button {
      min-height: 32px;
      border: 1px solid #475569;
      border-radius: 4px;
      background: #1f2937;
      color: #f8fafc;
      cursor: pointer;
      font: inherit;
      font-weight: 600;
    }
    #editor-panel button:hover:not(:disabled) { background: #334155; }
    #editor-panel button:disabled {
      opacity: 0.45;
      cursor: not-allowed;
    }
    #editor-actions {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 8px;
    }
    #save-button { grid-column: 1 / -1; }
    #editor-message {
      min-height: 18px;
      color: #94a3b8;
      font-size: 11px;
      overflow-wrap: anywhere;
    }
    .thumbnail {
      margin-bottom: 8px;
      padding: 4px;
      border: 2px solid transparent;
      border-radius: 4px;
      cursor: pointer;
      background: #fff;
    }
    .thumbnail.active { border-color: #4472c4; }
    .thumbnail:hover { border-color: #6090d0; }
    .thumb-label {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 4px;
      font-size: 10px;
      color: #888;
      padding: 2px 0;
      background: #16213e;
    }
    .thumb-title { min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .thumb-actions { display: flex; gap: 3px; }
    .thumb-action {
      min-width: 22px;
      min-height: 20px;
      border: 1px solid #334155;
      border-radius: 3px;
      background: #0f172a;
      color: #e5e7eb;
      cursor: pointer;
      font-size: 10px;
      font-weight: 700;
      line-height: 1;
    }
    .thumb-action:hover:not(:disabled) { background: #1d4ed8; }
    .thumb-action:disabled { opacity: 0.4; cursor: not-allowed; }
    .thumb-svg svg { width: 100%; height: auto; display: block; }
    #viewer {
      flex: 1;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 20px;
      overflow: auto;
    }
    #slide-container {
      background: #fff;
      border-radius: 4px;
      box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
      max-width: 100%;
      max-height: 100%;
      position: relative;
    }
    #slide-container svg { display: block; width: 100%; height: auto; }
    #selection-overlay {
      position: absolute;
      inset: 0;
      width: 100%;
      height: 100%;
      overflow: visible;
      touch-action: none;
      z-index: 2;
    }
    #text-editor-overlay {
      position: absolute;
      min-width: 40px;
      min-height: 28px;
      padding: 4px;
      border: 1px solid #2563eb;
      background: rgba(255, 255, 255, 0.96);
      color: #111827;
      z-index: 3;
      overflow: hidden;
      box-shadow: 0 0 0 1px rgba(37, 99, 235, 0.25);
    }
    .text-editor-paragraph {
      min-height: 20px;
      line-height: 1.25;
      white-space: pre-wrap;
      overflow-wrap: anywhere;
    }
    .text-editor-run {
      outline: none;
      white-space: pre-wrap;
    }
    .text-run-format-toolbar {
      display: grid;
      grid-template-columns: repeat(3, 24px) 44px 20px 28px 20px minmax(48px, 1fr) 20px;
      gap: 3px;
      align-items: center;
      margin-bottom: 4px;
    }
    .text-run-format-toolbar button,
    .text-run-format-toolbar input {
      min-height: 22px;
      border: 1px solid #94a3b8;
      border-radius: 4px;
      background: #f8fafc;
      color: #0f172a;
      font-size: 10px;
      font-weight: 650;
    }
    .text-run-format-toolbar button {
      padding: 0;
      min-width: 0;
    }
    .text-run-format-toolbar button[aria-pressed="true"] {
      background: #dbeafe;
      border-color: #2563eb;
    }
    .text-run-format-toolbar input {
      min-width: 0;
      padding: 2px 4px;
      font-weight: 500;
    }
    .text-run-format-toolbar input[type="color"] {
      padding: 1px;
    }
    .text-run-format-toolbar :disabled {
      opacity: 0.45;
      cursor: not-allowed;
    }
    .text-editor-actions {
      display: flex;
      justify-content: flex-end;
      gap: 4px;
      margin-top: 4px;
    }
    .text-editor-actions button {
      min-height: 24px;
      padding: 2px 8px;
      border: 1px solid #64748b;
      border-radius: 4px;
      background: #f8fafc;
      color: #0f172a;
      font-size: 11px;
      font-weight: 600;
    }
    .shape-hit-area {
      cursor: move;
      fill: transparent;
      pointer-events: all;
    }
    .selection-box {
      fill: none;
      stroke: #0f766e;
      stroke-width: 1.5;
      vector-effect: non-scaling-stroke;
      pointer-events: none;
    }
    .selection-handle {
      fill: #f8fafc;
      stroke: #0f766e;
      stroke-width: 1.5;
      cursor: nwse-resize;
      vector-effect: non-scaling-stroke;
      pointer-events: all;
    }
    .selection-handle[data-handle="ne"],
    .selection-handle[data-handle="sw"] {
      cursor: nesw-resize;
    }
    #info {
      padding: 4px 20px;
      background: #16213e;
      font-size: 11px;
      color: #666;
      text-align: center;
    }
`;
