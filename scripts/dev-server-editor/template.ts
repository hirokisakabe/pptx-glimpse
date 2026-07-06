import { createDevEditorClientScript } from "./client/index.js";
import { DEV_EDITOR_STYLES } from "./styles.js";

interface DevEditorSlide {
  readonly slideNumber: number;
  readonly svg: string;
  readonly width?: number;
  readonly height?: number;
  readonly handle?: unknown;
}

interface DevEditorHtmlOptions {
  readonly slides: readonly DevEditorSlide[];
  readonly pptxName: string;
  readonly emuPerPixel: number;
  readonly maxImageReplacementBytes: number;
}

export function generateDevEditorHtml(options: DevEditorHtmlOptions): string {
  const thumbnailsHtml = generateThumbnailsHtml(options.slides);
  const firstSvg = options.slides.length > 0 ? options.slides[0].svg : "<p>No slides</p>";
  const clientScript = createDevEditorClientScript({
    slides: options.slides,
    slideCount: options.slides.length,
    emuPerPixel: options.emuPerPixel,
    maxImageReplacementBytes: options.maxImageReplacementBytes,
  });

  return `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>pptx-glimpse dev - ${escapeHtml(options.pptxName)}</title>
  <style>
${DEV_EDITOR_STYLES}\
  </style>
</head>
<body>
  <div id="header">
    <h1>pptx-glimpse dev &mdash; ${escapeHtml(options.pptxName)}</h1>
    <span id="status">Connected</span>
  </div>
  <div id="main">
    <div id="sidebar">${thumbnailsHtml}</div>
    <div id="viewer">
      <div id="slide-container">${firstSvg}</div>
    </div>
    <div id="editor-panel">
      <label>Text run<select id="text-run-select"></select></label>
      <label>Text<input id="text-run-input" type="text"></label>
      <button id="apply-text-button" type="button">Apply</button>
      <label>Image<input id="image-replacement-input" data-testid="image-replacement-input" type="file" disabled></label>
      <div id="editor-actions">
        <button id="add-text-box-button" type="button">Add text box</button>
        <button id="add-connector-button" type="button">Add connector</button>
        <button id="delete-shape-button" type="button" disabled>Delete shape</button>
        <button id="undo-button" type="button">Undo</button>
        <button id="redo-button" type="button">Redo</button>
        <button id="save-button" type="button">Save</button>
      </div>
      <div id="editor-message"></div>
    </div>
  </div>
  <div id="info">Slide 1 / ${String(options.slides.length)}</div>
  <script>
${clientScript}\
  </script>
</body>
</html>`;
}

function generateThumbnailsHtml(slides: readonly DevEditorSlide[]): string {
  return slides
    .map(
      (s, i) =>
        `<div class="thumbnail${i === 0 ? " active" : ""}" data-index="${i}">` +
        `<div class="thumb-label"><span class="thumb-title">Slide ${String(s.slideNumber)}</span>` +
        `<span class="thumb-actions">` +
        `<button class="thumb-action" data-testid="duplicate-slide-${String(i)}" data-action="duplicate" data-index="${String(i)}" type="button" title="Duplicate slide">D</button>` +
        `<button class="thumb-action" data-testid="delete-slide-${String(i)}" data-action="delete" data-index="${String(i)}" type="button" title="Delete slide"${slides.length <= 1 ? " disabled" : ""}>X</button>` +
        `</span></div>` +
        `<div class="thumb-svg">${s.svg}</div>` +
        `</div>`,
    )
    .join("");
}

function escapeHtml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
