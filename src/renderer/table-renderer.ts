import type { TableElement } from "../model/table.js";
import type { Transform } from "../model/shape.js";
import { emuToPixels } from "../utils/emu.js";
import { buildTransformAttr } from "./transform.js";
import { renderFillAttrs, renderOutlineAttrs } from "./fill-renderer.js";
import { renderTextBody } from "./text-renderer.js";

export function renderTable(element: TableElement, defs: string[]): string {
  const { transform, table } = element;
  const transformAttr = buildTransformAttr(transform);

  const colWidths = table.columns.map((col) => emuToPixels(col.width));
  const rowHeights = table.rows.map((row) => emuToPixels(row.height));

  const parts: string[] = [];
  parts.push(`<g transform="${transformAttr}">`);

  let y = 0;
  for (let rowIdx = 0; rowIdx < table.rows.length; rowIdx++) {
    const row = table.rows[rowIdx];
    const rowH = rowHeights[rowIdx];

    let x = 0;
    let colIdx = 0;
    for (const cell of row.cells) {
      if (cell.hMerge || cell.vMerge) {
        x += colWidths[colIdx] ?? 0;
        colIdx++;
        continue;
      }

      const cellW = computeSpannedSize(colWidths, colIdx, cell.gridSpan);
      const cellH = computeSpannedSize(rowHeights, rowIdx, cell.rowSpan);

      // Cell background
      const fillResult = renderFillAttrs(cell.fill);
      if (fillResult.defs) defs.push(fillResult.defs);
      parts.push(
        `<rect x="${x}" y="${y}" width="${cellW}" height="${cellH}" ${fillResult.attrs}/>`,
      );

      // Cell borders
      if (cell.borders) {
        if (cell.borders.top) {
          const attrs = renderOutlineAttrs(cell.borders.top);
          parts.push(`<line x1="${x}" y1="${y}" x2="${x + cellW}" y2="${y}" ${attrs}/>`);
        }
        if (cell.borders.bottom) {
          const attrs = renderOutlineAttrs(cell.borders.bottom);
          parts.push(
            `<line x1="${x}" y1="${y + cellH}" x2="${x + cellW}" y2="${y + cellH}" ${attrs}/>`,
          );
        }
        if (cell.borders.left) {
          const attrs = renderOutlineAttrs(cell.borders.left);
          parts.push(`<line x1="${x}" y1="${y}" x2="${x}" y2="${y + cellH}" ${attrs}/>`);
        }
        if (cell.borders.right) {
          const attrs = renderOutlineAttrs(cell.borders.right);
          parts.push(
            `<line x1="${x + cellW}" y1="${y}" x2="${x + cellW}" y2="${y + cellH}" ${attrs}/>`,
          );
        }
      }

      // Cell text
      if (cell.textBody) {
        const cellTransform: Transform = {
          offsetX: 0,
          offsetY: 0,
          extentWidth: pixelsToEmu(cellW),
          extentHeight: pixelsToEmu(cellH),
          rotation: 0,
          flipH: false,
          flipV: false,
        };
        const textSvg = renderTextBody(cell.textBody, cellTransform);
        if (textSvg) {
          parts.push(`<g transform="translate(${x}, ${y})">${textSvg}</g>`);
        }
      }

      x += colWidths[colIdx] ?? 0;
      colIdx++;
    }

    y += rowH;
  }

  parts.push("</g>");
  return parts.join("");
}

function computeSpannedSize(sizes: number[], startIdx: number, span: number): number {
  let total = 0;
  for (let i = startIdx; i < startIdx + span && i < sizes.length; i++) {
    total += sizes[i];
  }
  return total;
}

function pixelsToEmu(px: number): number {
  return (px / 96) * 914400;
}
