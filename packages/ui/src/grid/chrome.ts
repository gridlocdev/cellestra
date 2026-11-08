import { DataScope, type CellRange, type DataValidation, type Worksheet } from '@cellestra/core';
import type { GridPalette } from './palette.ts';
import { HEADER_HEIGHT, HEADER_WIDTH, type Rect, type Viewport } from './viewport.ts';

/** A clickable control the grid draws over cells: a filter chevron or a validation-list arrow. */
export type ChromeButton =
  | {
      readonly kind: 'filter';
      readonly rect: Rect;
      readonly column: number;
      /** The table or AutoFilter range the chevron belongs to. */
      readonly range: CellRange;
      readonly hasHeader: boolean;
      /** This column currently has a filter condition (drawn filled). */
      readonly active: boolean;
    }
  | { readonly kind: 'validation'; readonly rect: Rect; readonly rule: DataValidation };

/**
 * The buttons to draw this frame, in screen px. Tables with a header row and the sheet AutoFilter
 * get a chevron in each visible header cell; a list-validated active cell gets a dropdown arrow.
 */
export function chromeButtons(
  sheet: Worksheet,
  v: Viewport,
  active: { readonly row: number; readonly column: number },
): ChromeButton[] {
  const out: ChromeButton[] = [];
  const visibleRow = (row: number) => v.visibleRows.some((b) => b.index === row);
  const regions: { range: CellRange; hasHeader: boolean }[] = sheet.tables
    .filter((t) => t.hasHeader)
    .map((t) => ({ range: t.range, hasHeader: true }));
  const auto = sheet.activeFilter;
  if (auto && !sheet.tables.some((t) => t.range.equals(auto))) regions.push({ range: auto, hasHeader: true });

  for (const { range, hasHeader } of regions) {
    if (!visibleRow(range.top)) continue;
    const spec = DataScope.effectiveFilter(sheet, range);
    for (const band of v.visibleColumns) {
      const column = band.index;
      if (column < range.left || column > range.right) continue;
      const cell = v.rangeRect(range.top, column, range.top, column);
      if (!onScreen(cell)) continue;
      const size = Math.min(14, cell.h - 2);
      const rect = { x: cell.x + cell.w - size - 2, y: cell.y + (cell.h - size) / 2, w: size, h: size };
      if (!inPane(v, range.top, column, rect)) continue;
      const cf = spec?.columns.get(column);
      const isActive = !!cf && (cf.allowedValues != null || (cf.condition ?? '') !== '');
      out.push({ kind: 'filter', rect, column, range, hasHeader, active: isActive });
    }
  }

  const rule = sheet.findValidation(active.row, active.column);
  if (rule?.kind === 'list') {
    const cell = v.rangeRect(active.row, active.column, active.row, active.column);
    if (onScreen(cell)) {
      // Inside the cell (clicks on the next column must select it) and clear of the fill handle.
      const size = Math.min(14, cell.h - 2);
      const rect = { x: cell.x + cell.w - size - 5, y: cell.y + (cell.h - size) / 2, w: size, h: size };
      if (inPane(v, active.row, active.column, rect)) out.push({ kind: 'validation', rect, rule });
    }
  }
  return out;
}

export function hitChrome(buttons: readonly ChromeButton[], x: number, y: number): ChromeButton | null {
  for (const b of buttons) {
    const r = b.rect;
    if (x >= r.x && x <= r.x + r.w && y >= r.y && y <= r.y + r.h) return b;
  }
  return null;
}

export function drawChrome(
  ctx: CanvasRenderingContext2D,
  buttons: readonly ChromeButton[],
  palette: GridPalette,
): void {
  ctx.save();
  ctx.beginPath();
  ctx.rect(HEADER_WIDTH, HEADER_HEIGHT, ctx.canvas.width, ctx.canvas.height);
  ctx.clip();
  for (const b of buttons) {
    const { x, y, w, h } = b.rect;
    const filled = b.kind === 'filter' && b.active;
    ctx.fillStyle = filled ? palette.activeOutline : palette.headerFill;
    ctx.fillRect(x, y, w, h);
    ctx.strokeStyle = filled
      ? palette.activeOutline
      : b.kind === 'filter'
        ? palette.tableHeaderBand
        : palette.gridLine;
    ctx.lineWidth = 1;
    ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);
    // A small chevron.
    ctx.strokeStyle = filled ? '#ffffff' : palette.headerText;
    ctx.lineWidth = 1.4;
    ctx.beginPath();
    const cx = x + w / 2;
    const cy = y + h / 2;
    const s = Math.max(2, w * 0.2);
    ctx.moveTo(cx - s, cy - s / 2);
    ctx.lineTo(cx, cy + s / 2);
    ctx.lineTo(cx + s, cy - s / 2);
    ctx.stroke();
  }
  ctx.restore();
}

function onScreen(r: Rect): boolean {
  return r.x + r.w > HEADER_WIDTH && r.y + r.h > HEADER_HEIGHT;
}

/**
 * Whether a button sits wholly inside its cell's pane. A scrolled cell partly under the frozen
 * rows or columns would otherwise draw its button over the frozen cells and take their clicks.
 */
function inPane(v: Viewport, row: number, column: number, r: Rect): boolean {
  const left = column < v.frozenColumns ? HEADER_WIDTH : v.paneX;
  const top = row < v.frozenRows ? HEADER_HEIGHT : v.paneY;
  return r.x >= left && r.y >= top;
}
