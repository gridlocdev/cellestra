import {
  Argb,
  Defaults,
  SheetDimensions,
  type BorderEdge,
  type Cell,
  type CellStyle,
  type CellValue,
  StylePool,
  type Worksheet,
} from '@cellestra/core';
import { drawChrome, type ChromeButton } from './chrome.ts';
import type { GridPalette } from './palette.ts';
import { HEADER_HEIGHT, HEADER_WIDTH, type Band, type Rect, type Viewport } from './viewport.ts';

export const CELL_PAD_X = 5;
export const CELL_PAD_Y = 2;
const PT_TO_PX = 96 / 72;

/** Selection as the renderer needs it: the active cell plus inclusive ranges. */
export interface RenderSelection {
  readonly row: number;
  readonly column: number;
  readonly ranges: readonly { top: number; left: number; bottom: number; right: number }[];
}

export interface RenderInput {
  readonly ctx: CanvasRenderingContext2D;
  readonly viewport: Viewport;
  readonly sheet: Worksheet;
  readonly styles: StylePool;
  readonly palette: GridPalette;
  readonly selection: RenderSelection;
  /** Formatted text for a cell's value (number formats, dates, locale). */
  readonly display: (value: CellValue, style: CellStyle) => DisplayText;
  /** Hide the content of the cell under the inline editor. */
  readonly editing?: { row: number; column: number } | null;
  readonly showFormulas?: boolean;
  /** Filter chevrons and the validation-list arrow, from `chromeButtons`. */
  readonly chrome?: readonly ChromeButton[];
}

export interface DisplayText {
  readonly text: string;
  /** Explicit colour from the number format (`[Red]`), as CSS. */
  readonly color?: string;
}

/** Draw one frame. Everything is in CSS px; the caller has already scaled for devicePixelRatio. */
export function renderGrid(input: RenderInput): void {
  const { ctx, viewport: v, palette } = input;
  ctx.save();
  ctx.fillStyle = palette.canvas;
  ctx.fillRect(0, 0, v.width, v.height);

  // Quadrants draw separately so frozen panes clip scrolled content correctly.
  const fc = v.visibleColumns.filter((b) => b.frozen);
  const sc = v.visibleColumns.filter((b) => !b.frozen);
  const fr = v.visibleRows.filter((b) => b.frozen);
  const sr = v.visibleRows.filter((b) => !b.frozen);
  const paneX = v.paneX;
  const paneY = v.paneY;
  drawQuadrant(input, sr, sc, { x: paneX, y: paneY, w: v.width - paneX, h: v.height - paneY });
  if (fr.length)
    drawQuadrant(input, fr, sc, { x: paneX, y: HEADER_HEIGHT, w: v.width - paneX, h: paneY - HEADER_HEIGHT });
  if (fc.length)
    drawQuadrant(input, sr, fc, { x: HEADER_WIDTH, y: paneY, w: paneX - HEADER_WIDTH, h: v.height - paneY });
  if (fr.length && fc.length)
    drawQuadrant(input, fr, fc, {
      x: HEADER_WIDTH,
      y: HEADER_HEIGHT,
      w: paneX - HEADER_WIDTH,
      h: paneY - HEADER_HEIGHT,
    });

  if (input.chrome?.length) drawChrome(ctx, input.chrome, palette);
  drawHeaders(input);

  if (v.frozenRows > 0 || v.frozenColumns > 0) {
    ctx.strokeStyle = palette.freezeDivider;
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    if (v.frozenRows > 0) line(ctx, 0, paneY, v.width, paneY);
    if (v.frozenColumns > 0) line(ctx, paneX, 0, paneX, v.height);
    ctx.stroke();
  }
  ctx.restore();
}

function drawQuadrant(input: RenderInput, rows: readonly Band[], cols: readonly Band[], clip: Rect): void {
  if (!rows.length || !cols.length || clip.w <= 0 || clip.h <= 0) return;
  const { ctx, sheet, styles, palette, viewport: v } = input;
  ctx.save();
  ctx.beginPath();
  ctx.rect(clip.x, clip.y, clip.w, clip.h);
  ctx.clip();

  const top = rows[0]!.index;
  const bottom = rows[rows.length - 1]!.index;
  const left = cols[0]!.index;
  const right = cols[cols.length - 1]!.index;
  const merges = sheet.merges.filter(
    (m) => m.range.bottom >= top && m.range.top <= bottom && m.range.right >= left && m.range.left <= right,
  );
  const coveredByMerge = (r: number, c: number) => merges.some((m) => m.covers(r, c));

  // 1. Fills (row / column default styles included via effectiveStyleId), table bands.
  for (const rb of rows) {
    for (const cb of cols) {
      if (coveredByMerge(rb.index, cb.index)) continue;
      const style = styles.get(sheet.effectiveStyleId(rb.index, cb.index));
      const fill = style.fill.background;
      if (fill !== null) {
        ctx.fillStyle = Argb.toCss(fill);
        ctx.fillRect(cb.start, rb.start, cb.size, rb.size);
      }
    }
  }
  for (const t of sheet.tables) {
    const r = t.range;
    if (r.bottom < top || r.top > bottom || r.right < left || r.left > right) continue;
    for (const rb of rows) {
      const row = rb.index;
      if (row < r.top || row > r.bottom) continue;
      const isHeader = t.hasHeader && row === r.top;
      const dataIndex = row - r.top - (t.hasHeader ? 1 : 0);
      const rect = v.rangeRect(row, r.left, row, r.right);
      if (isHeader) {
        ctx.fillStyle = palette.tableHeaderBand;
        ctx.fillRect(rect.x, rect.y, rect.w, rect.h);
      } else if (t.hasBandedRows && dataIndex % 2 === 1) {
        // Banding is opt-in: a translucent tint of the header band on alternate data rows.
        ctx.fillStyle = palette.tableRowBand;
        ctx.fillRect(rect.x, rect.y, rect.w, rect.h);
      }
      if (t.hasHeaderColumn && !isHeader) {
        const c = v.rangeRect(row, r.left, row, r.left);
        ctx.fillStyle = palette.tableHeaderColumnBand;
        ctx.fillRect(c.x, c.y, c.w, c.h);
      }
    }
  }

  // 2. Gridlines.
  ctx.strokeStyle = palette.gridLine;
  ctx.lineWidth = 1;
  ctx.beginPath();
  for (const rb of rows)
    line(ctx, clip.x, crisp(rb.start + rb.size), clip.x + clip.w, crisp(rb.start + rb.size));
  for (const cb of cols)
    line(ctx, crisp(cb.start + cb.size), clip.y, crisp(cb.start + cb.size), clip.y + clip.h);
  ctx.stroke();

  // Table outline, in the header band's colour so frame and header read as one piece.
  ctx.strokeStyle = palette.tableHeaderBand;
  ctx.lineWidth = 1.2;
  for (const t of sheet.tables) {
    const r = t.range;
    if (r.bottom < top || r.top > bottom || r.right < left || r.left > right) continue;
    const rect = v.rangeRect(r.top, r.left, r.bottom, r.right);
    ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.w - 1, rect.h - 1);
  }

  // 3. Merged areas: one rect each, painted over the gridlines inside them.
  for (const m of merges) {
    const r = v.rangeRect(m.range.top, m.range.left, m.range.bottom, m.range.right);
    const style = styles.get(sheet.effectiveStyleId(m.range.top, m.range.left));
    ctx.fillStyle = style.fill.background !== null ? Argb.toCss(style.fill.background) : palette.canvas;
    ctx.fillRect(r.x, r.y, r.w - 1, r.h - 1);
  }

  // 4. Text. Left-aligned text spills right across empty neighbours (and right-aligned spills left).
  const lastColumn = cols[cols.length - 1]!;
  const firstColumn = cols[0]!;
  for (const rb of rows) {
    for (let ci = 0; ci < cols.length; ci++) {
      const cb = cols[ci]!;
      const cell = sheet.find(rb.index, cb.index);
      if (!cell || cell.isBlank) continue;
      if (input.editing && input.editing.row === rb.index && input.editing.column === cb.index) continue;
      const merge = merges.find((m) => m.covers(rb.index, cb.index));
      if (merge && merge.isCovered(rb.index, cb.index)) continue;
      const rect = merge
        ? v.rangeRect(merge.range.top, merge.range.left, merge.range.bottom, merge.range.right)
        : { x: cb.start, y: rb.start, w: cb.size, h: rb.size };
      drawCellText(
        input,
        cell,
        rb.index,
        cb.index,
        rect,
        merge ? null : { cols, ci, firstColumn, lastColumn },
      );
    }
  }

  // 5. Borders.
  for (const rb of rows) {
    for (const cb of cols) {
      const cell = sheet.find(rb.index, cb.index);
      const id = cell ? cell.styleId : sheet.effectiveStyleId(rb.index, cb.index);
      if (id === StylePool.defaultId) continue;
      const b = styles.get(id).border;
      const merge = merges.find((m) => m.covers(rb.index, cb.index));
      const rect = merge
        ? v.rangeRect(merge.range.top, merge.range.left, merge.range.bottom, merge.range.right)
        : { x: cb.start, y: rb.start, w: cb.size, h: rb.size };
      edge(ctx, b.top, rect.x, rect.y, rect.x + rect.w, rect.y);
      edge(ctx, b.bottom, rect.x, rect.y + rect.h, rect.x + rect.w, rect.y + rect.h);
      edge(ctx, b.left, rect.x, rect.y, rect.x, rect.y + rect.h);
      edge(ctx, b.right, rect.x + rect.w, rect.y, rect.x + rect.w, rect.y + rect.h);
    }
  }

  drawSelection(input);
  ctx.restore();
}

function drawCellText(
  input: RenderInput,
  cell: Cell,
  row: number,
  column: number,
  rect: Rect,
  spill: { cols: readonly Band[]; ci: number; firstColumn: Band; lastColumn: Band } | null,
): void {
  const { ctx, sheet, styles, palette, viewport: v } = input;
  const style = styles.get(cell.styleId);
  const shown =
    input.showFormulas && cell.formula !== null
      ? { text: '=' + cell.formula }
      : input.display(cell.value, style);
  if (!shown.text) return;
  const font = style.font;
  const px = font.sizePoints * PT_TO_PX * v.zoom;
  ctx.font = `${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}${px}px ${cssFamily(font.family)}`;
  ctx.fillStyle = cell.hyperlink
    ? palette.hyperlink
    : (shown.color ??
      (font.color === Defaults.textColorArgb && !style.fill.background
        ? palette.defaultText
        : Argb.toCss(font.color)));

  const isNumberLike = cell.value.kind === 'number';
  const h = style.alignment.horizontal;
  const align: 'left' | 'center' | 'right' =
    h === 'general'
      ? isNumberLike
        ? 'right'
        : cell.value.kind === 'boolean' || cell.value.kind === 'error'
          ? 'center'
          : 'left'
      : h;
  const pad = CELL_PAD_X * v.zoom;
  const lineHeight = px * 1.2;
  const lines = style.alignment.wrapText ? wrap(ctx, shown.text, rect.w - 2 * pad) : shown.text.split('\n');

  // Numbers that don't fit show ### (PRD §8.8), text spills or is clipped.
  let clip = rect;
  if (!style.alignment.wrapText && lines.length === 1) {
    const width = ctx.measureText(lines[0]!).width;
    if (isNumberLike && width > rect.w - 2 * pad) {
      lines[0] = '#'.repeat(
        Math.max(1, Math.floor((rect.w - 2 * pad) / Math.max(1, ctx.measureText('#').width))),
      );
    } else if (!isNumberLike && spill && width > rect.w - 2 * pad) {
      clip = spillRect(sheet, row, rect, width + 2 * pad, align, spill);
    }
  }

  const total = lines.length * lineHeight;
  const vAlign = style.alignment.vertical;
  let y =
    vAlign === 'top'
      ? rect.y + CELL_PAD_Y * v.zoom
      : vAlign === 'middle'
        ? rect.y + (rect.h - total) / 2
        : rect.y + rect.h - total - CELL_PAD_Y * v.zoom;
  y = Math.max(rect.y, y);
  const x = align === 'left' ? rect.x + pad : align === 'right' ? rect.x + rect.w - pad : rect.x + rect.w / 2;

  ctx.save();
  ctx.beginPath();
  ctx.rect(clip.x, clip.y, clip.w, clip.h);
  ctx.clip();
  ctx.textAlign = align;
  ctx.textBaseline = 'top';
  for (const text of lines) {
    ctx.fillText(text, x, y + (lineHeight - px) / 2);
    if (font.underline || font.strikethrough || cell.hyperlink) {
      const w = ctx.measureText(text).width;
      const x0 = align === 'left' ? x : align === 'right' ? x - w : x - w / 2;
      ctx.fillRect(
        x0,
        y + (lineHeight - px) / 2 + (font.strikethrough ? px * 0.55 : px * 0.95),
        w,
        Math.max(1, v.zoom),
      );
    }
    y += lineHeight;
  }
  ctx.restore();
}

function spillRect(
  sheet: Worksheet,
  row: number,
  rect: Rect,
  needed: number,
  align: 'left' | 'center' | 'right',
  spill: { cols: readonly Band[]; ci: number },
): Rect {
  let x0 = rect.x;
  let x1 = rect.x + rect.w;
  const free = (b: Band | undefined) =>
    b !== undefined && !sheet.find(row, b.index)?.value.kind.match(/^(number|text|boolean|error)$/);
  if (align !== 'right') {
    for (let j = spill.ci + 1; x1 - x0 < needed && free(spill.cols[j]); j++)
      x1 = spill.cols[j]!.start + spill.cols[j]!.size;
  }
  if (align !== 'left') {
    for (let j = spill.ci - 1; x1 - x0 < needed && free(spill.cols[j]); j--) x0 = spill.cols[j]!.start;
  }
  return { x: x0, y: rect.y, w: x1 - x0, h: rect.h };
}

function drawSelection(input: RenderInput): void {
  const { ctx, viewport: v, palette, selection, sheet } = input;
  ctx.fillStyle = palette.selectionFill;
  for (const r of selection.ranges) {
    const single = r.top === r.bottom && r.left === r.right;
    if (single) continue;
    const rect = v.rangeRect(r.top, r.left, r.bottom, Math.min(r.right, v.columns.count - 1));
    ctx.fillRect(rect.x, rect.y, rect.w, rect.h);
  }
  const last = selection.ranges[selection.ranges.length - 1];
  if (last) {
    const rr = v.rangeRect(last.top, last.left, last.bottom, last.right);
    ctx.strokeStyle = palette.activeOutline;
    ctx.lineWidth = 1;
    ctx.strokeRect(rr.x + 0.5, rr.y + 0.5, rr.w - 1, rr.h - 1);
  }
  // Active cell: a 2px outline around it (or its whole merge), with the fill handle.
  const active = sheet.selectionRangeAt(selection.row, selection.column);
  const a = v.rangeRect(active.top, active.left, active.bottom, active.right);
  ctx.strokeStyle = palette.activeOutline;
  ctx.lineWidth = 2;
  ctx.strokeRect(a.x + 1, a.y + 1, a.w - 2, a.h - 2);
  if (last) {
    const rr = v.rangeRect(last.top, last.left, last.bottom, last.right);
    ctx.fillStyle = palette.activeOutline;
    ctx.fillRect(rr.x + rr.w - 4, rr.y + rr.h - 4, 7, 7);
  }
}

function drawHeaders(input: RenderInput): void {
  const { ctx, viewport: v, palette, selection } = input;
  const inSel = (pred: (r: RenderSelection['ranges'][number]) => boolean) => selection.ranges.some(pred);
  ctx.font = `${11 * Math.min(1.5, Math.max(0.75, v.zoom))}px system-ui, -apple-system, sans-serif`;
  ctx.textBaseline = 'middle';
  ctx.textAlign = 'center';

  ctx.fillStyle = palette.headerFill;
  ctx.fillRect(0, 0, v.width, HEADER_HEIGHT);
  for (const cb of v.visibleColumns) {
    if (!cb.frozen && cb.start + cb.size < v.paneX) continue;
    const active = inSel((r) => cb.index >= r.left && cb.index <= r.right);
    if (active) {
      ctx.fillStyle = palette.activeHeaderFill;
      ctx.fillRect(cb.start, 0, cb.size, HEADER_HEIGHT);
    }
    ctx.fillStyle = palette.headerText;
    ctx.fillText(SheetDimensions.columnToLetters(cb.index), cb.start + cb.size / 2, HEADER_HEIGHT / 2);
  }
  ctx.fillStyle = palette.headerFill;
  ctx.fillRect(0, HEADER_HEIGHT, HEADER_WIDTH, v.height);
  for (const rb of v.visibleRows) {
    if (!rb.frozen && rb.start + rb.size < v.paneY) continue;
    const active = inSel((r) => rb.index >= r.top && rb.index <= r.bottom);
    if (active) {
      ctx.fillStyle = palette.activeHeaderFill;
      ctx.fillRect(0, rb.start, HEADER_WIDTH, rb.size);
    }
    ctx.fillStyle = palette.headerText;
    ctx.fillText(String(rb.index + 1), HEADER_WIDTH / 2, rb.start + rb.size / 2);
  }
  ctx.fillStyle = palette.headerFill;
  ctx.fillRect(0, 0, HEADER_WIDTH, HEADER_HEIGHT);
  ctx.strokeStyle = palette.gridLine;
  ctx.lineWidth = 1;
  ctx.beginPath();
  for (const cb of v.visibleColumns)
    line(ctx, crisp(cb.start + cb.size), 0, crisp(cb.start + cb.size), HEADER_HEIGHT);
  for (const rb of v.visibleRows)
    line(ctx, 0, crisp(rb.start + rb.size), HEADER_WIDTH, crisp(rb.start + rb.size));
  line(ctx, 0, crisp(HEADER_HEIGHT), v.width, crisp(HEADER_HEIGHT));
  line(ctx, crisp(HEADER_WIDTH), 0, crisp(HEADER_WIDTH), v.height);
  ctx.stroke();
}

function edge(
  ctx: CanvasRenderingContext2D,
  e: BorderEdge,
  x0: number,
  y0: number,
  x1: number,
  y1: number,
): void {
  if (e.style === 'none') return;
  ctx.save();
  ctx.strokeStyle = Argb.toCss(e.color);
  ctx.lineWidth = e.style === 'thick' ? 3 : e.style === 'medium' ? 2 : 1;
  ctx.setLineDash(e.style === 'dashed' ? [4, 2] : e.style === 'dotted' ? [1, 2] : []);
  ctx.beginPath();
  const horizontal = y0 === y1;
  if (e.style === 'double') {
    const o = 1.5;
    if (horizontal) {
      line(ctx, x0, crisp(y0 - o), x1, crisp(y0 - o));
      line(ctx, x0, crisp(y0 + o), x1, crisp(y0 + o));
    } else {
      line(ctx, crisp(x0 - o), y0, crisp(x0 - o), y1);
      line(ctx, crisp(x0 + o), y0, crisp(x0 + o), y1);
    }
  } else if (horizontal) line(ctx, x0, crisp(y0), x1, crisp(y1));
  else line(ctx, crisp(x0), y0, crisp(x1), y1);
  ctx.stroke();
  ctx.restore();
}

function wrap(ctx: CanvasRenderingContext2D, text: string, width: number): string[] {
  const out: string[] = [];
  for (const para of text.split('\n')) {
    let current = '';
    for (const word of para.split(/(\s+)/)) {
      const next = current + word;
      if (current && ctx.measureText(next).width > width) {
        out.push(current.trimEnd());
        current = word.trimStart();
      } else current = next;
    }
    out.push(current);
  }
  return out;
}

function line(ctx: CanvasRenderingContext2D, x0: number, y0: number, x1: number, y1: number): void {
  ctx.moveTo(x0, y0);
  ctx.lineTo(x1, y1);
}

/** Half-pixel align a 1px line so it lands on one device row. */
function crisp(v: number): number {
  return Math.round(v) - 0.5;
}

function cssFamily(family: string): string {
  return `"${family.replace(/"/g, '')}", "Helvetica Neue", Helvetica, Arial, sans-serif`;
}
