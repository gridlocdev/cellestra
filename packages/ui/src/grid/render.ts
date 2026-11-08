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

/** Table frame thickness in CSS pixels (gridlines are 1). */
const TABLE_OUTLINE_WIDTH = 2;

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
  /** Transient marks drawn over the cells (drag previews, the copy marquee, formula references). */
  readonly overlay?: GridOverlay;
}

type Range = RenderSelection['ranges'][number];

export interface GridOverlay {
  /** Where a fill-handle drag will fill, outlined dashed while dragging. */
  readonly fill?: Range | null;
  /** The copy/cut source: marching ants until paste, Escape or an edit. `phase` animates the dashes. */
  readonly marquee?: { readonly range: Range; readonly cut: boolean; readonly phase: number } | null;
  /** The ranges the formula being typed refers to, each in its own colour. */
  readonly references?: readonly { readonly range: Range; readonly color: string }[];
  /** The cell or range being pointed at while typing a formula. */
  readonly point?: Range | null;
  /** A style shown over some cells without changing them (hover previews of fonts and colours). */
  readonly stylePreview?: {
    readonly ranges: readonly Range[];
    readonly transform: (style: CellStyle) => CellStyle;
  } | null;
}

/** The style a cell draws with: its own, with any hover preview applied over it. */
function styleAt(input: RenderInput, row: number, column: number, id: number): CellStyle {
  const style = input.styles.get(id);
  const p = input.overlay?.stylePreview;
  if (!p) return style;
  for (const r of p.ranges)
    if (row >= r.top && row <= r.bottom && column >= r.left && column <= r.right) return p.transform(style);
  return style;
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
  const { ctx, sheet, palette, viewport: v } = input;
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
      const style = styleAt(input, rb.index, cb.index, sheet.effectiveStyleId(rb.index, cb.index));
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

  // 3. Merged areas: one rect each, painted over the gridlines inside them.
  for (const m of merges) {
    const r = v.rangeRect(m.range.top, m.range.left, m.range.bottom, m.range.right);
    const style = styleAt(
      input,
      m.range.top,
      m.range.left,
      sheet.effectiveStyleId(m.range.top, m.range.left),
    );
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
      if (merges.length > 0 && coveredByMerge(rb.index, cb.index)) continue; // drawn once below
      const rect = { x: cb.start, y: rb.start, w: cb.size, h: rb.size };
      drawCellText(input, cell, rb.index, cb.index, rect, { cols, ci, firstColumn, lastColumn });
    }
  }
  // A merge's text belongs to its anchor, which may be scrolled out of view while the merge shows.
  for (const m of merges) {
    const { top: row, left: column } = m.range;
    const cell = sheet.find(row, column);
    if (!cell || cell.isBlank) continue;
    if (input.editing && input.editing.row === row && input.editing.column === column) continue;
    drawCellText(input, cell, row, column, v.rangeRect(row, column, m.range.bottom, m.range.right), null);
  }

  // 5. Borders.
  for (const rb of rows) {
    for (const cb of cols) {
      // A cell with the default style still shows its row's or column's borders.
      const id = sheet.effectiveStyleId(rb.index, cb.index);
      if (id === StylePool.defaultId) continue;
      const b = styleAt(input, rb.index, cb.index, id).border;
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

  // 6. Table outlines, over cell borders so a table always reads as one framed block. Drawn from
  // the table's range, so the frame grows with the table.
  ctx.strokeStyle = palette.tableOutline;
  ctx.lineWidth = TABLE_OUTLINE_WIDTH;
  ctx.setLineDash([]);
  for (const t of sheet.tables) {
    const r = t.range;
    if (r.bottom < top || r.top > bottom || r.right < left || r.left > right) continue;
    const rect = v.rangeRect(r.top, r.left, r.bottom, r.right);
    // Centred on the cell edges, like the gridlines it covers.
    ctx.strokeRect(rect.x, rect.y, rect.w, rect.h);
  }

  drawSelection(input);
  if (input.overlay) drawOverlay(input, input.overlay);
  ctx.restore();
}

/** Marching-ants dash lengths. */
export const MARQUEE_DASH = [4, 3] as const;

function drawOverlay(input: RenderInput, o: GridOverlay): void {
  const { ctx, viewport: v, palette } = input;
  const rectOf = (r: Range) => v.rangeRect(r.top, r.left, r.bottom, Math.min(r.right, v.columns.count - 1));
  ctx.save();
  if (o.point) {
    const p = rectOf(o.point);
    ctx.fillStyle = palette.selectionFill;
    ctx.fillRect(p.x, p.y, p.w, p.h);
  }
  ctx.lineWidth = 1.5;
  for (const ref of o.references ?? []) {
    const r = rectOf(ref.range);
    ctx.strokeStyle = ref.color;
    ctx.fillStyle = ref.color + '14';
    ctx.fillRect(r.x, r.y, r.w, r.h);
    ctx.strokeRect(r.x + 0.75, r.y + 0.75, r.w - 1.5, r.h - 1.5);
  }
  if (o.point) {
    const p = rectOf(o.point);
    ctx.strokeStyle = palette.activeOutline;
    ctx.lineWidth = 2;
    ctx.setLineDash([MARQUEE_DASH[0], MARQUEE_DASH[1]]);
    ctx.strokeRect(p.x + 1, p.y + 1, p.w - 2, p.h - 2);
  }
  if (o.marquee) {
    const r = rectOf(o.marquee.range);
    ctx.lineWidth = 2;
    ctx.setLineDash([MARQUEE_DASH[0], MARQUEE_DASH[1]]);
    ctx.lineDashOffset = -o.marquee.phase;
    // Cut is muted grey, copy is green: a glance tells them apart.
    ctx.strokeStyle = o.marquee.cut ? '#707070' : '#1e884e';
    ctx.strokeRect(r.x + 1, r.y + 1, r.w - 2, r.h - 2);
    ctx.lineDashOffset = 0;
  }
  if (o.fill) {
    const r = rectOf(o.fill);
    ctx.lineWidth = 1;
    ctx.setLineDash([3, 2]);
    ctx.strokeStyle = palette.activeOutline;
    ctx.strokeRect(r.x + 0.5, r.y + 0.5, r.w - 1, r.h - 1);
  }
  ctx.restore();
}

/** The fill handle's square, centred on the bottom-right corner of the last selected range. */
export function fillHandleRect(v: Viewport, selection: RenderSelection): Rect | null {
  const last = selection.ranges[selection.ranges.length - 1];
  if (!last) return null;
  // Whole columns or rows have no corner to drag on screen.
  if (last.bottom >= v.rows.count - 1 || last.right >= v.columns.count - 1) return null;
  const rr = v.rangeRect(last.top, last.left, last.bottom, last.right);
  return { x: rr.x + rr.w - 4, y: rr.y + rr.h - 4, w: 7, h: 7 };
}

/**
 * Where a checkbox cell's box sits: a square about the size of the text, placed by the cell's
 * alignment (centred unless aligned otherwise). The grid hit-tests clicks against the same square.
 */
export function checkboxRect(cell: Rect, style: CellStyle, zoom: number): Rect {
  const size = Math.max(6, Math.min(Math.round(CHECKBOX_SIZE * zoom), cell.w - 4, cell.h - 4));
  const pad = CELL_PAD_X * zoom;
  const h = style.alignment.horizontal;
  const x =
    h === 'left' ? cell.x + pad : h === 'right' ? cell.x + cell.w - pad - size : cell.x + (cell.w - size) / 2;
  const vAlign = style.alignment.vertical;
  const y =
    vAlign === 'top'
      ? cell.y + CELL_PAD_Y * zoom
      : vAlign === 'bottom' && cell.h > size + 8
        ? cell.y + cell.h - size - 3 * zoom
        : cell.y + (cell.h - size) / 2;
  return { x: Math.round(x) + 0.5, y: Math.round(y) + 0.5, w: size, h: size };
}

const CHECKBOX_SIZE = 13;

function drawCheckbox(
  ctx: CanvasRenderingContext2D,
  r: Rect,
  checked: boolean,
  palette: GridPalette,
  zoom: number,
): void {
  const radius = Math.min(3 * zoom, r.w / 4);
  ctx.save();
  ctx.beginPath();
  ctx.roundRect(r.x, r.y, r.w, r.h, radius);
  if (checked) {
    ctx.fillStyle = palette.activeOutline;
    ctx.fill();
    ctx.strokeStyle = '#ffffff';
    ctx.lineWidth = Math.max(1.5, 1.8 * zoom);
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.beginPath();
    ctx.moveTo(r.x + r.w * 0.24, r.y + r.h * 0.53);
    ctx.lineTo(r.x + r.w * 0.43, r.y + r.h * 0.72);
    ctx.lineTo(r.x + r.w * 0.77, r.y + r.h * 0.3);
    ctx.stroke();
  } else {
    ctx.fillStyle = palette.canvas;
    ctx.fill();
    ctx.strokeStyle = palette.checkboxBorder;
    ctx.lineWidth = Math.max(1, zoom);
    ctx.stroke();
  }
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
  const { ctx, sheet, palette, viewport: v } = input;
  // A cell with the default style takes its row's or column's (font, alignment, format).
  const id = cell.styleId !== StylePool.defaultId ? cell.styleId : sheet.effectiveStyleId(row, column);
  const style = styleAt(input, row, column, id);
  if (style.checkbox && cell.value.kind === 'boolean' && !(input.showFormulas && cell.formula !== null)) {
    drawCheckbox(ctx, checkboxRect(rect, style, v.zoom), cell.value.value, palette, v.zoom);
    return;
  }
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
    // A lone cell shows only its outline; in a multi-range selection every piece is shaded.
    if (single && selection.ranges.length === 1) continue;
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
  // Active cell: a 2px outline around it (or its whole merge), with the fill handle. A selection
  // of more than that is outlined as a whole, so the cell it started from isn't singled out.
  const active = sheet.selectionRangeAt(selection.row, selection.column);
  const onlyActive =
    selection.ranges.length <= 1 &&
    (!last ||
      (last.top === active.top &&
        last.left === active.left &&
        last.bottom === active.bottom &&
        last.right === active.right));
  if (onlyActive) {
    const a = v.rangeRect(active.top, active.left, active.bottom, active.right);
    ctx.strokeStyle = palette.activeOutline;
    ctx.lineWidth = 2;
    ctx.strokeRect(a.x + 1, a.y + 1, a.w - 2, a.h - 2);
  }
  const handle = input.overlay?.fill ? null : fillHandleRect(v, selection);
  if (handle) {
    ctx.fillStyle = palette.activeOutline;
    ctx.fillRect(handle.x, handle.y, handle.w, handle.h);
    ctx.strokeStyle = palette.canvas;
    ctx.lineWidth = 1;
    ctx.strokeRect(handle.x - 0.5, handle.y - 0.5, handle.w + 1, handle.h + 1);
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

/** Recent `wrap` results, keyed by font, width and text: a repaint or a row-height pass re-wraps the same cells. */
const wrapCache = new Map<string, string[]>();
const WRAP_CACHE_MAX = 50_000;

/**
 * Lines of `text` folded to `width`: at its own line breaks, then between words, and inside a word
 * too long for a line on its own (so nothing is cut off at the cell's edge).
 */
export function wrap(ctx: CanvasRenderingContext2D, text: string, width: number): string[] {
  const key = `${String(ctx.font)}\u0000${width}\u0000${text}`;
  const hit = wrapCache.get(key);
  if (hit) return hit;
  const out = wrapUncached(ctx, text, width);
  if (wrapCache.size >= WRAP_CACHE_MAX) wrapCache.clear();
  wrapCache.set(key, out);
  return out;
}

function wrapUncached(ctx: CanvasRenderingContext2D, text: string, width: number): string[] {
  // No room at all (a sliver of a column): one character a line would make the row absurdly tall.
  if (!(width > 0)) return text.split('\n');
  const out: string[] = [];
  const fits = (s: string) => ctx.measureText(s).width <= width;
  for (const para of text.split('\n')) {
    let current = '';
    for (const word of para.split(/(\s+)/)) {
      if (!word) continue;
      const next = current + word;
      if (fits(next)) {
        current = next;
        continue;
      }
      if (current.trim()) out.push(current.trimEnd());
      current = word.trimStart();
      // A word wider than the line: break it at the longest prefix that fits (at least one
      // character, never half a surrogate pair), found by binary search.
      while (current.length > 1 && !fits(current)) {
        let lo = 1;
        let hi = current.length - 1;
        while (lo < hi) {
          const mid = (lo + hi + 1) >> 1;
          if (fits(current.slice(0, mid))) lo = mid;
          else hi = mid - 1;
        }
        let cut = lo;
        const code = current.charCodeAt(cut - 1);
        if (code >= 0xd800 && code <= 0xdbff) cut = cut > 1 ? cut - 1 : cut + 1;
        out.push(current.slice(0, cut));
        current = current.slice(cut);
      }
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
