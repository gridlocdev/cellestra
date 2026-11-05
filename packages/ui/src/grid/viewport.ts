import type { Axis } from './axis.ts';

export const HEADER_WIDTH = 46;
export const HEADER_HEIGHT = 22;

/** A line drawn on screen: its index and its screen-space start and size (zoomed CSS px). */
export interface Band {
  readonly index: number;
  readonly start: number;
  readonly size: number;
  /** In the frozen pane (does not scroll). */
  readonly frozen: boolean;
}

export interface Rect {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

export type Hit =
  | { readonly zone: 'cell'; readonly row: number; readonly column: number }
  | { readonly zone: 'columnHeader'; readonly column: number }
  | { readonly zone: 'rowHeader'; readonly row: number }
  | { readonly zone: 'corner' };

/**
 * Where everything is on screen for one frame: the two axes, the frozen split, the scroll offset
 * of the scrollable pane (unzoomed px past the frozen lines), the zoom, and the viewport size.
 * Pure geometry — the renderer, hit-testing, and the inline editor all ask the same object.
 */
export class Viewport {
  readonly columns: Axis;
  readonly rows: Axis;
  readonly frozenColumns: number;
  readonly frozenRows: number;
  readonly scrollX: number;
  readonly scrollY: number;
  readonly zoom: number;
  readonly width: number;
  readonly height: number;
  readonly visibleColumns: readonly Band[];
  readonly visibleRows: readonly Band[];

  constructor(init: {
    columns: Axis;
    rows: Axis;
    frozenColumns: number;
    frozenRows: number;
    scrollX: number;
    scrollY: number;
    zoom: number;
    width: number;
    height: number;
  }) {
    this.columns = init.columns;
    this.rows = init.rows;
    this.frozenColumns = Math.min(init.frozenColumns, init.columns.count);
    this.frozenRows = Math.min(init.frozenRows, init.rows.count);
    this.scrollX = Math.max(0, init.scrollX);
    this.scrollY = Math.max(0, init.scrollY);
    this.zoom = init.zoom;
    this.width = init.width;
    this.height = init.height;
    this.visibleColumns = bands(
      this.columns,
      this.frozenColumns,
      this.scrollX,
      this.zoom,
      HEADER_WIDTH,
      this.width,
    );
    this.visibleRows = bands(this.rows, this.frozenRows, this.scrollY, this.zoom, HEADER_HEIGHT, this.height);
  }

  /** Screen x where the scrollable columns begin (right edge of the frozen pane). */
  get paneX(): number {
    return HEADER_WIDTH + this.columns.offset(this.frozenColumns) * this.zoom;
  }

  get paneY(): number {
    return HEADER_HEIGHT + this.rows.offset(this.frozenRows) * this.zoom;
  }

  /** Scrollable content size in unzoomed px. */
  get contentWidth(): number {
    return this.columns.total - this.columns.offset(this.frozenColumns);
  }

  get contentHeight(): number {
    return this.rows.total - this.rows.offset(this.frozenRows);
  }

  columnX(column: number): number {
    return lineStart(this.columns, this.frozenColumns, this.scrollX, this.zoom, HEADER_WIDTH, column);
  }

  rowY(row: number): number {
    return lineStart(this.rows, this.frozenRows, this.scrollY, this.zoom, HEADER_HEIGHT, row);
  }

  /** Screen rect of a cell range (inclusive), unclipped. */
  rangeRect(top: number, left: number, bottom: number, right: number): Rect {
    const x = this.columnX(left);
    const y = this.rowY(top);
    return {
      x,
      y,
      w: this.columnX(right) + this.columns.size(right) * this.zoom - x,
      h: this.rowY(bottom) + this.rows.size(bottom) * this.zoom - y,
    };
  }

  hitTest(x: number, y: number): Hit {
    const inHeaderCol = x < HEADER_WIDTH;
    const inHeaderRow = y < HEADER_HEIGHT;
    if (inHeaderCol && inHeaderRow) return { zone: 'corner' };
    const column = this.columnAt(x);
    const row = this.rowAt(y);
    if (inHeaderRow) return { zone: 'columnHeader', column };
    if (inHeaderCol) return { zone: 'rowHeader', row };
    return { zone: 'cell', row, column };
  }

  columnAt(x: number): number {
    return lineAt(this.columns, this.frozenColumns, this.scrollX, this.zoom, HEADER_WIDTH, x);
  }

  rowAt(y: number): number {
    return lineAt(this.rows, this.frozenRows, this.scrollY, this.zoom, HEADER_HEIGHT, y);
  }

  /** The scroll offset that brings a line fully into the scrollable pane (unchanged if it is frozen or visible). */
  scrollToReveal(axis: 'x' | 'y', index: number): number {
    const a = axis === 'x' ? this.columns : this.rows;
    const frozen = axis === 'x' ? this.frozenColumns : this.frozenRows;
    const scroll = axis === 'x' ? this.scrollX : this.scrollY;
    if (index < frozen) return scroll;
    const pane = (axis === 'x' ? this.width - this.paneX : this.height - this.paneY) / this.zoom;
    const start = a.offset(index) - a.offset(frozen);
    const end = start + a.size(index);
    if (start < scroll) return start;
    if (end > scroll + pane) return Math.max(0, Math.min(start, end - pane));
    return scroll;
  }
}

function lineStart(a: Axis, frozen: number, scroll: number, zoom: number, header: number, i: number): number {
  if (i < frozen) return header + a.offset(i) * zoom;
  return header + (a.offset(i) - scroll) * zoom;
}

function lineAt(a: Axis, frozen: number, scroll: number, zoom: number, header: number, px: number): number {
  const unzoomed = Math.max(0, (px - header) / zoom);
  const frozenEnd = a.offset(frozen);
  if (unzoomed < frozenEnd) return a.indexAt(unzoomed);
  return Math.max(frozen < a.count ? a.firstVisibleFrom(frozen) : 0, a.indexAt(unzoomed + scroll));
}

function bands(
  a: Axis,
  frozen: number,
  scroll: number,
  zoom: number,
  header: number,
  extent: number,
): Band[] {
  const out: Band[] = [];
  let pos = header;
  for (let i = 0; i < frozen && pos < extent; i++) {
    const s = a.size(i) * zoom;
    if (s > 0) out.push({ index: i, start: pos, size: s, frozen: true });
    pos += s;
  }
  if (frozen >= a.count) return out;
  let i = a.indexAt(a.offset(frozen) + scroll);
  let start = header + (a.offset(i) - scroll) * zoom;
  while (i < a.count && start < extent) {
    const s = a.size(i) * zoom;
    if (s > 0) {
      out.push({ index: i, start, size: s, frozen: false });
      start += s;
      i++;
    } else {
      // Jump a hidden run in one step (a hidden tail of a million rows is common).
      i = a.nextVisible(i);
    }
  }
  return out;
}
