import { Defaults } from '../defaults.ts';

let generation = 0;

/**
 * A counter that moves whenever any row or column's size, visibility or style changes, in any
 * sheet. A view that lays lines out (the grid's axes) rebuilds only when it moves, rather than
 * rescanning every row model after each edit.
 */
export function lineLayoutGeneration(): number {
  return generation;
}

/** Note a layout change made other than through a Row/Column setter (e.g. replacing the maps). */
export function bumpLineLayout(): void {
  generation++;
}

/**
 * Per-row layout. Only rows that differ from defaults are stored. A row can be hidden by hand and
 * hidden by a filter at once; the two are kept apart, as in LibreOffice, so clearing a filter shows
 * only what the filter hid, and a save writes `table:visibility="filter"` or `"collapse"`.
 */
export class Row {
  #height: number = Defaults.rowHeight;
  #manuallyHidden = false;
  #filtered = false;
  #styleId = 0;

  get height(): number {
    return this.#height;
  }
  set height(value: number) {
    if (value !== this.#height) generation++;
    this.#height = value;
  }
  /** Not shown, whether hidden by hand or by a filter. */
  get hidden(): boolean {
    return this.#manuallyHidden || this.#filtered;
  }
  /** Hidden with Hide Rows (PRD §8.14; ODF `table:visibility="collapse"`). */
  get manuallyHidden(): boolean {
    return this.#manuallyHidden;
  }
  set manuallyHidden(value: boolean) {
    if (value !== this.#manuallyHidden) generation++;
    this.#manuallyHidden = value;
  }
  /** Hidden by an AutoFilter or a table's filter (PRD §8.12; ODF `table:visibility="filter"`). */
  get filtered(): boolean {
    return this.#filtered;
  }
  set filtered(value: boolean) {
    if (value !== this.#filtered) generation++;
    this.#filtered = value;
  }
  get styleId(): number {
    return this.#styleId;
  }
  set styleId(value: number) {
    if (value !== this.#styleId) generation++;
    this.#styleId = value;
  }

  get isDefault(): boolean {
    return this.#height === Defaults.rowHeight && !this.hidden && this.#styleId === 0;
  }

  clone(): Row {
    const r = new Row();
    r.#height = this.#height;
    r.#manuallyHidden = this.#manuallyHidden;
    r.#filtered = this.#filtered;
    r.#styleId = this.#styleId;
    return r;
  }
}

/** Per-column layout. Only columns that differ from defaults are stored. */
export class Column {
  #width: number = Defaults.columnWidth;
  #hidden = false;
  #styleId = 0;

  get width(): number {
    return this.#width;
  }
  set width(value: number) {
    if (value !== this.#width) generation++;
    this.#width = value;
  }
  get hidden(): boolean {
    return this.#hidden;
  }
  set hidden(value: boolean) {
    if (value !== this.#hidden) generation++;
    this.#hidden = value;
  }
  get styleId(): number {
    return this.#styleId;
  }
  set styleId(value: number) {
    if (value !== this.#styleId) generation++;
    this.#styleId = value;
  }

  get isDefault(): boolean {
    return this.#width === Defaults.columnWidth && !this.#hidden && this.#styleId === 0;
  }

  clone(): Column {
    const c = new Column();
    c.#width = this.#width;
    c.#hidden = this.#hidden;
    c.#styleId = this.#styleId;
    return c;
  }
}
