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

/** Per-row layout. Only rows that differ from defaults are stored. */
export class Row {
  #height: number = Defaults.rowHeight;
  #hidden = false;
  #styleId = 0;

  get height(): number {
    return this.#height;
  }
  set height(value: number) {
    if (value !== this.#height) generation++;
    this.#height = value;
  }
  /** Manual or filter-driven (PRD §8.14 / ODF `table:visibility`). */
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
    return this.#height === Defaults.rowHeight && !this.#hidden && this.#styleId === 0;
  }

  clone(): Row {
    const r = new Row();
    r.#height = this.#height;
    r.#hidden = this.#hidden;
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
