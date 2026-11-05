import { Defaults } from '../defaults.ts';

/** Per-row layout. Only rows that differ from defaults are stored. */
export class Row {
  height: number = Defaults.rowHeight;
  /** Manual or filter-driven (PRD §8.14 / ODF `table:visibility`). */
  hidden = false;
  styleId = 0;

  get isDefault(): boolean {
    return this.height === Defaults.rowHeight && !this.hidden && this.styleId === 0;
  }

  clone(): Row {
    return Object.assign(new Row(), { height: this.height, hidden: this.hidden, styleId: this.styleId });
  }
}

/** Per-column layout. Only columns that differ from defaults are stored. */
export class Column {
  width: number = Defaults.columnWidth;
  hidden = false;
  styleId = 0;

  get isDefault(): boolean {
    return this.width === Defaults.columnWidth && !this.hidden && this.styleId === 0;
  }

  clone(): Column {
    return Object.assign(new Column(), { width: this.width, hidden: this.hidden, styleId: this.styleId });
  }
}
