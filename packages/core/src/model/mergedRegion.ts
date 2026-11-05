import type { CellAddress } from './cellAddress.ts';
import type { CellRange } from './cellRange.ts';

/** A merge: the top-left anchor holds the content and spans the rest (↔ ODF `number-*-spanned`). */
export class MergedRegion {
  readonly range: CellRange;

  constructor(range: CellRange) {
    this.range = range;
  }

  get anchor(): CellAddress {
    return this.range.topLeft;
  }

  covers(row: number, column: number): boolean {
    return this.range.contains(row, column);
  }

  /** True for every covered cell except the anchor. */
  isCovered(row: number, column: number): boolean {
    return this.covers(row, column) && !(row === this.range.top && column === this.range.left);
  }
}
