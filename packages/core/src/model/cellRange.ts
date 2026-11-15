import { CellAddress } from './cellAddress.ts';

/**
 * A rectangular block of cells defined by two corners. Corners keep their own `$` flags and sheet
 * for faithful formula round-tripping; top/left/bottom/right are always min/max-ordered.
 */
export class CellRange {
  readonly start: CellAddress;
  readonly end: CellAddress;
  readonly top: number;
  readonly bottom: number;
  readonly left: number;
  readonly right: number;

  constructor(start: CellAddress, end: CellAddress = start) {
    this.start = start;
    this.end = end;
    this.top = Math.min(start.row, end.row);
    this.bottom = Math.max(start.row, end.row);
    this.left = Math.min(start.column, end.column);
    this.right = Math.max(start.column, end.column);
  }

  /** Build from zero-based bounds (relative, unqualified). */
  static fromBounds(
    top: number,
    left: number,
    bottom: number,
    right: number,
    sheet: string | null = null,
  ): CellRange {
    return new CellRange(
      new CellAddress(left, top, false, false, sheet),
      new CellAddress(right, bottom, false, false, sheet),
    );
  }

  get rowCount(): number {
    return this.bottom - this.top + 1;
  }
  get columnCount(): number {
    return this.right - this.left + 1;
  }
  get cellCount(): number {
    return this.rowCount * this.columnCount;
  }
  get isSingleCell(): boolean {
    return this.rowCount === 1 && this.columnCount === 1;
  }
  /** The start corner's sheet (ranges live on one sheet). */
  get sheet(): string | null {
    return this.start.sheet;
  }

  contains(row: number, column: number): boolean {
    return row >= this.top && row <= this.bottom && column >= this.left && column <= this.right;
  }

  containsAddress(a: CellAddress): boolean {
    return this.contains(a.row, a.column);
  }

  containsRange(other: CellRange): boolean {
    return (
      other.top >= this.top &&
      other.bottom <= this.bottom &&
      other.left >= this.left &&
      other.right <= this.right
    );
  }

  intersects(other: CellRange): boolean {
    return (
      this.left <= other.right &&
      this.right >= other.left &&
      this.top <= other.bottom &&
      this.bottom >= other.top
    );
  }

  /** The parts of this range outside `other`, as up to four rectangles (above, below, left, right). */
  subtract(other: CellRange): CellRange[] {
    if (!this.intersects(other)) return [this];
    const out: CellRange[] = [];
    const top = Math.max(this.top, other.top);
    const bottom = Math.min(this.bottom, other.bottom);
    if (this.top < other.top) out.push(CellRange.fromBounds(this.top, this.left, other.top - 1, this.right));
    if (this.bottom > other.bottom)
      out.push(CellRange.fromBounds(other.bottom + 1, this.left, this.bottom, this.right));
    if (this.left < other.left) out.push(CellRange.fromBounds(top, this.left, bottom, other.left - 1));
    if (this.right > other.right) out.push(CellRange.fromBounds(top, other.right + 1, bottom, this.right));
    return out;
  }

  /** Same bounds and sheet, ignoring `$` flags. */
  sameBounds(other: CellRange): boolean {
    return (
      this.top === other.top &&
      this.left === other.left &&
      this.bottom === other.bottom &&
      this.right === other.right &&
      this.sheet === other.sheet
    );
  }

  equals(other: CellRange | null | undefined): boolean {
    return !!other && this.start.equals(other.start) && this.end.equals(other.end);
  }

  get topLeft(): CellAddress {
    return new CellAddress(this.left, this.top, false, false, this.sheet);
  }
  get bottomRight(): CellAddress {
    return new CellAddress(this.right, this.bottom, false, false, this.sheet);
  }

  /** Every (row, column) in the rectangle, row-major. */
  *positions(): IterableIterator<[row: number, column: number]> {
    for (let r = this.top; r <= this.bottom; r++) for (let c = this.left; c <= this.right; c++) yield [r, c];
  }

  withSheet(sheet: string | null): CellRange {
    return new CellRange(this.start.withSheet(sheet), this.end.withSheet(sheet));
  }

  toString(): string {
    if (this.isSingleCell && this.start.equals(this.end)) return this.start.toString();
    // The sheet qualifier belongs on the whole range, not repeated on the end corner.
    return `${this.start.toString()}:${this.end.withoutSheet().toString()}`;
  }

  static parse(text: string): CellRange {
    const r = CellRange.tryParse(text);
    if (!r) throw new SyntaxError(`'${text}' is not a valid range reference.`);
    return r;
  }

  static tryParse(text: string | null | undefined): CellRange | null {
    if (!text || !text.trim()) return null;
    const colon = text.indexOf(':');
    if (colon < 0) {
      const single = CellAddress.tryParse(text);
      return single ? new CellRange(single) : null;
    }
    const start = CellAddress.tryParse(text.slice(0, colon));
    // ODF writes an end corner on the start's sheet with an empty sheet part: `Sheet1.$B$2:.$B$7`.
    let end = CellAddress.tryParse(
      text
        .slice(colon + 1)
        .trim()
        .replace(/^\$?\./, ''),
    );
    if (!start || !end) return null;
    // A range lives on one sheet: an end-corner qualifier is only valid when it matches the start's.
    if (end.sheet !== null && end.sheet.toLowerCase() !== start.sheet?.toLowerCase()) return null;
    if (start.sheet !== null && end.sheet === null) end = end.withSheet(start.sheet);
    return new CellRange(start, end);
  }
}
