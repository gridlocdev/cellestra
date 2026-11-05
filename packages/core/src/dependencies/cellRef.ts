/**
 * A fully-qualified cell identity (sheet + zero-based position) used as a node in the dependency
 * graph. Unlike `CellAddress` it carries no `$` flags. Sheet names compare case-insensitively.
 */
export interface CellRef {
  readonly sheet: string;
  readonly row: number;
  readonly column: number;
}

export const CellRef = {
  of: (sheet: string, row: number, column: number): CellRef => Object.freeze({ sheet, row, column }),

  /** A string key equal for refs that are equal (sheet case-insensitive). */
  key: (ref: CellRef): string => `${ref.sheet.toLowerCase()}\u0000${ref.row}\u0000${ref.column}`,

  equals: (a: CellRef, b: CellRef): boolean =>
    a.row === b.row && a.column === b.column && a.sheet.toLowerCase() === b.sheet.toLowerCase(),

  toString: (ref: CellRef): string => `${ref.sheet}!R${ref.row}C${ref.column}`,
} as const;

/**
 * A resolved precedent rectangle too large to expand into the graph cell by cell; the engine keeps
 * it so an edit inside still marks its dependent formula.
 */
export interface RangeRef {
  readonly sheet: string;
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

export function rangeRefContains(range: RangeRef, cell: CellRef): boolean {
  return (
    cell.sheet.toLowerCase() === range.sheet.toLowerCase() &&
    cell.row >= range.top &&
    cell.row <= range.bottom &&
    cell.column >= range.left &&
    cell.column <= range.right
  );
}
