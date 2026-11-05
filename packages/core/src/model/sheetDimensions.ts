/**
 * Sheet size limits and A1 ↔ index conversions for column letters. Bounds match LibreOffice
 * Calc so files round-trip without clamping. All model indices are zero-based (A = 0, row 1 = 0).
 */
export const SheetDimensions = {
  maxColumns: 16384,
  maxRows: 1048576,

  isValidColumn: (column: number): boolean =>
    Number.isInteger(column) && column >= 0 && column < SheetDimensions.maxColumns,
  isValidRow: (row: number): boolean => Number.isInteger(row) && row >= 0 && row < SheetDimensions.maxRows,

  /** 0 → "A", 26 → "AA". */
  columnToLetters(column: number): string {
    if (!SheetDimensions.isValidColumn(column)) throw new RangeError(`Column index out of range: ${column}`);
    // Bijective base-26 (no zero digit), so offset by 1 each step.
    let out = '';
    let n = column;
    do {
      out = String.fromCharCode(65 + (n % 26)) + out;
      n = Math.floor(n / 26) - 1;
    } while (n >= 0);
    return out;
  },

  /** "A" → 0, case-insensitive; null when not a valid in-range label. */
  lettersToColumn(letters: string): number | null {
    if (letters.length === 0) return null;
    let result = 0;
    for (const raw of letters) {
      const c = raw.toUpperCase().charCodeAt(0);
      if (c < 65 || c > 90) return null;
      result = result * 26 + (c - 64);
      if (result > SheetDimensions.maxColumns) return null;
    }
    return result - 1;
  },
} as const;

/** Pack a (row, column) position into one safe integer, for Map keys. */
export function cellKey(row: number, column: number): number {
  return row * SheetDimensions.maxColumns + column;
}

export function keyRow(key: number): number {
  return Math.floor(key / SheetDimensions.maxColumns);
}

export function keyColumn(key: number): number {
  return key % SheetDimensions.maxColumns;
}
