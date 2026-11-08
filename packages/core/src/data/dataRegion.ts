import { CellRange } from '../model/cellRange.ts';
import type { Worksheet } from '../model/worksheet.ts';

/** Excel's "current region" and header heuristics. */
export const DataRegion = {
  /**
   * The block of populated cells around (row, column): grows while any adjacent row/column inside
   * the current span holds a value or formula (a blank cell that only carries a style doesn't count).
   * An empty start cell is its own region.
   */
  currentRegion(sheet: Worksheet, row: number, column: number): CellRange {
    if (!hasContent(sheet, row, column)) return CellRange.fromBounds(row, column, row, column);
    let top = row;
    let bottom = row;
    let left = column;
    let right = column;
    // Grow each side as far as it goes before re-checking the others: growing one row at a time
    // and rescanning the (often empty) neighbouring columns after each step is quadratic.
    let grew = true;
    while (grew) {
      grew = false;
      while (rowHasData(sheet, top - 1, left, right)) {
        top--;
        grew = true;
      }
      while (rowHasData(sheet, bottom + 1, left, right)) {
        bottom++;
        grew = true;
      }
      while (columnHasData(sheet, left - 1, top, bottom)) {
        left--;
        grew = true;
      }
      while (columnHasData(sheet, right + 1, top, bottom)) {
        right++;
        grew = true;
      }
    }
    return CellRange.fromBounds(top, left, bottom, right);
  },

  /** A first row of text over a body holding something that isn't text reads as a header. */
  looksLikeHeader(sheet: Worksheet, range: CellRange): boolean {
    if (range.rowCount < 2) return false;
    let headerAllText = true;
    let bodyHasNonText = false;
    for (let c = range.left; c <= range.right; c++) {
      const header = sheet.getValue(range.top, c);
      if (header.kind !== 'empty' && header.kind !== 'text') headerAllText = false;
      for (let r = range.top + 1; r <= range.bottom; r++) {
        const v = sheet.getValue(r, c);
        if (v.kind !== 'empty' && v.kind !== 'text') bodyHasNonText = true;
      }
    }
    return headerAllText && bodyHasNonText;
  },
} as const;

function rowHasData(sheet: Worksheet, row: number, left: number, right: number): boolean {
  if (row < 0) return false;
  for (let c = left; c <= right; c++) if (hasContent(sheet, row, c)) return true;
  return false;
}

function columnHasData(sheet: Worksheet, column: number, top: number, bottom: number): boolean {
  if (column < 0) return false;
  for (let r = top; r <= bottom; r++) if (hasContent(sheet, r, column)) return true;
  return false;
}

/** A value or a formula: formatting alone doesn't make a cell part of the data. */
function hasContent(sheet: Worksheet, row: number, column: number): boolean {
  const cell = sheet.find(row, column);
  return cell !== undefined && (cell.formula !== null || cell.value.kind !== 'empty');
}
