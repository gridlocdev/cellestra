import { CellAddress } from '../../model/cellAddress.ts';
import { CellRange } from '../../model/cellRange.ts';
import { MergedRegion } from '../../model/mergedRegion.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import type { XmlElement } from '../../xml/index.ts';
import {
  MAX_COLUMNS,
  MAX_ROWS,
  OdsContentTooLargeError,
  readPositive,
  REPEATED_MERGE_BUDGET,
  type RepeatBudget,
} from './odsReadLimits.ts';
import { Table } from './odsXmlNames.ts';

/** Merged cells: `table:number-columns-spanned` and `-rows-spanned` on the merge's top-left cell. */

// ── Reading ─────────────────────────────────────────────────────────────────

/**
 * The merges a `table:table-cell` starts, for each of its copies (`repeat` columns by `rowCount`
 * rows from `row`, `column`), added to `merges` in reading order.
 */
export function readMergeSpans(
  element: XmlElement,
  sheetName: string,
  row: number,
  rowCount: number,
  column: number,
  repeat: number,
  merges: MergedRegion[],
  repeats: RepeatBudget,
): void {
  const columnsSpanned = Math.min(
    readPositive(element, Table('number-columns-spanned')),
    MAX_COLUMNS - column,
  );
  const declaredRowsSpanned = readPositive(element, Table('number-rows-spanned'));
  if (columnsSpanned <= 1 && declaredRowsSpanned <= 1) return;
  // Each repeated copy starts a merge, but a copy inside the span of the one before it overlaps it
  // and would be dropped by addMerges: step over the spans instead, so only the merges that can
  // stand are made.
  const rowStep = Math.min(declaredRowsSpanned, rowCount);
  const columnStep = Math.min(columnsSpanned, repeat);
  if (rowCount > 1 || repeat > 1) {
    repeats.merges += Math.ceil(rowCount / rowStep) * Math.ceil(repeat / columnStep);
    if (repeats.merges > REPEATED_MERGE_BUDGET) {
      throw new OdsContentTooLargeError(
        `Sheet '${sheetName}' repeats merged cells past ${REPEATED_MERGE_BUDGET} in the workbook, the most repeats may create.`,
      );
    }
  }
  for (let r = row; r < row + rowCount; r += rowStep) {
    const rowsSpanned = Math.min(declaredRowsSpanned, MAX_ROWS - r);
    if (columnsSpanned <= 1 && rowsSpanned <= 1) continue;
    for (let c = column; c < column + repeat; c += columnStep) {
      const right = Math.min(c + columnsSpanned - 1, MAX_COLUMNS - 1);
      const range = new CellRange(new CellAddress(c, r), new CellAddress(right, r + rowsSpanned - 1));
      merges.push(new MergedRegion(range));
    }
  }
}

/**
 * Add the merges read from a table, skipping any that overlaps one already added: a corrupt file
 * with two overlapping spans loses the later one, not the whole load. `merges` arrive in reading
 * order (by top row), so only the merges still open at a candidate's top can overlap it.
 */
export function addMerges(sheet: Worksheet, merges: readonly MergedRegion[]): void {
  if (merges.length === 0) return;
  const accepted = [...sheet.merges];
  let open = [...accepted];
  for (const merge of merges) {
    const top = merge.range.top;
    if (open.some((m) => m.range.bottom < top)) open = open.filter((m) => m.range.bottom >= top);
    if (open.some((m) => m.range.intersects(merge.range))) continue;
    accepted.push(merge);
    open.push(merge);
  }
  sheet.replaceMerges(accepted);
}

// ── Writing ─────────────────────────────────────────────────────────────────

/**
 * The merges that cover a row, for rows visited top to bottom: merges start being reported at
 * their top row and stop after their bottom one, so each row costs the merges open there, not a
 * scan of every merge in the sheet.
 */
export class MergeSweep {
  private readonly sorted: readonly MergedRegion[];
  private next = 0;
  private open: MergedRegion[] = [];

  constructor(merges: readonly MergedRegion[]) {
    this.sorted = [...merges].sort((a, b) => a.range.top - b.range.top);
  }

  at(row: number): readonly MergedRegion[] {
    while (this.next < this.sorted.length && this.sorted[this.next]!.range.top <= row) {
      this.open.push(this.sorted[this.next++]!);
    }
    if (this.open.some((m) => m.range.bottom < row))
      this.open = this.open.filter((m) => m.range.bottom >= row);
    return this.open;
  }
}

/** Put the spans on the cell at (`row`, `column`) when it is `merge`'s top-left cell. */
export function writeMergeSpans(
  element: XmlElement,
  merge: MergedRegion | null,
  row: number,
  column: number,
): void {
  if (!merge || merge.range.top !== row || merge.range.left !== column) return;
  if (merge.range.columnCount > 1) element.setAttr(Table('number-columns-spanned'), merge.range.columnCount);
  if (merge.range.rowCount > 1) element.setAttr(Table('number-rows-spanned'), merge.range.rowCount);
}

/** A written cell that starts a merge. */
export function isSpanned(cell: XmlElement): boolean {
  return (
    cell.attr(Table('number-columns-spanned')) !== null || cell.attr(Table('number-rows-spanned')) !== null
  );
}
