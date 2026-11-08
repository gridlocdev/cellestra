import { ValueComparer } from '../formulas/evaluation/valueComparer.ts';
import { ReferenceAdjuster } from '../formulas/references/referenceAdjuster.ts';
import { EditRefusedError } from '../editing/editCommand.ts';
import { rewriteFormula } from '../editing/formulaRewrite.ts';
import type { Cell } from '../model/cell.ts';
import { CellRange } from '../model/cellRange.ts';
import { getHostLocale } from '../model/documentSettings.ts';
import type { Worksheet } from '../model/worksheet.ts';
import { CellValue } from '../values/cellValue.ts';
import { FilterEngine } from './filterEngine.ts';
import type { SortLevel, SortSpec } from './types.ts';

/** A sort spec; `hasHeader` defaults to true. */
export function sortSpec(levels: readonly SortLevel[], hasHeader = true): SortSpec {
  return { levels, hasHeader };
}

interface RowSnapshot {
  readonly originalRow: number;
  readonly cells: readonly (Cell | null)[];
}

/**
 * Sort (PRD §8.11): a stable multi-level sort of a range's data rows (the header stays put). Empty
 * cells always sort last; values order as the comparison operators do. Relative references into the
 * sorted rows follow their row, and an active filter is re-applied so visibility tracks the values.
 * A merge across the data rows refuses the sort (`EditRefusedError`) before anything moves.
 */
export const SortEngine = {
  sort(sheet: Worksheet, range: CellRange, spec: SortSpec, locale?: string): void {
    const first = range.top + (spec.hasHeader ? 1 : 0);
    if (first > range.bottom) return;
    const data = CellRange.fromBounds(first, range.left, range.bottom, range.right);
    // Merges don't move with their rows, so sorting under one would scramble it: refuse instead.
    if (sheet.merges.some((m) => m.range.intersects(data)))
      throw new EditRefusedError("Merged cells can't be sorted. Unmerge them, then sort again.");
    const rows: RowSnapshot[] = [];
    for (let row = first; row <= range.bottom; row++) {
      const cells: (Cell | null)[] = [];
      for (let c = 0; c < range.columnCount; c++)
        cells.push(sheet.find(row, range.left + c)?.clone() ?? null);
      rows.push({ originalRow: row, cells });
    }
    // Array.prototype.sort is stable, so equal keys keep their original order.
    const ordered = [...rows].sort((a, b) => compareRows(a, b, spec, range.left));
    ordered.forEach((snapshot, i) => {
      const target = first + i;
      const rowDelta = target - snapshot.originalRow;
      snapshot.cells.forEach((cell, c) => {
        if (cell && cell.formula !== null && rowDelta !== 0) {
          cell.formula = rewriteFormula(cell.formula, (expr) =>
            ReferenceAdjuster.forSortMove(expr, rowDelta, first, range.bottom),
          );
        }
        sheet.putCell(target, range.left + c, cell);
      });
    });
    reapplyActiveFilter(sheet, range, locale ?? getHostLocale());
  },
} as const;

function reapplyActiveFilter(sheet: Worksheet, sorted: CellRange, locale: string): void {
  if (sheet.activeFilterSpec && sheet.activeFilter && sheet.activeFilter.intersects(sorted)) {
    FilterEngine.apply(sheet, sheet.activeFilter, sheet.activeFilterSpec, locale);
  }
  for (const t of sheet.tables) {
    if (t.filter && t.range.intersects(sorted)) FilterEngine.apply(sheet, t.range, t.filter, locale);
  }
}

function compareRows(a: RowSnapshot, b: RowSnapshot, spec: SortSpec, rangeLeft: number): number {
  for (const level of spec.levels) {
    const index = level.column - rangeLeft;
    if (index < 0 || index >= a.cells.length) continue;
    const va = a.cells[index]?.value ?? CellValue.empty;
    const vb = b.cells[index]?.value ?? CellValue.empty;
    // Empty cells always sort last, regardless of direction.
    const ea = va.kind === 'empty';
    const eb = vb.kind === 'empty';
    if (ea !== eb) return ea ? 1 : -1;
    const result = ValueComparer.compare(va, vb);
    const cmp = typeof result === 'number' ? result : 0;
    if (cmp !== 0) return level.ascending ? cmp : -cmp;
  }
  return 0;
}
