import type { CellRef } from '../dependencies/cellRef.ts';
import { ReferenceCollector } from '../dependencies/referenceCollector.ts';
import type { Expr } from '../formulas/ast/expr.ts';
import { Parser } from '../formulas/parsing/parser.ts';
import { CellRange } from '../model/cellRange.ts';
import { ReferenceAdjuster } from '../formulas/references/referenceAdjuster.ts';
import type { NamedRange } from '../model/namedRange.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';
import { EditResult, requireSheet, type ChangedCell, type EditCommand } from './editCommand.ts';
import { rewriteFormula } from './formulaRewrite.ts';
import {
  captureReferences,
  defaultChartHome,
  restoreReferences,
  transformReferences,
  type ReferenceSnapshot,
} from './workbookReferences.ts';

/**
 * Delete a sheet; references to it elsewhere become `#REF!`: in formulas, and in everything else that
 * names cells (see `transformReferences`) — a named range on it (or scoped to it) goes, a validation
 * list sourced from it accepts nothing, a chart series plotting it goes. Undo re-inserts the same
 * sheet object at its old index and restores every other sheet's formula text and those references.
 * Throws on the last sheet.
 */
export class DeleteSheetCommand implements EditCommand {
  readonly label = 'Delete Sheet';
  private readonly sheetName: string;
  private index = 0;
  private removed: Worksheet | null = null;
  private otherFormulasBefore: [ChangedCell, string][] | null = null;
  private referencesBefore: ReferenceSnapshot | null = null;

  constructor(sheetName: string) {
    this.sheetName = sheetName;
  }

  apply(workbook: Workbook): EditResult {
    const sheet = requireSheet(workbook, this.sheetName);
    this.index = workbook.indexOf(sheet);
    this.removed = sheet;
    this.otherFormulasBefore = captureFormulas(workbook, sheet);
    this.referencesBefore = captureReferences(workbook);
    // Charts read unqualified references against the first data sheet, which may be this one.
    const chartHome = defaultChartHome(workbook);
    workbook.removeSheet(sheet);
    const deleted = this.sheetName;
    const transform = (expr: Expr, home: string): Expr =>
      ReferenceAdjuster.forDeleteSheet(expr, deleted, home);
    for (const other of workbook.worksheets) {
      for (const [, , cell] of other.enumerateCells()) {
        if (cell.formula === null) continue;
        cell.formula = rewriteFormula(cell.formula, (expr) => transform(expr, other.name));
      }
    }
    transformReferences(workbook, transform, chartHome);
    // A name scoped to the deleted sheet has no scope left, wherever its target is.
    const scopedHere = (n: NamedRange) =>
      n.scope !== null && n.scope !== undefined && n.scope.toLowerCase() === deleted.toLowerCase();
    if (workbook.namedRanges.some(scopedHere))
      workbook.replaceNamedRanges(workbook.namedRanges.filter((n) => !scopedHere(n)));
    return EditResult.structuralChange;
  }

  undo(workbook: Workbook): EditResult {
    if (!this.removed || !this.otherFormulasBefore || !this.referencesBefore) return EditResult.none;
    workbook.insertSheet(this.index, this.removed);
    restoreFormulas(workbook, this.otherFormulasBefore);
    restoreReferences(workbook, this.referencesBefore);
    return EditResult.structuralChange;
  }
}

/** What deleting a sheet would break elsewhere in the workbook. */
export interface SheetDeletionImpact {
  /** The sheet holds no values, formulas, comments, shapes or chart. */
  readonly empty: boolean;
  /** Formulas on other sheets that read it, directly or through a named range. */
  readonly formulas: number;
  /** Charts on other sheets that plot it, and validation lists on other sheets sourced from it. */
  readonly otherReferences: number;
}

/** What deleting `sheetName` would lose and break, without changing anything. */
export function sheetDeletionImpact(workbook: Workbook, sheetName: string): SheetDeletionImpact {
  const target = requireSheet(workbook, sheetName);
  const isDeleted = (sheet: string | null | undefined): boolean =>
    sheet !== null && sheet !== undefined && sheet.toLowerCase() === target.name.toLowerCase();

  let empty = !target.isChartSheet && target.chart === null;
  if (empty)
    for (const [, , cell] of target.enumerateCells()) {
      if (cell.formula !== null || cell.value.kind !== 'empty' || cell.unmodeled !== null) {
        empty = false;
        break;
      }
    }

  let formulas = 0;
  let otherReferences = 0;
  const chartHome = defaultChartHome(workbook);
  const rangeOnDeleted = (reference: string | null, home: string | undefined): boolean => {
    if (reference === null) return false;
    const range = CellRange.tryParse(reference.trim());
    return range !== null && isDeleted(range.sheet ?? home);
  };
  for (const sheet of workbook.worksheets) {
    if (sheet === target) continue;
    const resolve = (name: string): CellRange | null =>
      workbook.findNamedRange(name, sheet.name)?.target ?? null;
    for (const [, , cell] of sheet.enumerateCells()) {
      if (cell.formula === null) continue;
      const parsed = Parser.tryParse(cell.formula);
      if (!parsed.ok) continue;
      const read = new Map<string, CellRef>();
      ReferenceCollector.collect(parsed.expr, sheet.name, resolve, read);
      if ([...read.values()].some((ref) => isDeleted(ref.sheet))) formulas++;
    }
    const chart = sheet.chart;
    if (chart) {
      const home = chartHome ?? sheet.name;
      const refs = [
        chart.categoriesReference,
        ...chart.series.flatMap((x) => [
          x.valuesReference,
          x.nameReference,
          x.xValuesReference,
          x.bubbleSizeReference,
        ]),
      ];
      if (refs.some((r) => rangeOnDeleted(r, home))) otherReferences++;
    }
    for (const v of sheet.validations) if (v.listSource && isDeleted(v.listSource.sheet)) otherReferences++;
  }
  return { empty, formulas, otherReferences };
}

/** Every formula's text on every sheet except `except`, for structural undo. */
export function captureFormulas(workbook: Workbook, except: Worksheet | null): [ChangedCell, string][] {
  const out: [ChangedCell, string][] = [];
  for (const sheet of workbook.worksheets) {
    if (sheet === except) continue;
    for (const [row, column, cell] of sheet.enumerateCells()) {
      if (cell.formula !== null) out.push([{ sheet: sheet.name, row, column }, cell.formula]);
    }
  }
  return out;
}

export function restoreFormulas(
  workbook: Workbook,
  formulas: readonly (readonly [ChangedCell, string])[],
): void {
  for (const [ref, formula] of formulas) {
    const cell = workbook.findSheet(ref.sheet)?.find(ref.row, ref.column);
    if (cell) cell.formula = formula;
  }
}
