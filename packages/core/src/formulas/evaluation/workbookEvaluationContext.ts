import type { CellRange } from '../../model/cellRange.ts';
import { SheetDimensions } from '../../model/sheetDimensions.ts';
import type { Workbook } from '../../model/workbook.ts';
import { CellError } from '../../values/cellError.ts';
import { CellValue } from '../../values/cellValue.ts';
import { systemClock, type CivilDateTime } from './dateConversion.ts';
import type { EvaluationContext } from './evaluationContext.ts';

/**
 * An `EvaluationContext` backed by a live `Workbook`. Unqualified references resolve against
 * `currentSheet`, cross-sheet references by name. Clock and locale are injectable for determinism.
 */
export class WorkbookEvaluationContext implements EvaluationContext {
  readonly currentSheet: string;
  readonly locale: string;
  readonly #workbook: Workbook;
  readonly #clock: () => CivilDateTime;

  constructor(workbook: Workbook, currentSheet: string, locale?: string, clock?: () => CivilDateTime) {
    this.#workbook = workbook;
    this.currentSheet = currentSheet;
    this.locale = locale ?? workbook.settings.effectiveLocale;
    this.#clock = clock ?? systemClock;
  }

  now(): CivilDateTime {
    return this.#clock();
  }

  getCell(sheet: string | null, row: number, column: number): CellValue {
    const worksheet = this.#workbook.findSheet(sheet ?? this.currentSheet);
    if (!worksheet) return CellValue.error(CellError.reference(`Unknown sheet '${sheet ?? ''}'.`));
    if (!SheetDimensions.isValidRow(row) || !SheetDimensions.isValidColumn(column)) {
      return CellValue.error(CellError.reference('Reference is outside the sheet.'));
    }
    return worksheet.getValue(row, column);
  }

  resolveName(name: string): CellRange | null {
    return this.#workbook.findNamedRange(name, this.currentSheet)?.target ?? null;
  }
}
