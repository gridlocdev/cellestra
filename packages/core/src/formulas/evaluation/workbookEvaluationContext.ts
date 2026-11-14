import type { CellRange } from '../../model/cellRange.ts';
import { SheetDimensions } from '../../model/sheetDimensions.ts';
import type { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
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
  readonly formulaCell: { readonly row: number; readonly column: number } | undefined;
  readonly #workbook: Workbook;
  readonly #clock: () => CivilDateTime;
  /** Sheets by the name a reference used, so reading a large range looks each sheet up once. */
  readonly #sheets = new Map<string, Worksheet | undefined>();

  constructor(
    workbook: Workbook,
    currentSheet: string,
    locale?: string,
    clock?: () => CivilDateTime,
    formulaCell?: { readonly row: number; readonly column: number },
  ) {
    this.#workbook = workbook;
    this.currentSheet = currentSheet;
    this.formulaCell = formulaCell;
    this.locale = locale ?? workbook.settings.effectiveLocale;
    this.#clock = clock ?? systemClock;
  }

  now(): CivilDateTime {
    return this.#clock();
  }

  #sheet(name: string): Worksheet | undefined {
    if (this.#sheets.has(name)) return this.#sheets.get(name);
    const found = this.#workbook.findSheet(name);
    this.#sheets.set(name, found);
    return found;
  }

  /** Whether `sheet` (null: the current sheet) exists. */
  hasSheet(sheet: string | null): boolean {
    return this.#sheet(sheet ?? this.currentSheet) !== undefined;
  }

  getCell(sheet: string | null, row: number, column: number): CellValue {
    const worksheet = this.#sheet(sheet ?? this.currentSheet);
    if (!worksheet) return CellValue.error(CellError.reference(`Unknown sheet '${sheet ?? ''}'.`));
    if (!SheetDimensions.isValidRow(row) || !SheetDimensions.isValidColumn(column)) {
      return CellValue.error(CellError.reference('Reference is outside the sheet.'));
    }
    return worksheet.getValue(row, column);
  }

  cellExtent(
    sheet: string | null,
  ): { readonly lastRow: number; readonly lastColumn: number } | null | undefined {
    return this.#sheet(sheet ?? this.currentSheet)?.cellExtent();
  }

  sheetsBetween(first: string, last: string): readonly string[] | null {
    return this.#workbook.sheetsBetween(first, last)?.map((s) => s.name) ?? null;
  }

  rowHiddenBy(sheet: string | null, row: number): 'filtered' | 'hidden' | null {
    const worksheet = this.#sheet(sheet ?? this.currentSheet);
    if (!worksheet?.isRowHidden(row)) return null;
    // A hidden row inside a filtered range counts as filtered: the model doesn't record which hid it.
    const inFilter = (range: CellRange | null) => range !== null && row >= range.top && row <= range.bottom;
    return inFilter(worksheet.activeFilter) || worksheet.tables.some((t) => t.filter && inFilter(t.range))
      ? 'filtered'
      : 'hidden';
  }

  cellFormula(sheet: string | null, row: number, column: number): string | null {
    return this.#sheet(sheet ?? this.currentSheet)?.find(row, column)?.formula ?? null;
  }

  resolveName(name: string): CellRange | null {
    return this.#workbook.findNamedRange(name, this.currentSheet)?.target ?? null;
  }
}
