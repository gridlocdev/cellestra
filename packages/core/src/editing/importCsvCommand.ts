import type { DataValidation, FilterSpec, SortSpec } from '../data/types.ts';
import type { Cell } from '../model/cell.ts';
import type { CellRange } from '../model/cellRange.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet, WorksheetLayout } from '../model/worksheet.ts';
import type { CsvOptions } from '../serialization/csv/csvOptions.ts';
import { fillFromCsv } from '../serialization/csv/csvReader.ts';
import { EditResult, requireSheet, type EditCommand } from './editCommand.ts';

/** Everything `Worksheet.clear()` empties, kept so undo can put it back. */
interface SheetContents {
  readonly cells: readonly (readonly [row: number, column: number, cell: Cell])[];
  readonly layout: WorksheetLayout;
  readonly validations: readonly DataValidation[];
  readonly filter: CellRange | null;
  readonly filterSpec: FilterSpec | null;
  readonly sort: SortSpec | null;
}

/**
 * Replace a sheet's contents with CSV text (PRD §8.4): the sheet is emptied as if new (cells and
 * layout, like `Worksheet.clear()`), then filled from the CSV, numbers and dates typed with the
 * locale and everything else kept as text. The sheet keeps its name and place. Undo puts the old
 * contents back in one step. `afterFill` runs on the filled sheet within the same step (the app
 * fits the columns there); undo takes its layout changes back too.
 */
export class ImportCsvCommand implements EditCommand {
  readonly label = 'Import CSV';
  private readonly sheetName: string;
  private readonly text: string;
  private readonly options: CsvOptions;
  private readonly afterFill: ((workbook: Workbook, sheet: Worksheet) => void) | null;
  private before: SheetContents | null = null;

  constructor(
    sheetName: string,
    text: string,
    options: CsvOptions,
    afterFill: ((workbook: Workbook, sheet: Worksheet) => void) | null = null,
  ) {
    this.sheetName = sheetName;
    this.text = text;
    this.options = options;
    this.afterFill = afterFill;
  }

  apply(workbook: Workbook): EditResult {
    const sheet = requireSheet(workbook, this.sheetName);
    this.before = {
      cells: [...sheet.enumerateCells()].map(([row, column, cell]) => [row, column, cell.clone()] as const),
      layout: sheet.captureLayout(),
      validations: [...sheet.validations],
      filter: sheet.activeFilter,
      filterSpec: sheet.activeFilterSpec,
      sort: sheet.activeSort,
    };
    sheet.clear();
    fillFromCsv(workbook, sheet, this.text, this.options);
    this.afterFill?.(workbook, sheet);
    return EditResult.structuralChange;
  }

  undo(workbook: Workbook): EditResult {
    const before = this.before;
    if (!before) return EditResult.none;
    const sheet = requireSheet(workbook, this.sheetName);
    sheet.clear();
    for (const [row, column, cell] of before.cells) sheet.putCell(row, column, cell.clone());
    sheet.restoreLayout(before.layout);
    sheet.replaceValidations(before.validations);
    sheet.activeFilter = before.filter;
    sheet.activeFilterSpec = before.filterSpec;
    sheet.activeSort = before.sort;
    return EditResult.structuralChange;
  }
}
