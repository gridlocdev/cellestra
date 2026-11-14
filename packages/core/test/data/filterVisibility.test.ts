import { describe, expect, it } from 'vitest';
import { filterSpec } from '../../src/data/filterEngine.ts';
import { DataEditCommands } from '../../src/editing/dataEditCommands.ts';
import { EditCommands } from '../../src/editing/editCommands.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { Evaluator } from '../../src/formulas/evaluation/evaluator.ts';
import { WorkbookEvaluationContext } from '../../src/formulas/evaluation/workbookEvaluationContext.ts';
import { FunctionRegistry } from '../../src/formulas/functions/functionRegistry.ts';
import { Parser } from '../../src/formulas/parsing/parser.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { readEntry } from '../serialization/helpers.ts';
import { OdsXmlNames, readOds, writeOds } from '../../src/serialization/index.ts';

const R = (a1: string): CellRange => CellRange.parse(a1);

/** Score in A1, then 10, 90, 20, 95 in A2:A5; row 3 (90, index 2) hidden by hand. */
function setup(): { workbook: Workbook; history: EditHistory } {
  const workbook = Workbook.createNew();
  const sheet = workbook.activeSheet;
  sheet.setValue(0, 0, CellValue.text('Score'));
  [10, 90, 20, 95].forEach((n, i) => sheet.setValue(i + 1, 0, CellValue.number(n)));
  const history = new EditHistory(workbook);
  history.execute(EditCommands.setRowsHidden(sheet.name, 2, 2, true));
  return { workbook, history };
}

const over40 = filterSpec(new Map([[0, { condition: '>40' }]]));

describe('Filtered rows and rows hidden by hand', () => {
  it('a filter keeps a row hidden by hand hidden, even when the row passes', () => {
    const { workbook, history } = setup();
    const sheet = workbook.activeSheet;
    history.execute(DataEditCommands.applyFilter(sheet.name, R('A1:A5'), over40, 'en-US'));
    expect([1, 2, 3, 4].map((r) => sheet.isRowHidden(r))).toEqual([true, true, true, false]);
  });

  it('clearing the filter shows the filtered rows and keeps the row hidden by hand hidden', () => {
    const { workbook, history } = setup();
    const sheet = workbook.activeSheet;
    history.execute(DataEditCommands.applyFilter(sheet.name, R('A1:A5'), over40, 'en-US'));
    history.execute(DataEditCommands.clearFilter(sheet.name, R('A1:A5')));
    expect([1, 2, 3, 4].map((r) => sheet.isRowHidden(r))).toEqual([false, true, false, false]);
    history.undo();
    history.undo();
    expect([1, 2, 3, 4].map((r) => sheet.isRowHidden(r))).toEqual([false, true, false, false]);
  });

  it('saves filtered rows as visibility="filter" and rows hidden by hand as "collapse"', () => {
    const { workbook, history } = setup();
    const sheet = workbook.activeSheet;
    sheet.activeFilter = R('A1:A5');
    sheet.activeFilterSpec = over40;
    history.execute(DataEditCommands.applyFilter(sheet.name, R('A1:A5'), over40, 'en-US'));
    const bytes = writeOds(workbook);
    const xml = readEntry(bytes, OdsXmlNames.contentPart);
    const visibilities = [...xml.matchAll(/<table:table-row\b[^>]*>/g)].map(
      (m) => /table:visibility="(\w+)"/.exec(m[0])?.[1] ?? 'visible',
    );
    // Row 2 (10) filtered, row 3 hidden by hand (and passing), row 4 (20) filtered.
    expect(visibilities.slice(0, 5)).toEqual(['visible', 'filter', 'collapse', 'filter', 'visible']);

    // Reopened, clearing the filter still leaves the row hidden by hand hidden.
    const reopened = readOds(bytes);
    const again = reopened.activeSheet;
    expect([1, 2, 3, 4].map((r) => again.isRowHidden(r))).toEqual([true, true, true, false]);
    new EditHistory(reopened).execute(DataEditCommands.clearFilter(again.name, R('A1:A5')));
    expect([1, 2, 3, 4].map((r) => again.isRowHidden(r))).toEqual([false, true, false, false]);
  });

  it('SUBTOTAL 1–11 counts a row hidden by hand inside a filtered range; 101–111 leave it out', () => {
    const { workbook, history } = setup();
    const sheet = workbook.activeSheet;
    sheet.activeFilter = R('A1:A5');
    sheet.activeFilterSpec = over40;
    history.execute(DataEditCommands.applyFilter(sheet.name, R('A1:A5'), over40, 'en-US'));
    const evaluate = (formula: string): CellValue =>
      new Evaluator(
        new WorkbookEvaluationContext(workbook, sheet.name, 'en-US', undefined, { row: 9, column: 9 }),
        FunctionRegistry.createDefault(),
      ).evaluate(Parser.parse(formula));
    expect(evaluate('SUBTOTAL(9, A2:A5)')).toEqual(CellValue.number(185));
    expect(evaluate('SUBTOTAL(109, A2:A5)')).toEqual(CellValue.number(95));
  });
});
