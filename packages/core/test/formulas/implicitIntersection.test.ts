import { describe, expect, it } from 'vitest';
import { RecalculationEngine } from '../../src/dependencies/recalculationEngine.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { CellAddress } from '../../src/model/cellAddress.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';

/** A1:A3 = 10, 20, 30 and B1:B3 = 100, 200, 300; each formula goes in at its address, then the book recalculates. */
function shown(formulas: Record<string, string>): Record<string, string> {
  const workbook = Workbook.createNew();
  const sheet = workbook.activeSheet;
  [10, 20, 30].forEach((v, r) => sheet.setValue(r, 0, CellValue.number(v)));
  [100, 200, 300].forEach((v, r) => sheet.setValue(r, 1, CellValue.number(v)));
  workbook.addNamedRange({ name: 'Column', target: CellRange.parse('Sheet1.A1:A3') });
  for (const [at, formula] of Object.entries(formulas)) {
    const a = CellAddress.parse(at);
    sheet.setFormula(a.row, a.column, formula);
  }
  const engine = new RecalculationEngine(workbook);
  engine.rebuildGraph();
  engine.recalculateAll();
  const out: Record<string, string> = {};
  for (const at of Object.keys(formulas)) {
    const a = CellAddress.parse(at);
    const v = sheet.getValue(a.row, a.column);
    out[at] = v.kind === 'error' ? v.error.token : CellValue.toInvariantString(v);
  }
  return out;
}

// Each case is what LibreOffice Calc 26.2 shows for the same formula in the same cell.
describe('A range used as one value takes the cell in the formula’s row or column', () => {
  it.each([
    ['D2', 'A1:A3', '20'],
    ['D3', 'A1:A3*2', '60'],
    ['D5', 'A1:A3', '#VALUE!'],
    ['B7', 'A1:B1', '100'],
    ['D7', 'A1:B1', '#VALUE!'],
    ['E2', 'ABS(A1:A3)', '20'],
    ['E3', 'INDEX(A1:B3, 0, 1)', '30'],
    ['E4', 'INDEX(A1:B3, 0, 1)', '#VALUE!'],
    ['F2', 'IF(A1:A3>15, 1, 0)', '1'],
    ['G9', 'IFERROR(A1:A3, -1)', '-1'],
    ['H3', 'Column+1', '31'],
    ['H2', 'SUM(A1:A3)', '60'],
  ])('%s: =%s shows %s', (at, formula, expected) => {
    expect(shown({ [at]: formula })[at]).toBe(expected);
  });
});
