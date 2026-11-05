import { describe, expect, it } from 'vitest';
import { AutoSumPlanner, type AutoSumProposal } from '../../src/editing/autoSumPlanner.ts';
import { CellAddress } from '../../src/model/cellAddress.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import type { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';

const fresh = (): Worksheet => Workbook.createNew().activeSheet;

function planAt(sheet: Worksheet, activeCell: string): AutoSumProposal[] {
  const address = CellAddress.parse(activeCell);
  return AutoSumPlanner.plan(sheet, new CellRange(address), address.row, address.column);
}

describe('AutoSumPlanner', () => {
  it('single cell sums the contiguous run above', () => {
    const sheet = fresh();
    sheet.setValue(0, 0, CellValue.number(1));
    sheet.setValue(1, 0, CellValue.number(2));
    sheet.setValue(2, 0, CellValue.number(3));
    expect(planAt(sheet, 'A4')).toEqual([{ targetRow: 3, targetColumn: 0, formula: '=SUM(A1:A3)' }]);
  });

  it('run above stops at the first gap', () => {
    const sheet = fresh();
    sheet.setValue(0, 0, CellValue.number(1)); // A2 blank breaks the run
    sheet.setValue(2, 0, CellValue.number(3));
    sheet.setValue(3, 0, CellValue.number(4));
    const plan = planAt(sheet, 'A5');
    expect(plan).toHaveLength(1);
    expect(plan[0]!.formula).toBe('=SUM(A3:A4)');
  });

  it('falls back to the run to the left when nothing is above', () => {
    const sheet = fresh();
    sheet.setValue(2, 0, CellValue.number(1));
    sheet.setValue(2, 1, CellValue.number(2));
    sheet.setValue(2, 2, CellValue.number(3));
    const plan = planAt(sheet, 'D3');
    expect(plan).toHaveLength(1);
    expect(plan[0]!.formula).toBe('=SUM(A3:C3)');
  });

  it('prefers above over left when both are present', () => {
    const sheet = fresh();
    sheet.setValue(0, 2, CellValue.number(10));
    sheet.setValue(1, 2, CellValue.number(20));
    sheet.setValue(2, 0, CellValue.number(1));
    sheet.setValue(2, 1, CellValue.number(2));
    const plan = planAt(sheet, 'C3');
    expect(plan).toHaveLength(1);
    expect(plan[0]!.formula).toBe('=SUM(C1:C2)'); // above wins
  });

  it('no numbers above or left yields no proposal', () => {
    expect(planAt(fresh(), 'B2')).toEqual([]);
  });

  it('multi-cell selection totals each column into the row below', () => {
    const sheet = fresh();
    sheet.setValue(0, 0, CellValue.number(1));
    sheet.setValue(1, 0, CellValue.number(2));
    sheet.setValue(0, 1, CellValue.number(3));
    sheet.setValue(1, 1, CellValue.number(4));
    const plan = AutoSumPlanner.plan(sheet, CellRange.parse('A1:B2'), 0, 0);
    expect(plan).toHaveLength(2);
    expect(plan[0]!.formula).toBe('=SUM(A1:A2)');
    expect(plan[0]!.targetRow).toBe(2); // row 3 (below the selection)
    expect(plan[0]!.targetColumn).toBe(0);
    expect(plan[1]!.formula).toBe('=SUM(B1:B2)');
    expect(plan[1]!.targetColumn).toBe(1);
  });

  it('multi-cell selection skips non-numeric columns', () => {
    const sheet = fresh();
    sheet.setValue(0, 0, CellValue.text('Label'));
    sheet.setValue(1, 0, CellValue.text('Total'));
    sheet.setValue(0, 1, CellValue.number(3));
    sheet.setValue(1, 1, CellValue.number(4));
    const plan = AutoSumPlanner.plan(sheet, CellRange.parse('A1:B2'), 0, 0);
    expect(plan).toHaveLength(1); // only the numeric column B
    expect(plan[0]!.formula).toBe('=SUM(B1:B2)');
    expect(plan[0]!.targetColumn).toBe(1);
  });
});
