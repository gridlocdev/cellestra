import { describe, expect, it } from 'vitest';
import { CellRef } from '../../src/dependencies/cellRef.ts';
import { RecalculationEngine } from '../../src/dependencies/recalculationEngine.ts';
import { EditResult } from '../../src/editing/editCommand.ts';
import { moveSheet } from '../../src/editing/moveSheet.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';

/** Sheet1, Sheet2, Sheet3 with 1, 10, 100 in A1, and a sheet Out. */
function book() {
  const workbook = Workbook.createNew();
  const values = [1, 10, 100];
  workbook.worksheets[0]!.setValue(0, 0, CellValue.number(values[0]!));
  for (const [i, v] of values.slice(1).entries())
    workbook.addSheet(`Sheet${i + 2}`).setValue(0, 0, CellValue.number(v));
  const out = workbook.addSheet('Out');
  const engine = new RecalculationEngine(workbook);
  return { workbook, out, engine };
}

describe('moveSheet and 3D ranges', () => {
  it('a range spans the sheets now between its ends, and edits there recalculate it', () => {
    const { workbook, out, engine } = book();
    out.setFormula(0, 0, 'SUM(Sheet1.A1:Sheet2.A1)');
    engine.rebuildGraph();
    engine.recalculateAll();
    expect(out.getValue(0, 0)).toEqual(CellValue.number(11));

    moveSheet(workbook, 2, 1); // Sheet3 moves in between
    engine.apply(EditResult.structuralChange);
    expect(out.getValue(0, 0)).toEqual(CellValue.number(111));
    expect(out.find(0, 0)?.formula).toBe('SUM(Sheet1.A1:Sheet2.A1)');

    workbook.findSheet('Sheet3')!.setValue(0, 0, CellValue.number(1000));
    engine.recordEdits([CellRef.of('Sheet3', 0, 0)]);
    expect(out.getValue(0, 0)).toEqual(CellValue.number(1011));
  });

  it('when the last sheet moves before the first, the text swaps them to read in tab order', () => {
    const { workbook, out, engine } = book();
    out.setFormula(0, 0, 'SUM(Sheet1.A1:Sheet2.A1)');
    out.setFormula(1, 0, 'A1*2');
    workbook.addNamedRange({ name: 'Span', expression: 'Sheet1.$A$1:Sheet2.$A$1' });
    out.setFormula(2, 0, 'SUM(Span)');
    moveSheet(workbook, 1, 0); // Sheet2, Sheet1, Sheet3, Out
    engine.rebuildGraph();
    engine.recalculateAll();
    expect(out.find(0, 0)?.formula).toBe('SUM(Sheet2.A1:Sheet1.A1)');
    expect(out.find(1, 0)?.formula).toBe('A1*2');
    expect(workbook.findNamedRange('Span')?.expression).toBe('Sheet2.$A$1:Sheet1.$A$1');
    expect(out.getValue(0, 0)).toEqual(CellValue.number(11));
    expect(out.getValue(2, 0)).toEqual(CellValue.number(11));
  });
});
