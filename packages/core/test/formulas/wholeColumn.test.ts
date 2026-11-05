import { describe, expect, it } from 'vitest';
import { RecalculationEngine } from '../../src/dependencies/recalculationEngine.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';

function bookWith(fill: (set: (row: number, column: number, v: number | string) => void) => void) {
  const workbook = Workbook.createNew();
  const sheet = workbook.activeSheet;
  fill((row, column, v) =>
    typeof v === 'string' && v.startsWith('=')
      ? sheet.setFormula(row, column, v.slice(1))
      : sheet.setValue(row, column, typeof v === 'number' ? CellValue.number(v) : CellValue.text(v)),
  );
  const engine = new RecalculationEngine(workbook);
  engine.rebuildGraph();
  engine.recalculateAll();
  return { sheet, engine };
}

describe('whole-column references', () => {
  it('read only as far as the data, and fast', () => {
    const start = performance.now();
    const { sheet } = bookWith((set) => {
      for (let r = 0; r < 5; r++) set(r, 0, r + 1); // A1:A5 = 1..5
      set(0, 1, 'x');
      set(2, 1, 'x');
      for (let r = 0; r < 200; r++) set(r, 3, '=SUM(A1:A1048576)+COUNT(A1:A1048576)+MAX(A1:A1048576)');
      set(0, 4, '=COUNTIF(A1:A1048576,">2")');
      set(1, 4, '=SUMIF(B1:B1048576,"x",A1:A1048576)');
      set(2, 4, '=COUNTIF(B1:B1048576,"")');
      set(3, 4, '=VLOOKUP(4,A1:A1048576,1,FALSE())');
      set(4, 4, '=MATCH(9,A1:A1048576,1)');
      set(5, 4, '=MATCH("",B1:B1048576,0)');
      set(6, 4, '=SUMIF(B1:B1048576,"",A1:A1048576)');
    });
    expect(performance.now() - start).toBeLessThan(1000);
    const n = (row: number, column: number) => CellValue.asNumber(sheet.getValue(row, column));
    expect(n(0, 3)).toBe(15 + 5 + 5);
    expect(n(0, 4)).toBe(3);
    expect(n(1, 4)).toBe(1 + 3);
    // Every cell of B but the two "x" cells is blank.
    expect(n(2, 4)).toBe(1_048_576 - 2);
    expect(n(3, 4)).toBe(4);
    expect(n(4, 4)).toBe(5);
    expect(n(5, 4)).toBe(2);
    expect(n(6, 4)).toBe(2 + 4 + 5);
  });

  it('recalculates a whole-column formula after the formulas it reads', () => {
    const { sheet, engine } = bookWith((set) => {
      set(0, 1, '=SUM(A1:A1048576)'); // B1 reads A1:A1048576, which holds a formula further down
      set(0, 0, 1);
      set(9, 0, '=A1*10');
    });
    expect(CellValue.asNumber(sheet.getValue(0, 1))).toBe(11);
    sheet.setValue(0, 0, CellValue.number(2));
    engine.recordEdits([{ sheet: sheet.name, row: 0, column: 0 }]);
    expect(CellValue.asNumber(sheet.getValue(0, 1))).toBe(22);
  });

  it('keeps the sheet cell extent current as cells come and go', () => {
    const sheet = Workbook.createNew().activeSheet;
    expect(sheet.cellExtent()).toBeNull();
    sheet.setValue(4, 2, CellValue.number(1));
    sheet.setValue(9, 1, CellValue.number(1));
    expect(sheet.cellExtent()).toEqual({ lastRow: 9, lastColumn: 2 });
    sheet.setValue(9, 1, CellValue.empty);
    expect(sheet.cellExtent()).toEqual({ lastRow: 4, lastColumn: 2 });
    sheet.remove(4, 2);
    expect(sheet.cellExtent()).toBeNull();
  });
});
