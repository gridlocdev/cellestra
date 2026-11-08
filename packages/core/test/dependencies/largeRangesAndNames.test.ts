import { describe, expect, it } from 'vitest';
import { CellRef } from '../../src/dependencies/cellRef.ts';
import { RecalculationEngine } from '../../src/dependencies/recalculationEngine.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';

describe('large ranges', () => {
  it('running totals stay linear in graph size and correct', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    const rows = 4000;
    for (let r = 0; r < rows; r++) {
      sheet.setValue(r, 0, CellValue.number(1));
      sheet.setFormula(r, 1, `SUM($A$1:A${r + 1})`);
    }
    const engine = new RecalculationEngine(workbook);
    engine.rebuildGraph();
    engine.recalculateAll();
    expect(CellValue.asNumber(sheet.getValue(rows - 1, 1))).toBe(rows);

    // An edit near the top reaches every total below it, through both small and large ranges.
    sheet.setValue(1, 0, CellValue.number(11));
    engine.recordEdits([CellRef.of(sheet.name, 1, 0)]);
    expect(CellValue.asNumber(sheet.getValue(0, 1))).toBe(1);
    expect(CellValue.asNumber(sheet.getValue(1, 1))).toBe(12);
    expect(CellValue.asNumber(sheet.getValue(rows - 1, 1))).toBe(rows + 10);
    expect(engine.formulasReferencing([CellRef.of(sheet.name, 2999, 0)])).toHaveLength(rows - 2999);
    // A range past the expansion threshold keeps only its corners in the graph (it was one edge per
    // cell: 8 million edges for these 4000 totals).
    expect(engine.graph.getPrecedents(CellRef.of(sheet.name, rows - 1, 1))).toHaveLength(2);
  });

  it('a formula re-registered without its large range no longer recalculates on edits there', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setFormula(0, 1, 'SUM(A1:A5000)');
    const engine = new RecalculationEngine(workbook);
    engine.rebuildGraph();
    expect(engine.formulasReferencing([CellRef.of(sheet.name, 2500, 0)])).toHaveLength(1);
    sheet.setFormula(0, 1, '1');
    engine.recordEdits([CellRef.of(sheet.name, 0, 1)]);
    expect(engine.formulasReferencing([CellRef.of(sheet.name, 2500, 0)])).toHaveLength(0);
  });

  it('whole columns and rows evaluate and recalculate', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.number(2));
    sheet.setValue(9, 0, CellValue.number(3));
    sheet.setFormula(0, 2, 'SUM(A:A)');
    sheet.setFormula(1, 2, 'SUM(10:10)');
    const engine = new RecalculationEngine(workbook);
    engine.rebuildGraph();
    engine.recalculateAll();
    expect(CellValue.asNumber(sheet.getValue(0, 2))).toBe(5);
    expect(CellValue.asNumber(sheet.getValue(1, 2))).toBe(3);
    sheet.setValue(5000, 0, CellValue.number(10));
    engine.recordEdits([CellRef.of(sheet.name, 5000, 0)]);
    expect(CellValue.asNumber(sheet.getValue(0, 2))).toBe(15);
  });
});

describe('large ranges over formulas', () => {
  // Formulas inside each other's large ranges used to cost one ordering edge (and one index scan)
  // per formula per range: quadratic, seconds for a few thousand rows.
  it('8000 running totals over a formula column rebuild, recalculate and edit quickly', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    const rows = 8000;
    sheet.setValue(0, 0, CellValue.number(1));
    for (let r = 1; r < rows; r++) sheet.setFormula(r, 0, `A${r}+1`);
    // IFERROR never reads its fallback, so this times the dependency tracking, not 32M cell reads.
    for (let r = 0; r < rows; r++) sheet.setFormula(r, 1, `IFERROR(A${r + 1},SUM($A$1:A${r + 1}))`);
    const started = performance.now();
    const engine = new RecalculationEngine(workbook);
    engine.rebuildGraph();
    engine.recalculateAll();
    sheet.setValue(0, 0, CellValue.number(2));
    const edit = engine.recordEdits([CellRef.of(sheet.name, 0, 0)]);
    const elapsed = performance.now() - started;
    expect(edit.evaluated).toBe(2 * rows - 1);
    expect(CellValue.asNumber(sheet.getValue(rows - 1, 1))).toBe(rows + 1);
    expect(elapsed).toBeLessThan(3000); // ~0.3s here; it was several seconds
  });

  it('a large range waits for every formula inside it, wherever the range starts', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    const rows = 3000;
    // Column A: a chain; columns C/D: sums over windows of A that start part-way down.
    sheet.setValue(0, 0, CellValue.number(1));
    for (let r = 1; r < rows; r++) sheet.setFormula(r, 0, `A${r}+1`);
    sheet.setFormula(0, 2, `SUM(A1001:A${rows})`);
    sheet.setFormula(1, 2, `SUM(A1:A${rows})`);
    sheet.setFormula(2, 2, `SUM(A1500:B2700)`);
    sheet.setFormula(0, 3, `C1+C2`);
    const engine = new RecalculationEngine(workbook);
    engine.rebuildGraph();
    const sum = (from: number, to: number): number => ((from + to) * (to - from + 1)) / 2;
    engine.recalculateAll();
    expect(CellValue.asNumber(sheet.getValue(0, 2))).toBe(sum(1001, rows));
    expect(CellValue.asNumber(sheet.getValue(1, 2))).toBe(sum(1, rows));
    expect(CellValue.asNumber(sheet.getValue(2, 2))).toBe(sum(1500, 2700));
    sheet.setValue(0, 0, CellValue.number(2));
    engine.recordEdits([CellRef.of(sheet.name, 0, 0)]);
    expect(CellValue.asNumber(sheet.getValue(0, 3))).toBe(sum(1002, rows + 1) + sum(2, rows + 1));
    expect(CellValue.asNumber(sheet.getValue(2, 2))).toBe(sum(1501, 2701));
  });

  it('a large range over its own cell is still circular', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    for (let r = 0; r < 2000; r++) sheet.setFormula(r, 0, '1');
    sheet.setFormula(1500, 0, 'SUM(A1:A2000)');
    sheet.setFormula(0, 1, 'A1501+1');
    const engine = new RecalculationEngine(workbook);
    engine.rebuildGraph();
    const result = engine.recalculateAll();
    expect(result.circularCells.map((c) => [c.row, c.column]).sort()).toEqual([
      [0, 1],
      [1500, 0],
    ]);
    expect(CellValue.asNumber(sheet.getValue(1999, 0))).toBe(1);
  });
});

describe('sheet-local names in the dependency graph', () => {
  it('a local name shadows the global one and its edits recalculate', () => {
    const workbook = Workbook.createNew();
    const s1 = workbook.activeSheet;
    const s2 = workbook.addSheet('Sheet2');
    workbook.addNamedRange({ name: 'Rate', target: CellRange.parse('Sheet1.A1') });
    workbook.addNamedRange({ name: 'Rate', target: CellRange.parse('Sheet2.A1'), scope: 'Sheet2' });
    s1.setValue(0, 0, CellValue.number(1));
    s2.setValue(0, 0, CellValue.number(2));
    s1.setFormula(1, 0, 'Rate*10');
    s2.setFormula(1, 0, 'Rate*10');
    const engine = new RecalculationEngine(workbook);
    engine.rebuildGraph();
    engine.recalculateAll();
    expect(CellValue.asNumber(s1.getValue(1, 0))).toBe(10);
    expect(CellValue.asNumber(s2.getValue(1, 0))).toBe(20);

    s2.setValue(0, 0, CellValue.number(5));
    engine.recordEdits([CellRef.of('Sheet2', 0, 0)]);
    expect(CellValue.asNumber(s2.getValue(1, 0))).toBe(50);
    expect(CellValue.asNumber(s1.getValue(1, 0))).toBe(10);
  });
});
