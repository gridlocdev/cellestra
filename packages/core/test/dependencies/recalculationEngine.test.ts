import { describe, expect, it } from 'vitest';
import { CellRef } from '../../src/dependencies/cellRef.ts';
import { RecalculationEngine } from '../../src/dependencies/recalculationEngine.ts';
import { Workbook } from '../../src/model/workbook.ts';
import type { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';

function newBook(): { workbook: Workbook; sheet: Worksheet } {
  const workbook = Workbook.createNew();
  return { workbook, sheet: workbook.activeSheet };
}

function engineFor(workbook: Workbook): RecalculationEngine {
  const engine = new RecalculationEngine(workbook);
  engine.rebuildGraph();
  return engine;
}

const numberAt = (sheet: Worksheet, row: number, column: number): number =>
  CellValue.asNumber(sheet.getValue(row, column));

function errorAt(sheet: Worksheet, row: number, column: number): { token: string; reason: string } {
  const v = sheet.getValue(row, column);
  if (v.kind !== 'error') throw new Error(`Expected an error, got ${v.kind}.`);
  return { token: v.error.token, reason: v.error.reason };
}

describe('RecalculationEngine', () => {
  it('recalculateAll computes formula cells', () => {
    const { workbook, sheet } = newBook();
    sheet.setValue(0, 0, CellValue.number(1));
    sheet.setValue(1, 0, CellValue.number(2));
    sheet.setFormula(2, 0, 'A1+A2');

    engineFor(workbook).recalculateAll();

    expect(numberAt(sheet, 2, 0)).toBe(3);
  });

  it('recalculateAll orders a chain', () => {
    const { workbook, sheet } = newBook();
    sheet.setValue(0, 0, CellValue.number(1));
    sheet.setFormula(0, 1, 'A1+1'); // B1 = A1 + 1
    sheet.setFormula(0, 2, 'B1+1'); // C1 = B1 + 1

    engineFor(workbook).recalculateAll();

    expect(numberAt(sheet, 0, 1)).toBe(2);
    expect(numberAt(sheet, 0, 2)).toBe(3);
  });

  it('an edit propagates to dependents', () => {
    const { workbook, sheet } = newBook();
    sheet.setValue(0, 0, CellValue.number(1));
    sheet.setFormula(0, 1, 'A1+1');
    sheet.setFormula(0, 2, 'B1+1');
    const engine = engineFor(workbook);
    engine.recalculateAll();

    sheet.setValue(0, 0, CellValue.number(10));
    const result = engine.recordEdits([CellRef.of(sheet.name, 0, 0)]);

    expect(numberAt(sheet, 0, 1)).toBe(11);
    expect(numberAt(sheet, 0, 2)).toBe(12);
    expect(result.evaluated).toBe(2);
  });

  it('a circular reference is detected and reported', () => {
    const { workbook, sheet } = newBook();
    sheet.setFormula(0, 0, 'B1');
    sheet.setFormula(0, 1, 'A1');

    const result = engineFor(workbook).recalculateAll();

    expect(result.hasCircularReference).toBe(true);
    expect(errorAt(sheet, 0, 0)).toEqual({ token: '#REF!', reason: 'Circular reference.' });
    expect(errorAt(sheet, 0, 1).token).toBe('#REF!');
  });

  it('an incremental edit inside a large range recomputes the formula', () => {
    const { workbook, sheet } = newBook();
    // A range over 200k cells is summarized by its corners in the graph; an edit to an interior
    // cell must still recompute the formula on an incremental pass.
    sheet.setValue(0, 0, CellValue.number(1)); // A1
    sheet.setValue(299_999, 0, CellValue.number(1)); // A300000
    sheet.setFormula(0, 1, 'SUM(A1:A300000)'); // B1
    const engine = engineFor(workbook);
    engine.recalculateAll();
    expect(numberAt(sheet, 0, 1)).toBe(2);

    sheet.setValue(149_999, 0, CellValue.number(40)); // A150000 — interior, not a corner
    engine.recordEdits([CellRef.of(sheet.name, 149_999, 0)]);

    expect(numberAt(sheet, 0, 1)).toBe(42);
  });

  it('a direct self reference is detected as circular', () => {
    const { workbook, sheet } = newBook();
    sheet.setFormula(0, 0, 'A1+1');

    const result = engineFor(workbook).recalculateAll();

    expect(result.hasCircularReference).toBe(true);
    expect(errorAt(sheet, 0, 0).token).toBe('#REF!');
  });

  it('a range self reference is detected as circular', () => {
    const { workbook, sheet } = newBook();
    // A formula whose range covers its own cell is circular (Excel/Calc flag it the same way).
    sheet.setFormula(0, 0, 'SUM(A1:A3)');

    const result = engineFor(workbook).recalculateAll();

    expect(result.hasCircularReference).toBe(true);
    expect(errorAt(sheet, 0, 0).token).toBe('#REF!');
  });

  it('manual mode defers recalculation', () => {
    const { workbook, sheet } = newBook();
    workbook.settings.calculationMode = 'manual';
    sheet.setValue(0, 0, CellValue.number(1));
    sheet.setFormula(0, 1, 'A1+1', CellValue.number(99));

    const result = engineFor(workbook).recordEdits([CellRef.of(sheet.name, 0, 0)]);

    expect(result.evaluated).toBe(0);
    expect(numberAt(sheet, 0, 1)).toBe(99); // still the cached value
  });

  it('manual mode defers recalculation after a structural change, but re-plans the graph', () => {
    const { workbook, sheet } = newBook();
    workbook.settings.calculationMode = 'manual';
    sheet.setValue(0, 0, CellValue.number(1));
    sheet.setFormula(0, 1, 'A1+1', CellValue.number(99));
    const engine = engineFor(workbook);
    // As if a structural edit had rewritten the formula: it now reads A2.
    sheet.setFormula(0, 1, 'A2+1', CellValue.number(99));
    sheet.setValue(1, 0, CellValue.number(5));

    const result = engine.recordStructuralChange();
    expect(result.evaluated).toBe(0);
    expect(numberAt(sheet, 0, 1)).toBe(99);
    // The graph follows the new text, so an edit to A2 is tracked and a full recalc uses it.
    expect(engine.formulasReferencing([CellRef.of(sheet.name, 1, 0)])).toHaveLength(1);
    engine.recalculateAll();
    expect(numberAt(sheet, 0, 1)).toBe(6);

    workbook.settings.calculationMode = 'automatic';
    sheet.setValue(1, 0, CellValue.number(10));
    expect(engine.recordStructuralChange().evaluated).toBe(1);
    expect(numberAt(sheet, 0, 1)).toBe(11);
  });

  it('a cross-sheet reference recalculates', () => {
    const workbook = Workbook.createNew();
    const sheet1 = workbook.activeSheet;
    const sheet2 = workbook.addSheet();
    sheet2.setValue(0, 0, CellValue.number(42));
    sheet1.setFormula(0, 0, `${sheet2.name}.A1*2`);

    engineFor(workbook).recalculateAll();

    expect(numberAt(sheet1, 0, 0)).toBe(84);
  });

  it('a volatile function is tracked', () => {
    const { workbook, sheet } = newBook();
    sheet.setFormula(0, 0, 'TODAY()');

    const engine = engineFor(workbook);

    expect(engine.volatileCells).toContainEqual(CellRef.of(sheet.name, 0, 0));
    expect(engine.isVolatile(CellRef.of(sheet.name, 0, 0))).toBe(true);
  });

  it('recalculateVolatile re-evaluates volatile cells and their dependents', () => {
    const { workbook, sheet } = newBook();
    sheet.setFormula(0, 0, 'TODAY()');
    sheet.setFormula(0, 1, 'A1+1');
    let day = 7;
    const engine = new RecalculationEngine(workbook, undefined, () => ({
      year: 2026,
      month: 6,
      day,
      hour: 0,
      minute: 0,
      second: 0,
      millisecond: 0,
    }));
    engine.rebuildGraph();
    engine.recalculateAll();
    const before = numberAt(sheet, 0, 1);

    day = 8;
    engine.recalculateVolatile();

    expect(numberAt(sheet, 0, 1)).toBe(before + 1);
  });

  it('an unrelated edit recomputes volatile cells and what reads them', () => {
    const { workbook, sheet } = newBook();
    sheet.setFormula(0, 0, 'NOW()');
    sheet.setFormula(0, 1, 'A1*1440');
    sheet.setFormula(0, 2, 'B1+1');
    let minute = 0;
    const engine = new RecalculationEngine(workbook, undefined, () => ({
      year: 2026,
      month: 1,
      day: 1,
      hour: 0,
      minute,
      second: 0,
      millisecond: 0,
    }));
    engine.rebuildGraph();
    engine.recalculateAll();
    const before = numberAt(sheet, 0, 2);

    minute = 5;
    sheet.setValue(4, 4, CellValue.number(3));
    engine.recordEdits([CellRef.of(sheet.name, 4, 4)]);

    expect(numberAt(sheet, 0, 2)).toBeCloseTo(before + 5, 6);
  });

  it("keeps the value a file cached for a formula it can't parse", () => {
    const { workbook, sheet } = newBook();
    sheet.setFormula(0, 0, 'SUM(1;;', CellValue.number(42));
    sheet.setFormula(1, 0, 'SUM(1;;');

    engineFor(workbook).recalculateAll();

    expect(numberAt(sheet, 0, 0)).toBe(42);
    expect(errorAt(sheet, 1, 0).token).toBe('#VALUE!');
  });

  it('keeps the value a file cached for a formula calling a function it does not know', () => {
    // LibreOffice computed it, so its result beats #NAME? (and a save writes it back).
    const { workbook, sheet } = newBook();
    sheet.setValue(0, 1, CellValue.number(5));
    sheet.setFormula(0, 0, '1+NOSUCHFUNCTION(B1)', CellValue.number(42));
    sheet.setFormula(1, 0, 'NOSUCHFUNCTION(B1)');
    const engine = engineFor(workbook);
    engine.recalculateAll();

    expect(numberAt(sheet, 0, 0)).toBe(42);
    expect(errorAt(sheet, 1, 0).token).toBe('#NAME?');

    sheet.setValue(0, 1, CellValue.number(6));
    engine.recordEdits([CellRef.of(sheet.name, 0, 1)]);
    expect(numberAt(sheet, 0, 0)).toBe(42);
  });

  it('shows #NAME? for an unknown function once the formula is typed again', () => {
    const { workbook, sheet } = newBook();
    sheet.setFormula(0, 0, 'NOSUCHFUNCTION(1)', CellValue.number(42));
    const engine = engineFor(workbook);
    engine.recalculateAll();
    expect(numberAt(sheet, 0, 0)).toBe(42);

    // Entering a formula leaves no cached value.
    sheet.setFormula(0, 0, 'NOSUCHFUNCTION(2)', CellValue.empty);
    engine.recordEdits([CellRef.of(sheet.name, 0, 0)]);
    expect(errorAt(sheet, 0, 0).token).toBe('#NAME?');
  });

  it('SUMIF follows edits to the cells its sum range is reshaped onto', () => {
    const { workbook, sheet } = newBook();
    for (let r = 0; r < 3; r++) sheet.setValue(r, 0, CellValue.text('x'));
    sheet.setValue(0, 1, CellValue.number(1));
    sheet.setFormula(0, 3, 'SUMIF(A1:A3,"x",B1)');
    const engine = engineFor(workbook);
    engine.recalculateAll();
    expect(numberAt(sheet, 0, 3)).toBe(1);

    sheet.setValue(2, 1, CellValue.number(10));
    engine.recordEdits([CellRef.of(sheet.name, 2, 1)]);

    expect(numberAt(sheet, 0, 3)).toBe(11);
  });
});
