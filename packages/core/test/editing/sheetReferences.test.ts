import { createChart, createSeries } from '@cellestra/charts';
import { describe, expect, it } from 'vitest';
import { dataValidation, table } from '../../src/data/types.ts';
import { RecalculationEngine } from '../../src/dependencies/recalculationEngine.ts';
import { ClipboardBlock } from '../../src/editing/clipboardBlock.ts';
import { DeleteSheetCommand } from '../../src/editing/deleteSheetCommand.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { PasteBlockCommand } from '../../src/editing/pasteBlockCommand.ts';
import { RenameSheetCommand } from '../../src/editing/renameSheetCommand.ts';
import { StructuralEditCommand } from '../../src/editing/structuralEditCommand.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';

const R = (a1: string): CellRange => CellRange.parse(a1);
const str = (r: CellRange | null | undefined): string | null => r?.toString() ?? null;

function recalc(workbook: Workbook): void {
  const engine = new RecalculationEngine(workbook);
  engine.rebuildGraph();
  engine.recalculateAll();
}

describe('RenameSheetCommand', () => {
  function book() {
    const workbook = Workbook.createNew();
    const data = workbook.activeSheet;
    const other = workbook.addSheet('Other');
    data.setValue(0, 0, CellValue.number(4));
    other.setFormula(0, 0, 'Sheet1.A1*2');
    other.setFormula(1, 0, 'SUM(Sheet1.A1:A3)+Total');
    data.setFormula(1, 0, 'A1+Sheet1.A1');
    workbook.addNamedRange({ name: 'Total', target: R('Sheet1.A1') });
    workbook.addNamedRange({ name: 'Local', target: R('Sheet1.A1:A2'), scope: 'Sheet1' });
    other.addValidation(dataValidation('list', R('B1'), { listSource: R('Sheet1.A1:A3') }));
    const chart = workbook.addChartSheet(
      createChart({
        categoriesReference: 'Sheet1.A1:A3',
        series: [createSeries({ valuesReference: 'Sheet1.B1:B3', nameReference: 'Sheet1.B1' })],
      }),
    );
    return { workbook, data, other, chart };
  }

  it('rewrites formulas, names, validations and charts, and undoes', () => {
    const { workbook, data, other, chart } = book();
    const history = new EditHistory(workbook);
    history.execute(new RenameSheetCommand('Sheet1', 'Data 2024'));
    recalc(workbook);

    expect(data.name).toBe('Data 2024');
    expect(other.find(0, 0)!.formula).toBe("'Data 2024'.A1*2");
    expect(other.find(1, 0)!.formula).toBe("SUM('Data 2024'.A1:A3)+Total");
    expect(data.find(1, 0)!.formula).toBe("A1+'Data 2024'.A1");
    expect(CellValue.asNumber(other.getValue(0, 0))).toBe(8);
    expect(CellValue.asNumber(other.getValue(1, 0))).toBe(4 + 8 + 4);
    expect(str(workbook.findNamedRange('Total')!.target)).toBe("'Data 2024'.A1");
    const local = workbook.findNamedRange('Local', 'Data 2024')!;
    expect(local.scope).toBe('Data 2024');
    expect(str(local.target)).toBe("'Data 2024'.A1:A2");
    expect(str(other.validations[0]!.listSource)).toBe("'Data 2024'.A1:A3");
    expect(chart.chart!.categoriesReference).toBe("'Data 2024'.A1:A3");
    expect(chart.chart!.series[0]!.valuesReference).toBe("'Data 2024'.B1:B3");
    expect(chart.chart!.series[0]!.nameReference).toBe("'Data 2024'.B1");

    history.undo();
    expect(data.name).toBe('Sheet1');
    expect(other.find(0, 0)!.formula).toBe('Sheet1.A1*2');
    expect(str(workbook.findNamedRange('Total')!.target)).toBe('Sheet1.A1');
    expect(workbook.findNamedRange('Local', 'Sheet1')!.scope).toBe('Sheet1');
    expect(str(other.validations[0]!.listSource)).toBe('Sheet1.A1:A3');
    expect(chart.chart!.series[0]!.valuesReference).toBe('Sheet1.B1:B3');

    history.redo();
    expect(data.name).toBe('Data 2024');
    expect(other.find(0, 0)!.formula).toBe("'Data 2024'.A1*2");
  });

  it('rejects a duplicate name without changing anything', () => {
    const { workbook, other } = book();
    expect(() => new RenameSheetCommand('Sheet1', 'other').apply(workbook)).toThrow();
    expect(other.find(0, 0)!.formula).toBe('Sheet1.A1*2');
  });
});

describe('StructuralEditCommand moves references outside formulas', () => {
  function book() {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    workbook.addNamedRange({ name: 'Prices', target: R('Sheet1.B2:B5') });
    workbook.addNamedRange({ name: 'Gone', target: R('Sheet1.A3') });
    sheet.addValidation(dataValidation('list', R('C2:C5'), { listSource: R('Sheet1.E1:E3') }));
    sheet.activeFilter = R('A1:D5');
    sheet.activeFilterSpec = { hasHeader: true, columns: new Map([[3, { condition: '>1' }]]) };
    sheet.activeSort = { hasHeader: true, levels: [{ column: 3, ascending: true }] };
    sheet.addTable(
      table('T', R('F1:H5'), { filter: { hasHeader: true, columns: new Map([[7, { condition: 'x' }]]) } }),
    );
    const chart = workbook.addChartSheet(
      createChart({ series: [createSeries({ valuesReference: 'Sheet1.B2:B5' })] }),
    );
    return { workbook, sheet, chart, history: new EditHistory(workbook) };
  }

  it('inserting rows shifts names, validations, the filter and charts; undo restores them', () => {
    const { workbook, sheet, chart, history } = book();
    history.execute(new StructuralEditCommand(sheet.name, 'row', 'insert', 0, 2));
    expect(str(workbook.findNamedRange('Prices')!.target)).toBe('Sheet1.B4:B7');
    expect(str(sheet.validations[0]!.target)).toBe('C4:C7');
    expect(str(sheet.validations[0]!.listSource)).toBe('Sheet1.E3:E5');
    expect(str(sheet.activeFilter)).toBe('A3:D7');
    expect([...sheet.activeFilterSpec!.columns.keys()]).toEqual([3]);
    expect(chart.chart!.series[0]!.valuesReference).toBe('Sheet1.B4:B7');

    history.undo();
    expect(str(workbook.findNamedRange('Prices')!.target)).toBe('Sheet1.B2:B5');
    expect(str(sheet.validations[0]!.target)).toBe('C2:C5');
    expect(str(sheet.activeFilter)).toBe('A1:D5');
    expect(chart.chart!.series[0]!.valuesReference).toBe('Sheet1.B2:B5');
  });

  it('inserting columns re-keys filter, sort and table criteria', () => {
    const { workbook, sheet, chart, history } = book();
    history.execute(new StructuralEditCommand(sheet.name, 'column', 'insert', 1, 1));
    expect(str(workbook.findNamedRange('Prices')!.target)).toBe('Sheet1.C2:C5');
    expect(str(sheet.activeFilter)).toBe('A1:E5');
    expect([...sheet.activeFilterSpec!.columns.keys()]).toEqual([4]);
    expect(sheet.activeSort!.levels[0]!.column).toBe(4);
    expect(str(sheet.tables[0]!.range)).toBe('G1:I5');
    expect([...sheet.tables[0]!.filter!.columns.keys()]).toEqual([8]);
    expect(str(sheet.validations[0]!.target)).toBe('D2:D5');
    expect(chart.chart!.series[0]!.valuesReference).toBe('Sheet1.C2:C5');

    history.undo();
    expect([...sheet.activeFilterSpec!.columns.keys()]).toEqual([3]);
    expect([...sheet.tables[0]!.filter!.columns.keys()]).toEqual([7]);
    expect(sheet.activeSort!.levels[0]!.column).toBe(3);
  });

  it('deleting rows shrinks ranges and drops what was wholly deleted', () => {
    const { workbook, sheet, history } = book();
    history.execute(new StructuralEditCommand(sheet.name, 'row', 'delete', 2, 1));
    expect(str(workbook.findNamedRange('Prices')!.target)).toBe('Sheet1.B2:B4');
    expect(workbook.findNamedRange('Gone')).toBeUndefined();
    expect(str(sheet.validations[0]!.target)).toBe('C2:C4');
    history.undo();
    expect(str(workbook.findNamedRange('Gone')!.target)).toBe('Sheet1.A3');
  });

  it('deleting a filtered column drops its criteria', () => {
    const { sheet, chart, history } = book();
    history.execute(new StructuralEditCommand(sheet.name, 'column', 'delete', 1, 3));
    expect(str(sheet.activeFilter)).toBe('A1:A5');
    expect(sheet.activeFilterSpec!.columns.size).toBe(0);
    expect(sheet.activeSort!.levels).toHaveLength(0);
    // The chart's only series lived in the deleted columns.
    expect(chart.chart!.series).toHaveLength(0);
  });

  it('a whole-column reference stays whole after rows are deleted', () => {
    const { sheet, history } = book();
    sheet.setFormula(9, 9, 'SUM(A:A)+SUM(2:2)');
    history.execute(new StructuralEditCommand(sheet.name, 'row', 'delete', 0, 3));
    expect(sheet.find(6, 9)!.formula).toBe('SUM(A:A)+SUM(#REF!)');
  });
});

describe('cut and paste to another sheet', () => {
  it('re-homes references inside the block and pins the rest to the source sheet', () => {
    const workbook = Workbook.createNew();
    const s1 = workbook.activeSheet;
    const s2 = workbook.addSheet('Sheet2');
    s1.setValue(0, 0, CellValue.number(7));
    s1.setFormula(1, 0, 'A1*2');
    s1.setFormula(2, 0, 'B5+1');
    s1.setValue(4, 1, CellValue.number(100));
    const history = new EditHistory(workbook);
    history.execute(new PasteBlockCommand('Sheet2', [0, 2], ClipboardBlock.capture(s1, R('A1:A3'), true)));
    expect(s2.find(1, 2)!.formula).toBe('C1*2');
    expect(s2.find(2, 2)!.formula).toBe('Sheet1.B5+1');
    recalc(workbook);
    expect(CellValue.asNumber(s2.getValue(1, 2))).toBe(14);
    expect(CellValue.asNumber(s2.getValue(2, 2))).toBe(101);

    history.undo();
    expect(s1.find(1, 0)!.formula).toBe('A1*2');
    expect(s1.find(2, 0)!.formula).toBe('B5+1');
    expect(s2.find(1, 2)).toBeUndefined();
  });
});

describe('DeleteSheetCommand fixes references outside formulas', () => {
  function book() {
    const workbook = Workbook.createNew();
    const history = new EditHistory(workbook);
    const main = workbook.activeSheet;
    workbook.addSheet('Data');
    workbook.addNamedRange({ name: 'Src', target: R('Data.A1:A3') });
    workbook.addNamedRange({ name: 'Local', target: R('Sheet1.A1'), scope: 'Data' });
    workbook.addNamedRange({ name: 'Kept', target: R('Sheet1.B1:B2') });
    main.addValidation(dataValidation('list', R('B1:B5'), { listSource: R('Data.A1:A3') }));
    const chartSheet = workbook.addChartSheet(
      createChart({
        categoriesReference: 'Data.A1:A3',
        series: [
          createSeries({ valuesReference: 'Data.$B$1:.$B$3' }),
          createSeries({ valuesReference: 'Sheet1.C1:C3' }),
        ],
      }),
    );
    return { workbook, history, main, chartSheet };
  }

  it('names on (or scoped to) the sheet go, list sources break, and chart series on it go', () => {
    const { workbook, history, main, chartSheet } = book();
    history.execute(new DeleteSheetCommand('Data'));

    expect(workbook.namedRanges.map((n) => n.name)).toEqual(['Kept']);
    const rule = main.validations[0]!;
    expect(str(rule.target)).toBe('B1:B5');
    expect(rule.listSource ?? null).toBeNull(); // #REF!: the list accepts nothing
    expect(chartSheet.chart!.series.map((s) => s.valuesReference)).toEqual(['Sheet1.C1:C3']);
    expect(chartSheet.chart!.categoriesReference).toBeNull();

    // A new sheet of the same name doesn't pick the old references back up.
    workbook.addSheet('Data');
    expect(workbook.namedRanges.some((n) => n.name === 'Src')).toBe(false);
  });

  it('undo puts every reference back', () => {
    const { workbook, history, main, chartSheet } = book();
    history.execute(new DeleteSheetCommand('Data'));
    history.undo();

    expect(workbook.namedRanges.map((n) => n.name).sort()).toEqual(['Kept', 'Local', 'Src']);
    expect(str(main.validations[0]!.listSource)).toBe('Data.A1:A3');
    expect(chartSheet.chart!.series).toHaveLength(2);
    expect(chartSheet.chart!.categoriesReference).toBe('Data.A1:A3');
  });
});
