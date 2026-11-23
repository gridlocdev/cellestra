import { describe, expect, it } from 'vitest';
import { CellRef } from '../../src/dependencies/cellRef.ts';
import { RecalculationEngine } from '../../src/dependencies/recalculationEngine.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { RenameSheetCommand } from '../../src/editing/renameSheetCommand.ts';
import { StructuralEditCommand } from '../../src/editing/structuralEditCommand.ts';
import { readNameDefinition } from '../../src/formulas/evaluation/namedExpressions.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { readOds, writeOds } from '../../src/serialization/index.ts';
import { CellValue } from '../../src/values/cellValue.ts';

/** A workbook with `formulas` in column B of Sheet1, calculated. */
function calculated(workbook: Workbook): RecalculationEngine {
  const engine = new RecalculationEngine(workbook);
  engine.rebuildGraph();
  engine.recalculateAll();
  return engine;
}

const shown = (workbook: Workbook, a1: string, sheet = workbook.activeSheet): CellValue => {
  const at = CellRange.parse(a1);
  return sheet.getValue(at.top, at.left);
};

describe('named expressions (a name that holds a value or formula)', () => {
  it('a value name works in formulas, as TaxRate = 0.2', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    workbook.addNamedRange({ name: 'TaxRate', expression: '0.2' });
    sheet.setValue(0, 0, CellValue.number(150));
    sheet.setFormula(0, 1, 'A1*TaxRate');
    sheet.setFormula(1, 1, 'taxrate+1');
    calculated(workbook);
    expect(shown(workbook, 'B1')).toEqual(CellValue.number(30));
    expect(shown(workbook, 'B2')).toEqual(CellValue.number(1.2));
  });

  it('an expression that reads cells recalculates when they change', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    workbook.addNamedRange({ name: 'Doubled', expression: 'Sheet1.$A$1*2' });
    sheet.setValue(0, 0, CellValue.number(5));
    sheet.setFormula(0, 1, 'Doubled+1');
    const engine = calculated(workbook);
    expect(shown(workbook, 'B1')).toEqual(CellValue.number(11));
    sheet.setValue(0, 0, CellValue.number(7));
    engine.recordEdits([CellRef.of(sheet.name, 0, 0)]);
    expect(shown(workbook, 'B1')).toEqual(CellValue.number(15));
  });

  it('a name can use another name, and a range expression works where a range is needed', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    workbook.addNamedRange({ name: 'Rate', expression: '0.5' });
    workbook.addNamedRange({ name: 'Half', expression: 'Rate*10' });
    workbook.addNamedRange({ name: 'Both', expression: '$A$1:$A$2~$A$4' });
    for (const [r, v] of [
      [0, 1],
      [1, 2],
      [3, 4],
    ] as const)
      sheet.setValue(r, 0, CellValue.number(v));
    sheet.setFormula(0, 1, 'Half');
    sheet.setFormula(1, 1, 'SUM(Both)');
    calculated(workbook);
    expect(shown(workbook, 'B1')).toEqual(CellValue.number(5));
    expect(shown(workbook, 'B2')).toEqual(CellValue.number(7));
  });

  it('a name that uses itself is an error, not a hang', () => {
    const workbook = Workbook.createNew();
    workbook.addNamedRange({ name: 'Loop', expression: 'Other+1' });
    workbook.addNamedRange({ name: 'Other', expression: 'Loop*2' });
    workbook.activeSheet.setFormula(0, 0, 'Loop');
    calculated(workbook);
    expect(shown(workbook, 'A1').kind).toBe('error');
  });

  it('a sheet-local expression shadows the global one on its sheet', () => {
    const workbook = Workbook.createNew();
    const other = workbook.addSheet('Other');
    workbook.addNamedRange({ name: 'Rate', expression: '1' });
    workbook.addNamedRange({ name: 'Rate', expression: '2', scope: 'Other' });
    workbook.activeSheet.setFormula(0, 0, 'Rate');
    other.setFormula(0, 0, 'Rate');
    calculated(workbook);
    expect(shown(workbook, 'A1')).toEqual(CellValue.number(1));
    expect(shown(workbook, 'A1', other)).toEqual(CellValue.number(2));
  });

  it('survives a save and reopen as a table:named-expression, and still calculates', () => {
    const workbook = Workbook.createNew();
    workbook.addNamedRange({ name: 'TaxRate', expression: '0.2' });
    workbook.addNamedRange({ name: 'Local', expression: 'Sheet1.$A$1*3', scope: 'Sheet1' });
    workbook.activeSheet.setValue(0, 0, CellValue.number(10));
    workbook.activeSheet.setFormula(0, 1, 'A1*TaxRate+Local');
    const bytes = writeOds(workbook);
    const reopened = readOds(bytes);
    expect(reopened.findNamedRange('TaxRate')?.expression).toBe('0.2');
    expect(reopened.findNamedRange('Local', 'Sheet1')?.scope).toBe('Sheet1');
    calculated(reopened);
    expect(shown(reopened, 'B1')).toEqual(CellValue.number(32));
  });

  it('references inside an expression follow inserted rows and a renamed sheet, and undo', () => {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    workbook.addNamedRange({ name: 'Base', expression: 'Sheet1.$A$2*2' });
    const history = new EditHistory(workbook);
    history.execute(new StructuralEditCommand(sheet.name, 'row', 'insert', 0, 3));
    expect(workbook.findNamedRange('Base')?.expression).toBe('Sheet1.$A$5*2');
    history.execute(new RenameSheetCommand('Sheet1', 'Data'));
    expect(workbook.findNamedRange('Base')?.expression).toBe('Data.$A$5*2');
    history.undo();
    history.undo();
    expect(workbook.findNamedRange('Base')?.expression).toBe('Sheet1.$A$2*2');
  });
});

describe('readNameDefinition', () => {
  it.each([
    ['B2:D9', { target: 'B2:D9' }],
    ['$A$1', { target: 'A1' }],
    ['=Sheet1.$B$1:$B$3', { target: 'Sheet1.B1:B3' }],
    ['0.2', { expression: '0.2' }],
    ['=0.2', { expression: '0.2' }],
    ['= A1 * 2', { expression: 'A1*2' }],
    ['="text"', { expression: '"text"' }],
  ])('%s', (typed, expected) => {
    const read = readNameDefinition(typed);
    expect(read && ('target' in read ? { target: read.target.toString() } : read)).toEqual(expected);
  });

  it.each(['', '=', '=1+', 'SUM('])('%j is neither', (typed) => {
    expect(readNameDefinition(typed)).toBeNull();
  });
});
