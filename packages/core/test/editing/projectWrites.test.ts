import { describe, expect, it } from 'vitest';
import { ClipboardBlock } from '../../src/editing/clipboardBlock.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { FillCommand } from '../../src/editing/fillCommand.ts';
import { PasteBlockCommand } from '../../src/editing/pasteBlockCommand.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import type { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';

// The validation pre-flight projections must match what the commands actually write, or paste/fill
// could be blocked (or allowed) for the wrong cells.

function setup() {
  const workbook = Workbook.createNew();
  return { workbook, history: new EditHistory(workbook), sheet: workbook.activeSheet };
}

interface Projected {
  readonly row: number;
  readonly column: number;
  readonly value: CellValue;
  readonly isFormula: boolean;
}

function expectApplied(sheet: Worksheet, projected: readonly Projected[]): void {
  for (const { row, column, value, isFormula } of projected) {
    const cell = sheet.find(row, column);
    expect(cell).toBeTruthy();
    expect(CellValue.equals(value, cell!.value)).toBe(true);
    expect(cell!.formula !== null).toBe(isFormula);
  }
}

describe('ProjectWrites', () => {
  it('paste projection matches applied cells', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(10));
    sheet.setFormula(0, 1, 'A1*2', CellValue.number(20));

    const block = ClipboardBlock.capture(sheet, CellRange.parse('A1:B1'), false);
    const command = new PasteBlockCommand(sheet.name, [2, 0], block);

    const projected = [...command.projectWrites()];
    history.execute(command);

    expect(projected.length).toBeGreaterThan(0);
    expectApplied(sheet, projected);
  });

  it('fill projection matches applied cells', () => {
    const { workbook, history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(7));

    const command = new FillCommand(sheet.name, CellRange.parse('A1'), CellRange.parse('A1:A3'));

    const projected = [...command.projectWrites(workbook)];
    history.execute(command);

    expect(projected).toHaveLength(2);
    expectApplied(sheet, projected);
  });
});
