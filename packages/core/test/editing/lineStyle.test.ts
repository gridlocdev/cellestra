import { describe, expect, it } from 'vitest';
import { DataRegion } from '../../src/data/dataRegion.ts';
import { EditCommands } from '../../src/editing/editCommands.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { FormatBlock } from '../../src/editing/formatBlock.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import type { CellStyle } from '../../src/styling/cellStyle.ts';
import { CellValue } from '../../src/values/cellValue.ts';

// Whole rows and whole columns are styled through their line models (cell → row → column), never
// with a cell per position, and every position ends up looking like the change of what it was.

const R = (a1: string): CellRange => CellRange.parse(a1);
const bold = (s: CellStyle): CellStyle => ({ ...s, font: { ...s.font, bold: true } });

function setup() {
  const workbook = Workbook.createNew();
  const sheet = workbook.activeSheet;
  const history = new EditHistory(workbook);
  const red = workbook.styles.intern({ ...workbook.styles.get(0), fill: { background: 0xffff0000 } });
  const look = (a1: string): CellStyle => {
    const a = R(a1);
    return workbook.styles.get(sheet.effectiveStyleId(a.top, a.left));
  };
  return { workbook, sheet, history, red, look };
}

describe('whole-row styling', () => {
  it('styles one whole row through its row model, with no cell per column', () => {
    const { workbook, sheet, history, look } = setup();
    sheet.setValue(0, 0, CellValue.text('x'));
    history.execute(EditCommands.applyStyle(workbook, sheet.name, R('A1:XFD1'), bold));
    expect(sheet.cellCount).toBe(1);
    expect(look('Z1').font.bold).toBe(true);
    expect(look('A1').font.bold).toBe(true);
    expect(look('A2').font.bold).toBeFalsy();
  });

  it('styles every row of a tall whole-row selection, not just the used range', () => {
    const { workbook, sheet, history, look } = setup();
    sheet.setValue(0, 0, CellValue.text('x'));
    history.execute(EditCommands.applyStyle(workbook, sheet.name, R('A1:XFD100'), bold));
    expect(sheet.cellCount).toBe(1);
    expect(look('B50').font.bold).toBe(true);
    expect(look('B100').font.bold).toBe(true);
    expect(look('B101').font.bold).toBeFalsy();
  });

  it('keeps a styled column showing through where the new row style crosses it', () => {
    const { workbook, sheet, history, red, look } = setup();
    sheet.getOrCreateColumn(2).styleId = red; // C
    history.execute(EditCommands.applyStyle(workbook, sheet.name, R('A2:XFD2'), bold));
    expect(look('C2').font.bold).toBe(true);
    expect(look('C2').fill.background).toBe(0xffff0000);
    expect(look('D2').fill.background).not.toBe(0xffff0000);
    expect(look('C3').font.bold).toBeFalsy();
  });

  it('undo restores the rows, and the crossing cells it made', () => {
    const { workbook, sheet, history, red, look } = setup();
    sheet.getOrCreateColumn(2).styleId = red;
    history.execute(EditCommands.applyStyle(workbook, sheet.name, R('A2:XFD3'), bold));
    history.undo();
    expect(look('B2').font.bold).toBeFalsy();
    expect(look('C2').font.bold).toBeFalsy();
    expect(look('C2').fill.background).toBe(0xffff0000);
    expect(sheet.cellCount).toBe(0);
  });
});

describe('whole-column styling over styled rows', () => {
  it('a column style shows on rows that carry a row style', () => {
    const { workbook, sheet, history, red, look } = setup();
    sheet.getOrCreateRow(2).styleId = red; // row 3
    history.execute(EditCommands.applyStyle(workbook, sheet.name, R('A1:A1048576'), bold));
    expect(look('A2').font.bold).toBe(true);
    expect(look('A3').font.bold).toBe(true);
    expect(look('A3').fill.background).toBe(0xffff0000); // the row's fill stays
    expect(look('B3').font.bold).toBeFalsy();
    expect(sheet.cellCount).toBe(1); // only the crossing

    history.undo();
    expect(look('A3').font.bold).toBeFalsy();
    expect(sheet.cellCount).toBe(0);
  });

  it('the whole sheet restyles columns and styled rows, with no crossing cells', () => {
    const { workbook, sheet, history, red, look } = setup();
    sheet.getOrCreateRow(2).styleId = red;
    history.execute(EditCommands.applyStyle(workbook, sheet.name, R('A1:XFD1048576'), bold));
    expect(sheet.cellCount).toBe(0);
    expect(look('Q3').font.bold).toBe(true);
    expect(look('Q3').fill.background).toBe(0xffff0000);
    expect(look('Q9').font.bold).toBe(true);
  });
});

describe('format painter onto whole lines', () => {
  it('paints a whole column as a column style', () => {
    const { workbook, sheet, history, look } = setup();
    sheet.getOrCreate(0, 0).styleId = workbook.styles.intern(bold(workbook.styles.get(0)));
    const block = FormatBlock.capture(workbook, sheet, R('A1:A1'));
    history.execute(block.applyTo(sheet.name, R('C1:C1048576')));
    expect(sheet.cellCount).toBe(1);
    expect(look('C500000').font.bold).toBe(true);
    history.undo();
    expect(look('C500000').font.bold).toBeFalsy();
  });

  it('paints an oversized block only where the sheet is used, keeping the tiling', () => {
    const { workbook, sheet, history, look } = setup();
    sheet.getOrCreate(0, 0).styleId = workbook.styles.intern(bold(workbook.styles.get(0)));
    sheet.setValue(3, 3, CellValue.number(1)); // D4
    const block = FormatBlock.capture(workbook, sheet, R('A1:A2')); // bold, plain
    history.execute(block.applyTo(sheet.name, R('B1:D1048576')));
    expect(look('D3').font.bold).toBe(true); // row 3 is the first row of a tile
    expect(look('D4').font.bold).toBeFalsy();
    expect(look('D5').font.bold).toBeFalsy(); // past the used range: untouched
    expect(sheet.cellCount).toBeLessThan(20);
  });
});

describe('current region and formatting', () => {
  it('ignores blank cells that only carry a style', () => {
    const { workbook, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('x'));
    sheet.setValue(1, 0, CellValue.number(1));
    for (let c = 1; c < 5; c++)
      sheet.getOrCreate(0, c).styleId = workbook.styles.intern(bold(workbook.styles.get(0)));
    expect(DataRegion.currentRegion(sheet, 0, 0).toString()).toBe('A1:A2');
  });
});
