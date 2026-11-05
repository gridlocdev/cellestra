import { describe, expect, it } from 'vitest';
import type { EditCommand } from '../../src/editing/editCommand.ts';
import { EditCommands } from '../../src/editing/editCommands.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { CellAddress } from '../../src/model/cellAddress.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { SheetDimensions } from '../../src/model/sheetDimensions.ts';
import { Workbook } from '../../src/model/workbook.ts';
import type { Worksheet } from '../../src/model/worksheet.ts';
import { readOds, writeOds } from '../../src/serialization/index.ts';
import { CellStyles } from '../../src/styling/cellStyle.ts';
import { CellValue } from '../../src/values/cellValue.ts';

// Formatting a whole column is one default style on the column, not a style stamped onto each of its
// 1,048,576 cells — ODF's table:default-cell-style-name.

function setup() {
  const workbook = Workbook.createNew();
  return { workbook, history: new EditHistory(workbook), sheet: workbook.activeSheet };
}

const wholeColumn = (column: number): CellRange =>
  new CellRange(new CellAddress(column, 0), new CellAddress(column, SheetDimensions.maxRows - 1));

const wrap = (workbook: Workbook, sheet: Worksheet, range: CellRange): EditCommand =>
  EditCommands.applyStyle(
    workbook,
    sheet.name,
    range,
    (s) => ({ ...s, alignment: { ...s.alignment, wrapText: true } }),
    'Wrap Text',
  );

const bold = (workbook: Workbook, sheet: Worksheet, a1: string): EditCommand =>
  EditCommands.applyStyle(workbook, sheet.name, CellRange.parse(a1), (s) => ({
    ...s,
    font: { ...s.font, bold: true },
  }));

describe('Column default styles', () => {
  it('wrapping a whole column materializes no cells', () => {
    const { workbook, history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('hello'));

    history.execute(wrap(workbook, sheet, wholeColumn(0)));

    expect([...sheet.enumerateCells()]).toHaveLength(1);
    expect(sheet.getUsedRange()!.equals(CellRange.parse('A1'))).toBe(true);
    expect(workbook.styles.get(sheet.findColumn(0)!.styleId).alignment.wrapText).toBe(true);
  });

  it('a blank cell in the column reads as wrapped', () => {
    const { workbook, history, sheet } = setup();
    history.execute(wrap(workbook, sheet, wholeColumn(0)));

    expect(workbook.styles.get(sheet.effectiveStyleId(500, 0)).alignment.wrapText).toBe(true);
    expect(workbook.styles.get(sheet.effectiveStyleId(500, 1)).alignment.wrapText).toBe(false);
    expect(sheet.find(500, 0)).toBeFalsy();
  });

  it('a cell with a style of its own takes the transform directly', () => {
    const { workbook, history, sheet } = setup();
    sheet.setValue(3, 0, CellValue.text('bold already'));
    history.execute(bold(workbook, sheet, 'A4'));

    history.execute(wrap(workbook, sheet, wholeColumn(0)));

    const style = workbook.styles.get(sheet.effectiveStyleId(3, 0));
    expect(style.alignment.wrapText).toBe(true);
    expect(style.font.bold).toBe(true);
  });

  it('undo restores the column and the cells it touched', () => {
    const { workbook, history, sheet } = setup();
    sheet.setValue(3, 0, CellValue.text('x'));
    history.execute(bold(workbook, sheet, 'A4'));

    history.execute(wrap(workbook, sheet, wholeColumn(0)));
    history.undo();

    expect(workbook.styles.get(sheet.effectiveStyleId(500, 0)).alignment.wrapText).toBe(false);
    const style = workbook.styles.get(sheet.effectiveStyleId(3, 0));
    expect(style.alignment.wrapText).toBe(false);
    expect(style.font.bold).toBe(true);
  });

  it("a cell's own style wins over its column and row", () => {
    const { workbook, sheet } = setup();
    const d = CellStyles.default;
    const wrapped = workbook.styles.intern({ ...d, alignment: { ...d.alignment, wrapText: true } });
    const boldId = workbook.styles.intern({ ...d, font: { ...d.font, bold: true } });
    const italic = workbook.styles.intern({ ...d, font: { ...d.font, italic: true } });

    sheet.getOrCreateColumn(0).styleId = wrapped;
    expect(sheet.effectiveStyleId(0, 0)).toBe(wrapped);

    sheet.getOrCreateRow(0).styleId = boldId;
    expect(sheet.effectiveStyleId(0, 0)).toBe(boldId);

    sheet.setValue(0, 0, CellValue.text('v'));
    sheet.find(0, 0)!.styleId = italic;
    expect(sheet.effectiveStyleId(0, 0)).toBe(italic);
  });

  it("a column's default style round-trips through ods", () => {
    const { workbook, history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('hello'));
    history.execute(wrap(workbook, sheet, wholeColumn(0)));

    const reopened = readOds(writeOds(workbook));
    const reopenedSheet = reopened.activeSheet;

    expect(reopened.styles.get(reopenedSheet.effectiveStyleId(500, 0)).alignment.wrapText).toBe(true);
    expect(reopened.styles.get(reopenedSheet.effectiveStyleId(500, 1)).alignment.wrapText).toBe(false);
    expect([...reopenedSheet.enumerateCells()]).toHaveLength(1);
  });

  it('a formatted but empty column survives a save', () => {
    const { workbook, history, sheet } = setup();
    history.execute(wrap(workbook, sheet, wholeColumn(4)));

    const reopened = readOds(writeOds(workbook));

    expect(reopened.styles.get(reopened.activeSheet.effectiveStyleId(0, 4)).alignment.wrapText).toBe(true);
  });
});
