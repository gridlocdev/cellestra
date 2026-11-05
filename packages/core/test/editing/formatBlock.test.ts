import { describe, expect, it } from 'vitest';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { FormatBlock } from '../../src/editing/formatBlock.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { MergedRegion } from '../../src/model/mergedRegion.ts';
import { Workbook } from '../../src/model/workbook.ts';
import type { Worksheet } from '../../src/model/worksheet.ts';
import { CellStyles, Fills, Fonts } from '../../src/styling/cellStyle.ts';
import { Argb } from '../../src/styling/color.ts';
import { StylePool } from '../../src/styling/stylePool.ts';
import { CellValue } from '../../src/values/cellValue.ts';

const R = (a1: string): CellRange => CellRange.parse(a1);

function setup() {
  const workbook = Workbook.createNew();
  return { workbook, history: new EditHistory(workbook), sheet: workbook.activeSheet };
}

const style = (sheet: Worksheet, row: number, column: number): number =>
  sheet.getOrCreate(row, column).styleId;
const yellow = Argb.fromHex('#FFFF00');

function internBoldYellow(workbook: Workbook): number {
  return workbook.styles.intern({
    ...CellStyles.default,
    font: { ...Fonts.default, bold: true },
    fill: Fills.solid(yellow),
  });
}

const internBold = (workbook: Workbook): number =>
  workbook.styles.intern({ ...CellStyles.default, font: { ...Fonts.default, bold: true } });

describe('FormatBlock', () => {
  it('painting copies only formatting, not values', () => {
    const { workbook, history, sheet } = setup();
    const sourceStyle = internBoldYellow(workbook);
    sheet.setValue(0, 0, CellValue.text('source'));
    sheet.getOrCreate(0, 0).styleId = sourceStyle;
    sheet.setValue(1, 1, CellValue.number(42));

    const stamp = FormatBlock.capture(workbook, sheet, R('A1'));
    history.execute(stamp.applyTo(sheet.name, R('B2')));

    const painted = workbook.styles.get(style(sheet, 1, 1));
    expect(painted.font.bold).toBe(true);
    expect(painted.fill.background).toBe(yellow);
    expect(CellValue.asNumber(sheet.getValue(1, 1))).toBe(42);
    expect(sheet.getValue(0, 0)).toEqual(CellValue.text('source'));
  });

  it('painting tiles a multi-cell pattern over the target', () => {
    const { workbook, history, sheet } = setup();
    sheet.getOrCreate(0, 0).styleId = internBold(workbook);

    const stamp = FormatBlock.capture(workbook, sheet, R('A1:B1'));
    history.execute(stamp.applyTo(sheet.name, R('A3:D3')));

    expect(workbook.styles.get(style(sheet, 2, 0)).font.bold).toBe(true);
    expect(workbook.styles.get(style(sheet, 2, 1)).font.bold).toBe(false);
    expect(workbook.styles.get(style(sheet, 2, 2)).font.bold).toBe(true);
    expect(workbook.styles.get(style(sheet, 2, 3)).font.bold).toBe(false);
  });

  it('painting is undoable and restores styles', () => {
    const { workbook, history, sheet } = setup();
    sheet.getOrCreate(0, 0).styleId = internBold(workbook);

    const stamp = FormatBlock.capture(workbook, sheet, R('A1'));
    history.execute(stamp.applyTo(sheet.name, R('B2')));
    expect(workbook.styles.get(style(sheet, 1, 1)).font.bold).toBe(true);

    history.undo();
    expect(sheet.getOrCreate(1, 1).styleId).toBe(StylePool.defaultId);
  });

  it('painting replicates merge from the source', () => {
    const { workbook, history, sheet } = setup();
    sheet.addMerge(new MergedRegion(R('A1:B1')));

    const stamp = FormatBlock.capture(workbook, sheet, R('A1:B1'));
    history.execute(stamp.applyTo(sheet.name, R('A3:B3')));

    const merge = sheet.findMerge(2, 0);
    expect(merge).toBeTruthy();
    expect(merge!.range.equals(R('A3:B3'))).toBe(true);

    history.undo();
    expect(sheet.findMerge(2, 0)).toBeFalsy();
  });

  it('painting clears an existing target merge when source has one', () => {
    const { workbook, history, sheet } = setup();
    sheet.addMerge(new MergedRegion(R('A1:B1')));
    sheet.addMerge(new MergedRegion(R('A3:A4')));

    const stamp = FormatBlock.capture(workbook, sheet, R('A1:B1'));
    history.execute(stamp.applyTo(sheet.name, R('A3:B3')));

    const merge = sheet.findMerge(2, 0);
    expect(merge).toBeTruthy();
    expect(merge!.range.equals(R('A3:B3'))).toBe(true);

    history.undo();
    const restored = sheet.findMerge(2, 0);
    expect(restored).toBeTruthy();
    expect(restored!.range.equals(R('A3:A4'))).toBe(true);
  });

  // ---- Drag-to-paint range resolution ----

  it.each([
    [2, 1, 2, 1, 'B3', 'B3'],
    [2, 1, 4, 3, 'B3', 'D5'],
    [4, 3, 2, 1, 'B3', 'D5'],
  ])(
    'paintTarget(%i,%i,%i,%i) is the bounding box %s:%s',
    (anchorRow, anchorColumn, row, column, topLeft, bottomRight) => {
      const target = FormatBlock.paintTarget(anchorRow, anchorColumn, row, column);
      expect(target.equals(R(`${topLeft}:${bottomRight}`))).toBe(true);
    },
  );

  it('dragging paints the stamp over the resolved range, leaving values', () => {
    const { workbook, history, sheet } = setup();
    const sourceStyle = internBoldYellow(workbook);
    sheet.setValue(0, 0, CellValue.text('source'));
    sheet.getOrCreate(0, 0).styleId = sourceStyle;
    sheet.setValue(3, 2, CellValue.number(7));

    const stamp = FormatBlock.capture(workbook, sheet, R('A1'));
    const painted = FormatBlock.paintTarget(2, 1, 4, 3);
    history.execute(stamp.applyTo(sheet.name, painted));

    for (const [row, column] of painted.positions()) {
      const s = workbook.styles.get(style(sheet, row, column));
      expect(s.font.bold).toBe(true);
      expect(s.fill.background).toBe(yellow);
    }

    expect(CellValue.asNumber(sheet.getValue(3, 2))).toBe(7);
    expect(sheet.getValue(2, 0).kind).toBe('empty');
    expect(workbook.styles.get(style(sheet, 2, 0)).font.bold).toBe(false);
  });

  it('dragging replicates a source merge over the painted range', () => {
    const { workbook, history, sheet } = setup();
    sheet.addMerge(new MergedRegion(R('A1:B1')));

    const stamp = FormatBlock.capture(workbook, sheet, R('A1:B1'));
    const painted = FormatBlock.paintTarget(2, 0, 2, 1);
    history.execute(stamp.applyTo(sheet.name, painted));

    const merge = sheet.findMerge(2, 0);
    expect(merge).toBeTruthy();
    expect(merge!.range.equals(R('A3:B3'))).toBe(true);
  });
});
