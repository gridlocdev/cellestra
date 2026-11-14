import { describe, expect, it } from 'vitest';
import { ClipboardBlock } from '../../src/editing/clipboardBlock.ts';
import { EditCommands } from '../../src/editing/editCommands.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { PasteBlockCommand } from '../../src/editing/pasteBlockCommand.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { SheetDimensions } from '../../src/model/sheetDimensions.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { Office, Text } from '../../src/serialization/ods/odsXmlNames.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { XmlElement } from '../../src/xml/index.ts';

const R = (a1: string): CellRange => CellRange.parse(a1);
const num = (v: CellValue): number => CellValue.asNumber(v);

function setup() {
  const workbook = Workbook.createNew();
  return { workbook, history: new EditHistory(workbook), sheet: workbook.activeSheet };
}

describe('PasteBlockCommand', () => {
  it('copy-paste preserves formula and style and shifts references', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(10));
    sheet.setFormula(0, 1, 'A1*2', CellValue.number(20));

    const block = ClipboardBlock.capture(sheet, R('B1'), false);
    history.execute(new PasteBlockCommand(sheet.name, [1, 1], block));

    expect(sheet.find(1, 1)!.formula).toBe('A2*2');
  });

  it('cut-paste moves cells and keeps formula unchanged', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(10));
    sheet.setFormula(0, 1, 'A1*2', CellValue.number(20));

    const block = ClipboardBlock.capture(sheet, R('B1'), true);
    history.execute(new PasteBlockCommand(sheet.name, [5, 5], block));

    expect(sheet.find(0, 1)).toBeFalsy();
    expect(sheet.find(5, 5)!.formula).toBe('A1*2');

    history.undo();
    expect(sheet.find(0, 1)!.formula).toBe('A1*2');
    expect(sheet.find(5, 5)).toBeFalsy();
  });

  it('cut-paste makes external references follow the moved cell', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(10));
    sheet.setFormula(0, 2, 'A1*2', CellValue.number(20));

    const block = ClipboardBlock.capture(sheet, R('A1'), true);
    history.execute(new PasteBlockCommand(sheet.name, [4, 0], block));

    expect(sheet.find(0, 2)!.formula).toBe('A5*2');

    history.undo();
    expect(sheet.find(0, 2)!.formula).toBe('A1*2');
  });

  it('cut-paste follows even absolute references', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(10));
    sheet.setFormula(0, 2, '$A$1*2', CellValue.number(20));

    const block = ClipboardBlock.capture(sheet, R('A1'), true);
    history.execute(new PasteBlockCommand(sheet.name, [4, 0], block));

    expect(sheet.find(0, 2)!.formula).toBe('$A$5*2');
  });

  it('cut-paste adjusts references between cells that move together', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 1, CellValue.number(3));
    sheet.setFormula(1, 1, 'B1+1', CellValue.number(4));

    const block = ClipboardBlock.capture(sheet, R('B1:B2'), true);
    history.execute(new PasteBlockCommand(sheet.name, [0, 3], block));

    expect(sheet.find(1, 3)!.formula).toBe('D1+1');
  });

  it('cut-paste leaves references outside the moved cell untouched', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(10));
    sheet.setFormula(0, 1, 'A1*2', CellValue.number(20));

    const block = ClipboardBlock.capture(sheet, R('B1'), true);
    history.execute(new PasteBlockCommand(sheet.name, [5, 5], block));

    expect(sheet.find(5, 5)!.formula).toBe('A1*2');
  });

  it('single cell fills a multi-cell selection', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(7));

    const block = ClipboardBlock.capture(sheet, R('A1'), false);
    history.execute(new PasteBlockCommand(sheet.name, R('C1:D2'), block));

    expect(num(sheet.getValue(0, 2))).toBe(7);
    expect(num(sheet.getValue(0, 3))).toBe(7);
    expect(num(sheet.getValue(1, 2))).toBe(7);
    expect(num(sheet.getValue(1, 3))).toBe(7);
  });

  it('block tiles across an exact-multiple selection', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(1));
    sheet.setValue(0, 1, CellValue.number(2));

    const block = ClipboardBlock.capture(sheet, R('A1:B1'), false);
    history.execute(new PasteBlockCommand(sheet.name, R('A3:D3'), block));

    expect(num(sheet.getValue(2, 0))).toBe(1);
    expect(num(sheet.getValue(2, 1))).toBe(2);
    expect(num(sheet.getValue(2, 2))).toBe(1);
    expect(num(sheet.getValue(2, 3))).toBe(2);
  });

  it('mismatched shapes paste once and flag', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(1));
    sheet.setValue(0, 1, CellValue.number(2));

    const block = ClipboardBlock.capture(sheet, R('A1:B1'), false);
    const command = new PasteBlockCommand(sheet.name, R('A3:C3'), block);
    history.execute(command);

    expect(command.shapeMismatch).toBe(true);
    expect(num(sheet.getValue(2, 0))).toBe(1);
    expect(num(sheet.getValue(2, 1))).toBe(2);
    expect(sheet.getValue(2, 2).kind).toBe('empty');
  });

  it('paste carries the style index', () => {
    const { workbook, history, sheet } = setup();
    history.execute(
      EditCommands.applyStyle(workbook, sheet.name, R('A1'), (s) => ({
        ...s,
        font: { ...s.font, bold: true },
      })),
    );
    const styledId = sheet.find(0, 0)!.styleId;

    const block = ClipboardBlock.capture(sheet, R('A1'), false);
    history.execute(new PasteBlockCommand(sheet.name, [2, 2], block));

    expect(sheet.find(2, 2)!.styleId).toBe(styledId);
  });

  // ---- Paste Special modes ---------------------------------------------

  it('paste values writes value without formula and keeps destination style', () => {
    const { workbook, history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(10));
    sheet.setFormula(0, 1, 'A1*2', CellValue.number(20));

    history.execute(
      EditCommands.applyStyle(workbook, sheet.name, R('B2'), (s) => ({
        ...s,
        font: { ...s.font, italic: true },
      })),
    );
    const destStyle = sheet.find(1, 1)!.styleId;

    const block = ClipboardBlock.capture(sheet, R('B1'), false);
    history.execute(new PasteBlockCommand(sheet.name, [1, 1], block, 'values'));

    const dest = sheet.find(1, 1)!;
    expect(dest.formula).toBeNull();
    expect(num(dest.value)).toBe(20);
    expect(dest.styleId).toBe(destStyle);
  });

  it('paste formulas writes shifted formula and keeps destination style', () => {
    const { workbook, history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(10));
    sheet.setFormula(0, 1, 'A1*2', CellValue.number(20));

    history.execute(
      EditCommands.applyStyle(workbook, sheet.name, R('B2'), (s) => ({
        ...s,
        font: { ...s.font, italic: true },
      })),
    );
    const destStyle = sheet.find(1, 1)!.styleId;

    const block = ClipboardBlock.capture(sheet, R('B1'), false);
    history.execute(new PasteBlockCommand(sheet.name, [1, 1], block, 'formulas'));

    const dest = sheet.find(1, 1)!;
    expect(dest.formula).toBe('A2*2');
    expect(dest.styleId).toBe(destStyle);
  });

  it('paste formats applies source style and leaves destination value', () => {
    const { workbook, history, sheet } = setup();
    history.execute(
      EditCommands.applyStyle(workbook, sheet.name, R('A1'), (s) => ({
        ...s,
        font: { ...s.font, bold: true },
      })),
    );
    const sourceStyle = sheet.find(0, 0)!.styleId;
    sheet.setValue(2, 2, CellValue.number(99));

    const block = ClipboardBlock.capture(sheet, R('A1'), false);
    history.execute(new PasteBlockCommand(sheet.name, [2, 2], block, 'formats'));

    const dest = sheet.find(2, 2)!;
    expect(num(dest.value)).toBe(99);
    expect(dest.styleId).toBe(sourceStyle);
  });

  it('paste special is undoable', () => {
    const { history, sheet } = setup();
    sheet.setFormula(0, 1, '1+1', CellValue.number(2));
    sheet.setValue(1, 1, CellValue.number(5));

    const block = ClipboardBlock.capture(sheet, R('B1'), false);
    history.execute(new PasteBlockCommand(sheet.name, [1, 1], block, 'values'));
    expect(num(sheet.getValue(1, 1))).toBe(2);

    history.undo();
    expect(num(sheet.getValue(1, 1))).toBe(5);
  });

  it('refuses a paste that would run past the last column or row, writing nothing', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(1));
    sheet.setValue(0, 1, CellValue.number(2));
    const block = ClipboardBlock.capture(sheet, R('A1:B1'), false);
    const lastColumn = SheetDimensions.maxColumns - 1;

    expect(() => history.execute(new PasteBlockCommand(sheet.name, [3, lastColumn], block))).toThrow(
      /doesn't fit/,
    );
    expect(sheet.find(3, lastColumn)).toBeFalsy();
    expect(sheet.find(4, 0)).toBeFalsy();
    expect(history.canUndo).toBe(false);

    const tall = ClipboardBlock.capture(sheet, R('A1:A2'), false);
    expect(() =>
      history.execute(new PasteBlockCommand(sheet.name, [SheetDimensions.maxRows - 1, 0], tall)),
    ).toThrow(/doesn't fit/);
  });

  describe('a comment and rich text travel with the cell', () => {
    function commented() {
      const { workbook, history, sheet } = setup();
      sheet.setValue(0, 0, CellValue.text('Hello'));
      const cell = sheet.find(0, 0)!;
      const comment = new XmlElement(Office('annotation'), new XmlElement(Text('p'), 'Check this'));
      cell.unmodeled = { before: [comment], after: [] };
      cell.richText = { text: 'Hello', hyperlink: null, paragraphs: [new XmlElement(Text('p'), 'Hello')] };
      return { workbook, history, sheet, unmodeled: cell.unmodeled, richText: cell.richText };
    }

    it('on a cut and paste', () => {
      const { history, sheet, unmodeled, richText } = commented();
      history.execute(
        new PasteBlockCommand(sheet.name, [3, 3], ClipboardBlock.capture(sheet, R('A1'), true)),
      );
      expect(sheet.find(0, 0)).toBeFalsy();
      expect(sheet.find(3, 3)!.unmodeled).toBe(unmodeled);
      expect(sheet.find(3, 3)!.richText).toBe(richText);

      history.undo();
      expect(sheet.find(0, 0)!.unmodeled).toBe(unmodeled);
      expect(sheet.find(3, 3)).toBeFalsy();
    });

    it('on a copy and paste', () => {
      const { history, sheet, unmodeled, richText } = commented();
      history.execute(
        new PasteBlockCommand(sheet.name, [3, 3], ClipboardBlock.capture(sheet, R('A1'), false)),
      );
      expect(sheet.find(0, 0)!.unmodeled).toBe(unmodeled);
      expect(sheet.find(3, 3)!.unmodeled).toBe(unmodeled);
      expect(sheet.find(3, 3)!.richText).toBe(richText);
    });

    it('not on a paste of values only', () => {
      const { history, sheet } = commented();
      const block = ClipboardBlock.capture(sheet, R('A1'), false);
      history.execute(new PasteBlockCommand(sheet.name, [3, 3], block, 'values'));
      expect(sheet.find(3, 3)!.value).toEqual(CellValue.text('Hello'));
      expect(sheet.find(3, 3)!.unmodeled).toBeNull();
      expect(sheet.find(3, 3)!.richText).toBeNull();
    });
  });
});
