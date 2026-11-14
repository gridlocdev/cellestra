import { describe, expect, it } from 'vitest';
import type { CellSnapshotCommand } from '../../src/editing/cellSnapshotCommand.ts';
import type { EditCommand } from '../../src/editing/editCommand.ts';
import { EditCommands } from '../../src/editing/editCommands.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { CellAddress } from '../../src/model/cellAddress.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { SheetDimensions } from '../../src/model/sheetDimensions.ts';
import { Workbook } from '../../src/model/workbook.ts';
import type { CellStyle } from '../../src/styling/cellStyle.ts';
import { StylePool } from '../../src/styling/stylePool.ts';
import { CellValue } from '../../src/values/cellValue.ts';

function setup() {
  const workbook = Workbook.createNew();
  return { workbook, history: new EditHistory(workbook), sheet: workbook.activeSheet };
}

const bold = (s: CellStyle): CellStyle => ({ ...s, font: { ...s.font, bold: true } });
const wholeSheet = (): CellRange =>
  new CellRange(
    new CellAddress(0, 0),
    new CellAddress(SheetDimensions.maxColumns - 1, SheetDimensions.maxRows - 1),
  );

describe('clear commands', () => {
  it('clear formatting resets style but keeps content', () => {
    const { workbook, history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(42));
    history.execute(EditCommands.applyStyle(workbook, sheet.name, CellRange.parse('A1'), bold));
    expect(sheet.find(0, 0)!.styleId).not.toBe(StylePool.defaultId);

    history.execute(EditCommands.clearFormatting(workbook, sheet.name, CellRange.parse('A1')));

    expect(sheet.find(0, 0)!.styleId).toBe(StylePool.defaultId); // formatting gone
    expect(CellValue.asNumber(sheet.getValue(0, 0))).toBe(42); // content stays

    history.undo();
    expect(workbook.styles.get(sheet.find(0, 0)!.styleId).font.bold).toBe(true); // restored
  });

  it('clear all drops content, formula, style and hyperlink', () => {
    const { workbook, history, sheet } = setup();
    sheet.setFormula(0, 0, '1+1', CellValue.number(2));
    sheet.find(0, 0)!.hyperlink = { target: 'https://example.com' };
    history.execute(EditCommands.applyStyle(workbook, sheet.name, CellRange.parse('A1'), bold));

    history.execute(EditCommands.clearAll(sheet.name, CellRange.parse('A1')));

    expect(sheet.find(0, 0)).toBeUndefined(); // cell pruned entirely

    history.undo();
    const restored = sheet.find(0, 0)!;
    expect(restored.formula).toBe('1+1');
    expect(restored.hyperlink).not.toBeNull();
    expect(workbook.styles.get(restored.styleId).font.bold).toBe(true);
  });

  it('clear all on the whole sheet is bounded to the used range and completes quickly', () => {
    // The whole-sheet selection is ≈17 billion cells; a clear that materialized every position would hang.
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(1));
    sheet.setValue(2, 3, CellValue.text('x'));

    const start = Date.now();
    history.execute(EditCommands.clearAll(sheet.name, wholeSheet()));
    expect(Date.now() - start).toBeLessThan(2000);
    expect(sheet.find(0, 0)).toBeUndefined();
    expect(sheet.find(2, 3)).toBeUndefined();

    history.undo(); // bounded undo restores exactly what was cleared
    expect(CellValue.asNumber(sheet.getValue(0, 0))).toBe(1);
    expect(sheet.getValue(2, 3)).toEqual(CellValue.text('x'));
  });

  it('clear contents on the whole sheet keeps styles and is bounded', () => {
    const { workbook, history, sheet } = setup();
    sheet.setValue(5, 5, CellValue.number(7));
    history.execute(EditCommands.applyStyle(workbook, sheet.name, CellRange.parse('F6'), bold));

    const start = Date.now();
    history.execute(EditCommands.clearContents(sheet.name, wholeSheet()));
    expect(Date.now() - start).toBeLessThan(2000);
    expect(sheet.getValue(5, 5).kind).toBe('empty');
    expect(workbook.styles.get(sheet.find(5, 5)!.styleId).font.bold).toBe(true); // style kept
  });

  it('apply style on the whole sheet is recorded per column, not per cell', () => {
    // A whole-sheet selection spans every row, so it lands on the columns' default cell styles.
    const { workbook, history, sheet } = setup();
    sheet.setValue(1, 1, CellValue.number(9));

    const start = Date.now();
    history.execute(EditCommands.applyStyle(workbook, sheet.name, wholeSheet(), bold, 'Bold'));
    expect(Date.now() - start).toBeLessThan(2000);
    expect(workbook.styles.get(sheet.effectiveStyleId(1, 1)).font.bold).toBe(true);
    expect([...sheet.enumerateCells()]).toHaveLength(1); // no cell was materialized to carry the format

    history.undo();
    expect(workbook.styles.get(sheet.effectiveStyleId(1, 1)).font.bold).toBe(false);
  });

  describe('on a sparse sheet the undo snapshot holds only the stored cells', () => {
    // A1:A20000 plus one value in AX1: the used range is A1:AX20000, a million positions.
    function sparse() {
      const { workbook, history, sheet } = setup();
      for (let r = 0; r < 20_000; r++) sheet.setValue(r, 0, CellValue.number(r));
      sheet.setValue(0, 49, CellValue.text('far'));
      sheet.setFormula(5, 0, 'A1+1', CellValue.number(1));
      sheet.getOrCreate(7, 0).styleId = workbook.styles.intern(bold(workbook.styles.get(0)));
      return { workbook, history, sheet };
    }
    const snapshotSize = (command: EditCommand): number =>
      (command as CellSnapshotCommand)['snapshot']!.length;

    it('clear contents', () => {
      const { history, sheet } = sparse();
      const styleId = sheet.find(7, 0)!.styleId;
      const command = EditCommands.clearContents(sheet.name, wholeSheet());
      const result = history.execute(command);
      expect(snapshotSize(command)).toBe(20_001);
      expect(result.changedCells).toHaveLength(20_001);
      expect(sheet.getValue(3, 0).kind).toBe('empty');
      expect(sheet.find(7, 0)!.styleId).toBe(styleId);

      history.undo();
      expect(sheet.cellCount).toBe(20_001);
      expect(sheet.getValue(3, 0)).toEqual(CellValue.number(3));
      expect(sheet.getValue(0, 49)).toEqual(CellValue.text('far'));
      expect(sheet.find(5, 0)!.formula).toBe('A1+1');
      expect(sheet.find(7, 0)!.styleId).toBe(styleId);
    });

    it('clear all', () => {
      const { history, sheet } = sparse();
      const command = EditCommands.clearAll(sheet.name, CellRange.parse('A1:AX20000'));
      history.execute(command);
      expect(snapshotSize(command)).toBe(20_001);
      expect(sheet.cellCount).toBe(0);

      history.undo();
      expect(sheet.cellCount).toBe(20_001);
      expect(sheet.getValue(19_999, 0)).toEqual(CellValue.number(19_999));
      expect(sheet.find(7, 0)!.styleId).not.toBe(StylePool.defaultId);
    });

    it('insert checkboxes takes back the cells it filled', () => {
      const { workbook, history, sheet } = setup();
      sheet.setValue(1, 1, CellValue.number(5));
      history.execute(EditCommands.insertCheckboxes(workbook, sheet.name, CellRange.parse('A1:C3')));
      expect(sheet.getValue(0, 0)).toEqual(CellValue.boolean(false));
      history.undo();
      expect(sheet.cellCount).toBe(1);
      expect(sheet.getValue(1, 1)).toEqual(CellValue.number(5));
    });
  });
});
