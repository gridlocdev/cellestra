import { describe, expect, it } from 'vitest';
import { DeleteSheetCommand } from '../../src/editing/deleteSheetCommand.ts';
import { EditRefusedError } from '../../src/editing/editCommand.ts';
import { EditCommands } from '../../src/editing/editCommands.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { StructuralEditCommand } from '../../src/editing/structuralEditCommand.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import type { Hyperlink } from '../../src/model/hyperlink.ts';
import { MergedRegion } from '../../src/model/mergedRegion.ts';
import { SheetDimensions } from '../../src/model/sheetDimensions.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { NumberFormats, type CellStyle } from '../../src/styling/cellStyle.ts';
import { CellValue } from '../../src/values/cellValue.ts';

const INVARIANT = 'en-US';
const bold = (s: CellStyle): CellStyle => ({ ...s, font: { ...s.font, bold: true } });
const num = (v: CellValue): number => CellValue.asNumber(v);

function setup() {
  const workbook = Workbook.createNew();
  return { workbook, history: new EditHistory(workbook), sheet: workbook.activeSheet };
}

describe('editing', () => {
  it('set value is undoable', () => {
    const { history, sheet } = setup();
    history.execute(EditCommands.setValue(sheet.name, 0, 0, CellValue.number(5)));
    expect(num(sheet.getValue(0, 0))).toBe(5);
    history.undo();
    expect(sheet.getValue(0, 0).kind).toBe('empty');
    history.redo();
    expect(num(sheet.getValue(0, 0))).toBe(5);
  });

  it('setHyperlink sets and undoes', () => {
    const { workbook, history, sheet } = setup();
    const link: Hyperlink = { target: 'https://example.com', tooltip: 'Example' };
    history.execute(EditCommands.setHyperlink(workbook, sheet.name, 0, 0, link));
    expect(sheet.find(0, 0)!.hyperlink).toEqual(link);
    history.undo();
    expect(sheet.find(0, 0)?.hyperlink ?? null).toBeNull();
    history.redo();
    expect(sheet.find(0, 0)!.hyperlink!.target).toBe('https://example.com');
  });

  it('setHyperlink null clears an existing link and undo restores it', () => {
    const { workbook, history, sheet } = setup();
    // A cell with a value and a link, so clearing the link must keep the value.
    history.execute(EditCommands.setValue(sheet.name, 0, 0, CellValue.text('Anchor')));
    history.execute(EditCommands.setHyperlink(workbook, sheet.name, 0, 0, { target: 'https://a.test' }));
    history.execute(EditCommands.setHyperlink(workbook, sheet.name, 0, 0, null));
    expect(sheet.find(0, 0)!.hyperlink).toBeNull();
    expect(sheet.getValue(0, 0)).toEqual(CellValue.text('Anchor'));
    history.undo(); // restores the cleared link
    expect(sheet.find(0, 0)!.hyperlink!.target).toBe('https://a.test');
  });

  it('setHyperlink null on an otherwise blank cell prunes it', () => {
    const { workbook, history, sheet } = setup();
    history.execute(EditCommands.setHyperlink(workbook, sheet.name, 0, 0, { target: 'https://a.test' }));
    history.execute(EditCommands.setHyperlink(workbook, sheet.name, 0, 0, null));
    expect(sheet.find(0, 0)).toBeUndefined(); // no value, formula, style or link → pruned
  });

  it('setInput chooses formula or value', () => {
    const { history, sheet } = setup();
    history.execute(EditCommands.setInput(sheet.name, 0, 0, '=A2+1', INVARIANT));
    history.execute(EditCommands.setInput(sheet.name, 1, 0, '42', INVARIANT));
    expect(sheet.find(0, 0)!.formula).toBe('A2+1');
    expect(num(sheet.getValue(1, 0))).toBe(42);
  });

  it("rewriting a cell's text preserves its alignment", () => {
    // Find & Replace rewrites a matched cell's text through setInput; an explicit alignment must survive.
    const { workbook, history, sheet } = setup();
    history.execute(EditCommands.setInput(sheet.name, 0, 0, 'Total', INVARIANT));
    history.execute(
      EditCommands.applyStyle(workbook, sheet.name, CellRange.parse('A1'), (s) => ({
        ...s,
        alignment: { ...s.alignment, horizontal: 'center', vertical: 'middle' },
      })),
    );
    history.execute(EditCommands.setInput(sheet.name, 0, 0, 'Sum', INVARIANT));
    const cell = sheet.find(0, 0)!;
    expect(cell.value).toEqual(CellValue.text('Sum'));
    const alignment = workbook.styles.get(cell.styleId).alignment;
    expect(alignment.horizontal).toBe('center');
    expect(alignment.vertical).toBe('middle');
  });

  it('setInput stamps a date format on a typed date', () => {
    const { workbook, history, sheet } = setup();
    history.execute(EditCommands.setInput(sheet.name, 0, 0, '6/7/2026', 'en-US', workbook));
    expect(sheet.getValue(0, 0).kind).toBe('number'); // stored as a serial
    expect(NumberFormats.kind(workbook.styles.get(sheet.find(0, 0)!.styleId).numberFormat)).toBe('date');
    history.undo();
    expect(sheet.getValue(0, 0).kind).toBe('empty'); // value and the stamped style both revert
  });

  it('inserted row inherits formatting from the row above', () => {
    const { workbook, history, sheet } = setup();
    history.execute(EditCommands.applyStyle(workbook, sheet.name, CellRange.parse('A2'), bold));
    const styledId = sheet.find(1, 0)!.styleId;
    expect(styledId).not.toBe(0);

    history.execute(new StructuralEditCommand(sheet.name, 'row', 'insert', 2, 1));
    const inherited = sheet.find(2, 0);
    expect(inherited).toBeDefined();
    expect(inherited!.styleId).toBe(styledId); // formatting carried down
    expect(inherited!.value.kind).toBe('empty'); // but no value

    history.undo();
    expect(sheet.find(2, 0)).toBeUndefined(); // the inherited styled cell is removed on undo
  });

  it('deleting a referenced sheet breaks dependents and undoes', () => {
    const { workbook, history, sheet } = setup();
    workbook.addSheet('Data');
    history.execute(EditCommands.setFormula(sheet.name, 0, 0, 'Data.A1+1'));
    history.execute(new DeleteSheetCommand('Data'));
    expect(workbook.findSheet('Data') ?? null).toBeNull();
    expect(sheet.find(0, 0)!.formula).toBe('#REF!+1');
    history.undo();
    expect(workbook.findSheet('Data') ?? null).not.toBeNull();
    expect(sheet.find(0, 0)!.formula).toBe('Data.A1+1');
  });

  it('apply style toggles bold and undoes', () => {
    const { workbook, history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.text('x'));
    history.execute(EditCommands.applyStyle(workbook, sheet.name, CellRange.parse('A1'), bold, 'Bold'));
    expect(workbook.styles.get(sheet.find(0, 0)!.styleId).font.bold).toBe(true);
    history.undo();
    expect(workbook.styles.get(sheet.find(0, 0)!.styleId).font.bold).toBe(false);
  });

  it('history respects the cap', () => {
    const { history, sheet } = setup();
    for (let i = 0; i < EditHistory.maxDepth + 50; i++) {
      history.execute(EditCommands.setValue(sheet.name, 0, 0, CellValue.number(i)));
    }
    let undone = 0;
    while (history.canUndo) {
      history.undo();
      undone++;
    }
    expect(undone).toBe(EditHistory.maxDepth);
  });

  it('insert rows shifts cells and adjusts formulas', () => {
    const { history, sheet } = setup();
    sheet.setValue(0, 0, CellValue.number(1)); // A1
    sheet.setValue(4, 0, CellValue.number(5)); // A5
    sheet.setFormula(5, 0, 'A5*2', CellValue.empty); // A6 references A5

    history.execute(new StructuralEditCommand(sheet.name, 'row', 'insert', 2, 1));
    expect(num(sheet.getValue(5, 0))).toBe(5); // A5 -> A6
    expect(sheet.find(6, 0)!.formula).toBe('A6*2'); // formula adjusted

    history.undo();
    expect(num(sheet.getValue(4, 0))).toBe(5); // back to A5
    expect(sheet.find(5, 0)!.formula).toBe('A5*2');
  });

  it('insert that would push data off the sheet is refused and changes nothing', () => {
    const { history, sheet } = setup();
    const last = SheetDimensions.maxRows - 1;
    sheet.setValue(last - 1, 0, CellValue.number(111));
    sheet.setValue(last, 0, CellValue.number(222));

    expect(() => history.execute(new StructuralEditCommand(sheet.name, 'row', 'insert', 0, 1))).toThrow(
      EditRefusedError,
    );
    expect(num(sheet.getValue(last - 1, 0))).toBe(111);
    expect(num(sheet.getValue(last, 0))).toBe(222);
    expect(history.canUndo).toBe(false);
  });

  it('insert pushing only formatting off the sheet drops it without overwriting survivors', () => {
    const { workbook, history, sheet } = setup();
    const last = SheetDimensions.maxRows - 1;
    sheet.setValue(last - 1, 0, CellValue.number(111));
    sheet.getOrCreate(last, 0).styleId = workbook.styles.intern({
      ...workbook.styles.get(0),
      font: { ...workbook.styles.get(0).font, bold: true },
    });

    history.execute(new StructuralEditCommand(sheet.name, 'row', 'insert', 0, 1));
    expect(num(sheet.getValue(last, 0))).toBe(111); // survivor kept, not overwritten
    expect(sheet.getValue(last - 1, 0).kind).toBe('empty'); // its old slot vacated
  });

  it('insert columns is refused when a merge or table would be cut off at the right edge', () => {
    const { history, sheet } = setup();
    const lastColumn = SheetDimensions.maxColumns - 1;
    sheet.addMerge(new MergedRegion(CellRange.fromBounds(0, lastColumn - 1, 0, lastColumn)));
    expect(() => history.execute(new StructuralEditCommand(sheet.name, 'column', 'insert', 3, 1))).toThrow(
      EditRefusedError,
    );
    expect(sheet.merges[0]!.range.right).toBe(lastColumn);
  });

  it('delete rows breaks a reference to a deleted cell', () => {
    const { history, sheet } = setup();
    sheet.setValue(2, 0, CellValue.number(9)); // A3
    sheet.setFormula(0, 1, 'A3+1', CellValue.empty); // B1 references A3
    history.execute(new StructuralEditCommand(sheet.name, 'row', 'delete', 2, 1));
    expect(sheet.find(0, 1)!.formula).toBe('#REF!+1');
  });

  it('merge and unmerge are undoable', () => {
    const { history, sheet } = setup();
    history.execute(EditCommands.mergeCells(sheet.name, CellRange.parse('A1:B2')));
    expect(sheet.findMerge(0, 0)).toBeDefined();
    history.undo();
    expect(sheet.findMerge(0, 0)).toBeUndefined();
  });

  it('hide columns is undoable', () => {
    const { history, sheet } = setup();
    history.execute(EditCommands.setColumnsHidden(sheet.name, 1, 2, true));
    expect(sheet.isColumnHidden(1)).toBe(true);
    expect(sheet.isColumnHidden(2)).toBe(true);
    history.undo();
    expect(sheet.isColumnHidden(1)).toBe(false);
  });
});
