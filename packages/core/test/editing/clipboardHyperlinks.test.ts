import { describe, expect, it } from 'vitest';
import { ClipboardBlock } from '../../src/editing/clipboardBlock.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { FillCommand } from '../../src/editing/fillCommand.ts';
import { PasteBlockCommand } from '../../src/editing/pasteBlockCommand.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { CellValue } from '../../src/values/cellValue.ts';

const R = (a1: string): CellRange => CellRange.parse(a1);
const LINK = { target: 'https://example.com/', tooltip: 'Example' };

function setup() {
  const workbook = Workbook.createNew();
  const sheet = workbook.activeSheet;
  sheet.setValue(0, 0, CellValue.text('Example'));
  sheet.getOrCreate(0, 0).hyperlink = LINK;
  return { history: new EditHistory(workbook), sheet };
}

describe('Hyperlinks travel with their cells', () => {
  it('on copy and paste', () => {
    const { history, sheet } = setup();
    history.execute(new PasteBlockCommand(sheet.name, [2, 2], ClipboardBlock.capture(sheet, R('A1'), false)));
    expect(sheet.find(2, 2)?.hyperlink).toEqual(LINK);
    expect(sheet.find(0, 0)?.hyperlink).toEqual(LINK);
  });

  it('on cut and paste', () => {
    const { history, sheet } = setup();
    history.execute(new PasteBlockCommand(sheet.name, [2, 2], ClipboardBlock.capture(sheet, R('A1'), true)));
    expect(sheet.find(2, 2)?.hyperlink).toEqual(LINK);
    history.undo();
    expect(sheet.find(0, 0)?.hyperlink).toEqual(LINK);
  });

  it('on fill', () => {
    const { history, sheet } = setup();
    history.execute(new FillCommand(sheet.name, R('A1'), R('A1:A3')));
    expect(sheet.find(2, 0)?.hyperlink).toEqual(LINK);
  });

  it('but not on a formats-only paste', () => {
    const { history, sheet } = setup();
    history.execute(
      new PasteBlockCommand(sheet.name, [2, 2], ClipboardBlock.capture(sheet, R('A1'), false), 'formats'),
    );
    expect(sheet.find(2, 2)?.hyperlink ?? null).toBeNull();
  });
});
