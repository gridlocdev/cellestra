import { describe, expect, it } from 'vitest';
import { ClipboardBlock } from '../../src/editing/clipboardBlock.ts';
import { PastePlanner } from '../../src/editing/pastePlanner.ts';
import { CellAddress } from '../../src/model/cellAddress.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { SheetDimensions } from '../../src/model/sheetDimensions.ts';
import { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';

// What happens after ⌘A has expanded to the whole sheet (PRD §10.1): a 16384 × 1048576 rectangle that
// no operation may turn into one entry per cell.
const wholeSheet = (): CellRange =>
  new CellRange(
    new CellAddress(0, 0),
    new CellAddress(SheetDimensions.maxColumns - 1, SheetDimensions.maxRows - 1),
  );

function sheetWithABlock(): Worksheet {
  const sheet = new Worksheet('Sheet1');
  sheet.setValue(0, 0, CellValue.number(1)); // A1
  sheet.setValue(0, 1, CellValue.number(2)); // B1
  sheet.setValue(1, 0, CellValue.number(3)); // A2
  sheet.setValue(1, 1, CellValue.number(4)); // B2
  return sheet;
}

describe('whole-sheet selection', () => {
  it('copying the whole sheet captures only the used range', () => {
    const block = ClipboardBlock.capture(sheetWithABlock(), wholeSheet(), false);
    expect(block.rows).toBe(2);
    expect(block.columns).toBe(2);
    expect(block.sourceRange.equals(CellRange.parse('A1:B2'))).toBe(true);
    expect(CellValue.asNumber(block.cells[1]![1]!.value)).toBe(4);
  });

  it('copying the whole sheet of an empty sheet captures one cell', () => {
    const block = ClipboardBlock.capture(new Worksheet('Sheet1'), wholeSheet(), false);
    expect(block.rows).toBe(1);
    expect(block.columns).toBe(1);
    expect(block.cells[0]![0]).toBeNull();
  });

  it('a deliberately selected block of blanks still copies as blanks', () => {
    const block = ClipboardBlock.capture(sheetWithABlock(), CellRange.parse('D1:E3'), false);
    expect(block.rows).toBe(3);
    expect(block.columns).toBe(2);
    expect(block.cells[0]![0]).toBeNull();
  });

  it('filling the whole sheet from one cell anchors instead', () => {
    const plan = PastePlanner.plan(1, 1, wholeSheet(), false);
    expect(plan.anchors).toEqual([[0, 0]]);
    expect(plan.tiled).toBe(false);
    expect(plan.shapeMismatch).toBe(false);
  });

  it('tiling into the whole sheet anchors instead', () => {
    const plan = PastePlanner.plan(2, 2, wholeSheet(), false);
    expect(plan.anchors).toHaveLength(1);
    expect(plan.tiled).toBe(false);
  });

  it('an ordinary fill still covers every selected cell', () => {
    const plan = PastePlanner.plan(1, 1, CellRange.parse('A1:C2'), false);
    expect(plan.anchors).toHaveLength(6);
    expect(plan.tiled).toBe(true);
  });
});
