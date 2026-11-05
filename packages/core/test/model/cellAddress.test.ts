import { describe, expect, it } from 'vitest';
import { CellAddress } from '../../src/model/cellAddress.ts';
import { SheetDimensions } from '../../src/model/sheetDimensions.ts';

describe('CellAddress', () => {
  it.each([
    ['A1', 0, 0],
    ['B7', 1, 6],
    ['Z1', 25, 0],
    ['AA1', 26, 0],
    ['AB10', 27, 9],
    ['XFD1048576', 16383, 1048575],
  ])('parse reads column and row of %s', (text, expectedColumn, expectedRow) => {
    const address = CellAddress.parse(text);
    expect(address.column).toBe(expectedColumn);
    expect(address.row).toBe(expectedRow);
  });

  it.each([
    ['$A$1', true, true],
    ['$A1', true, false],
    ['A$1', false, true],
    ['A1', false, false],
  ])('parse reads absolute flags of %s', (text, colAbs, rowAbs) => {
    const address = CellAddress.parse(text);
    expect(address.columnAbsolute).toBe(colAbs);
    expect(address.rowAbsolute).toBe(rowAbs);
  });

  it.each([
    ['Sheet2.A1', 'Sheet2'],
    ["'My Sheet'.B2", 'My Sheet'],
    ["'O''Brien'.C3", "O'Brien"],
  ])('parse reads the sheet qualifier of %s', (text, expectedSheet) => {
    expect(CellAddress.parse(text).sheet).toBe(expectedSheet);
  });

  it.each(['A1', '$A$1', 'Sheet2.B7', "'My Sheet'.B7"])('toString round-trips parse: %s', (text) => {
    expect(CellAddress.parse(text).toString()).toBe(text);
  });

  it.each(['', '1A', 'A', '1', 'A0', 'A1B', '$'])('tryParse rejects garbage %j', (text) => {
    expect(CellAddress.tryParse(text)).toBeNull();
  });

  it('shift moves relative parts but not absolute', () => {
    expect(CellAddress.parse('$B$2').shift(3, 4).equals(CellAddress.parse('$B$2'))).toBe(true);
    expect(CellAddress.parse('B2').shift(3, 4).equals(CellAddress.parse('E6'))).toBe(true);
    expect(CellAddress.parse('$B2').shift(3, 4).equals(CellAddress.parse('$B6'))).toBe(true);
  });

  it('column letters round-trip', () => {
    for (let column = 0; column < 16384; column += 257) {
      expect(SheetDimensions.lettersToColumn(SheetDimensions.columnToLetters(column))).toBe(column);
    }
  });
});
