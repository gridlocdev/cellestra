import { describe, expect, it } from 'vitest';
import { StickyTableHeader } from '../../src/data/stickyTableHeader.ts';
import { table } from '../../src/data/types.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';

/** A table over A1:B10 with a header row. */
function sheet(hasHeader = true): Worksheet {
  const s = new Worksheet('Sheet1');
  s.setValue(0, 0, CellValue.text('Team'));
  s.setValue(0, 1, CellValue.text('Points'));
  s.addTable(table('T', CellRange.parse('A1:B10'), { hasHeader }));
  return s;
}

describe('StickyTableHeader', () => {
  it('scrolled into the body, the header row labels the column', () => {
    expect(StickyTableHeader.labelRow(sheet(), 1, 4)).toBe(0);
  });

  it('with the real header on screen there is nothing to repeat', () => {
    expect(StickyTableHeader.labelRow(sheet(), 1, 0)).toBeNull();
  });

  it('scrolled past the last row the letters come back', () => {
    expect(StickyTableHeader.labelRow(sheet(), 1, 10)).toBeNull();
  });

  it('a column outside the table is never relabelled', () => {
    expect(StickyTableHeader.labelRow(sheet(), 5, 4)).toBeNull();
  });

  it('a headerless table has no title to lend', () => {
    expect(StickyTableHeader.labelRow(sheet(false), 1, 4)).toBeNull();
  });
});
