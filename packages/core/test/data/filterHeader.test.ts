import { describe, expect, it } from 'vitest';
import { filterSpec, FilterEngine } from '../../src/data/filterEngine.ts';
import { sortSpec, SortEngine } from '../../src/data/sortEngine.ts';
import { CellRange } from '../../src/model/cellRange.ts';
import { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { numberAt, textAt } from './helpers.ts';

/** A header row stays put when sorting and is never hidden or offered as a filter value (PRD §8.11–§8.12). */
function scoreSheet(): Worksheet {
  const sheet = new Worksheet('Sheet1');
  sheet.setValue(0, 0, CellValue.text('Score'));
  sheet.setValue(1, 0, CellValue.number(30));
  sheet.setValue(2, 0, CellValue.number(10));
  sheet.setValue(3, 0, CellValue.number(20));
  return sheet;
}

describe('FilterHeader', () => {
  it('distinct values exclude the header', () => {
    const values = FilterEngine.distinctValues(scoreSheet(), CellRange.parse('A1:A4'), 0, true).map((v) =>
      CellValue.toInvariantString(v),
    );
    expect(values).not.toContain('Score');
    expect(values).toEqual(['30', '10', '20']);
  });

  it('sorting with a header leaves row one in place', () => {
    const sheet = scoreSheet();
    SortEngine.sort(sheet, CellRange.parse('A1:A4'), sortSpec([{ column: 0, ascending: true }], true));
    expect(textAt(sheet, 0, 0)).toBe('Score');
    expect(numberAt(sheet, 1, 0)).toBe(10);
    expect(numberAt(sheet, 3, 0)).toBe(30);
  });

  it('filtering with a header never hides row one', () => {
    const sheet = scoreSheet();
    const spec = filterSpec(new Map([[0, { condition: '>15' }]]), true);
    FilterEngine.apply(sheet, CellRange.parse('A1:A4'), spec, 'en-US');
    expect(sheet.isRowHidden(0)).toBe(false);
    expect(sheet.isRowHidden(2)).toBe(true);
  });
});
