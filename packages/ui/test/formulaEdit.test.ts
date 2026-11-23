import { describe, expect, it } from 'vitest';
import { CellRange, CellValue, Workbook, table } from '@cellestra/core';
import {
  REFERENCE_COLORS,
  columnCompletion,
  fillDownTarget,
  fillTarget,
  formulaReferences,
  functionPrefix,
  functionSuggestions,
  isReferenceExpected,
  referenceText,
  signatureHint,
} from '../src/grid/formulaEdit.ts';

describe('isReferenceExpected', () => {
  it.each([
    ['=', 1, true],
    ['=SUM(', 5, true],
    ['=A1+', 4, true],
    ['=SUM(A1, ', 9, true],
    ['=A1', 3, false],
    ['hello', 5, false],
    ['=SUM(A1)', 8, false],
  ])('%s at %i → %s', (text, caret, expected) => {
    expect(isReferenceExpected(text, caret)).toBe(expected);
  });
});

describe('referenceText', () => {
  it('writes a cell and a range', () => {
    expect(referenceText({ top: 0, left: 0, bottom: 0, right: 0 })).toBe('A1');
    expect(referenceText({ top: 1, left: 1, bottom: 4, right: 2 })).toBe('B2:C5');
  });
});

describe('formulaReferences', () => {
  it('colours each reference in order and skips other sheets', () => {
    const refs = formulaReferences('=SUM(A1:B2)+C3*Other.D4+E5', 'Sheet1');
    expect(refs).toEqual([
      { range: { top: 0, left: 0, bottom: 1, right: 1 }, color: REFERENCE_COLORS[0] },
      { range: { top: 2, left: 2, bottom: 2, right: 2 }, color: REFERENCE_COLORS[1] },
      { range: { top: 4, left: 4, bottom: 4, right: 4 }, color: REFERENCE_COLORS[3] },
    ]);
  });

  it('outlines a 3D range on every sheet it spans, in tab order', () => {
    const order = ['Sheet1', 'Middle', 'Sheet3', 'After'];
    const between = (first: string, last: string): readonly string[] | null => {
      const a = order.indexOf(first);
      const b = order.indexOf(last);
      return a < 0 || b < 0 ? null : order.slice(Math.min(a, b), Math.max(a, b) + 1);
    };
    const text = '=SUM(Sheet1.A1:Sheet3.B2)+C3';
    const span = { range: { top: 0, left: 0, bottom: 1, right: 1 }, color: REFERENCE_COLORS[0] };
    const c3 = { range: { top: 2, left: 2, bottom: 2, right: 2 }, color: REFERENCE_COLORS[1] };
    expect(formulaReferences(text, 'Middle', between)).toEqual([span, c3]);
    expect(formulaReferences(text, 'sheet3', between)).toEqual([span, c3]);
    // A sheet outside the span shows only its own reference, which keeps its colour slot.
    expect(formulaReferences(text, 'After', between)).toEqual([c3]);
    // Without the tab order, only the end sheets are known to be in it.
    expect(formulaReferences(text, 'Middle')).toEqual([c3]);
    expect(formulaReferences(text, 'Sheet1')).toEqual([span, c3]);
  });

  it('is empty for text and broken formulas', () => {
    expect(formulaReferences('A1', 'Sheet1')).toEqual([]);
    expect(formulaReferences('=SUM(A1', 'Sheet1')).toEqual([]);
  });
});

describe('signatureHint', () => {
  it('marks the argument the caret is in', () => {
    const hint = signatureHint('=ROUND(A1, ', 11)!;
    expect(hint.meta.name).toBe('ROUND');
    expect(hint.parts.find((p) => p.active)?.text).toContain('digits');
  });

  it('keeps the repeatable argument active past the end', () => {
    const hint = signatureHint('=SUM(1, 2, 3, ', 14)!;
    expect(hint.parts.find((p) => p.active)?.text).toContain('…');
  });

  it('is null outside a call', () => {
    expect(signatureHint('=A1+', 4)).toBeNull();
    expect(signatureHint('SUM(', 4)).toBeNull();
  });
});

describe('function suggestions', () => {
  it('finds the name being typed', () => {
    expect(functionPrefix('=SU', 3)).toEqual({ start: 1, prefix: 'SU' });
    expect(functionPrefix('=A1+vl', 6)).toEqual({ start: 4, prefix: 'vl' });
    expect(functionPrefix('="SU', 4)).toBeNull();
    expect(functionPrefix('SU', 2)).toBeNull();
  });

  it('lists matching functions alphabetically', () => {
    const names = functionSuggestions('su').map((m) => m.name);
    expect(names).toContain('SUM');
    expect(names).toEqual([...names].sort());
    expect(names.every((n) => n.startsWith('SU'))).toBe(true);
  });
});

describe('columnCompletion', () => {
  const sheetWith = (values: (string | null)[]) => {
    const wb = Workbook.createNew();
    const sheet = wb.worksheets[0]!;
    values.forEach((v, row) => {
      if (v !== null) sheet.setValue(row, 0, CellValue.text(v));
    });
    return sheet;
  };

  it('offers the nearest matching entry above, case-sensitively', () => {
    const sheet = sheetWith(['Apple', 'Apricot', 'Banana']);
    expect(columnCompletion(sheet, 3, 0, 'Ap')).toBe('Apricot');
    expect(columnCompletion(sheet, 3, 0, 'ap')).toBeNull();
    expect(columnCompletion(sheet, 3, 0, 'Apricot')).toBeNull();
  });

  it('stops at a gap and at a table header', () => {
    expect(columnCompletion(sheetWith(['Apple', null, 'Banana']), 3, 0, 'Ap')).toBeNull();
    const sheet = sheetWith(['Fruit', 'Banana']);
    sheet.addTable(table('T', CellRange.parse('A1:A3')));
    expect(columnCompletion(sheet, 2, 0, 'Fr')).toBeNull();
    expect(columnCompletion(sheet, 2, 0, 'Ba')).toBe('Banana');
  });
});

describe('fill targets', () => {
  it('covers the source and the pointer in both axes', () => {
    expect(fillTarget({ top: 1, left: 1, bottom: 2, right: 1 }, 5, 3)).toEqual({
      top: 1,
      left: 1,
      bottom: 5,
      right: 3,
    });
    expect(fillTarget({ top: 4, left: 2, bottom: 4, right: 2 }, 1, 2)).toEqual({
      top: 1,
      left: 2,
      bottom: 4,
      right: 2,
    });
  });

  it('fills down to the neighbouring column’s data', () => {
    const wb = Workbook.createNew();
    const sheet = wb.worksheets[0]!;
    for (let r = 0; r < 5; r++) sheet.setValue(r, 0, CellValue.number(r));
    expect(fillDownTarget(sheet, { top: 0, left: 1, bottom: 0, right: 1 })).toEqual({
      top: 0,
      left: 1,
      bottom: 4,
      right: 1,
    });
    // From column A it looks right.
    expect(fillDownTarget(sheet, { top: 0, left: 0, bottom: 4, right: 0 })).toBeNull();
  });
});
