import { CellStyles, Workbook } from '@cellestra/core';
import { describe, expect, it } from 'vitest';
import { csvLosses } from '../src/session/csvLoss.ts';
import { WorkbookSession } from '../src/session/workbookSession.ts';

describe('csvLosses', () => {
  it('is empty for plain values, dates included', () => {
    const s = new WorkbookSession();
    s.load(new TextEncoder().encode('a,2024-01-02\n1,2\n'), null, 'x.csv');
    expect(csvLosses(s.workbook.peek(), s.activeSheet.peek())).toEqual([]);
  });

  it('names other sheets, formulas and formatting', () => {
    const wb = Workbook.createNew();
    const sheet = wb.activeSheet;
    sheet.setFormula(0, 0, '=1+1');
    sheet.getOrCreate(1, 0).styleId = wb.styles.intern({
      ...CellStyles.default,
      font: { ...CellStyles.default.font, bold: true },
    });
    wb.addSheet();
    expect(csvLosses(wb, sheet)).toEqual(['1 other sheet(s)', 'formulas', 'formatting']);
  });
});

describe('WorkbookSession.markSaved', () => {
  it('stays dirty when edits landed after the saved bytes were serialized', () => {
    const s = new WorkbookSession();
    s.touch();
    const atSerialize = s.editCount;
    s.touch();
    s.markSaved({ store: 'nas', name: 'a.ods', id: 'r:a.ods', version: 'v2' }, atSerialize);
    expect(s.modified.peek()).toBe(true);
    s.markSaved({ store: 'nas', name: 'a.ods', id: 'r:a.ods', version: 'v3' }, s.editCount);
    expect(s.modified.peek()).toBe(false);
  });
});
