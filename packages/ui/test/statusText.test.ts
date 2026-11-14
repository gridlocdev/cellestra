import { describe, expect, it } from 'vitest';
import { EditResult } from '@cellestra/core';
import {
  calculationBadge,
  clockTime,
  describeCells,
  narrateEdit,
  narrateHistory,
  pastTense,
  saveStateText,
} from '../src/session/statusText.ts';

const cell = (row: number, column: number, sheet = 'Sheet1') => ({ sheet, row, column });

describe('pastTense', () => {
  it('turns a leading verb into the past and keeps other labels', () => {
    expect(pastTense('Paste')).toEqual({ text: 'Pasted', verb: true });
    expect(pastTense('Insert Rows')).toEqual({ text: 'Inserted rows', verb: true });
    expect(pastTense('Set Value')).toEqual({ text: 'Set value', verb: true });
    expect(pastTense('Bold')).toEqual({ text: 'Bold', verb: false });
    expect(pastTense('Page Setup')).toEqual({ text: 'Page Setup', verb: false });
  });
});

describe('describeCells', () => {
  it('names one cell, a filled rectangle, or a count', () => {
    expect(describeCells([])).toBeNull();
    expect(describeCells([cell(1, 1)])).toBe('B2');
    const block = [];
    for (let r = 0; r < 20; r++) for (let c = 0; c < 3; c++) block.push(cell(r, c));
    expect(describeCells(block)).toBe('A1:C20');
    expect(describeCells([cell(0, 0), cell(5, 5)])).toBe('2 cells');
    expect(describeCells([cell(0, 0), cell(0, 0, 'Other')])).toBe('2 cells');
  });

  it('counts a cell reported twice once', () => {
    expect(describeCells([cell(0, 0), cell(0, 1), cell(0, 1)])).toBe('A1:B1');
  });
});

describe('narrateEdit', () => {
  const sel = { top: 1, left: 1, bottom: 3, right: 2 };
  it('reads like the PRD examples', () => {
    expect(narrateEdit('Paste', EditResult.cells([cell(0, 0), cell(0, 1)]), sel)).toBe('Pasted A1:B1');
    expect(narrateEdit('Sort', EditResult.cells([cell(0, 0), cell(4, 2)]), sel)).toBe('Sorted 2 cells');
    expect(narrateEdit('Insert Rows', EditResult.structuralChange, sel)).toBe('Inserted rows');
  });

  it('falls back to the selection when the edit reports no cells', () => {
    expect(narrateEdit('Bold', EditResult.none, sel)).toBe('Bold · B2:C4');
    expect(narrateEdit('Format Cells', EditResult.none, sel)).toBe('Formatted cells B2:C4');
    expect(narrateEdit('Bold', EditResult.none, null)).toBe('Bold');
  });

  it('names the step undo and redo moved over', () => {
    expect(narrateHistory('undo', 'Paste')).toBe('Undid Paste');
    expect(narrateHistory('redo', 'Sort')).toBe('Redid Sort');
  });
});

describe('save state and calculation badge', () => {
  const at = new Date(2026, 10, 14, 9, 5);
  it('formats the time as HH:mm', () => {
    expect(clockTime(at)).toBe('09:05');
  });

  it('says saved, edited, or not saved', () => {
    expect(saveStateText(false, true, at)).toBe('Saved 09:05');
    expect(saveStateText(true, true, at)).toBe('Edited');
    expect(saveStateText(true, false, null)).toBe('Not saved');
    expect(saveStateText(false, false, null)).toBe('');
  });

  it('labels the calculation mode', () => {
    expect(calculationBadge('automatic')).toBe('Auto');
    expect(calculationBadge('manual')).toBe('Manual');
  });
});
