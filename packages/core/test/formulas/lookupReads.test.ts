import { afterEach, describe, expect, it, vi } from 'vitest';
import { Worksheet } from '../../src/model/worksheet.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { evaluate } from './formulaHarness.ts';

const ROWS = 50_000;

function column(sheet: Worksheet): void {
  for (let r = 0; r < ROWS; r++) sheet.setValue(r, 0, CellValue.number(r + 1));
  for (let r = 0; r < ROWS; r++) sheet.setValue(r, 1, CellValue.text(`v${r + 1}`));
}

/** Evaluate with `find` counted only while the formula runs. */
function readsFor(formula: string): { value: CellValue; reads: number } {
  let spy: ReturnType<typeof vi.spyOn> | null = null;
  const value = evaluate(formula, (sheet) => {
    column(sheet);
    spy = vi.spyOn(Worksheet.prototype, 'find');
  });
  const reads = spy!.mock.calls.length;
  return { value, reads };
}

afterEach(() => vi.restoreAllMocks());

describe('Exact lookups stop at the first match', () => {
  it('MATCH', () => {
    const { value, reads } = readsFor('MATCH(3, A1:A50000, 0)');
    expect(value).toEqual(CellValue.number(3));
    expect(reads).toBeLessThan(100);
  });

  it('VLOOKUP', () => {
    const { value, reads } = readsFor('VLOOKUP(3, A1:B50000, 2, FALSE)');
    expect(value).toEqual(CellValue.text('v3'));
    expect(reads).toBeLessThan(100);
  });
});
