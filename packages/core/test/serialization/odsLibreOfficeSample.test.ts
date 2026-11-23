import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { RecalculationEngine } from '../../src/dependencies/index.ts';
import { CellAddress } from '../../src/model/cellAddress.ts';
import type { Workbook } from '../../src/model/workbook.ts';
import { parseCsv } from '../../src/serialization/csv/csvReader.ts';
import { readOds, writeOds } from '../../src/serialization/index.ts';
import { NumberFormatter } from '../../src/styling/numberFormatter.ts';

/**
 * A file LibreOffice saved (scripts/make-libreoffice-sample.py builds it), and what LibreOffice
 * shows in each of its cells after recalculating (the same script's `--check`).
 */
const fixture = (name: string) => new URL(`./fixtures/${name}`, import.meta.url);
const bytes = new Uint8Array(readFileSync(fixture('libreoffice-sample.ods')));
const libreOffice = shownByLibreOffice(readFileSync(fixture('libreoffice-sample.shown.txt'), 'utf8'));

/** Calc's Err:502 (a bad argument) has no standard code; Calc saves it to .xlsx as #VALUE!. */
function sameError(shown: string): string {
  return shown === 'Err:502' ? '#VALUE!' : shown;
}

/** `== Sheet` headings, each followed by that sheet as CSV, into address → shown text. */
function shownByLibreOffice(text: string): Map<string, Map<string, string>> {
  const sheets = new Map<string, Map<string, string>>();
  for (const part of text.split(/^== /m).slice(1)) {
    const newline = part.indexOf('\n');
    const cells = new Map<string, string>();
    parseCsv(part.slice(newline + 1)).forEach((row, r) =>
      row.forEach((value, c) => {
        if (value !== '') cells.set(new CellAddress(c, r).toString(), sameError(value));
      }),
    );
    sheets.set(part.slice(0, newline), cells);
  }
  return sheets;
}

/** Every non-empty cell as Cellestra shows it, per sheet. */
function shownByCellestra(workbook: Workbook): Map<string, Map<string, string>> {
  const locale = workbook.settings.effectiveLocale;
  const sheets = new Map<string, Map<string, string>>();
  for (const sheet of workbook.worksheets) {
    const cells = new Map<string, string>();
    for (const [row, column, cell] of sheet.enumerateCells()) {
      const code = workbook.styles.get(sheet.effectiveStyleId(row, column)).numberFormat.code;
      const text = NumberFormatter.format(cell.value, code, locale);
      if (text !== '') cells.set(new CellAddress(column, row).toString(), sameError(text));
    }
    sheets.set(sheet.name, cells);
  }
  return sheets;
}

function expectShownLikeLibreOffice(workbook: Workbook): void {
  const ours = shownByCellestra(workbook);
  for (const [name, cells] of libreOffice) {
    expect({ sheet: name, cells: Object.fromEntries(ours.get(name) ?? []) }).toEqual({
      sheet: name,
      cells: Object.fromEntries(cells),
    });
  }
}

describe('a file LibreOffice saved', () => {
  it('opens with its sheets, layout and names', () => {
    const workbook = readOds(bytes);
    expect(workbook.worksheets.map((s) => s.name)).toEqual(['Data', 'Lookup', 'Q2 Sales', 'Empty']);
    const data = workbook.worksheets[0]!;
    expect(data.merges.map((m) => m.range.toString())).toEqual(['A1:H1']);
    expect(data.isColumnHidden(8)).toBe(true);
    expect(data.isRowHidden(10)).toBe(true);
    expect(data.validations).toHaveLength(1);
    expect(data.find(1, 0)?.formula).toBeNull();
    expect(data.find(2, 6)?.formula).toBe('ROUND(B3*C3*(1-E3), 2)');
    expect(workbook.namedRanges.map((n) => n.name)).toContain('Prices');
  });

  it('shows every cell as LibreOffice does', () => {
    expectShownLikeLibreOffice(readOds(bytes));
  });

  it('works every formula out to the same result as LibreOffice', () => {
    const workbook = readOds(bytes);
    for (const sheet of workbook.worksheets)
      for (const [, , cell] of sheet.enumerateCells()) if (cell.isFormula) cell.value = { kind: 'empty' };
    const engine = new RecalculationEngine(workbook);
    engine.rebuildGraph();
    engine.recalculateAll();
    expectShownLikeLibreOffice(workbook);
  });

  it('keeps every cell when saved and reopened', () => {
    const reopened = readOds(writeOds(readOds(bytes)));
    expectShownLikeLibreOffice(reopened);
    const formulas = (wb: Workbook) =>
      wb.worksheets.flatMap((s) =>
        [...s.enumerateCells()].map(([r, c, cell]) => `${s.name}!${r},${c}=${cell.formula}`),
      );
    expect(formulas(reopened)).toEqual(formulas(readOds(bytes)));
  });
});
