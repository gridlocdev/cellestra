import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { RecalculationEngine } from '../../src/dependencies/index.ts';
import { CellAddress } from '../../src/model/cellAddress.ts';
import type { Workbook } from '../../src/model/workbook.ts';
import { parseCsv } from '../../src/serialization/csv/csvReader.ts';
import { readOds, writeOds, XlsxReader } from '../../src/serialization/index.ts';
import { NumberFormatter } from '../../src/styling/numberFormatter.ts';

/**
 * The LibreOffice sample saved as `.xlsx` by LibreOffice (scripts/make-libreoffice-sample.py writes
 * both), checked against what LibreOffice shows for the `.ods` (PRD §8.19).
 */
const fixture = (name: string) => new URL(`./fixtures/${name}`, import.meta.url);
const bytes = new Uint8Array(readFileSync(fixture('libreoffice-sample.xlsx')));
const libreOffice = shownByLibreOffice(readFileSync(fixture('libreoffice-sample.shown.txt'), 'utf8'));

/** Calc's Err:502 (a bad argument) has no standard code; Calc saves it to .xlsx as #VALUE!. */
function sameError(shown: string): string {
  return shown === 'Err:502' ? '#VALUE!' : shown;
}

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

describe('an .xlsx LibreOffice saved', () => {
  it('opens with its sheets, layout and names', () => {
    const { workbook, report } = XlsxReader.read(bytes);
    expect(workbook.worksheets.map((s) => s.name)).toEqual(['Data', 'Lookup', 'Q2 Sales', 'Empty']);
    const data = workbook.worksheets[0]!;
    expect(data.merges.map((m) => m.range.toString())).toEqual(['A1:H1']);
    expect(data.isColumnHidden(8)).toBe(true);
    expect(data.isRowHidden(10)).toBe(true);
    expect(data.find(2, 6)?.formula).toBe('ROUND(B3*C3*(1-E3), 2)');
    expect(workbook.namedRanges.map((n) => n.name)).toContain('Prices');
    expect(report.leftOut.has('dataValidation')).toBe(true);
    expect(report.leftOut.has('comments')).toBe(true);
  });

  it('shows every cell as LibreOffice does', () => {
    expectShownLikeLibreOffice(XlsxReader.read(bytes).workbook);
  });

  it('works every formula out to the same result as LibreOffice', () => {
    const { workbook } = XlsxReader.read(bytes);
    for (const sheet of workbook.worksheets)
      for (const [, , cell] of sheet.enumerateCells()) if (cell.isFormula) cell.value = { kind: 'empty' };
    const engine = new RecalculationEngine(workbook);
    engine.rebuildGraph();
    engine.recalculateAll();
    expectShownLikeLibreOffice(workbook);
  });

  it('keeps every cell and formula when saved as .ods and reopened', () => {
    const imported = XlsxReader.read(bytes).workbook;
    const reopened = readOds(writeOds(imported));
    expectShownLikeLibreOffice(reopened);
    const formulas = (wb: Workbook) =>
      wb.worksheets.flatMap((s) =>
        [...s.enumerateCells()]
          .filter(([, , cell]) => cell.isFormula)
          .map(([r, c, cell]) => `${s.name}!${r},${c}=${cell.formula}`),
      );
    expect(formulas(reopened)).toEqual(formulas(imported));
    expect(formulas(imported)).toEqual(formulas(readOds(readFileSync(fixture('libreoffice-sample.ods')))));
  });
});
