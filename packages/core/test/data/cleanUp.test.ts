import { describe, expect, it } from 'vitest';
import { CleanUp } from '../../src/data/cleanUp.ts';
import { Defaults } from '../../src/defaults.ts';
import { Workbook } from '../../src/model/workbook.ts';
import type { Worksheet } from '../../src/model/worksheet.ts';
import { readOds, writeOds } from '../../src/serialization/index.ts';
import { CellStyles, NumberFormats, type CellStyle } from '../../src/styling/cellStyle.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { numberAt, textAt } from './helpers.ts';

/**
 * Clean Up (PRD §8.17): the formatting it infers, the two data tidies it may make, and everything it
 * leaves alone (formulas, chosen number formats, identifiers that merely look numeric).
 */
// A real locale: "$" amounts must parse for the currency case to test anything.
const US = 'en-US';

function book(): [Workbook, Worksheet] {
  const workbook = Workbook.createNew();
  workbook.settings.locale = US;
  return [workbook, workbook.activeSheet];
}

function fill(sheet: Worksheet, ...rows: string[][]): void {
  rows.forEach((cells, r) =>
    cells.forEach((text, c) => {
      if (text.length > 0) sheet.setValue(r, c, CellValue.text(text));
    }),
  );
}

function run(workbook: Workbook, sheet: Worksheet): void {
  CleanUp.forSheet(workbook, sheet, US).toCommand(CleanUp.sheetLabel)?.apply(workbook);
}

const styleOf = (workbook: Workbook, sheet: Worksheet, row: number, column: number): CellStyle =>
  workbook.styles.get(sheet.effectiveStyleId(row, column));

// Region | Revenue / West | 1200.50 / East | 900
function sales(): [Workbook, Worksheet] {
  const [workbook, sheet] = book();
  fill(sheet, ['Region', 'Revenue'], ['West', '1200.50'], ['East', '900']);
  return [workbook, sheet];
}

function restyle(workbook: Workbook, sheet: Worksheet, row: number, column: number, style: CellStyle): void {
  sheet.getOrCreate(row, column).styleId = workbook.styles.intern(style);
}

describe('CleanUp', () => {
  it('header row is emphasized and frozen', () => {
    const [workbook, sheet] = sales();
    run(workbook, sheet);
    const header = styleOf(workbook, sheet, 0, 0);
    expect(header.font.bold).toBe(true);
    expect(header.fill.background).toBe(Defaults.tableHeaderFillArgb);
    expect(header.font.color).toBe(Defaults.tableHeaderTextArgb);
    expect(header.alignment.horizontal).toBe('center');
    expect(sheet.frozenRows).toBe(1);
    expect(sheet.frozenColumns).toBe(0);
  });

  it('numbers stored as text become numbers and the column gets one format', () => {
    const [workbook, sheet] = sales();
    run(workbook, sheet);
    expect(numberAt(sheet, 1, 1)).toBe(1200.5);
    expect(numberAt(sheet, 2, 1)).toBe(900);
    const format = styleOf(workbook, sheet, 1, 1).numberFormat.code;
    expect(format).toBe('#,##0.00');
    expect(styleOf(workbook, sheet, 2, 1).numberFormat.code).toBe(format);
  });

  it('a column of whole numbers below a thousand is left general', () => {
    const [workbook, sheet] = book();
    fill(sheet, ['Count'], ['3'], ['12']);
    run(workbook, sheet);
    expect(NumberFormats.isGeneral(styleOf(workbook, sheet, 1, 0).numberFormat)).toBe(true);
  });

  it('large whole numbers get a thousands separator', () => {
    const [workbook, sheet] = book();
    fill(sheet, ['Count'], ['4200'], ['12']);
    run(workbook, sheet);
    expect(styleOf(workbook, sheet, 1, 0).numberFormat.code).toBe('#,##0');
  });

  it('percentages and currency are read as what they are', () => {
    const [workbook, sheet] = book();
    fill(sheet, ['Share', 'Price'], ['45.5%', '$12.00'], ['54.5%', '$8.50']);
    run(workbook, sheet);
    expect(numberAt(sheet, 1, 0)).toBeCloseTo(0.455, 10);
    expect(styleOf(workbook, sheet, 1, 0).numberFormat.code).toBe('0.0%');
    expect(numberAt(sheet, 1, 1)).toBe(12);
    expect(NumberFormats.kind(styleOf(workbook, sheet, 1, 1).numberFormat)).toBe('currency');
  });

  it('dates typed as text become dates', () => {
    const [workbook, sheet] = book();
    fill(sheet, ['When'], ['2026-01-05'], ['2026-02-06']);
    run(workbook, sheet);
    expect(sheet.getValue(1, 0).kind).toBe('number');
    expect(NumberFormats.kind(styleOf(workbook, sheet, 1, 0).numberFormat)).toBe('date');
  });

  it('an identifier column keeps its leading zeros', () => {
    const [workbook, sheet] = book();
    fill(sheet, ['Part'], ['007'], ['0912']);
    run(workbook, sheet);
    expect(textAt(sheet, 1, 0)).toBe('007');
    expect(NumberFormats.isGeneral(styleOf(workbook, sheet, 1, 0).numberFormat)).toBe(true);
  });

  it('phone numbers with a leading plus stay text', () => {
    const [workbook, sheet] = book();
    fill(sheet, ['Phone'], ['+15551234567'], ['+447700900123']);
    run(workbook, sheet);
    expect(textAt(sheet, 1, 0)).toBe('+15551234567');
    expect(textAt(sheet, 2, 0)).toBe('+447700900123');
  });

  it('codes that look like scientific notation stay text', () => {
    const [workbook, sheet] = book();
    fill(sheet, ['Code'], ['12E3'], ['4e10']);
    run(workbook, sheet);
    expect(textAt(sheet, 1, 0)).toBe('12E3');
    expect(textAt(sheet, 2, 0)).toBe('4e10');
  });

  it('one stray identifier stops the whole column being coerced', () => {
    const [workbook, sheet] = book();
    fill(sheet, ['Code'], ['12'], ['007']);
    run(workbook, sheet);
    expect(textAt(sheet, 1, 0)).toBe('12');
    expect(textAt(sheet, 2, 0)).toBe('007');
  });

  // One column takes one format, and a date under a numeric one is a five-digit serial.
  it('a column mixing dates and numbers is left alone', () => {
    const [workbook, sheet] = book();
    fill(sheet, ['Value'], ['2026-01-05'], ['42']);
    run(workbook, sheet);
    expect(textAt(sheet, 1, 0)).toBe('2026-01-05');
    expect(textAt(sheet, 2, 0)).toBe('42');
    expect(NumberFormats.isGeneral(styleOf(workbook, sheet, 1, 0).numberFormat)).toBe(true);
  });

  it('a mixed column is never half converted', () => {
    const [workbook, sheet] = book();
    fill(sheet, ['Note'], ['12'], ['pending']);
    run(workbook, sheet);
    expect(textAt(sheet, 1, 0)).toBe('12');
    expect(textAt(sheet, 2, 0)).toBe('pending');
  });

  it('surrounding whitespace is trimmed even when the column stays text', () => {
    const [workbook, sheet] = book();
    fill(sheet, ['Name'], ['  West  '], ['East']);
    run(workbook, sheet);
    expect(textAt(sheet, 1, 0)).toBe('West');
  });

  it('a whitespace-only cell is cleared', () => {
    const [workbook, sheet] = book();
    fill(sheet, ['Name'], ['West'], ['   ']);
    run(workbook, sheet);
    expect(sheet.getValue(2, 0).kind).toBe('empty');
  });

  it('formulas are never rewritten', () => {
    const [workbook, sheet] = book();
    fill(sheet, ['Region', 'Revenue'], ['West', '1200'], ['East', '900']);
    sheet.setFormula(3, 1, 'SUM(B2:B3)', CellValue.number(2100));
    run(workbook, sheet);
    expect(sheet.find(3, 1)?.formula).toBe('SUM(B2:B3)');
    expect(numberAt(sheet, 3, 1)).toBe(2100);
  });

  it('a chosen number format survives', () => {
    const [workbook, sheet] = sales();
    restyle(workbook, sheet, 1, 1, { ...CellStyles.default, numberFormat: NumberFormats.of('0.0000') });
    run(workbook, sheet);
    expect(styleOf(workbook, sheet, 1, 1).numberFormat.code).toBe('0.0000');
    expect(styleOf(workbook, sheet, 2, 1).numberFormat.code).toBe('#,##0.00');
  });

  it('body alignment returns to the type-driven default', () => {
    const [workbook, sheet] = sales();
    const d = CellStyles.default;
    restyle(workbook, sheet, 1, 0, { ...d, alignment: { ...d.alignment, horizontal: 'right' } });
    run(workbook, sheet);
    expect(styleOf(workbook, sheet, 1, 0).alignment.horizontal).toBe('general');
  });

  it('columns are fitted to their contents', () => {
    const [workbook, sheet] = book();
    fill(sheet, ['Region'], ['A considerably longer label than the default width'], ['East']);
    run(workbook, sheet);
    expect(sheet.getColumnWidth(0)).toBeGreaterThan(Defaults.columnWidth);
  });

  it('a wrapped cell does not drag its column wide', () => {
    const [workbook, sheet] = book();
    fill(sheet, ['Note'], ['a very long wrapped note that folds at the column edge'], ['ok']);
    const d = CellStyles.default;
    restyle(workbook, sheet, 1, 0, { ...d, alignment: { ...d.alignment, wrapText: true } });
    run(workbook, sheet);
    expect(sheet.getColumnWidth(0)).toBeLessThan(200);
  });

  it('planning changes nothing until the command runs', () => {
    const [workbook, sheet] = sales();
    const plan = CleanUp.forSheet(workbook, sheet, US);
    expect(plan.isEmpty).toBe(false);
    expect(textAt(sheet, 1, 1)).toBe('1200.50');
    expect(sheet.frozenRows).toBe(0);
  });

  it('undo puts back values, styles, widths and panes', () => {
    const [workbook, sheet] = sales();
    const width = sheet.getColumnWidth(0);
    const command = CleanUp.forSheet(workbook, sheet, US).toCommand(CleanUp.sheetLabel)!;
    command.apply(workbook);
    command.undo(workbook);
    expect(textAt(sheet, 1, 1)).toBe('1200.50');
    expect(styleOf(workbook, sheet, 0, 0).font.bold).toBe(false);
    expect(sheet.getColumnWidth(0)).toBe(width);
    expect(sheet.frozenRows).toBe(0);
  });

  it('running it twice is a no-op the second time', () => {
    const [workbook, sheet] = sales();
    run(workbook, sheet);
    const second = CleanUp.forSheet(workbook, sheet, US);
    expect(second.isEmpty, second.summary).toBe(true);
  });

  it('each data block on a sheet is cleaned on its own', () => {
    const [workbook, sheet] = book();
    fill(sheet, ['Region', 'Revenue'], ['West', '1200']);
    sheet.setValue(5, 0, CellValue.text('Owner'));
    sheet.setValue(5, 1, CellValue.text('Tasks'));
    sheet.setValue(6, 0, CellValue.text('Ada'));
    sheet.setValue(6, 1, CellValue.text('4'));
    run(workbook, sheet);
    expect(styleOf(workbook, sheet, 0, 0).font.bold).toBe(true);
    expect(styleOf(workbook, sheet, 5, 0).font.bold).toBe(true);
    expect(numberAt(sheet, 6, 1)).toBe(4);
    // Only the topmost block's header is what the frozen pane holds still.
    expect(sheet.frozenRows).toBe(1);
  });

  it('a whole-workbook clean up reports a line per sheet', () => {
    const [workbook, first] = sales();
    const second = workbook.addSheet('Sheet2');
    second.setValue(0, 0, CellValue.text('Name'));
    second.setValue(1, 0, CellValue.text('Ada'));
    second.setValue(0, 1, CellValue.text('Score'));
    second.setValue(1, 1, CellValue.text('9'));
    const plan = CleanUp.forWorkbook(workbook, US);
    plan.toCommand(CleanUp.workbookLabel)!.apply(workbook);
    expect(plan.notes).toHaveLength(2);
    expect(plan.notes.some((n) => n.sheet === first.name)).toBe(true);
    expect(plan.notes.some((n) => n.sheet === 'Sheet2')).toBe(true);
    expect(numberAt(second, 1, 1)).toBe(9);
  });

  // A formatting move Clean Up can make but the file cannot keep would make it report work every time.
  it('nothing is left to do after a save and reopen', () => {
    const [workbook, sheet] = sales();
    CleanUp.forSheet(workbook, sheet, US).toCommand(CleanUp.sheetLabel)!.apply(workbook);
    const reopened = readOds(writeOds(workbook));
    const again = CleanUp.forSheet(reopened, reopened.activeSheet, US);
    expect(again.isEmpty, again.summary).toBe(true);
  });

  it('an empty sheet has nothing to clean', () => {
    const [workbook, sheet] = book();
    const plan = CleanUp.forSheet(workbook, sheet, US);
    expect(plan.isEmpty).toBe(true);
    expect(plan.summary).toBe('Nothing to clean up.');
    expect(plan.toCommand(CleanUp.sheetLabel)).toBeNull();
  });
});
