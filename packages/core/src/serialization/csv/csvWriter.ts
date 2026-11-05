import type { Cell } from '../../model/cell.ts';
import type { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import { NumberFormatter } from '../../styling/numberFormatter.ts';
import { CsvOptions } from './csvOptions.ts';

/**
 * Write a worksheet to CSV. Exports the **displayed/computed** values, never formulas (PRD §8.4):
 * numbers and dates are formatted with each cell's number format and the locale. RFC 4180 quoting;
 * LF line endings. The workbook is passed because a sheet does not know its workbook (styles and
 * the document locale live there).
 */
export function writeCsv(
  workbook: Workbook,
  sheet: Worksheet,
  options: CsvOptions = CsvOptions.default,
): string {
  const locale = options.locale ?? workbook.settings.effectiveLocale;
  const used = sheet.getUsedRange();
  if (!used) return '';

  // Walk only the stored cells: visiting every row × column of the used range took minutes for a
  // sheet with one value in A1 and one in XFD1048576.
  const byRow = new Map<number, Array<[column: number, cell: Cell]>>();
  for (const [row, column, cell] of sheet.enumerateCells()) {
    if (cell.value.kind === 'empty' || row > used.bottom || column > used.right) continue;
    let cells = byRow.get(row);
    if (!cells) byRow.set(row, (cells = []));
    cells.push([column, cell]);
  }

  const { delimiter } = options;
  const blankLine = delimiter.repeat(used.right) + '\n';
  const parts: string[] = [];
  for (let row = 0; row <= used.bottom; row++) {
    const cells = byRow.get(row);
    if (!cells) {
      parts.push(blankLine);
      continue;
    }
    cells.sort((a, b) => a[0] - b[0]);
    let line = '';
    let at = 0;
    for (const [column, cell] of cells) {
      line += delimiter.repeat(column - at) + quote(displayValue(workbook, cell, locale), delimiter);
      at = column;
    }
    parts.push(line + delimiter.repeat(used.right - at) + '\n'); // LF on export
  }
  return parts.join('');
}

function displayValue(workbook: Workbook, cell: Cell, locale: string): string {
  const code = workbook.styles.get(cell.styleId).numberFormat.code;
  const value = cell.value;
  if (
    value.kind === 'number' &&
    Number.isFinite(value.value) &&
    (!code || code === '@' || code.toLowerCase() === 'general')
  ) {
    return generalNumber(value.value, locale);
  }
  return NumberFormatter.format(value, code, locale);
}

/**
 * A General number as CSV carries it: 15 significant digits, the precision a spreadsheet stores.
 * The on-screen General display rounds to ten decimals, which exported 1E-20 as 0 and lost the
 * tail of every long fraction; a CSV is data for the next program, not a cell display.
 */
function generalNumber(n: number, locale: string): string {
  const text = String(Number(n.toPrecision(15)));
  const decimal = decimalSeparator(locale);
  return decimal === '.' ? text : text.replace('.', decimal);
}

const decimalSeparators = new Map<string, string>();

function decimalSeparator(locale: string): string {
  let decimal = decimalSeparators.get(locale);
  if (decimal === undefined) {
    try {
      decimal =
        new Intl.NumberFormat(locale).formatToParts(1.5).find((p) => p.type === 'decimal')?.value ?? '.';
    } catch {
      decimal = '.';
    }
    decimalSeparators.set(locale, decimal);
  }
  return decimal;
}

function quote(field: string, delimiter: string): string {
  const mustQuote =
    field.includes(delimiter) || field.includes('"') || field.includes('\n') || field.includes('\r');
  return mustQuote ? '"' + field.replaceAll('"', '""') + '"' : field;
}
