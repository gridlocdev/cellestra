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

  let out = '';
  for (let row = 0; row <= used.bottom; row++) {
    for (let column = 0; column <= used.right; column++) {
      if (column > 0) out += options.delimiter;
      out += quote(displayValue(workbook, sheet, row, column, locale), options.delimiter);
    }
    out += '\n'; // LF on export
  }
  return out;
}

function displayValue(
  workbook: Workbook,
  sheet: Worksheet,
  row: number,
  column: number,
  locale: string,
): string {
  const cell = sheet.find(row, column);
  if (!cell || cell.value.kind === 'empty') return '';
  const format = workbook.styles.get(cell.styleId).numberFormat;
  return NumberFormatter.format(cell.value, format.code, locale);
}

function quote(field: string, delimiter: string): string {
  const mustQuote =
    field.includes(delimiter) || field.includes('"') || field.includes('\n') || field.includes('\r');
  return mustQuote ? '"' + field.replaceAll('"', '""') + '"' : field;
}
