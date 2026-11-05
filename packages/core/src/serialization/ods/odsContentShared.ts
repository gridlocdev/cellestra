import { CellAddress } from '../../model/cellAddress.ts';
import type { CellRange } from '../../model/cellRange.ts';

/**
 * The unit conversions and ODF address formatters shared by the `content.xml` reader, writer,
 * validation, and database-range halves.
 */

export const PX_PER_INCH = 96.0;
export const CM_PER_INCH = 2.54;

/** Per sheet name: each validation rule's target and the name it is written under. */
export type ValidationNames = ReadonlyMap<
  string,
  ReadonlyArray<{ readonly range: CellRange; readonly name: string }>
>;

/** `$Sheet1.$A$1` or `$Sheet1.$A$1:$Sheet1.$B$2`. */
export function odfRangeAddress(range: CellRange, fallbackSheet: string): string {
  const sheet = range.sheet ?? fallbackSheet;
  const start = odfCellAddress(new CellAddress(range.left, range.top, false, false, sheet), fallbackSheet);
  if (range.rowCount === 1 && range.columnCount === 1) return start;
  const end = odfCellAddress(new CellAddress(range.right, range.bottom, false, false, sheet), fallbackSheet);
  return `${start}:${end}`;
}

export function odfCellAddress(address: CellAddress, fallbackSheet: string): string {
  const sheet = address.sheet ?? fallbackSheet;
  return `$${odfSheetName(sheet)}.$${address.columnLabel}$${address.rowLabel}`;
}

/**
 * A sheet name as an ODF address writes it: bare when it is a plain identifier, otherwise in single
 * quotes with any quote doubled (`'My Sheet'`, `'It''s'`). An unquoted `$My Sheet.$A$1` is not an
 * address at all to LibreOffice, so named ranges, validations and filters on it were lost.
 */
export function odfSheetName(sheet: string): string {
  const plain = /^[\p{L}_][\p{L}\p{N}_]*$/u.test(sheet) && !/^[A-Za-z]{1,3}\d+$/.test(sheet);
  return plain ? sheet : `'${sheet.replaceAll("'", "''")}'`;
}

export function pxToCm(px: number): number {
  return (px / PX_PER_INCH) * CM_PER_INCH;
}

export function cmToPx(cm: number): number {
  return (cm / CM_PER_INCH) * PX_PER_INCH;
}

/** .NET's `ToString("0.###", Invariant)`. */
export function formatUpTo3(value: number): string {
  return String(Number.parseFloat(value.toFixed(3)));
}
