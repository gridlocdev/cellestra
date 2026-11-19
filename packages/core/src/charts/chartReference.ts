import { findSheetSeparator } from '../model/cellAddress.ts';
import { CellRange } from '../model/cellRange.ts';
import { SheetDimensions } from '../model/sheetDimensions.ts';

const COLUMN = /^(\$?)([A-Za-z]+)$/;
const ROW = /^(\$?)([0-9]+)$/;

/** `Data.B` → `Data.` and `B`: the sheet prefix (with its dot) and the rest. */
function splitSheet(text: string): [string, string] {
  const dot = findSheetSeparator(text);
  return dot >= 0 ? [text.slice(0, dot + 1), text.slice(dot + 1)] : ['', text];
}

/**
 * A chart reference with whole columns or rows written out as cells, so every reader of the file
 * understands it: `Data.A:C` → `Data.A1:C1048576`, `Data.2:3` → `Data.A2:XFD3`. Anything else
 * comes back trimmed, as typed.
 */
export function expandWholeLines(text: string): string {
  const t = text.trim();
  const colon = t.indexOf(':');
  if (colon < 0) return t;
  const [prefix, start] = splitSheet(t.slice(0, colon));
  const [endPrefix, end] = splitSheet(t.slice(colon + 1));
  const lastRow = SheetDimensions.maxRows;
  const a = COLUMN.exec(start);
  const b = COLUMN.exec(end);
  if (a && b) {
    if (SheetDimensions.lettersToColumn(a[2]!) === null || SheetDimensions.lettersToColumn(b[2]!) === null)
      return t;
    return `${prefix}${a[1]}${a[2]}${a[1]}1:${endPrefix}${b[1]}${b[2]}${b[1]}${lastRow}`;
  }
  const r = ROW.exec(start);
  const s = ROW.exec(end);
  if (r && s) {
    const valid = (row: string) => SheetDimensions.isValidRow(Number.parseInt(row, 10) - 1);
    if (!valid(r[2]!) || !valid(s[2]!)) return t;
    const lastColumn = SheetDimensions.columnToLetters(SheetDimensions.maxColumns - 1);
    return `${prefix}${r[1]}A${r[1]}${r[2]}:${endPrefix}${s[1]}${lastColumn}${s[1]}${s[2]}`;
  }
  return t;
}

/** The cell range a chart reference names, whole columns and rows included; null when it names none. */
export function parseChartRange(text: string | null | undefined): CellRange | null {
  return text ? CellRange.tryParse(expandWholeLines(text)) : null;
}
