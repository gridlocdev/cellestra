import {
  Argb,
  BorderEdges,
  Borders,
  CellAddress,
  Defaults,
  Fonts,
  NumberFormats,
  NumberFormatter,
  SheetDimensions,
  type BorderEdge,
  type CellRef,
  type CellStyle,
  type DataValidation,
  type NamedRange,
  type Table,
  type Workbook,
  type Worksheet,
} from '@cellestra/core';
import type { Console } from './io.ts';
import { toJson } from './json.ts';

/**
 * One cell as a script or an agent wants to see it: the address, the typed value, the display
 * string the GUI would show, and — only when present — the formula source, the error token, and
 * the number-format code. Nulls are omitted on the wire, so a plain number is three keys.
 */
export interface CellReport {
  readonly cell: string;
  readonly value: number | string | boolean | null;
  readonly display: string;
  readonly formula: string | null;
  readonly error: string | null;
  readonly format: string | null;
  readonly style?: Record<string, unknown> | null;
}

export type StyleReport = Record<string, unknown>;

/**
 * The version of the JSON shapes this build speaks. Bump on a breaking change to any reply or to
 * the `apply` document; additive keys don't bump it. Scripts read it from `man --json`.
 */
export const CONTRACT_VERSION = 1;

/** Serialize a payload to stdout as one indented JSON document. */
export function writeReport(con: Console, payload: unknown): void {
  con.outLine(toJson(payload));
}

export function address(sheet: string, row: number, column: number): string {
  return new CellAddress(column, row, false, false, sheet).toString();
}

/** Describe one cell; an empty position still reports (value null, display ""). */
export function cellReport(
  workbook: Workbook,
  sheet: Worksheet,
  row: number,
  column: number,
  locale: string,
): CellReport {
  const cell = sheet.find(row, column);
  const where = address(sheet.name, row, column);
  if (!cell) return { cell: where, value: null, display: '', formula: null, error: null, format: null };
  const style = workbook.styles.get(cell.styleId);
  const format = NumberFormats.isGeneral(style.numberFormat) ? null : style.numberFormat.code;
  const value = cell.value;
  const typed =
    value.kind === 'number' || value.kind === 'text' || value.kind === 'boolean' ? value.value : null;
  const display =
    value.kind === 'empty' ? '' : NumberFormatter.format(value, style.numberFormat.code, locale);
  return {
    cell: where,
    value: typed,
    display,
    formula: cell.formula !== null ? '=' + cell.formula : null,
    error: value.kind === 'error' ? value.error.token : null,
    format,
    style: styleReport(style),
  };
}

export function cellReportAt(workbook: Workbook, ref: CellRef, locale: string): CellReport {
  const sheet = workbook.findSheet(ref.sheet);
  if (!sheet) throw new Error(`Sheet '${ref.sheet}' not found.`);
  return cellReport(workbook, sheet, ref.row, ref.column, locale);
}

const edgeName = (e: BorderEdge): string => e.style;
const sameEdge = (a: BorderEdge, b: BorderEdge): boolean =>
  a.style === b.style && a.color >>> 0 === b.color >>> 0;

/**
 * A cell's formatting as the keys `apply`'s `style` op takes — only what differs from the default,
 * so a plain cell has no `style` at all and a bold header is `{"bold": true}`.
 */
export function styleReport(style: CellStyle): StyleReport | null {
  const d: StyleReport = {};
  const font = style.font;
  const basis = Fonts.default;
  if (font.bold) d['bold'] = true;
  if (font.italic) d['italic'] = true;
  if (font.underline) d['underline'] = true;
  if (font.strikethrough) d['strikethrough'] = true;
  if (font.family !== basis.family) d['font'] = font.family;
  if (font.sizePoints !== basis.sizePoints) d['size'] = font.sizePoints;
  if (font.color >>> 0 !== basis.color >>> 0 && font.color !== 0) d['color'] = Argb.toHex(font.color);
  if (style.fill.background !== null) d['fill'] = Argb.toHex(style.fill.background);
  if (style.alignment.horizontal !== 'general') d['align'] = style.alignment.horizontal;
  if (style.alignment.vertical !== 'bottom') d['valign'] = style.alignment.vertical;
  if (style.alignment.wrapText) d['wrap'] = true;
  if (style.checkbox) d['checkbox'] = true;
  if (Borders.hasAnyEdge(style.border)) {
    const b = style.border;
    const uniform = sameEdge(b.left, b.right) && sameEdge(b.right, b.top) && sameEdge(b.top, b.bottom);
    d['border'] = uniform
      ? edgeName(b.left)
      : {
          left: edgeName(b.left),
          right: edgeName(b.right),
          top: edgeName(b.top),
          bottom: edgeName(b.bottom),
        };
    const visible = [b.left, b.right, b.top, b.bottom].find(BorderEdges.isVisible);
    const color = visible?.color ?? Argb.black;
    if (color >>> 0 !== Argb.black >>> 0) d['borderColor'] = Argb.toHex(color);
  }
  return Object.keys(d).length === 0 ? null : d;
}

/** Every cell in the workbook whose value is an error, in sheet/row/column order. */
export function errorReports(workbook: Workbook, locale: string): CellReport[] {
  const errors: CellReport[] = [];
  for (const sheet of workbook.worksheets) {
    const cells = [...sheet.enumerateCells()].filter(([, , cell]) => cell.value.kind === 'error');
    cells.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
    for (const [row, column] of cells) errors.push(cellReport(workbook, sheet, row, column, locale));
  }
  return errors;
}

export function sheetReport(sheet: Worksheet): Record<string, unknown> {
  const used = sheet.getUsedRange();
  return {
    name: sheet.name,
    usedRange: used?.toString() ?? null,
    rows: used?.rowCount ?? 0,
    columns: used?.columnCount ?? 0,
    cells: sheet.cellCount,
    chart: sheet.isChartSheet ? true : null,
    frozenRows: sheet.frozenRows > 0 ? sheet.frozenRows : null,
    frozenColumns: sheet.frozenColumns > 0 ? sheet.frozenColumns : null,
    tables: sheet.tables.length > 0 ? sheet.tables.map(tableReport) : null,
    merges: sheet.merges.length > 0 ? sheet.merges.map((m) => m.range.toString()) : null,
    widths: layout(
      sheet.columns,
      (c) => c.width !== Defaults.columnWidth,
      (c) => c.width,
      (i) => SheetDimensions.columnToLetters(i),
    ),
    heights: layout(
      sheet.rows,
      (r) => r.height !== Defaults.rowHeight,
      (r) => r.height,
      (i) => String(i + 1),
    ),
    hiddenColumns: hidden(sheet.columns, (i) => SheetDimensions.columnToLetters(i)),
    hiddenRows: hidden(sheet.rows, (i) => String(i + 1)),
    validations: sheet.validations.length > 0 ? sheet.validations.map(validationReport) : null,
  };
}

export function tableReport(t: Table): Record<string, unknown> {
  return {
    name: t.name,
    range: t.range.toString(),
    header: t.hasHeader,
    headerColumn: t.hasHeaderColumn ? true : null,
    bandedRows: t.hasBandedRows ? true : null,
  };
}

export function validationReport(v: DataValidation): Record<string, unknown> {
  const listLike = v.kind === 'list' || v.kind === 'any';
  const twoBounds = v.operator === 'between' || v.operator === 'notBetween';
  return {
    range: v.target.toString(),
    kind: v.kind.toLowerCase(),
    operator: listLike ? null : v.operator.toLowerCase(),
    value1: listLike ? null : v.operand1,
    value2: twoBounds && !listLike ? v.operand2 : null,
    list: v.listValues ?? null,
    source: v.listSource?.toString() ?? null,
    response: v.response === 'stop' ? null : v.response,
  };
}

export function namedRangeReport(named: NamedRange): Record<string, unknown> {
  return { name: named.name, target: named.target.toString(), scope: named.scope ?? 'workbook' };
}

function layout<T>(
  lines: ReadonlyMap<number, T>,
  differs: (line: T) => boolean,
  size: (line: T) => number,
  label: (index: number) => string,
): Record<string, number> | null {
  const entries = [...lines].filter(([, v]) => differs(v)).sort((a, b) => a[0] - b[0]);
  if (entries.length === 0) return null;
  // ODS stores cm; don't echo the conversion noise.
  return Object.fromEntries(entries.map(([k, v]) => [label(k), Math.round(size(v) * 10) / 10]));
}

function hidden<T extends { readonly hidden: boolean }>(
  lines: ReadonlyMap<number, T>,
  label: (index: number) => string,
): string[] | null {
  const list = [...lines]
    .filter(([, v]) => v.hidden)
    .sort((a, b) => a[0] - b[0])
    .map(([k]) => label(k));
  return list.length === 0 ? null : list;
}
