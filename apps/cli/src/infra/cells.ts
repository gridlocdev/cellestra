import { CellAddress, CellRange, NumberFormatter, type Workbook, type Worksheet } from '@cellestra/core';
import { CliError } from './errors.ts';

/** Addressing helpers shared by every command: sheets, cells, ranges, names, and tables. */

export function pickSheet(workbook: Workbook, name: string | null | undefined): Worksheet {
  if (!name) return workbook.activeSheet;
  const sheet = workbook.findSheet(name);
  if (!sheet) throw new CliError(`No sheet named '${name}'. Available: ${sheetList(workbook)}.`);
  return sheet;
}

export interface ResolvedCell {
  readonly sheet: Worksheet;
  readonly address: CellAddress;
}

export interface ResolvedRange {
  readonly sheet: Worksheet;
  readonly range: CellRange;
}

export function resolveCell(
  workbook: Workbook,
  reference: string,
  defaultSheet: string | null = null,
): ResolvedCell {
  const address = CellAddress.tryParse(reference);
  if (!address) {
    const named = resolveName(workbook, reference, defaultSheet);
    if (named) {
      if (named.range.rowCount !== 1 || named.range.columnCount !== 1) {
        throw new CliError(`'${reference}' is a range (${named.range.toString()}), not a single cell.`);
      }
      return { sheet: named.sheet, address: named.range.topLeft };
    }
    throw new CliError(
      `'${reference}' is not a valid cell reference (e.g. Sheet1.B2 or B2) or a known name.`,
    );
  }
  const sheet =
    address.sheet !== null ? requireSheet(workbook, address.sheet) : pickSheet(workbook, defaultSheet);
  return { sheet, address };
}

export function resolveRange(
  workbook: Workbook,
  reference: string,
  defaultSheet: string | null = null,
): ResolvedRange {
  const range = CellRange.tryParse(reference);
  if (!range) {
    const named = resolveName(workbook, reference, defaultSheet);
    if (named) return named;
    throw new CliError(
      `'${reference}' is not a valid range (e.g. Sheet1.A1:C10 or A1:C10), a named range, or a table.`,
    );
  }
  const sheet =
    range.sheet !== null ? requireSheet(workbook, range.sheet) : pickSheet(workbook, defaultSheet);
  return { sheet, range };
}

function requireSheet(workbook: Workbook, name: string): Worksheet {
  const sheet = workbook.findSheet(name);
  if (!sheet) throw new CliError(`No sheet named '${name}'.`);
  return sheet;
}

/** A named range (scoped to the default sheet first) or a table name, as a sheet plus a bare range. */
export function resolveName(
  workbook: Workbook,
  rawName: string,
  defaultSheet: string | null,
): ResolvedRange | null {
  const name = rawName.trim();
  if (name.length === 0) return null;
  const home = !defaultSheet ? workbook.activeSheet : workbook.findSheet(defaultSheet);
  const named = workbook.findNamedRange(name, home?.name ?? null);
  if (named) {
    const target = named.target;
    let sheet: Worksheet;
    if (target.sheet !== null) {
      const found = workbook.findSheet(target.sheet);
      if (!found) throw new CliError(`Named range '${name}' points at a missing sheet '${target.sheet}'.`);
      sheet = found;
    } else {
      sheet = home ?? workbook.activeSheet;
    }
    return { sheet, range: target.withSheet(null) };
  }
  // Tables: the default sheet first, then any sheet — a table name is unique within its sheet.
  const candidates = home ? [home, ...workbook.worksheets.filter((w) => w !== home)] : workbook.worksheets;
  const lower = name.toLowerCase();
  for (const sheet of candidates) {
    const table = sheet.tables.find((t) => t.name.toLowerCase() === lower);
    if (table) return { sheet, range: table.range.withSheet(null) };
  }
  return null;
}

export function sheetList(workbook: Workbook): string {
  return workbook.worksheets.map((s) => s.name).join(', ');
}

/** The string the app shows in the cell ("" for an empty one). */
export function display(
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

export function formulaOrValue(
  workbook: Workbook,
  sheet: Worksheet,
  row: number,
  column: number,
  locale: string,
): string {
  const cell = sheet.find(row, column);
  return cell && cell.formula !== null ? '=' + cell.formula : display(workbook, sheet, row, column, locale);
}

export function formulaAndValue(
  workbook: Workbook,
  sheet: Worksheet,
  row: number,
  column: number,
  locale: string,
): string {
  const cell = sheet.find(row, column);
  return cell && cell.formula !== null
    ? `=${cell.formula} → ${display(workbook, sheet, row, column, locale)}`
    : display(workbook, sheet, row, column, locale);
}

export function buildGrid(
  workbook: Workbook,
  sheet: Worksheet,
  range: CellRange,
  locale: string,
  formulas = false,
  both = false,
): string[][] {
  const rows: string[][] = [];
  for (let r = 0; r < range.rowCount; r++) {
    const cols: string[] = [];
    for (let c = 0; c < range.columnCount; c++) {
      const row = range.top + r;
      const column = range.left + c;
      cols.push(
        both
          ? formulaAndValue(workbook, sheet, row, column, locale)
          : formulas
            ? formulaOrValue(workbook, sheet, row, column, locale)
            : display(workbook, sheet, row, column, locale),
      );
    }
    rows.push(cols);
  }
  return rows;
}

export function usedRangeOrA1(sheet: Worksheet): CellRange {
  return sheet.getUsedRange() ?? new CellRange(new CellAddress(0, 0));
}
