import { SheetDimensions, type SortLevel, type Worksheet } from '@cellestra/core';
import { CliError } from '../infra/errors.ts';

/** Row, column, and span spellings shared by the verbs and the `apply` ops. */

const INTEGER = /^[+-]?\d+$/;

/** A column reference — a letter (A, AB) or a 1-based number — as a 0-based index. */
export function parseColumn(raw: string): number {
  const token = raw.trim();
  if (INTEGER.test(token)) {
    const oneBased = Number(token);
    if (oneBased < 1) throw new CliError(`Column number '${token}' must be 1 or greater.`);
    return oneBased - 1;
  }
  const column = SheetDimensions.lettersToColumn(token);
  if (column !== null) return column;
  throw new CliError(`'${token}' is not a valid column (use a letter like B or a 1-based number).`);
}

/** A 1-based row number as a 0-based index. */
export function parseRow(token: string): number {
  const trimmed = token.trim();
  const oneBased = INTEGER.test(trimmed) ? Number(trimmed) : Number.NaN;
  if (!Number.isFinite(oneBased) || oneBased < 1) {
    throw new CliError(`'${token}' is not a valid row number (1 or greater).`);
  }
  return oneBased - 1;
}

export function parseRowSpan(token: string): [first: number, last: number] {
  const [a, b] = splitSpan(token);
  const first = parseRow(a);
  return order(first, b === null ? first : parseRow(b));
}

export function parseColumnSpan(token: string): [first: number, last: number] {
  const [a, b] = splitSpan(token);
  const first = parseColumn(a);
  return order(first, b === null ? first : parseColumn(b));
}

function splitSpan(token: string): [string, string | null] {
  const colon = token.indexOf(':');
  if (colon < 0) return [token.trim(), null];
  return [token.slice(0, colon).trim(), token.slice(colon + 1).trim()];
}

function order(a: number, b: number): [number, number] {
  return a <= b ? [a, b] : [b, a];
}

/** `B,A:desc` → sort levels. */
export function parseSortLevels(spec: string): SortLevel[] {
  const levels: SortLevel[] = [];
  for (const token of splitList(spec)) {
    const colon = token.indexOf(':');
    const column = colon < 0 ? token : token.slice(0, colon);
    const direction =
      colon < 0
        ? null
        : token
            .slice(colon + 1)
            .trim()
            .toLowerCase();
    const ascending = direction === null || !['desc', 'descending', 'd'].includes(direction);
    levels.push({ column: parseColumn(column.trim()), ascending });
  }
  if (levels.length === 0)
    throw new CliError('--by needs at least one column, e.g. --by B or --by B,A:desc.');
  return levels;
}

/** Split on commas, trimming each entry and dropping empty ones. */
export function splitList(text: string): string[] {
  return text
    .split(',')
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
}

/**
 * Refuse an insert or delete the sheet can't hold: a count past the sheet's edge, or an insert that
 * would push stored cells off the last row or column (the engine would drop them silently).
 */
export function checkStructuralEdit(
  sheet: Worksheet,
  axis: 'row' | 'column',
  operation: 'insert' | 'delete',
  index: number,
  count: number,
): void {
  const isRow = axis === 'row';
  const size = isRow ? SheetDimensions.maxRows : SheetDimensions.maxColumns;
  const noun = isRow ? 'row' : 'column';
  const label = (i: number): string => (isRow ? String(i + 1) : SheetDimensions.columnToLetters(i));
  if (index >= size)
    throw new CliError(`${noun} ${index + 1} is past the sheet's last ${noun} (${label(size - 1)}).`);
  const room = size - index;
  if (count > room) {
    throw new CliError(
      `count ${count} reaches past the sheet's last ${noun} (${label(size - 1)}); at most ${room} fit from ${noun} ${label(index)}.`,
    );
  }
  if (operation !== 'insert') return;
  const extent = sheet.cellExtent();
  const last = extent === null ? -1 : isRow ? extent.lastRow : extent.lastColumn;
  if (last >= index && last + count >= size) {
    throw new CliError(
      `inserting ${count} ${noun}(s) at ${label(index)} would push data in ${noun} ${label(last)} off the sheet ` +
        `(the last ${noun} is ${label(size - 1)}); clear or move it first.`,
    );
  }
}
