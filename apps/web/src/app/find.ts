import { CellAddress, type Cell, type Worksheet } from '@cellestra/core';

/** A cell Find matched: where it is and what it shows. */
export interface FindMatch {
  readonly row: number;
  readonly column: number;
  /** `B7`. */
  readonly address: string;
  /** What the cell shows, or its formula when only the formula holds the query. */
  readonly text: string;
}

/**
 * Every cell whose shown text or formula contains `query` (ignoring case), in row-major order.
 * `shown` formats a cell's value as the grid draws it. Walks the stored cells rather than every
 * position of the used range: one stray cell at XFD1048576 would make that 17 billion lookups.
 */
export function findMatches(
  sheet: Worksheet,
  query: string,
  shown: (row: number, column: number, cell: Cell) => string,
): FindMatch[] {
  if (!query) return [];
  const q = query.toLowerCase();
  const matches: FindMatch[] = [];
  for (const [row, column, cell] of sheet.enumerateCells()) {
    if (cell.value.kind === 'empty' && cell.formula === null) continue;
    const text = shown(row, column, cell);
    const formula = cell.formula === null ? '' : '=' + cell.formula;
    const inText = text.toLowerCase().includes(q);
    if (!inText && !formula.toLowerCase().includes(q)) continue;
    const address = new CellAddress(column, row).toString();
    matches.push({ row, column, address, text: inText ? text : formula });
  }
  return matches.sort((a, b) => a.row - b.row || a.column - b.column);
}

/**
 * The match after (or before, `backwards`) the cell `at` in row-major order, wrapping around the
 * sheet. Null when there are none.
 */
export function nextMatch(
  matches: readonly FindMatch[],
  at: { readonly row: number; readonly column: number },
  backwards = false,
): FindMatch | null {
  if (matches.length === 0) return null;
  const after = (m: FindMatch) => m.row > at.row || (m.row === at.row && m.column > at.column);
  const before = (m: FindMatch) => m.row < at.row || (m.row === at.row && m.column < at.column);
  return backwards
    ? ([...matches].reverse().find(before) ?? matches[matches.length - 1]!)
    : (matches.find(after) ?? matches[0]!);
}

/** What Replace edits in a cell: its formula (with `=`) or its text, or null for other values. */
export function replaceableText(cell: Cell | undefined): string | null {
  if (!cell) return null;
  if (cell.formula !== null) return '=' + cell.formula;
  return cell.value.kind === 'text' ? cell.value.value : null;
}

/** `text` with every `query` (ignoring case) replaced, or null when it holds none. */
export function replaceIn(text: string, query: string, replacement: string): string | null {
  if (!query) return null;
  const pattern = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
  if (!pattern.test(text)) return null;
  pattern.lastIndex = 0;
  return text.replace(pattern, () => replacement);
}
