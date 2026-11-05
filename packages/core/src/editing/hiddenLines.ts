import { SheetDimensions } from '../model/sheetDimensions.ts';
import type { Worksheet } from '../model/worksheet.ts';

/** An inclusive run of line indices. */
export type LineRun = readonly [first: number, last: number];

/**
 * Which hidden rows/columns an "Unhide" over a selection should reveal: the hidden lines inside it,
 * extended outward through hidden neighbours (so selecting B:D reveals a hidden A..E run whole).
 */
export const HiddenLines = {
  rowsToUnhide: (sheet: Worksheet, top: number, bottom: number): LineRun[] =>
    runs(hiddenKeys(sheet.rows), top, bottom, SheetDimensions.maxRows),

  columnsToUnhide: (sheet: Worksheet, left: number, right: number): LineRun[] =>
    runs(hiddenKeys(sheet.columns), left, right, SheetDimensions.maxColumns),
} as const;

function hiddenKeys(lines: ReadonlyMap<number, { readonly hidden: boolean }>): Set<number> {
  const set = new Set<number>();
  for (const [index, line] of lines) if (line.hidden) set.add(index);
  return set;
}

function runs(hidden: ReadonlySet<number>, first: number, last: number, max: number): LineRun[] {
  if (hidden.size === 0) return [];
  // Reach outward only while the next line is hidden: bounded by the hidden count, not the distance.
  while (first > 0 && hidden.has(first - 1)) first--;
  while (last < max - 1 && hidden.has(last + 1)) last++;
  const out: [number, number][] = [];
  for (const line of [...hidden].filter((l) => l >= first && l <= last).sort((a, b) => a - b)) {
    const prev = out.at(-1);
    if (prev && prev[1] === line - 1) prev[1] = line;
    else out.push([line, line]);
  }
  return out;
}
