import { SheetDimensions } from '../model/sheetDimensions.ts';
import type { Worksheet } from '../model/worksheet.ts';

/** An inclusive run of line indices. */
export type LineRun = readonly [first: number, last: number];

/**
 * Which hidden rows/columns an "Unhide" over a selection should reveal: the hidden lines inside it,
 * extended outward through hidden neighbours (so selecting B:D reveals a hidden A..E run whole).
 * Rows a filter hid are the filter's: the reach passes through them, but only rows hidden by hand
 * are revealed (clearing the filter shows the others).
 */
export const HiddenLines = {
  rowsToUnhide: (sheet: Worksheet, top: number, bottom: number): LineRun[] =>
    runs(
      keys(sheet.rows, (r) => r.hidden),
      keys(sheet.rows, (r) => r.manuallyHidden),
      top,
      bottom,
      SheetDimensions.maxRows,
    ),

  columnsToUnhide: (sheet: Worksheet, left: number, right: number): LineRun[] => {
    const hidden = keys(sheet.columns, (c) => c.hidden);
    return runs(hidden, hidden, left, right, SheetDimensions.maxColumns);
  },
} as const;

function keys<T>(lines: ReadonlyMap<number, T>, test: (line: T) => boolean): Set<number> {
  const set = new Set<number>();
  for (const [index, line] of lines) if (test(line)) set.add(index);
  return set;
}

/** The `revealed` lines in the selection widened through the `hidden` lines touching its edges. */
function runs(
  hidden: ReadonlySet<number>,
  revealed: ReadonlySet<number>,
  first: number,
  last: number,
  max: number,
): LineRun[] {
  if (revealed.size === 0) return [];
  // Reach outward only while the next line is hidden: bounded by the hidden count, not the distance.
  while (first > 0 && hidden.has(first - 1)) first--;
  while (last < max - 1 && hidden.has(last + 1)) last++;
  const out: [number, number][] = [];
  for (const line of [...revealed].filter((l) => l >= first && l <= last).sort((a, b) => a - b)) {
    const prev = out.at(-1);
    if (prev && prev[1] === line - 1) prev[1] = line;
    else out.push([line, line]);
  }
  return out;
}
