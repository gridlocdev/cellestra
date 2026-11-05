/**
 * The single home for new-workbook and cell defaults (PRD §8.8). Defined once so the model,
 * styling, and UI never disagree. Sizes are CSS pixels.
 */
export const Defaults = {
  firstSheetName: 'Sheet1',
  sheetNamePrefix: 'Sheet',
  chartSheetNamePrefix: 'Chart',
  fontFamily: 'Helvetica Neue',
  fontSizePoints: 10,
  /** ≈ 8 characters at the default font. */
  columnWidth: 64,
  rowHeight: 20,
  /** Opaque black, ARGB. */
  textColorArgb: 0xff000000,
  /** Clean Up's header-row fill (#C6C6C6) — matches the grid's light table-header band. */
  tableHeaderFillArgb: 0xffc6c6c6,
  tableHeaderTextArgb: 0xff000000,
} as const;

/** Next unused default sheet name (Sheet2, Sheet3, …), case-insensitive. */
export function nextSheetName(existing: Iterable<string>): string {
  return nextName(Defaults.sheetNamePrefix, existing);
}

/** Next unused chart-sheet name (Chart1, Chart2, …). */
export function nextChartSheetName(existing: Iterable<string>): string {
  return nextName(Defaults.chartSheetNamePrefix, existing);
}

function nextName(prefix: string, existing: Iterable<string>): string {
  const taken = new Set([...existing].map((n) => n.toLowerCase()));
  for (let n = 1; ; n++) {
    const candidate = `${prefix}${n}`;
    if (!taken.has(candidate.toLowerCase())) return candidate;
  }
}
