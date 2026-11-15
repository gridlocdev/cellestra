import type { ChartAggregate } from '@cellestra/charts';

/** One series' values as read from the sheet: the numbers (`NaN` for a gap) and which cells hold anything. */
export interface GroupingSeries {
  readonly values: readonly number[];
  /** Whether each cell holds a value of any kind, for `count`. */
  readonly filled: readonly boolean[];
}

/** Categories after grouping, with each series' values combined to match. */
export interface GroupedData {
  readonly categories: readonly string[];
  readonly values: readonly (readonly number[])[];
  /**
   * How many source rows each category merged, set only when grouping merged or counted rows; a
   * chart left as it is has none (every category is one row).
   */
  readonly rowCounts?: readonly number[];
}

function combine(aggregate: ChartAggregate, rows: readonly number[], series: GroupingSeries): number {
  if (aggregate === 'count') return rows.filter((r) => series.filled[r] === true).length;
  const numbers = rows.map((r) => series.values[r] ?? NaN).filter((n) => !Number.isNaN(n));
  if (numbers.length === 0) return NaN;
  switch (aggregate) {
    case 'sum':
      return numbers.reduce((a, b) => a + b, 0);
    case 'average':
      return numbers.reduce((a, b) => a + b, 0) / numbers.length;
    case 'min':
      return Math.min(...numbers);
    case 'max':
      return Math.max(...numbers);
  }
}

/** The category shown for rows whose label is blank. */
export const BLANK_CATEGORY = '(blank)';

/** Whether no series has anything at `row`: every value a gap, and (for count) no cell filled. */
function rowIsEmpty(series: readonly GroupingSeries[], row: number, aggregate: ChartAggregate): boolean {
  return series.every(
    (s) => Number.isNaN(s.values[row] ?? NaN) && (aggregate !== 'count' || s.filled[row] !== true),
  );
}

/**
 * Merging rows that share a category label. Labels are trimmed, so " North" and "North" are one
 * category. Rows group by their label in the order each label first appears. A row with a blank
 * (empty or whitespace) label and nothing in any series, such as an empty row at the bottom of a
 * table, is left out; the other blank-label rows merge into one "(blank)" category. Each series
 * combines its rows with the chart's aggregate: sum, average, min and max skip gaps (a group with
 * no numbers stays a gap), and count counts the non-empty cells. A chart with no blank, repeated
 * or padded labels is left as it is, unless it counts.
 */
export const ChartGrouping = {
  group(
    categories: readonly string[],
    series: readonly GroupingSeries[],
    aggregate: ChartAggregate,
  ): GroupedData {
    const groups = new Map<string, number[]>();
    const order: number[][] = [];
    const labels: string[] = [];
    let changed = aggregate === 'count';
    const longest = Math.max(categories.length, ...series.map((s) => s.values.length));
    for (let row = 0; row < longest; row++) {
      const raw = categories[row] ?? '';
      const trimmed = raw.trim();
      // Rows past the last label (a values range longer than the categories) count as blank.
      const label = trimmed === '' ? BLANK_CATEGORY : trimmed;
      if (trimmed !== raw || row >= categories.length) changed = true;
      if (trimmed === '' && rowIsEmpty(series, row, aggregate)) {
        changed = true;
        continue;
      }
      const existing = groups.get(label);
      if (existing) {
        existing.push(row);
        changed = true;
        continue;
      }
      const rows = [row];
      groups.set(label, rows);
      order.push(rows);
      labels.push(label);
    }
    if (!changed) return { categories, values: series.map((s) => s.values) };
    return {
      categories: labels,
      values: series.map((s) => order.map((rows) => combine(aggregate, rows, s))),
      rowCounts: order.map((rows) => rows.length),
    };
  },
} as const;
