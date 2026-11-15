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

/**
 * Merging rows that share a category label. Rows group by their exact label in the order each label
 * first appears; a blank label never merges. Each series combines its rows with the chart's
 * aggregate: sum, average, min and max skip gaps (a group with no numbers stays a gap), and count
 * counts the non-empty cells. A chart whose labels are all different is left as it is, unless it
 * counts.
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
    categories.forEach((label, row) => {
      const existing = label === '' ? undefined : groups.get(label);
      if (existing) {
        existing.push(row);
        return;
      }
      const rows = [row];
      if (label !== '') groups.set(label, rows);
      order.push(rows);
      labels.push(label);
    });
    if (order.length === categories.length && aggregate !== 'count')
      return { categories, values: series.map((s) => s.values) };
    // Rows past the last label (a values range longer than the categories) stay rows of their own.
    const longest = Math.max(0, ...series.map((s) => s.values.length));
    for (let row = categories.length; row < longest; row++) {
      order.push([row]);
      labels.push('');
    }
    return {
      categories: labels,
      values: series.map((s) => order.map((rows) => combine(aggregate, rows, s))),
    };
  },
} as const;
