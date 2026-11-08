import type { CellRef, RangeRef } from './cellRef.ts';

/** Large ranges wider than this many columns are indexed under `WIDE` rather than per column. */
const WIDE_RANGE_COLUMNS = 64;
const WIDE = -1;

/** The index buckets a large range sits in: each of its columns, or `WIDE`. */
function indexColumns(r: RangeRef): number[] {
  if (r.right - r.left + 1 > WIDE_RANGE_COLUMNS) return [WIDE];
  const out: number[] = [];
  for (let c = r.left; c <= r.right; c++) out.push(c);
  return out;
}

interface Built {
  readonly tops: Int32Array;
  readonly lefts: Int32Array;
  readonly rights: Int32Array;
  readonly keys: string[];
  /** Max bottom row per tree node; leaves at `size + i`; -1 for an empty leaf. */
  readonly tree: Int32Array;
  readonly size: number;
}

/** Per-search state for `LargeRangeIndex.forEachOver`: the ranges already visited, per bucket. */
export type RangeSearchSession = Map<unknown, Int32Array>;

/**
 * The large ranges over one column (or the wide ones of a sheet), searchable by row: the entries
 * sorted by top row, with a max-heap-shaped tree of their bottom rows, so the ranges over a row are
 * found in time proportional to how many there are, not to how many the column holds.
 */
class Bucket {
  /** Each formula's ranges in this bucket. */
  readonly ranges = new Map<string, RangeRef[]>();
  #built: Built | null = null;

  invalidate(): void {
    this.#built = null;
  }

  #build(): Built {
    if (this.#built) return this.#built;
    const entries: { key: string; range: RangeRef }[] = [];
    for (const [key, ranges] of this.ranges) for (const range of ranges) entries.push({ key, range });
    entries.sort((a, b) => a.range.top - b.range.top);
    let size = 1;
    while (size < entries.length) size *= 2;
    const tree = new Int32Array(2 * size).fill(-1);
    const tops = new Int32Array(entries.length);
    const lefts = new Int32Array(entries.length);
    const rights = new Int32Array(entries.length);
    const keys: string[] = [];
    entries.forEach(({ key, range }, i) => {
      tops[i] = range.top;
      lefts[i] = range.left;
      rights[i] = range.right;
      keys.push(key);
      tree[size + i] = range.bottom;
    });
    for (let i = size - 1; i >= 1; i--) tree[i] = Math.max(tree[2 * i]!, tree[2 * i + 1]!);
    return (this.#built = { tops, lefts, rights, keys, tree, size });
  }

  /**
   * Visit the formula key of every range here containing (`row`, `column`). With `claimed` (one
   * per search session), a range already visited in the session is skipped — so a breadth-first
   * walk visits each range at most once however many cells it asks about.
   */
  forEachOver(
    row: number,
    column: number,
    visit: (formulaKey: string) => void,
    claimed?: RangeSearchSession,
  ): void {
    const built = this.#build();
    if (built.keys.length === 0) return;
    let tree = built.tree;
    if (claimed) {
      let own = claimed.get(this);
      if (!own) claimed.set(this, (own = built.tree.slice()));
      tree = own;
    }
    // Only the entries with top <= row can contain it: a prefix of the sorted entries.
    let lo = 0;
    let hi = built.tops.length;
    while (lo < hi) {
      const mid = (lo + hi) >>> 1;
      if (built.tops[mid]! <= row) lo = mid + 1;
      else hi = mid;
    }
    const limit = lo;
    if (limit === 0) return;
    const { size } = built;
    const descend = (node: number, nodeLo: number, nodeHi: number): void => {
      if (nodeLo >= limit || tree[node]! < row) return;
      if (node >= size) {
        const i = node - size;
        if (column < built.lefts[i]! || column > built.rights[i]!) return; // a wide range, other columns
        if (claimed) {
          tree[node] = -1;
          for (let p = node >> 1; p >= 1; p >>= 1) tree[p] = Math.max(tree[2 * p]!, tree[2 * p + 1]!);
        }
        visit(built.keys[i]!);
        return;
      }
      const mid = (nodeLo + nodeHi) >>> 1;
      descend(2 * node, nodeLo, mid);
      descend(2 * node + 1, mid, nodeHi);
    };
    descend(1, 0, size);
  }
}

/**
 * The large precedent ranges of every formula, indexed by sheet and then by each column they cover
 * (ranges wider than `WIDE_RANGE_COLUMNS` sit under one `WIDE` bucket instead), so the formulas
 * reading one cell through a large range are found without scanning every large-range formula.
 */
export class LargeRangeIndex {
  readonly #byFormula = new Map<string, RangeRef[]>();
  readonly #buckets = new Map<string, Map<number, Bucket>>();

  clear(): void {
    this.#byFormula.clear();
    this.#buckets.clear();
  }

  /** Every formula with a large range, and those ranges. */
  entries(): Iterable<[string, readonly RangeRef[]]> {
    return this.#byFormula.entries();
  }

  /** Replace the large ranges of the formula at `key` (an empty list removes it). */
  set(key: string, ranges: readonly RangeRef[]): void {
    for (const r of this.#byFormula.get(key) ?? []) {
      const columns = this.#buckets.get(r.sheet.toLowerCase());
      if (!columns) continue;
      for (const column of indexColumns(r)) {
        const bucket = columns.get(column);
        if (!bucket) continue;
        bucket.ranges.delete(key);
        bucket.invalidate();
        if (bucket.ranges.size === 0) columns.delete(column);
      }
    }
    if (ranges.length === 0) {
      this.#byFormula.delete(key);
      return;
    }
    this.#byFormula.set(key, [...ranges]);
    for (const r of ranges) {
      const sheet = r.sheet.toLowerCase();
      let columns = this.#buckets.get(sheet);
      if (!columns) this.#buckets.set(sheet, (columns = new Map()));
      for (const column of indexColumns(r)) {
        let bucket = columns.get(column);
        if (!bucket) columns.set(column, (bucket = new Bucket()));
        let list = bucket.ranges.get(key);
        if (!list) bucket.ranges.set(key, (list = []));
        list.push(r);
        bucket.invalidate();
      }
    }
  }

  /**
   * Call `visit` with the key of every formula that has a large range containing `cell` (once per
   * containing range). Pass the same `session` map across calls to visit each range only once.
   */
  forEachOver(cell: CellRef, visit: (formulaKey: string) => void, session?: RangeSearchSession): void {
    const columns = this.#buckets.get(cell.sheet.toLowerCase());
    if (!columns) return;
    columns.get(cell.column)?.forEachOver(cell.row, cell.column, visit, session);
    columns.get(WIDE)?.forEachOver(cell.row, cell.column, visit, session);
  }

  /** A fresh search session for `forEachOver`. */
  static session(): RangeSearchSession {
    return new Map();
  }
}
