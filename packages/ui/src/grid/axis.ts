/**
 * One grid dimension (columns or rows) as positions in unzoomed CSS pixels. Most lines keep the
 * default size, so the axis stores only the exceptions (custom sizes, hidden = 0) sorted by index,
 * with prefix sums of their difference from the default: offset and hit-test are O(log n) without
 * a per-line array over 1,048,576 rows.
 */
export class Axis {
  readonly count: number;
  readonly defaultSize: number;
  private readonly indexes: Int32Array;
  private readonly sizes: Float64Array;
  /** deltaBefore[k] = Σ (sizes[j] − default) for j < k. */
  private readonly deltaBefore: Float64Array;

  constructor(
    count: number,
    defaultSize: number,
    exceptions: Iterable<readonly [index: number, size: number]>,
  ) {
    this.count = count;
    this.defaultSize = defaultSize;
    const list = [...exceptions]
      .filter(([i, s]) => i >= 0 && i < count && s !== defaultSize)
      .sort((a, b) => a[0] - b[0]);
    this.indexes = Int32Array.from(list, ([i]) => i);
    this.sizes = Float64Array.from(list, ([, s]) => Math.max(0, s));
    this.deltaBefore = new Float64Array(list.length + 1);
    for (let k = 0; k < list.length; k++)
      this.deltaBefore[k + 1] = this.deltaBefore[k]! + this.sizes[k]! - defaultSize;
  }

  /** Number of exceptions with index < i. */
  private rank(i: number): number {
    let lo = 0;
    let hi = this.indexes.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.indexes[mid]! < i) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }

  size(i: number): number {
    const k = this.rank(i);
    return this.indexes[k] === i ? this.sizes[k]! : this.defaultSize;
  }

  /** Position of line i's leading edge; offset(count) is the total length. */
  offset(i: number): number {
    const clamped = Math.max(0, Math.min(i, this.count));
    return clamped * this.defaultSize + this.deltaBefore[this.rank(clamped)]!;
  }

  get total(): number {
    return this.offset(this.count);
  }

  /** The visible line containing position px (clamped to the axis). Hidden lines are never returned. */
  indexAt(px: number): number {
    if (px <= 0) return this.firstVisibleFrom(0);
    let lo = 0;
    let hi = this.count - 1;
    // Largest i with offset(i) <= px — a zero-size line shares its offset with the next, so the
    // search naturally lands on the visible one after it.
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.offset(mid) <= px) lo = mid;
      else hi = mid - 1;
    }
    return this.size(lo) > 0 ? lo : this.lastVisibleBefore(lo);
  }

  isHidden(i: number): boolean {
    return this.size(i) === 0;
  }

  /** First visible line at or after i (i itself when none: a fully hidden tail). */
  firstVisibleFrom(i: number): number {
    for (let j = i; j < this.count; j++) if (this.size(j) > 0) return j;
    return Math.min(i, this.count - 1);
  }

  lastVisibleBefore(i: number): number {
    for (let j = i; j >= 0; j--) if (this.size(j) > 0) return j;
    return this.firstVisibleFrom(i);
  }

  /** Step from i by delta visible lines, clamped to the axis. */
  step(i: number, delta: number): number {
    let j = i;
    const dir = Math.sign(delta);
    for (let n = Math.abs(delta); n > 0;) {
      const next = j + dir;
      if (next < 0 || next >= this.count) break;
      j = next;
      if (this.size(j) > 0) n--;
    }
    return this.size(j) > 0 ? j : i;
  }
}
