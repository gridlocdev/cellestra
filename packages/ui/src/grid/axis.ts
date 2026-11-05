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
  /**
   * For a hidden exception k, the first line after its run of consecutive hidden lines (runEnd) and
   * the last line before it (runStart): skipping a hidden run is one lookup, not a walk.
   */
  private readonly runEnd: Int32Array;
  private readonly runStart: Int32Array;

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
    const n = list.length;
    this.runEnd = new Int32Array(n);
    this.runStart = new Int32Array(n);
    for (let k = n - 1; k >= 0; k--) {
      if (this.sizes[k]! > 0) continue;
      const nextHiddenAdjacent =
        k + 1 < n && this.sizes[k + 1]! === 0 && this.indexes[k + 1] === this.indexes[k]! + 1;
      this.runEnd[k] = nextHiddenAdjacent ? this.runEnd[k + 1]! : this.indexes[k]! + 1;
    }
    for (let k = 0; k < n; k++) {
      if (this.sizes[k]! > 0) continue;
      const prevHiddenAdjacent =
        k > 0 && this.sizes[k - 1]! === 0 && this.indexes[k - 1] === this.indexes[k]! - 1;
      this.runStart[k] = prevHiddenAdjacent ? this.runStart[k - 1]! : this.indexes[k]! - 1;
    }
  }

  /** The first visible line at or after i, or `count` when there is none. */
  nextVisible(i: number): number {
    if (i < 0) i = 0;
    if (i >= this.count) return this.count;
    const k = this.rank(i);
    if (this.indexes[k] !== i || this.sizes[k]! > 0) return i;
    const after = this.runEnd[k]!;
    // A default-size line always shows; a custom-size exception right after the run shows too.
    return after >= this.count ? this.count : after;
  }

  /** The last visible line at or before i, or −1 when there is none. */
  prevVisible(i: number): number {
    if (i >= this.count) i = this.count - 1;
    if (i < 0) return -1;
    const k = this.rank(i);
    if (this.indexes[k] !== i || this.sizes[k]! > 0) return i;
    return this.runStart[k]!;
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

  /** First visible line at or after i (else the last visible before it; i itself when none). */
  firstVisibleFrom(i: number): number {
    const j = this.nextVisible(i);
    if (j < this.count) return j;
    const back = this.prevVisible(i);
    return back >= 0 ? back : Math.min(Math.max(0, i), this.count - 1);
  }

  /** Last visible line at or before i (else the first visible after it). */
  lastVisibleBefore(i: number): number {
    const j = this.prevVisible(i);
    return j >= 0 ? j : this.firstVisibleFrom(i);
  }

  /**
   * Step from i by delta visible lines, clamped to the axis. When there is no visible line in that
   * direction the result is i itself (which may be hidden: the caller decides what to do).
   */
  step(i: number, delta: number): number {
    let j = i;
    for (let n = Math.abs(delta); n > 0; n--) {
      const next = delta > 0 ? this.nextVisible(j + 1) : this.prevVisible(j - 1);
      if (next < 0 || next >= this.count) break;
      j = next;
    }
    return j;
  }
}
