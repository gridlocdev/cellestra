/**
 * Counts clicks on pointer presses. `pointerdown` carries the click count in `detail` in some
 * browsers but reports 0 in others (Chromium, Safari), so a double click on a header edge or the
 * fill handle was never seen. When the browser gives no count, a press soon after the last one and
 * near it counts as the next click.
 */
export class ClickCounter {
  /** The longest gap between presses of one multi-click, as most systems default it. */
  static readonly interval = 500;
  /** How far the pointer may move between presses, in CSS pixels. */
  static readonly slop = 4;

  #last: { time: number; x: number; y: number; count: number } | null = null;

  /** Record a press at `time` (ms) and (x, y); returns 1 for a single click, 2 for a double, … */
  press(time: number, x: number, y: number, detail: number): number {
    const last = this.#last;
    let count: number;
    if (detail > 0) count = detail;
    else if (
      last &&
      time - last.time <= ClickCounter.interval &&
      Math.abs(x - last.x) <= ClickCounter.slop &&
      Math.abs(y - last.y) <= ClickCounter.slop
    )
      count = last.count + 1;
    else count = 1;
    this.#last = { time, x, y, count };
    return count;
  }
}
