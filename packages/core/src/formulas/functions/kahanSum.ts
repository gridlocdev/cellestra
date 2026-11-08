import { approxEqual } from '../evaluation/valueComparer.ts';

/**
 * Compensated summation the way LibreOffice Calc 7.2+ sums (`KahanSum`, sc/inc/kahan.hxx): a
 * Neumaier (improved Kahan) running sum and error, with the most recent non-zero addend held back.
 * `get()` adds that last value like `rtl::math::approxAdd`, so when it cancels the rest within
 * `approxEqual` the total is exactly 0 — SUM(0.1, 0.2, -0.3) = 0 — and otherwise it is the
 * compensated total. Zeros are skipped.
 *
 * Calc feeds values one at a time for argument lists, SUMIF and AVERAGE; Cellestra does the same
 * for SUM over a range too. (Calc's fast path for SUM over a cell range sums each column block on
 * its own and, on non-x86 builds, folds the held-back value in without the snap (`KahanSum::add(const
 * KahanSum&)`), so SUM over cells holding 0.1, 0.2 and -0.3 is 2.8E-17 in Calc on ARM but 0 on x86
 * and in Cellestra.)
 */
export class KahanSum {
  #sum = 0;
  #error = 0;
  #memory = 0;

  /** One Neumaier step into the running sum and error (`sumNeumaierNormal`). */
  #step(value: number): void {
    const sum = this.#sum;
    const t = sum + value;
    this.#error += Math.abs(sum) >= Math.abs(value) ? sum - t + value : value - t + sum;
    this.#sum = t;
  }

  add(value: number): void {
    if (value === 0) return;
    if (this.#memory === 0) {
      this.#memory = value;
      return;
    }
    this.#step(this.#memory);
    this.#memory = value;
  }

  /** The total. */
  get(): number {
    const total = this.#sum + this.#error;
    const memory = this.#memory;
    if (memory === 0) return total;
    if (((memory < 0 && total > 0) || (total < 0 && memory > 0)) && approxEqual(memory, -total)) return 0;
    const sum = this.#sum;
    const t = sum + memory;
    const error = this.#error + (Math.abs(sum) >= Math.abs(memory) ? sum - t + memory : memory - t + sum);
    return t + error;
  }
}
