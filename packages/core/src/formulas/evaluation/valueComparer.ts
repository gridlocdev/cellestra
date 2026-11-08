import type { CellError } from '../../values/cellError.ts';
import { CellValue } from '../../values/cellValue.ts';

const RANK_NUMBER = 0;
const RANK_TEXT = 1;
const RANK_BOOLEAN = 2;

function rank(v: CellValue): number {
  if (v.kind === 'boolean') return RANK_BOOLEAN;
  if (v.kind === 'text') return RANK_TEXT;
  return RANK_NUMBER; // number and (normalized) empty
}

function blankAs(other: CellValue): CellValue {
  if (other.kind === 'text') return CellValue.text('');
  if (other.kind === 'boolean') return CellValue.boolean(false);
  return CellValue.number(0);
}

const EPSILON = Math.pow(2, -48);
/** Half of the 15th significant decimal digit, the precision Calc displays numbers to. */
const HALF_15TH_SIGNIFICAND = 5e-15;

/** 10^exponent, correctly rounded like the literal table LibreOffice uses (`getN10Exp`). */
const powerOfTen = (exponent: number): number => Number(`1e${exponent}`);

/** An integer a double holds exactly (below 2^53). Expects a non-negative value. */
const isRepresentableInteger = (absValue: number): boolean =>
  absValue < 2 ** 53 && Number.isInteger(absValue);

/**
 * Equal within LibreOffice's `rtl::math::approxEqual` (sal/rtl/math.cxx), so binary noise does not
 * break equality: 0.1+0.2 = 0.3. The tolerance is the larger of 2^-48 relative to the smaller
 * magnitude and half the 15th significant decimal digit of it. Zero equals only zero, values of
 * opposite sign never match, and two different integers a double holds exactly never match
 * (1E15 ≠ 1E15+1).
 */
export function approxEqual(a: number, b: number): boolean {
  if (a === b) return true;
  if (a === 0 || b === 0 || a < 0 !== b < 0) return false;
  const d = Math.abs(a - b);
  if (!Number.isFinite(d)) return false;
  const absA = Math.abs(a);
  const absB = Math.abs(b);
  const min = Math.min(absA, absB);
  const threshold = Math.max(min * EPSILON, powerOfTen(Math.floor(Math.log10(min))) * HALF_15TH_SIGNIFICAND);
  if (d >= threshold) return false;
  return !(isRepresentableInteger(absA) && isRepresentableInteger(absB));
}

/**
 * a + b, but 0 when the operands have opposite signs and cancel within `approxEqual`, so
 * 0.1+0.2-0.3 is exactly 0. LibreOffice's `rtl::math::approxAdd` (include/rtl/math.hxx), which
 * Calc's + operator uses (ScInterpreter::CalculateAddSub, sc/source/core/tool/interpr5.cxx).
 */
export function approxAdd(a: number, b: number): number {
  if (((a < 0 && b > 0) || (b < 0 && a > 0)) && approxEqual(a, -b)) return 0;
  return a + b;
}

/** a − b, but 0 when the operands share a sign and are equal within `approxEqual`: `rtl::math::approxSub`. */
export function approxSub(a: number, b: number): number {
  if (((a < 0 && b < 0) || (a > 0 && b > 0)) && approxEqual(a, b)) return 0;
  return a - b;
}

/** .NET OrdinalIgnoreCase: compare upper-cased UTF-16 code units. */
export function compareOrdinalIgnoreCase(a: string, b: string): number {
  const x = a.toUpperCase();
  const y = b.toUpperCase();
  return x < y ? -1 : x > y ? 1 : 0;
}

/** LibreOffice has no boolean type in comparisons: TRUE is 1 and FALSE is 0, so 1=TRUE() is TRUE. */
const booleanAsNumber = (v: CellValue): CellValue =>
  v.kind === 'boolean' ? CellValue.number(v.value ? 1 : 0) : v;

function relational(a: CellValue, b: CellValue, predicate: (c: number) => boolean): CellValue {
  const cmp = ValueComparer.compare(booleanAsNumber(a), booleanAsNumber(b));
  return typeof cmp === 'number' ? CellValue.boolean(predicate(cmp)) : CellValue.error(cmp);
}

/**
 * Ordering rules for the comparison operators. Within a kind, numbers compare numerically and text
 * case-insensitively, numbers within `approxEqual`. Across kinds the rank is numbers < text <
 * booleans (sorting); the comparison operators treat booleans as the numbers 1 and 0 first. A blank
 * takes the shape of whatever it is compared against (0, "", or FALSE). Errors propagate.
 */
export const ValueComparer = {
  /** Negative / 0 / positive, or the error of the first erroneous operand. */
  compare(left: CellValue, right: CellValue): number | CellError {
    if (left.kind === 'error') return left.error;
    if (right.kind === 'error') return right.error;
    if (left.kind === 'empty' && right.kind !== 'empty') left = blankAs(right);
    else if (right.kind === 'empty' && left.kind !== 'empty') right = blankAs(left);
    else if (left.kind === 'empty' && right.kind === 'empty') left = right = CellValue.number(0);

    const lr = rank(left);
    const rr = rank(right);
    if (lr !== rr) return lr < rr ? -1 : 1;
    if (lr === RANK_TEXT) {
      return compareOrdinalIgnoreCase((left as { value: string }).value, (right as { value: string }).value);
    }
    const a = CellValue.asNumber(left);
    const b = CellValue.asNumber(right);
    return approxEqual(a, b) ? 0 : a < b ? -1 : 1;
  },

  equal: (a: CellValue, b: CellValue): CellValue => relational(a, b, (c) => c === 0),
  notEqual: (a: CellValue, b: CellValue): CellValue => relational(a, b, (c) => c !== 0),
  less: (a: CellValue, b: CellValue): CellValue => relational(a, b, (c) => c < 0),
  lessEqual: (a: CellValue, b: CellValue): CellValue => relational(a, b, (c) => c <= 0),
  greater: (a: CellValue, b: CellValue): CellValue => relational(a, b, (c) => c > 0),
  greaterEqual: (a: CellValue, b: CellValue): CellValue => relational(a, b, (c) => c >= 0),
} as const;
