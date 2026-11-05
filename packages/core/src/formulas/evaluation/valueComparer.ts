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

/** .NET OrdinalIgnoreCase: compare upper-cased UTF-16 code units. */
export function compareOrdinalIgnoreCase(a: string, b: string): number {
  const x = a.toUpperCase();
  const y = b.toUpperCase();
  return x < y ? -1 : x > y ? 1 : 0;
}

function relational(a: CellValue, b: CellValue, predicate: (c: number) => boolean): CellValue {
  const cmp = ValueComparer.compare(a, b);
  return typeof cmp === 'number' ? CellValue.boolean(predicate(cmp)) : CellValue.error(cmp);
}

/**
 * Ordering rules for the comparison operators. Within a kind, numbers compare numerically and text
 * case-insensitively. Across kinds the rank is numbers < text < booleans. A blank takes the shape of
 * whatever it is compared against (0, "", or FALSE). Errors propagate.
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
    return a < b ? -1 : a > b ? 1 : 0;
  },

  equal: (a: CellValue, b: CellValue): CellValue => relational(a, b, (c) => c === 0),
  notEqual: (a: CellValue, b: CellValue): CellValue => relational(a, b, (c) => c !== 0),
  less: (a: CellValue, b: CellValue): CellValue => relational(a, b, (c) => c < 0),
  lessEqual: (a: CellValue, b: CellValue): CellValue => relational(a, b, (c) => c <= 0),
  greater: (a: CellValue, b: CellValue): CellValue => relational(a, b, (c) => c > 0),
  greaterEqual: (a: CellValue, b: CellValue): CellValue => relational(a, b, (c) => c >= 0),
} as const;
