import { CellError } from './cellError.ts';

/**
 * A typed, immutable cell value: empty, number, text, boolean, or error.
 * A discriminated union of frozen plain objects; `EMPTY`, `TRUE`, `FALSE` are shared singletons.
 */
export type CellValue = EmptyValue | NumberValue | TextValue | BooleanValue | ErrorValue;

export interface EmptyValue {
  readonly kind: 'empty';
}
export interface NumberValue {
  readonly kind: 'number';
  readonly value: number;
}
export interface TextValue {
  readonly kind: 'text';
  readonly value: string;
}
export interface BooleanValue {
  readonly kind: 'boolean';
  readonly value: boolean;
}
export interface ErrorValue {
  readonly kind: 'error';
  readonly error: CellError;
}

export type CellValueKind = CellValue['kind'];

const EMPTY: EmptyValue = Object.freeze({ kind: 'empty' });
const TRUE: BooleanValue = Object.freeze({ kind: 'boolean', value: true });
const FALSE: BooleanValue = Object.freeze({ kind: 'boolean', value: false });

export const CellValue = {
  empty: EMPTY as CellValue,
  number(value: number): CellValue {
    return { kind: 'number', value };
  },
  text(value: string): CellValue {
    return { kind: 'text', value };
  },
  boolean(value: boolean): CellValue {
    return value ? TRUE : FALSE;
  },
  error(error: CellError): CellValue {
    return { kind: 'error', error };
  },

  isEmpty: (v: CellValue): v is EmptyValue => v.kind === 'empty',
  isNumber: (v: CellValue): v is NumberValue => v.kind === 'number',
  isText: (v: CellValue): v is TextValue => v.kind === 'text',
  isBoolean: (v: CellValue): v is BooleanValue => v.kind === 'boolean',
  isError: (v: CellValue): v is ErrorValue => v.kind === 'error',

  /** Numeric payload of a number or boolean (TRUE = 1). Throws for other kinds. */
  asNumber(v: CellValue): number {
    if (v.kind === 'number') return v.value;
    if (v.kind === 'boolean') return v.value ? 1 : 0;
    throw new TypeError(`CellValue of kind ${v.kind} is not numeric.`);
  },

  /**
   * Culture-invariant, lossless string form: stored as the cached value in `.ods` and used for
   * tab-separated clipboard text. Display formatting is the styling layer's job.
   */
  toInvariantString(v: CellValue): string {
    switch (v.kind) {
      case 'empty':
        return '';
      case 'number':
        return roundTripNumber(v.value);
      case 'text':
        return v.value;
      case 'boolean':
        return v.value ? 'TRUE' : 'FALSE';
      case 'error':
        return v.error.token;
    }
  },

  equals(a: CellValue, b: CellValue): boolean {
    if (a.kind !== b.kind) return false;
    switch (a.kind) {
      case 'empty':
        return true;
      case 'number':
        return Object.is(a.value, (b as NumberValue).value) || a.value === (b as NumberValue).value;
      case 'text':
        return a.value === (b as TextValue).value;
      case 'boolean':
        return a.value === (b as BooleanValue).value;
      case 'error':
        return a.error.equals((b as ErrorValue).error);
    }
  },
} as const;

/**
 * .NET "R" round-trip formatting: shortest string that parses back to the same double, never
 * localized. JS `String(n)` is already shortest-round-trip; .NET writes exponents as `E+XX`.
 */
export function roundTripNumber(n: number): string {
  if (Number.isNaN(n)) return 'NaN';
  if (n === Infinity) return 'Infinity';
  if (n === -Infinity) return '-Infinity';
  const s = String(n);
  const e = s.indexOf('e');
  if (e < 0) return s;
  const mantissa = s.slice(0, e);
  const exp = s.slice(e + 1);
  const sign = exp.startsWith('-') ? '-' : '+';
  const digits = exp.replace(/^[+-]/, '').padStart(2, '0');
  return `${mantissa}E${sign}${digits}`;
}
