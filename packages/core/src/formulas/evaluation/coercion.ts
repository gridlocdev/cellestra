import { CellError } from '../../values/cellError.ts';
import type { CellValue } from '../../values/cellValue.ts';
import { fixedDigits, scientificG15 } from './decimalDigits.ts';

/** `double.TryParse(…, NumberStyles.Float, InvariantCulture)`: sign, digits, one `.`, exponent; no grouping. */
const INVARIANT_FLOAT = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/;

/** Parse invariant float text (surrounding whitespace allowed); null when it is not a number. */
export function parseInvariantFloat(text: string): number | null {
  const t = text.trim();
  return INVARIANT_FLOAT.test(t) ? Number(t) : null;
}

/**
 * Type-coercion rules shared by the evaluator and the functions. Arithmetic coerces loosely
 * (blank → 0, "12" → 12, TRUE → 1); errors always propagate. Each `to*` returns the coerced value or
 * the `CellError` to surface — never throws.
 */
export const Coercion = {
  /** Coerce for arithmetic. Blank → 0, text parsed invariantly, boolean → 1/0. */
  toNumber(value: CellValue): number | CellError {
    switch (value.kind) {
      case 'empty':
        return 0;
      case 'number':
        return value.value;
      case 'boolean':
        return value.value ? 1 : 0;
      case 'text':
        return parseInvariantFloat(value.value) ?? CellError.value(`'${value.value}' is not a number.`);
      case 'error':
        return value.error;
    }
  },

  /** Coerce to text for concatenation and text functions. */
  toText(value: CellValue): string | CellError {
    switch (value.kind) {
      case 'empty':
        return '';
      case 'text':
        return value.value;
      case 'number':
        return Coercion.numberToText(value.value);
      case 'boolean':
        return value.value ? 'TRUE' : 'FALSE';
      case 'error':
        return value.error;
    }
  },

  /** Coerce to a logical value for AND/OR/IF. Numbers: non-zero → true; text "TRUE"/"FALSE". */
  toBoolean(value: CellValue): boolean | CellError {
    switch (value.kind) {
      case 'boolean':
        return value.value;
      case 'number':
        return value.value !== 0;
      case 'empty':
        return false;
      case 'text': {
        const upper = value.value.toUpperCase();
        if (upper === 'TRUE') return true;
        if (upper === 'FALSE') return false;
        return CellError.value(`'${value.value}' is not a logical value.`);
      }
      case 'error':
        return value.error;
    }
  },

  /**
   * Render a number for text contexts without float noise or thousands separators (Calc-style):
   * integers in full decimal below 1e21, 15 significant digits fixed-point inside ~1e-10 … 1e15,
   * scientific (G15) outside that band.
   */
  numberToText(value: number): string {
    if (Number.isNaN(value)) return 'NaN';
    if (!Number.isFinite(value)) return value > 0 ? 'Infinity' : '-Infinity';
    if (value === 0) return '0';
    const abs = Math.abs(value);
    const sign = value < 0 ? '-' : '';
    if (value === Math.floor(value) && abs < 1e21) return sign + fixedDigits(abs, 0).integer;
    const exponent = Math.floor(Math.log10(abs));
    if (exponent >= 15 || exponent <= -11) return scientificG15(value);
    const decimals = Math.max(0, 15 - 1 - exponent);
    const { integer, fraction } = fixedDigits(abs, decimals);
    const trimmed = fraction.replace(/0+$/, '');
    return sign + integer + (trimmed ? '.' + trimmed : '');
  },
} as const;

export const isCellError = (v: unknown): v is CellError => v instanceof CellError;
