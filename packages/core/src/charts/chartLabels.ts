import type { CellValue } from '../values/cellValue.ts';
import { fixedDigits } from '../formulas/evaluation/decimalDigits.ts';

/** .NET `ToString("0.######", InvariantCulture)`. */
function invariantLabel(n: number): string {
  if (!Number.isFinite(n)) return Number.isNaN(n) ? 'NaN' : n > 0 ? 'Infinity' : '-Infinity';
  const { integer, fraction } = fixedDigits(Math.abs(n), 6);
  const trimmed = fraction.replace(/0+$/, '');
  const body = integer + (trimmed ? '.' + trimmed : '');
  return n < 0 && /[1-9]/.test(body) ? '-' + body : body;
}

/** A cell's value as a chart label: text as is, numbers without locale, nothing for gaps. */
export function labelOf(value: CellValue): string {
  switch (value.kind) {
    case 'text':
      return value.value;
    case 'number':
      return invariantLabel(value.value);
    case 'boolean':
      return value.value ? 'TRUE' : 'FALSE';
    case 'empty':
    case 'error':
    default:
      return '';
  }
}
