import type { CellRange } from '../model/cellRange.ts';
import { CellValue } from '../values/cellValue.ts';
import type { DataValidation } from './types.ts';

/** Resolves a range (a list rule's source) to its values. */
export type RangeResolver = (range: CellRange) => Iterable<CellValue>;

/** Whether a rule applies at (row, column). */
export function validationCovers(rule: DataValidation, row: number, column: number): boolean {
  return rule.target.contains(row, column);
}

/** Data validation checks on entry (PRD §8.13). Empty is always allowed; required-ness is separate. */
export const Validator = {
  isValid(
    rule: DataValidation,
    value: CellValue,
    _locale: string,
    resolveRange?: RangeResolver | null,
  ): boolean {
    if (value.kind === 'empty') return true;
    switch (rule.kind) {
      case 'any':
        return true;
      case 'list':
        return matchesList(rule, value, resolveRange ?? null);
      case 'textLength':
        return compareNumber(
          rule,
          (value.kind === 'text' ? value.value : CellValue.toInvariantString(value)).length,
        );
      case 'wholeNumber': {
        const n = numberOf(value);
        return n !== null && n === Math.trunc(n) && compareNumber(rule, n);
      }
      case 'decimal':
      case 'date': {
        const n = numberOf(value);
        return n !== null && compareNumber(rule, n);
      }
    }
  },
} as const;

/** List matching uses the invariant text form, case-insensitively, so it is stable across locales. */
function matchesList(rule: DataValidation, value: CellValue, resolveRange: RangeResolver | null): boolean {
  const text = (value.kind === 'text' ? value.value : CellValue.toInvariantString(value)).toUpperCase();
  if (rule.listValues?.some((v) => v.toUpperCase() === text)) return true;
  if (rule.listSource && resolveRange) {
    for (const candidate of resolveRange(rule.listSource)) {
      const t = candidate.kind === 'text' ? candidate.value : CellValue.toInvariantString(candidate);
      if (t.toUpperCase() === text) return true;
    }
  }
  return false;
}

function compareNumber(rule: DataValidation, v: number): boolean {
  const a = rule.operand1;
  const b = rule.operand2;
  switch (rule.operator) {
    case 'between':
      return v >= a && v <= b;
    case 'notBetween':
      return v < a || v > b;
    case 'equal':
      return v === a;
    case 'notEqual':
      return v !== a;
    case 'greaterThan':
      return v > a;
    case 'lessThan':
      return v < a;
    case 'greaterOrEqual':
      return v >= a;
    case 'lessOrEqual':
      return v <= a;
  }
}

function numberOf(value: CellValue): number | null {
  return value.kind === 'number' || value.kind === 'boolean' ? CellValue.asNumber(value) : null;
}
