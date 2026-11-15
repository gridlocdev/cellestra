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

/** A list value's text in its invariant form, trimmed. */
const listText = (v: CellValue): string =>
  (v.kind === 'text' ? v.value : CellValue.toInvariantString(v)).trim();

/**
 * List matching uses the invariant text form, trimmed and case-insensitive, so it is stable across
 * locales. A blank (whitespace) entry is allowed, as an empty cell is.
 */
function matchesList(rule: DataValidation, value: CellValue, resolveRange: RangeResolver | null): boolean {
  const text = listText(value).toUpperCase();
  if (text === '') return true;
  if (rule.listValues?.some((v) => v.trim().toUpperCase() === text)) return true;
  if (rule.listSource && resolveRange) {
    for (const candidate of resolveRange(rule.listSource)) {
      if (candidate.kind !== 'empty' && listText(candidate).toUpperCase() === text) return true;
    }
  }
  return false;
}

/**
 * The items a list rule's dropdown offers: its typed values, or its source range's cells, trimmed,
 * without blank or whitespace items, and without repeats (ignoring case; the first spelling wins).
 */
export function listItems(rule: DataValidation, resolveRange: RangeResolver | null): string[] {
  const raw: string[] = [];
  if (rule.listValues && rule.listValues.length > 0) raw.push(...rule.listValues);
  else if (rule.listSource && resolveRange)
    for (const v of resolveRange(rule.listSource)) if (v.kind !== 'empty') raw.push(listText(v));
  const seen = new Set<string>();
  const out: string[] = [];
  for (const item of raw) {
    const text = item.trim();
    const key = text.toUpperCase();
    if (text === '' || seen.has(key)) continue;
    seen.add(key);
    out.push(text);
  }
  return out;
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
