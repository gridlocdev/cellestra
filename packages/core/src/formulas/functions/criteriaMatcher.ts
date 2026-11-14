import { CellValue } from '../../values/cellValue.ts';
import { approxEqual, booleanAsNumber, compareOrdinalIgnoreCase } from '../evaluation/valueComparer.ts';
import { localeInfo } from '../evaluation/localeInfo.ts';
import { parseNumberAny } from '../evaluation/valueParsing.ts';

type Predicate = (value: CellValue) => boolean;
type Operator = '=' | '<>' | '>' | '>=' | '<' | '<=';

const OPERATORS: readonly Operator[] = ['<=', '>=', '<>', '=', '<', '>'];

function splitOperator(text: string): [Operator, string] {
  for (const op of OPERATORS) if (text.startsWith(op)) return [op, text.slice(op.length)];
  return ['=', text];
}

const asText = (v: CellValue): string => (v.kind === 'text' ? v.value : CellValue.toInvariantString(v));

/** A cell's number for a numeric criterion: TRUE and FALSE count as 1 and 0, as in Calc. */
function numberOf(value: CellValue): number | null {
  const v = booleanAsNumber(value);
  return v.kind === 'number' ? v.value : null;
}

/** The number a criteria operand stands for: a number, or TRUE/FALSE as 1/0. */
function operandNumber(operand: string, locale: string): number | null {
  const upper = operand.trim().toUpperCase();
  if (upper === 'TRUE') return 1;
  if (upper === 'FALSE') return 0;
  return parseNumberAny(operand, localeInfo(locale));
}

function compareNumeric(value: CellValue, op: Operator, operand: number): boolean {
  const v = numberOf(value);
  if (v === null) return op === '<>';
  switch (op) {
    case '=':
      return approxEqual(v, operand);
    case '<>':
      return !approxEqual(v, operand);
    case '>':
      return v > operand;
    case '>=':
      return v >= operand;
    case '<':
      return v < operand;
    case '<=':
      return v <= operand;
  }
}

/** A text ordering criterion (`"<b"`) compares text cells only: numbers, booleans and blanks never match (Calc). */
function compareText(value: CellValue, op: Operator, operand: string): boolean {
  if (value.kind !== 'text') return false;
  const cmp = compareOrdinalIgnoreCase(value.value, operand);
  switch (op) {
    case '>':
      return cmp > 0;
    case '>=':
      return cmp >= 0;
    case '<':
      return cmp < 0;
    case '<=':
      return cmp <= 0;
    default:
      return false;
  }
}

function wildcardToRegex(pattern: string): RegExp {
  let source = '^';
  for (const c of pattern) {
    if (c === '*') source += '.*';
    else if (c === '?') source += '.';
    else source += c.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&');
  }
  return new RegExp(source + '$', 'iu');
}

/**
 * Case-insensitive equality with `*`/`?` wildcards and no operators — the exact-match test of
 * MATCH/VLOOKUP. Only reads `text`; callers decide which cell kinds take part.
 */
export function textPattern(pattern: string): (text: string) => boolean {
  if (!pattern.includes('*') && !pattern.includes('?')) {
    const target = pattern.toUpperCase();
    return (t) => t.toUpperCase() === target;
  }
  const re = wildcardToRegex(pattern);
  return (t) => re.test(t);
}

function textMatcher(pattern: string): Predicate {
  if (!pattern.includes('*') && !pattern.includes('?')) {
    const target = pattern.toUpperCase();
    return (v) => asText(v).toUpperCase() === target;
  }
  const re = wildcardToRegex(pattern);
  return (v) => re.test(asText(v));
}

/**
 * A predicate built from a SUMIF/COUNTIF criteria value: a comparison (`">5"`, `"<>0"`), a
 * number/text equality, or a text pattern with `*`/`?` wildcards. Comparisons against numbers use
 * numeric ordering; text matches are case-insensitive.
 */
export class CriteriaMatcher {
  readonly #predicate: Predicate;

  private constructor(predicate: Predicate) {
    this.#predicate = predicate;
  }

  matches(value: CellValue): boolean {
    return this.#predicate(value);
  }

  static build(criteria: CellValue, locale: string): CriteriaMatcher {
    // An empty criteria cell means 0, as in Calc: it matches zeros, not blanks. TRUE and FALSE
    // are 1 and 0 on both sides.
    if (criteria.kind === 'number' || criteria.kind === 'empty' || criteria.kind === 'boolean') {
      const target = numberOf(criteria) ?? 0;
      return new CriteriaMatcher((v) => {
        const n = numberOf(v);
        return n !== null && approxEqual(n, target);
      });
    }
    const text = asText(criteria);
    const [op, operand] = splitOperator(text);
    const numeric = operandNumber(operand, locale);
    if (numeric !== null) return new CriteriaMatcher((v) => compareNumeric(v, op, numeric));
    switch (op) {
      case '>':
      case '>=':
      case '<':
      case '<=':
        return new CriteriaMatcher((v) => compareText(v, op, operand));
      case '<>': {
        const m = textMatcher(operand);
        return new CriteriaMatcher((v) => !m(v));
      }
      default:
        return new CriteriaMatcher(textMatcher(operand));
    }
  }
}
