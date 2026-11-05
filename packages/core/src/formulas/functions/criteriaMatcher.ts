import { CellValue } from '../../values/cellValue.ts';
import { compareOrdinalIgnoreCase } from '../evaluation/valueComparer.ts';
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

function compareNumeric(value: CellValue, op: Operator, operand: number): boolean {
  if (value.kind !== 'number') return op === '<>';
  const v = value.value;
  switch (op) {
    case '=':
      return v === operand;
    case '<>':
      return v !== operand;
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

function compareText(value: CellValue, op: Operator, operand: string): boolean {
  const cmp = compareOrdinalIgnoreCase(asText(value), operand);
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
    if (criteria.kind === 'number') {
      const target = criteria.value;
      return new CriteriaMatcher((v) => v.kind === 'number' && v.value === target);
    }
    if (criteria.kind === 'boolean') {
      const target = criteria.value;
      return new CriteriaMatcher((v) => v.kind === 'boolean' && v.value === target);
    }
    const text = asText(criteria);
    const [op, operand] = splitOperator(text);
    const numeric = parseNumberAny(operand, localeInfo(locale));
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
