import { roundTripNumber, type CellValue } from '../values/cellValue.ts';
import type { BinaryExpr, BinaryOperator, Expr, UnaryExpr } from './ast/expr.ts';

export const BINARY_SYMBOL: Readonly<Record<BinaryOperator, string>> = {
  add: '+',
  subtract: '-',
  multiply: '*',
  divide: '/',
  power: '^',
  concat: '&',
  equal: '=',
  notEqual: '<>',
  less: '<',
  lessEqual: '<=',
  greater: '>',
  greaterEqual: '>=',
};

/** Binary-operator precedence shared by the A1 writer and the OpenFormula translator (power differs). */
export function binaryPrecedence(op: BinaryOperator, power: number): number {
  switch (op) {
    case 'concat':
      return 2;
    case 'add':
    case 'subtract':
      return 3;
    case 'multiply':
    case 'divide':
      return 4;
    case 'power':
      return power;
    default:
      return 1; // comparisons
  }
}

/** A literal as formula text (shared by both writers; booleans differ in OpenFormula). */
export function literalText(value: CellValue, booleanSuffix = ''): string {
  switch (value.kind) {
    case 'number':
      return roundTripNumber(value.value);
    case 'boolean':
      return (value.value ? 'TRUE' : 'FALSE') + booleanSuffix;
    case 'text':
      return `"${value.value.replaceAll('"', '""')}"`;
    case 'error':
      return value.error.token;
    case 'empty':
      return '';
  }
}

// Unary +/- binds looser than ^ (5 < power's 6) so -2^2 round-trips without spurious parens;
// percent (postfix) binds tighter than everything.
function precedence(expr: Expr): number {
  if (expr.kind === 'binary') return binaryPrecedence(expr.operator, 6);
  if (expr.kind === 'unary') return expr.operator === 'percent' ? 7 : 5;
  return 8;
}

function withParens(expr: Expr, minPrecedence: number): string {
  const rendered = render(expr);
  return precedence(expr) < minPrecedence ? `(${rendered})` : rendered;
}

function renderUnary(u: UnaryExpr): string {
  switch (u.operator) {
    case 'negate':
      return '-' + withParens(u.operand, 5);
    case 'plus':
      return '+' + withParens(u.operand, 5);
    case 'percent':
      return withParens(u.operand, 7) + '%';
  }
}

function renderBinary(b: BinaryExpr): string {
  const p = precedence(b);
  return `${withParens(b.left, p)}${BINARY_SYMBOL[b.operator]}${withParens(b.right, p + 1)}`;
}

function render(expr: Expr): string {
  switch (expr.kind) {
    case 'literal':
      return literalText(expr.value);
    case 'reference':
      return expr.address.toString();
    case 'range':
      return expr.range.toString();
    case 'name':
      return expr.name;
    case 'refError':
      return '#REF!';
    case 'unary':
      return renderUnary(expr);
    case 'binary':
      return renderBinary(expr);
    case 'call':
      return `${expr.name}(${expr.args.map(render).join(', ')})`;
  }
}

/**
 * Renders a formula AST back to canonical A1 text (without the leading `=`), adding the minimum
 * parentheses needed to preserve grouping.
 */
export const FormulaWriter = {
  write: (expr: Expr): string => render(expr),
} as const;
