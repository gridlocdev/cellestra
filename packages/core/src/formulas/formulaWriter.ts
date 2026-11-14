import type { CellAddress } from '../model/cellAddress.ts';
import { quoteSheetIfNeeded } from '../model/cellAddress.ts';
import type { CellRange } from '../model/cellRange.ts';
import { SheetDimensions } from '../model/sheetDimensions.ts';
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
  union: '~',
  intersect: '!',
};

/** Binary-operator precedence shared by the A1 writer and the OpenFormula translator. */
export function binaryPrecedence(op: BinaryOperator): number {
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
      return 5;
    case 'union':
      return 8;
    case 'intersect':
      return 9;
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

/**
 * An inline array as formula text with the given column and row separators (A1: `,` and `;`;
 * OpenFormula: `;` and `|`). Booleans are written bare in both, as Calc does inside an array.
 */
export function arrayText(
  rows: readonly (readonly CellValue[])[],
  columnSeparator: string,
  rowSeparator: string,
): string {
  return `{${rows.map((row) => row.map((v) => literalText(v)).join(columnSeparator)).join(rowSeparator)}}`;
}

/**
 * A range that spans every row with absolute row corners is a whole-column range (`A:C`); every
 * column with absolute column corners, a whole-row range (`1:3`). Null for an ordinary range.
 * `sheetPrefix` renders the start's sheet and `endPrefix` goes before the end, so both writers share it.
 */
export function wholeLinesText(
  range: CellRange,
  sheetPrefix: (sheet: string | null) => string,
  endPrefix: string,
): string | null {
  const { start, end } = range;
  const col = (a: CellAddress): string => `${a.columnAbsolute ? '$' : ''}${a.columnLabel}`;
  const row = (a: CellAddress): string => `${a.rowAbsolute ? '$' : ''}${a.rowLabel}`;
  if (range.top === 0 && range.bottom === SheetDimensions.maxRows - 1 && start.rowAbsolute && end.rowAbsolute)
    return `${sheetPrefix(range.sheet)}${col(start)}:${endPrefix}${col(end)}`;
  if (
    range.left === 0 &&
    range.right === SheetDimensions.maxColumns - 1 &&
    start.columnAbsolute &&
    end.columnAbsolute
  )
    return `${sheetPrefix(range.sheet)}${row(start)}:${endPrefix}${row(end)}`;
  return null;
}

const a1SheetPrefix = (sheet: string | null): string =>
  sheet === null ? '' : `${quoteSheetIfNeeded(sheet)}.`;

/**
 * Precedence of a node as written: binary operators by `binaryPrecedence`, then a prefix sign (6,
 * tighter than ^, so (-2)^2 is written -2^2 and -(2^2) keeps its parentheses), then postfix percent,
 * then the reference operators (union 8, intersection 9), which bind tightest as in Calc.
 */
export function writtenPrecedence(expr: Expr): number {
  if (expr.kind === 'binary') return binaryPrecedence(expr.operator);
  if (expr.kind === 'unary') return expr.operator === 'percent' ? 7 : 6;
  return 10;
}

function withParens(expr: Expr, minPrecedence: number): string {
  const rendered = render(expr);
  return writtenPrecedence(expr) < minPrecedence ? `(${rendered})` : rendered;
}

function renderUnary(u: UnaryExpr): string {
  switch (u.operator) {
    case 'negate':
      return '-' + withParens(u.operand, 6);
    case 'plus':
      return '+' + withParens(u.operand, 6);
    case 'percent':
      return withParens(u.operand, 7) + '%';
  }
}

function renderBinary(b: BinaryExpr): string {
  const p = writtenPrecedence(b);
  return `${withParens(b.left, p)}${BINARY_SYMBOL[b.operator]}${withParens(b.right, p + 1)}`;
}

function render(expr: Expr): string {
  switch (expr.kind) {
    case 'literal':
      return literalText(expr.value);
    case 'array':
      return arrayText(expr.rows, ',', ';');
    case 'reference':
      return expr.address.toString();
    case 'range':
      return wholeLinesText(expr.range, a1SheetPrefix, '') ?? expr.range.toString();
    case 'sheetSpan': {
      const last = a1SheetPrefix(expr.lastSheet);
      const { range } = expr;
      return (
        wholeLinesText(range, a1SheetPrefix, last) ??
        `${range.start.toString()}:${last}${range.end.withoutSheet().toString()}`
      );
    }
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
