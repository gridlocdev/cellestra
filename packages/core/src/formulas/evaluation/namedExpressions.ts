import { CellRange } from '../../model/cellRange.ts';
import type { Workbook } from '../../model/workbook.ts';
import type { Expr } from '../ast/expr.ts';
import { FormulaWriter } from '../formulaWriter.ts';
import { Parser } from '../parsing/parser.ts';

const MAX_CACHE = 512;
const parsed = new Map<string, Expr | null>();

/** An expression's source parsed once; null when it does not parse. */
export function parseNameExpression(source: string): Expr | null {
  if (parsed.has(source)) return parsed.get(source)!;
  const result = Parser.tryParse(source);
  const expr = result.ok ? result.expr : null;
  if (parsed.size >= MAX_CACHE) parsed.clear();
  parsed.set(source, expr);
  return expr;
}

/**
 * The parsed expression a name holds as a formula on `sheet` sees it (a sheet-local name shadows a
 * global one), or null when the name is unknown, holds a range, or its expression does not parse.
 */
export function namedExpressionOf(workbook: Workbook, name: string, sheet: string | null): Expr | null {
  const named = workbook.findNamedRange(name, sheet);
  return named?.expression === undefined ? null : parseNameExpression(named.expression);
}

/** What a name holds, as typed in the Name Manager or the CLI. */
export type NameDefinition = { readonly target: CellRange } | { readonly expression: string };

/**
 * Read a typed definition: a range (`B2:D9`, `Sheet1.$A$1`) names cells; anything else that parses
 * as a formula (`0.2`, `=Sheet1.B1*2`, a leading `=` optional) is an expression. Null when it is
 * neither.
 */
export function readNameDefinition(text: string): NameDefinition | null {
  const typed = text.trim();
  const formula = typed.startsWith('=') ? typed.slice(1).trim() : typed;
  if (formula === '') return null;
  const range = CellRange.tryParse(formula.replace(/\$/g, ''));
  if (range) return { target: range };
  const result = Parser.tryParse(formula);
  return result.ok ? { expression: FormulaWriter.write(result.expr) } : null;
}
