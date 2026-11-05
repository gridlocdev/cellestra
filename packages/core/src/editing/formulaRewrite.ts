import type { Expr } from '../formulas/ast/expr.ts';
import { FormulaWriter } from '../formulas/formulaWriter.ts';
import { Parser } from '../formulas/parsing/parser.ts';

/**
 * Parse a formula (source without `=`), transform its AST, and write it back. A formula that does
 * not parse is returned unchanged — the edit never destroys what the user typed.
 */
export function rewriteFormula(formula: string, transform: (expr: Expr) => Expr): string {
  const parsed = Parser.tryParse(formula);
  return parsed.ok ? FormulaWriter.write(transform(parsed.expr)) : formula;
}
