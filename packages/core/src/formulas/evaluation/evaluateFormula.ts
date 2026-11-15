import type { Workbook } from '../../model/workbook.ts';
import type { CellValue } from '../../values/cellValue.ts';
import type { Expr } from '../ast/expr.ts';
import { FunctionRegistry } from '../functions/functionRegistry.ts';
import type { CivilDateTime } from './dateConversion.ts';
import { Evaluator } from './evaluator.ts';
import { WorkbookEvaluationContext } from './workbookEvaluationContext.ts';

/** Where a formula is evaluated: the sheet unqualified references read, and what it may assume. */
export interface FormulaSite {
  readonly workbook: Workbook;
  readonly sheet: string;
  /** The formula's own cell, for implicit intersection; none for a formula typed outside a cell. */
  readonly cell?: { readonly row: number; readonly column: number };
  /** Defaults to the workbook's locale. */
  readonly locale?: string;
  /** Defaults to the system clock (TODAY, NOW). */
  readonly clock?: () => CivilDateTime;
  /** Defaults to the standard library. */
  readonly functions?: FunctionRegistry;
}

/** Evaluate a parsed formula against a workbook: the one way to compute a formula's value. */
export function evaluateFormula(expr: Expr, site: FormulaSite): CellValue {
  const context = new WorkbookEvaluationContext(
    site.workbook,
    site.sheet,
    site.locale ?? site.workbook.settings.effectiveLocale,
    site.clock,
    site.cell,
  );
  return new Evaluator(context, site.functions ?? FunctionRegistry.standard).evaluate(expr);
}
