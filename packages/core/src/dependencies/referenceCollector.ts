import type { CellRange } from '../model/cellRange.ts';
import type { Expr } from '../formulas/ast/expr.ts';
import { FunctionRegistry } from '../formulas/functions/functionRegistry.ts';
import { CellRef, type RangeRef } from './cellRef.ts';

/** Largest range expanded cell by cell; bigger ranges are summarized by their corners. */
export const MAX_EXPANDED_CELLS = 200_000;

function expand(
  range: CellRange,
  homeSheet: string,
  destination: Map<string, CellRef>,
  largeRanges: RangeRef[] | undefined,
): void {
  const sheet = range.sheet ?? homeSheet;
  const add = (row: number, column: number): void => {
    const ref = CellRef.of(sheet, row, column);
    destination.set(CellRef.key(ref), ref);
  };
  if (range.cellCount > MAX_EXPANDED_CELLS) {
    // Summarize by the corners to bound graph size; record the full extent so an interior edit
    // still marks the formula.
    add(range.top, range.left);
    add(range.bottom, range.right);
    largeRanges?.push({ sheet, top: range.top, left: range.left, bottom: range.bottom, right: range.right });
    return;
  }
  for (let r = range.top; r <= range.bottom; r++) for (let c = range.left; c <= range.right; c++) add(r, c);
}

/**
 * Walks a formula AST to find the cells it reads (its precedents) and whether it is volatile.
 * Ranges and named ranges are expanded to their member cells.
 */
export const ReferenceCollector = {
  maxExpandedCells: MAX_EXPANDED_CELLS,

  /**
   * Collect the cells `expr` reads into `destination` (keyed by `CellRef.key`). Ranges larger than
   * `MAX_EXPANDED_CELLS` add only their corners there and, when `largeRanges` is given, their full
   * extent there.
   */
  collect(
    expr: Expr,
    homeSheet: string,
    resolveName: (name: string) => CellRange | null,
    destination: Map<string, CellRef>,
    largeRanges?: RangeRef[],
  ): void {
    switch (expr.kind) {
      case 'reference': {
        const ref = CellRef.of(expr.address.sheet ?? homeSheet, expr.address.row, expr.address.column);
        destination.set(CellRef.key(ref), ref);
        break;
      }
      case 'range':
        expand(expr.range, homeSheet, destination, largeRanges);
        break;
      case 'name': {
        const named = resolveName(expr.name);
        if (named) expand(named, homeSheet, destination, largeRanges);
        break;
      }
      case 'unary':
        ReferenceCollector.collect(expr.operand, homeSheet, resolveName, destination, largeRanges);
        break;
      case 'binary':
        ReferenceCollector.collect(expr.left, homeSheet, resolveName, destination, largeRanges);
        ReferenceCollector.collect(expr.right, homeSheet, resolveName, destination, largeRanges);
        break;
      case 'call':
        for (const a of expr.args)
          ReferenceCollector.collect(a, homeSheet, resolveName, destination, largeRanges);
        break;
      default:
        break;
    }
  },

  isVolatile(expr: Expr): boolean {
    switch (expr.kind) {
      case 'call':
        return FunctionRegistry.isVolatile(expr.name) || expr.args.some(ReferenceCollector.isVolatile);
      case 'unary':
        return ReferenceCollector.isVolatile(expr.operand);
      case 'binary':
        return ReferenceCollector.isVolatile(expr.left) || ReferenceCollector.isVolatile(expr.right);
      default:
        return false;
    }
  },
} as const;
