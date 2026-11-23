import type { CellRange } from '../model/cellRange.ts';
import type { Expr } from '../formulas/ast/expr.ts';
import { FunctionRegistry } from '../formulas/functions/functionRegistry.ts';
import { CellRef, type RangeRef } from './cellRef.ts';

/**
 * Largest range expanded cell by cell; bigger ranges are summarized by their corners and tracked as
 * a large range. Kept small so running totals (`SUM($A$1:A4000)` down a column) stay linear in
 * graph size rather than one edge per cell per formula.
 */
export const MAX_EXPANDED_CELLS = 1024;

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
   * extent there. `sheetsBetween` gives the sheets a 3D range (`Sheet1.A1:Sheet3.B2`) spans.
   * `nameExpression` gives the expression a name holds (`TaxRate` = `0.2`), whose reads count too.
   */
  collect(
    expr: Expr,
    homeSheet: string,
    resolveName: (name: string) => CellRange | null,
    destination: Map<string, CellRef>,
    largeRanges?: RangeRef[],
    sheetsBetween?: (first: string, last: string) => readonly string[] | null,
    functions: FunctionRegistry = FunctionRegistry.standard,
    nameExpression?: (name: string) => Expr | null,
  ): void {
    const recurse = (e: Expr, names = nameExpression): void =>
      ReferenceCollector.collect(
        e,
        homeSheet,
        resolveName,
        destination,
        largeRanges,
        sheetsBetween,
        functions,
        names,
      );
    switch (expr.kind) {
      case 'reference': {
        const ref = CellRef.of(expr.address.sheet ?? homeSheet, expr.address.row, expr.address.column);
        destination.set(CellRef.key(ref), ref);
        break;
      }
      case 'range':
        expand(expr.range, homeSheet, destination, largeRanges);
        break;
      case 'sheetSpan': {
        // Every sheet between the two ends; without the sheet order, at least the ends.
        const first = expr.range.sheet!;
        const sheets = sheetsBetween?.(first, expr.lastSheet) ?? [first, expr.lastSheet];
        for (const sheet of sheets) expand(expr.range.withSheet(sheet), homeSheet, destination, largeRanges);
        break;
      }
      case 'name': {
        const named = resolveName(expr.name);
        if (named) {
          expand(named, homeSheet, destination, largeRanges);
          break;
        }
        const body = nameExpression?.(expr.name);
        // Inside the name's own expression the name is gone, so a name that uses itself stops.
        const key = expr.name.toLowerCase();
        if (body) recurse(body, (n) => (n.toLowerCase() === key ? null : nameExpression!(n)));
        break;
      }
      case 'unary':
        recurse(expr.operand);
        break;
      case 'binary':
        recurse(expr.left);
        recurse(expr.right);
        break;
      case 'call': {
        for (const a of expr.args) recurse(a);
        const extra = functions.get(expr.name)?.extraReads?.(expr.args, resolveName) ?? [];
        for (const range of extra) expand(range, homeSheet, destination, largeRanges);
        break;
      }
      case 'array':
      case 'literal':
      case 'refError':
      default:
        break;
    }
  },

  /** Whether `expr` calls a volatile function anywhere (TODAY, NOW, SUBTOTAL). */
  isVolatile(expr: Expr, functions: FunctionRegistry = FunctionRegistry.standard): boolean {
    const volatile = (e: Expr): boolean => ReferenceCollector.isVolatile(e, functions);
    switch (expr.kind) {
      case 'call':
        return functions.get(expr.name)?.volatile === true || expr.args.some(volatile);
      case 'unary':
        return volatile(expr.operand);
      case 'binary':
        return volatile(expr.left) || volatile(expr.right);
      case 'array':
      case 'literal':
      case 'name':
      case 'range':
      case 'reference':
      case 'refError':
      case 'sheetSpan':
      default:
        return false;
    }
  },
} as const;
