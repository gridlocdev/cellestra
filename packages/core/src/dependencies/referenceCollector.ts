import { CellRange } from '../model/cellRange.ts';
import { SheetDimensions } from '../model/sheetDimensions.ts';
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

/** The cells an argument names, as a range: a reference, a range or a named range. */
function asRange(expr: Expr, resolveName: (name: string) => CellRange | null): CellRange | null {
  if (expr.kind === 'reference') {
    const { row, column, sheet } = expr.address;
    return CellRange.fromBounds(row, column, row, column, sheet);
  }
  if (expr.kind === 'range') return expr.range;
  if (expr.kind === 'name') return resolveName(expr.name);
  return null;
}

/**
 * SUMIF reads its sum range from its top-left, sized like the criteria range (Excel/Calc), so
 * `SUMIF(A1:A3, "x", B1)` reads B1:B3. Returns that reshaped range, or null for any other call.
 */
function reshapedSumRange(
  name: string,
  args: readonly Expr[],
  resolveName: (name: string) => CellRange | null,
): CellRange | null {
  if (name.toUpperCase() !== 'SUMIF' || args.length < 3) return null;
  const criteria = asRange(args[0]!, resolveName);
  const sum = asRange(args[2]!, resolveName);
  if (!criteria || !sum) return null;
  const bottom = Math.min(sum.top + criteria.rowCount, SheetDimensions.maxRows) - 1;
  const right = Math.min(sum.left + criteria.columnCount, SheetDimensions.maxColumns) - 1;
  return CellRange.fromBounds(sum.top, sum.left, bottom, right, sum.sheet);
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
   */
  collect(
    expr: Expr,
    homeSheet: string,
    resolveName: (name: string) => CellRange | null,
    destination: Map<string, CellRef>,
    largeRanges?: RangeRef[],
    sheetsBetween?: (first: string, last: string) => readonly string[] | null,
  ): void {
    const recurse = (e: Expr): void =>
      ReferenceCollector.collect(e, homeSheet, resolveName, destination, largeRanges, sheetsBetween);
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
        if (named) expand(named, homeSheet, destination, largeRanges);
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
        const reshaped = reshapedSumRange(expr.name, expr.args, resolveName);
        if (reshaped) expand(reshaped, homeSheet, destination, largeRanges);
        break;
      }
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
