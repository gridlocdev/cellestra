import { CellAddress } from '../../model/cellAddress.ts';
import { CellRange } from '../../model/cellRange.ts';
import { SheetDimensions } from '../../model/sheetDimensions.ts';
import { Expr, type SheetSpanExpr } from '../ast/expr.ts';

type MapAddress = (a: CellAddress) => CellAddress | null;
type MapRange = (r: CellRange) => CellRange | null;
/** A 3D range's fix-up. Left out, a 3D range stays as it is. */
type MapSpan = (span: SheetSpanExpr) => Expr;

const keepSpan: MapSpan = (span) => span;

const sameSheet = (a: string, b: string): boolean => a.toLowerCase() === b.toLowerCase();

function rewrite(expr: Expr, mapAddress: MapAddress, mapRange: MapRange, mapSpan = keepSpan): Expr {
  const recurse = (e: Expr): Expr => rewrite(e, mapAddress, mapRange, mapSpan);
  switch (expr.kind) {
    case 'reference': {
      const mapped = mapAddress(expr.address);
      return mapped ? Expr.reference(mapped) : Expr.refError;
    }
    case 'range': {
      const mapped = mapRange(expr.range);
      return mapped ? Expr.range(mapped) : Expr.refError;
    }
    case 'sheetSpan':
      return mapSpan(expr);
    case 'unary':
      return Expr.unary(expr.operator, recurse(expr.operand));
    case 'binary':
      return Expr.binary(expr.operator, recurse(expr.left), recurse(expr.right));
    case 'call':
      return Expr.call(expr.name, expr.args.map(recurse));
    default:
      return expr; // literals, names, ref-errors are unchanged
  }
}

const withCoord = (a: CellAddress, rowAxis: boolean, value: number): CellAddress =>
  rowAxis
    ? new CellAddress(a.column, value, a.columnAbsolute, a.rowAbsolute, a.sheet)
    : new CellAddress(value, a.row, a.columnAbsolute, a.rowAbsolute, a.sheet);

function shiftAddress(a: CellAddress, columnDelta: number, rowDelta: number): CellAddress | null {
  const column = a.columnAbsolute ? a.column : a.column + columnDelta;
  const row = a.rowAbsolute ? a.row : a.row + rowDelta;
  if (!SheetDimensions.isValidColumn(column) || !SheetDimensions.isValidRow(row)) return null;
  return new CellAddress(column, row, a.columnAbsolute, a.rowAbsolute, a.sheet);
}

/**
 * A one-axis insert or delete mapping old line indices to new ones. Insert(at, n) pushes lines
 * at/after `at` down by n; Delete(at, n) removes lines at … at+n−1 (mapped to null) and pulls later
 * lines up by n.
 */
interface AxisEdit {
  readonly insert: boolean;
  readonly at: number;
  readonly count: number;
  /** Exclusive upper bound for inserts. */
  readonly limit: number;
}

function axisMap(e: AxisEdit, coord: number): number | null {
  if (e.insert) {
    const shifted = coord >= e.at ? coord + e.count : coord;
    // An insert that pushes a reference past the sheet edge has nowhere to land → #REF!.
    return shifted >= e.limit ? null : shifted;
  }
  if (coord < e.at) return coord;
  return coord < e.at + e.count ? null : coord - e.count;
}

/** The smallest and largest old indices in [from, to] that survive, or null if none do. */
function survivingExtremes(e: AxisEdit, from: number, to: number): [number, number] | null {
  if (e.insert) return [from, to];
  const deleteEnd = e.at + e.count;
  const min = from < e.at ? from : Math.max(from, deleteEnd) <= to ? Math.max(from, deleteEnd) : null;
  const max = to >= deleteEnd ? to : to < e.at ? to : e.at - 1 >= from ? e.at - 1 : null;
  return min !== null && max !== null ? [min, max] : null;
}

function axisRewrite(
  expr: Expr,
  homeSheet: string,
  editedSheet: string,
  rowAxis: boolean,
  edit: AxisEdit,
): Expr {
  const affects = (sheet: string | null): boolean => sameSheet(sheet ?? homeSheet, editedSheet);

  const mapAddress: MapAddress = (a) => {
    if (!affects(a.sheet)) return a;
    const mapped = axisMap(edit, rowAxis ? a.row : a.column);
    return mapped === null ? null : withCoord(a, rowAxis, mapped);
  };

  const mapRange: MapRange = (range) => {
    if (!affects(range.sheet)) return range;
    // Work on the normalized span; the whole range is #REF! only when every line of it is deleted.
    const min = rowAxis ? range.top : range.left;
    const max = rowAxis ? range.bottom : range.right;
    const lastLine = (rowAxis ? SheetDimensions.maxRows : SheetDimensions.maxColumns) - 1;
    // A whole column (A:A) or whole row (1:1) still spans the whole sheet after any insert/delete.
    if (min === 0 && max === lastLine) return range;
    const surviving = survivingExtremes(edit, min, max);
    if (!surviving) return null;
    // An insert can push a corner past the edge: the lower corner pushed off → #REF!; only the
    // upper corner spilling over → clamp it to the last line (the range shrinks).
    const newMin = axisMap(edit, surviving[0]);
    if (newMin === null) return null;
    const newMax = axisMap(edit, surviving[1]) ?? lastLine;
    // Re-attach each corner's $ flags: the corner at the span's min gets newMin.
    const corner = (c: CellAddress): CellAddress =>
      withCoord(c, rowAxis, (rowAxis ? c.row : c.column) === min ? newMin : newMax);
    return new CellRange(corner(range.start), corner(range.end));
  };

  return rewrite(expr, mapAddress, mapRange);
}

const insert = (at: number, count: number, limit: number): AxisEdit => ({ insert: true, at, count, limit });
const del = (at: number, count: number): AxisEdit => ({
  insert: false,
  at,
  count,
  limit: Number.MAX_SAFE_INTEGER,
});

/**
 * Where a line lands when lines `first`…`last` move one step by `delta` (-1 or 1): the moved lines
 * shift by `delta`, and the line they step onto takes the free place at the other end.
 */
export function lineMoveCoord(coord: number, first: number, last: number, delta: -1 | 1): number {
  if (coord >= first && coord <= last) return coord + delta;
  if (delta < 0 && coord === first - 1) return last;
  if (delta > 0 && coord === last + 1) return first;
  return coord;
}

/**
 * How the span `lo`…`hi` fares in that move: it `moves` (wholly inside the moved lines, or the
 * swapped line alone), `stays` (clear of both, or covering both), or is `split` across the edge.
 */
export function lineMoveSpan(
  lo: number,
  hi: number,
  first: number,
  last: number,
  delta: -1 | 1,
): 'moves' | 'stays' | 'split' {
  const other = delta < 0 ? first - 1 : last + 1;
  const start = Math.min(first, other);
  const end = Math.max(last, other);
  if (hi < start || lo > end || (lo <= start && hi >= end)) return 'stays';
  if ((lo >= first && hi <= last) || (lo === other && hi === other)) return 'moves';
  return 'split';
}

/**
 * Rewrites the references inside a formula AST — the one place reference fix-ups live. On
 * copy/fill, relative refs shift and absolute (`$`) refs stay. On insert/delete row/column, refs move
 * with their target and refs into deleted cells become `#REF!`. Serialize with `FormulaWriter`.
 */
export const ReferenceAdjuster = {
  /** Copy/fill from one position to another: shift relative parts by the delta. */
  forCopy(expr: Expr, columnDelta: number, rowDelta: number): Expr {
    const mapRange = (r: CellRange): CellRange | null => {
      const s = shiftAddress(r.start, columnDelta, rowDelta);
      const e = shiftAddress(r.end, columnDelta, rowDelta);
      return s && e ? new CellRange(s, e) : null;
    };
    return rewrite(
      expr,
      (a) => shiftAddress(a, columnDelta, rowDelta),
      mapRange,
      (span) => {
        const mapped = mapRange(span.range);
        return mapped ? Expr.sheetSpan(mapped, span.lastSheet) : Expr.refError;
      },
    );
  },

  forInsertRows: (expr: Expr, homeSheet: string, editedSheet: string, atRow: number, count: number): Expr =>
    axisRewrite(expr, homeSheet, editedSheet, true, insert(atRow, count, SheetDimensions.maxRows)),

  forDeleteRows: (expr: Expr, homeSheet: string, editedSheet: string, atRow: number, count: number): Expr =>
    axisRewrite(expr, homeSheet, editedSheet, true, del(atRow, count)),

  forInsertColumns: (
    expr: Expr,
    homeSheet: string,
    editedSheet: string,
    atColumn: number,
    count: number,
  ): Expr =>
    axisRewrite(expr, homeSheet, editedSheet, false, insert(atColumn, count, SheetDimensions.maxColumns)),

  forDeleteColumns: (
    expr: Expr,
    homeSheet: string,
    editedSheet: string,
    atColumn: number,
    count: number,
  ): Expr => axisRewrite(expr, homeSheet, editedSheet, false, del(atColumn, count)),

  /**
   * Lines `first`…`last` on `sheet` moved one step (`delta` is -1 or 1) along the row or column axis,
   * swapping places with the line next to them. References follow their cells, `$`-absolute ones too.
   * A range follows when it sits wholly in the moved lines or is the swapped line alone; any other
   * range (covering both, or neither) keeps its bounds.
   */
  forLineMove(
    expr: Expr,
    homeSheet: string,
    sheet: string,
    rowAxis: boolean,
    first: number,
    last: number,
    delta: -1 | 1,
  ): Expr {
    const onSheet = (s: string | null): boolean => sameSheet(s ?? homeSheet, sheet);
    const coord = (a: CellAddress): number => (rowAxis ? a.row : a.column);
    const mapAddress = (a: CellAddress): CellAddress =>
      onSheet(a.sheet) ? withCoord(a, rowAxis, lineMoveCoord(coord(a), first, last, delta)) : a;
    const mapRange = (r: CellRange): CellRange => {
      if (!onSheet(r.sheet)) return r;
      const lo = rowAxis ? r.top : r.left;
      const hi = rowAxis ? r.bottom : r.right;
      if (lineMoveSpan(lo, hi, first, last, delta) === 'moves')
        return new CellRange(mapAddress(r.start), mapAddress(r.end));
      return r;
    };
    return rewrite(expr, mapAddress, mapRange);
  },

  /**
   * A row moved during a sort: a relative reference into the sorted data rows shifts by the same
   * delta the formula's row moved. Absolute, cross-sheet, and out-of-data refs stay put.
   */
  forSortMove(expr: Expr, rowDelta: number, dataTop: number, dataBottom: number): Expr {
    const map = (a: CellAddress): CellAddress => {
      if (a.sheet !== null || a.rowAbsolute || a.row < dataTop || a.row > dataBottom) return a;
      const newRow = a.row + rowDelta;
      return SheetDimensions.isValidRow(newRow)
        ? new CellAddress(a.column, newRow, a.columnAbsolute, a.rowAbsolute, a.sheet)
        : a;
    };
    return rewrite(expr, map, (r) => new CellRange(map(r.start), map(r.end)));
  },

  /**
   * Cells moved (cut-paste or nudge): references into `sourceRange` on `sourceSheet` follow the
   * cells — shifted by the deltas and re-homed on `destSheet` — even `$`-absolute ones. A range
   * follows only when wholly inside the moved block. `homeSheet` resolves unqualified refs.
   */
  forMove(
    expr: Expr,
    homeSheet: string,
    sourceSheet: string,
    sourceRange: CellRange,
    destSheet: string,
    rowDelta: number,
    columnDelta: number,
  ): Expr {
    const destSheetRef = sameSheet(destSheet, homeSheet) ? null : destSheet;
    const inSource = (a: CellAddress): boolean =>
      sameSheet(a.sheet ?? homeSheet, sourceSheet) &&
      a.column >= sourceRange.left &&
      a.column <= sourceRange.right &&
      a.row >= sourceRange.top &&
      a.row <= sourceRange.bottom;
    const moved = (a: CellAddress): CellAddress =>
      new CellAddress(
        a.column + columnDelta,
        a.row + rowDelta,
        a.columnAbsolute,
        a.rowAbsolute,
        destSheetRef,
      );
    return rewrite(
      expr,
      (a) => (inSource(a) ? moved(a) : a),
      (r) => (inSource(r.start) && inSource(r.end) ? new CellRange(moved(r.start), moved(r.end)) : r),
    );
  },

  /**
   * Two equal-shaped, non-overlapping blocks swapped: block A is `blockA`, block B is A translated by
   * the deltas. References into A go to B and into B go to A, in a single pass. A range follows only
   * when wholly within one block.
   */
  forSwap(
    expr: Expr,
    homeSheet: string,
    sheet: string,
    blockA: CellRange,
    rowDelta: number,
    columnDelta: number,
  ): Expr {
    const aTop = blockA.top;
    const aLeft = blockA.left;
    const aBottom = blockA.bottom;
    const aRight = blockA.right;
    const bTop = aTop + rowDelta;
    const bLeft = aLeft + columnDelta;
    const bBottom = aBottom + rowDelta;
    const bRight = aRight + columnDelta;
    const onSheet = (s: string | null): boolean => sameSheet(s ?? homeSheet, sheet);
    const inA = (row: number, col: number): boolean =>
      row >= aTop && row <= aBottom && col >= aLeft && col <= aRight;
    const inB = (row: number, col: number): boolean =>
      row >= bTop && row <= bBottom && col >= bLeft && col <= bRight;
    const shift = (a: CellAddress, dr: number, dc: number): CellAddress =>
      new CellAddress(a.column + dc, a.row + dr, a.columnAbsolute, a.rowAbsolute, a.sheet);

    const mapAddress = (a: CellAddress): CellAddress => {
      if (!onSheet(a.sheet)) return a;
      if (inA(a.row, a.column)) return shift(a, rowDelta, columnDelta);
      return inB(a.row, a.column) ? shift(a, -rowDelta, -columnDelta) : a;
    };
    const mapRange = (r: CellRange): CellRange => {
      if (onSheet(r.sheet) && inA(r.top, r.left) && inA(r.bottom, r.right)) {
        return new CellRange(shift(r.start, rowDelta, columnDelta), shift(r.end, rowDelta, columnDelta));
      }
      if (onSheet(r.sheet) && inB(r.top, r.left) && inB(r.bottom, r.right)) {
        return new CellRange(shift(r.start, -rowDelta, -columnDelta), shift(r.end, -rowDelta, -columnDelta));
      }
      return r;
    };
    return rewrite(expr, mapAddress, mapRange);
  },

  /**
   * A worksheet renamed: every reference that names `oldName` explicitly now names `newName`.
   * Unqualified references resolve against their own sheet and need no change.
   */
  forRenameSheet(expr: Expr, oldName: string, newName: string): Expr {
    const renamed = (a: CellAddress): CellAddress =>
      a.sheet !== null && sameSheet(a.sheet, oldName) ? a.withSheet(newName) : a;
    const renamedRange = (r: CellRange): CellRange =>
      r.sheet !== null && sameSheet(r.sheet, oldName) ? new CellRange(renamed(r.start), renamed(r.end)) : r;
    return rewrite(expr, renamed, renamedRange, (span) =>
      Expr.sheetSpan(renamedRange(span.range), sameSheet(span.lastSheet, oldName) ? newName : span.lastSheet),
    );
  },

  /**
   * Give every unqualified reference an explicit `sheet`, so the formula still means the same cells
   * after it moves to another sheet (a cut pasted across sheets).
   */
  qualify(expr: Expr, sheet: string): Expr {
    const own = (a: CellAddress): CellAddress => (a.sheet === null ? a.withSheet(sheet) : a);
    return rewrite(expr, own, (r) => (r.sheet === null ? new CellRange(own(r.start), own(r.end)) : r));
  },

  /**
   * A worksheet deleted: any reference whose effective sheet is the deleted one becomes #REF!. A 3D
   * range loses only that sheet: deleting a sheet inside it changes nothing, deleting an end moves
   * the end to the next sheet inward, as in Calc. That needs `sheetOrder`, the tab order before the
   * delete; without it a 3D range that ends on the deleted sheet becomes #REF!.
   */
  forDeleteSheet(expr: Expr, deletedSheet: string, homeSheet: string, sheetOrder?: readonly string[]): Expr {
    const affected = (sheet: string | null): boolean => sameSheet(sheet ?? homeSheet, deletedSheet);
    const narrow = (span: SheetSpanExpr): Expr => {
      const first = span.range.sheet!;
      if (!affected(first) && !affected(span.lastSheet)) return span;
      const index = (name: string): number => sheetOrder?.findIndex((s) => sameSheet(s, name)) ?? -1;
      const a = index(first);
      const b = index(span.lastSheet);
      if (a < 0 || b < 0) return Expr.refError;
      const left = sheetOrder!.slice(Math.min(a, b), Math.max(a, b) + 1).filter((s) => !affected(s));
      if (left.length === 0) return Expr.refError;
      const range = span.range.withSheet(left[0]!);
      return left.length === 1 ? Expr.range(range) : Expr.sheetSpan(range, left[left.length - 1]!);
    };
    return rewrite(
      expr,
      (a) => (affected(a.sheet) ? null : a),
      (r) => (affected(r.sheet) ? null : r),
      narrow,
    );
  },
} as const;
