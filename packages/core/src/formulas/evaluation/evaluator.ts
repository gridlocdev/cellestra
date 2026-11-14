import { CellRange } from '../../model/cellRange.ts';
import { CellError } from '../../values/cellError.ts';
import { CellValue } from '../../values/cellValue.ts';
import type { BinaryExpr, Expr, FunctionCallExpr, NameExpr, UnaryExpr } from '../ast/expr.ts';
import type { FunctionEvaluator } from '../functions/formulaFunction.ts';
import { acceptsArgCount } from '../functions/functionMetadata.ts';
import type { FunctionRegistry } from '../functions/functionRegistry.ts';
import { RangeValues } from '../functions/rangeValues.ts';
import { Coercion } from './coercion.ts';
import type { EvaluationContext } from './evaluationContext.ts';
import { approxAdd, approxSub, ValueComparer } from './valueComparer.ts';

/** Guard against non-finite results, which surface as #NUM! (as in LibreOffice Calc). */
function finite(value: number): CellValue {
  return Number.isFinite(value)
    ? CellValue.number(value)
    : CellValue.error(CellError.number('The result is not a finite number.'));
}

/** Functions that receive error arguments as values to inspect, rather than failing on them. */
const TRAPS_ERRORS = new Set(['IFERROR']);

/** Both operands as numbers, or the error value to return. */
function numbers(left: CellValue, right: CellValue): [number, number] | CellValue {
  if (left.kind === 'error') return left;
  if (right.kind === 'error') return right;
  const a = Coercion.toNumber(left);
  if (a instanceof CellError) return CellValue.error(a);
  const b = Coercion.toNumber(right);
  if (b instanceof CellError) return CellValue.error(b);
  return [a, b];
}

function arithmetic(left: CellValue, right: CellValue, op: (a: number, b: number) => number): CellValue {
  const n = numbers(left, right);
  return Array.isArray(n) ? finite(op(n[0], n[1])) : n;
}

function divide(left: CellValue, right: CellValue): CellValue {
  const n = numbers(left, right);
  if (!Array.isArray(n)) return n;
  if (n[1] === 0) return CellValue.error(CellError.divideByZero());
  return finite(n[0] / n[1]);
}

function concat(left: CellValue, right: CellValue): CellValue {
  const a = Coercion.toText(left);
  if (a instanceof CellError) return CellValue.error(a);
  const b = Coercion.toText(right);
  if (b instanceof CellError) return CellValue.error(b);
  return CellValue.text(a + b);
}

/**
 * Walks a parsed formula and produces a `CellValue`. Errors are values, not exceptions, so they flow
 * through arithmetic and out to the cell. The same instance serves functions as their evaluator.
 */
export class Evaluator implements FunctionEvaluator {
  readonly context: EvaluationContext;
  readonly #functions: FunctionRegistry;

  constructor(context: EvaluationContext, functions: FunctionRegistry) {
    this.context = context;
    this.#functions = functions;
  }

  evaluate(expr: Expr): CellValue {
    switch (expr.kind) {
      case 'literal':
        return expr.value;
      case 'reference':
        return this.context.getCell(expr.address.sheet, expr.address.row, expr.address.column);
      case 'range':
        return this.#missingSheet(expr) ?? this.intersect(expr.range);
      case 'name':
        return this.#name(expr);
      case 'refError':
        return CellValue.error(CellError.reference('A referenced cell was deleted.'));
      case 'unary':
        return this.#unary(expr);
      case 'binary':
        return this.#binary(expr);
      case 'call':
        return this.#call(expr);
    }
  }

  #name(node: NameExpr): CellValue {
    const range = this.context.resolveName(node.name);
    if (!range) return CellValue.error(CellError.name(`Unknown name '${node.name}'.`));
    return this.intersect(range);
  }

  intersect(range: CellRange): CellValue {
    if (range.isSingleCell) return this.context.getCell(range.sheet, range.top, range.left);
    const at = this.context.formulaCell;
    if (at) {
      if (range.left === range.right && at.row >= range.top && at.row <= range.bottom)
        return this.context.getCell(range.sheet, at.row, range.left);
      if (range.top === range.bottom && at.column >= range.left && at.column <= range.right)
        return this.context.getCell(range.sheet, range.top, at.column);
    }
    return CellValue.error(
      CellError.value('A range where one value is needed must share a row or column with the formula.'),
    );
  }

  #unary(node: UnaryExpr): CellValue {
    const operand = this.evaluate(node.operand);
    if (operand.kind === 'error') return operand;
    const value = Coercion.toNumber(operand);
    if (value instanceof CellError) return CellValue.error(value);
    switch (node.operator) {
      case 'negate':
        return CellValue.number(-value);
      case 'plus':
        return CellValue.number(value);
      case 'percent':
        return CellValue.number(value / 100);
    }
  }

  #binary(node: BinaryExpr): CellValue {
    if (node.operator === 'union' || node.operator === 'intersect') return this.#referenceValue(node);
    const left = this.evaluate(node.left);
    const right = this.evaluate(node.right);
    switch (node.operator) {
      case 'add':
        return arithmetic(left, right, approxAdd);
      case 'subtract':
        return arithmetic(left, right, approxSub);
      case 'multiply':
        return arithmetic(left, right, (a, b) => a * b);
      case 'divide':
        return divide(left, right);
      case 'power':
        return arithmetic(left, right, Math.pow);
      case 'concat':
        return concat(left, right);
      case 'equal':
        return ValueComparer.equal(left, right);
      case 'notEqual':
        return ValueComparer.notEqual(left, right);
      case 'less':
        return ValueComparer.less(left, right);
      case 'lessEqual':
        return ValueComparer.lessEqual(left, right);
      case 'greater':
        return ValueComparer.greater(left, right);
      case 'greaterEqual':
        return ValueComparer.greaterEqual(left, right);
    }
  }

  /** A union or intersection where one value is needed: its one area, intersected with the formula's cell. */
  #referenceValue(node: BinaryExpr): CellValue {
    const areas = this.#areas(node);
    if (!Array.isArray(areas)) return areas;
    if (areas.length !== 1)
      return CellValue.error(CellError.value('A list of ranges cannot be used where one value is needed.'));
    return this.intersect(areas[0]!);
  }

  /**
   * The areas a reference expression names: one for a reference, range, name or INDEX, several for
   * a union, and the shared cells of each pair for an intersection. An error value when the
   * expression is not a reference (#VALUE!), names a missing sheet (#REF!), or the intersection is
   * empty (#NULL!).
   */
  #areas(expr: Expr): CellRange[] | CellValue {
    switch (expr.kind) {
      case 'reference':
        return this.#missingSheet(expr) ?? [new CellRange(expr.address)];
      case 'range':
        return this.#missingSheet(expr) ?? [expr.range];
      case 'name': {
        const named = this.context.resolveName(expr.name);
        if (!named) return CellValue.error(CellError.name(`Unknown name '${expr.name}'.`));
        return this.#missingSheet(expr) ?? [named];
      }
      case 'refError':
        return CellValue.error(CellError.reference('A referenced cell was deleted.'));
      case 'call': {
        const fn = this.#functions.get(expr.name);
        if (!fn?.reference || !acceptsArgCount(fn.metadata, expr.args.length)) break;
        const target = fn.reference(expr.args, this);
        return target instanceof CellRange ? [target] : target;
      }
      case 'binary': {
        if (expr.operator !== 'union' && expr.operator !== 'intersect') break;
        const left = this.#areas(expr.left);
        if (!Array.isArray(left)) return left;
        const right = this.#areas(expr.right);
        if (!Array.isArray(right)) return right;
        if (expr.operator === 'union') return [...left, ...right];
        const shared: CellRange[] = [];
        for (const a of left)
          for (const b of right) {
            const both = this.#intersection(a, b);
            if (both) shared.push(both);
          }
        return shared.length > 0 ? shared : CellValue.error(CellError.null());
      }
      default:
        break;
    }
    return CellValue.error(CellError.value('A reference operator needs references on both sides.'));
  }

  /** The cells two ranges share (on the same sheet), or null. */
  #intersection(a: CellRange, b: CellRange): CellRange | null {
    const sheetOf = (r: CellRange): string => (r.sheet ?? this.context.currentSheet).toLowerCase();
    if (sheetOf(a) !== sheetOf(b) || !a.intersects(b)) return null;
    return CellRange.fromBounds(
      Math.max(a.top, b.top),
      Math.max(a.left, b.left),
      Math.min(a.bottom, b.bottom),
      Math.min(a.right, b.right),
      a.sheet ?? b.sheet,
    );
  }

  #call(node: FunctionCallExpr): CellValue {
    const fn = this.#functions.get(node.name);
    if (!fn) return CellValue.error(CellError.name(`Unknown function '${node.name}'.`));
    if (!acceptsArgCount(fn.metadata, node.args.length)) {
      return CellValue.error(
        CellError.value(`${node.name} does not accept ${node.args.length} argument(s).`),
      );
    }
    // A range on a sheet that no longer exists is one #REF!, not a million empty cells to scan.
    if (!TRAPS_ERRORS.has(node.name)) {
      for (const arg of node.args) {
        const missing = this.#missingSheet(arg);
        if (missing) return missing;
      }
    }
    return fn.invoke(node.args, this);
  }

  /** #REF! when `expr` is a reference, range or name on a sheet that does not exist; else null. */
  #missingSheet(expr: Expr): CellValue | null {
    let sheet: string | null;
    if (expr.kind === 'range') sheet = expr.range.sheet;
    else if (expr.kind === 'reference') sheet = expr.address.sheet;
    else if (expr.kind === 'name') sheet = this.context.resolveName(expr.name)?.sheet ?? null;
    else return null;
    if (sheet === null || this.context.hasSheet?.(sheet) !== false) return null;
    return CellValue.error(CellError.reference(`Unknown sheet '${sheet}'.`));
  }

  // ---- FunctionEvaluator ----------------------------------------------

  *enumerateValues(expr: Expr): Iterable<CellValue> {
    const areas = expr.kind === 'reference' ? null : this.tryEvaluateAreas(expr);
    if (!areas) {
      yield this.evaluate(expr);
      return;
    }
    // Every caller skips empty cells, so the empty margin past the sheet's data is left out.
    for (const area of areas) yield* area.usedValues();
  }

  tryEvaluateAreas(expr: Expr): readonly RangeValues[] | null {
    if (expr.kind !== 'binary') {
      const range = this.tryEvaluateRange(expr);
      return range ? [range] : null;
    }
    if (expr.operator !== 'union' && expr.operator !== 'intersect') return null;
    const areas = this.#areas(expr);
    return Array.isArray(areas) ? areas.map((a) => new RangeValues(a, this.context)) : null;
  }

  tryEvaluateRange(expr: Expr): RangeValues | null {
    switch (expr.kind) {
      case 'range':
        return new RangeValues(expr.range, this.context);
      case 'reference':
        return new RangeValues(new CellRange(expr.address), this.context);
      case 'name': {
        const named = this.context.resolveName(expr.name);
        return named ? new RangeValues(named, this.context) : null;
      }
      case 'binary': {
        // A union of one area, or an intersection that leaves one, is a plain range.
        const areas = this.tryEvaluateAreas(expr);
        return areas?.length === 1 ? areas[0]! : null;
      }
      case 'call': {
        // A reference-returning function (INDEX); its errors surface when it is evaluated instead.
        const fn = this.#functions.get(expr.name);
        if (!fn?.reference || !acceptsArgCount(fn.metadata, expr.args.length)) return null;
        const target = fn.reference(expr.args, this);
        return target instanceof CellRange ? new RangeValues(target, this.context) : null;
      }
      default:
        return null;
    }
  }
}
