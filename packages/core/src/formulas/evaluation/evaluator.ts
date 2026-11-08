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
        return (
          this.#missingSheet(expr) ??
          CellValue.error(CellError.value('A range is not valid in this position.'))
        );
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
    if (range.isSingleCell) return this.context.getCell(range.sheet, range.top, range.left);
    return CellValue.error(CellError.value(`The name '${node.name}' is a range, not a single value.`));
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

  enumerateValues(expr: Expr): Iterable<CellValue> {
    const range = expr.kind === 'reference' ? null : this.tryEvaluateRange(expr);
    // Every caller skips empty cells, so the empty margin past the sheet's data is left out.
    return range ? range.usedValues() : [this.evaluate(expr)];
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
      default:
        return null;
    }
  }
}
