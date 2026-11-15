import type { Expr } from '../formulas/ast/expr.ts';
import { Evaluator } from '../formulas/evaluation/evaluator.ts';
import { Coercion } from '../formulas/evaluation/coercion.ts';
import { ValueComparer } from '../formulas/evaluation/valueComparer.ts';
import { WorkbookEvaluationContext } from '../formulas/evaluation/workbookEvaluationContext.ts';
import { FormulaWriter } from '../formulas/formulaWriter.ts';
import { FunctionRegistry } from '../formulas/functions/functionRegistry.ts';
import { OpenFormulaTranslator } from '../formulas/openFormulaTranslator.ts';
import { Parser } from '../formulas/parsing/parser.ts';
import { ReferenceAdjuster } from '../formulas/references/referenceAdjuster.ts';
import { CellError } from '../values/cellError.ts';
import { CellValue } from '../values/cellValue.ts';
import type { Workbook } from '../model/workbook.ts';
import type { CellStyle, CellStyleMap } from './cellStyle.ts';

/** How a cell's value is compared in a `cell-content()` condition. */
export type ConditionOperator = '<' | '>' | '<=' | '>=' | '=' | '!=';

/** A `style:condition`, parsed. Operands are formula expressions anchored at the map's base cell. */
export type MapCondition =
  | { readonly kind: 'compare'; readonly operator: ConditionOperator; readonly value: Expr }
  | { readonly kind: 'between' | 'notBetween'; readonly low: Expr; readonly high: Expr }
  | { readonly kind: 'formula'; readonly formula: Expr };

const OPERATORS: readonly ConditionOperator[] = ['<=', '>=', '!=', '<', '>', '='];

/**
 * Parse an ODF `style:condition` (`cell-content()>5`, `cell-content-is-between(1,[.B1])`,
 * `is-true-formula([.A1]>0)`). Null for anything else, which then never applies.
 */
export function parseMapCondition(condition: string): MapCondition | null {
  const text = condition.trim().replace(/^of:/, '');
  const compare = /^cell-content\(\)\s*/.exec(text);
  if (compare) {
    const rest = text.slice(compare[0].length);
    const operator = OPERATORS.find((op) => rest.startsWith(op));
    if (!operator) return null;
    const value = parseOperand(rest.slice(operator.length));
    return value ? { kind: 'compare', operator, value } : null;
  }
  const between = /^cell-content-is-(not-)?between\((.*)\)$/s.exec(text);
  if (between) {
    const args = splitArguments(between[2]!);
    if (args.length !== 2) return null;
    const low = parseOperand(args[0]!);
    const high = parseOperand(args[1]!);
    if (!low || !high) return null;
    return { kind: between[1] ? 'notBetween' : 'between', low, high };
  }
  const formula = /^is-true-formula\((.*)\)$/s.exec(text);
  if (formula) {
    const expr = parseOperand(formula[1]!);
    return expr ? { kind: 'formula', formula: expr } : null;
  }
  return null;
}

/** Write a condition back in the ODF form Calc reads. */
export function writeMapCondition(condition: MapCondition): string {
  switch (condition.kind) {
    case 'compare':
      return `cell-content()${condition.operator}${writeOperand(condition.value)}`;
    case 'between':
    case 'notBetween':
      return `cell-content-is-${condition.kind === 'notBetween' ? 'not-' : ''}between(${writeOperand(
        condition.low,
      )},${writeOperand(condition.high)})`;
    case 'formula':
      return `is-true-formula(${writeOperand(condition.formula)})`;
  }
}

/** An operand in OpenFormula syntax (`5`, `"x"`, `[.A1]`), parsed. */
function parseOperand(odf: string): Expr | null {
  if (!odf.trim()) return null;
  const parsed = Parser.tryParse(OpenFormulaTranslator.fromOpenFormula(`of:=${odf.trim()}`));
  return parsed.ok ? parsed.expr : null;
}

/** An operand written in OpenFormula syntax. */
export function writeOperand(expr: Expr): string {
  return OpenFormulaTranslator.toOpenFormula(FormulaWriter.write(expr)).slice('of:='.length);
}

/** Split at top-level commas or semicolons, outside quotes, brackets and parentheses. */
function splitArguments(text: string): string[] {
  const out: string[] = [];
  let depth = 0;
  let quoted = false;
  let start = 0;
  for (let i = 0; i < text.length; i++) {
    const c = text[i]!;
    if (c === '"') quoted = !quoted;
    else if (quoted) continue;
    else if (c === '(' || c === '[') depth++;
    else if (c === ')' || c === ']') depth--;
    else if ((c === ',' || c === ';') && depth === 0) {
      out.push(text.slice(start, i));
      start = i + 1;
    }
  }
  out.push(text.slice(start));
  return out;
}

const parsedConditions = new Map<string, MapCondition | null>();
const PARSED_MAX = 10_000;

function parsedCondition(condition: string): MapCondition | null {
  if (parsedConditions.has(condition)) return parsedConditions.get(condition)!;
  const parsed = parseMapCondition(condition);
  if (parsedConditions.size >= PARSED_MAX) parsedConditions.clear();
  parsedConditions.set(condition, parsed);
  return parsed;
}

/** The base cell of a map: its sheet (null when unqualified), row and column. */
function baseCell(map: CellStyleMap): { sheet: string | null; row: number; column: number } | null {
  if (map.baseCellAddress === null) return null;
  const expr = parseOperand(`[${map.baseCellAddress}]`);
  return expr?.kind === 'reference'
    ? { sheet: expr.address.sheet, row: expr.address.row, column: expr.address.column }
    : null;
}

/** Text compares only with text (an empty operand takes the cell's kind; errors fail elsewhere). */
const mismatched = (cell: CellValue, operand: CellValue): boolean =>
  operand.kind !== 'empty' &&
  operand.kind !== 'error' &&
  (cell.kind === 'text') !== (operand.kind === 'text');

/**
 * Evaluates conditional maps (legacy `style:map` conditional formats) for cells of one workbook.
 * Relative references in a condition are relative to the map's base cell, so a rule written for
 * A1 reads B1 when it checks B1, as in Calc.
 */
export class ConditionalStyles {
  readonly #workbook: Workbook;
  readonly #contexts = new Map<string, WorkbookEvaluationContext>();
  readonly #bases = new Map<string, ReturnType<typeof baseCell>>();

  constructor(workbook: Workbook) {
    this.#workbook = workbook;
  }

  /**
   * The style a cell shows: the applied style of the first map whose condition holds, else
   * `style` itself. `sheet` is the cell's sheet, which unqualified references resolve against.
   */
  resolve(style: CellStyle, sheet: string, row: number, column: number): CellStyle {
    const maps = style.conditionalMaps;
    if (!maps?.length) return style;
    for (const map of maps) {
      if (map.applyStyle === null) continue;
      if (this.holds(map, sheet, row, column)) return map.applyStyle;
    }
    return style;
  }

  /** Whether `map`'s condition holds for the cell at (row, column) of `sheet`. */
  holds(map: CellStyleMap, sheet: string, row: number, column: number): boolean {
    const condition = parsedCondition(map.condition);
    if (!condition) return false;
    let base = this.#bases.get(map.baseCellAddress ?? '');
    if (base === undefined) {
      base = baseCell(map);
      this.#bases.set(map.baseCellAddress ?? '', base);
    }
    const dr = base ? row - base.row : 0;
    const dc = base ? column - base.column : 0;
    const evaluate = (expr: Expr): CellValue =>
      this.#evaluator(sheet).evaluate(dr || dc ? ReferenceAdjuster.forCopy(expr, dc, dr) : expr);

    if (condition.kind === 'formula') {
      const result = evaluate(condition.formula);
      const truth = Coercion.toBoolean(result);
      return !(truth instanceof CellError) && truth;
    }
    const value = this.#workbook.findSheet(sheet)?.getValue(row, column) ?? CellValue.empty;
    // An empty cell is not "less than 5": a rule describes what is in the cell.
    if (value.kind === 'empty' || value.kind === 'error') return false;
    if (condition.kind === 'compare') {
      const operand = evaluate(condition.value);
      // As in Calc, text is neither more nor less than a number: it is only "not equal" to one.
      if (mismatched(value, operand)) return condition.operator === '!=';
      const c = ValueComparer.compare(value, operand);
      if (c instanceof CellError) return false;
      switch (condition.operator) {
        case '<':
          return c < 0;
        case '>':
          return c > 0;
        case '<=':
          return c <= 0;
        case '>=':
          return c >= 0;
        case '=':
          return c === 0;
        case '!=':
          return c !== 0;
      }
    }
    const lowValue = evaluate(condition.low);
    const highValue = evaluate(condition.high);
    if (mismatched(value, lowValue) || mismatched(value, highValue)) return condition.kind === 'notBetween';
    const low = ValueComparer.compare(value, lowValue);
    const high = ValueComparer.compare(value, highValue);
    if (low instanceof CellError || high instanceof CellError) return false;
    const inside = low >= 0 && high <= 0;
    return condition.kind === 'between' ? inside : !inside;
  }

  #evaluator(sheet: string): Evaluator {
    let context = this.#contexts.get(sheet);
    if (!context) {
      context = new WorkbookEvaluationContext(this.#workbook, sheet);
      this.#contexts.set(sheet, context);
    }
    return new Evaluator(context, FunctionRegistry.standard);
  }
}
