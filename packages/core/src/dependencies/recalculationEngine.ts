import type { CellRange } from '../model/cellRange.ts';
import type { Workbook } from '../model/workbook.ts';
import { CellError } from '../values/cellError.ts';
import { CellValue } from '../values/cellValue.ts';
import type { Expr } from '../formulas/ast/expr.ts';
import { systemClock, type CivilDateTime } from '../formulas/evaluation/dateConversion.ts';
import { Evaluator } from '../formulas/evaluation/evaluator.ts';
import { WorkbookEvaluationContext } from '../formulas/evaluation/workbookEvaluationContext.ts';
import { FunctionRegistry } from '../formulas/functions/functionRegistry.ts';
import { Parser } from '../formulas/parsing/parser.ts';
import { CellRef, rangeRefContains, type RangeRef } from './cellRef.ts';
import { DependencyGraph } from './dependencyGraph.ts';
import { ReferenceCollector } from './referenceCollector.ts';

/** The outcome of a recalculation pass. */
export interface RecalculationResult {
  readonly evaluated: number;
  readonly circularCells: readonly CellRef[];
  readonly hasCircularReference: boolean;
}

const result = (evaluated: number, circularCells: readonly CellRef[]): RecalculationResult => ({
  evaluated,
  circularCells,
  hasCircularReference: circularCells.length > 0,
});

/**
 * Recomputes formula cells in dependency order, touching only what an edit affects. Honours the
 * workbook's calculation mode, recomputes volatile functions every pass, and reports circular
 * references as `#REF!` "Circular reference." on the offending cells.
 */
export class RecalculationEngine {
  readonly #workbook: Workbook;
  readonly #functions: FunctionRegistry;
  readonly #clock: () => CivilDateTime;
  readonly #graph = new DependencyGraph();
  /** Formula cells by key: their ref and parsed AST (null when unparseable). */
  readonly #formulas = new Map<string, { ref: CellRef; ast: Expr | null }>();
  readonly #volatile = new Map<string, CellRef>();
  /** Formula cells whose precedents include a range too large to expand into the graph. */
  readonly #largeRangePrecedents = new Map<string, RangeRef[]>();
  readonly #parseCache = new Map<string, Expr | null>();

  constructor(workbook: Workbook, functions?: FunctionRegistry, clock?: () => CivilDateTime) {
    this.#workbook = workbook;
    this.#functions = functions ?? FunctionRegistry.createDefault();
    this.#clock = clock ?? systemClock;
  }

  /** The formula cells that contain a volatile function (TODAY/NOW). */
  get volatileCells(): readonly CellRef[] {
    return [...this.#volatile.values()];
  }

  isVolatile(cell: CellRef): boolean {
    return this.#volatile.has(CellRef.key(cell));
  }

  /** The dependency graph (read-only use). */
  get graph(): DependencyGraph {
    return this.#graph;
  }

  /** Scan the whole workbook, rebuilding the dependency graph from every formula cell. */
  rebuildGraph(): void {
    this.#graph.clear();
    this.#formulas.clear();
    this.#volatile.clear();
    this.#largeRangePrecedents.clear();
    for (const sheet of this.#workbook.worksheets) {
      for (const [row, column, cell] of sheet.enumerateCells()) {
        if (cell.isFormula) this.#registerFormula(CellRef.of(sheet.name, row, column), cell.formula!);
      }
    }
  }

  /** Recompute every formula cell in dependency order. */
  recalculateAll(): RecalculationResult {
    return this.#evaluate(new Set(this.#formulas.keys()));
  }

  /**
   * Update the graph for cells whose content just changed, then recompute them and everything
   * downstream (plus volatile cells). In manual mode nothing is evaluated, but the graph stays
   * current so a later full recalc is correct.
   */
  recordEdits(changed: Iterable<CellRef>): RecalculationResult {
    const seeds = new Map<string, CellRef>();
    for (const cell of changed) {
      this.#syncFormulaState(cell);
      seeds.set(CellRef.key(cell), cell);
    }
    if (this.#workbook.settings.calculationMode === 'manual') return result(0, []);
    const affected = this.#collectAffected(seeds.entries());
    for (const k of this.#volatile.keys()) affected.add(k);
    return this.#evaluate(affected);
  }

  /**
   * Every formula cell that reads any of `cells`, via a stored edge or a summarized large precedent
   * range — the reverse index a move/swap uses to rewrite only the formulas that reference them.
   */
  formulasReferencing(cells: Iterable<CellRef>): CellRef[] {
    const found = new Map<string, CellRef>();
    for (const cell of cells) {
      for (const [k, dependent] of this.#graph.dependentEntries(CellRef.key(cell))) found.set(k, dependent);
      for (const [k, ranges] of this.#largeRangePrecedents) {
        if (ranges.some((r) => rangeRefContains(r, cell))) found.set(k, this.#formulas.get(k)!.ref);
      }
    }
    return [...found.values()];
  }

  /** Recompute volatile cells and their dependents. */
  recalculateVolatile(): RecalculationResult {
    const affected = this.#collectAffected(this.#volatile.entries());
    for (const k of this.#volatile.keys()) affected.add(k);
    return this.#evaluate(affected);
  }

  // ---- graph maintenance ------------------------------------------------

  #syncFormulaState(cell: CellRef): void {
    const model = this.#workbook.findSheet(cell.sheet)?.find(cell.row, cell.column);
    if (model?.isFormula) {
      this.#registerFormula(cell, model.formula!);
      return;
    }
    const key = CellRef.key(cell);
    this.#graph.removeFormula(cell);
    this.#formulas.delete(key);
    this.#volatile.delete(key);
    this.#largeRangePrecedents.delete(key);
  }

  #registerFormula(cell: CellRef, formula: string): void {
    const key = CellRef.key(cell);
    const ast = this.#parseCached(formula);
    this.#formulas.set(key, { ref: cell, ast });
    if (!ast) {
      this.#graph.setPrecedents(cell, []);
      this.#volatile.delete(key);
      this.#largeRangePrecedents.delete(key);
      return;
    }
    const precedents = new Map<string, CellRef>();
    const largeRanges: RangeRef[] = [];
    ReferenceCollector.collect(ast, cell.sheet, (n) => this.#resolveName(n), precedents, largeRanges);
    this.#graph.setPrecedents(cell, precedents.values());
    if (largeRanges.length > 0) this.#largeRangePrecedents.set(key, largeRanges);
    else this.#largeRangePrecedents.delete(key);
    if (ReferenceCollector.isVolatile(ast)) this.#volatile.set(key, cell);
    else this.#volatile.delete(key);
  }

  #parseCached(formula: string): Expr | null {
    if (this.#parseCache.has(formula)) return this.#parseCache.get(formula)!;
    const parsed = Parser.tryParse(formula);
    const ast = parsed.ok ? parsed.expr : null;
    this.#parseCache.set(formula, ast);
    return ast;
  }

  /** Global-scope names only (the graph is built outside any one sheet's scope). */
  #resolveName(name: string): CellRange | null {
    return this.#workbook.findNamedRange(name)?.target ?? null;
  }

  // ---- evaluation -------------------------------------------------------

  #collectAffected(seeds: Iterable<[string, CellRef]>): Set<string> {
    const affected = new Set<string>();
    const queue: [string, CellRef][] = [];
    for (const [key, seed] of seeds) {
      if (this.#formulas.has(key) && !affected.has(key)) {
        affected.add(key);
        queue.push([key, seed]);
      }
      this.#addDependents(key, seed, affected, queue);
    }
    for (let i = 0; i < queue.length; i++) {
      const [key, cell] = queue[i]!;
      this.#addDependents(key, cell, affected, queue);
    }
    return affected;
  }

  /** Enqueue every formula reading `cell`: via a stored edge, or a large range containing it. */
  #addDependents(key: string, cell: CellRef, affected: Set<string>, queue: [string, CellRef][]): void {
    for (const [dk, dependent] of this.#graph.dependentEntries(key)) {
      if (!affected.has(dk)) {
        affected.add(dk);
        queue.push([dk, dependent]);
      }
    }
    for (const [fk, ranges] of this.#largeRangePrecedents) {
      if (!affected.has(fk) && ranges.some((r) => rangeRefContains(r, cell))) {
        affected.add(fk);
        queue.push([fk, this.#formulas.get(fk)!.ref]);
      }
    }
  }

  #evaluate(nodes: Set<string>): RecalculationResult {
    if (nodes.size === 0) return result(0, []);
    const { ordered, circular } = this.#topologicalSort(nodes);
    for (const key of ordered) this.#evaluateCell(key);
    const circularCells: CellRef[] = [];
    for (const key of circular) {
      const ref = this.#refOf(key);
      if (!ref) continue;
      circularCells.push(ref);
      this.#writeResult(ref, CellValue.error(CellError.circular()));
    }
    return result(ordered.length, circularCells);
  }

  #refOf(key: string): CellRef | null {
    return this.#formulas.get(key)?.ref ?? null;
  }

  #evaluateCell(key: string): void {
    const entry = this.#formulas.get(key);
    if (!entry) return;
    if (!entry.ast) {
      this.#writeResult(entry.ref, CellValue.error(CellError.value('The formula could not be parsed.')));
      return;
    }
    const context = new WorkbookEvaluationContext(
      this.#workbook,
      entry.ref.sheet,
      this.#workbook.settings.effectiveLocale,
      this.#clock,
    );
    this.#writeResult(entry.ref, new Evaluator(context, this.#functions).evaluate(entry.ast));
  }

  #writeResult(cell: CellRef, value: CellValue): void {
    const model = this.#workbook.findSheet(cell.sheet)?.find(cell.row, cell.column);
    if (model?.isFormula) model.value = value;
  }

  /** Kahn's algorithm over the node set; nodes left in a cycle come back as `circular`. */
  #topologicalSort(nodes: Set<string>): { ordered: string[]; circular: string[] } {
    const inDegree = new Map<string, number>();
    for (const node of nodes) {
      let degree = 0;
      for (const p of this.#graph.precedentKeys(node)) if (nodes.has(p)) degree++;
      inDegree.set(node, degree);
    }
    const ready: string[] = [];
    for (const [node, degree] of inDegree) if (degree === 0) ready.push(node);
    const ordered: string[] = [];
    for (let i = 0; i < ready.length; i++) {
      const node = ready[i]!;
      ordered.push(node);
      for (const [dk] of this.#graph.dependentEntries(node)) {
        const degree = inDegree.get(dk);
        if (degree === undefined) continue;
        inDegree.set(dk, degree - 1);
        if (degree - 1 === 0) ready.push(dk);
      }
    }
    const circular: string[] = [];
    if (ordered.length !== nodes.size) {
      const placed = new Set(ordered);
      for (const node of nodes) if (!placed.has(node)) circular.push(node);
    }
    return { ordered, circular };
  }
}
