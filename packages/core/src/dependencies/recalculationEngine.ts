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
import { CellRef, type RangeRef } from './cellRef.ts';
import { DependencyGraph } from './dependencyGraph.ts';
import { LargeRangeIndex, type RangeSearchSession } from './largeRangeIndex.ts';
import { ReferenceCollector } from './referenceCollector.ts';

/** The outcome of a recalculation pass. */
export interface RecalculationResult {
  readonly evaluated: number;
  readonly circularCells: readonly CellRef[];
  readonly hasCircularReference: boolean;
  /** The formula cells this pass wrote a result to (circular ones included). */
  readonly evaluatedCells: readonly CellRef[];
}

/** The parse cache is dropped when it grows past this many distinct formula texts. */
const MAX_PARSE_CACHE = 50_000;

const result = (
  evaluated: number,
  circularCells: readonly CellRef[],
  evaluatedCells: readonly CellRef[] = [],
): RecalculationResult => ({
  evaluated,
  circularCells,
  hasCircularReference: circularCells.length > 0,
  evaluatedCells,
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
  /** The precedent ranges too large to expand into the graph, indexed by sheet, column and row. */
  readonly #largeRanges = new LargeRangeIndex();
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
    this.#largeRanges.clear();
    this.#parseCache.clear();
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
    // Volatile cells recompute on every edit, so whatever reads them does too.
    for (const [k, cell] of this.#volatile) seeds.set(k, cell);
    return this.#evaluate(this.#collectAffected(seeds.entries()));
  }

  /**
   * Every formula cell that reads any of `cells`, via a stored edge or a summarized large precedent
   * range — the reverse index a move/swap uses to rewrite only the formulas that reference them.
   */
  formulasReferencing(cells: Iterable<CellRef>): CellRef[] {
    const found = new Map<string, CellRef>();
    for (const cell of cells) {
      for (const [k, dependent] of this.#graph.dependentEntries(CellRef.key(cell))) found.set(k, dependent);
      this.#largeRanges.forEachOver(cell, (k) => found.set(k, this.#formulas.get(k)!.ref));
    }
    return [...found.values()];
  }

  /** Recompute volatile cells and their dependents. */
  recalculateVolatile(): RecalculationResult {
    return this.#evaluate(this.#collectAffected(this.#volatile.entries()));
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
    this.#setLargeRanges(key, []);
  }

  #registerFormula(cell: CellRef, formula: string): void {
    const key = CellRef.key(cell);
    const ast = this.#parseCached(formula);
    this.#formulas.set(key, { ref: cell, ast });
    if (!ast) {
      this.#graph.setPrecedents(cell, []);
      this.#volatile.delete(key);
      this.#setLargeRanges(key, []);
      return;
    }
    const precedents = new Map<string, CellRef>();
    const largeRanges: RangeRef[] = [];
    const resolve = (n: string): CellRange | null => this.#resolveName(n, cell.sheet);
    ReferenceCollector.collect(ast, cell.sheet, resolve, precedents, largeRanges);
    this.#graph.setPrecedents(cell, precedents.values());
    this.#setLargeRanges(key, largeRanges);
    if (ReferenceCollector.isVolatile(ast)) this.#volatile.set(key, cell);
    else this.#volatile.delete(key);
  }

  #parseCached(formula: string): Expr | null {
    if (this.#parseCache.has(formula)) return this.#parseCache.get(formula)!;
    const parsed = Parser.tryParse(formula);
    const ast = parsed.ok ? parsed.expr : null;
    if (this.#parseCache.size >= MAX_PARSE_CACHE) this.#parseCache.clear();
    this.#parseCache.set(formula, ast);
    return ast;
  }

  /** A name as the formula on `sheet` sees it: a sheet-local name shadows a global one. */
  #resolveName(name: string, sheet: string): CellRange | null {
    return this.#workbook.findNamedRange(name, sheet)?.target ?? null;
  }

  /** Replace a formula's large precedent ranges, keeping the index in step. */
  #setLargeRanges(key: string, ranges: RangeRef[]): void {
    this.#largeRanges.set(key, ranges);
  }

  // ---- evaluation -------------------------------------------------------

  #collectAffected(seeds: Iterable<[string, CellRef]>): Set<string> {
    const affected = new Set<string>();
    const queue: [string, CellRef][] = [];
    // One search session for the whole walk, so each large range is reported at most once.
    const session = LargeRangeIndex.session();
    for (const [key, seed] of seeds) {
      if (this.#formulas.has(key) && !affected.has(key)) {
        affected.add(key);
        queue.push([key, seed]);
      }
      this.#addDependents(key, seed, affected, queue, session);
    }
    for (let i = 0; i < queue.length; i++) {
      const [key, cell] = queue[i]!;
      this.#addDependents(key, cell, affected, queue, session);
    }
    return affected;
  }

  /** Enqueue every formula reading `cell`: via a stored edge, or a large range containing it. */
  #addDependents(
    key: string,
    cell: CellRef,
    affected: Set<string>,
    queue: [string, CellRef][],
    session: RangeSearchSession,
  ): void {
    for (const [dk, dependent] of this.#graph.dependentEntries(key)) {
      if (!affected.has(dk)) {
        affected.add(dk);
        queue.push([dk, dependent]);
      }
    }
    this.#largeRanges.forEachOver(
      cell,
      (fk) => {
        if (affected.has(fk)) return;
        affected.add(fk);
        queue.push([fk, this.#formulas.get(fk)!.ref]);
      },
      session,
    );
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
    const evaluatedCells: CellRef[] = [...circularCells];
    for (const key of ordered) {
      const ref = this.#refOf(key);
      if (ref) evaluatedCells.push(ref);
    }
    return result(ordered.length, circularCells, evaluatedCells);
  }

  #refOf(key: string): CellRef | null {
    return this.#formulas.get(key)?.ref ?? null;
  }

  #evaluateCell(key: string): void {
    const entry = this.#formulas.get(key);
    if (!entry) return;
    if (!entry.ast) {
      // Keep a value the file cached (it came from an app that could read the formula); only a
      // formula with nothing cached shows the parse failure.
      const model = this.#workbook.findSheet(entry.ref.sheet)?.find(entry.ref.row, entry.ref.column);
      if (model && model.value.kind !== 'empty') return;
      this.#writeResult(entry.ref, CellValue.error(CellError.value('The formula could not be parsed.')));
      return;
    }
    const context = new WorkbookEvaluationContext(
      this.#workbook,
      entry.ref.sheet,
      this.#workbook.settings.effectiveLocale,
      this.#clock,
      { row: entry.ref.row, column: entry.ref.column },
    );
    this.#writeResult(entry.ref, new Evaluator(context, this.#functions).evaluate(entry.ast));
  }

  #writeResult(cell: CellRef, value: CellValue): void {
    const model = this.#workbook.findSheet(cell.sheet)?.find(cell.row, cell.column);
    if (model?.isFormula) model.value = value;
  }

  /** The nodes grouped by sheet, then column, each column's list sorted by row. */
  #nodesBySheetAndColumn(nodes: Set<string>): Map<string, Map<number, ColumnNodes>> {
    const out = new Map<string, Map<number, ColumnNodes>>();
    for (const node of nodes) {
      const ref = this.#refOf(node);
      if (!ref) continue;
      let columns = out.get(ref.sheet.toLowerCase());
      if (!columns) out.set(ref.sheet.toLowerCase(), (columns = new Map()));
      let list = columns.get(ref.column);
      if (!list) columns.set(ref.column, (list = { keys: [], rows: [], segments: new Map() }));
      list.keys.push(node);
      list.rows.push(ref.row);
    }
    for (const columns of out.values()) {
      for (const list of columns.values()) {
        const order = list.rows.map((_, i) => i).sort((a, b) => list.rows[a]! - list.rows[b]!);
        list.keys = order.map((i) => list.keys[i]!);
        list.rows = order.map((i) => list.rows[i]!);
      }
    }
    return out;
  }

  /**
   * Kahn's algorithm over the node set; nodes left in a cycle come back as `circular`.
   *
   * A range too large for the graph keeps only its corners there, so the formulas inside it are
   * ordered here instead. Rather than one edge from every formula inside to the range's reader
   * (quadratic for running totals), each column's formulas, sorted by row, get a segment tree of
   * virtual nodes: a segment is done when both halves are, and a range waits on the O(log n)
   * segments that tile the rows it covers.
   */
  #topologicalSort(nodes: Set<string>): { ordered: string[]; circular: string[] } {
    const inDegree = new Map<string, number>();
    for (const node of nodes) {
      let degree = 0;
      for (const p of this.#graph.precedentKeys(node)) if (nodes.has(p)) degree++;
      inDegree.set(node, degree);
    }
    /** Extra edges: from a node (real or a segment) to the nodes waiting on it. */
    const rangeDependents = new Map<string, string[]>();
    const addEdge = (from: string, to: string): void => {
      let list = rangeDependents.get(from);
      if (!list) rangeDependents.set(from, (list = []));
      list.push(to);
      inDegree.set(to, inDegree.get(to)! + 1);
    };
    const virtual = new Set<string>();
    let nextVirtual = 0;
    /** The node standing for sorted entries [lo, hi) of a column: the formula itself, or a segment. */
    const segment = (list: ColumnNodes, treeIndex: number, lo: number, hi: number): string => {
      if (hi - lo === 1) return list.keys[lo]!;
      const existing = list.segments.get(treeIndex);
      if (existing !== undefined) return existing;
      const key = `\u0001segment${nextVirtual++}`;
      virtual.add(key);
      inDegree.set(key, 0);
      list.segments.set(treeIndex, key);
      const mid = (lo + hi) >>> 1;
      addEdge(segment(list, 2 * treeIndex, lo, mid), key);
      addEdge(segment(list, 2 * treeIndex + 1, mid, hi), key);
      return key;
    };
    /** Make `reader` wait on the sorted entries [from, to) of a column. */
    const waitOn = (list: ColumnNodes, reader: string, from: number, to: number): void => {
      const descend = (treeIndex: number, lo: number, hi: number): void => {
        if (to <= lo || hi <= from) return;
        if (from <= lo && hi <= to) {
          addEdge(segment(list, treeIndex, lo, hi), reader);
          return;
        }
        const mid = (lo + hi) >>> 1;
        descend(2 * treeIndex, lo, mid);
        descend(2 * treeIndex + 1, mid, hi);
      };
      descend(1, 0, list.keys.length);
    };

    let bySheet: Map<string, Map<number, ColumnNodes>> | null = null;
    for (const [lk, ranges] of this.#largeRanges.entries()) {
      if (!nodes.has(lk)) continue;
      bySheet ??= this.#nodesBySheetAndColumn(nodes);
      const columns = bySheet;
      for (const r of ranges) {
        const sheetColumns = columns.get(r.sheet.toLowerCase());
        if (!sheetColumns) continue;
        const visit = (column: number): void => {
          const list = sheetColumns.get(column);
          if (!list) return;
          const from = lowerBound(list.rows, r.top);
          const to = lowerBound(list.rows, r.bottom + 1);
          if (from < to) waitOn(list, lk, from, to);
        };
        if (r.right - r.left + 1 > sheetColumns.size) {
          for (const column of sheetColumns.keys()) if (column >= r.left && column <= r.right) visit(column);
        } else {
          for (let column = r.left; column <= r.right; column++) visit(column);
        }
      }
    }
    const ready: string[] = [];
    for (const [node, degree] of inDegree) if (degree === 0) ready.push(node);
    const ordered: string[] = [];
    const release = (dk: string): void => {
      const degree = inDegree.get(dk);
      if (degree === undefined) return;
      inDegree.set(dk, degree - 1);
      if (degree - 1 === 0) ready.push(dk);
    };
    for (let i = 0; i < ready.length; i++) {
      const node = ready[i]!;
      if (!virtual.has(node)) {
        ordered.push(node);
        for (const [dk] of this.#graph.dependentEntries(node)) release(dk);
      }
      for (const dk of rangeDependents.get(node) ?? []) release(dk);
    }
    const circular: string[] = [];
    if (ordered.length !== nodes.size) {
      const placed = new Set(ordered);
      for (const node of nodes) if (!placed.has(node)) circular.push(node);
    }
    return { ordered, circular };
  }
}

/** One column's nodes sorted by row, and the segment nodes built over them (by tree index). */
interface ColumnNodes {
  keys: string[];
  rows: number[];
  readonly segments: Map<number, string>;
}

/** The first index whose value is >= `target` in an ascending list. */
function lowerBound(values: readonly number[], target: number): number {
  let lo = 0;
  let hi = values.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (values[mid]! < target) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}
