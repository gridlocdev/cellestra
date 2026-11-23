import {
  CellAddress,
  CellRange,
  CellValue,
  type Expr,
  FormulaWriter,
  InputParser,
  namedRangeDefinition,
  NumberFormats,
  Parser,
  ReferenceAdjuster,
  SheetDimensions,
  Validator,
  type Workbook,
  type Worksheet,
} from '@cellestra/core';
import { display, pickSheet, resolveCell } from '../infra/cells.ts';
import { CliError } from '../infra/errors.ts';
import { resolveLocale } from '../infra/locale.ts';
import { address, writeReport } from '../infra/report.ts';
import { load, recalculate } from '../infra/workspace.ts';
import type { CommandSpec } from '../spec.ts';
import { fileArg, jsonOpt, localeOpt, sheetOpt } from './options.ts';

/**
 * Looking at a workbook the way someone debugging it does: `trace` follows a cell's formula to what
 * it reads (or who reads it), and `lint` finds the mistakes that are not errors.
 */

// ---- trace ------------------------------------------------------------------

/** A cell, or a whole range kept as one node so `SUM(D2:D4)` reads as one line. */
class Node {
  readonly sheet: string;
  readonly range: CellRange;

  constructor(sheet: string, range: CellRange) {
    this.sheet = sheet;
    this.range = range;
  }

  get isCell(): boolean {
    return this.range.rowCount === 1 && this.range.columnCount === 1;
  }

  get address(): string {
    return this.isCell
      ? address(this.sheet, this.range.top, this.range.left)
      : `${this.sheet}.${this.range.withSheet(null).toString()}`;
  }
}

interface TraceNode {
  address: string;
  formula: string | null;
  display: string | null;
  error: string | null;
  cells: number | null;
  truncated: boolean | null;
  cycle: boolean | null;
  precedents: TraceNode[] | null;
  dependents: TraceNode[] | null;
}

const NODE_LIMIT = 2000;

/** A ceiling on how much of a fan-out tree one trace prints, so a hub cell can't run away. */
class NodeBudget {
  #left = NODE_LIMIT;
  exhausted = false;

  take(count: number): boolean {
    if (count > this.#left) {
      this.exhausted = true;
      return false;
    }
    this.#left -= count;
    return true;
  }
}

const cellNode = (sheet: string, row: number, column: number): Node =>
  new Node(sheet, new CellRange(new CellAddress(column, row)));

interface IndexedFormula {
  readonly sheet: Worksheet;
  readonly row: number;
  readonly column: number;
  readonly expr: Expr;
  readonly reads: Node[];
}

const posKey = (row: number, column: number): string => `${row},${column}`;

/** Every formula in the workbook, parsed once, with the cells and ranges it reads. */
class FormulaIndex {
  readonly #formulas = new Map<string, IndexedFormula>();
  readonly #readers = new Map<string, { sheet: string; row: number; column: number; readers: Node[] }>();
  readonly #rangeReaders: [read: Node, formula: Node][] = [];
  readonly #bySheet = new Map<string, Map<string, { row: number; column: number; expr: Expr }>>();

  static build(workbook: Workbook): FormulaIndex {
    const index = new FormulaIndex();
    for (const sheet of workbook.worksheets) {
      for (const [row, column, cell] of sheet.enumerateCells()) {
        if (cell.formula === null) continue;
        const parsed = Parser.tryParse(cell.formula);
        if (!parsed.ok) continue;
        const expr = parsed.expr;
        const reads: Node[] = [];
        collectReads(expr, sheet.name, workbook, reads);
        index.#formulas.set(address(sheet.name, row, column).toUpperCase(), {
          sheet,
          row,
          column,
          expr,
          reads,
        });
        const sheetKey = sheet.name.toUpperCase();
        let onSheet = index.#bySheet.get(sheetKey);
        if (!onSheet) index.#bySheet.set(sheetKey, (onSheet = new Map()));
        onSheet.set(posKey(row, column), { row, column, expr });
        const self = cellNode(sheet.name, row, column);
        for (const read of reads) {
          if (read.isCell) {
            const s = read.sheet.toUpperCase();
            const key = `${s}\u0000${read.range.top}\u0000${read.range.left}`;
            let entry = index.#readers.get(key);
            if (!entry) {
              entry = { sheet: s, row: read.range.top, column: read.range.left, readers: [] };
              index.#readers.set(key, entry);
            }
            entry.readers.push(self);
          } else {
            index.#rangeReaders.push([read, self]);
          }
        }
      }
    }
    return index;
  }

  onSheet(sheet: Worksheet): Map<string, { row: number; column: number; expr: Expr }> {
    return this.#bySheet.get(sheet.name.toUpperCase()) ?? new Map();
  }

  precedentsOf(node: Node): readonly Node[] {
    return node.isCell ? (this.#formulas.get(node.address.toUpperCase())?.reads ?? []) : [];
  }

  /** Formula cells that read any cell of `node` — a lookup, not a sweep. */
  dependentsOf(node: Node): Node[] {
    const result: Node[] = [];
    const seen = new Set<string>();
    const add = (reader: Node): void => {
      const key = reader.address.toUpperCase();
      if (seen.has(key)) return;
      seen.add(key);
      result.push(reader);
    };
    const nodeSheet = node.sheet.toUpperCase();
    if (node.isCell) {
      const direct = this.#readers.get(`${nodeSheet}\u0000${node.range.top}\u0000${node.range.left}`);
      for (const reader of direct?.readers ?? []) add(reader);
    } else {
      for (const entry of this.#readers.values()) {
        if (entry.sheet === nodeSheet && node.range.contains(entry.row, entry.column)) {
          for (const reader of entry.readers) add(reader);
        }
      }
    }
    for (const [read, formula] of this.#rangeReaders) {
      if (read.sheet.toUpperCase() === nodeSheet && read.range.intersects(node.range)) add(formula);
    }
    return result;
  }

  get all(): Iterable<IndexedFormula> {
    return this.#formulas.values();
  }
}

function collectReads(expr: Expr, home: string, workbook: Workbook, into: Node[]): void {
  switch (expr.kind) {
    case 'reference':
      into.push(new Node(expr.address.sheet ?? home, new CellRange(expr.address.withoutSheet().relative())));
      break;
    case 'range':
      into.push(new Node(expr.range.sheet ?? home, expr.range.withSheet(null)));
      break;
    case 'name': {
      const target = workbook.findNamedRange(expr.name, home)?.target;
      if (target) into.push(new Node(target.sheet ?? home, target.withSheet(null)));
      break;
    }
    case 'unary':
      collectReads(expr.operand, home, workbook, into);
      break;
    case 'binary':
      collectReads(expr.left, home, workbook, into);
      collectReads(expr.right, home, workbook, into);
      break;
    case 'call':
      for (const a of expr.args) collectReads(a, home, workbook, into);
      break;
    case 'array':
    case 'literal':
    case 'refError':
    case 'sheetSpan':
    default:
      break;
  }
}

function walk(
  workbook: Workbook,
  index: FormulaIndex,
  node: Node,
  downstream: boolean,
  maxDepth: number,
  depth: number,
  visited: Set<string>,
  locale: string,
  budget: NodeBudget,
): TraceNode {
  const sheet = workbook.findSheet(node.sheet);
  const cell = node.isCell ? sheet?.find(node.range.top, node.range.left) : undefined;
  const key = node.address.toUpperCase();
  const cycle = visited.has(key);
  visited.add(key);
  const shown =
    !cell || cell.value.kind === 'empty'
      ? null
      : display(workbook, sheet!, node.range.top, node.range.left, locale);
  let children: TraceNode[] | null = null;
  let truncated = false;
  if (!cycle) {
    const next = downstream ? index.dependentsOf(node) : index.precedentsOf(node);
    if (next.length > 0) {
      if (depth < maxDepth && budget.take(next.length)) {
        children = next.map((n) =>
          walk(workbook, index, n, downstream, maxDepth, depth + 1, visited, locale, budget),
        );
      } else {
        truncated = true;
      }
    }
  }
  // A cell may appear again on another branch; only a true loop is a cycle.
  visited.delete(key);
  return {
    address: node.address,
    formula: cell && cell.formula !== null ? '=' + cell.formula : null,
    display: shown,
    error: cell && cell.value.kind === 'error' ? cell.value.error.token : null,
    cells: node.isCell ? null : node.range.rowCount * node.range.columnCount,
    truncated: truncated ? true : null,
    cycle: cycle ? true : null,
    precedents: downstream ? null : children,
    dependents: downstream ? children : null,
  };
}

function printTree(lines: string[], node: TraceNode, downstream: boolean, indent: number): void {
  let line = ' '.repeat(indent * 2) + node.address;
  if (node.cells !== null) line += ` (${node.cells} cells)`;
  if (node.formula !== null) line += '  ' + node.formula;
  if (node.display !== null) line += '  → ' + node.display;
  if (node.cycle === true) line += '  [cycle]';
  if (node.truncated === true) line += '  …';
  lines.push(line);
  for (const child of (downstream ? node.dependents : node.precedents) ?? []) {
    printTree(lines, child, downstream, indent + 1);
  }
}

export const trace: CommandSpec = {
  name: 'trace',
  description: 'Follow a formula to its precedents (what it reads) or dependents (who reads it)',
  args: [
    fileArg(),
    { name: 'cell', description: 'The cell to trace from, e.g. Sheet1.D5 (or a single-cell name)' },
  ],
  options: [
    {
      name: '--dependents',
      description: 'Follow who reads this cell, instead of what it reads (precedents)',
      kind: 'flag',
    },
    {
      name: '--depth',
      description: 'How many levels to follow (0 = all the way)',
      type: 'int',
      default: '1',
    },
    localeOpt,
    jsonOpt,
  ],
  run(p, con) {
    const locale = resolveLocale(p.str('--locale'));
    const loaded = load(con, p.arg('file'), null, locale);
    recalculate(loaded.workbook);
    const effective = locale ?? loaded.workbook.settings.effectiveLocale;
    const workbook = loaded.workbook;
    const { sheet, address: a } = resolveCell(workbook, p.arg('cell'));
    const downstream = p.flag('--dependents');
    const maxDepth = p.num('--depth') ?? 1;
    if (maxDepth < 0) throw new CliError('--depth must be 0 (all) or a positive number.');
    const index = FormulaIndex.build(workbook);
    const budget = new NodeBudget();
    const tree = walk(
      workbook,
      index,
      cellNode(sheet.name, a.row, a.column),
      downstream,
      maxDepth === 0 ? Number.MAX_SAFE_INTEGER : maxDepth,
      0,
      new Set(),
      effective,
      budget,
    );
    if (p.flag('--json')) {
      writeReport(con, {
        cell: tree.address,
        direction: downstream ? 'dependents' : 'precedents',
        depth: maxDepth,
        truncated: budget.exhausted ? true : null,
        tree,
      });
      return 0;
    }
    const lines: string[] = [];
    printTree(lines, tree, downstream, 0);
    for (const line of lines) con.outLine(line);
    if (budget.exhausted) con.errLine(`note: stopped after ${NODE_LIMIT} nodes; narrow it with --depth.`);
    return 0;
  },
};

// ---- lint -------------------------------------------------------------------

interface Finding {
  readonly severity: string;
  readonly code: string;
  readonly cell: string;
  readonly message: string;
}

const finding = (severity: string, code: string, cell: string, message: string): Finding => ({
  severity,
  code,
  cell,
  message,
});

const rank = (severity: string): number => (severity === 'error' ? 0 : severity === 'warning' ? 1 : 2);

function checkStoredAsText(workbook: Workbook, sheet: Worksheet, locale: string, findings: Finding[]): void {
  for (const [row, column, cell] of sheet.enumerateCells()) {
    if (cell.value.kind !== 'text' || cell.formula !== null) continue;
    const text = cell.value.value.trim();
    // An explicit text format means it's text on purpose.
    if (text.length === 0 || NumberFormats.kind(workbook.styles.get(cell.styleId).numberFormat) === 'text')
      continue;
    const parsed = InputParser.parse(text, locale);
    if (parsed.formula !== null || parsed.value.kind !== 'number') continue;
    const kind = parsed.suggestedFormat ? NumberFormats.kind(parsed.suggestedFormat) : null;
    const date = kind === 'date' || kind === 'dateTime' || kind === 'time';
    findings.push(
      finding(
        'warning',
        date ? 'date-as-text' : 'number-as-text',
        address(sheet.name, row, column),
        `"${text}" is text; it won't add up or sort as a ${date ? 'date' : 'number'} — re-enter it, or apply a set op with a typed value`,
      ),
    );
  }
}

function checkTableHeaders(sheet: Worksheet, findings: Finding[]): void {
  for (const table of sheet.tables.filter((t) => t.hasHeader)) {
    for (let column = table.range.left; column <= table.range.right; column++) {
      const value = sheet.getValue(table.range.top, column);
      if (value.kind !== 'text' || value.value.trim().length === 0) {
        findings.push(
          finding(
            'warning',
            'table-header',
            address(sheet.name, table.range.top, column),
            `table '${table.name}' header is ${value.kind === 'empty' ? 'empty' : 'not text'}`,
          ),
        );
      }
    }
  }
}

function* resolveValues(sheet: Worksheet, range: CellRange): IterableIterator<CellValue> {
  for (const [row, column] of range.positions()) yield sheet.getValue(row, column);
}

function checkValidationRules(sheet: Worksheet, locale: string, findings: Finding[]): void {
  if (sheet.validations.length === 0) return;
  for (const [row, column, cell] of sheet.enumerateCells()) {
    if (cell.formula !== null || cell.value.kind === 'empty') continue;
    const rule = sheet.findValidation(row, column);
    if (!rule) continue;
    if (!Validator.isValid(rule, cell.value, locale, (range) => resolveValues(sheet, range))) {
      findings.push(
        finding(
          'warning',
          'validation',
          address(sheet.name, row, column),
          `"${CellValue.toInvariantString(cell.value)}" violates the ${rule.kind.toLowerCase()} rule on ${rule.target.toString()}`,
        ),
      );
    }
  }
}

function isEmptyNode(workbook: Workbook, node: Node): boolean {
  const sheet = workbook.findSheet(node.sheet);
  if (!sheet) return true;
  // A small range: probe it. A big one (a whole column, say): walk what's populated instead.
  if (node.range.rowCount * node.range.columnCount <= 4096) {
    for (const [row, column] of node.range.positions()) {
      if (sheet.getValue(row, column).kind !== 'empty') return false;
    }
    return true;
  }
  for (const [row, column, cell] of sheet.enumerateCells()) {
    if (node.range.contains(row, column) && cell.value.kind !== 'empty') return false;
  }
  return true;
}

const AGGREGATES = new Set(['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA']);

/** Formula-level checks: all-empty inputs, and an aggregate over a run that stops just short of more data. */
function checkFormulas(
  workbook: Workbook,
  index: FormulaIndex,
  sheets: readonly Worksheet[],
  findings: Finding[],
): void {
  const names = new Set(sheets.map((s) => s.name.toUpperCase()));
  for (const { sheet, row, column, expr, reads } of index.all) {
    if (!names.has(sheet.name.toUpperCase())) continue;
    const here = address(sheet.name, row, column);
    if (reads.length > 0 && reads.every((r) => isEmptyNode(workbook, r))) {
      findings.push(
        finding(
          'warning',
          'empty-inputs',
          here,
          `=${sheet.find(row, column)!.formula} reads only empty cells`,
        ),
      );
    }
    if (expr.kind !== 'call' || !AGGREGATES.has(expr.name.toUpperCase())) continue;
    for (const argument of expr.args) {
      if (argument.kind !== 'range') continue;
      const range = argument.range;
      const target = workbook.findSheet(range.sheet ?? sheet.name);
      if (!target) continue;
      // A single column (or row) run: is the next cell past its end populated and not the formula itself?
      if (range.columnCount === 1 && range.rowCount > 1) {
        const next = range.bottom + 1;
        if (
          !(target === sheet && next === row && range.left === column) &&
          target.getValue(next, range.left).kind === 'number'
        ) {
          findings.push(
            finding(
              'warning',
              'short-range',
              here,
              `${expr.name}(${range.withSheet(null).toString()}) stops at row ${range.bottom + 1} but ${address(target.name, next, range.left)} holds a number`,
            ),
          );
        }
      } else if (range.rowCount === 1 && range.columnCount > 1) {
        const next = range.right + 1;
        if (
          !(target === sheet && next === column && range.top === row) &&
          target.getValue(range.top, next).kind === 'number'
        ) {
          findings.push(
            finding(
              'warning',
              'short-range',
              here,
              `${expr.name}(${range.withSheet(null).toString()}) stops at column ${SheetDimensions.columnToLetters(range.right)} but ${address(target.name, range.top, next)} holds a number`,
            ),
          );
        }
      }
    }
  }
}

/**
 * A formula that breaks the pattern of its neighbours: the cells above and below (or left and
 * right) agree once their relative references are shifted onto this row, and this one differs.
 */
function checkFormulaPatterns(index: FormulaIndex, sheet: Worksheet, findings: Finding[]): void {
  const formulas = index.onSheet(sheet);
  for (const { row, column, expr } of formulas.values()) {
    const here = FormulaWriter.write(expr);
    if (
      breaks(formulas, posKey(row - 1, column), posKey(row + 1, column), 0, 1, here) ||
      breaks(formulas, posKey(row, column - 1), posKey(row, column + 1), 1, 0, here)
    ) {
      findings.push(
        finding(
          'warning',
          'inconsistent-formula',
          address(sheet.name, row, column),
          `=${sheet.find(row, column)!.formula} differs from the pattern of its neighbours`,
        ),
      );
    }
  }
}

function breaks(
  formulas: Map<string, { expr: Expr }>,
  before: string,
  after: string,
  columnDelta: number,
  rowDelta: number,
  here: string,
): boolean {
  const a = formulas.get(before);
  const b = formulas.get(after);
  if (!a || !b) return false;
  const fromBefore = FormulaWriter.write(
    ReferenceAdjuster.forCopy(a.expr, columnDelta, rowDelta),
  ).toUpperCase();
  const fromAfter = FormulaWriter.write(
    ReferenceAdjuster.forCopy(b.expr, -columnDelta, -rowDelta),
  ).toUpperCase();
  return fromBefore === fromAfter && fromBefore !== here.toUpperCase();
}

function collectNames(expr: Expr, into: Set<string>): void {
  switch (expr.kind) {
    case 'name':
      into.add(expr.name.toUpperCase());
      break;
    case 'unary':
      collectNames(expr.operand, into);
      break;
    case 'binary':
      collectNames(expr.left, into);
      collectNames(expr.right, into);
      break;
    case 'call':
      for (const a of expr.args) collectNames(a, into);
      break;
    case 'array':
    case 'literal':
    case 'range':
    case 'reference':
    case 'refError':
    case 'sheetSpan':
    default:
      break;
  }
}

function checkUnusedNames(workbook: Workbook, index: FormulaIndex, findings: Finding[]): void {
  if (workbook.namedRanges.length === 0) return;
  const used = new Set<string>();
  for (const { expr } of index.all) collectNames(expr, used);
  for (const named of workbook.namedRanges) {
    if (!used.has(named.name.toUpperCase())) {
      findings.push(
        finding(
          'info',
          'unused-name',
          named.name,
          `named range ${named.name} = ${namedRangeDefinition(named)} is not used by any formula`,
        ),
      );
    }
  }
}

export const lint: CommandSpec = {
  name: 'lint',
  description:
    'Find likely mistakes that are not errors: numbers as text, short ranges, inconsistent formulas, bad headers, rule violations',
  args: [fileArg()],
  options: [
    sheetOpt,
    { name: '--strict', description: 'Exit 1 on warnings too, not only on errors', kind: 'flag' },
    localeOpt,
    jsonOpt,
  ],
  run(p, con) {
    const locale = resolveLocale(p.str('--locale'));
    const loaded = load(con, p.arg('file'), null, locale);
    const result = recalculate(loaded.workbook);
    const effective = locale ?? loaded.workbook.settings.effectiveLocale;
    const workbook = loaded.workbook;
    const only = p.str('--sheet');
    if (only !== null) pickSheet(workbook, only); // validates the name
    const same = (name: string): boolean => only === null || name.toUpperCase() === only.toUpperCase();
    const findings: Finding[] = [];
    const sheets = workbook.worksheets.filter((s) => !s.isChartSheet && same(s.name));
    const index = FormulaIndex.build(workbook);

    // Errors — what validate reports, so lint is a superset.
    for (const s of sheets) {
      for (const [row, column, cell] of s.enumerateCells()) {
        if (cell.value.kind === 'error') {
          findings.push(
            finding(
              'error',
              'error-value',
              address(s.name, row, column),
              `${cell.value.error.token}${cell.formula !== null ? ' from =' + cell.formula : ''}`,
            ),
          );
        }
      }
    }
    for (const c of result.circularCells.filter((c) => same(c.sheet))) {
      findings.push(
        finding('error', 'circular', address(c.sheet, c.row, c.column), 'part of a circular reference'),
      );
    }
    for (const s of sheets) {
      checkStoredAsText(workbook, s, effective, findings);
      checkTableHeaders(s, findings);
      checkValidationRules(s, effective, findings);
      checkFormulaPatterns(index, s, findings);
    }
    checkFormulas(workbook, index, sheets, findings);
    checkUnusedNames(workbook, index, findings);

    findings.sort((a, b) => rank(a.severity) - rank(b.severity));
    const errors = findings.filter((f) => f.severity === 'error').length;
    const warnings = findings.filter((f) => f.severity === 'warning').length;
    const failed = errors > 0 || (p.flag('--strict') && warnings > 0);
    if (p.flag('--json')) {
      writeReport(con, {
        ok: !failed,
        file: p.arg('file'),
        errors,
        warnings,
        infos: findings.length - errors - warnings,
        findings,
      });
      return failed ? 1 : 0;
    }
    for (const f of findings) con.outLine(`${f.severity}\t${f.cell}\t${f.code}\t${f.message}`);
    con.outLine(
      findings.length === 0
        ? 'OK — nothing to report.'
        : `${errors} error(s), ${warnings} warning(s), ${findings.length - errors - warnings} note(s).`,
    );
    return failed ? 1 : 0;
  },
};
