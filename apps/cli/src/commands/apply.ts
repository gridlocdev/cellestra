import {
  Argb,
  BorderEdges,
  Borders,
  CellAddress,
  CellRange,
  CellValue,
  DataEditCommands,
  DeleteSheetCommand,
  EditCommands,
  Fills,
  InputParser,
  namedRangeError,
  tableNameError,
  planCellEntry,
  RenameSheetCommand,
  StructuralEditCommand,
  TableCommands,
  dataValidation,
  type ArgbColor,
  type BorderLineStyle,
  type CellStyle,
  type ColumnFilter,
  type DataValidation,
  type EditCommand,
  type Fill,
  type HorizontalAlignment,
  type StructuralAxis,
  type StructuralOperation,
  type Table,
  type ValidationKind,
  type ValidationOperator,
  type ValidationResponse,
  type VerticalAlignment,
  type Worksheet,
} from '@cellestra/core';
import {
  pickSheet,
  resolveCell,
  resolveRange,
  type ResolvedCell,
  type ResolvedRange,
} from '../infra/cells.ts';
import { EditSession } from '../infra/editSession.ts';
import { CliError, isExpectedError } from '../infra/errors.ts';
import type { Console } from '../infra/io.ts';
import { resolveLocale } from '../infra/locale.ts';
import { address } from '../infra/report.ts';
import { readInputText } from '../infra/workspace.ts';
import { APPLY_REFERENCE } from '../referenceTexts.ts';
import type { CommandSpec, Parsed } from '../spec.ts';
import { addChart, chartSheet, updateChart } from './chartOps.ts';
import { editTarget, fileArg, inPlaceOpt, jsonOpt, localeOpt, outputOpt } from './options.ts';
import {
  checkStructuralEdit,
  parseColumn,
  parseColumnSpan,
  parseRow,
  parseRowSpan,
  parseSortLevels,
} from './parsers.ts';
import { conditionalOp, pageSetupOp } from './printing.ts';
import { copyOrMove, fill, filter, paint } from './rangeOps.ts';

/**
 * `apply`: one JSON document of typed operations, one load, one recalculation, one atomic save, one
 * report. Every op drives the same core builder the matching verb (and the app) does; the verbs are
 * thin fronts that build one op and dispatch it here, so a verb and its op can never disagree.
 */

export { APPLY_REFERENCE };

type Json = null | boolean | number | string | Json[] | { [key: string]: Json };
type JsonObject = { [key: string]: Json };

const isObject = (v: unknown): v is JsonObject => typeof v === 'object' && v !== null && !Array.isArray(v);

/** A key, matched case-insensitively (the first match wins). */
function property(element: JsonObject, name: string): { found: boolean; value: Json | undefined } {
  const lower = name.toLowerCase();
  for (const key of Object.keys(element)) {
    if (key.toLowerCase() === lower) return { found: true, value: element[key] };
  }
  return { found: false, value: undefined };
}

/** One op's fields, read leniently (case-insensitive names, numbers as strings) with clear errors. */
export class Op {
  readonly #element: JsonObject;
  readonly index: number;

  constructor(element: JsonObject, index: number) {
    this.#element = element;
    this.index = index;
  }

  /** Present with a non-null value. */
  has(name: string): boolean {
    const p = property(this.#element, name);
    return p.found && p.value !== null && p.value !== undefined;
  }

  raw(name: string): Json | undefined {
    return property(this.#element, name).value;
  }

  /** True when the key is present, even with a null value (so `"fill": null` can mean "remove"). */
  present(name: string): boolean {
    return property(this.#element, name).found;
  }

  string(name: string): string | null {
    const v = this.raw(name);
    if (v === null || v === undefined) return null;
    if (typeof v === 'string') return v;
    if (typeof v === 'number') return String(v);
    if (typeof v === 'boolean') return v ? 'true' : 'false';
    throw new CliError(`"${name}" must be a string.`);
  }

  require(name: string): string {
    const v = this.string(name);
    if (v === null) throw new CliError(`needs "${name}".`);
    return v;
  }

  bool(name: string): boolean | null {
    const v = this.raw(name);
    if (v === null || v === undefined) return null;
    if (typeof v === 'boolean') return v;
    if (typeof v === 'string') {
      const t = v.trim().toLowerCase();
      if (t === 'true') return true;
      if (t === 'false') return false;
    }
    throw new CliError(`"${name}" must be true or false.`);
  }

  number(name: string): number | null {
    const v = this.raw(name);
    if (v === null || v === undefined) return null;
    if (typeof v === 'number') return v;
    if (typeof v === 'string' && v.trim() !== '') {
      const n = Number(v.trim());
      if (!Number.isNaN(n)) return n;
    }
    throw new CliError(`"${name}" must be a number.`);
  }

  int(name: string, fallback: number): number {
    const n = this.number(name);
    if (n === null) return fallback;
    if (!Number.isInteger(n) || n < -2147483648 || n > 2147483647) {
      throw new CliError(`"${name}" must be a whole number.`);
    }
    return n;
  }

  strings(name: string): string[] {
    const v = this.raw(name);
    if (v === null || v === undefined) return [];
    if (typeof v === 'string') {
      // A lone string is the verb's comma-separated spelling ("Red,Green,Blue").
      return v
        .split(',')
        .map((s) => s.trim())
        .filter((s) => s.length > 0);
    }
    if (!Array.isArray(v)) throw new CliError(`"${name}" must be an array of strings.`);
    return v.map((e) => (typeof e === 'string' ? e : JSON.stringify(e)));
  }
}

// ---- The apply verb ----------------------------------------------------------

export const apply: CommandSpec = {
  name: 'apply',
  description: 'Apply a batch of typed operations from JSON: cells, styles, layout, structure, data, charts',
  args: [
    fileArg(),
    fileArg('ops', 'JSON file of operations (or - for stdin) — see `cellestra schema apply`'),
  ],
  options: [
    inPlaceOpt,
    outputOpt,
    { name: '--dry-run', description: 'Apply and report, but write nothing', kind: 'flag' },
    localeOpt,
    jsonOpt,
  ],
  run(p, con) {
    const path = p.arg('file');
    const preview = p.flag('--dry-run');
    const target = preview ? path : editTarget(path, p.flag('--in-place'), p.str('--output'));
    const ops = parseDocument(con, p.arg('ops'));
    return EditSession.run(
      con,
      path,
      target,
      resolveLocale(p.str('--locale')),
      p.flag('--json'),
      (session) => {
        ops.forEach(({ element, kind }, i) => {
          try {
            run(session, new Op(element, i), kind);
          } catch (e) {
            // Core's own refusals (duplicate sheet, bad name, …) are user errors too, and the user
            // needs to know which op raised them.
            if (isExpectedError(e)) throw new CliError(`op ${i + 1} (${kind}): ${e.message}`);
            throw e;
          }
        });
        session.field('applied', ops.length);
        session.note(`Applied ${ops.length} op(s).`);
      },
      'ods',
      preview,
    );
  },
};

function parseDocument(con: Console, path: string): { element: JsonObject; kind: string }[] {
  const text = readInputText(con, path);
  let root: unknown;
  try {
    root = JSON.parse(text);
  } catch (e) {
    throw new CliError(`Invalid ops JSON: ${(e as Error).message}`);
  }
  if (isObject(root)) {
    const list = property(root, 'ops');
    if (list.found) root = list.value;
  }
  if (!Array.isArray(root)) {
    throw new CliError('The ops document must be an array of operations, or {"ops": [...]}.');
  }
  const ops: { element: JsonObject; kind: string }[] = [];
  for (const element of root as unknown[]) {
    const n = ops.length + 1;
    if (!isObject(element)) throw new CliError(`Op ${n} is not an object.`);
    const op = property(element, 'op').value;
    let kind: string;
    if (typeof op === 'string') kind = op.trim().toLowerCase();
    else if (property(element, 'cell').found)
      kind = 'set'; // the original {cell, value|formula} form
    else throw new CliError(`Op ${n} has no "op".`);
    ops.push({ element, kind });
  }
  return ops;
}

/**
 * Run one op given as plain fields — how the standalone verbs (`style`, `width`, `freeze`, …) share
 * the ops' single implementation.
 */
export function runOp(session: EditSession, kind: string, fields: Record<string, unknown>): void {
  run(session, new Op(JSON.parse(JSON.stringify(fields)) as JsonObject, 0), kind);
}

/**
 * The verbs' front: resolve the destination, drop absent fields (an absent option must be an
 * absent key, not a null — `"fill": null` means "remove the fill"), and run the one op.
 */
export function dispatch(
  p: Parsed,
  con: Console,
  kind: string,
  fields: Record<string, unknown>,
  locale: string | null = null,
): number {
  const path = p.arg('file');
  const target = editTarget(path, p.flag('--in-place'), p.str('--output'));
  // `undefined` is always dropped; null too, unless the verb asks for "remove" with it.
  const present = Object.fromEntries(Object.entries(fields).filter(([, v]) => v !== undefined && v !== null));
  for (const [k, v] of Object.entries(fields)) if (v === REMOVE) present[k] = null;
  return EditSession.run(con, path, target, locale, p.flag('--json'), (session) =>
    runOp(session, kind, present),
  );
}

/** A field value a verb passes to `dispatch` to send an explicit null ("remove it"). */
export const REMOVE = Symbol('remove');

// ---- Dispatch ----------------------------------------------------------------

function run(session: EditSession, op: Op, kind: string): void {
  const wb = session.workbook;
  switch (kind) {
    case 'set':
      return set(session, op);
    case 'clear':
      return clear(session, op);
    case 'style':
      return style(session, op);
    case 'conditional':
      return conditionalOp(session, op);
    case 'page-setup':
    case 'pagesetup':
      return pageSetupOp(session, op);
    case 'width':
      return width(session, op);
    case 'height':
      return height(session, op);
    case 'freeze':
      return freeze(session, op);
    case 'merge': {
      const { sheet, range } = rangeOf(session, op, 'range');
      session.apply(EditCommands.mergeCells(sheet.name, range));
      return;
    }
    case 'unmerge': {
      const { sheet, address: at } = cellOf(session, op, 'cell');
      session.apply(EditCommands.unmergeCells(sheet.name, at.row, at.column));
      return;
    }
    case 'hide':
      return visibility(session, op, true);
    case 'show':
      return visibility(session, op, false);
    case 'add-sheet':
      return addSheet(session, op);
    case 'rename-sheet': {
      const sheet = pickSheet(wb, op.require('name'));
      const to = op.require('to');
      session.apply(new RenameSheetCommand(sheet.name, to));
      session.field('sheet', to);
      return;
    }
    case 'duplicate-sheet': {
      const copy = wb.duplicateSheet(pickSheet(wb, op.require('name')));
      const newName = op.string('as');
      if (newName !== null) wb.renameSheet(copy, newName);
      session.field('sheet', copy.name);
      return;
    }
    case 'move-sheet': {
      const sheet = pickSheet(wb, op.require('name'));
      const to = op.int('to', 0);
      if (to < 1 || to > wb.sheetCount) {
        throw new CliError(`"to" must be a 1-based position from 1 to ${wb.sheetCount}.`);
      }
      wb.moveSheet(wb.indexOf(sheet), to - 1);
      return;
    }
    case 'settings':
      return settings(session, op);
    case 'copy':
    case 'move': {
      const from = rangeOf(session, op, 'from');
      const to = rangeOf(session, op, 'to');
      copyOrMove(
        session,
        from.sheet,
        from.range,
        to.sheet,
        to.range,
        kind === 'move',
        op.string('mode') ?? 'all',
      );
      return;
    }
    case 'fill': {
      const { sheet, range } = rangeOf(session, op, 'range');
      let source: CellRange | null = null;
      const fromText = op.string('from');
      if (fromText !== null) {
        const from = resolveRange(wb, fromText, sheet.name);
        if (from.sheet !== sheet)
          throw new CliError("a fill's source must be on the same sheet as the range.");
        source = from.range;
      }
      fill(session, sheet, range, source);
      return;
    }
    case 'paint': {
      const from = rangeOf(session, op, 'from');
      const to = rangeOf(session, op, 'to');
      paint(session, from.sheet, from.range, to.sheet, to.range);
      return;
    }
    case 'filter': {
      const { sheet, range } = rangeOf(session, op, 'range');
      const columns = new Map<number, ColumnFilter>();
      const spec = op.raw('columns');
      if (isObject(spec)) {
        for (const [key, value] of Object.entries(spec)) {
          const column = parseColumn(key);
          const criteria = new Op(isObject(value) ? value : {}, op.index);
          const allowed = criteria.strings('values');
          columns.set(column, {
            allowedValues: allowed.length > 0 ? new Set(allowed) : null,
            condition: criteria.string('condition'),
          });
        }
      }
      filter(session, sheet, range, op.bool('header') ?? true, op.bool('clear') ?? false, columns);
      return;
    }
    case 'delete-sheet':
      session.apply(new DeleteSheetCommand(pickSheet(wb, op.require('name')).name));
      return;
    case 'insert-rows':
      return structural(session, op, 'row', 'insert');
    case 'delete-rows':
      return structural(session, op, 'row', 'delete');
    case 'insert-cols':
      return structural(session, op, 'column', 'insert');
    case 'delete-cols':
      return structural(session, op, 'column', 'delete');
    case 'sort':
      return sort(session, op);
    case 'remove-duplicates':
      return removeDuplicates(session, op);
    case 'split':
      return split(session, op);
    case 'table':
      return tableOp(session, op);
    case 'validation':
      return validation(session, op);
    case 'table-remove': {
      const { sheet, table } = findTable(session, op);
      session.apply(DataEditCommands.removeTable(sheet.name, table));
      session.field('table', table.name);
      return;
    }
    case 'table-options':
      return tableOptions(session, op);
    case 'rename-table': {
      const { sheet, table } = findTable(session, op);
      const to = op.require('to');
      const error = tableNameError(wb, to.trim(), table);
      if (error !== null) throw new CliError(error);
      session.apply(DataEditCommands.renameTable(wb, sheet.name, table, to));
      session.field('table', to.trim());
      return;
    }
    case 'chart-remove': {
      const sheet = chartSheet(wb, op.require('sheet'));
      session.apply(new DeleteSheetCommand(sheet.name));
      session.field('sheet', sheet.name);
      return;
    }
    case 'chart-update':
      updateChart(
        session,
        op.require('sheet'),
        op.string('type'),
        chartLook(op),
        op.bool('legend'),
        chartData(op),
      );
      return;
    case 'name':
      return name(session, op);
    case 'unname':
      if (!wb.removeNamedRange(op.require('name'))) {
        throw new CliError(`no named range called '${op.string('name')}'.`);
      }
      return;
    case 'hyperlink':
      return hyperlink(session, op);
    case 'chart':
      addChart(
        session,
        op.string('type') ?? 'column',
        chartLook(op),
        !(op.bool('legend') ?? true),
        chartData(op),
      );
      return;
    default:
      throw new CliError(`unknown op '${kind}' — \`cellestra schema apply\` lists them.`);
  }
}

const chartLook = (op: Op) => ({
  title: op.string('title'),
  xTitle: op.string('xTitle'),
  yTitle: op.string('yTitle'),
  stacking: op.string('stacking'),
  aggregate: op.string('aggregate'),
});

const chartData = (op: Op) => ({
  values: op.strings('values'),
  names: op.strings('names'),
  categories: op.string('categories'),
  xValues: op.string('xValues'),
  bubbleSizes: op.string('bubbleSizes'),
  from: op.string('from'),
});

// ---- Cells -------------------------------------------------------------------

function set(session: EditSession, op: Op): void {
  const wb = session.workbook;
  if (op.has('values')) {
    const anchor = op.string('at') ?? op.string('range') ?? op.string('cell');
    if (anchor === null) throw new CliError('a block needs "at" (its top-left cell).');
    const { sheet, range } = resolveRange(wb, anchor, op.string('sheet'));
    const rows = op.raw('values');
    if (!Array.isArray(rows)) throw new CliError('"values" must be an array of rows (arrays).');
    rows.forEach((row, r) => {
      if (!Array.isArray(row)) throw new CliError(`"values" row ${r + 1} is not an array.`);
      row.forEach((value, c) => setOne(session, sheet, range.top + r, range.left + c, value));
    });
    return;
  }
  const { sheet, address: at } = resolveCell(wb, op.require('cell'), op.string('sheet'));
  if (op.has('formula')) {
    // The original {"cell", "formula"} form: the leading = is optional there.
    const formula = op.require('formula');
    enter(session, sheet, at.row, at.column, formula.startsWith('=') ? formula : '=' + formula);
    return;
  }
  setOne(session, sheet, at.row, at.column, op.raw('value'));
}

function setOne(
  session: EditSession,
  sheet: Worksheet,
  row: number,
  column: number,
  value: Json | undefined,
): void {
  if (value === null || value === undefined) {
    session.apply(EditCommands.clearContents(sheet.name, new CellRange(new CellAddress(column, row))));
    session.touched(sheet.name, row, column);
  } else if (typeof value === 'number') {
    session.apply(EditCommands.setValue(sheet.name, row, column, CellValue.number(value)));
  } else if (typeof value === 'boolean') {
    session.apply(EditCommands.setValue(sheet.name, row, column, CellValue.boolean(value)));
  } else if (typeof value === 'string') {
    enter(session, sheet, row, column, value);
  } else {
    throw new CliError(
      `a cell value must be a number, string, boolean, or null (at ${address(sheet.name, row, column)}).`,
    );
  }
}

/**
 * Typing into the app: the same parse (formulas, numbers, dates with a date format stamped, text),
 * and a table beside the cell grows to take it in. A value that breaks the cell's validation rule
 * is still written, with a note, so a script can record what it was given (`lint` reports it).
 */
function enter(session: EditSession, sheet: Worksheet, row: number, column: number, text: string): void {
  const entry = planCellEntry(session.workbook, sheet.name, row, column, text, session.locale, (range) =>
    readRange(sheet, range),
  );
  session.apply(entry.command);
  if (entry.brokenRule) session.note(`${address(sheet.name, row, column)} breaks its validation rule.`);
}

function* readRange(sheet: Worksheet, range: CellRange): Iterable<CellValue> {
  for (const [row, column] of range.positions()) yield sheet.getValue(row, column);
}

function clear(session: EditSession, op: Op): void {
  const { sheet, range } = rangeOf(session, op, 'range');
  let edit: EditCommand;
  switch ((op.string('what') ?? 'contents').toLowerCase()) {
    case 'contents':
    case 'values':
      edit = EditCommands.clearContents(sheet.name, range);
      break;
    case 'formats':
    case 'formatting':
    case 'styles':
      edit = EditCommands.clearFormatting(session.workbook, sheet.name, range);
      break;
    case 'all':
      edit = EditCommands.clearAll(sheet.name, range);
      break;
    default:
      throw new CliError('"what" must be contents, formats, or all.');
  }
  session.apply(edit);
}

// ---- Formatting --------------------------------------------------------------

type StyleChange = (s: CellStyle) => CellStyle;

function style(session: EditSession, op: Op): void {
  const { sheet, range } = rangeOf(session, op, 'range');
  const changes: StyleChange[] = [];
  const font = (patch: Partial<CellStyle['font']>): void => {
    changes.push((s) => ({ ...s, font: { ...s.font, ...patch } }));
  };
  const align = (patch: Partial<CellStyle['alignment']>): void => {
    changes.push((s) => ({ ...s, alignment: { ...s.alignment, ...patch } }));
  };

  const bold = op.bool('bold');
  if (bold !== null) font({ bold });
  const italic = op.bool('italic');
  if (italic !== null) font({ italic });
  const underline = op.bool('underline');
  if (underline !== null) font({ underline });
  const strikethrough = op.bool('strikethrough');
  if (strikethrough !== null) font({ strikethrough });
  const family = op.string('font');
  if (family !== null) font({ family });
  const size = op.number('size');
  if (size !== null) {
    if (size <= 0) throw new CliError('"size" must be a positive number of points.');
    font({ sizePoints: size });
  }
  const color = op.string('color');
  if (color !== null) font({ color: parseColor(color, 'color') });
  if (op.present('fill')) {
    const text = op.string('fill');
    const value: Fill =
      text === null || text.toLowerCase() === 'none' ? Fills.none : Fills.solid(parseColor(text, 'fill'));
    changes.push((s) => ({ ...s, fill: value }));
  }
  const alignText = op.string('align');
  if (alignText !== null) {
    let horizontal: HorizontalAlignment;
    switch (alignText.trim().toLowerCase()) {
      case 'general':
        horizontal = 'general';
        break;
      case 'left':
        horizontal = 'left';
        break;
      case 'center':
      case 'centre':
        horizontal = 'center';
        break;
      case 'right':
        horizontal = 'right';
        break;
      default:
        throw new CliError('"align" must be general, left, center, or right.');
    }
    align({ horizontal });
  }
  const valignText = op.string('valign');
  if (valignText !== null) {
    let vertical: VerticalAlignment;
    switch (valignText.trim().toLowerCase()) {
      case 'top':
        vertical = 'top';
        break;
      case 'middle':
      case 'center':
        vertical = 'middle';
        break;
      case 'bottom':
        vertical = 'bottom';
        break;
      default:
        throw new CliError('"valign" must be top, middle, or bottom.');
    }
    align({ vertical });
  }
  const wrap = op.bool('wrap');
  if (wrap !== null) align({ wrapText: wrap });
  const rotation = op.number('rotation');
  if (rotation !== null) {
    if (!Number.isFinite(rotation)) throw new CliError('"rotation" must be a number of degrees.');
    const degrees = ((Math.round(rotation) % 360) + 360) % 360;
    // As in Format Cells: unrotated is the absence of both keys, so it matches the default style.
    changes.push((s) => {
      if ((s.alignment.rotation ?? 0) === degrees) return s;
      const { rotation: _r, rotationAlign: _a, ...rest } = s.alignment;
      return { ...s, alignment: degrees === 0 ? rest : { ...s.alignment, rotation: degrees } };
    });
  }
  const format = op.string('format');
  if (format !== null) changes.push((s) => ({ ...s, numberFormat: { code: format } }));
  if (op.present('border')) {
    const text = (op.string('border') ?? 'none').trim().toLowerCase();
    const lines: readonly BorderLineStyle[] = [
      'none',
      'thin',
      'medium',
      'thick',
      'dashed',
      'dotted',
      'double',
    ];
    const line = lines.find((l) => l === text);
    if (line === undefined) {
      throw new CliError('"border" must be none, thin, medium, thick, dashed, dotted, or double.');
    }
    const borderColor = op.string('borderColor');
    const edge = BorderEdges.of(
      line,
      borderColor !== null ? parseColor(borderColor, 'borderColor') : Argb.black,
    );
    const border = line === 'none' ? Borders.none : { left: edge, right: edge, top: edge, bottom: edge };
    changes.push((s) => ({ ...s, border }));
  }
  const checkbox = op.bool('checkbox');
  if (changes.length === 0 && checkbox === null) {
    throw new CliError(
      'a style op needs at least one of: bold, italic, underline, strikethrough, font, size, color, fill, align, valign, wrap, rotation, format, border, checkbox.',
    );
  }

  if (changes.length > 0) {
    session.apply(
      EditCommands.applyStyle(session.workbook, sheet.name, range, (s) =>
        changes.reduce((acc, change) => change(acc), s),
      ),
    );
  }
  // A checkbox also needs a value: empty cells get FALSE so they show an unticked box.
  if (checkbox === true) session.apply(EditCommands.insertCheckboxes(session.workbook, sheet.name, range));
  if (checkbox === false) session.apply(EditCommands.removeCheckboxes(session.workbook, sheet.name, range));

  // applyStyle reports no cells (a style moves no values), but the report should still say what
  // was touched: every cell that now exists inside the range. Walk whichever is smaller — the
  // range, or the populated cells.
  if (range.rowCount * range.columnCount <= 4096) {
    for (const [row, column] of range.positions()) {
      if (sheet.find(row, column)) session.touched(sheet.name, row, column);
    }
  } else {
    for (const [row, column] of sheet.enumerateCells()) {
      if (range.contains(row, column)) session.touched(sheet.name, row, column);
    }
  }
}

function parseColor(text: string, field: string): ArgbColor {
  const color = Argb.tryFromHex(text);
  if (color === null) throw new CliError(`"${field}" must be a #RRGGBB color, not '${text}'.`);
  return color;
}

function width(session: EditSession, op: Op): void {
  const sheet = pickSheet(session.workbook, op.string('sheet'));
  const columns = op.string('columns') ?? op.string('column');
  if (columns === null) throw new CliError('needs "columns" (e.g. "B" or "B:D").');
  const [first, last] = parseColumnSpan(columns);
  const value = op.number('width');
  if (value === null) throw new CliError('needs "width" (points).');
  if (value <= 0) throw new CliError('"width" must be positive.');
  for (let column = first; column <= last; column++) {
    session.apply(EditCommands.setColumnWidth(sheet.name, column, value));
  }
}

function height(session: EditSession, op: Op): void {
  const sheet = pickSheet(session.workbook, op.string('sheet'));
  const rows = op.string('rows') ?? op.string('row');
  if (rows === null) throw new CliError('needs "rows" (e.g. "1" or "2:5").');
  const [first, last] = parseRowSpan(rows);
  const value = op.number('height');
  if (value === null) throw new CliError('needs "height" (points).');
  if (value <= 0) throw new CliError('"height" must be positive.');
  for (let row = first; row <= last; row++) session.apply(EditCommands.setRowHeight(sheet.name, row, value));
}

function freeze(session: EditSession, op: Op): void {
  const sheet = pickSheet(session.workbook, op.string('sheet'));
  const rows = op.int('rows', 0);
  const columns = op.int('columns', 0);
  if (rows < 0 || columns < 0) throw new CliError('"rows" and "columns" must be 0 or more.');
  session.apply(EditCommands.setFrozenPanes(sheet.name, rows, columns));
}

function visibility(session: EditSession, op: Op, hidden: boolean): void {
  const sheet = pickSheet(session.workbook, op.string('sheet'));
  const rows = op.string('rows');
  const columns = op.string('columns');
  if (rows !== null) {
    const [first, last] = parseRowSpan(rows);
    session.apply(EditCommands.setRowsHidden(sheet.name, first, last, hidden));
  } else if (columns !== null) {
    const [first, last] = parseColumnSpan(columns);
    session.apply(EditCommands.setColumnsHidden(sheet.name, first, last, hidden));
  } else {
    throw new CliError('needs "rows" or "columns".');
  }
}

// ---- Structure ---------------------------------------------------------------

function addSheet(session: EditSession, op: Op): void {
  const wb = session.workbook;
  const sheetName = op.require('name');
  const existing = wb.findSheet(sheetName);
  if (existing) {
    if (op.bool('ifMissing') === true) {
      session.field('sheet', existing.name); // already there: nothing to do, and not an error
      return;
    }
    throw new CliError(`a sheet named '${sheetName}' already exists (add "ifMissing": true to allow that).`);
  }
  const sheet = wb.addSheet(sheetName);
  if (op.has('at')) {
    const at = op.int('at', 0);
    if (at < 1 || at > wb.sheetCount) {
      throw new CliError(`"at" must be a 1-based position from 1 to ${wb.sheetCount}.`);
    }
    wb.moveSheet(wb.indexOf(sheet), at - 1);
  }
  session.field('sheet', sheet.name);
}

function settings(session: EditSession, op: Op): void {
  const wb = session.workbook;
  let any = false;
  const locale = op.string('locale');
  if (locale !== null) {
    wb.settings.locale = resolveLocale(locale);
    any = true;
  }
  const calc = op.string('calc');
  if (calc !== null) {
    switch (calc.trim().toLowerCase()) {
      case 'auto':
      case 'automatic':
        wb.settings.calculationMode = 'automatic';
        break;
      case 'manual':
        wb.settings.calculationMode = 'manual';
        break;
      default:
        throw new CliError('"calc" must be automatic or manual.');
    }
    any = true;
  }
  const active = op.string('activeSheet');
  if (active !== null) {
    wb.settings.activeSheetIndex = wb.indexOf(pickSheet(wb, active));
    any = true;
  }
  if (!any) throw new CliError('a settings op needs at least one of: locale, calc, activeSheet.');
}

function structural(
  session: EditSession,
  op: Op,
  axis: StructuralAxis,
  operation: StructuralOperation,
): void {
  const sheet = pickSheet(session.workbook, op.string('sheet'));
  const at = op.require('at');
  const index = axis === 'row' ? parseRow(at) : parseColumn(at);
  const count = op.int('count', 1);
  if (count < 1) throw new CliError('"count" must be at least 1.');
  checkStructuralEdit(sheet, axis, operation, index, count);
  session.apply(new StructuralEditCommand(sheet.name, axis, operation, index, count));
}

// ---- Data --------------------------------------------------------------------

function sort(session: EditSession, op: Op): void {
  const { sheet, range } = rangeOf(session, op, 'range');
  const levels = parseSortLevels(op.require('by'));
  session.apply(
    DataEditCommands.sort(
      sheet.name,
      range,
      { levels, hasHeader: op.bool('header') ?? true },
      session.locale,
    ),
  );
}

function removeDuplicates(session: EditSession, op: Op): void {
  const { sheet, range } = rangeOf(session, op, 'range');
  const keyColumns = op.strings('columns').map(parseColumn);
  const edit = DataEditCommands.removeDuplicates(sheet.name, range, keyColumns, op.bool('header') ?? true);
  session.apply(edit.command);
  session.field('removed', edit.removedCount());
  session.note(`Removed ${edit.removedCount()} duplicate row(s).`);
}

function split(session: EditSession, op: Op): void {
  const { sheet, range } = rangeOf(session, op, 'range');
  const delimiter = op.string('delimiter') ?? ',';
  if ([...delimiter].length !== 1) throw new CliError('"delimiter" must be a single character.');
  session.apply(DataEditCommands.textToColumns(sheet, range, delimiter));
}

function tableOp(session: EditSession, op: Op): void {
  const { sheet, range } = rangeOf(session, op, 'range');
  const plan = TableCommands.create(session.workbook, sheet.name, range, {
    name: op.string('name'),
    header: op.bool('header') ?? true,
    replace: op.bool('replace') === true,
  });
  if (!plan.ok) {
    throw new CliError(
      `the range overlaps table '${plan.overlapping[0]!.name}'; a cell can belong to only one table (add "replace": true to redefine it).`,
    );
  }
  session.apply(plan.command);
  session.field('table', plan.table.name);
}

function validation(session: EditSession, op: Op): void {
  const { sheet, range } = rangeOf(session, op, 'range');
  const existing = sheet.validations.filter((v) => v.target.intersects(range));
  if (op.bool('remove') === true) {
    if (existing.length === 0) throw new CliError(`no validation rule covers ${range.toString()}.`);
    session.apply(DataEditCommands.replaceValidation(sheet.name, range, null));
    session.field('removed', existing.length);
    return;
  }
  // One rule per cell: the new rule replaces the old ones over the range, which keep their other cells.
  session.apply(DataEditCommands.replaceValidation(sheet.name, range, buildValidation(session, op, range)));
}

function buildValidation(session: EditSession, op: Op, range: CellRange): DataValidation {
  const kindText = op.string('kind');
  let kind: ValidationKind;
  switch (kindText?.trim().toLowerCase()) {
    case undefined:
      if (op.has('list') || op.has('source')) {
        kind = 'list';
        break;
      }
      throw new CliError('needs "list" or "source" (a list rule), or "kind" whole|decimal|date|text-length.');
    case 'list':
      kind = 'list';
      break;
    case 'whole':
    case 'whole-number':
    case 'wholenumber':
    case 'integer':
      kind = 'wholeNumber';
      break;
    case 'decimal':
    case 'number':
      kind = 'decimal';
      break;
    case 'date':
      kind = 'date';
      break;
    case 'text-length':
    case 'textlength':
    case 'length':
      kind = 'textLength';
      break;
    case 'any':
      kind = 'any';
      break;
    default:
      throw new CliError('"kind" must be list, whole, decimal, date, or text-length.');
  }

  let response: ValidationResponse;
  switch ((op.string('response') ?? 'stop').trim().toLowerCase()) {
    case 'stop':
      response = 'stop';
      break;
    case 'warning':
    case 'warn':
      response = 'warning';
      break;
    case 'information':
    case 'info':
      response = 'information';
      break;
    case 'none':
      response = 'none';
      break;
    default:
      throw new CliError('"response" must be stop, warning, information, or none.');
  }
  const base = { inputMessage: op.string('inputMessage'), errorMessage: op.string('errorMessage'), response };

  if (kind === 'list') {
    const source = op.string('source');
    if (source !== null) {
      const listSource = CellRange.tryParse(source);
      if (!listSource) throw new CliError(`'${source}' is not a valid range for "source".`);
      return dataValidation(kind, range, { ...base, listSource });
    }
    const list = op.strings('list');
    if (list.length === 0) {
      throw new CliError('a list rule needs "list" (an array of allowed values) or "source" (a range).');
    }
    return dataValidation(kind, range, { ...base, listValues: list });
  }
  if (kind === 'any') return dataValidation(kind, range, base);

  const operators: Record<string, ValidationOperator> = {
    between: 'between',
    'not-between': 'notBetween',
    notbetween: 'notBetween',
    equal: 'equal',
    '=': 'equal',
    '==': 'equal',
    'not-equal': 'notEqual',
    notequal: 'notEqual',
    '<>': 'notEqual',
    '!=': 'notEqual',
    greater: 'greaterThan',
    'greater-than': 'greaterThan',
    '>': 'greaterThan',
    less: 'lessThan',
    'less-than': 'lessThan',
    '<': 'lessThan',
    'greater-or-equal': 'greaterOrEqual',
    '>=': 'greaterOrEqual',
    'less-or-equal': 'lessOrEqual',
    '<=': 'lessOrEqual',
  };
  const comparison = Object.hasOwn(operators, (op.string('operator') ?? 'between').trim().toLowerCase())
    ? operators[(op.string('operator') ?? 'between').trim().toLowerCase()]!
    : null;
  if (comparison === null) {
    throw new CliError(
      '"operator" must be between, not-between, equal, not-equal, greater, less, greater-or-equal, or less-or-equal.',
    );
  }
  const twoBounds = comparison === 'between' || comparison === 'notBetween';
  const value1 = bound(session, op, 'value1', kind);
  if (value1 === null) {
    throw new CliError(`a ${kindText} rule needs "value1"${twoBounds ? ' and "value2"' : ''}.`);
  }
  let value2 = 0;
  if (twoBounds) {
    const upper = bound(session, op, 'value2', kind);
    if (upper === null) throw new CliError('a between rule needs "value2" (the upper bound).');
    value2 = upper;
  }
  return dataValidation(kind, range, { ...base, operator: comparison, operand1: value1, operand2: value2 });
}

/** A rule bound: a number, or for a date rule any date the document's locale accepts (ISO always works). */
function bound(session: EditSession, op: Op, field: string, kind: ValidationKind): number | null {
  if (!op.has(field)) return null;
  const raw = op.raw(field);
  if (kind === 'date' && typeof raw === 'string') {
    const parsed = InputParser.parse(raw, session.locale);
    if (parsed.formula === null && parsed.value.kind === 'number') return parsed.value.value;
    throw new CliError(`"${field}" is not a date: '${raw}'.`);
  }
  return op.number(field);
}

function findTable(session: EditSession, op: Op): { sheet: Worksheet; table: Table } {
  const reference = op.require('table');
  const { sheet, range } = rangeOf(session, op, 'table'); // a table name resolves to its range; so does A1-style
  const lower = reference.toLowerCase();
  const table =
    sheet.tables.find((t) => t.name.toLowerCase() === lower) ??
    sheet.tables.find((t) => t.range.intersects(range));
  if (!table) throw new CliError(`no table at or named '${reference}'.`);
  return { sheet, table };
}

/** The sheet's stored copy: core swaps tables by identity, so each step must pass the live object. */
function current(sheet: Worksheet, name: string): Table {
  return sheet.tables.find((t) => t.name === name)!;
}

function tableOptions(session: EditSession, op: Op): void {
  const found = findTable(session, op);
  const sheet = found.sheet;
  let table = found.table;
  let any = false;
  const newRange = op.string('range');
  if (newRange !== null) {
    const target = resolveRange(session.workbook, newRange, sheet.name);
    if (target.sheet !== sheet) {
      throw new CliError("a table can't move to another sheet; remove it and create it there.");
    }
    if (sheet.tables.some((t) => t !== table && t.range.intersects(target.range))) {
      throw new CliError('the new range overlaps another table.');
    }
    const resized: Table = { ...table, range: target.range };
    session.apply(DataEditCommands.replaceTable(sheet.name, table, resized));
    table = current(sheet, table.name);
    any = true;
  }
  const banded = op.bool('bandedRows');
  if (banded !== null) {
    session.apply(DataEditCommands.setTableBandedRows(sheet.name, table, banded));
    table = current(sheet, table.name);
    any = true;
  }
  const headerColumn = op.bool('headerColumn');
  if (headerColumn !== null) {
    session.apply(DataEditCommands.setTableHeaderColumn(session.workbook, sheet.name, table, headerColumn));
    any = true;
  }
  if (!any) throw new CliError('a table-options op needs at least one of: range, bandedRows, headerColumn.');
  session.field('table', table.name);
}

function name(session: EditSession, op: Op): void {
  const { sheet, range } = rangeOf(session, op, 'range');
  const named = op.require('name');
  const error = namedRangeError(session.workbook, named);
  if (error !== null) throw new CliError(error);
  // Stored fully sheet-qualified, so the name resolves whatever sheet is active.
  session.workbook.addNamedRange({ name: named, target: range.withSheet(range.sheet ?? sheet.name) });
}

function hyperlink(session: EditSession, op: Op): void {
  const { sheet, address: at } = cellOf(session, op, 'cell');
  if (!op.present('url')) throw new CliError('needs "url" (a link, or null to remove the cell\'s link).');
  const url = op.string('url');
  const link = url === null ? null : { target: url, tooltip: op.string('tooltip') };
  session.apply(EditCommands.setHyperlink(session.workbook, sheet.name, at.row, at.column, link));
}

// ---- Address helpers ---------------------------------------------------------

function rangeOf(session: EditSession, op: Op, field: string): ResolvedRange {
  return resolveRange(session.workbook, op.require(field), op.string('sheet'));
}

function cellOf(session: EditSession, op: Op, field: string): ResolvedCell {
  return resolveCell(session.workbook, op.require(field), op.string('sheet'));
}
