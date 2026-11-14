import { existsSync } from 'node:fs';
import {
  CellAddress,
  CellRange,
  CellValue,
  CsvOptions,
  DeleteSheetCommand,
  EditCommands,
  FunctionRegistry,
  Workbook,
  fillFromCsv,
  parseCsv,
  writeCsv,
  type FunctionMetadata,
  type Worksheet,
} from '@cellestra/core';
import { display, formulaOrValue, pickSheet, resolveCell, resolveRange } from '../infra/cells.ts';
import { EditSession } from '../infra/editSession.ts';
import { CliError } from '../infra/errors.ts';
import type { Console } from '../infra/io.ts';
import { toCompactJson, toJson } from '../infra/json.ts';
import { resolveLocale } from '../infra/locale.ts';
import {
  address,
  cellReport,
  namedRangeReport,
  sheetReport,
  styleReport,
  writeReport,
  type CellReport,
} from '../infra/report.ts';
import {
  launch,
  load,
  readInputText,
  recalculate,
  saveText,
  type LoadedWorkbook,
} from '../infra/workspace.ts';
import { isFlag, type CommandSpec, type Parsed } from '../spec.ts';
import { dispatch } from './apply.ts';
import { capitalize } from './baseline.ts';
import { describeChart } from './chartOps.ts';
import {
  csvOptions,
  delimiterOpt,
  editTarget,
  fileArg,
  inPlaceOpt,
  jsonOpt,
  localeOpt,
  outputOpt,
  sheetOpt,
} from './options.ts';
import { splitList } from './parsers.ts';

/** The nice-to-haves: new, import/export, stats, find, diff, functions, sheet management, open, completions. */

export const newCommand: CommandSpec = {
  name: 'new',
  description: 'Create a new workbook',
  args: [fileArg('file', 'Workbook to create (.ods)')],
  options: [{ name: '--sheets', description: 'Comma-separated sheet names (default: Sheet1)' }, jsonOpt],
  run(p, con) {
    const path = p.arg('file');
    if (path !== '-' && existsSync(path))
      throw new CliError(`'${path}' already exists — refusing to overwrite.`);
    const names = splitList(p.str('--sheets') ?? '');
    const workbook = Workbook.createNew();
    if (names.length > 0) {
      workbook.renameSheet(workbook.activeSheet, names[0]!);
      for (const name of names.slice(1)) workbook.addSheet(name);
    }
    return EditSession.runLoaded(con, { workbook, package: null }, path, null, p.flag('--json'), (session) =>
      session.field(
        'sheets',
        workbook.worksheets.map((w) => w.name),
      ),
    );
  },
};

export const importCommand: CommandSpec = {
  name: 'import',
  description: 'Import a CSV into a sheet of an .ods workbook',
  args: [fileArg('csv', 'CSV file to import (or - for stdin)')],
  options: [
    { name: '--into', description: 'Target .ods workbook (created if absent)', required: true },
    { name: '--sheet', aliases: ['-s'], description: 'Sheet to fill (created/replaced)', required: true },
    delimiterOpt,
    localeOpt,
    jsonOpt,
  ],
  run(p, con) {
    const target = p.str('--into')!;
    const sheetName = p.str('--sheet')!;
    const locale = resolveLocale(p.str('--locale'));
    const csv = csvOptions(p.str('--delimiter'));
    const text = readInputText(con, p.arg('csv'));
    const loaded: LoadedWorkbook = existsSync(target)
      ? load(con, target, 'ods', locale)
      : { workbook: Workbook.createNew(), package: null };
    return EditSession.runLoaded(con, loaded, target, locale, p.flag('--json'), (session) => {
      const workbook = session.workbook;
      let worksheet = workbook.findSheet(sheetName);
      if (!worksheet) worksheet = workbook.addSheet(sheetName);
      else clearSheet(worksheet);
      fillFromCsv(workbook, worksheet, text, { delimiter: csv?.delimiter ?? ',', locale });
      session.field('sheet', worksheet.name);
      session.field('range', worksheet.getUsedRange()?.toString() ?? null);
      session.note(`Imported ${worksheet.cellCount} cell(s) into '${worksheet.name}'.`);
    });
  },
};

/**
 * Drop a grid at an anchor — the missing middle between `set` (one cell) and `import` (a whole
 * sheet). Each field is parsed like typing, so numbers, dates, and `=`formulas land typed.
 */
export const put: CommandSpec = {
  name: 'put',
  description: 'Write a CSV/TSV grid into a sheet at a cell (values and formulas, parsed like typing)',
  args: [fileArg(), { name: 'at', description: 'Top-left cell to write at, e.g. Sheet1.B3' }],
  options: [
    { name: '--from', description: 'CSV/TSV file to read the grid from (or - for stdin)', required: true },
    delimiterOpt,
    inPlaceOpt,
    outputOpt,
    localeOpt,
    jsonOpt,
  ],
  run(p, con) {
    const path = p.arg('file');
    const target = editTarget(path, p.flag('--in-place'), p.str('--output'));
    const source = p.str('--from')!;
    const csv =
      csvOptions(p.str('--delimiter')) ?? (source.toLowerCase().endsWith('.tsv') ? CsvOptions.tab : null);
    const rows = parseCsv(readInputText(con, source), csv ?? CsvOptions.default);
    return EditSession.run(
      con,
      path,
      target,
      resolveLocale(p.str('--locale')),
      p.flag('--json'),
      (session) => {
        const { sheet, address: anchor } = resolveCell(session.workbook, p.arg('at'));
        let written = 0;
        rows.forEach((row, r) => {
          row.forEach((text, c) => {
            session.apply(
              EditCommands.setInput(
                sheet.name,
                anchor.row + r,
                anchor.column + c,
                text,
                session.locale,
                session.workbook,
              ),
            );
            written++;
          });
        });
        let width = 0;
        for (const row of rows) width = Math.max(width, row.length);
        const covered =
          rows.length === 0
            ? null
            : new CellRange(
                new CellAddress(anchor.column, anchor.row),
                new CellAddress(anchor.column + width - 1, anchor.row + rows.length - 1),
              );
        session.field('sheet', sheet.name);
        session.field('range', covered?.toString() ?? null);
        session.note(`Wrote ${written} cell(s) at ${address(sheet.name, anchor.row, anchor.column)}.`);
      },
    );
  },
};

export const exportCommand: CommandSpec = {
  name: 'export',
  description: 'Export a sheet to CSV',
  args: [fileArg('file', 'Source .ods workbook'), fileArg('output', 'Destination CSV (or - for stdout)')],
  options: [sheetOpt, delimiterOpt],
  run(p, con) {
    const csv = csvOptions(p.str('--delimiter'));
    const loaded = load(con, p.arg('file'), 'ods', null);
    recalculate(loaded.workbook);
    const worksheet = pickSheet(loaded.workbook, p.str('--sheet'));
    saveText(con, writeCsv(loaded.workbook, worksheet, csv ?? CsvOptions.default), p.arg('output'));
    return 0;
  },
};

// ---- stats -------------------------------------------------------------------

/** .NET's `ToString("0.##########", InvariantCulture)`: up to ten decimals, trailing zeros trimmed. */
export function formatNumber(n: number): string {
  if (!Number.isFinite(n)) return String(n);
  let text = n.toFixed(10);
  if (text.includes('e')) return String(n);
  text = text.replace(/\.?0+$/, '');
  return text === '-0' ? '0' : text;
}

function median(numbers: readonly number[]): number | null {
  if (numbers.length === 0) return null;
  const sorted = [...numbers].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 1 ? sorted[mid]! : (sorted[mid - 1]! + sorted[mid]!) / 2;
}

/** Sample standard deviation (n − 1), the same as the STDEV function; null below two numbers. */
function sampleStdev(numbers: readonly number[]): number | null {
  if (numbers.length < 2) return null;
  const mean = numbers.reduce((a, b) => a + b, 0) / numbers.length;
  return Math.sqrt(numbers.reduce((a, n) => a + (n - mean) * (n - mean), 0) / (numbers.length - 1));
}

export const stats: CommandSpec = {
  name: 'stats',
  description: 'Describe a range: count, blanks, types, sum, average, median, stdev, min, max, distinct',
  args: [
    fileArg(),
    { name: 'range', description: 'Range to aggregate, e.g. Sheet1.A1:A100 (a name or table works too)' },
  ],
  options: [localeOpt, jsonOpt],
  run(p, con) {
    const locale = resolveLocale(p.str('--locale'));
    const loaded = load(con, p.arg('file'), null, locale);
    recalculate(loaded.workbook);
    const { sheet, range } = resolveRange(loaded.workbook, p.arg('range'));
    let count = 0;
    let texts = 0;
    let booleans = 0;
    let errors = 0;
    let formulas = 0;
    const numbers: number[] = [];
    const distinct = new Set<string>();
    // Walk only the populated cells, not every position in the bounding box — so
    // `stats Sheet1.A1:A1048576` over a sparse column stays fast.
    for (const [row, column, cell] of sheet.enumerateCells()) {
      if (!range.contains(row, column) || cell.value.kind === 'empty') continue;
      count++;
      distinct.add(CellValue.toInvariantString(cell.value));
      if (cell.isFormula) formulas++;
      switch (cell.value.kind) {
        case 'number':
          numbers.push(cell.value.value);
          break;
        case 'text':
          texts++;
          break;
        case 'boolean':
          booleans++;
          break;
        case 'error':
          errors++;
          break;
      }
    }
    const total = range.rowCount * range.columnCount;
    const blanks = total - count;
    const any = numbers.length > 0;
    const sum = any ? numbers.reduce((a, b) => a + b, 0) : null;
    const average = any ? sum! / numbers.length : null;
    // Loops, not Math.min(...numbers): a big range has more values than the stack holds arguments.
    let min: number | null = null;
    let max: number | null = null;
    for (const n of numbers) {
      if (min === null || n < min) min = n;
      if (max === null || n > max) max = n;
    }
    const med = median(numbers);
    const stdev = sampleStdev(numbers);
    if (p.flag('--json')) {
      writeReport(con, {
        sheet: sheet.name,
        range: range.toString(),
        cells: total,
        count,
        blanks,
        types: { number: numbers.length, text: texts, boolean: booleans, error: errors, formula: formulas },
        distinct: distinct.size,
        sum,
        average,
        median: med,
        stdev,
        min,
        max,
      });
      return 0;
    }
    const n = (v: number | null): string => (v === null ? '' : formatNumber(v));
    const lines: [string, string | number][] = [
      ['cells', total],
      ['count', count],
      ['blanks', blanks],
      ['numeric', numbers.length],
      ['text', texts],
      ['boolean', booleans],
      ['error', errors],
      ['formula', formulas],
      ['distinct', distinct.size],
      ['sum', n(sum)],
      ['average', n(average)],
      ['median', n(med)],
      ['stdev', n(stdev)],
      ['min', n(min)],
      ['max', n(max)],
    ];
    for (const [label, value] of lines) con.outLine(`${label}\t${value}`);
    return 0;
  },
};

// ---- find --------------------------------------------------------------------

const escapeRegex = (text: string): string => text.replace(/[.*+?^${}()|[\]\\/-]/g, '\\$&');

export const find: CommandSpec = {
  name: 'find',
  description: 'Search cells by value or formula — and optionally replace',
  args: [fileArg(), { name: 'term', description: 'Text to search for (case-insensitive unless --regex)' }],
  options: [
    { name: '--regex', description: 'Treat the term as a regular expression', kind: 'flag' },
    { name: '--in', description: 'Where to search: values | formulas | both', default: 'both' },
    {
      name: '--replace',
      description:
        'Replace every match with this text, in text cells and formulas (needs --in-place or -o); $1 groups with --regex',
    },
    sheetOpt,
    inPlaceOpt,
    outputOpt,
    localeOpt,
    jsonOpt,
  ],
  run(p, con) {
    const where = (p.str('--in') ?? 'both').toLowerCase();
    const searchValues = where === 'values' || where === 'both';
    const searchFormulas = where === 'formulas' || where === 'both';
    if (!searchValues && !searchFormulas) throw new CliError('--in must be values, formulas, or both.');
    const termText = p.arg('term');
    let source: string;
    let flags: string;
    if (p.flag('--regex')) {
      source = termText;
      flags = '';
    } else {
      source = escapeRegex(termText);
      flags = 'i';
    }
    let test: RegExp;
    let replaceAll: RegExp;
    try {
      test = new RegExp(source, flags);
      replaceAll = new RegExp(source, flags + 'g');
    } catch (e) {
      throw new CliError(`Invalid regex '${termText}': ${(e as Error).message}`);
    }
    const isMatch = (text: string): boolean => test.test(text);
    const replacement = p.str('--replace');
    const path = p.arg('file');
    const only = p.str('--sheet');
    const inScope = (s: Worksheet): boolean => only === null || s.name.toLowerCase() === only.toLowerCase();
    const json = p.flag('--json');
    const locale = resolveLocale(p.str('--locale'));

    if (replacement !== null) {
      const target = editTarget(path, p.flag('--in-place'), p.str('--output'));
      return EditSession.run(con, path, target, locale, json, (session) => {
        let replaced = 0;
        for (const worksheet of session.workbook.worksheets.filter(inScope)) {
          // Snapshot first: the edits mutate the sheet we're walking.
          const hits = [...worksheet.enumerateCells()];
          for (const [row, column, cell] of hits) {
            // Match what `find` matched: the formula as it's shown, leading '=' and all.
            if (searchFormulas && cell.formula !== null && isMatch('=' + cell.formula)) {
              const formula = ('=' + cell.formula).replace(replaceAll, replacement);
              session.apply(
                EditCommands.setFormula(
                  worksheet.name,
                  row,
                  column,
                  formula.startsWith('=') ? formula.slice(1) : formula,
                ),
              );
              replaced++;
            } else if (
              searchValues &&
              cell.formula === null &&
              cell.value.kind === 'text' &&
              isMatch(cell.value.value)
            ) {
              const text = cell.value.value.replace(replaceAll, replacement);
              session.apply(EditCommands.setValue(worksheet.name, row, column, CellValue.text(text)));
              replaced++;
            }
          }
        }
        session.field('replaced', replaced);
        session.note(`Replaced in ${replaced} cell(s).`);
      });
    }

    const loaded = load(con, path, null, locale);
    recalculate(loaded.workbook);
    const effective = locale ?? loaded.workbook.settings.effectiveLocale;
    let found = 0;
    const matches: CellReport[] = [];
    for (const worksheet of loaded.workbook.worksheets.filter(inScope)) {
      for (const [row, column, cell] of worksheet.enumerateCells()) {
        const match =
          (searchFormulas && cell.formula !== null && isMatch('=' + cell.formula)) ||
          (searchValues && isMatch(display(loaded.workbook, worksheet, row, column, effective)));
        if (!match) continue;
        found++;
        if (json) {
          matches.push(cellReport(loaded.workbook, worksheet, row, column, effective));
          continue;
        }
        con.outLine(
          `${address(worksheet.name, row, column)}\t${formulaOrValue(loaded.workbook, worksheet, row, column, effective)}`,
        );
      }
    }
    if (json) writeReport(con, { term: termText, matches });
    return found > 0 ? 0 : 1;
  },
};

// ---- diff --------------------------------------------------------------------

export const diff: CommandSpec = {
  name: 'diff',
  description:
    'Compare two workbooks: values/formulas, structure (sheets, layout, tables, charts, names), and with --styles formatting (exit 1 if they differ)',
  args: [fileArg('before', 'First workbook'), fileArg('after', 'Second workbook')],
  options: [
    { name: '--format', aliases: ['-f'], description: 'unified (default)', default: 'unified' },
    {
      name: '--styles',
      description: 'Also compare cell formatting (bold, fill, number format, …)',
      kind: 'flag',
    },
    jsonOpt,
  ],
  run(p, con) {
    const a = load(con, p.arg('before'), null, null);
    const b = load(con, p.arg('after'), null, null);
    recalculate(a.workbook);
    recalculate(b.workbook);
    const locale = 'en-US'; // .NET's invariant culture formats numbers the way en-US does
    const json = p.flag('--json');
    const compareStyles = p.flag('--styles');
    const addedSheets: string[] = [];
    const removedSheets: string[] = [];
    const cells: Record<string, unknown>[] = [];
    const structure: Record<string, unknown>[] = [];
    const sheetNames: string[] = [];
    for (const s of [...a.workbook.worksheets, ...b.workbook.worksheets]) {
      if (!sheetNames.some((n) => n.toLowerCase() === s.name.toLowerCase())) sheetNames.push(s.name);
    }
    const line = (text: string): void => {
      if (!json) con.outLine(text);
    };

    for (const name of sheetNames) {
      const sheetA = a.workbook.findSheet(name);
      const sheetB = b.workbook.findSheet(name);
      if (!sheetA) {
        addedSheets.push(name);
        line(`+ sheet ${name}`);
        continue;
      }
      if (!sheetB) {
        removedSheets.push(name);
        line(`- sheet ${name}`);
        continue;
      }
      for (const [row, column] of usedUnion(sheetA, sheetB)) {
        const before = formulaOrValue(a.workbook, sheetA, row, column, locale);
        const after = formulaOrValue(b.workbook, sheetB, row, column, locale);
        const cell = address(name, row, column);
        if (before !== after) {
          cells.push({ cell, before, after });
          line(`${cell}: "${before}" -> "${after}"`);
        }
        if (compareStyles) {
          const styleA = styleText(a.workbook, sheetA, row, column);
          const styleB = styleText(b.workbook, sheetB, row, column);
          if (styleA !== styleB) {
            cells.push({ cell, styleBefore: styleA, styleAfter: styleB });
            line(`${cell} style: ${styleA} -> ${styleB}`);
          }
        }
      }
      // Layout and structure: the same description `sheets --json` prints, compared key by key.
      compareObjects(name, sheetReport(sheetA), sheetReport(sheetB), structure, line, [
        'cells',
        'rows',
        'columns',
        'usedRange',
      ]);
      if (sheetA.isChartSheet && sheetB.isChartSheet) {
        compareObjects(name, describeChart(sheetA), describeChart(sheetB), structure, line, ['sheet']);
      }
    }
    compareObjects(
      '(workbook)',
      { namedRanges: a.workbook.namedRanges.map(namedRangeReport) },
      { namedRanges: b.workbook.namedRanges.map(namedRangeReport) },
      structure,
      line,
      [],
    );
    const differences = addedSheets.length + removedSheets.length + cells.length + structure.length;
    if (json) {
      writeReport(con, {
        same: differences === 0,
        differences,
        addedSheets,
        removedSheets,
        cells,
        structure,
      });
    }
    return differences > 0 ? 1 : 0;
  },
};

function styleText(workbook: Workbook, sheet: Worksheet, row: number, column: number): string {
  const cell = sheet.find(row, column);
  const style = cell ? styleReport(workbook.styles.get(cell.styleId)) : null;
  return style === null ? '{}' : toCompactJson(style);
}

/** Report every top-level property whose JSON differs between two descriptions of the same thing. */
function compareObjects(
  scope: string,
  before: Record<string, unknown>,
  after: Record<string, unknown>,
  into: Record<string, unknown>[],
  line: (text: string) => void,
  skip: readonly string[],
): void {
  // Round-trip through the report spelling first, so a null property is absent on both sides.
  const docA = JSON.parse(toJson(before)) as Record<string, unknown>;
  const docB = JSON.parse(toJson(after)) as Record<string, unknown>;
  const keys = [...new Set([...Object.keys(docA), ...Object.keys(docB)])].filter((k) => !skip.includes(k));
  for (const key of keys) {
    const valueA = Object.hasOwn(docA, key) ? rawText(docA[key]) : null;
    const valueB = Object.hasOwn(docB, key) ? rawText(docB[key]) : null;
    if (valueA !== valueB) {
      into.push({ sheet: scope, property: key, before: valueA, after: valueB });
      line(`${scope} ${key}: ${compact(valueA)} -> ${compact(valueB)}`);
    }
  }
}

/** A top-level property's text as it sits in the indented document (one level in). */
const rawText = (value: unknown): string => JSON.stringify(value, null, 2).replaceAll('\n', '\n  ');

const compact = (json: string | null): string => (json === null ? '(none)' : json.replace(/\s*\n\s*/g, ' '));

/** Reading order, so a diff of the same two files always lists its cells the same way. */
function usedUnion(a: Worksheet, b: Worksheet): [number, number][] {
  const seen = new Map<string, [number, number]>();
  for (const sheet of [a, b]) {
    for (const [row, column] of sheet.enumerateCells()) seen.set(`${row},${column}`, [row, column]);
  }
  return [...seen.values()].sort((x, y) => x[0] - y[0] || x[1] - y[1]);
}

// ---- functions ---------------------------------------------------------------

function functionJson(m: FunctionMetadata): Record<string, unknown> {
  return {
    name: m.name,
    signature: m.signature,
    category: capitalize(m.category),
    returns: capitalize(m.returnType),
    summary: m.summary,
    arguments: m.args.map((a) => ({
      name: a.name,
      kind: capitalize(a.type),
      required: !a.optional,
      repeatable: a.repeatable,
      description: a.description ?? null,
    })),
    example: m.example,
  };
}

export const functions: CommandSpec = {
  name: 'functions',
  description: 'List supported functions, or describe one',
  args: [{ name: 'name', description: 'A function to describe (omit to list all)', arity: 'optional' }],
  options: [jsonOpt],
  run(p, con) {
    const registry = FunctionRegistry.createDefault();
    const wanted = p.argOrNull('name');
    const json = p.flag('--json');
    if (wanted === null || wanted === '') {
      if (json) {
        writeReport(con, registry.allMetadata().map(functionJson));
        return 0;
      }
      for (const meta of registry.allMetadata())
        con.outLine(`${meta.signature.padEnd(40)}  ${capitalize(meta.category)}`);
      return 0;
    }
    const fn = registry.get(wanted);
    if (!fn) throw new CliError(`No function named '${wanted}'.`);
    const m = fn.metadata;
    if (json) {
      writeReport(con, functionJson(m));
      return 0;
    }
    con.outLine(m.signature);
    con.outLine(`Category: ${capitalize(m.category)}`);
    con.outLine(`Returns:  ${capitalize(m.returnType)}`);
    con.outLine(m.summary);
    if (m.args.length > 0) {
      con.outLine('Arguments:');
      for (const arg of m.args) {
        const flags = (arg.optional ? 'optional' : 'required') + (arg.repeatable ? ', repeatable' : '');
        con.outLine(
          `  ${arg.name} (${capitalize(arg.type)}, ${flags})${arg.description === undefined ? '' : ' — ' + arg.description}`,
        );
      }
    }
    con.outLine(`Example: ${m.example}`);
    return 0;
  },
};

// ---- Sheets ------------------------------------------------------------------

/** The edit contract every writing verb follows: nothing overwrites without --in-place or -o. */
function mutate(p: Parsed, con: Console, edit: (session: EditSession) => void): number {
  const path = p.arg('file');
  const target = editTarget(path, p.flag('--in-place'), p.str('--output'));
  return EditSession.run(con, path, target, null, p.flag('--json'), edit);
}

export const addSheet: CommandSpec = {
  name: 'add-sheet',
  description: 'Add a worksheet',
  args: [fileArg(), { name: 'name', description: 'New sheet name' }],
  options: [
    { name: '--at', description: '1-based position for the new sheet (default: last)', type: 'int' },
    {
      name: '--if-missing',
      description: 'Succeed quietly if a sheet of that name already exists',
      kind: 'flag',
    },
    inPlaceOpt,
    outputOpt,
    jsonOpt,
  ],
  run: (p, con) =>
    dispatch(p, con, 'add-sheet', {
      name: p.arg('name'),
      at: p.num('--at'),
      ifMissing: p.flag('--if-missing') ? true : null,
    }),
};

export const renameSheet: CommandSpec = {
  name: 'rename-sheet',
  description: 'Rename a worksheet',
  args: [
    fileArg(),
    { name: 'old', description: 'Existing sheet name' },
    { name: 'new', description: 'New sheet name' },
  ],
  options: [inPlaceOpt, outputOpt, jsonOpt],
  run: (p, con) =>
    mutate(p, con, (session) => {
      session.workbook.renameSheet(pickSheet(session.workbook, p.arg('old')), p.arg('new'));
      session.field('sheet', p.arg('new'));
    }),
};

export const deleteSheet: CommandSpec = {
  name: 'delete-sheet',
  description: 'Delete a worksheet (dependent formulas become #REF!)',
  args: [fileArg(), { name: 'name', description: 'Sheet to delete' }],
  options: [inPlaceOpt, outputOpt, jsonOpt],
  run: (p, con) =>
    mutate(p, con, (session) => {
      const sheet = pickSheet(session.workbook, p.arg('name'));
      session.apply(new DeleteSheetCommand(sheet.name));
      session.field('sheet', sheet.name);
    }),
};

/** The explicit spelling of what `cellestra <file>` also does. */
export const open: CommandSpec = {
  name: 'open',
  description: 'Open a workbook in the Cellestra desktop app',
  args: [fileArg('file', 'Workbook to open in the Cellestra GUI')],
  run: (p, con) => launch(con, p.arg('file')),
};

// ---- completions -------------------------------------------------------------

const verbs = (parent: CommandSpec): CommandSpec[] =>
  [...(parent.subcommands ?? [])].sort((a, b) => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0));

const ordinal = (a: string, b: string): number => (a < b ? -1 : a > b ? 1 : 0);

function optionNames(command: CommandSpec): string[] {
  return (command.options ?? []).flatMap((o) => [...(o.aliases ?? []), o.name]).sort(ordinal);
}

const quote = (text: string): string =>
  text.replaceAll("'", "'\\''").replaceAll('[', '\\[').replaceAll(']', '\\]');

/** zsh: a _describe'd command list, then per-command _arguments with every option and its help; groups dispatch again. */
function zshCompletion(root: CommandSpec): string {
  let sb = '#compdef cellestra\n\n';
  sb += '_cellestra_options() {\n  local -a opts\n  case "$1" in\n';
  for (const verb of verbs(root)) {
    if ((verb.subcommands ?? []).length > 0) {
      for (const sub of verbs(verb)) sb += `    '${verb.name} ${sub.name}') opts=(${zshOptions(sub)}) ;;\n`;
      continue;
    }
    sb += `    '${verb.name}') opts=(${zshOptions(verb)}) ;;\n`;
  }
  sb += "  esac\n  _arguments -s $opts '*:file:_files'\n}\n\n";
  sb += '_cellestra() {\n  local -a cmds\n  cmds=(\n';
  for (const verb of verbs(root)) sb += `    '${verb.name}:${quote(verb.description)}'\n`;
  sb += "  )\n  if (( CURRENT == 2 )); then\n    _describe 'command' cmds\n    return\n  fi\n";
  sb += '  local cmd=$words[2]\n  case "$cmd" in\n';
  for (const group of verbs(root).filter((v) => (v.subcommands ?? []).length > 0)) {
    sb += `    ${group.name})\n      if (( CURRENT == 3 )); then\n        local -a subs\n        subs=(\n`;
    for (const sub of verbs(group)) sb += `          '${sub.name}:${quote(sub.description)}'\n`;
    sb +=
      '        )\n        _describe \'subcommand\' subs\n      else\n        _cellestra_options "$cmd $words[3]"\n      fi\n      ;;\n';
  }
  sb += '    *) _cellestra_options "$cmd" ;;\n  esac\n}\n\ncompdef _cellestra cellestra\n';
  return sb;
}

function zshOptions(command: CommandSpec): string {
  const parts: string[] = [];
  for (const option of command.options ?? []) {
    const help = quote(option.description);
    const value = isFlag(option) ? '' : ':value:';
    for (const name of [...(option.aliases ?? []), option.name]) parts.push(`'${name}[${help}]${value}'`);
  }
  return parts.join(' ');
}

/** bash: commands at word 1, subcommands for groups, each command's options after a dash; files otherwise. */
function bashCompletion(root: CommandSpec): string {
  let sb =
    '_cellestra() {\n  local cur="${COMP_WORDS[COMP_CWORD]}"\n  local cmd="${COMP_WORDS[1]}"\n  local opts=""\n';
  sb += `  if [[ $COMP_CWORD -eq 1 ]]; then\n    COMPREPLY=( $(compgen -W "${verbs(root)
    .map((v) => v.name)
    .join(' ')}" -- "$cur") )\n    return\n  fi\n`;
  sb += '  case "$cmd" in\n';
  for (const verb of verbs(root)) {
    if ((verb.subcommands ?? []).length > 0) {
      sb += `    ${verb.name})\n      if [[ $COMP_CWORD -eq 2 ]]; then\n        COMPREPLY=( $(compgen -W "${verbs(
        verb,
      )
        .map((s) => s.name)
        .join(' ')}" -- "$cur") )\n        return\n      fi\n      case "\${COMP_WORDS[2]}" in\n`;
      for (const sub of verbs(verb)) sb += `        ${sub.name}) opts="${optionNames(sub).join(' ')}" ;;\n`;
      sb += '      esac\n      ;;\n';
      continue;
    }
    sb += `    ${verb.name}) opts="${optionNames(verb).join(' ')}" ;;\n`;
  }
  sb +=
    '  esac\n  if [[ "$cur" == -* ]]; then\n    COMPREPLY=( $(compgen -W "$opts" -- "$cur") )\n  else\n    COMPREPLY=( $(compgen -f -- "$cur") )\n  fi\n}\ncomplete -o filenames -F _cellestra cellestra\n';
  return sb;
}

export const completions: CommandSpec = {
  name: 'completions',
  description: 'Print a shell completion script (commands, subcommands, and every option)',
  args: [{ name: 'shell', description: 'zsh | bash' }],
  run(p, con, root) {
    let script: string;
    switch (p.arg('shell').toLowerCase()) {
      case 'zsh':
        script = zshCompletion(root);
        break;
      case 'bash':
        script = bashCompletion(root);
        break;
      default:
        throw new CliError('Supported shells: zsh, bash.');
    }
    con.outLine(script);
    return 0;
  },
};

// ---- Helpers -----------------------------------------------------------------

function clearSheet(sheet: Worksheet): void {
  // Snapshot populated positions first (can't mutate while enumerating), then clear them.
  const positions = [...sheet.enumerateCells()].map(([row, column]) => [row, column] as const);
  for (const [row, column] of positions) sheet.putCell(row, column, null);
}
