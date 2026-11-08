import {
  CellAddress,
  Defaults,
  EditCommands,
  Evaluator,
  FunctionRegistry,
  NumberFormatter,
  parseErrorToString,
  Parser,
  Workbook,
  WorkbookEvaluationContext,
  writeCsv,
  type CellRange,
  type Worksheet,
} from '@cellestra/core';
import {
  buildGrid,
  display,
  formulaOrValue,
  pickSheet,
  resolveCell,
  resolveRange,
  usedRangeOrA1,
} from '../infra/cells.ts';
import { EditSession } from '../infra/editSession.ts';
import { CliError } from '../infra/errors.ts';
import { resolveLocale } from '../infra/locale.ts';
import { parseFormat, render } from '../infra/output.ts';
import {
  address,
  cellReport,
  errorReports,
  namedRangeReport,
  sheetReport,
  writeReport,
} from '../infra/report.ts';
import {
  detectFormat,
  load,
  recalculate,
  saveOds,
  saveText,
  type LoadedWorkbook,
} from '../infra/workspace.ts';
import type { CommandSpec } from '../spec.ts';
import {
  csvOptions,
  delimiterOpt,
  editTarget,
  fileArg,
  formatOpt,
  formatOverride,
  formulaOpt,
  fromOpt,
  inPlaceOpt,
  jsonOpt,
  localeOpt,
  outputOpt,
  rangeOpt,
  sheetOpt,
  toOpt,
} from './options.ts';

/** The genuinely useful core: convert, cat, get, set, recalc, sheets, info, eval, validate. */

export const capitalize = (s: string): string => s.charAt(0).toUpperCase() + s.slice(1);

export const convert: CommandSpec = {
  name: 'convert',
  description: 'Convert between .ods and CSV (both directions)',
  args: [
    fileArg('input', 'Source file (or - for stdin)'),
    fileArg('output', 'Destination file (or - for stdout)'),
  ],
  options: [fromOpt, toOpt, sheetOpt, localeOpt, delimiterOpt],
  run(p, con) {
    const inputPath = p.arg('input');
    const outputPath = p.arg('output');
    const locale = resolveLocale(p.str('--locale'));
    const csv = csvOptions(p.str('--delimiter'));
    const loaded = load(con, inputPath, formatOverride(p.str('--from')), locale, csv);
    const outFormat = detectFormat(outputPath, formatOverride(p.str('--to')));
    const sheetName = p.str('--sheet');
    if (outFormat === 'ods') {
      if (sheetName !== null && loaded.package === null) {
        loaded.workbook.renameSheet(loaded.workbook.activeSheet, sheetName);
      }
      recalculate(loaded.workbook);
      saveOds(con, loaded.workbook, loaded.package, outputPath);
    } else {
      const worksheet = pickSheet(loaded.workbook, sheetName);
      recalculate(loaded.workbook);
      saveText(
        con,
        writeCsv(loaded.workbook, worksheet, { delimiter: csv?.delimiter ?? ',', locale }),
        outputPath,
      );
    }
    return 0;
  },
};

export const cat: CommandSpec = {
  name: 'cat',
  description: 'Print a sheet or range to the terminal',
  args: [fileArg()],
  options: [
    sheetOpt,
    rangeOpt,
    formatOpt,
    formulaOpt,
    {
      name: '--both',
      description: 'Show formula cells as `=A1*2 → 20` (source and result together)',
      kind: 'flag',
    },
    localeOpt,
    fromOpt,
    jsonOpt,
  ],
  run(p, con) {
    const locale = resolveLocale(p.str('--locale'));
    const loaded = load(con, p.arg('file'), formatOverride(p.str('--from')), locale);
    recalculate(loaded.workbook);
    const effective = locale ?? loaded.workbook.settings.effectiveLocale;
    const rangeText = p.str('--range');
    let worksheet: Worksheet;
    let range: CellRange;
    if (rangeText !== null) {
      ({ sheet: worksheet, range } = resolveRange(loaded.workbook, rangeText, p.str('--sheet')));
    } else {
      worksheet = pickSheet(loaded.workbook, p.str('--sheet'));
      range = usedRangeOrA1(worksheet);
    }
    const grid = buildGrid(
      loaded.workbook,
      worksheet,
      range,
      effective,
      p.flag('--formula'),
      p.flag('--both'),
    );
    const format = p.flag('--json') ? 'json' : parseFormat(p.str('--format'));
    if (format === 'json') {
      // The grid of display strings for a quick look, plus every populated cell in full.
      const cells = [...worksheet.enumerateCells()]
        .filter(([r, c]) => range.contains(r, c))
        .sort((a, b) => a[0] - b[0] || a[1] - b[1])
        .map(([r, c]) => cellReport(loaded.workbook, worksheet, r, c, effective));
      writeReport(con, { sheet: worksheet.name, range: range.toString(), rows: grid, cells });
      return 0;
    }
    con.out(render(format, worksheet.name, range, grid));
    return 0;
  },
};

export const get: CommandSpec = {
  name: 'get',
  description: 'Read cell values (or --formula for sources): one cell, several, or a range',
  args: [
    fileArg(),
    {
      name: 'cells',
      description: 'Cell(s) or range(s): Sheet1.B7, B7, A1:C3, a named range, a table name — several at once',
      arity: 'many',
    },
  ],
  options: [formulaOpt, localeOpt, jsonOpt],
  run(p, con) {
    const locale = resolveLocale(p.str('--locale'));
    const loaded = load(con, p.arg('file'), null, locale);
    recalculate(loaded.workbook);
    const effective = locale ?? loaded.workbook.settings.effectiveLocale;
    const references = p.args('cells');
    const targets: [Worksheet, number, number][] = [];
    for (const reference of references) {
      if (CellAddress.tryParse(reference)) {
        const { sheet, address: a } = resolveCell(loaded.workbook, reference);
        targets.push([sheet, a.row, a.column]);
        continue;
      }
      const { sheet, range } = resolveRange(loaded.workbook, reference);
      if (range.rowCount * range.columnCount > 1_000_000) {
        throw new CliError(
          `'${reference}' covers ${range.rowCount}×${range.columnCount} cells — use \`cat --range\` for a block that size.`,
        );
      }
      for (const [r, c] of range.positions()) targets.push([sheet, r, c]);
    }
    const single = references.length === 1 && targets.length === 1;
    if (p.flag('--json')) {
      const reports = targets.map(([s, r, c]) => cellReport(loaded.workbook, s, r, c, effective));
      writeReport(con, single ? reports[0] : reports);
      return 0;
    }
    const sources = p.flag('--formula');
    for (const [sheet, row, column] of targets) {
      const text = sources
        ? formulaOrValue(loaded.workbook, sheet, row, column, effective)
        : display(loaded.workbook, sheet, row, column, effective);
      con.outLine(single ? text : `${address(sheet.name, row, column)}\t${text}`);
    }
    return 0;
  },
};

/**
 * `set file B2 42` is the original two-argument form; otherwise every argument is a `cell=value`
 * pair split at its first `=`, so a formula reads `B2==SUM(A1:A9)`.
 */
function parseSetPairs(args: readonly string[]): [string, string][] {
  if (args.length === 2 && !args[0]!.includes('=')) return [[args[0]!, args[1]!]];
  return args.map((argument) => {
    const split = argument.indexOf('=');
    if (split <= 0) {
      throw new CliError(
        `'${argument}' is not a cell=value pair (e.g. B2=42, or B2==SUM(A1:A9) for a formula).`,
      );
    }
    return [argument.slice(0, split), argument.slice(split + 1)];
  });
}

export const set: CommandSpec = {
  name: 'set',
  description: 'Set one or more cells and save the workbook (atomic)',
  args: [
    fileArg(),
    {
      name: 'edits',
      description:
        'Either `<cell> <value>` or any number of `cell=value` pairs (a formula: `B7==SUM(B1:B6)`)',
      arity: 'many',
    },
  ],
  options: [inPlaceOpt, outputOpt, localeOpt, jsonOpt],
  run(p, con) {
    const path = p.arg('file');
    const target = editTarget(path, p.flag('--in-place'), p.str('--output'));
    const pairs = parseSetPairs(p.args('edits'));
    return EditSession.run(
      con,
      path,
      target,
      resolveLocale(p.str('--locale')),
      p.flag('--json'),
      (session) => {
        for (const [reference, value] of pairs) {
          const { sheet, address: a } = resolveCell(session.workbook, reference);
          session.apply(
            EditCommands.setInput(sheet.name, a.row, a.column, value, session.locale, session.workbook),
          );
        }
      },
      null,
    );
  },
};

export const recalc: CommandSpec = {
  name: 'recalc',
  description: "Recalculate and refresh the file's cached values",
  args: [fileArg()],
  options: [inPlaceOpt, outputOpt, localeOpt, jsonOpt],
  run(p, con) {
    const path = p.arg('file');
    const target = editTarget(path, p.flag('--in-place'), p.str('--output'));
    return EditSession.run(
      con,
      path,
      target,
      resolveLocale(p.str('--locale')),
      p.flag('--json'),
      (session) => {
        let formulas = 0;
        for (const sheet of session.workbook.worksheets) {
          for (const [, , cell] of sheet.enumerateCells()) if (cell.formula !== null) formulas++;
        }
        session.note(`Recalculated ${formulas} formula cell(s).`);
      },
      null,
    );
  },
};

export const sheets: CommandSpec = {
  name: 'sheets',
  description: 'List worksheets with their used dimensions',
  args: [fileArg()],
  options: [jsonOpt],
  run(p, con) {
    const loaded = load(con, p.arg('file'), null, null);
    if (p.flag('--json')) {
      writeReport(con, loaded.workbook.worksheets.map(sheetReport));
      return 0;
    }
    for (const worksheet of loaded.workbook.worksheets) {
      const used = worksheet.getUsedRange();
      con.outLine(`${worksheet.name}\t${used ? used.toString() : 'empty'}\t${worksheet.cellCount} cells`);
    }
    return 0;
  },
};

const populated = (wb: Workbook): number => wb.worksheets.reduce((n, s) => n + s.cellCount, 0);

export const info: CommandSpec = {
  name: 'info',
  description: 'Workbook overview: sheets, named ranges, locale, cell counts',
  args: [fileArg()],
  options: [jsonOpt],
  run(p, con) {
    const loaded = load(con, p.arg('file'), null, null);
    const workbook = loaded.workbook;
    const calculation = capitalize(workbook.settings.calculationMode);
    if (p.flag('--json')) {
      writeReport(con, {
        file: p.arg('file'),
        sheets: workbook.worksheets.map(sheetReport),
        namedRanges: workbook.namedRanges.map(namedRangeReport),
        locale: workbook.settings.locale,
        calculation,
        populatedCells: populated(workbook),
      });
      return 0;
    }
    con.outLine(`Sheets: ${workbook.sheetCount}`);
    for (const worksheet of workbook.worksheets) {
      const used = worksheet.getUsedRange();
      con.outLine(`  ${worksheet.name}  ${used ? used.toString() : 'empty'}  ${worksheet.cellCount} cells`);
    }
    con.outLine(`Named ranges: ${workbook.namedRanges.length}`);
    for (const named of workbook.namedRanges) {
      con.outLine(`  ${named.name} = ${named.target.toString()}  (${named.scope ?? 'workbook'})`);
    }
    con.outLine(`Locale: ${workbook.settings.locale ?? '(unspecified — system default)'}`);
    con.outLine(`Calculation: ${calculation}`);
    con.outLine(`Populated cells: ${populated(workbook)}`);
    return 0;
  },
};

export const evalCommand: CommandSpec = {
  name: 'eval',
  description: 'Evaluate a formula with no file — pure engine',
  args: [{ name: 'formula', description: 'A formula to evaluate, e.g. "=ROUND(355/113, 4)"' }],
  options: [localeOpt, jsonOpt],
  run(p, con) {
    const workbook = new Workbook();
    workbook.addSheet(Defaults.firstSheetName);
    const locale = resolveLocale(p.str('--locale')) ?? workbook.settings.effectiveLocale;
    let text = p.arg('formula').trim();
    if (text.startsWith('=')) text = text.slice(1);
    const asJson = p.flag('--json');
    const parsed = Parser.tryParse(text);
    if (!parsed.ok) {
      const message = parseErrorToString(parsed.error);
      if (asJson) {
        writeReport(con, { ok: false, formula: '=' + text, error: 'parse', message });
        return 1;
      }
      throw new CliError(`Parse error: ${message}`);
    }
    workbook.settings.locale = locale;
    const context = new WorkbookEvaluationContext(workbook, workbook.activeSheet.name, locale);
    const evaluator = new Evaluator(context, FunctionRegistry.createDefault());
    const value = evaluator.evaluate(parsed.expr);
    const shown =
      value.kind === 'error' ? value.error.token : NumberFormatter.format(value, 'General', locale);
    if (asJson) {
      const typed =
        value.kind === 'number' || value.kind === 'text' || value.kind === 'boolean' ? value.value : null;
      writeReport(con, {
        ok: value.kind !== 'error',
        formula: '=' + text,
        value: typed,
        display: shown,
        error: value.kind === 'error' ? value.error.token : null,
      });
      return value.kind === 'error' ? 1 : 0;
    }
    if (value.kind === 'error') {
      con.errLine(shown);
      return 1;
    }
    con.outLine(shown);
    return 0;
  },
};

export const validate: CommandSpec = {
  name: 'validate',
  description: 'Check the workbook opens and is free of formula/#REF!/circular errors',
  args: [fileArg()],
  options: [jsonOpt],
  run(p, con) {
    const path = p.arg('file');
    const asJson = p.flag('--json');
    let loaded: LoadedWorkbook;
    try {
      loaded = load(con, path, null, null);
    } catch (e) {
      if (!(e instanceof CliError)) throw e;
      if (asJson) {
        writeReport(con, {
          ok: false,
          file: path,
          opens: false,
          message: e.message,
          errors: [],
          circular: [],
        });
        return 1;
      }
      con.errLine(e.message);
      return 1;
    }
    const result = recalculate(loaded.workbook);
    const errors = errorReports(loaded.workbook, loaded.workbook.settings.effectiveLocale);
    const circular = result.circularCells.map((c) => address(c.sheet, c.row, c.column));
    const ok = errors.length === 0 && !result.hasCircularReference;
    if (asJson) {
      writeReport(con, { ok, file: path, opens: true, errors, circular });
      return ok ? 0 : 1;
    }
    for (const problem of errors.slice(0, 50)) con.errLine(`${problem.cell}: ${problem.error}`);
    if (ok) {
      con.outLine('OK');
      return 0;
    }
    if (result.hasCircularReference) con.errLine(`${circular.length} circular reference(s).`);
    con.errLine(`${errors.length} cell error(s).`);
    return 1;
  },
};
