import { CleanUp, cleanUpNoteText, type StructuralAxis } from '@cellestra/core';
import { pickSheet } from '../infra/cells.ts';
import { EditSession } from '../infra/editSession.ts';
import { CliError } from '../infra/errors.ts';
import { resolveLocale } from '../infra/locale.ts';
import { namedRangeReport, writeReport } from '../infra/report.ts';
import { load } from '../infra/workspace.ts';
import type { CommandSpec, OptionSpec } from '../spec.ts';
import { dispatch, REMOVE } from './apply.ts';
import { editTarget, fileArg, inPlaceOpt, jsonOpt, localeOpt, outputOpt, sheetOpt } from './options.ts';

/**
 * Data verbs: sort, remove-duplicates, clean-up, split, table, data-validation, named ranges,
 * hyperlinks, insert/delete rows and columns, hide/show, merge/unmerge.
 */

const noHeaderOpt: OptionSpec = {
  name: '--no-header',
  description: 'Treat the first row as data, not a header',
  kind: 'flag',
};

export const sort: CommandSpec = {
  name: 'sort',
  description: 'Sort a range by one or more columns',
  args: [fileArg(), { name: 'range', description: 'Range to sort, e.g. Sheet1.A1:D100' }],
  options: [
    {
      name: '--by',
      description: 'Sort keys: column letters/numbers, :desc to reverse (e.g. B,A:desc)',
      required: true,
    },
    noHeaderOpt,
    inPlaceOpt,
    outputOpt,
    localeOpt,
    jsonOpt,
  ],
  run: (p, con) =>
    dispatch(
      p,
      con,
      'sort',
      { range: p.arg('range'), by: p.str('--by'), header: p.flag('--no-header') ? false : null },
      resolveLocale(p.str('--locale')),
    ),
};

export const removeDuplicates: CommandSpec = {
  name: 'remove-duplicates',
  description: 'Drop repeated rows from a range',
  args: [fileArg(), { name: 'range', description: 'Range to de-dup, e.g. Sheet1.A1:D100' }],
  options: [
    { name: '--columns', description: 'Key columns to compare (default: all), e.g. A,C' },
    noHeaderOpt,
    inPlaceOpt,
    outputOpt,
    jsonOpt,
  ],
  run: (p, con) =>
    dispatch(p, con, 'remove-duplicates', {
      range: p.arg('range'),
      columns: p.str('--columns'),
      header: p.flag('--no-header') ? false : null,
    }),
};

/**
 * The headless face of the app's Clean Up commands. Whole workbook by default — `--sheet` narrows
 * it to one. The plan comes from the same core reader the app uses, so `--dry-run` reports exactly
 * what the app would do (with the estimating text measurer: there is no display here).
 */
export const cleanUp: CommandSpec = {
  name: 'clean-up',
  description: 'Format a workbook from its contents: header, number formats, alignment, and column widths',
  args: [fileArg()],
  options: [
    sheetOpt,
    { name: '--dry-run', description: 'Report what would change, and write nothing', kind: 'flag' },
    inPlaceOpt,
    outputOpt,
    localeOpt,
    jsonOpt,
  ],
  run(p, con) {
    const path = p.arg('file');
    const preview = p.flag('--dry-run');
    // Resolve the destination before reading the file, so "refusing to overwrite" is the first
    // thing you hear — and is never said at all for a --dry-run.
    const target = preview ? null : editTarget(path, p.flag('--in-place'), p.str('--output'));
    const locale = resolveLocale(p.str('--locale'));
    const loaded = load(con, path, 'ods', locale);
    const effective = locale ?? loaded.workbook.settings.effectiveLocale;
    const name = p.str('--sheet');
    const plan =
      name === null
        ? CleanUp.forWorkbook(loaded.workbook, effective)
        : CleanUp.forSheet(loaded.workbook, pickSheet(loaded.workbook, name), effective);
    const json = p.flag('--json');
    const notes = plan.notes.map(cleanUpNoteText);

    if (preview) {
      if (json) {
        writeReport(con, { ok: true, file: path, dryRun: true, notes });
        return 0;
      }
      con.outLine(plan.isEmpty ? 'Nothing to clean up.' : notes.join('\n'));
      return 0;
    }
    if (plan.isEmpty) {
      if (json) {
        writeReport(con, { ok: true, file: path, notes, changed: [], errors: [] });
        return 0;
      }
      con.outLine('Nothing to clean up.');
      return 0;
    }
    if (!json) con.outLine(notes.join('\n'));
    return EditSession.runLoaded(con, loaded, target!, locale, json, (session) => {
      for (const edit of plan.edits) session.apply(edit);
      session.field('notes', notes);
    });
  },
};

export const split: CommandSpec = {
  name: 'split',
  description: 'Split a delimited column into multiple columns (text to columns)',
  args: [fileArg(), { name: 'column', description: 'Single-column range to split, e.g. Sheet1.A1:A100' }],
  options: [
    { name: '--delimiter', aliases: ['-d'], description: 'Delimiter to split on (default ,)', default: ',' },
    inPlaceOpt,
    outputOpt,
    jsonOpt,
  ],
  run(p, con) {
    const delimiter = p.str('--delimiter') ?? ',';
    if ([...delimiter].length !== 1) throw new CliError('--delimiter must be a single character.');
    return dispatch(p, con, 'split', { range: p.arg('column'), delimiter });
  },
};

// ---- Tables & validation (fronts on the apply ops) ---------------------------

function parseBool(value: string | null, option: string): boolean | null {
  if (value === null) return null;
  switch (value.trim().toLowerCase()) {
    case 'true':
    case 'on':
    case 'yes':
      return true;
    case 'false':
    case 'off':
    case 'no':
      return false;
    default:
      throw new CliError(`${option} must be true or false.`);
  }
}

export const table: CommandSpec = {
  name: 'table',
  description: 'Turn a range into a table (database range), or change/remove one',
  args: [
    fileArg(),
    {
      name: 'range',
      description:
        "Range to turn into a table (or an existing table's name/range for the option forms), e.g. Sheet1.A1:D20",
    },
  ],
  options: [
    { name: '--name', description: 'Table name (default: the next free Table N)' },
    { name: '--no-header', description: 'The range has no header row', kind: 'flag' },
    { name: '--replace', description: 'Redefine a table already on the range', kind: 'flag' },
    { name: '--remove', description: 'Remove the table (cells and formatting stay)', kind: 'flag' },
    { name: '--banded', description: 'true | false — stripe alternate data rows' },
    { name: '--header-column', description: 'true | false — treat the first column as row labels' },
    { name: '--resize', description: 'New range for the table, e.g. A1:D30' },
    inPlaceOpt,
    outputOpt,
    jsonOpt,
  ],
  run(p, con) {
    const range = p.arg('range');
    if (p.flag('--remove')) return dispatch(p, con, 'table-remove', { table: range });
    const banded = p.str('--banded');
    const headerColumn = p.str('--header-column');
    const resize = p.str('--resize');
    if (banded !== null || headerColumn !== null || resize !== null) {
      return dispatch(p, con, 'table-options', {
        table: range,
        bandedRows: parseBool(banded, '--banded'),
        headerColumn: parseBool(headerColumn, '--header-column'),
        range: resize,
      });
    }
    return dispatch(p, con, 'table', {
      range,
      name: p.str('--name'),
      header: p.flag('--no-header') ? false : null,
      replace: p.flag('--replace') ? true : null,
    });
  },
};

export const renameTable: CommandSpec = {
  name: 'rename-table',
  description: 'Rename a table; charts bound to it follow the new name',
  args: [
    fileArg(),
    { name: 'table', description: 'The table: its name or a cell in it' },
    { name: 'new', description: 'New table name (letters, digits and _; unique in the workbook)' },
  ],
  options: [inPlaceOpt, outputOpt, jsonOpt],
  run: (p, con) => dispatch(p, con, 'rename-table', { table: p.arg('table'), to: p.arg('new') }),
};

export const dataValidation: CommandSpec = {
  name: 'data-validation',
  description: 'Restrict a range: a list of allowed values, or a number/date/text-length rule',
  args: [fileArg(), { name: 'range', description: 'Range to validate, e.g. Sheet1.A1:A20' }],
  options: [
    { name: '--list', description: 'Allowed values, comma-separated (e.g. Red,Green,Blue)' },
    { name: '--source', description: 'A cell range whose values are the allowed list (e.g. Lists.A1:A8)' },
    { name: '--kind', description: 'whole | decimal | date | text-length (with --op and --value/--max)' },
    {
      name: '--op',
      description:
        'between | not-between | equal | not-equal | greater | less | greater-or-equal | less-or-equal (default between)',
    },
    {
      name: '--value',
      description: 'The bound (or the lower bound for between); a date may be ISO, e.g. 2026-01-31',
    },
    { name: '--max', description: 'The upper bound for between / not-between' },
    { name: '--response', description: 'stop | warning | information | none (default stop)' },
    { name: '--message', description: 'Input prompt shown while editing a cell' },
    { name: '--error', description: 'Message shown when input is rejected' },
    { name: '--remove', description: 'Remove the rule(s) on the range', kind: 'flag' },
    inPlaceOpt,
    outputOpt,
    localeOpt,
    jsonOpt,
  ],
  run: (p, con) =>
    dispatch(
      p,
      con,
      'validation',
      {
        range: p.arg('range'),
        remove: p.flag('--remove') ? true : null,
        list: p.str('--list'),
        source: p.str('--source'),
        kind: p.str('--kind'),
        operator: p.str('--op'),
        value1: p.str('--value'),
        value2: p.str('--max'),
        response: p.str('--response'),
        inputMessage: p.str('--message'),
        errorMessage: p.str('--error'),
      },
      resolveLocale(p.str('--locale')),
    ),
};

// ---- Named ranges ------------------------------------------------------------

export const name: CommandSpec = {
  name: 'name',
  description: 'Manage named ranges',
  subcommands: [
    {
      name: 'list',
      description: "List the workbook's named ranges",
      args: [fileArg()],
      options: [jsonOpt],
      run(p, con) {
        const loaded = load(con, p.arg('file'), 'ods', null);
        if (p.flag('--json')) {
          writeReport(con, loaded.workbook.namedRanges.map(namedRangeReport));
          return 0;
        }
        for (const named of loaded.workbook.namedRanges) {
          con.outLine(`${named.name}\t${named.target.toString()}\t${named.scope ?? 'workbook'}`);
        }
        return 0;
      },
    },
    {
      name: 'add',
      description: 'Define a workbook-scoped named range',
      args: [
        fileArg(),
        { name: 'name', description: 'Name to define, e.g. TaxRate' },
        { name: 'range', description: 'Range it points to, e.g. Sheet1.B1:B10' },
      ],
      options: [inPlaceOpt, outputOpt, jsonOpt],
      run: (p, con) => dispatch(p, con, 'name', { name: p.arg('name'), range: p.arg('range') }),
    },
    {
      name: 'remove',
      description: 'Remove a named range',
      args: [fileArg(), { name: 'name', description: 'Name to remove' }],
      options: [inPlaceOpt, outputOpt, jsonOpt],
      run: (p, con) => dispatch(p, con, 'unname', { name: p.arg('name') }),
    },
  ],
};

export const hyperlink: CommandSpec = {
  name: 'hyperlink',
  description: "Add, change, or remove a cell's hyperlink",
  args: [
    fileArg(),
    { name: 'cell', description: 'Cell to link, e.g. Sheet1.B2' },
    { name: 'url', description: 'Link target (omit with --remove)', arity: 'optional' },
  ],
  options: [
    { name: '--tooltip', description: 'Hover tooltip' },
    { name: '--remove', description: "Remove the cell's hyperlink", kind: 'flag' },
    inPlaceOpt,
    outputOpt,
    jsonOpt,
  ],
  run(p, con) {
    if (p.flag('--remove')) return dispatch(p, con, 'hyperlink', { cell: p.arg('cell'), url: REMOVE });
    const url = p.argOrNull('url');
    if (url === null || url.trim() === '') {
      throw new CliError('Give a url, or pass --remove to clear the link.');
    }
    return dispatch(p, con, 'hyperlink', { cell: p.arg('cell'), url, tooltip: p.str('--tooltip') });
  },
};

// ---- Insert / delete rows & columns ------------------------------------------

function structuralCommand(verb: string, description: string, axis: StructuralAxis): CommandSpec {
  const isRow = axis === 'row';
  return {
    name: verb,
    description,
    args: [
      fileArg(),
      { name: 'at', description: isRow ? 'Row number (1-based), e.g. 3' : 'Column letter or number, e.g. B' },
    ],
    options: [
      { name: '--count', description: 'How many lines (default 1)', type: 'int', default: '1' },
      sheetOpt,
      inPlaceOpt,
      outputOpt,
      jsonOpt,
    ],
    run: (p, con) =>
      dispatch(p, con, verb, { sheet: p.str('--sheet'), at: p.arg('at'), count: p.num('--count') }),
  };
}

export const insertRows = structuralCommand('insert-rows', 'Insert rows before a row', 'row');
export const deleteRows = structuralCommand('delete-rows', 'Delete rows', 'row');
export const insertCols = structuralCommand('insert-cols', 'Insert columns before a column', 'column');
export const deleteCols = structuralCommand('delete-cols', 'Delete columns', 'column');

// ---- Hide / unhide rows & columns --------------------------------------------

function visibilityCommand(verb: string, description: string, isRow: boolean, hidden: boolean): CommandSpec {
  return {
    name: verb,
    description,
    args: [
      fileArg(),
      { name: 'lines', description: isRow ? 'Row(s), e.g. 3 or 3:5' : 'Column(s), e.g. B or B:D' },
    ],
    options: [sheetOpt, inPlaceOpt, outputOpt, jsonOpt],
    run: (p, con) =>
      dispatch(p, con, hidden ? 'hide' : 'show', {
        sheet: p.str('--sheet'),
        [isRow ? 'rows' : 'columns']: p.arg('lines'),
      }),
  };
}

export const hideRows = visibilityCommand('hide-rows', 'Hide rows', true, true);
export const showRows = visibilityCommand('show-rows', 'Unhide rows', true, false);
export const hideCols = visibilityCommand('hide-cols', 'Hide columns', false, true);
export const showCols = visibilityCommand('show-cols', 'Unhide columns', false, false);

// ---- Merge / unmerge ---------------------------------------------------------

export const merge: CommandSpec = {
  name: 'merge',
  description: 'Merge a range into a single cell',
  args: [fileArg(), { name: 'range', description: 'Range to merge into one cell, e.g. Sheet1.A1:B2' }],
  options: [inPlaceOpt, outputOpt, jsonOpt],
  run: (p, con) => dispatch(p, con, 'merge', { range: p.arg('range') }),
};

export const unmerge: CommandSpec = {
  name: 'unmerge',
  description: 'Split a merged region back into individual cells',
  args: [fileArg(), { name: 'cell', description: 'Any cell of the merged region, e.g. Sheet1.A1' }],
  options: [inPlaceOpt, outputOpt, jsonOpt],
  run: (p, con) => dispatch(p, con, 'unmerge', { cell: p.arg('cell') }),
};
