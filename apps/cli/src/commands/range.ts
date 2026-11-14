import { EditSession } from '../infra/editSession.ts';
import { CliError } from '../infra/errors.ts';
import { resolveLocale } from '../infra/locale.ts';
import type { CommandSpec } from '../spec.ts';
import { dispatch, runOp } from './apply.ts';
import { editTarget, fileArg, inPlaceOpt, jsonOpt, localeOpt, outputOpt } from './options.ts';

/**
 * Block verbs — the desktop's clipboard, fill handle, format painter, and AutoFilter. Each is a front
 * on the matching `apply` op.
 */

function transfer(verb: string, description: string, cut: boolean): CommandSpec {
  return {
    name: verb,
    description,
    args: [
      fileArg(),
      { name: 'from', description: 'Source range, e.g. Sheet1.A1:B3 (a name or table works too)' },
      { name: 'to', description: 'Destination top-left cell or range, e.g. Sheet2.D1' },
    ],
    options: [
      {
        name: '--mode',
        description: 'all | values | formats | formulas (what to paste — copy only)',
        default: 'all',
      },
      inPlaceOpt,
      outputOpt,
      jsonOpt,
    ],
    run: (p, con) =>
      dispatch(p, con, verb, { from: p.arg('from'), to: p.arg('to'), mode: cut ? null : p.str('--mode') }),
  };
}

export const copy = transfer(
  'copy',
  "Copy a block to a destination (relative references shift, like the app's paste)",
  false,
);
export const move = transfer(
  'move',
  'Move a block to a destination (formulas travel unchanged; the source is cleared)',
  true,
);

export const fill: CommandSpec = {
  name: 'fill',
  description: 'Fill a range by repeating a source block (formulas shift like the fill handle)',
  args: [fileArg(), { name: 'range', description: "Range to fill, e.g. D2:D20 — the fill handle's target" }],
  options: [
    {
      name: '--from',
      description:
        "Source block to repeat (default: the range's first row, or first column when it is wider than tall)",
    },
    inPlaceOpt,
    outputOpt,
    jsonOpt,
  ],
  run: (p, con) => dispatch(p, con, 'fill', { range: p.arg('range'), from: p.str('--from') }),
};

export const paint: CommandSpec = {
  name: 'paint',
  description: 'Copy formatting from one range onto another (format painter)',
  args: [
    fileArg(),
    { name: 'from', description: 'Range whose formatting to copy, e.g. A1:D1' },
    { name: 'to', description: 'Range to format like it (the pattern tiles)' },
  ],
  options: [inPlaceOpt, outputOpt, jsonOpt],
  run: (p, con) => dispatch(p, con, 'paint', { from: p.arg('from'), to: p.arg('to') }),
};

export const filter: CommandSpec = {
  name: 'filter',
  description: "AutoFilter a range: hide the rows that don't match (never deletes data)",
  args: [fileArg(), { name: 'range', description: 'Range (or table) to filter, e.g. A1:D100' }],
  options: [
    {
      name: '--where',
      description:
        'A condition on a column: "B >5", "C *ink*", "D <>0" (SUMIF-style; repeat for more columns)',
      kind: 'multi',
    },
    {
      name: '--values',
      description: 'Allowed values for a column: "C Red,Blue" (repeat for more columns)',
      kind: 'multi',
    },
    { name: '--clear', description: 'Remove the filter: show every row again', kind: 'flag' },
    { name: '--no-header', description: 'Treat the first row as data, not a header', kind: 'flag' },
    inPlaceOpt,
    outputOpt,
    localeOpt,
    jsonOpt,
  ],
  run(p, con) {
    // Keyed case-insensitively; the first spelling of a column wins.
    const columns = new Map<string, { key: string; criteria: Record<string, unknown> }>();
    const merge = (column: string, key: string, value: unknown): void => {
      const entry = columns.get(column.toLowerCase()) ?? { key: column, criteria: {} };
      entry.criteria[key] = value;
      columns.set(column.toLowerCase(), entry);
    };
    for (const spec of p.multi('--where')) {
      const [column, rest] = splitColumn(spec, '--where');
      merge(column, 'condition', rest);
    }
    for (const spec of p.multi('--values')) {
      const [column, rest] = splitColumn(spec, '--values');
      merge(
        column,
        'values',
        rest.split(',').map((s) => s.trim()),
      );
    }
    const fields: Record<string, unknown> = {
      range: p.arg('range'),
      header: p.flag('--no-header') ? false : null,
      clear: p.flag('--clear') ? true : null,
      columns:
        columns.size > 0 ? Object.fromEntries([...columns.values()].map((e) => [e.key, e.criteria])) : null,
    };
    const path = p.arg('file');
    const target = editTarget(path, p.flag('--in-place'), p.str('--output'));
    const present = Object.fromEntries(Object.entries(fields).filter(([, v]) => v !== null));
    return EditSession.run(con, path, target, resolveLocale(p.str('--locale')), p.flag('--json'), (session) =>
      runOp(session, 'filter', present),
    );
  },
};

function splitColumn(spec: string, option: string): [column: string, criteria: string] {
  const space = spec.indexOf(' ');
  if (space <= 0 || space === spec.length - 1) {
    throw new CliError(`${option} needs "<column> <criteria>", e.g. ${option} "B >5".`);
  }
  return [spec.slice(0, space).trim(), spec.slice(space + 1).trim()];
}
