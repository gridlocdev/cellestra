import type { CommandSpec, OptionSpec } from '../spec.ts';
import { dispatch } from './apply.ts';
import { fileArg, inPlaceOpt, jsonOpt, outputOpt, sheetOpt } from './options.ts';

/**
 * The formatting and layout verbs a person at a shell reaches for — `style`, `width`, `height`,
 * `freeze`, `clear`, `settings`, and the sheet-arranging pair — each a thin front on the matching
 * `apply` op, so a verb and its op can never disagree.
 */

/** A tri-state switch: `--bold` sets, `--bold false` (or off/no) clears, absent leaves alone. */
export const onOff = (name: string): OptionSpec => ({
  name,
  description: `${name.slice(2)} on (\`${name} false\` turns it off)`,
  kind: 'optionalValue',
});

const editOptions = [inPlaceOpt, outputOpt, jsonOpt] as const;

export const style: CommandSpec = {
  name: 'style',
  description: 'Format a range: font, fill, alignment, number format, border',
  args: [
    fileArg(),
    { name: 'range', description: 'Range (or cell, name, table) to format, e.g. Sheet1.A1:D1' },
  ],
  options: [
    onOff('--bold'),
    onOff('--italic'),
    onOff('--underline'),
    onOff('--strikethrough'),
    onOff('--wrap'),
    {
      name: '--checkbox',
      description:
        'show TRUE/FALSE as a checkbox; empty cells become FALSE (`--checkbox false` turns it off)',
      kind: 'optionalValue',
    },
    { name: '--font', description: 'Font family' },
    { name: '--size', description: 'Font size in points', type: 'number' },
    { name: '--color', description: 'Text color, #RRGGBB' },
    { name: '--fill', description: 'Background, #RRGGBB (or none)' },
    { name: '--align', description: 'general | left | center | right' },
    { name: '--valign', description: 'top | middle | bottom' },
    { name: '--number-format', description: 'Number format code, e.g. "#,##0.00" or General' },
    {
      name: '--border',
      description: 'none | thin | medium | thick | dashed | dotted | double (all four edges)',
    },
    { name: '--border-color', description: 'Border color, #RRGGBB' },
    ...editOptions,
  ],
  run: (p, con) =>
    dispatch(p, con, 'style', {
      range: p.arg('range'),
      bold: p.onOff('--bold'),
      italic: p.onOff('--italic'),
      underline: p.onOff('--underline'),
      strikethrough: p.onOff('--strikethrough'),
      wrap: p.onOff('--wrap'),
      checkbox: p.onOff('--checkbox'),
      font: p.str('--font'),
      size: p.num('--size'),
      color: p.str('--color'),
      fill: p.str('--fill'),
      align: p.str('--align'),
      valign: p.str('--valign'),
      format: p.str('--number-format'),
      border: p.str('--border'),
      borderColor: p.str('--border-color'),
    }),
};

export const width: CommandSpec = {
  name: 'width',
  description: 'Set column width(s)',
  args: [
    fileArg(),
    { name: 'columns', description: 'Column or span, e.g. B or B:D' },
    { name: 'width', description: 'Width in points', type: 'number' },
  ],
  options: [sheetOpt, ...editOptions],
  run: (p, con) =>
    dispatch(p, con, 'width', {
      columns: p.arg('columns'),
      width: p.argNumber('width'),
      sheet: p.str('--sheet'),
    }),
};

export const height: CommandSpec = {
  name: 'height',
  description: 'Set row height(s)',
  args: [
    fileArg(),
    { name: 'rows', description: 'Row or span, e.g. 1 or 2:5' },
    { name: 'height', description: 'Height in points', type: 'number' },
  ],
  options: [sheetOpt, ...editOptions],
  run: (p, con) =>
    dispatch(p, con, 'height', {
      rows: p.arg('rows'),
      height: p.argNumber('height'),
      sheet: p.str('--sheet'),
    }),
};

export const freeze: CommandSpec = {
  name: 'freeze',
  description: 'Freeze panes (--rows 0 --columns 0 unfreezes)',
  args: [fileArg()],
  options: [
    { name: '--rows', description: 'Rows to freeze at the top (0 = none)', type: 'int' },
    { name: '--columns', description: 'Columns to freeze at the left (0 = none)', type: 'int' },
    sheetOpt,
    ...editOptions,
  ],
  run: (p, con) =>
    dispatch(p, con, 'freeze', {
      rows: p.num('--rows') ?? 0,
      columns: p.num('--columns') ?? 0,
      sheet: p.str('--sheet'),
    }),
};

export const clear: CommandSpec = {
  name: 'clear',
  description: "Clear a range's contents, formats, or everything",
  args: [fileArg(), { name: 'range', description: 'Range to clear, e.g. Sheet1.A2:D20' }],
  options: [{ name: '--what', description: 'contents | formats | all', default: 'contents' }, ...editOptions],
  run: (p, con) => dispatch(p, con, 'clear', { range: p.arg('range'), what: p.str('--what') }),
};

export const settings: CommandSpec = {
  name: 'settings',
  description: 'Set document settings: locale, calculation mode, active sheet',
  args: [fileArg()],
  options: [
    { name: '--locale', description: 'Document locale, e.g. en-US (parsing and display)' },
    { name: '--calc', description: 'automatic | manual' },
    { name: '--active-sheet', description: 'The sheet the app opens on' },
    ...editOptions,
  ],
  run: (p, con) =>
    dispatch(p, con, 'settings', {
      locale: p.str('--locale'),
      calc: p.str('--calc'),
      activeSheet: p.str('--active-sheet'),
    }),
};

export const duplicateSheet: CommandSpec = {
  name: 'duplicate-sheet',
  description: 'Copy a worksheet, contents and formatting',
  args: [fileArg(), { name: 'name', description: 'Sheet to copy' }],
  options: [
    { name: '--as', description: "Name for the copy (default: the app's next free name)" },
    ...editOptions,
  ],
  run: (p, con) => dispatch(p, con, 'duplicate-sheet', { name: p.arg('name'), as: p.str('--as') }),
};

export const moveSheet: CommandSpec = {
  name: 'move-sheet',
  description: 'Reorder a worksheet',
  args: [
    fileArg(),
    { name: 'name', description: 'Sheet to move' },
    { name: 'to', description: 'New 1-based position', type: 'int' },
  ],
  options: [...editOptions],
  run: (p, con) => dispatch(p, con, 'move-sheet', { name: p.arg('name'), to: p.argNumber('to') }),
};
