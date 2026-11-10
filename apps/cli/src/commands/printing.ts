import { spawnSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  ConditionalStyles,
  EditCommands,
  NumberFormatter,
  PAPER_SIZES,
  PageSetups,
  SetPageSetupCommand,
  type CellStyle,
  type PageSetup,
} from '@cellestra/core';
import {
  FORMAT_PRESETS,
  RULE_KINDS,
  rulesOf,
  rulesTransform,
  type FormatPresetId,
  type RuleDraft,
  type RuleKind,
} from '@cellestra/ui/conditional-formats';
import { printHtml } from '@cellestra/ui/print';
import { pickSheet, resolveRange } from '../infra/cells.ts';
import type { EditSession } from '../infra/editSession.ts';
import { CliError } from '../infra/errors.ts';
import { resolveLocale } from '../infra/locale.ts';
import { writeReport } from '../infra/report.ts';
import { atomicWrite, load, recalculate, saveText } from '../infra/workspace.ts';
import type { CommandSpec } from '../spec.ts';
import { dispatch, type Op } from './apply.ts';
import { fileArg, inPlaceOpt, jsonOpt, localeOpt, outputOpt, sheetOpt } from './options.ts';

/**
 * Conditional formats, page setup and printing: the headless faces of Format ▸ Conditional
 * Formatting, File ▸ Page Setup and File ▸ Export PDF. Rules, presets and the printed page come from
 * the same code the app uses, so a file set up here looks the same there.
 */

// ---- Conditional formats -------------------------------------------------------

/** The spellings `when` accepts for each rule kind. */
const WHEN: Record<string, RuleKind> = {
  '>': 'greater',
  gt: 'greater',
  greater: 'greater',
  '<': 'less',
  lt: 'less',
  less: 'less',
  '>=': 'greaterOrEqual',
  ge: 'greaterOrEqual',
  'at-least': 'greaterOrEqual',
  '<=': 'lessOrEqual',
  le: 'lessOrEqual',
  'at-most': 'lessOrEqual',
  '=': 'equal',
  eq: 'equal',
  equal: 'equal',
  '!=': 'notEqual',
  '<>': 'notEqual',
  ne: 'notEqual',
  'not-equal': 'notEqual',
  between: 'between',
  'not-between': 'notBetween',
  formula: 'formula',
};

/** The `when` a rule kind is reported as (the first, symbolic spelling). */
export const WHEN_OF: Record<RuleKind, string> = {
  greater: '>',
  less: '<',
  greaterOrEqual: '>=',
  lessOrEqual: '<=',
  equal: '=',
  notEqual: '!=',
  between: 'between',
  notBetween: 'not-between',
  formula: 'formula',
};

/** Preset ids in kebab case, as the CLI spells them (`light-red`). */
const kebab = (id: string) => id.replace(/[A-Z]/g, (c) => `-${c.toLowerCase()}`);
export const FORMAT_NAMES = FORMAT_PRESETS.map((p) => kebab(p.id));

function parseFormat(text: string | null): Exclude<FormatPresetId, 'keep'> {
  if (text === null) return 'lightRed';
  const t = text.trim().toLowerCase().replace(/_/g, '-');
  const preset = FORMAT_PRESETS.find((p) => kebab(p.id) === t || p.id.toLowerCase() === t);
  if (!preset) throw new CliError(`"format" must be one of: ${FORMAT_NAMES.join(', ')}.`);
  return preset.id;
}

function parseWhen(text: string): RuleKind {
  const kind = WHEN[text.trim().toLowerCase()];
  if (!kind) throw new CliError(`"when" must be one of: ${Object.keys(WHEN).join(' ')}.`);
  return kind;
}

const operandText = (v: unknown, name: string): string => {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number' || typeof v === 'boolean') return String(v).toUpperCase();
  if (typeof v === 'string') return v;
  throw new CliError(`"${name}" must be a number or text.`);
};

/** One rule from its JSON fields ({when, value, value2, format}). */
function ruleOf(fields: Record<string, unknown>, index: number): RuleDraft {
  const get = (name: string) => {
    const key = Object.keys(fields).find((k) => k.toLowerCase() === name.toLowerCase());
    return key === undefined ? undefined : fields[key];
  };
  const when = get('when');
  if (typeof when !== 'string') throw new CliError(`rule ${index + 1} needs "when".`);
  const kind = parseWhen(when);
  const format = get('format');
  if (format !== undefined && format !== null && typeof format !== 'string')
    throw new CliError(`rule ${index + 1}: "format" must be a string.`);
  return {
    kind,
    value1: operandText(get('value') ?? get('formula'), 'value'),
    value2: operandText(get('value2'), 'value2'),
    format: parseFormat((format as string | null | undefined) ?? null),
  };
}

/**
 * `conditional`: give a range conditional-format rules. `rules` replaces the range's rules (an empty
 * list removes them); with `add`, they go after the rules the range's first cell already has.
 */
export function conditionalOp(session: EditSession, op: Op): void {
  const { sheet, range } = resolveRange(session.workbook, op.require('range'), op.string('sheet'));
  const raw = op.raw('rules');
  if (raw === undefined || raw === null) throw new CliError('needs "rules" (an empty list removes them).');
  const list = Array.isArray(raw) ? raw : [raw];
  const rules = list.map((r, i) => {
    if (typeof r !== 'object' || r === null || Array.isArray(r))
      throw new CliError(`rule ${i + 1} must be an object.`);
    return ruleOf(r as Record<string, unknown>, i);
  });
  const existing = op.bool('add')
    ? rulesOf(session.workbook.styles.get(sheet.find(range.top, range.left)?.styleId ?? 0))
    : [];
  const result = rulesTransform([...existing, ...rules], {
    sheet: sheet.name,
    row: range.top,
    column: range.left,
  });
  if ('error' in result) throw new CliError(result.error);
  session.apply(EditCommands.applyStyle(session.workbook, sheet.name, range, result.transform));
  for (const [row, column] of sheet.enumerateCells())
    if (range.contains(row, column)) session.touched(sheet.name, row, column);
  session.field('rules', rules.length + existing.length);
}

/** A style's conditional rules as `conditional` takes them; null when it has none. */
export function conditionalReport(style: CellStyle): Record<string, unknown>[] | null {
  const rules = rulesOf(style);
  if (rules.length === 0) return null;
  return rules.map((r) => ({
    when: WHEN_OF[r.kind],
    // A formula rule reads back the way `value` takes it, with its = (relative to the range's top-left).
    value: r.kind === 'formula' && !r.value1.startsWith('=') ? `=${r.value1}` : r.value1,
    value2: RULE_KINDS.find((k) => k.kind === r.kind)!.operands === 2 ? r.value2 : null,
    format: r.format === 'keep' ? 'custom' : kebab(r.format),
  }));
}

export const conditional: CommandSpec = {
  name: 'conditional',
  description: 'Add conditional formatting to a range (--clear removes it)',
  args: [fileArg(), { name: 'range', description: 'Range to format, e.g. Sheet1.B2:B90' }],
  options: [
    {
      name: '--when',
      description: 'Condition: > < >= <= = != between not-between formula',
    },
    { name: '--value', description: 'Value, cell (B1), formula (=B1*2) or text; for formula: A2>B2' },
    { name: '--value2', description: 'Upper bound for between / not-between' },
    { name: '--format', description: `Format: ${FORMAT_NAMES.join(' | ')} [default: light-red]` },
    { name: '--replace', description: "Replace the range's rules instead of adding to them", kind: 'flag' },
    { name: '--clear', description: 'Remove every rule from the range', kind: 'flag' },
    sheetOpt,
    inPlaceOpt,
    outputOpt,
    jsonOpt,
  ],
  run(p, con) {
    if (p.flag('--clear'))
      return dispatch(p, con, 'conditional', { range: p.arg('range'), sheet: p.str('--sheet'), rules: [] });
    const when = p.str('--when');
    if (when === null) throw new CliError('Give --when (or --clear).');
    return dispatch(p, con, 'conditional', {
      range: p.arg('range'),
      sheet: p.str('--sheet'),
      add: !p.flag('--replace'),
      rules: [{ when, value: p.str('--value'), value2: p.str('--value2'), format: p.str('--format') }],
    });
  },
};

// ---- Page setup ----------------------------------------------------------------

const PAPER_NAMES = PAPER_SIZES.map((p) => p.id);

function paperSize(text: string): { width: number; height: number } {
  const t = text.trim().toLowerCase();
  const preset = PAPER_SIZES.find((p) => p.id === t);
  if (preset) return preset;
  const custom = /^(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)$/.exec(t);
  if (custom) return { width: Number(custom[1]), height: Number(custom[2]) };
  throw new CliError(`"paper" must be one of ${PAPER_NAMES.join(', ')}, or WIDTHxHEIGHT in mm (210x297).`);
}

function count(op: Op, name: string, min: number): number | null {
  const n = op.number(name);
  if (n === null) return null;
  if (!Number.isInteger(n) || n < min || n > 999)
    throw new CliError(`"${name}" must be a whole number from ${min} to 999.`);
  return n;
}

/**
 * `page-setup`: how a sheet prints. Fields left out keep the sheet's current setup (or the default,
 * A4 portrait with 20 mm margins at 100%); `reset` goes back to the default.
 */
export function pageSetupOp(session: EditSession, op: Op): void {
  const sheet = pickSheet(session.workbook, op.string('sheet'));
  if (op.bool('reset')) {
    session.apply(new SetPageSetupCommand(sheet.name, null));
    return;
  }
  let s: PageSetup = sheet.pageSetup ?? PageSetups.default;
  const paper = op.string('paper');
  if (paper !== null) {
    const size = paperSize(paper);
    s = PageSetups.orient({ ...s, paperWidth: size.width, paperHeight: size.height }, s.orientation);
  }
  const orientation = op.string('orientation');
  if (orientation !== null) {
    const o = orientation.trim().toLowerCase();
    if (o !== 'portrait' && o !== 'landscape')
      throw new CliError('"orientation" must be portrait or landscape.');
    s = PageSetups.orient(s, o);
  }
  if (op.present('margins')) {
    const raw = op.raw('margins');
    const mm = (v: unknown, side: string): number => {
      const n = typeof v === 'string' ? Number(v) : v;
      if (typeof n !== 'number' || !Number.isFinite(n) || n < 0 || n > 200)
        throw new CliError(`the ${side} margin must be 0–200 mm.`);
      return n;
    };
    if (typeof raw === 'object' && raw !== null && !Array.isArray(raw)) {
      const m = raw as Record<string, unknown>;
      s = {
        ...s,
        margins: {
          top: m['top'] === undefined ? s.margins.top : mm(m['top'], 'top'),
          bottom: m['bottom'] === undefined ? s.margins.bottom : mm(m['bottom'], 'bottom'),
          left: m['left'] === undefined ? s.margins.left : mm(m['left'], 'left'),
          right: m['right'] === undefined ? s.margins.right : mm(m['right'], 'right'),
        },
      };
    } else {
      const all = mm(raw, 'every');
      s = { ...s, margins: { top: all, bottom: all, left: all, right: all } };
    }
  }
  if (
    s.paperWidth - s.margins.left - s.margins.right < 20 ||
    s.paperHeight - s.margins.top - s.margins.bottom < 20
  )
    throw new CliError('the margins leave too little of the page to print on.');

  const percent = op.number('scale');
  const fitWidth = count(op, 'fitWidth', 0);
  const fitHeight = count(op, 'fitHeight', 0);
  const pages = count(op, 'pages', 1);
  const scales = [percent !== null, fitWidth !== null || fitHeight !== null, pages !== null].filter(Boolean);
  if (scales.length > 1) throw new CliError('give one of "scale", "fitWidth"/"fitHeight", or "pages".');
  if (percent !== null) {
    if (percent < 10 || percent > 400) throw new CliError('"scale" must be 10–400 (percent).');
    s = { ...s, scale: { kind: 'percent', percent } };
  } else if (fitWidth !== null || fitHeight !== null) {
    const width = fitWidth ?? 0;
    const height = fitHeight ?? 0;
    if (width === 0 && height === 0) throw new CliError('"fitWidth" and "fitHeight" can\'t both be 0.');
    s = { ...s, scale: { kind: 'fit', width, height } };
  } else if (pages !== null) s = { ...s, scale: { kind: 'pages', pages } };

  const gridlines = op.bool('gridlines');
  if (gridlines !== null) s = { ...s, gridlines };
  const headers = op.bool('headers');
  if (headers !== null) s = { ...s, headers };
  const center = op.string('center');
  if (center !== null) {
    const c = center.trim().toLowerCase();
    if (!['none', 'horizontal', 'vertical', 'both'].includes(c))
      throw new CliError('"center" must be none, horizontal, vertical, or both.');
    s = {
      ...s,
      centerHorizontally: c === 'horizontal' || c === 'both',
      centerVertically: c === 'vertical' || c === 'both',
    };
  }
  const order = op.string('order');
  if (order !== null) {
    const o = order.trim().toLowerCase();
    if (o === 'down' || o === 'down-then-over') s = { ...s, pageOrder: 'downThenOver' };
    else if (o === 'over' || o === 'over-then-down') s = { ...s, pageOrder: 'overThenDown' };
    else throw new CliError('"order" must be down-then-over or over-then-down.');
  }
  session.apply(new SetPageSetupCommand(sheet.name, s));
}

/** A sheet's page setup as `page-setup` takes it; null when it has none of its own or it is the default. */
export function pageSetupReport(setup: PageSetup | null): Record<string, unknown> | null {
  if (setup === null || PageSetups.equal(setup, PageSetups.default)) return null;
  const m = setup.margins;
  const uniform = m.top === m.bottom && m.bottom === m.left && m.left === m.right;
  const paper = PageSetups.paperOf(setup);
  const short = Math.min(setup.paperWidth, setup.paperHeight);
  const long = Math.max(setup.paperWidth, setup.paperHeight);
  const sc = setup.scale;
  return {
    paper: paper?.id ?? `${short}x${long}`,
    orientation: setup.orientation,
    margins: uniform ? m.top : { top: m.top, bottom: m.bottom, left: m.left, right: m.right },
    scale: sc.kind === 'percent' ? sc.percent : null,
    fitWidth: sc.kind === 'fit' ? sc.width : null,
    fitHeight: sc.kind === 'fit' ? sc.height : null,
    pages: sc.kind === 'pages' ? sc.pages : null,
    gridlines: setup.gridlines,
    headers: setup.headers,
    center:
      setup.centerHorizontally && setup.centerVertically
        ? 'both'
        : setup.centerHorizontally
          ? 'horizontal'
          : setup.centerVertically
            ? 'vertical'
            : 'none',
    order: setup.pageOrder === 'downThenOver' ? 'down-then-over' : 'over-then-down',
  };
}

export const pageSetup: CommandSpec = {
  name: 'page-setup',
  description: 'Set how a sheet prints: paper, orientation, margins, scale or fit, grid lines, headers',
  args: [fileArg()],
  options: [
    { name: '--paper', description: `${PAPER_NAMES.join(' | ')} | WIDTHxHEIGHT (mm)` },
    { name: '--orientation', description: 'portrait | landscape' },
    { name: '--margins', description: 'All four margins in mm', type: 'number' },
    { name: '--margin-top', description: 'Top margin in mm', type: 'number' },
    { name: '--margin-bottom', description: 'Bottom margin in mm', type: 'number' },
    { name: '--margin-left', description: 'Left margin in mm', type: 'number' },
    { name: '--margin-right', description: 'Right margin in mm', type: 'number' },
    { name: '--scale', description: 'Zoom in percent, 10–400', type: 'number' },
    { name: '--fit-width', description: 'Fit to this many pages across (0 = any)', type: 'int' },
    { name: '--fit-height', description: 'Fit to this many pages down (0 = any)', type: 'int' },
    { name: '--pages', description: 'Fit on at most this many pages', type: 'int' },
    {
      name: '--gridlines',
      description: 'print grid lines (`--gridlines false` turns them off)',
      kind: 'optionalValue',
    },
    {
      name: '--headers',
      description: 'print row numbers and column letters (`--headers false` turns them off)',
      kind: 'optionalValue',
    },
    { name: '--center', description: 'none | horizontal | vertical | both' },
    { name: '--order', description: 'down-then-over | over-then-down' },
    { name: '--reset', description: 'Back to the default page (A4 portrait, 20 mm, 100%)', kind: 'flag' },
    sheetOpt,
    inPlaceOpt,
    outputOpt,
    jsonOpt,
  ],
  run(p, con) {
    const side = ['top', 'bottom', 'left', 'right'] as const;
    const sides = Object.fromEntries(
      side.map((k) => [k, p.num(`--margin-${k}`)]).filter(([, v]) => v !== null),
    );
    const all = p.num('--margins');
    return dispatch(p, con, 'page-setup', {
      sheet: p.str('--sheet'),
      reset: p.flag('--reset') || null,
      paper: p.str('--paper'),
      orientation: p.str('--orientation'),
      margins:
        Object.keys(sides).length > 0
          ? { ...(all !== null ? { top: all, bottom: all, left: all, right: all } : {}), ...sides }
          : all,
      scale: p.num('--scale'),
      fitWidth: p.num('--fit-width'),
      fitHeight: p.num('--fit-height'),
      pages: p.num('--pages'),
      gridlines: p.onOff('--gridlines'),
      headers: p.onOff('--headers'),
      center: p.str('--center'),
      order: p.str('--order'),
    });
  },
};

// ---- print ---------------------------------------------------------------------

/** Chromium-family browsers that can print HTML to PDF headlessly, in the order they are tried. */
function browserCandidates(): string[] {
  const env = process.env['CELLESTRA_BROWSER'];
  if (env) return [env];
  switch (process.platform) {
    case 'darwin':
      return ['Google Chrome', 'Chromium', 'Microsoft Edge', 'Brave Browser', 'Vivaldi'].map(
        (app) => `/Applications/${app}.app/Contents/MacOS/${app}`,
      );
    case 'win32': {
      const roots = [
        process.env['PROGRAMFILES'],
        process.env['PROGRAMFILES(X86)'],
        process.env['LOCALAPPDATA'],
      ];
      const rel = [
        'Google\\Chrome\\Application\\chrome.exe',
        'Microsoft\\Edge\\Application\\msedge.exe',
        'BraveSoftware\\Brave-Browser\\Application\\brave.exe',
      ];
      return roots.flatMap((r) => (r ? rel.map((x) => join(r, x)) : []));
    }
    default:
      return [
        'google-chrome',
        'google-chrome-stable',
        'chromium',
        'chromium-browser',
        'microsoft-edge',
        'brave-browser',
      ].flatMap((name) => ['/usr/bin', '/usr/local/bin', '/snap/bin'].map((d) => join(d, name)));
  }
}

function findBrowser(explicit: string | null): string {
  const candidates = explicit ? [explicit] : browserCandidates();
  const found = candidates.find((c) => existsSync(c));
  if (!found)
    throw new CliError(
      explicit
        ? `No browser at ${explicit}.`
        : 'Making a PDF needs Chrome, Chromium, Edge or Brave. Set CELLESTRA_BROWSER (or --browser) to one, or write .html instead.',
    );
  return found;
}

/** Print an HTML document to PDF with a headless Chromium (its own throwaway profile). */
function htmlToPdf(browser: string, html: string, output: string): void {
  const dir = mkdtempSync(join(tmpdir(), 'cellestra-print-'));
  try {
    const page = join(dir, 'page.html');
    const pdf = join(dir, 'out.pdf');
    writeFileSync(page, html);
    // The printer is async (DevTools), so it runs as its own Node process; see pdfPrinter.ts.
    const printer = fileURLToPath(new URL('../infra/pdfPrinter.ts', import.meta.url));
    const run = spawnSync(process.execPath, [printer, browser, page, pdf, join(dir, 'profile')], {
      encoding: 'utf8',
      timeout: 120_000,
    });
    if (!existsSync(pdf))
      throw new CliError(
        (
          run.stderr?.trim() || `The browser didn't make a PDF${run.error ? `: ${run.error.message}` : '.'}`
        ).trim(),
      );
    atomicWrite(output, readFileSync(pdf));
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

export const print: CommandSpec = {
  name: 'print',
  description: 'Print a sheet to PDF (or printable HTML) using its page setup',
  args: [
    fileArg('file', 'Source .ods workbook'),
    fileArg('output', 'Destination .pdf, or .html (- for HTML on stdout)'),
  ],
  options: [
    sheetOpt,
    { name: '--range', description: 'Print only this range, e.g. A1:F40 (default: every used cell)' },
    {
      name: '--browser',
      description:
        'Chromium-family browser to make the PDF with (default: $CELLESTRA_BROWSER, then Chrome, Edge, Brave…)',
    },
    localeOpt,
    jsonOpt,
  ],
  run(p, con) {
    const output = p.arg('output');
    const asPdf = output !== '-' && output.toLowerCase().endsWith('.pdf');
    if (!asPdf && output !== '-' && !/\.html?$/i.test(output))
      throw new CliError('The output must end in .pdf or .html (or be - for HTML on stdout).');
    const loaded = load(con, p.arg('file'), 'ods', null);
    const workbook = loaded.workbook;
    recalculate(workbook);
    const rangeText = p.str('--range');
    const resolved = rangeText !== null ? resolveRange(workbook, rangeText, p.str('--sheet')) : null;
    const sheet = resolved?.sheet ?? pickSheet(workbook, p.str('--sheet'));
    if (sheet.isChartSheet) throw new CliError(`'${sheet.name}' is a chart sheet; print a sheet with cells.`);
    const locale = resolveLocale(p.str('--locale')) ?? workbook.settings.effectiveLocale;
    const conditional = new ConditionalStyles(workbook);
    const r = resolved?.range;
    const title =
      p.arg('file') === '-'
        ? sheet.name
        : `${resolve(p.arg('file'))
            .split(/[\\/]/)
            .pop()!
            .replace(/\.[^.]+$/, '')} — ${sheet.name}`;
    const doc = printHtml({
      sheet,
      styles: workbook.styles,
      setup: sheet.pageSetup,
      range: r ? { top: r.top, left: r.left, bottom: r.bottom, right: r.right } : null,
      display: (value, style) => NumberFormatter.format(value, style.numberFormat.code, locale),
      conditional: (style, row, column) => conditional.resolve(style, sheet.name, row, column),
      title,
    });
    if (asPdf) htmlToPdf(findBrowser(p.str('--browser')), doc.html, output);
    else saveText(con, doc.html, output);
    if (p.flag('--json'))
      writeReport(con, { output, sheet: sheet.name, pages: doc.pages, format: asPdf ? 'pdf' : 'html' });
    else if (output !== '-')
      con.errLine(`Wrote ${output}: ${sheet.name}, ${doc.pages} page${doc.pages === 1 ? '' : 's'}.`);
    return 0;
  },
};
