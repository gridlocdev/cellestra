import { Defaults } from '../defaults.ts';
import { CellSnapshotCommand } from '../editing/cellSnapshotCommand.ts';
import type { EditCommand, Position } from '../editing/editCommand.ts';
import { EditCommands } from '../editing/editCommands.ts';
import { InputParser } from '../editing/inputParser.ts';
import { localeInfo } from '../formulas/evaluation/localeInfo.ts';
import { parseNumberAny } from '../formulas/evaluation/valueParsing.ts';
import type { CellRange } from '../model/cellRange.ts';
import { cellKey, keyColumn, keyRow } from '../model/sheetDimensions.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';
import { Fills, NumberFormats, styleKey, type CellStyle, type NumberFormat } from '../styling/cellStyle.ts';
import type { NumberFormatKind } from '../styling/numberFormatParser.ts';
import { NumberFormatter } from '../styling/numberFormatter.ts';
import { approximateTextMeasurer, type TextMeasurer } from '../styling/textMeasurer.ts';
import { CellValue } from '../values/cellValue.ts';
import { DataRegion } from './dataRegion.ts';

/** One line of a Clean Up report: what was done on which sheet. */
export interface CleanUpNote {
  readonly sheet: string;
  readonly description: string;
}

export function cleanUpNoteText(note: CleanUpNote): string {
  return `${note.sheet}: ${note.description}`;
}

/** What Clean Up would do — computed without touching the workbook; `toCommand` makes it one undo step. */
export class CleanUpPlan {
  readonly notes: readonly CleanUpNote[];
  readonly edits: readonly EditCommand[];

  constructor(notes: readonly CleanUpNote[], edits: readonly EditCommand[]) {
    this.notes = notes;
    this.edits = edits;
  }

  get isEmpty(): boolean {
    return this.edits.length === 0;
  }

  get summary(): string {
    return this.notes.length === 0 ? 'Nothing to clean up.' : this.notes.map(cleanUpNoteText).join('; ');
  }

  toCommand(label: string): EditCommand | null {
    return this.edits.length === 0 ? null : EditCommands.composite(label, this.edits);
  }
}

// Matches the grid's own auto-fit so Clean Up and a border double-click land on the same width.
const WIDTH_PADDING = 10;
const MIN_WIDTH = 24;
const MAX_WIDTH = 600;
const THOUSANDS_THRESHOLD = 1000;
const CONVENTIONAL_DECIMALS = 2;
const MAX_INFERRED_DECIMALS = 4;
const MAX_SAFE_DIGITS = 15;
const BARE_CURRENCY = /^[$€£¥¤]+$/;
const ANY_CURRENCY = /[$€£¥¤]/;
// "+15551234567", "+44 7700 900123": nobody types a plain amount with a leading plus.
const PHONE_LIKE = /^\+[0-9][0-9 ()-]*$/;
// "12E3", "4e10": part numbers and lot codes far more often than scientific notation typed as text.
const EXPONENT_LIKE = /^[0-9]+(\.[0-9]+)?[eE][+-]?[0-9]+$/;

/**
 * Clean Up (PRD §8.17): per data block, style and freeze a header row, coerce numbers/dates typed as
 * text when a column is uniformly numeric, trim whitespace, give each column one inferred number
 * format, reset body alignment, and fit widths. It never rewrites formulas, overrides a chosen
 * number format, or coerces identifiers (leading zeros, more than 15 digits, "+" phone numbers, "12E3" codes).
 */
export const CleanUp = {
  sheetLabel: 'Clean Up Sheet',
  workbookLabel: 'Clean Up Workbook',

  forSheet: (workbook: Workbook, sheet: Worksheet, locale?: string, measurer?: TextMeasurer): CleanUpPlan =>
    CleanUp.plan(workbook, [sheet], locale, measurer),

  forWorkbook: (workbook: Workbook, locale?: string, measurer?: TextMeasurer): CleanUpPlan =>
    CleanUp.plan(workbook, workbook.worksheets, locale, measurer),

  plan(
    workbook: Workbook,
    sheets: Iterable<Worksheet>,
    locale?: string,
    measurer?: TextMeasurer,
  ): CleanUpPlan {
    const loc = locale ?? workbook.settings.effectiveLocale;
    const measure = measurer ?? approximateTextMeasurer;
    const notes: CleanUpNote[] = [];
    const edits: EditCommand[] = [];
    // A chart sheet carries a chart, not a grid.
    for (const sheet of sheets)
      if (!sheet.isChartSheet) planSheet(workbook, sheet, loc, measure, notes, edits);
    return new CleanUpPlan(notes, edits);
  },

  /** The sheet's data blocks: current regions of populated cells, top-to-bottom, lone cells excluded. */
  blocks(sheet: Worksheet): CellRange[] {
    const claimed = new Set<number>();
    const blocks: CellRange[] = [];
    const populated = [...sheet.enumerateCells()]
      .filter(([, , cell]) => cell.value.kind !== 'empty')
      .map(([row, column]) => cellKey(row, column))
      .sort((a, b) => a - b);
    for (const key of populated) {
      if (claimed.has(key)) continue;
      claimed.add(key);
      const region = DataRegion.currentRegion(sheet, keyRow(key), keyColumn(key));
      for (const [r, c] of region.positions()) claimed.add(cellKey(r, c));
      if (region.rowCount * region.columnCount > 1) blocks.push(region);
    }
    return blocks;
  },
} as const;

type CurrentStyle = (row: number, column: number) => CellStyle;
type Restyle = (row: number, column: number, style: CellStyle) => boolean;

function planSheet(
  workbook: Workbook,
  sheet: Worksheet,
  locale: string,
  measurer: TextMeasurer,
  notes: CleanUpNote[],
  edits: EditCommand[],
): void {
  const blocks = CleanUp.blocks(sheet);
  if (blocks.length === 0) return;
  const values = new Map<number, CellValue>();
  const styles = new Map<number, CellStyle>();
  const widths = new Map<number, number>();
  let tidied = 0;
  let formattedColumns = 0;
  let headerRows = 0;
  let realigned = 0;
  let freezeRows: number | null = null;

  const current: CurrentStyle = (row, column) =>
    styles.get(cellKey(row, column)) ?? workbook.styles.get(sheet.effectiveStyleId(row, column));
  // Reports whether this was a change, so a second run has nothing to say.
  const restyle: Restyle = (row, column, style) => {
    if (styleKey(current(row, column)) === styleKey(style)) return false;
    styles.set(cellKey(row, column), style);
    return true;
  };

  for (const block of blocks) {
    const hasHeader = blockHasHeader(sheet, block, locale);
    if (hasHeader) {
      if (styleHeader(block, current, restyle)) headerRows++;
      // Only the topmost block's header is what the frozen pane holds still.
      freezeRows ??= block.top + 1;
    }
    const bodyTop = hasHeader ? block.top + 1 : block.top;
    for (let column = block.left; column <= block.right; column++) {
      const scan = scanColumn(sheet, column, bodyTop, block.bottom, locale);
      tidied += applyColumnTidy(column, scan, values);
      if (applyColumnFormat(column, bodyTop, block.bottom, scan, locale, values, sheet, current, restyle)) {
        formattedColumns++;
      }
      realigned += alignBody(column, bodyTop, block.bottom, sheet, current, restyle);
    }
    fitColumns(sheet, block, locale, measurer, values, current, widths);
  }

  const freezeTo: number | null = freezeRows;
  let froze = false;
  if (freezeTo !== null && (sheet.frozenRows !== freezeTo || sheet.frozenColumns !== 0)) {
    edits.push(EditCommands.setFrozenPanes(sheet.name, freezeTo, 0));
    froze = true;
  }

  const touched: Position[] = [...new Set([...values.keys(), ...styles.keys()])]
    .sort((a, b) => a - b)
    .map((k) => [keyRow(k), keyColumn(k)]);
  if (touched.length > 0) {
    edits.unshift(
      new CellSnapshotCommand('Clean Up Cells', sheet.name, touched, (ws) => {
        for (const [key, value] of values) ws.setValue(keyRow(key), keyColumn(key), value);
        for (const [key, style] of styles) {
          const row = keyRow(key);
          const column = keyColumn(key);
          const cell = ws.getOrCreate(row, column);
          cell.styleId = workbook.styles.intern(style);
          ws.pruneIfBlank(row, column, cell);
        }
      }),
    );
  }
  for (const [column, width] of [...widths].sort((a, b) => a[0] - b[0])) {
    edits.push(EditCommands.setColumnWidth(sheet.name, column, width));
  }

  const parts: string[] = [];
  addPart(parts, tidied, 'cell', 'tidied');
  addPart(parts, formattedColumns, 'column', 'formatted');
  addPart(parts, realigned, 'cell', 'realigned');
  addPart(parts, headerRows, 'header row', 'styled');
  if (froze) parts.push('froze the header row');
  addPart(parts, widths.size, 'column', 'resized');
  if (parts.length > 0) notes.push({ sheet: sheet.name, description: parts.join(', ') });
}

function addPart(parts: string[], count: number, noun: string, verb: string): void {
  if (count > 0) parts.push(count === 1 ? `${verb} 1 ${noun}` : `${verb} ${count} ${noun}s`);
}

// ---- Header -----------------------------------------------------------------

function blockHasHeader(sheet: Worksheet, block: CellRange, locale: string): boolean {
  if (block.rowCount < 2) return false;
  let headerAllLabels = true;
  let bodyHasNumbers = false;
  for (let column = block.left; column <= block.right; column++) {
    const header = sheet.getValue(block.top, column);
    if (header.kind !== 'empty' && (header.kind !== 'text' || readsAsNumber(header.value, locale))) {
      headerAllLabels = false;
    }
    for (let row = block.top + 1; row <= block.bottom; row++) {
      const v = sheet.getValue(row, column);
      if (v.kind === 'empty') continue;
      bodyHasNumbers ||= v.kind !== 'text' || readsAsNumber(v.value, locale);
    }
  }
  return headerAllLabels && bodyHasNumbers;
}

function readsAsNumber(text: string, locale: string): boolean {
  const trimmed = text.trim();
  return !isIdentifier(trimmed) && tryCoerce(trimmed, locale) !== null;
}

function styleHeader(block: CellRange, current: CurrentStyle, restyle: Restyle): boolean {
  let changed = false;
  for (let column = block.left; column <= block.right; column++) {
    const s = current(block.top, column);
    const style: CellStyle = {
      ...s,
      font: { ...s.font, bold: true, color: Defaults.tableHeaderTextArgb },
      fill: Fills.solid(Defaults.tableHeaderFillArgb),
      alignment: { ...s.alignment, horizontal: 'center' },
    };
    changed = restyle(block.top, column, style) || changed;
  }
  return changed;
}

// ---- Column analysis ----------------------------------------------------------

interface Coercion {
  readonly value: CellValue;
  readonly dateFormat: NumberFormat | null;
  readonly decimals: number;
  readonly kind: NumberFormatKind;
}

class ColumnScan {
  readonly trims = new Map<number, string>();
  readonly coercions = new Map<number, Coercion>();
  allNumeric = true;
  blocked = false;
  numericCount = 0;
  dateCount = 0;
  percentCount = 0;
  currencyCount = 0;
  allIntegers = true;
  maxMagnitude = 0;
  decimals = 0;
  dateCode: string | null = null;

  /** Dates and plain numbers never mix: one column takes one format. */
  get isUniformlyNumeric(): boolean {
    return (
      this.allNumeric &&
      !this.blocked &&
      this.numericCount > 0 &&
      (this.dateCount === 0 || this.dateCount === this.numericCount)
    );
  }
}

function scanColumn(
  sheet: Worksheet,
  column: number,
  top: number,
  bottom: number,
  locale: string,
): ColumnScan {
  const scan = new ColumnScan();
  for (let row = top; row <= bottom; row++) {
    const cell = sheet.find(row, column);
    if (!cell || cell.value.kind === 'empty') continue;
    // A formula's result votes on the format, but the formula itself is never rewritten.
    if (cell.formula !== null || cell.value.kind !== 'text') {
      contribute(scan, cell.value, null);
      continue;
    }
    const text = cell.value.value;
    const trimmed = text.trim();
    if (trimmed.length !== text.length) scan.trims.set(row, trimmed);
    if (trimmed.length === 0) continue; // the trim clears it; an empty cell votes on nothing
    if (isIdentifier(trimmed)) {
      scan.blocked = true;
      scan.allNumeric = false;
      continue;
    }
    const coercion = tryCoerce(trimmed, locale);
    if (coercion) {
      scan.coercions.set(row, coercion);
      contribute(scan, coercion.value, coercion);
    } else {
      scan.allNumeric = false;
    }
  }
  return scan;
}

function contribute(scan: ColumnScan, value: CellValue, coercion: Coercion | null): void {
  if (value.kind !== 'number') {
    scan.allNumeric = false;
    return;
  }
  scan.numericCount++;
  const n = value.value;
  scan.maxMagnitude = Math.max(scan.maxMagnitude, Math.abs(n));
  if (n !== Math.floor(n)) scan.allIntegers = false;
  if (coercion) {
    // Decimals from what was typed: "45.5%" is one place on screen though the value is 0.455.
    scan.decimals = Math.max(scan.decimals, coercion.decimals);
    if (coercion.kind === 'date') {
      scan.dateCount++;
      scan.dateCode ??= coercion.dateFormat?.code ?? null;
    } else if (coercion.kind === 'percent') scan.percentCount++;
    else if (coercion.kind === 'currency') scan.currencyCount++;
    return;
  }
  scan.decimals = Math.max(scan.decimals, numberFractionDigits(n));
}

function applyColumnTidy(column: number, scan: ColumnScan, values: Map<number, CellValue>): number {
  const coerce = scan.isUniformlyNumeric && scan.coercions.size > 0;
  let tidied = 0;
  if (coerce) {
    for (const [row, c] of scan.coercions) {
      values.set(cellKey(row, column), c.value);
      tidied++;
    }
  }
  for (const [row, trimmed] of scan.trims) {
    if (coerce && scan.coercions.has(row)) continue; // the coercion already replaced it
    values.set(cellKey(row, column), trimmed.length === 0 ? CellValue.empty : CellValue.text(trimmed));
    tidied++;
  }
  return tidied;
}

function applyColumnFormat(
  column: number,
  top: number,
  bottom: number,
  scan: ColumnScan,
  locale: string,
  values: ReadonlyMap<number, CellValue>,
  sheet: Worksheet,
  current: CurrentStyle,
  restyle: Restyle,
): boolean {
  const format = inferFormat(scan, locale);
  if (!format) return false;
  let applied = false;
  for (let row = top; row <= bottom; row++) {
    const planned = values.get(cellKey(row, column));
    const cell = sheet.find(row, column);
    const willHoldValue =
      planned !== undefined ? planned.kind !== 'empty' : !!cell && cell.value.kind !== 'empty';
    if (!willHoldValue) continue;
    const style = current(row, column);
    if (!NumberFormats.isGeneral(style.numberFormat)) continue; // a chosen format survives
    applied = restyle(row, column, { ...style, numberFormat: format }) || applied;
  }
  return applied;
}

function inferFormat(scan: ColumnScan, locale: string): NumberFormat | null {
  if (!scan.isUniformlyNumeric) return null;
  if (scan.dateCount === scan.numericCount && scan.dateCode !== null) return NumberFormats.of(scan.dateCode);
  // As precise as anything typed, capped: a stray 12-decimal float is not a column's precision.
  const decimals = Math.min(scan.decimals, MAX_INFERRED_DECIMALS);
  const zeros = (n: number): string => '0'.repeat(n);
  if (scan.percentCount > 0) return NumberFormats.of(decimals === 0 ? '0%' : `0.${zeros(decimals)}%`);
  if (scan.currencyCount > 0) {
    // Money shows its cents even when every amount is round.
    return NumberFormats.of(
      `${currencySymbolCode(locale)}#,##0.${zeros(Math.max(decimals, CONVENTIONAL_DECIMALS))}`,
    );
  }
  if (scan.allIntegers) return scan.maxMagnitude >= THOUSANDS_THRESHOLD ? NumberFormats.of('#,##0') : null;
  // Decimals read to two places, so 900 lines up under 1,200.50.
  return NumberFormats.of(`#,##0.${zeros(Math.max(decimals, CONVENTIONAL_DECIMALS))}`);
}

function currencySymbolCode(locale: string): string {
  const symbol = localeInfo(locale).currencySymbol;
  return BARE_CURRENCY.test(symbol) ? symbol : `"${symbol}"`;
}

function alignBody(
  column: number,
  top: number,
  bottom: number,
  sheet: Worksheet,
  current: CurrentStyle,
  restyle: Restyle,
): number {
  let realigned = 0;
  for (let row = top; row <= bottom; row++) {
    const cell = sheet.find(row, column);
    if (!cell || cell.value.kind === 'empty') continue;
    const style = current(row, column);
    if (
      style.alignment.horizontal !== 'general' &&
      restyle(row, column, { ...style, alignment: { ...style.alignment, horizontal: 'general' } })
    ) {
      realigned++;
    }
  }
  return realigned;
}

// ---- Widths -------------------------------------------------------------------

function fitColumns(
  sheet: Worksheet,
  block: CellRange,
  locale: string,
  measurer: TextMeasurer,
  values: ReadonlyMap<number, CellValue>,
  current: CurrentStyle,
  widths: Map<number, number>,
): void {
  for (let column = block.left; column <= block.right; column++) {
    let widest = 0;
    for (let row = block.top; row <= block.bottom; row++) {
      const value = values.get(cellKey(row, column)) ?? sheet.getValue(row, column);
      if (value.kind === 'empty') continue;
      const style = current(row, column);
      if (style.alignment.wrapText) continue; // wrapped text folds; it doesn't drag the column wide
      const text = NumberFormatter.format(value, style.numberFormat.code, locale);
      widest = Math.max(widest, measurer.measure(text, style.font));
    }
    const target =
      widest > 0 ? Math.min(MAX_WIDTH, Math.max(MIN_WIDTH, widest + WIDTH_PADDING)) : Defaults.columnWidth;
    if (Math.abs(target - sheet.getColumnWidth(column)) > 0.5) widths.set(column, target);
  }
}

// ---- Reading what a text cell really is ------------------------------------------

/**
 * Text that only looks numeric: digits with a leading zero (and more than one digit) or more than 15
 * digits, a phone number written with a leading "+", or a code in exponent form such as "12E3".
 */
function isIdentifier(text: string): boolean {
  if (PHONE_LIKE.test(text) || EXPONENT_LIKE.test(text)) return true;
  if (!/^[0-9]+$/.test(text)) return false;
  return text.length > MAX_SAFE_DIGITS || (text.length > 1 && text[0] === '0');
}

function tryCoerce(text: string, locale: string): Coercion | null {
  if (text.length === 0) return null;
  const info = localeInfo(locale);
  if (text.endsWith('%')) {
    const body = text.slice(0, -1).trimEnd();
    const percent = body.length === 0 ? null : parseNumberAny(body, info);
    if (percent === null) return null;
    return {
      value: CellValue.number(percent / 100),
      dateFormat: null,
      decimals: textFractionDigits(body, locale),
      kind: 'percent',
    };
  }
  const parsed = InputParser.parse(text, locale);
  if (parsed.formula !== null || parsed.value.kind !== 'number') return null;
  const kind: NumberFormatKind = parsed.suggestedFormat
    ? 'date'
    : looksLikeCurrency(text, info.currencySymbol)
      ? 'currency'
      : 'number';
  return {
    value: parsed.value,
    dateFormat: parsed.suggestedFormat,
    decimals: textFractionDigits(text, locale),
    kind,
  };
}

function looksLikeCurrency(text: string, symbol: string): boolean {
  return (symbol.length > 0 && text.includes(symbol)) || ANY_CURRENCY.test(text);
}

/** Digits after the last decimal separator as typed. */
function textFractionDigits(text: string, locale: string): number {
  const separator = localeInfo(locale).decimal;
  const at = text.lastIndexOf(separator);
  if (at < 0) return 0;
  let digits = 0;
  for (let i = at + separator.length; i < text.length && /[0-9]/.test(text[i]!); i++) digits++;
  return digits;
}

/** Decimal places of a number written to at most 10 places ("0.##########"), so float noise doesn't count. */
function numberFractionDigits(n: number): number {
  const fixed = Math.abs(n).toFixed(10).replace(/0+$/, '');
  const dot = fixed.indexOf('.');
  return dot < 0 || dot === fixed.length - 1 ? 0 : fixed.length - dot - 1;
}
