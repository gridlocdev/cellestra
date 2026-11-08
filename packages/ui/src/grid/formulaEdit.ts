import {
  FormulaSignatureScanner,
  FunctionRegistry,
  Parser,
  SheetDimensions,
  type Expr,
  type FunctionMetadata,
  type Worksheet,
} from '@cellestra/core';
import type { GridRange } from '../session/selection.ts';

/**
 * Pure helpers behind the inline editor's formula aids: pointing at cells, the coloured reference
 * outlines, the argument hint, function-name suggestions, column type-ahead, and the fill target.
 * Kept free of the DOM so they can be tested on their own.
 */

/** The one function registry the evaluator also uses, so hints can't drift from behaviour. */
const registry = FunctionRegistry.createDefault();

/** Colours cycled across a formula's references, in the order they appear. */
export const REFERENCE_COLORS = ['#1a73e8', '#1e884e', '#e8710a', '#8e24aa', '#00839b', '#c62828'] as const;

const REFERENCE_OPENERS = new Set(['=', '+', '-', '*', '/', '^', '(', ',', ':', '<', '>', '&', '%', ';']);

/**
 * Whether a cell reference may go at `caret`: the text is a formula and the previous non-space
 * character is an operator, an opening parenthesis, or an argument separator.
 */
export function isReferenceExpected(text: string, caret: number): boolean {
  if (!text.startsWith('=')) return false;
  for (let i = Math.min(caret, text.length) - 1; i >= 0; i--) {
    const c = text[i]!;
    if (c === ' ') continue;
    return REFERENCE_OPENERS.has(c);
  }
  return false;
}

/** `A1` or `A1:B3` for a range in the editor's own sheet. */
export function referenceText(r: GridRange): string {
  const a = (row: number, column: number) => `${SheetDimensions.columnToLetters(column)}${row + 1}`;
  return r.top === r.bottom && r.left === r.right
    ? a(r.top, r.left)
    : `${a(r.top, r.left)}:${a(r.bottom, r.right)}`;
}

/** The ranges a formula refers to on `sheetName`, each with its colour. Empty when it doesn't parse. */
export function formulaReferences(text: string, sheetName: string): { range: GridRange; color: string }[] {
  if (text.length < 2 || !text.startsWith('=')) return [];
  const parsed = Parser.tryParse(text);
  if (!parsed.ok) return [];
  const found: { range: GridRange; sheet: string | null }[] = [];
  const walk = (e: Expr): void => {
    switch (e.kind) {
      case 'reference':
        found.push({
          range: {
            top: e.address.row,
            left: e.address.column,
            bottom: e.address.row,
            right: e.address.column,
          },
          sheet: e.address.sheet,
        });
        return;
      case 'range':
        found.push({
          range: { top: e.range.top, left: e.range.left, bottom: e.range.bottom, right: e.range.right },
          sheet: e.range.start.sheet,
        });
        return;
      case 'unary':
        walk(e.operand);
        return;
      case 'binary':
        walk(e.left);
        walk(e.right);
        return;
      case 'call':
        e.args.forEach(walk);
        return;
      default:
        return;
    }
  };
  walk(parsed.expr);
  const out: { range: GridRange; color: string }[] = [];
  found.forEach((f, i) => {
    // A reference to another sheet can't be shown here, but it still takes its colour slot.
    if (f.sheet !== null && f.sheet.toLowerCase() !== sheetName.toLowerCase()) return;
    out.push({ range: f.range, color: REFERENCE_COLORS[i % REFERENCE_COLORS.length]! });
  });
  return out;
}

/** One piece of the argument hint; `active` marks the argument the caret is in. */
export interface SignaturePart {
  readonly text: string;
  readonly active?: boolean;
}

/** The argument hint for the call around `caret`, such as `SUM(number1, **number2, …**)`; null outside one. */
export function signatureHint(
  text: string,
  caret: number,
): { meta: FunctionMetadata; parts: SignaturePart[] } | null {
  if (!text.startsWith('=')) return null;
  const site = FormulaSignatureScanner.find(text, caret);
  if (!site) return null;
  const fn = registry.get(site.functionName);
  if (!fn) return null;
  const meta = fn.metadata;
  const args = meta.args;
  // A repeatable trailing argument covers any further positions.
  let active = site.argumentIndex;
  if (args.length > 0 && args[args.length - 1]!.repeatable && active >= args.length) active = args.length - 1;
  const parts: SignaturePart[] = [{ text: `${meta.name}(` }];
  args.forEach((arg, i) => {
    if (i > 0) parts.push({ text: ', ' });
    let label = arg.repeatable ? `${arg.name}, …` : arg.name;
    if (arg.optional) label = `[${label}]`;
    parts.push(i === active ? { text: label, active: true } : { text: label });
  });
  parts.push({ text: ')' });
  return { meta, parts };
}

/**
 * The function name being typed at `caret` (letters right before it, after an operator or the
 * `=`), with where it starts; null when the caret isn't at the end of a name.
 */
export function functionPrefix(text: string, caret: number): { start: number; prefix: string } | null {
  if (!text.startsWith('=')) return null;
  // Inside a string literal nothing is a name.
  let quotes = 0;
  for (let i = 0; i < caret; i++) if (text[i] === '"') quotes++;
  if (quotes % 2 === 1) return null;
  let start = caret;
  while (start > 1 && /[A-Za-z0-9_.]/.test(text[start - 1]!)) start--;
  const prefix = text.slice(start, caret);
  if (!/^[A-Za-z]/.test(prefix)) return null;
  if (start > 1 && !REFERENCE_OPENERS.has(text[start - 1]!) && text[start - 1] !== ' ') return null;
  // A name followed by `(` is already a call.
  if (text[caret] === '(') return null;
  return { start, prefix };
}

/** Functions whose names start with `prefix` (case-insensitive), alphabetical, at most `limit`. */
export function functionSuggestions(prefix: string, limit = 8): FunctionMetadata[] {
  const p = prefix.toUpperCase();
  return registry
    .allMetadata()
    .filter((m) => m.name.startsWith(p))
    .sort((a, b) => a.name.localeCompare(b.name))
    .slice(0, limit);
}

/**
 * The nearest text entry above `row` in `column` that starts with `prefix` and is longer than it
 * (case-sensitive), or null. Walks up from the cell and stops at the first gap or a table header:
 * the contiguous block the cell belongs to is the list it is an entry of.
 */
export function columnCompletion(
  sheet: Worksheet,
  row: number,
  column: number,
  prefix: string,
): string | null {
  if (!prefix || prefix.startsWith('=')) return null;
  for (let r = row - 1; r >= 0; r--) {
    const table = sheet.tables.find(
      (t) => t.hasHeader && t.range.top === r && column >= t.range.left && column <= t.range.right,
    );
    if (table) return null;
    const cell = sheet.find(r, column);
    if (!cell || cell.value.kind === 'empty') return null;
    if (cell.value.kind !== 'text') continue;
    const v = cell.value.value;
    if (v.length > prefix.length && v.startsWith(prefix)) return v;
  }
  return null;
}

/** The range a fill-handle drag covers: the bounding box of the source and the cell under the pointer. */
export function fillTarget(source: GridRange, row: number, column: number): GridRange {
  return {
    top: Math.min(source.top, row),
    left: Math.min(source.left, column),
    bottom: Math.max(source.bottom, row),
    right: Math.max(source.right, column),
  };
}

/**
 * Double-clicking the fill handle fills down as far as the neighbouring column has data (the
 * left neighbour, or the right one from column A). Null when there is nothing to fill.
 */
export function fillDownTarget(sheet: Worksheet, source: GridRange): GridRange | null {
  const neighbour = source.left > 0 ? source.left - 1 : source.right + 1;
  let bottom = source.bottom;
  for (let row = source.bottom + 1; row < SheetDimensions.maxRows; row++) {
    const cell = sheet.find(row, neighbour);
    if (!cell || cell.value.kind === 'empty') break;
    bottom = row;
  }
  return bottom > source.bottom ? { ...source, bottom } : null;
}
