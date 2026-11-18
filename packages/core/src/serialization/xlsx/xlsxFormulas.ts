import type { Expr } from '../../formulas/ast/expr.ts';
import { FormulaWriter } from '../../formulas/formulaWriter.ts';
import { FunctionRegistry } from '../../formulas/functions/functionRegistry.ts';
import { Parser } from '../../formulas/parsing/parser.ts';
import { ReferenceAdjuster } from '../../formulas/references/referenceAdjuster.ts';
import { quoteSheetIfNeeded } from '../../model/cellAddress.ts';

/** Prefixes Excel writes before functions newer than 2007 (`_xlfn.CONCAT`, `_xlfn._xlws.SORT`). */
const FUNCTION_PREFIX = /_xl(fn|ws)\./gi;

const NAME_CHAR = /[\p{L}\p{Nd}_.\\]/u;
const CELL_OR_RANGE = /^(\$?[A-Za-z]{1,3}\$?\d+)(:\$?[A-Za-z]{1,3}\$?\d+)?/;

/**
 * Excel formula text (no `=`) → Cellestra's A1 syntax: `Sheet2!A1` → `Sheet2.A1`, `'My Sheet'!A1`
 * → `'My Sheet'.A1`, `Sheet1:Sheet3!A1:B2` → `Sheet1.A1:Sheet3.B2`, and the `_xlfn.` prefixes go.
 * Null when the formula uses what Cellestra has no syntax for: structured references
 * (`Table1[Total]`) and other workbooks (`[1]Sheet1!A1`).
 */
export function excelToInternal(formula: string): string | null {
  const body = formula.replace(FUNCTION_PREFIX, '');
  let out = '';
  for (let i = 0; i < body.length; i++) {
    const c = body[i]!;
    if (c === '"') {
      const end = closeQuote(body, i, '"');
      out += body.slice(i, end + 1);
      i = end;
    } else if (c === '[') {
      return null;
    } else if (c === "'") {
      const end = closeQuote(body, i, "'");
      const name = body.slice(i + 1, end).replaceAll("''", "'");
      if (body[end + 1] !== '!') return null;
      const sheet = sheetPrefix(name, body.slice(end + 2));
      if (sheet === null) return null;
      out += sheet.text;
      i = end + 1 + sheet.consumed;
    } else if (NAME_CHAR.test(c) && !NAME_CHAR.test(body[i - 1] ?? '')) {
      let end = i;
      while (end < body.length && NAME_CHAR.test(body[end]!)) end++;
      // `Sheet1:Sheet3!` names a span of sheets.
      const span = /^:([\p{L}\p{Nd}_.\\]+)!/u.exec(body.slice(end));
      if (body[end] === '!' || span) {
        const name = span ? `${body.slice(i, end)}:${span[1]}` : body.slice(i, end);
        const rest = body.slice(end + (span ? span[0].length : 1));
        const sheet = sheetPrefix(name, rest);
        if (sheet === null) return null;
        out += sheet.text;
        i = end + (span ? span[0].length : 1) + sheet.consumed - 1;
      } else {
        out += body.slice(i, end);
        i = end - 1;
      }
    } else {
      out += c;
    }
  }
  return out;
}

/** The index of the quote closing the one at `start` (a doubled quote is part of the text). */
function closeQuote(text: string, start: number, quote: string): number {
  for (let i = start + 1; i < text.length; i++) {
    if (text[i] !== quote) continue;
    if (text[i + 1] === quote) i++;
    else return i;
  }
  return text.length - 1;
}

/**
 * The Cellestra text for a reference after `name!`: `Sheet.A1` or `Sheet.A1:B2`, or for a span of
 * sheets `First.A1:Last.B2`. `consumed` is how much of `rest` the reference used; a plain sheet
 * prefix (the reference is left to the main loop) consumes nothing.
 */
function sheetPrefix(name: string, rest: string): { text: string; consumed: number } | null {
  const colon = name.indexOf(':');
  if (colon < 0) return { text: `${quoteSheetIfNeeded(name)}.`, consumed: 0 };
  const ref = CELL_OR_RANGE.exec(rest);
  if (!ref) return null;
  const first = quoteSheetIfNeeded(name.slice(0, colon));
  const last = quoteSheetIfNeeded(name.slice(colon + 1));
  const end = (ref[2] ?? `:${ref[1]}`).slice(1);
  return { text: `${first}.${ref[1]}:${last}.${end}`, consumed: ref[0].length };
}

/** Why a formula is kept as its cached value instead (PRD §8.19), or null when it can be kept. */
export type FormulaRefusal = 'syntax' | 'function';

/** A formula from the file made ready for the model: Cellestra text, or why it can't be one. */
export type ImportedFormula =
  | { readonly ok: true; readonly formula: string; readonly expr: Expr }
  | { readonly ok: false; readonly reason: FormulaRefusal };

/**
 * Translate, parse and check an Excel formula: it must parse, every function it calls must be in the
 * registry, and every name it reads must have been imported (`hasName`), or recalculation would turn
 * the file's result into an error.
 */
export function importFormula(excel: string, hasName: (name: string) => boolean): ImportedFormula {
  const internal = excelToInternal(excel);
  if (internal === null) return { ok: false, reason: 'syntax' };
  const parsed = Parser.tryParse(internal);
  if (!parsed.ok) return { ok: false, reason: 'syntax' };
  if (!computable(parsed.expr, hasName)) return { ok: false, reason: 'function' };
  return { ok: true, formula: FormulaWriter.write(parsed.expr), expr: parsed.expr };
}

/** A shared formula's text for a cell `rows`/`columns` away from the one that holds it. */
export function shiftFormula(expr: Expr, rows: number, columns: number): string {
  return FormulaWriter.write(ReferenceAdjuster.forCopy(expr, columns, rows));
}

function computable(expr: Expr, hasName: (name: string) => boolean): boolean {
  const registry = FunctionRegistry.standard;
  const visit = (e: Expr): boolean => {
    switch (e.kind) {
      case 'call':
        return registry.has(e.name) && e.args.every(visit);
      case 'unary':
        return visit(e.operand);
      case 'binary':
        return visit(e.left) && visit(e.right);
      case 'name':
        return hasName(e.name);
      case 'literal':
      case 'array':
      case 'reference':
      case 'range':
      case 'sheetSpan':
      case 'refError':
        return true;
    }
  };
  return visit(expr);
}
