import { quoteSheetIfNeeded, type CellAddress } from '../model/cellAddress.ts';
import type { CellRange } from '../model/cellRange.ts';
import type { BinaryExpr, Expr, UnaryExpr } from './ast/expr.ts';
import {
  BINARY_SYMBOL,
  binaryPrecedence,
  FormulaWriter,
  literalText,
  wholeLinesText,
} from './formulaWriter.ts';
import { Parser } from './parsing/parser.ts';

const PREFIX = 'of:=';
const ERROR_REFS = new Set(['#REF!', '#NAME?', '#N/A', '#VALUE!', '#DIV/0!']);

// ---- A1 AST → OpenFormula --------------------------------------------------

// Readers disagree on whether unary minus binds tighter than ^ (OpenFormula and Calc: yes; Cellestra,
// PRD §8: no, so -2^2 = -4). The writer never leaves that to the reader: a sign on a power is written
// `-(2^2)` and a power of a signed value `(-2)^2`, which every reader evaluates the same way.

function precedence(expr: Expr): number {
  if (expr.kind === 'binary') return binaryPrecedence(expr.operator, 5);
  if (expr.kind === 'unary') return expr.operator === 'percent' ? 7 : 6;
  return 8;
}

function withParens(expr: Expr, minPrecedence: number): string {
  const rendered = render(expr);
  return precedence(expr) < minPrecedence ? `(${rendered})` : rendered;
}

function renderUnary(u: UnaryExpr): string {
  switch (u.operator) {
    case 'negate':
      return '-' + withParens(u.operand, 6);
    case 'plus':
      return '+' + withParens(u.operand, 6);
    case 'percent':
      return withParens(u.operand, 7) + '%';
  }
}

function renderBinary(b: BinaryExpr): string {
  const p = precedence(b);
  const signedBase = b.operator === 'power' && b.left.kind === 'unary' && b.left.operator !== 'percent';
  const left = signedBase ? `(${render(b.left)})` : withParens(b.left, p);
  return `${left}${BINARY_SYMBOL[b.operator]}${withParens(b.right, p + 1)}`;
}

const sheetDot = (sheet: string | null): string => (sheet === null ? '.' : `${quoteSheetIfNeeded(sheet)}.`);

const cell = (a: CellAddress): string =>
  `${a.columnAbsolute ? '$' : ''}${a.columnLabel}${a.rowAbsolute ? '$' : ''}${a.rowLabel}`;

const renderAddress = (a: CellAddress): string => `[${sheetDot(a.sheet)}${cell(a)}]`;

const renderRange = (r: CellRange): string =>
  `[${wholeLinesText(r, sheetDot, '.') ?? `${sheetDot(r.start.sheet)}${cell(r.start)}:.${cell(r.end)}`}]`;

function render(expr: Expr): string {
  switch (expr.kind) {
    case 'literal':
      return literalText(expr.value, '()');
    case 'reference':
      return renderAddress(expr.address);
    case 'range':
      return renderRange(expr.range);
    case 'name':
      return expr.name;
    case 'refError':
      return '[#REF!]';
    case 'unary':
      return renderUnary(expr);
    case 'binary':
      return renderBinary(expr);
    case 'call':
      return `${expr.name}(${expr.args.map(render).join(';')})`; // OpenFormula argument separator
  }
}

// ---- OpenFormula text → internal A1 ----------------------------------------

function stripPrefix(formula: string): string {
  const body = formula.trimStart();
  if (body.startsWith(PREFIX)) return body.slice(PREFIX.length);
  return body.startsWith('=') ? body.slice(1) : body;
}

function convertReferencesAndSeparators(body: string): string {
  let out = '';
  for (let i = 0; i < body.length; i++) {
    const c = body[i]!;
    if (c === '"') {
      // Copy a string literal verbatim (doubled quotes included).
      out += '"';
      i++;
      for (; i < body.length; i++) {
        out += body[i];
        if (body[i] === '"') {
          if (body[i + 1] === '"') {
            out += '"';
            i++;
            continue;
          }
          break;
        }
      }
      if (i >= body.length) i = body.length - 1;
    } else if (c === '[') {
      const close = body.indexOf(']', i);
      if (close < 0) {
        out += c;
        continue;
      }
      const inner = body.slice(i + 1, close); // ".A1", ".A1:.A3", "Sheet2.A1", "#REF!"
      out += ERROR_REFS.has(inner) ? '#REF!' : inner.split(':').map(convertCorner).join(':');
      i = close;
    } else if (c === ';') {
      out += ','; // OpenFormula → internal argument separator
    } else {
      out += c;
    }
  }
  return out;
}

function convertCorner(corner: string): string {
  // ".A1" → "A1"; "Sheet2.A1" stays; "$Sheet2.A1" → "Sheet2.A1" (no syntax for an absolute sheet).
  if (corner.startsWith('.')) return corner.slice(1);
  return corner.startsWith('$') ? corner.slice(1) : corner;
}

/**
 * The original OpenFormula text of formulas the parser could not read, by the internal text they
 * were loaded as, so saving writes them back exactly as they came (bounded: oldest entries go).
 */
const unparsedOriginals = new Map<string, string>();
const MAX_UNPARSED_ORIGINALS = 10_000;

function rememberUnparsed(internal: string, original: string): void {
  unparsedOriginals.delete(internal);
  unparsedOriginals.set(internal, original);
  if (unparsedOriginals.size > MAX_UNPARSED_ORIGINALS) {
    const oldest = unparsedOriginals.keys().next().value;
    if (oldest !== undefined) unparsedOriginals.delete(oldest);
  }
}

/**
 * Converts between internal A1 formula text and ODF OpenFormula: bracketed sheet-dot references
 * (`[.A1]`, `[Sheet2.A1]`), `;` argument separators, and the `of:=` prefix. Always go through here
 * when reading/writing `.ods` formulas.
 */
export const OpenFormulaTranslator = {
  /** Internal A1 formula (no leading `=`) → OpenFormula with the `of:=` prefix. */
  toOpenFormula(internalFormula: string): string {
    const parsed = Parser.tryParse(internalFormula);
    if (parsed.ok) return PREFIX + render(parsed.expr);
    // Unparseable: the OpenFormula it was loaded from, unchanged; else the text as typed.
    return unparsedOriginals.get(internalFormula) ?? PREFIX + internalFormula;
  },

  /** OpenFormula string → internal A1 formula (no leading `=`). */
  fromOpenFormula(openFormula: string): string {
    const internal = convertReferencesAndSeparators(stripPrefix(openFormula));
    // Normalize through the parser when possible; keep the converted text otherwise so formulas
    // using unmodeled functions still survive the round-trip.
    const parsed = Parser.tryParse(internal);
    if (parsed.ok) return FormulaWriter.write(parsed.expr);
    rememberUnparsed(internal, openFormula);
    return internal;
  },
} as const;
