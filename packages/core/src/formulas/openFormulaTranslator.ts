import type { UnparsedFormulaSource } from '../model/cell.ts';
import { quoteSheetIfNeeded, type CellAddress } from '../model/cellAddress.ts';
import type { CellRange } from '../model/cellRange.ts';
import type { BinaryExpr, Expr, UnaryExpr } from './ast/expr.ts';
import {
  BINARY_SYMBOL,
  FormulaWriter,
  literalText,
  wholeLinesText,
  writtenPrecedence,
} from './formulaWriter.ts';
import { Parser } from './parsing/parser.ts';

const PREFIX = 'of:=';
const ERROR_REFS = new Set(['#REF!', '#NAME?', '#N/A', '#VALUE!', '#DIV/0!']);

/**
 * Excel 2010+ functions LibreOffice Calc stores under the `COM.MICROSOFT.` namespace in OpenFormula
 * (unprefixed, Calc reads them as #NAME?). Written with the prefix, read back without it.
 */
const MICROSOFT_PREFIX = 'COM.MICROSOFT.';
const MICROSOFT_FUNCTIONS = new Set([
  'CONCAT',
  'TEXTJOIN',
  'IFS',
  'SWITCH',
  'MAXIFS',
  'MINIFS',
  'XLOOKUP',
  'XMATCH',
]);
const MICROSOFT_CALL = /^COM\.MICROSOFT\.([A-Z][A-Z0-9.]*)(?=\s*\()/i;

// ---- A1 AST → OpenFormula --------------------------------------------------

// OpenFormula, LibreOffice Calc and Cellestra all bind a prefix sign tighter than ^ (-2^2 = 4). The
// writer still spells a signed base `(-2)^2`, so no reader is left to guess.

function withParens(expr: Expr, minPrecedence: number): string {
  const rendered = render(expr);
  return writtenPrecedence(expr) < minPrecedence ? `(${rendered})` : rendered;
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
  const p = writtenPrecedence(b);
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
    case 'call': {
      const name = MICROSOFT_FUNCTIONS.has(expr.name) ? MICROSOFT_PREFIX + expr.name : expr.name;
      return `${name}(${expr.args.map(render).join(';')})`; // OpenFormula argument separator
    }
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
      const inner = body.slice(i + 1, close); // ".A1", ".A1:.A3", "Sheet2.A1", "#REF!", ".#REF!"
      // Calc writes a reference broken by a deletion as [#REF!], [.#REF!], [.A1:.#REF!] or
      // [$#REF!.A1]: any of them is one broken reference.
      out +=
        ERROR_REFS.has(inner) || inner.toUpperCase().includes('#REF!')
          ? '#REF!'
          : inner.split(':').map(convertCorner).join(':');
      i = close;
    } else if (c === ';') {
      out += ','; // OpenFormula → internal argument separator
    } else if ((c === 'C' || c === 'c') && !/[\p{L}\p{Nd}_.$]/u.test(body[i - 1] ?? '')) {
      const call = MICROSOFT_CALL.exec(body.slice(i, i + 64));
      const name = call?.[1]?.toUpperCase();
      if (call && name && MICROSOFT_FUNCTIONS.has(name)) {
        out += name;
        i += call[0].length - 1;
      } else out += c;
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
 * Converts between internal A1 formula text and ODF OpenFormula: bracketed sheet-dot references
 * (`[.A1]`, `[Sheet2.A1]`), `;` argument separators, and the `of:=` prefix. Always go through here
 * when reading/writing `.ods` formulas.
 */
export const OpenFormulaTranslator = {
  /**
   * Internal A1 formula (no leading `=`) → OpenFormula with the `of:=` prefix. A formula the parser
   * cannot read is written as the file gave it when `source` is its loaded text, else as typed.
   */
  toOpenFormula(internalFormula: string, source: UnparsedFormulaSource | null = null): string {
    const parsed = Parser.tryParse(internalFormula);
    if (parsed.ok) return PREFIX + render(parsed.expr);
    return source?.formula === internalFormula ? source.openFormula : PREFIX + internalFormula;
  },

  /** OpenFormula string → internal A1 formula (no leading `=`). */
  fromOpenFormula(openFormula: string): string {
    return OpenFormulaTranslator.read(openFormula).formula;
  },

  /**
   * OpenFormula string → internal A1 formula, plus the original text to keep with the cell when the
   * parser cannot read it, so a save writes it back unchanged.
   */
  read(openFormula: string): { formula: string; source: UnparsedFormulaSource | null } {
    const internal = convertReferencesAndSeparators(stripPrefix(openFormula));
    // Normalize through the parser when possible; keep the converted text otherwise so formulas
    // using unmodeled functions still survive the round-trip.
    const parsed = Parser.tryParse(internal);
    if (parsed.ok) return { formula: FormulaWriter.write(parsed.expr), source: null };
    return { formula: internal, source: { formula: internal, openFormula } };
  },
} as const;
