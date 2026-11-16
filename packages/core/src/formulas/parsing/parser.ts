import { CellAddress, findSheetSeparator, unquoteSheet } from '../../model/cellAddress.ts';
import { SheetDimensions } from '../../model/sheetDimensions.ts';
import { CellRange } from '../../model/cellRange.ts';
import { CellError } from '../../values/cellError.ts';
import { CellValue } from '../../values/cellValue.ts';
import { Expr, type BinaryOperator } from '../ast/expr.ts';
import { Lexer, type Token, type TokenType } from '../lexing/lexer.ts';
import { isReferenceExpr, treeDepth } from './exprTree.ts';
import { FormulaParseException, type ParseError } from './parseError.ts';

export type ParseResult =
  { readonly ok: true; readonly expr: Expr } | { readonly ok: false; readonly error: ParseError };

const COMPARISON: Partial<Record<TokenType, BinaryOperator>> = {
  equal: 'equal',
  notEqual: 'notEqual',
  less: 'less',
  lessEqual: 'lessEqual',
  greater: 'greater',
  greaterEqual: 'greaterEqual',
};

const DESCRIBE: Partial<Record<TokenType, string>> = {
  endOfInput: 'end of formula',
  rightParen: "')'",
  leftParen: "'('",
  comma: "','",
  number: 'a number',
  text: 'a string',
  name: 'a name',
  error: 'an error value',
};

function describe(type: TokenType): string {
  return DESCRIBE[type] ?? `'${type}'`;
}

const NUMBER = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/;

/** Most tokens in one formula (LibreOffice Calc's limit, `MAXCODE`). */
export const MAX_FORMULA_TOKENS = 8192;
/** Most parentheses, function calls and signs one formula may nest inside each other. */
export const MAX_FORMULA_NESTING = 256;
/**
 * Deepest a parsed formula's tree may be, counting operator chains too (`1+1+…+1` nests one level
 * per `+`). The evaluator, writers and reference walkers recurse over the tree; this keeps them
 * well inside the call stack.
 */
export const MAX_FORMULA_DEPTH = 1024;

/**
 * A recursive-descent parser for Cellestra formulas. Precedence (lowest → highest): comparison,
 * concatenation (`&`), add/subtract, multiply/divide, power, unary +/-, percent, then reference/range.
 * As in OpenFormula and LibreOffice Calc (and Excel) a sign binds tighter than `^`, so `-2^2 = 4`.
 */
export class Parser {
  private readonly tokens: readonly Token[];
  private index = 0;
  private depth = 0;

  constructor(tokens: readonly Token[]) {
    this.tokens = tokens;
  }

  /** Parse a formula, with or without its leading `=`. Throws `FormulaParseException`. */
  static parse(formula: string): Expr {
    let body = formula.trimStart();
    if (body.startsWith('=')) body = body.slice(1);
    const tokens = Lexer.tokenize(body);
    if (tokens.length - 1 > MAX_FORMULA_TOKENS)
      throw new FormulaParseException(
        `The formula is too long (more than ${MAX_FORMULA_TOKENS} parts).`,
        tokens[MAX_FORMULA_TOKENS]!.position,
      );
    const parser = new Parser(tokens);
    const expr = parser.parseExpression();
    parser.expect('endOfInput');
    if (treeDepth(expr) > MAX_FORMULA_DEPTH)
      throw new FormulaParseException('The formula is nested too deeply.', 0);
    return expr;
  }

  /** Parse without throwing: a formula that cannot be read comes back as its `ParseError`. */
  static tryParse(formula: string): ParseResult {
    try {
      return { ok: true, expr: Parser.parse(formula) };
    } catch (e) {
      if (e instanceof FormulaParseException) return { ok: false, error: e.error };
      // The depth limits should make this unreachable; never let a hostile formula crash a caller.
      if (e instanceof RangeError)
        return { ok: false, error: { reason: 'The formula is nested too deeply.', position: 0 } };
      throw e;
    }
  }

  /** Enter one level of nesting (a parenthesis, call or sign); throws past `MAX_FORMULA_NESTING`. */
  private enter(): void {
    if (++this.depth > MAX_FORMULA_NESTING)
      throw new FormulaParseException('The formula is nested too deeply.', this.current.position);
  }

  private leave(): void {
    this.depth--;
  }

  private get current(): Token {
    return this.tokens[this.index]!;
  }

  private advance(): Token {
    return this.tokens[this.index++]!;
  }

  private match(type: TokenType): boolean {
    if (this.current.type === type) {
      this.index++;
      return true;
    }
    return false;
  }

  private expect(type: TokenType): Token {
    if (this.current.type !== type)
      throw new FormulaParseException(
        `Expected ${describe(type)} but found ${describe(this.current.type)}.`,
        this.current.position,
      );
    return this.advance();
  }

  parseExpression(): Expr {
    return this.parseComparison();
  }

  private parseComparison(): Expr {
    let left = this.parseConcat();
    for (;;) {
      const op = COMPARISON[this.current.type];
      if (!op) return left;
      this.advance();
      left = Expr.binary(op, left, this.parseConcat());
    }
  }

  private parseConcat(): Expr {
    let left = this.parseAddSub();
    while (this.current.type === 'ampersand') {
      this.advance();
      left = Expr.binary('concat', left, this.parseAddSub());
    }
    return left;
  }

  private parseAddSub(): Expr {
    let left = this.parseMulDiv();
    while (this.current.type === 'plus' || this.current.type === 'minus') {
      const op = this.advance().type === 'plus' ? 'add' : 'subtract';
      left = Expr.binary(op, left, this.parseMulDiv());
    }
    return left;
  }

  private parseMulDiv(): Expr {
    let left = this.parsePower();
    while (this.current.type === 'star' || this.current.type === 'slash') {
      const op = this.advance().type === 'star' ? 'multiply' : 'divide';
      left = Expr.binary(op, left, this.parsePower());
    }
    return left;
  }

  private parsePower(): Expr {
    // Left-associative (2^3^2 = 64). Both operands may carry a sign: -2^2 = (-2)^2 = 4 and
    // 2^-2 = 0.25, as in LibreOffice Calc.
    let left = this.parseUnary();
    while (this.current.type === 'caret') {
      this.advance();
      left = Expr.binary('power', left, this.parseUnary());
    }
    return left;
  }

  private parseUnary(): Expr {
    // A prefix sign binds tighter than every binary operator, ^ included (OpenFormula / Calc order),
    // but looser than a postfix % (-5% = -(5%)).
    if (this.current.type === 'minus' || this.current.type === 'plus') {
      const op = this.advance().type === 'minus' ? 'negate' : 'plus';
      this.enter();
      const operand = this.parseUnary();
      this.leave();
      return Expr.unary(op, operand);
    }
    return this.parsePostfix();
  }

  private parsePostfix(): Expr {
    let expr = this.parseUnion();
    while (this.current.type === 'percent') {
      this.advance();
      expr = Expr.unary('percent', expr);
    }
    return expr;
  }

  /** The reference union `~`, looser than intersection but tighter than every other operator. */
  private parseUnion(): Expr {
    let left = this.parseIntersection();
    while (this.current.type === 'tilde') {
      this.advance();
      left = Expr.binary('union', left, this.parseIntersection());
    }
    return left;
  }

  /**
   * The reference intersection `!`, which binds tightest after `:` (as in Calc and OpenFormula). A
   * space between two references is the same operator, as in Calc and Excel (`A1:C3 B2:D4`); it is
   * written back as `!`.
   */
  private parseIntersection(): Expr {
    let left = this.parsePrimary();
    for (;;) {
      if (this.match('bang')) left = Expr.binary('intersect', left, this.parsePrimary());
      else if (this.atSpaceIntersection(left)) {
        const at = this.current.position;
        const right = this.parsePrimary();
        if (!isReferenceExpr(right))
          throw new FormulaParseException('Only references can be intersected with a space.', at);
        left = Expr.binary('intersect', left, right);
      } else return left;
    }
  }

  /** After a reference, whitespace and then something that can start another reference. */
  private atSpaceIntersection(left: Expr): boolean {
    const { type, spaced } = this.current;
    return (
      spaced === true &&
      (type === 'name' || type === 'leftParen' || type === 'number') &&
      isReferenceExpr(left)
    );
  }

  /**
   * After `(` and its first expression: a `,` or `;` makes the parentheses a reference list, a
   * union of every item (Calc's `(A1:A2;C1)`). Each item must be a reference.
   */
  private parseReferenceList(first: Expr): Expr {
    const isSeparator = (): boolean => this.current.type === 'comma' || this.current.type === 'semicolon';
    if (!isSeparator()) return first;
    this.requireReference(first);
    let list = first;
    while (isSeparator()) {
      this.advance();
      const item = this.parseExpression();
      this.requireReference(item);
      list = Expr.binary('union', list, item);
    }
    return list;
  }

  private requireReference(expr: Expr): void {
    if (isReferenceExpr(expr)) return;
    throw new FormulaParseException(
      'Only references can be listed inside parentheses.',
      this.current.position,
    );
  }

  private parsePrimary(): Expr {
    const token = this.current;
    switch (token.type) {
      case 'number': {
        this.advance();
        // `1:3` is a whole-row range.
        if (this.current.type === 'colon') {
          const rows = this.tryWholeLines(token.text);
          if (rows) return rows;
        }
        const value = parseNumber(token);
        // Too large for a double (1E400): Calc keeps the formula and shows an error, so does this.
        if (!Number.isFinite(value))
          return Expr.literal(CellValue.error(CellError.number(`${token.text} is too large a number.`)));
        return Expr.literal(CellValue.number(value));
      }
      case 'text':
        this.advance();
        return Expr.literal(CellValue.text(token.text));
      case 'error': {
        this.advance();
        const code = CellError.codeFromToken(token.text)!;
        if (code === 'reference') return Expr.refError;
        return Expr.literal(CellValue.error(new CellError(code, `The formula contains ${token.text}.`)));
      }
      case 'leftParen': {
        this.advance();
        this.enter();
        const inner = this.parseReferenceList(this.parseExpression());
        this.expect('rightParen');
        this.leave();
        return inner;
      }
      case 'name':
        return this.parseNameOperand();
      case 'leftBrace':
        return this.parseArray();
      case 'ampersand':
      case 'bang':
      case 'caret':
      case 'colon':
      case 'comma':
      case 'endOfInput':
      case 'equal':
      case 'greater':
      case 'greaterEqual':
      case 'less':
      case 'lessEqual':
      case 'minus':
      case 'notEqual':
      case 'percent':
      case 'plus':
      case 'rightBrace':
      case 'rightParen':
      case 'semicolon':
      case 'slash':
      case 'star':
      case 'tilde':
      default:
        throw new FormulaParseException(`Unexpected ${describe(token.type)}.`, token.position);
    }
  }

  /**
   * An inline array of constants: `,` separates columns and `;` rows (Calc's en-US separators). Every
   * row must have as many elements as the first.
   */
  private parseArray(): Expr {
    const open = this.expect('leftBrace');
    const rows: CellValue[][] = [];
    let row: CellValue[] = [this.parseArrayElement()];
    for (;;) {
      if (this.match('comma')) row.push(this.parseArrayElement());
      else if (this.match('semicolon')) {
        rows.push(row);
        row = [this.parseArrayElement()];
      } else break;
    }
    rows.push(row);
    this.expect('rightBrace');
    if (rows.some((r) => r.length !== rows[0]!.length))
      throw new FormulaParseException('Every row of an array must have the same length.', open.position);
    return Expr.array(rows);
  }

  /** One array element: a signed number, a string, TRUE or FALSE (with or without `()`), or an error. */
  private parseArrayElement(): CellValue {
    const token = this.current;
    let sign = 1;
    if (token.type === 'minus' || token.type === 'plus') {
      this.advance();
      if (token.type === 'minus') sign = -1;
      if (this.current.type !== 'number')
        throw new FormulaParseException('Expected a number after the sign.', this.current.position);
    }
    const element = this.current;
    switch (element.type) {
      case 'number': {
        this.advance();
        const value = sign * parseNumber(element);
        return Number.isFinite(value)
          ? CellValue.number(value)
          : CellValue.error(CellError.number(`${element.text} is too large a number.`));
      }
      case 'text':
        this.advance();
        return CellValue.text(element.text);
      case 'error': {
        this.advance();
        const code = CellError.codeFromToken(element.text)!;
        return CellValue.error(new CellError(code, `The array contains ${element.text}.`));
      }
      case 'name': {
        const upper = element.text.toUpperCase();
        if (upper !== 'TRUE' && upper !== 'FALSE') break;
        this.advance();
        // OpenFormula may spell the constant as the function TRUE().
        if (this.current.type === 'leftParen') {
          this.advance();
          this.expect('rightParen');
        }
        return CellValue.boolean(upper === 'TRUE');
      }
      case 'ampersand':
      case 'bang':
      case 'caret':
      case 'colon':
      case 'comma':
      case 'endOfInput':
      case 'equal':
      case 'greater':
      case 'greaterEqual':
      case 'leftBrace':
      case 'leftParen':
      case 'less':
      case 'lessEqual':
      case 'minus':
      case 'notEqual':
      case 'percent':
      case 'plus':
      case 'rightBrace':
      case 'rightParen':
      case 'semicolon':
      case 'slash':
      case 'star':
      case 'tilde':
      default:
        break;
    }
    throw new FormulaParseException(
      'An array may only hold numbers, text, TRUE, FALSE and error values.',
      element.position,
    );
  }

  private parseNameOperand(): Expr {
    const nameToken = this.advance();

    if (this.current.type === 'leftParen') return this.parseFunctionCall(nameToken);

    const address = CellAddress.tryParse(nameToken.text);
    if (address) {
      if (this.current.type === 'colon') {
        this.advance();
        return this.parseRangeEnd(address);
      }
      return Expr.reference(address);
    }
    // `A:C`, `$B:$B`, `Sheet2.A:A`, `$1:$1`: whole columns or rows.
    if (this.current.type === 'colon') {
      const lines = this.tryWholeLines(nameToken.text);
      if (lines) return lines;
    }

    const upper = nameToken.text.toUpperCase();
    if (upper === 'TRUE') return Expr.literal(CellValue.boolean(true));
    if (upper === 'FALSE') return Expr.literal(CellValue.boolean(false));

    if (!isValidName(nameToken.text))
      throw new FormulaParseException(
        `'${nameToken.text}' is not a valid reference or name.`,
        nameToken.position,
      );
    return Expr.name(nameToken.text);
  }

  private parseRangeEnd(start: CellAddress): Expr {
    const end = this.current.type === 'name' ? CellAddress.tryParse(this.current.text) : null;
    if (!end) throw new FormulaParseException("Expected a cell reference after ':'.", this.current.position);
    if (end.sheet !== null && end.sheet.toLowerCase() !== start.sheet?.toLowerCase()) {
      // Only a range qualified on both corners can span sheets; `A1:Sheet2.B2` has no first sheet.
      if (start.sheet === null)
        throw new FormulaParseException(
          'A range across sheets must name the sheet on both corners.',
          this.current.position,
        );
      this.advance();
      // A range from one sheet to another (`Sheet1.A1:Sheet3.B2`) spans every sheet between them.
      return Expr.sheetSpan(new CellRange(start, end.withSheet(start.sheet)), end.sheet);
    }
    this.advance();
    // A sheet qualifier on the left corner applies to the whole range.
    const qualified = start.sheet !== null && end.sheet === null ? end.withSheet(start.sheet) : end;
    return Expr.range(new CellRange(start, qualified));
  }

  /**
   * After a start token and before its `:`, read a whole-column (`A:C`) or whole-row (`1:3`) range.
   * Returns null, consuming nothing, when the two sides are not lines of the same kind. The range is
   * stored with absolute rows (columns) so it keeps spanning the whole sheet when copied.
   */
  private tryWholeLines(startText: string): Expr | null {
    const start = parseLine(startText);
    const endToken = this.tokens[this.index + 1];
    if (!start || !endToken || (endToken.type !== 'name' && endToken.type !== 'number')) return null;
    const end = parseLine(endToken.text);
    if (!end || end.kind !== start.kind) return null;
    // `Sheet1.A:Sheet3.C` spans sheets; an end on another sheet needs a sheet on the start too.
    const acrossSheets = end.sheet !== null && end.sheet.toLowerCase() !== start.sheet?.toLowerCase();
    if (acrossSheets && start.sheet === null) return null;
    this.index += 2; // the colon and the end
    const sheet = start.sheet;
    let range: CellRange;
    if (start.kind === 'column') {
      const lastRow = SheetDimensions.maxRows - 1;
      range = new CellRange(
        new CellAddress(start.index, 0, start.absolute, true, sheet),
        new CellAddress(end.index, lastRow, end.absolute, true, sheet),
      );
    } else {
      const lastColumn = SheetDimensions.maxColumns - 1;
      range = new CellRange(
        new CellAddress(0, start.index, true, start.absolute, sheet),
        new CellAddress(lastColumn, end.index, true, end.absolute, sheet),
      );
    }
    return acrossSheets ? Expr.sheetSpan(range, end.sheet!) : Expr.range(range);
  }

  private parseFunctionCall(nameToken: Token): Expr {
    this.expect('leftParen');
    this.enter();
    const args: Expr[] = [];
    if (this.current.type !== 'rightParen') {
      // An omitted argument (`IF(A1>0,,1)`) is an empty value, as in LibreOffice.
      const argument = (): Expr =>
        this.current.type === 'comma' || this.current.type === 'rightParen'
          ? Expr.literal(CellValue.empty)
          : this.parseExpression();
      args.push(argument());
      while (this.match('comma')) args.push(argument());
    }
    this.expect('rightParen');
    this.leave();
    return Expr.call(nameToken.text.toUpperCase(), args);
  }
}

interface Line {
  readonly kind: 'column' | 'row';
  readonly index: number;
  readonly absolute: boolean;
  readonly sheet: string | null;
}

/** `A`, `$C`, `Sheet2.B` (a column) or `1`, `$3`, `Sheet2.4` (a row); null otherwise. */
function parseLine(text: string): Line | null {
  let span = text;
  let sheet: string | null = null;
  const dot = findSheetSeparator(span);
  if (dot >= 0) {
    // As in `CellAddress`, a leading `$` marks an absolute sheet (`$Sheet1.A:$Sheet3.C`).
    const prefix = span.slice(0, dot);
    sheet = unquoteSheet(prefix.startsWith('$') ? prefix.slice(1) : prefix);
    if (sheet === null) return null;
    span = span.slice(dot + 1);
  }
  const m = /^(\$?)(?:([A-Za-z]+)|([0-9]+))$/.exec(span);
  if (!m) return null;
  const absolute = m[1] === '$';
  if (m[2] !== undefined) {
    const index = SheetDimensions.lettersToColumn(m[2]);
    return index === null ? null : { kind: 'column', index, absolute, sheet };
  }
  const index = Number.parseInt(m[3]!, 10) - 1;
  return SheetDimensions.isValidRow(index) ? { kind: 'row', index, absolute, sheet } : null;
}

function parseNumber(token: Token): number {
  const value = NUMBER.test(token.text) ? Number(token.text) : NaN;
  if (Number.isNaN(value))
    throw new FormulaParseException(`'${token.text}' is not a valid number.`, token.position);
  return value;
}

/** A defined name: starts with a letter or underscore, then letters/digits/underscores (no dots). */
function isValidName(text: string): boolean {
  return /^[\p{L}_][\p{L}\p{Nd}_]*$/u.test(text);
}
