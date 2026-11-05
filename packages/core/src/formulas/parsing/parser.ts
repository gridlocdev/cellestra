import { CellAddress } from '../../model/cellAddress.ts';
import { CellRange } from '../../model/cellRange.ts';
import { CellValue } from '../../values/cellValue.ts';
import { Expr, type BinaryOperator } from '../ast/expr.ts';
import { Lexer, type Token, type TokenType } from '../lexing/lexer.ts';
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
};

function describe(type: TokenType): string {
  return DESCRIBE[type] ?? `'${type}'`;
}

const NUMBER = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/;

/**
 * A recursive-descent parser for Cellestra formulas. Precedence (lowest → highest): comparison,
 * concatenation (`&`), add/subtract, multiply/divide, unary +/-, power, percent, then reference/range.
 */
export class Parser {
  private readonly tokens: readonly Token[];
  private index = 0;

  constructor(tokens: readonly Token[]) {
    this.tokens = tokens;
  }

  /** Parse a formula, with or without its leading `=`. Throws `FormulaParseException`. */
  static parse(formula: string): Expr {
    let body = formula.trimStart();
    if (body.startsWith('=')) body = body.slice(1);
    const parser = new Parser(Lexer.tokenize(body));
    const expr = parser.parseExpression();
    parser.expect('endOfInput');
    return expr;
  }

  static tryParse(formula: string): ParseResult {
    try {
      return { ok: true, expr: Parser.parse(formula) };
    } catch (e) {
      if (e instanceof FormulaParseException) return { ok: false, error: e.error };
      throw e;
    }
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
    let left = this.parseUnary();
    while (this.current.type === 'star' || this.current.type === 'slash') {
      const op = this.advance().type === 'star' ? 'multiply' : 'divide';
      left = Expr.binary(op, left, this.parseUnary());
    }
    return left;
  }

  private parseUnary(): Expr {
    // Unary +/- binds looser than ^ but tighter than * and / (OpenFormula / LibreOffice Calc order),
    // so -2^2 = -(2^2) = -4. Deliberately differs from Excel, which gives 4.
    if (this.current.type === 'minus') {
      this.advance();
      return Expr.unary('negate', this.parseUnary());
    }
    if (this.current.type === 'plus') {
      this.advance();
      return Expr.unary('plus', this.parseUnary());
    }
    return this.parsePower();
  }

  private parsePower(): Expr {
    // Left-associative (2^3^2 = 64). A sign right after ^ (2^-3) binds only to the following postfix.
    let left = this.parsePostfix();
    while (this.current.type === 'caret') {
      this.advance();
      left = Expr.binary('power', left, this.parsePowerOperand());
    }
    return left;
  }

  private parsePowerOperand(): Expr {
    if (this.current.type === 'minus') {
      this.advance();
      return Expr.unary('negate', this.parsePowerOperand());
    }
    if (this.current.type === 'plus') {
      this.advance();
      return Expr.unary('plus', this.parsePowerOperand());
    }
    return this.parsePostfix();
  }

  private parsePostfix(): Expr {
    let expr = this.parsePrimary();
    while (this.current.type === 'percent') {
      this.advance();
      expr = Expr.unary('percent', expr);
    }
    return expr;
  }

  private parsePrimary(): Expr {
    const token = this.current;
    switch (token.type) {
      case 'number':
        this.advance();
        return Expr.literal(CellValue.number(parseNumber(token)));
      case 'text':
        this.advance();
        return Expr.literal(CellValue.text(token.text));
      case 'leftParen': {
        this.advance();
        const inner = this.parseExpression();
        this.expect('rightParen');
        return inner;
      }
      case 'name':
        return this.parseNameOperand();
      default:
        throw new FormulaParseException(`Unexpected ${describe(token.type)}.`, token.position);
    }
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
    this.advance();
    // A sheet qualifier on the left corner applies to the whole range.
    const qualified = start.sheet !== null && end.sheet === null ? end.withSheet(start.sheet) : end;
    return Expr.range(new CellRange(start, qualified));
  }

  private parseFunctionCall(nameToken: Token): Expr {
    this.expect('leftParen');
    const args: Expr[] = [];
    if (this.current.type !== 'rightParen') {
      args.push(this.parseExpression());
      while (this.match('comma')) args.push(this.parseExpression());
    }
    this.expect('rightParen');
    return Expr.call(nameToken.text.toUpperCase(), args);
  }
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
