import { FormulaParseException } from '../parsing/parseError.ts';

/** The kinds of lexical token in a Cellestra formula (user-facing A1 syntax). */
export type TokenType =
  | 'number'
  /** A double-quoted string literal; embedded quotes are doubled (`""`). */
  | 'text'
  /** A function name, named range, boolean keyword, or cell reference — the parser classifies it. */
  | 'name'
  | 'plus'
  | 'minus'
  | 'star'
  | 'slash'
  | 'caret'
  | 'ampersand'
  | 'percent'
  | 'equal'
  | 'notEqual'
  | 'less'
  | 'lessEqual'
  | 'greater'
  | 'greaterEqual'
  | 'leftParen'
  | 'rightParen'
  | 'comma'
  /** The range operator `:`. */
  | 'colon'
  | 'endOfInput';

/** One token: `text` is the source slice (decoded for strings); `position` its zero-based start offset. */
export interface Token {
  readonly type: TokenType;
  readonly text: string;
  readonly position: number;
}

const SINGLE: Readonly<Record<string, TokenType>> = {
  '+': 'plus',
  '-': 'minus',
  '*': 'star',
  '/': 'slash',
  '^': 'caret',
  '&': 'ampersand',
  '%': 'percent',
  '(': 'leftParen',
  ')': 'rightParen',
  ',': 'comma',
  ':': 'colon',
  '=': 'equal',
};

const isAsciiDigit = (c: string | undefined): boolean => c !== undefined && c >= '0' && c <= '9';
const isLetter = (c: string): boolean => /^\p{L}$/u.test(c);
const isNameChar = (c: string): boolean => /^[\p{L}\p{Nd}_$.]$/u.test(c);
const isWhiteSpace = (c: string): boolean => /^\s$/u.test(c);

/**
 * Turns a formula body (the text after the leading `=`) into tokens. Numbers always use `.` as the
 * decimal point and `,` as the argument separator regardless of locale — locale only affects display.
 */
export class Lexer {
  private readonly source: string;
  private pos = 0;

  constructor(source: string) {
    this.source = source ?? '';
  }

  static tokenize(source: string): Token[] {
    const lexer = new Lexer(source);
    const tokens: Token[] = [];
    let token: Token;
    do {
      token = lexer.next();
      tokens.push(token);
    } while (token.type !== 'endOfInput');
    return tokens;
  }

  next(): Token {
    this.skipWhitespace();
    const src = this.source;
    if (this.pos >= src.length) return { type: 'endOfInput', text: '', position: this.pos };

    const start = this.pos;
    const c = src[this.pos]!;
    if (c === '<') return this.lessOrCompound();
    if (c === '>') return this.greaterOrCompound();
    if (c === '"') return this.readString();
    const single = SINGLE[c];
    if (single) {
      this.pos++;
      return { type: single, text: c, position: start };
    }
    if (isAsciiDigit(c) || (c === '.' && isAsciiDigit(src[this.pos + 1]))) return this.readNumber();
    if (isLetter(c) || c === '_' || c === '$' || c === "'") return this.readName();
    throw new FormulaParseException(`Unexpected character '${c}'.`, start);
  }

  private lessOrCompound(): Token {
    const start = this.pos++;
    const n = this.source[this.pos];
    if (n === '=') {
      this.pos++;
      return { type: 'lessEqual', text: '<=', position: start };
    }
    if (n === '>') {
      this.pos++;
      return { type: 'notEqual', text: '<>', position: start };
    }
    return { type: 'less', text: '<', position: start };
  }

  private greaterOrCompound(): Token {
    const start = this.pos++;
    if (this.source[this.pos] === '=') {
      this.pos++;
      return { type: 'greaterEqual', text: '>=', position: start };
    }
    return { type: 'greater', text: '>', position: start };
  }

  private readString(): Token {
    const src = this.source;
    const start = this.pos;
    this.pos++; // opening quote
    let out = '';
    while (this.pos < src.length) {
      const c = src[this.pos]!;
      if (c === '"') {
        // A doubled quote ("") is an escaped quote character.
        if (src[this.pos + 1] === '"') {
          out += '"';
          this.pos += 2;
          continue;
        }
        this.pos++; // closing quote
        return { type: 'text', text: out, position: start };
      }
      out += c;
      this.pos++;
    }
    throw new FormulaParseException('Unterminated string literal.', start);
  }

  private readNumber(): Token {
    const src = this.source;
    const start = this.pos;
    while (isAsciiDigit(src[this.pos])) this.pos++;
    if (src[this.pos] === '.') {
      this.pos++;
      while (isAsciiDigit(src[this.pos])) this.pos++;
    }
    if (src[this.pos] === 'e' || src[this.pos] === 'E') {
      const expStart = this.pos;
      this.pos++;
      if (src[this.pos] === '+' || src[this.pos] === '-') this.pos++;
      if (!isAsciiDigit(src[this.pos])) {
        // Not actually an exponent (e.g. a trailing 'E'); rewind.
        this.pos = expStart;
      } else {
        while (isAsciiDigit(src[this.pos])) this.pos++;
      }
    }
    return { type: 'number', text: src.slice(start, this.pos), position: start };
  }

  private readName(): Token {
    const src = this.source;
    const start = this.pos;
    if (src[this.pos] === "'") this.readQuotedSheetPrefix(start);
    while (this.pos < src.length && isNameChar(src[this.pos]!)) this.pos++;
    if (this.pos === start) throw new FormulaParseException('Invalid name.', start);
    return { type: 'name', text: src.slice(start, this.pos), position: start };
  }

  private readQuotedSheetPrefix(start: number): void {
    const src = this.source;
    this.pos++; // opening quote
    while (this.pos < src.length) {
      if (src[this.pos] === "'") {
        if (src[this.pos + 1] === "'") {
          this.pos += 2; // escaped quote inside the sheet name
          continue;
        }
        this.pos++; // closing quote
        if (src[this.pos] !== '.')
          throw new FormulaParseException("A quoted sheet name must be followed by '.cell'.", start);
        this.pos++; // the dot; the A1 part is read by the caller's name loop
        return;
      }
      this.pos++;
    }
    throw new FormulaParseException('Unterminated sheet name.', start);
  }

  private skipWhitespace(): void {
    while (this.pos < this.source.length && isWhiteSpace(this.source[this.pos]!)) this.pos++;
  }
}
