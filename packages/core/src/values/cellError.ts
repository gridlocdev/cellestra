/** The canonical spreadsheet error codes (PRD §8, Formulas). */
export type CellErrorCode =
  'divideByZero' | 'reference' | 'value' | 'name' | 'notAvailable' | 'number' | 'null';

const TOKENS: Record<CellErrorCode, string> = {
  divideByZero: '#DIV/0!',
  reference: '#REF!',
  value: '#VALUE!',
  name: '#NAME?',
  notAvailable: '#N/A',
  number: '#NUM!',
  null: '#NULL!',
};

/**
 * An error carried by a cell or an evaluation: a canonical code (for the `#…!` token shown in the
 * cell) plus a plain-English reason shown on hover.
 */
export class CellError {
  readonly code: CellErrorCode;
  readonly reason: string;

  constructor(code: CellErrorCode, reason: string) {
    this.code = code;
    this.reason = reason;
  }

  /** The short token displayed in the cell, e.g. `#DIV/0!`. */
  get token(): string {
    return TOKENS[this.code];
  }

  equals(other: CellError | null | undefined): boolean {
    return !!other && other.code === this.code && other.reason === this.reason;
  }

  toString(): string {
    return this.token;
  }

  static divideByZero(reason = 'Division by zero.'): CellError {
    return new CellError('divideByZero', reason);
  }
  static reference(reason = 'Invalid cell reference.'): CellError {
    return new CellError('reference', reason);
  }
  /** A circular reference surfaces as #REF! with an explanatory reason. */
  static circular(reason = 'Circular reference.'): CellError {
    return new CellError('reference', reason);
  }
  static value(reason = 'A value is of the wrong type.'): CellError {
    return new CellError('value', reason);
  }
  static name(reason = 'Unknown name.'): CellError {
    return new CellError('name', reason);
  }
  static notAvailable(reason = 'Value not available.'): CellError {
    return new CellError('notAvailable', reason);
  }
  /** A number that cannot be represented, or an argument outside a function's domain (#NUM!). */
  static number(reason = 'The result is not a valid number.'): CellError {
    return new CellError('number', reason);
  }
  /** An empty intersection of ranges (#NULL!); only ever typed or loaded, never computed here. */
  static null(reason = 'The ranges do not intersect.'): CellError {
    return new CellError('null', reason);
  }

  /** Parse a token back to a code (`#N/A` → notAvailable); null if not an error token. */
  static codeFromToken(token: string): CellErrorCode | null {
    for (const [code, t] of Object.entries(TOKENS)) if (t === token) return code as CellErrorCode;
    return null;
  }
}
