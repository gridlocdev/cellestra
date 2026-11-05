/**
 * A formula parse problem: a plain-English `reason` plus the zero-based `position` in the formula
 * source where it was found.
 */
export interface ParseError {
  readonly reason: string;
  readonly position: number;
}

export function parseErrorToString(e: ParseError): string {
  return `${e.reason} (at position ${e.position})`;
}

/** Thrown by the lexer/parser when a formula cannot be read; `Parser.tryParse` turns it into a value. */
export class FormulaParseException extends Error {
  readonly error: ParseError;

  constructor(reason: string, position: number) {
    const error = { reason, position };
    super(parseErrorToString(error));
    this.name = 'FormulaParseException';
    this.error = error;
  }
}
