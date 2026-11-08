/**
 * A mistake the user can fix (bad address, missing file, refusing to overwrite): printed as
 * `cellestra: <message>` with exit code 1, never as a stack trace.
 */
export class CliError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'CliError';
  }
}

/**
 * Whether an exception is the user's to fix rather than a bug. The engine reports bad input
 * (duplicate sheet, invalid name, overlapping merge, malformed color) with plain `Error`,
 * `RangeError`, or `SyntaxError`; file-system failures carry an errno `code`. A `TypeError` is a
 * bug and keeps its stack trace.
 */
export function isExpectedError(e: unknown): e is Error {
  if (e instanceof CliError) return true;
  if (!(e instanceof Error)) return false;
  if ((e as NodeJS.ErrnoException).code !== undefined) return true;
  return (
    e.constructor === Error ||
    e instanceof RangeError ||
    e instanceof SyntaxError ||
    /^(FormulaParseException|Ods\w*TooLargeError|XmlParseError)$/.test(e.name)
  );
}
