/** Success; for validate, diff and find, nothing to report. */
export const EXIT_OK = 0;
/** A problem was found (validate, diff, lint --strict, find), or the input was bad (missing file). */
export const EXIT_PROBLEM = 1;
/** The command line itself was wrong: unknown option or command, missing argument, unparsable value. */
export const EXIT_USAGE = 2;

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

/** A command line that doesn't parse or names no command: like `CliError`, but exit code 2. */
export class UsageError extends CliError {
  constructor(message: string) {
    super(message);
    this.name = 'UsageError';
  }
}

/**
 * Whether an exception is the user's to fix rather than a bug. The engine reports bad input
 * (duplicate sheet, invalid name, overlapping merge, malformed color) with plain `Error`,
 * `RangeError`, or `SyntaxError`; file-system failures carry an errno `code` (`ENOENT`, `EACCES`).
 * A `TypeError`, a Node internal error (`ERR_*` codes), and the runtime's own `RangeError`s (stack
 * overflow, invalid array length) are bugs and keep their stack trace.
 */
export function isExpectedError(e: unknown): e is Error {
  if (e instanceof CliError) return true;
  if (!(e instanceof Error) || e instanceof TypeError) return false;
  const code = (e as NodeJS.ErrnoException).code;
  if (code !== undefined) return typeof code === 'string' && ERRNO_CODE.test(code);
  if (e instanceof RangeError) return !RUNTIME_RANGE_ERROR.test(e.message);
  return (
    e.constructor === Error ||
    e instanceof SyntaxError ||
    /^(FormulaParseException|Ods\w*TooLargeError|XmlParseError)$/.test(e.name)
  );
}

/** An errno code from a system call (`ENOENT`, `EISDIR`), not a Node `ERR_*` code. */
const ERRNO_CODE = /^E[A-Z0-9]+$/;

/** RangeErrors V8 throws by itself: they mean the CLI has a bug, not that the input was bad. */
const RUNTIME_RANGE_ERROR =
  /^(Maximum call stack size exceeded|Invalid (array|string|typed array) length|Invalid (count|time|code point) value|Array buffer allocation failed|.* argument must be between)/;
