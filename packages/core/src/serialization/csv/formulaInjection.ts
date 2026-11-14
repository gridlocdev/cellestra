/**
 * CSV formula injection (OWASP "CSV Injection"). A text cell such as `=HYPERLINK(...)` or
 * `+cmd|' /C calc'!A0` is data here, but Excel, LibreOffice and Google Sheets run it as a formula
 * when they open the CSV. Export prefixes such text with an apostrophe, the OWASP-recommended
 * defuse that every spreadsheet shows as a plain text marker; import strips that apostrophe again so
 * our own CSV round-trips. Only text cells are touched: a number cell of -5 is written as -5.
 */

// The characters that start a formula in a spreadsheet, plus tab and CR, which some strip first.
const TRIGGERS = new Set(['=', '+', '-', '@', '\t', '\r']);
// A signed number or phone number written as text ("-12.5", "+1 555 123 4567") runs nothing.
const PLAIN_SIGNED_NUMBER = /^[+-][0-9][0-9 .,()-]*$/;

export function defuseFormulaText(text: string): string {
  const first = text[0];
  if (first === undefined || !TRIGGERS.has(first)) return text;
  if (PLAIN_SIGNED_NUMBER.test(text)) return text;
  return "'" + text;
}

/** The text a field held before `defuseFormulaText`, or null when it was not defused. */
export function restoreFormulaText(field: string): string | null {
  const next = field[1];
  return field[0] === "'" && next !== undefined && TRIGGERS.has(next) ? field.slice(1) : null;
}
