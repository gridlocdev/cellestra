const isReferenceChar = (c: string): boolean => /^[\p{L}\p{Nd}$]$/u.test(c);
const isLetter = (c: string): boolean => /^\p{L}$/u.test(c);
const isDigit = (c: string): boolean => /^\p{Nd}$/u.test(c);

interface Reference {
  readonly columnAbsolute: boolean;
  readonly column: string;
  readonly rowAbsolute: boolean;
  readonly row: string;
}

function parseReference(token: string): Reference | null {
  let i = 0;
  const columnAbsolute = token[i] === '$';
  if (columnAbsolute) i++;
  const columnStart = i;
  while (i < token.length && isLetter(token[i]!)) i++;
  if (i === columnStart) return null; // no column letters
  const column = token.slice(columnStart, i);
  const rowAbsolute = token[i] === '$';
  if (rowAbsolute) i++;
  const rowStart = i;
  while (i < token.length && isDigit(token[i]!)) i++;
  if (i === rowStart || i !== token.length) return null; // no row digits, or trailing junk
  return { columnAbsolute, column, rowAbsolute, row: token.slice(rowStart, i) };
}

function nextForm(r: Reference): string {
  // A1 → $A$1 → A$1 → $A1 → A1
  let col: boolean;
  let row: boolean;
  if (!r.columnAbsolute && !r.rowAbsolute) [col, row] = [true, true];
  else if (r.columnAbsolute && r.rowAbsolute) [col, row] = [false, true];
  else if (!r.columnAbsolute && r.rowAbsolute) [col, row] = [true, false];
  else [col, row] = [false, false];
  return (col ? '$' : '') + r.column + (row ? '$' : '') + r.row;
}

/**
 * Cycles the absolute/relative markers of the cell reference at the caret (F4):
 * `A1 → $A$1 → A$1 → $A1 → A1`. Pure text manipulation.
 */
export const ReferenceCycler = {
  /**
   * The text with the reference at `caret` cycled, and the caret moved to the reference's end.
   * Unchanged when the caret is not on a cell reference.
   */
  cycle(text: string, caret: number): { text: string; caret: number } {
    if (!text) return { text, caret };
    caret = Math.max(0, Math.min(caret, text.length));
    let start = caret;
    while (start > 0 && isReferenceChar(text[start - 1]!)) start--;
    let end = caret;
    while (end < text.length && isReferenceChar(text[end]!)) end++;
    const reference = end > start ? parseReference(text.slice(start, end)) : null;
    if (!reference) return { text, caret };
    const cycled = nextForm(reference);
    return { text: text.slice(0, start) + cycled + text.slice(end), caret: start + cycled.length };
  },
} as const;
