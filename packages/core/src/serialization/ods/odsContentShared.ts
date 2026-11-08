import { CellAddress } from '../../model/cellAddress.ts';
import type { CellRange } from '../../model/cellRange.ts';
import { XmlElement, XmlText } from '../../xml/index.ts';
import { Office, Text } from './odsXmlNames.ts';

/**
 * The unit conversions and ODF address formatters shared by the `content.xml` reader, writer,
 * validation, and database-range halves.
 */

export const PX_PER_INCH = 96.0;
export const CM_PER_INCH = 2.54;

/** Per sheet name: each validation rule's target and the name it is written under. */
export type ValidationNames = ReadonlyMap<
  string,
  ReadonlyArray<{ readonly range: CellRange; readonly name: string }>
>;

/** `$Sheet1.$A$1` or `$Sheet1.$A$1:$Sheet1.$B$2`. */
export function odfRangeAddress(range: CellRange, fallbackSheet: string): string {
  const sheet = range.sheet ?? fallbackSheet;
  const start = odfCellAddress(new CellAddress(range.left, range.top, false, false, sheet), fallbackSheet);
  if (range.rowCount === 1 && range.columnCount === 1) return start;
  const end = odfCellAddress(new CellAddress(range.right, range.bottom, false, false, sheet), fallbackSheet);
  return `${start}:${end}`;
}

export function odfCellAddress(address: CellAddress, fallbackSheet: string): string {
  const sheet = address.sheet ?? fallbackSheet;
  return `$${odfSheetName(sheet)}.$${address.columnLabel}$${address.rowLabel}`;
}

/**
 * A sheet name as an ODF address writes it: bare when it is a plain identifier, otherwise in single
 * quotes with any quote doubled (`'My Sheet'`, `'It''s'`). An unquoted `$My Sheet.$A$1` is not an
 * address at all to LibreOffice, so named ranges, validations and filters on it were lost.
 */
export function odfSheetName(sheet: string): string {
  const plain = /^[\p{L}_][\p{L}\p{N}_]*$/u.test(sheet) && !/^[A-Za-z]{1,3}\d+$/.test(sheet);
  return plain ? sheet : `'${sheet.replaceAll("'", "''")}'`;
}

export function pxToCm(px: number): number {
  return (px / PX_PER_INCH) * CM_PER_INCH;
}

export function cmToPx(cm: number): number {
  return (cm / CM_PER_INCH) * PX_PER_INCH;
}

/** .NET's `ToString("0.###", Invariant)`. */
export function formatUpTo3(value: number): string {
  return String(Number.parseFloat(value.toFixed(3)));
}

// ---- paragraph text (cells, comments, validation messages) ----------------

/** The most spaces one `text:s` may stand for (a hostile `text:c` would otherwise allocate gigabytes). */
const MAX_SPACE_RUN = 65536;

/**
 * The text of one `text:p` (ODF 6.1): `text:s` is `text:c` spaces (default 1), `text:tab` a tab,
 * `text:line-break` a newline, and spans, links and other inline markup contribute their own text.
 * Whitespace-only text nodes are content too (a cell holding " " is a space). Only paragraphs are
 * joined with "\n", by the caller. Iterative, so deeply nested spans cannot overflow the stack.
 */
export function paragraphText(paragraph: XmlElement): string {
  let text = '';
  const stack: Array<{ e: XmlElement; i: number }> = [{ e: paragraph, i: 0 }];
  while (stack.length > 0) {
    const top = stack[stack.length - 1]!;
    if (top.i >= top.e.nodes.length) {
      stack.pop();
      continue;
    }
    const node = top.e.nodes[top.i++]!;
    if (node instanceof XmlText) {
      text += node.value;
      continue;
    }
    if (!(node instanceof XmlElement)) continue;
    if (node.uri === Text.uri) {
      if (node.local === 's') {
        const count = spaceCount(node.attr(Text('c')));
        text += ' '.repeat(Math.min(count, MAX_SPACE_RUN));
        continue;
      }
      if (node.local === 'tab') {
        text += '\t';
        continue;
      }
      if (node.local === 'line-break') {
        text += '\n';
        continue;
      }
      // A footnote or an annotation inside the paragraph is not part of the cell's text.
      if (node.local === 'note') continue;
    }
    if (node.is(Office('annotation'))) continue;
    stack.push({ e: node, i: 0 });
  }
  return text;
}

/** `text:c`: a positive count, 1 when absent or malformed. */
function spaceCount(text: string | null): number {
  const t = text?.trim() ?? '';
  if (!/^[+-]?\d+$/.test(t)) return 1;
  const n = Number(t);
  return n > 0 ? n : 1;
}

/**
 * Append a line of cell text the way ODF white-space handling (ODF 6.1.2) reads it back: a reader
 * collapses runs of spaces and drops tabs, so a run of n spaces is a space plus `text:s c=n-1` (a
 * leading space is all `text:s`), and a tab is `text:tab`. Plain text stays plain.
 */
export function appendParagraphText(target: XmlElement, line: string): void {
  if (!/ {2}|^ | $|\t/.test(line)) {
    if (line.length > 0) target.add(line);
    if (target.nodes.length === 0) target.add(new XmlText(''));
    return;
  }
  // Mixed content from here on: the empty text node keeps the writer from indenting it.
  target.add(new XmlText(''));
  let i = 0;
  while (i < line.length) {
    const c = line[i]!;
    if (c === '\t') {
      target.add(new XmlElement(Text('tab')));
      i++;
      continue;
    }
    if (c === ' ') {
      let n = 1;
      while (line[i + n] === ' ') n++;
      // Only a space between words survives as itself; leading, trailing and after-tab ones do not.
      const literal = i > 0 && line[i - 1] !== '\t' && i + n < line.length ? 1 : 0;
      if (literal) target.add(' ');
      const rest = n - literal;
      if (rest > 0) {
        const s = new XmlElement(Text('s'));
        if (rest > 1) s.setAttr(Text('c'), rest);
        target.add(s);
      }
      i += n;
      continue;
    }
    let j = i;
    while (j < line.length && line[j] !== ' ' && line[j] !== '\t') j++;
    target.add(line.slice(i, j));
    i = j;
  }
}
