import { SheetDimensions } from './sheetDimensions.ts';

/**
 * A cell location in A1 notation. Zero-based column/row, absolute-vs-relative per axis
 * (`$A$1` vs `A1`), optionally sheet-qualified (`Sheet2.A1`). Two addresses are `equals` only
 * when every part matches; compare `samePosition` to ignore `$` flags and sheet.
 */
export class CellAddress {
  readonly column: number;
  readonly row: number;
  readonly columnAbsolute: boolean;
  readonly rowAbsolute: boolean;
  readonly sheet: string | null;

  constructor(
    column: number,
    row: number,
    columnAbsolute = false,
    rowAbsolute = false,
    sheet: string | null = null,
  ) {
    if (!SheetDimensions.isValidColumn(column)) throw new RangeError(`Column index out of range: ${column}`);
    if (!SheetDimensions.isValidRow(row)) throw new RangeError(`Row index out of range: ${row}`);
    this.column = column;
    this.row = row;
    this.columnAbsolute = columnAbsolute;
    this.rowAbsolute = rowAbsolute;
    this.sheet = sheet;
  }

  static of(row: number, column: number): CellAddress {
    return new CellAddress(column, row);
  }

  get columnLabel(): string {
    return SheetDimensions.columnToLetters(this.column);
  }

  /** One-based row label as shown in the UI. */
  get rowLabel(): number {
    return this.row + 1;
  }

  withSheet(sheet: string | null): CellAddress {
    return new CellAddress(this.column, this.row, this.columnAbsolute, this.rowAbsolute, sheet);
  }

  withoutSheet(): CellAddress {
    return this.withSheet(null);
  }

  /** Strip absolute flags and sheet. */
  relative(): CellAddress {
    return new CellAddress(this.column, this.row);
  }

  /** Offset the relative parts; absolute parts stay put (ReferenceAdjuster semantics). Throws if off-sheet. */
  shift(columnDelta: number, rowDelta: number): CellAddress {
    const c = this.columnAbsolute ? this.column : this.column + columnDelta;
    const r = this.rowAbsolute ? this.row : this.row + rowDelta;
    return new CellAddress(c, r, this.columnAbsolute, this.rowAbsolute, this.sheet);
  }

  samePosition(other: CellAddress): boolean {
    return this.row === other.row && this.column === other.column;
  }

  equals(other: CellAddress | null | undefined): boolean {
    return (
      !!other &&
      this.row === other.row &&
      this.column === other.column &&
      this.columnAbsolute === other.columnAbsolute &&
      this.rowAbsolute === other.rowAbsolute &&
      this.sheet === other.sheet
    );
  }

  toString(): string {
    const a1 = `${this.columnAbsolute ? '$' : ''}${this.columnLabel}${this.rowAbsolute ? '$' : ''}${this.rowLabel}`;
    return this.sheet === null ? a1 : `${quoteSheetIfNeeded(this.sheet)}.${a1}`;
  }

  /** Parse `B7`, `$B$7`, `Sheet1.B7`, or `'My Sheet'.B7`. Throws on bad input. */
  static parse(text: string): CellAddress {
    const a = CellAddress.tryParse(text);
    if (!a) throw new SyntaxError(`'${text}' is not a valid cell reference.`);
    return a;
  }

  static tryParse(text: string | null | undefined): CellAddress | null {
    if (!text || !text.trim()) return null;
    let span = text.trim();
    let sheet: string | null = null;
    const dot = findSheetSeparator(span);
    if (dot >= 0) {
      sheet = unquoteSheet(span.slice(0, dot));
      if (sheet === null) return null;
      span = span.slice(dot + 1);
    }
    return parseA1(span, sheet);
  }
}

const A1 = /^(\$?)([A-Za-z]+)(\$?)([0-9]+)$/;

function parseA1(span: string, sheet: string | null): CellAddress | null {
  const m = A1.exec(span);
  if (!m) return null;
  const column = SheetDimensions.lettersToColumn(m[2]!);
  if (column === null) return null;
  const oneBased = Number.parseInt(m[4]!, 10);
  if (!(oneBased >= 1) || !SheetDimensions.isValidRow(oneBased - 1)) return null;
  return new CellAddress(column, oneBased - 1, m[1] === '$', m[3] === '$', sheet);
}

/** The LAST unquoted dot splits an optional sheet prefix. */
export function findSheetSeparator(span: string): number {
  let inQuote = false;
  let lastDot = -1;
  for (let i = 0; i < span.length; i++) {
    const c = span[i];
    if (c === "'") inQuote = !inQuote;
    else if (c === '.' && !inQuote) lastDot = i;
  }
  return lastDot;
}

/** `'It''s'` → `It's`; bare names pass through; empty → null. */
export function unquoteSheet(sheet: string): string | null {
  if (sheet.length === 0) return null;
  if (sheet.length >= 2 && sheet.startsWith("'") && sheet.endsWith("'")) {
    // ODF quotes embedded apostrophes by doubling them.
    const name = sheet.slice(1, -1).replaceAll("''", "'");
    return name.length > 0 ? name : null;
  }
  return sheet;
}

/**
 * Quote a sheet name unless it reads back as a plain name: it must be letters, digits and `_` only,
 * must not start with a digit (`'2024'.A1`), and must not itself look like a cell reference
 * on this sheet's grid (`'AB12'.A1`; `Sheet1` is fine, SHEET is past XFD) or a boolean keyword.
 */
export function quoteSheetIfNeeded(sheet: string): string {
  const needs =
    sheet.length === 0 ||
    /[^\p{L}\p{N}_]/u.test(sheet) ||
    /^\p{N}/u.test(sheet) ||
    parseA1(sheet, null) !== null ||
    /^(TRUE|FALSE)$/i.test(sheet);
  return needs ? `'${sheet.replaceAll("'", "''")}'` : sheet;
}
