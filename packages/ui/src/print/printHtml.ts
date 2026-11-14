import {
  Argb,
  PageSetups,
  SheetDimensions,
  StylePool,
  type BorderEdge,
  type CellStyle,
  type CellValue,
  type PageSetup,
  type Worksheet,
} from '@cellestra/core';

/** Cells to print, inclusive. */
export interface PrintRange {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

export interface PrintInput {
  readonly sheet: Worksheet;
  readonly styles: StylePool;
  /** The sheet's page setup; null prints on the default page. */
  readonly setup: PageSetup | null;
  /** What to print; null prints every used cell of the sheet. */
  readonly range: PrintRange | null;
  /** A value as the grid shows it in `style` (number format applied). */
  readonly display: (value: CellValue, style: CellStyle) => string;
  /** The style conditional formats give a cell. */
  readonly conditional?: (style: CellStyle, row: number, column: number) => CellStyle;
  /**
   * Rows their content makes taller than their stored height (a big font, wrapped text), as the
   * grid measures them (`contentRowHeights`). A row prints at the larger of the two.
   */
  readonly contentHeights?: ReadonlyMap<number, number>;
  readonly title: string;
}

export interface PrintDocument {
  readonly html: string;
  readonly pages: number;
  readonly setup: PageSetup;
}

const PX_PER_MM = 96 / 25.4;
const HEADER_COLUMN_PX = 40;
const HEADER_ROW_PX = 20;
const CELL_PAD_X = 5;

/**
 * A sheet as a self-contained HTML document for printing or saving as PDF: one block per page, each
 * a table sized and scaled by the page setup, in the setup's page order. Pages are cut here rather
 * than by the browser, so a sheet wider than the paper goes onto further pages instead of being cut
 * off, and "fit to N pages" can be honoured.
 */
export function printHtml(input: PrintInput): PrintDocument {
  const setup = input.setup ?? PageSetups.default;
  const { sheet } = input;
  const range = input.range ?? usedRange(sheet);
  const rows = lines(range.top, range.bottom, (r) => sheet.isRowHidden(r));
  const columns = lines(range.left, range.right, (c) => sheet.isColumnHidden(c));
  const width = (c: number) => sheet.getColumnWidth(c);
  const height = rowHeight(input);

  const pageW = (setup.paperWidth - setup.margins.left - setup.margins.right) * PX_PER_MM;
  const pageH = (setup.paperHeight - setup.margins.top - setup.margins.bottom) * PX_PER_MM;
  const extraW = setup.headers ? HEADER_COLUMN_PX : 0;
  const extraH = setup.headers ? HEADER_ROW_PX : 0;
  const layout = (scale: number) => ({
    scale,
    columnPages: chunks(columns, width, pageW / scale - extraW),
    rowPages: chunks(rows, height, pageH / scale - extraH),
  });

  const totalW = sum(columns, width) + extraW;
  const totalH = sum(rows, height) + extraH;
  const s = setup.scale;
  let pages: ReturnType<typeof layout>;
  if (s.kind === 'percent') pages = layout(Math.max(0.1, s.percent / 100));
  else if (s.kind === 'fit') {
    let scale = 1;
    if (s.width > 0) scale = Math.min(scale, (s.width * pageW) / totalW);
    if (s.height > 0) scale = Math.min(scale, (s.height * pageH) / totalH);
    pages = layout(scale);
    // Page cuts fall between lines, so the ideal scale can still spill: shrink until it fits.
    while (
      pages.scale > 0.1 &&
      ((s.width > 0 && pages.columnPages.length > s.width) ||
        (s.height > 0 && pages.rowPages.length > s.height))
    )
      pages = layout(pages.scale * 0.97);
  } else {
    pages = layout(1);
    while (pages.scale > 0.1 && pages.columnPages.length * pages.rowPages.length > s.pages)
      pages = layout(pages.scale * 0.97);
  }

  const order: [readonly number[], readonly number[]][] = [];
  if (setup.pageOrder === 'downThenOver') {
    for (const cs of pages.columnPages) for (const rs of pages.rowPages) order.push([rs, cs]);
  } else {
    for (const rs of pages.rowPages) for (const cs of pages.columnPages) order.push([rs, cs]);
  }

  const classes = new Map<string, string>();
  const body = order
    .map(([rs, cs]) => page(input, setup, rs, cs, pages.scale, pageW, pageH, classes))
    .join('');
  const css = [...classes].map(([rule, name]) => `.${name}{${rule}}`).join('\n');
  const m = setup.margins;
  const html =
    `<!doctype html><html><head><meta charset="utf-8"><title>${escape(input.title)}</title><style>` +
    `@page{size:${setup.paperWidth}mm ${setup.paperHeight}mm;margin:${m.top}mm ${m.right}mm ${m.bottom}mm ${m.left}mm}` +
    `html,body{margin:0;padding:0;background:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact}` +
    `.page{display:flex;overflow:hidden;box-sizing:border-box;break-after:page}` +
    `.page:last-child{break-after:auto}` +
    `table{border-collapse:collapse;table-layout:fixed;flex:none}` +
    `td,th{padding:0 ${CELL_PAD_X}px;overflow:visible;white-space:pre;box-sizing:border-box}` +
    `th{background:#eee;color:#444;font:normal 9pt Helvetica,Arial,sans-serif;border:0.5px solid #999;text-align:center}` +
    `.r90{display:inline-block;writing-mode:vertical-rl;transform:rotate(180deg)}` +
    `.r270{display:inline-block;writing-mode:vertical-rl}` +
    `\n${css}</style></head><body>${body}</body></html>`;
  return { html, pages: order.length, setup };
}

/** A row's printed height: its stored height, or what its content needs when that is more. */
function rowHeight(input: PrintInput): (row: number) => number {
  const content = input.contentHeights;
  return (r) => Math.max(input.sheet.getRowHeight(r), content?.get(r) ?? 0);
}

/** The used cells of a sheet, merges included; just A1 when it is empty. */
export function usedRange(sheet: Worksheet): PrintRange {
  const e = sheet.cellExtent();
  let bottom = e?.lastRow ?? 0;
  let right = e?.lastColumn ?? 0;
  for (const m of sheet.merges) {
    bottom = Math.max(bottom, m.range.bottom);
    right = Math.max(right, m.range.right);
  }
  return { top: 0, left: 0, bottom, right };
}

function lines(from: number, to: number, hidden: (i: number) => boolean): number[] {
  const out: number[] = [];
  for (let i = from; i <= to; i++) if (!hidden(i)) out.push(i);
  return out;
}

const sum = (xs: readonly number[], size: (i: number) => number) => xs.reduce((t, i) => t + size(i), 0);

/** Lines cut into runs that fit `room` each; a line larger than the room gets a page to itself. */
function chunks(xs: readonly number[], size: (i: number) => number, room: number): number[][] {
  const out: number[][] = [];
  let current: number[] = [];
  let used = 0;
  for (const i of xs) {
    const s = size(i);
    if (current.length > 0 && used + s > room) {
      out.push(current);
      current = [];
      used = 0;
    }
    current.push(i);
    used += s;
  }
  if (current.length > 0 || out.length === 0) out.push(current);
  return out;
}

function page(
  input: PrintInput,
  setup: PageSetup,
  rows: readonly number[],
  columns: readonly number[],
  scale: number,
  pageW: number,
  pageH: number,
  classes: Map<string, string>,
): string {
  const { sheet } = input;
  const height = rowHeight(input);
  const rowSet = new Set(rows);
  const columnSet = new Set(columns);
  const covered = new Set<string>();
  const spans = new Map<
    string,
    { rowspan: number; colspan: number; anchorRow: number; anchorColumn: number }
  >();
  for (const m of sheet.merges) {
    const r = m.range;
    const inRows = rows.filter((x) => x >= r.top && x <= r.bottom);
    const inColumns = columns.filter((x) => x >= r.left && x <= r.right);
    if (!inRows.length || !inColumns.length) continue;
    // The part of a merge on this page is drawn from its first visible cell here.
    for (const x of inRows) for (const y of inColumns) covered.add(`${x},${y}`);
    spans.set(`${inRows[0]},${inColumns[0]}`, {
      rowspan: inRows.length,
      colspan: inColumns.length,
      anchorRow: r.top,
      anchorColumn: r.left,
    });
  }

  const cols =
    (setup.headers ? `<col style="width:${HEADER_COLUMN_PX}px">` : '') +
    columns.map((c) => `<col style="width:${sheet.getColumnWidth(c)}px">`).join('');
  const head = setup.headers
    ? `<tr style="height:${HEADER_ROW_PX}px"><th></th>${columns
        .map((c) => `<th>${SheetDimensions.columnToLetters(c)}</th>`)
        .join('')}</tr>`
    : '';
  const body = rows
    .map((r) => {
      let tr = `<tr style="height:${height(r)}px">`;
      if (setup.headers) tr += `<th>${r + 1}</th>`;
      for (const c of columns) {
        const span = spans.get(`${r},${c}`);
        if (!span && covered.has(`${r},${c}`)) continue;
        // A merge's content belongs to its top-left cell, wherever that is.
        const row = span ? span.anchorRow : r;
        const column = span ? span.anchorColumn : c;
        // Shown once: on the page holding its top-left cell (a merge cut by a page edge is blank past it).
        const showContent = !span || (rowSet.has(row) && columnSet.has(column));
        tr += cell(input, setup, row, column, span, showContent, classes);
      }
      return tr + '</tr>';
    })
    .join('');
  const tableW = sum(columns, (c) => sheet.getColumnWidth(c)) + (setup.headers ? HEADER_COLUMN_PX : 0);
  const justify = setup.centerHorizontally ? 'center' : 'flex-start';
  const align = setup.centerVertically ? 'center' : 'flex-start';
  // A hair under the printable area, so rounding never pushes a page onto the next sheet of paper.
  return (
    `<div class="page" style="width:${Math.floor(pageW)}px;height:${Math.floor(pageH) - 2}px;` +
    `justify-content:${justify};align-items:${align}">` +
    `<table style="width:${tableW}px;zoom:${scale.toFixed(4)}"><colgroup>${cols}</colgroup>${head}${body}</table></div>`
  );
}

function cell(
  input: PrintInput,
  setup: PageSetup,
  row: number,
  column: number,
  span: { rowspan: number; colspan: number } | undefined,
  showContent: boolean,
  classes: Map<string, string>,
): string {
  const { sheet } = input;
  const stored = sheet.find(row, column);
  const id =
    stored && stored.styleId !== StylePool.defaultId ? stored.styleId : sheet.effectiveStyleId(row, column);
  let style = input.styles.get(id);
  if (style.conditionalMaps?.length && input.conditional) style = input.conditional(style, row, column);

  const value = stored?.value;
  let text = '';
  if (showContent && value && value.kind !== 'empty') {
    text =
      style.checkbox && value.kind === 'boolean' ? (value.value ? '☑' : '☐') : input.display(value, style);
  }
  const horizontal =
    style.alignment.horizontal !== 'general'
      ? style.alignment.horizontal
      : value?.kind === 'number'
        ? 'right'
        : value?.kind === 'boolean' || value?.kind === 'error'
          ? 'center'
          : 'left';
  const rule = cellCss(style, horizontal, setup.gridlines);
  let name = classes.get(rule);
  if (name === undefined) {
    name = `s${classes.size}`;
    classes.set(rule, name);
  }
  const attrs =
    (span && span.rowspan > 1 ? ` rowspan="${span.rowspan}"` : '') +
    (span && span.colspan > 1 ? ` colspan="${span.colspan}"` : '');
  const rotation = style.alignment.rotation ?? 0;
  const content =
    rotation === 90 || rotation === 270 ? `<span class="r${rotation}">${escape(text)}</span>` : escape(text);
  return `<td class="${name}"${attrs}>${content}</td>`;
}

function cellCss(style: CellStyle, horizontal: string, gridlines: boolean): string {
  const f = style.font;
  const decoration = [f.underline ? 'underline' : '', f.strikethrough ? 'line-through' : ''].filter(Boolean);
  const grid = gridlines ? '0.5px solid #c0c0c0' : 'none';
  const b = style.border;
  return [
    `font:${f.italic ? 'italic ' : ''}${f.bold ? 'bold ' : ''}${f.sizePoints}pt ${cssFamily(f.family)}`,
    `color:${Argb.isTransparent(f.color) ? '#000' : Argb.toCss(f.color)}`,
    style.fill.background !== null && !Argb.isTransparent(style.fill.background)
      ? `background:${Argb.toCss(style.fill.background)}`
      : '',
    decoration.length ? `text-decoration:${decoration.join(' ')}` : '',
    `text-align:${horizontal}`,
    `vertical-align:${style.alignment.vertical === 'middle' ? 'middle' : style.alignment.vertical}`,
    style.alignment.wrapText ? 'white-space:pre-wrap;overflow:hidden' : '',
    `border-top:${edge(b.top, grid)}`,
    `border-bottom:${edge(b.bottom, grid)}`,
    `border-left:${edge(b.left, grid)}`,
    `border-right:${edge(b.right, grid)}`,
  ]
    .filter(Boolean)
    .join(';');
}

function edge(e: BorderEdge, fallback: string): string {
  if (e.style === 'none') return fallback;
  const color = Argb.toCss(e.color);
  switch (e.style) {
    case 'thin':
      return `1px solid ${color}`;
    case 'medium':
      return `2px solid ${color}`;
    case 'thick':
      return `3px solid ${color}`;
    case 'dashed':
      return `1px dashed ${color}`;
    case 'dotted':
      return `1px dotted ${color}`;
    case 'double':
      return `3px double ${color}`;
  }
}

function cssFamily(family: string): string {
  return `"${family.replace(/["<>]/g, '')}", "Helvetica Neue", Helvetica, Arial, sans-serif`;
}

function escape(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
