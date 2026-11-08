import { SheetDimensions, type CellRange } from '@cellestra/core';
import { CliError } from './errors.ts';
import { toJson } from './json.ts';

export type OutputFormat = 'table' | 'csv' | 'tsv' | 'json';

export function parseFormat(value: string | null | undefined): OutputFormat {
  switch (value?.toLowerCase()) {
    case undefined:
    case '':
    case 'table':
      return 'table';
    case 'csv':
      return 'csv';
    case 'tsv':
      return 'tsv';
    case 'json':
      return 'json';
    default:
      throw new CliError(`Unknown format '${value}'. Use table, csv, tsv, or json.`);
  }
}

export function render(format: OutputFormat, sheetName: string, range: CellRange, cells: string[][]): string {
  switch (format) {
    case 'csv':
      return delimited(cells, ',', true);
    case 'tsv':
      return delimited(cells, '\t', false);
    case 'json':
      return toJson({ sheet: sheetName, range: range.toString(), rows: cells }) + '\n';
    default:
      return table(range, cells);
  }
}

function delimited(cells: string[][], delimiter: string, quote: boolean): string {
  let out = '';
  for (const row of cells) {
    out += row.map((f) => (quote ? quoteCsv(f, delimiter) : f.replaceAll('\t', ' '))).join(delimiter) + '\n';
  }
  return out;
}

function quoteCsv(field: string, delimiter: string): string {
  const mustQuote =
    field.includes(delimiter) || field.includes('"') || field.includes('\n') || field.includes('\r');
  return mustQuote ? '"' + field.replaceAll('"', '""') + '"' : field;
}

/** A boxed grid with column letters across the top and row numbers down a gutter. */
function table(range: CellRange, cells: string[][]): string {
  const columns = range.columnCount;
  const rows = range.rowCount;
  const headers = Array.from({ length: columns }, (_, c) => SheetDimensions.columnToLetters(range.left + c));
  const gutterHeader = '';
  const rowLabels = Array.from({ length: rows }, (_, r) => String(range.top + r + 1));
  const gutter = Math.max(gutterHeader.length, ...rowLabels.map((l) => l.length), 0);
  const widths = headers.map((h, c) => Math.max(h.length, ...cells.map((row) => row[c]!.length)));
  let out = appendRow(gutter, widths, gutterHeader, headers);
  out += '─'.repeat(gutter) + '─┼' + widths.map((w) => '─'.repeat(w + 2)).join('┼') + '\n';
  for (let r = 0; r < rows; r++) out += appendRow(gutter, widths, rowLabels[r]!, cells[r]!);
  return out;
}

function appendRow(gutter: number, widths: number[], label: string, values: readonly string[]): string {
  return (
    label.padStart(gutter) + ' │' + widths.map((w, c) => ' ' + values[c]!.padEnd(w) + ' ').join('│') + '\n'
  );
}
