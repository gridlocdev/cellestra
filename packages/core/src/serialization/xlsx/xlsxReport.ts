/** What an `.xlsx` can hold that Cellestra leaves out when it opens one (PRD §8.19). */
export type XlsxLeftOut =
  | 'drawings'
  | 'conditionalFormatting'
  | 'dataValidation'
  | 'comments'
  | 'pivotTables'
  | 'tables'
  | 'autoFilters'
  | 'macros'
  | 'chartSheets'
  | 'otherSheets'
  | 'names';

const LABELS: Readonly<Record<XlsxLeftOut, string>> = {
  drawings: 'charts, images and shapes',
  conditionalFormatting: 'conditional formatting',
  dataValidation: 'data validation',
  comments: 'comments',
  pivotTables: 'pivot tables',
  tables: 'tables (their cells are kept)',
  autoFilters: 'AutoFilters',
  macros: 'macros',
  chartSheets: 'chart sheets',
  otherSheets: 'macro and dialog sheets',
  names: 'named ranges that aren’t a single range',
};

/** What an `.xlsx` import changed or left out, for the notice and the CLI's warning. */
export interface XlsxImportReport {
  readonly leftOut: ReadonlySet<XlsxLeftOut>;
  /** Formulas kept as their cached result: unparsed, an unknown function or name, or an array formula. */
  readonly formulasAsValues: number;
  /** Hidden sheets, which come in visible (Cellestra has no hidden sheets). */
  readonly hiddenSheetsShown: number;
}

/** The report being filled while a workbook is read. */
export class XlsxReportBuilder {
  readonly leftOut = new Set<XlsxLeftOut>();
  formulasAsValues = 0;
  hiddenSheetsShown = 0;

  build(): XlsxImportReport {
    return {
      leftOut: new Set(this.leftOut),
      formulasAsValues: this.formulasAsValues,
      hiddenSheetsShown: this.hiddenSheetsShown,
    };
  }
}

const plural = (n: number, one: string, many: string): string => `${n} ${n === 1 ? one : many}`;

/**
 * The report as sentences (“Left out: comments, AutoFilters. 3 formulas are kept as values.”), or an
 * empty string when everything came across.
 */
export function describeXlsxImport(report: XlsxImportReport): string {
  const parts: string[] = [];
  const kinds = (Object.keys(LABELS) as XlsxLeftOut[]).filter((k) => report.leftOut.has(k));
  if (kinds.length > 0) parts.push(`Left out: ${kinds.map((k) => LABELS[k]).join(', ')}.`);
  if (report.formulasAsValues > 0)
    parts.push(
      `${plural(report.formulasAsValues, 'formula is', 'formulas are')} kept as values, since Cellestra can’t calculate ${report.formulasAsValues === 1 ? 'it' : 'them'}.`,
    );
  if (report.hiddenSheetsShown > 0)
    parts.push(`${plural(report.hiddenSheetsShown, 'hidden sheet is', 'hidden sheets are')} shown.`);
  return parts.join(' ');
}
