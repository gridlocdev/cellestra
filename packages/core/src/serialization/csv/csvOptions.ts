/**
 * CSV import/export options (PRD §8.4). UTF-8 throughout; comma is the default delimiter (semicolon
 * and tab are selectable); RFC 4180 quoting; export writes LF line endings.
 */
export interface CsvOptions {
  /** The field separator, one character. */
  readonly delimiter: string;
  /**
   * The BCP 47 locale used to infer numbers and dates on import and to format values on export.
   * Omitted: the workbook's effective locale.
   */
  readonly locale?: string | null;
}

export const CsvOptions = {
  default: Object.freeze({ delimiter: ',' }) as CsvOptions,
  comma: Object.freeze({ delimiter: ',' }) as CsvOptions,
  semicolon: Object.freeze({ delimiter: ';' }) as CsvOptions,
  tab: Object.freeze({ delimiter: '\t' }) as CsvOptions,
  /** The delimiter a file's name implies: tab for a .tsv, comma for anything else. */
  forFileName(name: string): CsvOptions {
    return /\.tsv$/i.test(name) ? CsvOptions.tab : CsvOptions.comma;
  },
} as const;
