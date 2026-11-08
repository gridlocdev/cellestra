import type { CsvOptions } from '@cellestra/core';
import { CliError } from '../infra/errors.ts';
import type { WorkbookFormat } from '../infra/workspace.ts';
import type { ArgumentSpec, OptionSpec } from '../spec.ts';

/** The options and arguments shared across commands. */

export const fileArg = (name = 'file', description = 'Workbook file (or - for stdin)'): ArgumentSpec => ({
  name,
  description,
});

export const sheetOpt: OptionSpec = {
  name: '--sheet',
  aliases: ['-s'],
  description: 'Worksheet name (default: the active/first sheet)',
};

export const rangeOpt: OptionSpec = {
  name: '--range',
  aliases: ['-r'],
  description: 'Cell range, e.g. A1:D20 (default: the used range)',
};

export const formatOpt: OptionSpec = {
  name: '--format',
  aliases: ['-f'],
  description: 'table | csv | tsv | json',
  default: 'table',
};

export const localeOpt: OptionSpec = {
  name: '--locale',
  description: 'Locale override for parsing/formatting, e.g. en-US',
};

export const fromOpt: OptionSpec = { name: '--from', description: 'Input format override: ods | csv' };
export const toOpt: OptionSpec = { name: '--to', description: 'Output format override: ods | csv' };

export const delimiterOpt: OptionSpec = {
  name: '--delimiter',
  aliases: ['-d'],
  description: 'CSV delimiter (default ,)',
};

export const inPlaceOpt: OptionSpec = {
  name: '--in-place',
  aliases: ['-i'],
  description: 'Edit the file in place (atomic write)',
  kind: 'flag',
};

export const outputOpt: OptionSpec = {
  name: '--output',
  aliases: ['-o'],
  description: 'Write the result to this file instead of in place',
};

export const formulaOpt: OptionSpec = {
  name: '--formula',
  description: 'Show formula sources instead of computed values',
  kind: 'flag',
};

export const jsonOpt: OptionSpec = {
  name: '--json',
  description: 'Print a JSON report to stdout (for scripts and agents)',
  kind: 'flag',
};

export function formatOverride(value: string | null | undefined): WorkbookFormat | null {
  switch (value?.toLowerCase()) {
    case undefined:
    case '':
      return null;
    case 'ods':
      return 'ods';
    case 'csv':
    case 'tsv':
      return 'csv';
    default:
      throw new CliError(`Unknown format '${value}'. Use ods or csv.`);
  }
}

export function csvOptions(delimiter: string | null | undefined): CsvOptions | null {
  if (delimiter === null || delimiter === undefined || delimiter === '') return null;
  if ([...delimiter].length !== 1) throw new CliError('--delimiter must be a single character.');
  return { delimiter };
}

/** The destination for an edit command: explicit -o, or the input file with --in-place. */
export function editTarget(inputPath: string, inPlace: boolean, output: string | null): string {
  if (output !== null) return output;
  if (inPlace) return inputPath;
  throw new CliError('Refusing to overwrite without --in-place or -o <file>.');
}
