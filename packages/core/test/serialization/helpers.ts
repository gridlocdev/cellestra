import { strFromU8, unzipSync } from 'fflate';
import type { Workbook } from '../../src/model/workbook.ts';
import { readOds, readOdsWithPackage, writeOds } from '../../src/serialization/index.ts';
import type { OdsReadResult } from '../../src/serialization/index.ts';

/** save → reopen through `.ods` bytes. */
export function roundTrip(workbook: Workbook): Workbook {
  return readOds(writeOds(workbook));
}

export function roundTripWithPackage(workbook: Workbook): OdsReadResult {
  return readOdsWithPackage(writeOds(workbook));
}

/** One zip entry's text, or throws when the entry is missing. */
export function readEntry(bytes: Uint8Array, name: string): string {
  const entries = unzipSync(bytes, { filter: (file) => file.name === name });
  const data = entries[name];
  if (!data) throw new Error(`No zip entry named '${name}'.`);
  return strFromU8(data);
}

export function entryNames(bytes: Uint8Array): string[] {
  return Object.keys(unzipSync(bytes));
}
