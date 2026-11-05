/**
 * File I/O: `.ods` (ODF 1.3 spreadsheet) and CSV. Byte-oriented, with no DOM or Node APIs, so it
 * runs unchanged in a browser, a worker, or Node.
 */
import type { Workbook } from '../model/workbook.ts';
import type { OdsPackage } from './ods/odsPackage.ts';
import { OdsReader, type OdsReadResult } from './ods/odsReader.ts';
import { OdsWriter } from './ods/odsWriter.ts';

export * from './ods/index.ts';
export * from './csv/index.ts';

/** Read `.ods` bytes into a workbook. */
export function readOds(bytes: Uint8Array): Workbook {
  return OdsReader.read(bytes);
}

/** Read `.ods` bytes, also returning the package so a later `writeOds` can preserve unmodeled content. */
export function readOdsWithPackage(bytes: Uint8Array): OdsReadResult {
  return OdsReader.readWithPackage(bytes);
}

/**
 * Write a workbook to `.ods` bytes. Pass the `package` from `readOdsWithPackage` to keep the source
 * file's styles.xml, meta.xml, unmodeled parts, and unknown spreadsheet children verbatim.
 */
export function writeOds(workbook: Workbook, source: OdsPackage | null = null): Uint8Array {
  return OdsWriter.write(workbook, source);
}
