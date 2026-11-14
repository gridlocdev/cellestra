import { strFromU8, unzipSync } from 'fflate';
import type { Workbook } from '../../src/model/workbook.ts';
import {
  OdsPackage,
  OdsXmlNames,
  readOds,
  readOdsWithPackage,
  writeOds,
} from '../../src/serialization/index.ts';
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

const ODS_NAMESPACES = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"',
  'xmlns:xlink="http://www.w3.org/1999/xlink"',
  'xmlns:dc="http://purl.org/dc/elements/1.1/"',
  'xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"',
  'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"',
  'xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"',
].join(' ');

/** A hand-built `.ods` whose spreadsheet body is `tables`, with optional automatic styles and parts. */
export function odsWith(
  tables: string,
  parts: { readonly automaticStyles?: string; readonly extra?: Record<string, string> } = {},
): Uint8Array {
  const content = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${ODS_NAMESPACES} office:version="1.3"><office:automatic-styles>${parts.automaticStyles ?? ''}</office:automatic-styles><office:body><office:spreadsheet>${tables}</office:spreadsheet></office:body></office:document-content>`;
  const pkg = new OdsPackage();
  pkg.setText(OdsXmlNames.mimetypePart, OdsXmlNames.mimeType);
  pkg.setText(OdsXmlNames.contentPart, content);
  for (const [name, text] of Object.entries(parts.extra ?? {})) pkg.setText(name, text);
  return pkg.write();
}
