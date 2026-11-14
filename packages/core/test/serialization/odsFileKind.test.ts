import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { OdsUnsupportedFileError, readOds } from '../../src/serialization/index.ts';
import { odsWith } from './helpers.ts';

/** Files that are not plain spreadsheets are refused with a message a person can act on. */
function zip(entries: Record<string, string>): Uint8Array {
  return zipSync(Object.fromEntries(Object.entries(entries).map(([name, text]) => [name, strToU8(text)])));
}

const CONTENT = '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>';
const manifest = (rootType: string, extra = ''): string =>
  `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="${rootType}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml">${extra}</manifest:file-entry></manifest:manifest>`;

describe('ODF file kinds', () => {
  it('refuses a text document', () => {
    const bytes = zip({
      mimetype: 'application/vnd.oasis.opendocument.text',
      'content.xml': CONTENT,
      'META-INF/manifest.xml': manifest('application/vnd.oasis.opendocument.text'),
    });
    expect(() => readOds(bytes)).toThrow(OdsUnsupportedFileError);
    expect(() => readOds(bytes)).toThrow('This is a text document, not a spreadsheet.');
  });

  it('refuses a presentation named only in the manifest', () => {
    const bytes = zip({
      'content.xml': CONTENT,
      'META-INF/manifest.xml': manifest('application/vnd.oasis.opendocument.presentation'),
    });
    expect(() => readOds(bytes)).toThrow('This is a presentation, not a spreadsheet.');
  });

  it('refuses a password-protected spreadsheet', () => {
    const bytes = zip({
      mimetype: 'application/vnd.oasis.opendocument.spreadsheet',
      'content.xml': 'not xml: encrypted bytes',
      'META-INF/manifest.xml': manifest(
        'application/vnd.oasis.opendocument.spreadsheet',
        '<manifest:encryption-data manifest:checksum-type="SHA1/1K" manifest:checksum="x"/>',
      ),
    });
    expect(() => readOds(bytes)).toThrow('This file is password-protected.');
  });

  it('refuses a zip that is not an OpenDocument file', () => {
    const bytes = zip({ '[Content_Types].xml': '<Types/>', 'xl/workbook.xml': '<workbook/>' });
    expect(() => readOds(bytes)).toThrow(OdsUnsupportedFileError);
  });

  it('refuses bytes that are not a zip package', () => {
    expect(() => readOds(strToU8('name,score\nAda,99\n'))).toThrow(OdsUnsupportedFileError);
  });

  it('opens a spreadsheet template', () => {
    const bytes = zip({
      mimetype: 'application/vnd.oasis.opendocument.spreadsheet-template',
      'content.xml': CONTENT,
    });
    expect(readOds(bytes).sheetCount).toBe(1);
  });

  it('opens a plain spreadsheet', () => {
    expect(readOds(odsWith('<table:table table:name="A"/>')).sheetCount).toBe(1);
  });
});
