import { tryParseXml } from '../../xml/index.ts';
import { OdsUnsupportedFileError, type OdsPackage } from './odsPackage.ts';
import { Manifest, OdsXmlNames as N } from './odsXmlNames.ts';

const ODF_PREFIX = 'application/vnd.oasis.opendocument.';
const SPREADSHEET_TYPES = new Set([N.mimeType, `${N.mimeType}-template`]);

/** What to call each other ODF document type, by the part after the ODF media-type prefix. */
const KIND_NAMES: Readonly<Record<string, string>> = {
  text: 'a text document',
  'text-template': 'a text document template',
  'text-master': 'a master document',
  'text-web': 'a web page',
  presentation: 'a presentation',
  'presentation-template': 'a presentation template',
  graphics: 'a drawing',
  'graphics-template': 'a drawing template',
  chart: 'a chart',
  formula: 'a formula document',
  image: 'an image',
  base: 'a database',
};

/**
 * Refuse a package that is not a plain spreadsheet before reading it: a password-protected file
 * (its content.xml is ciphertext) or another kind of document. Opening those as an empty sheet
 * would let a save overwrite the real document. The type comes from the `mimetype` entry, else
 * the manifest's root entry; a package naming neither must at least have a content.xml.
 */
export function checkSpreadsheetPackage(pkg: OdsPackage): void {
  const manifestXml = pkg.getText(N.manifestPart);
  const manifest = manifestXml === null ? null : tryParseXml(manifestXml);
  const entries = manifest?.root.descendants(Manifest('file-entry')) ?? [];
  if (entries.some((entry) => entry.descendants(Manifest('encryption-data')).length > 0)) {
    throw new OdsUnsupportedFileError(
      'This file is password-protected. Save it without a password in LibreOffice, then open it here.',
    );
  }

  const rootEntry = entries.find((entry) => entry.attr(Manifest('full-path')) === '/');
  const type = pkg.getText(N.mimetypePart)?.trim() || rootEntry?.attr(Manifest('media-type')) || null;
  if (type === null) {
    if (pkg.contains(N.contentPart)) return;
    throw new OdsUnsupportedFileError('This file is not an OpenDocument spreadsheet.');
  }
  if (SPREADSHEET_TYPES.has(type)) return;
  const kind = type.startsWith(ODF_PREFIX) ? KIND_NAMES[type.slice(ODF_PREFIX.length)] : undefined;
  throw new OdsUnsupportedFileError(
    kind === undefined
      ? `This file is not a spreadsheet (its type is ${type}).`
      : `This is ${kind}, not a spreadsheet.`,
  );
}
