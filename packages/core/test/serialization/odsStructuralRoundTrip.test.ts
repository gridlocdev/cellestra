import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { readOdsWithPackage, writeOds } from '../../src/serialization/index.ts';
import { tryParseXml, type XmlElement } from '../../src/xml/index.ts';
import { readEntry } from './helpers.ts';

/**
 * Files LibreOffice saved, opened and saved again by Cellestra without an edit. Every element and
 * attribute name the original has must still be in the saved file, at the same place in the tree:
 * content Cellestra neither models nor keeps verbatim disappears from this comparison first.
 * scripts/make-libreoffice-sample.py and scripts/make-libreoffice-outline.py build the fixtures.
 */
const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(new URL(`./fixtures/${name}`, import.meta.url)));

const PARTS = ['content.xml', 'styles.xml'];

/**
 * What a save may change on purpose. Names are `prefix:local`, with LibreOffice's prefixes.
 * - `calcext:value-type` repeats `office:value-type`; `office:scripts` empty is no content.
 * - Default cell styles named "Default" are the default, so they are not written, on cells
 *   (`covered-table-cell@table:style-name` too) or as a line's default cell style.
 * - `style:use-optimal-row-height` and `-column-width` say nothing beyond the default, and a row
 *   of the default, optimal height needs no row style at all.
 * - An empty `table:named-expressions` holds nothing.
 * - A formula's text result is written as its paragraph, not also as `office:string-value`.
 * - `number:min-decimal-places` repeats `number:decimal-places` (LibreOffice's extension for
 *   formats that drop trailing zeros, which Cellestra's formats do not do).
 * - `number:boolean-style` is TRUE/FALSE, which a boolean cell shows without it.
 * - `table:number-rows-spanned="1"` beside a column span is no row span.
 */
const EXPECTED_CHANGES: readonly RegExp[] = [
  /@calcext:value-type$/,
  /@table:default-cell-style-name$/,
  /covered-table-cell@table:style-name$/,
  /^content\.xml office:document-content\/office:scripts$/,
  /\/table:named-expressions$/,
  /@style:use-optimal-row-height$/,
  /@style:use-optimal-column-width$/,
  /table:table-row@table:style-name$/,
  /@office:string-value$/,
  /@number:min-decimal-places$/,
  /\/number:boolean-style(\/number:boolean|@style:name)?$/,
  /@table:number-rows-spanned$/,
];

/**
 * Changes one fixture may show that would hide a loss in another.
 * - The sample's rows and columns all say `fo:break-before="auto"`, the default; the outline
 *   fixture's manual breaks must still come back.
 */
const EXPECTED_IN: Readonly<Record<string, readonly RegExp[]>> = {
  'libreoffice-sample.ods': [/@fo:break-before$/],
};

/** Every element path and `path@attribute` in a part, with prefixes taken from namespace URIs. */
function shape(part: string, xml: string): Set<string> {
  const doc = tryParseXml(xml, { preserveWhitespace: true });
  if (!doc) throw new Error(`${part} does not parse`);
  const out = new Set<string>();
  const walk = (e: XmlElement, parent: string): void => {
    const path = parent === '' ? name(e.name) : `${parent}/${name(e.name)}`;
    out.add(`${part} ${path}`);
    for (const a of e.attributes)
      if (!isNamespaceDeclaration(a.name)) out.add(`${part} ${path}@${name(a.name)}`);
    for (const child of e.elements()) walk(child, path);
  };
  walk(doc.root, '');
  return out;
}

const XMLNS = 'http://www.w3.org/2000/xmlns/';
const isNamespaceDeclaration = (n: { prefix: string; uri: string; local: string }): boolean =>
  n.uri === XMLNS || n.prefix === 'xmlns' || (n.prefix === '' && n.local === 'xmlns');

/** LibreOffice's prefix for a namespace, so files that pick other prefixes still compare. */
const PREFIXES = new Map<string, string>([
  ['urn:oasis:names:tc:opendocument:xmlns:office:1.0', 'office'],
  ['urn:oasis:names:tc:opendocument:xmlns:style:1.0', 'style'],
  ['urn:oasis:names:tc:opendocument:xmlns:text:1.0', 'text'],
  ['urn:oasis:names:tc:opendocument:xmlns:table:1.0', 'table'],
  ['urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0', 'fo'],
  ['urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0', 'number'],
  ['urn:oasis:names:tc:opendocument:xmlns:script:1.0', 'script'],
  ['http://openoffice.org/2009/office', 'officeooo'],
  ['urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0', 'calcext'],
  ['urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0', 'loext'],
]);

function name(n: { prefix: string; uri: string; local: string }): string {
  return `${PREFIXES.get(n.uri) ?? n.prefix}:${n.local}`;
}

/** What the original has that a save without edits lost, less the changes made on purpose. */
function lostOnSave(bytes: Uint8Array, expected: readonly RegExp[]): string[] {
  const { workbook, package: pkg } = readOdsWithPackage(bytes);
  const saved = writeOds(workbook, pkg);
  const lost: string[] = [];
  for (const part of PARTS) {
    const before = shape(part, readEntry(bytes, part));
    const after = shape(part, readEntry(saved, part));
    for (const entry of before) {
      if (!after.has(entry) && !expected.some((r) => r.test(entry))) lost.push(entry);
    }
  }
  return lost.sort();
}

describe('A LibreOffice file saved without edits keeps its structure', () => {
  it.each(['libreoffice-sample.ods', 'libreoffice-outline.ods'])('%s', (file) => {
    expect(lostOnSave(fixture(file), [...EXPECTED_CHANGES, ...(EXPECTED_IN[file] ?? [])])).toEqual([]);
  });
});
