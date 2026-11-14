import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { Workbook } from '../../src/model/workbook.ts';
import {
  OdsPackage,
  OdsPackageTooLargeError,
  OdsUnsupportedFileError,
  OdsXmlNames,
  readContentXml,
  readOds,
  writeOds,
} from '../../src/serialization/index.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { odsWith, readEntry } from './helpers.ts';

const TABLES = `<table:table table:name="Sheet1" table:protected="true">
    <table:table-column table:number-columns-repeated="2"/>
    <table:table-row><table:table-cell office:value-type="float" office:value="5"><office:annotation><text:p>Note</text:p></office:annotation><text:p>5</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>a <text:span>rich</text:span> é😀</text:p></table:table-cell></table:table-row>
    <table:table-row table:number-rows-repeated="3"><table:table-cell table:number-columns-spanned="2" table:formula="of:=1+[.A1]" office:value-type="float" office:value="6"><text:p>6</text:p></table:table-cell><table:covered-table-cell/></table:table-row>
    <table:table-row table:number-rows-repeated="1048570"><table:table-cell table:number-columns-repeated="1024"/></table:table-row>
  </table:table>
  <table:named-expressions><table:named-range table:name="Five" table:cell-range-address="$Sheet1.$A$1"/></table:named-expressions>`;

/** The bytes of a package whose content.xml is stored (method 0) or deflated (method 8). */
function packageWith(content: string, level: 0 | 6): Uint8Array {
  return zipSync({
    [OdsXmlNames.mimetypePart]: [strToU8(OdsXmlNames.mimeType), { level: 0 }],
    [OdsXmlNames.contentPart]: [strToU8(content), { level }],
  });
}

describe('streamed .ods reading', () => {
  it('reads content.xml the same in pieces of any size', () => {
    const bytes = odsWith(TABLES);
    const content = readEntry(bytes, OdsXmlNames.contentPart);
    const saved = (chunks: Iterable<string>): string => {
      const workbook = new Workbook();
      const pkg = OdsPackage.read(bytes);
      readContentXml(chunks, workbook, pkg.foreignSpreadsheetChildren, pkg);
      return readEntry(writeOds(workbook, pkg), OdsXmlNames.contentPart);
    };
    const whole = saved([content]);
    expect(whole).toContain('office:annotation');
    for (const size of [1, 5, 64]) {
      const pieces: string[] = [];
      for (let i = 0; i < content.length; i += size) pieces.push(content.slice(i, i + size));
      expect(saved(pieces)).toBe(whole);
    }
  });

  it('inflates content.xml a piece at a time, stored or deflated, with or without a BOM', () => {
    // Hex digits compress only to about half, so the deflated part is many pieces of input.
    let seed = 1;
    const digits = Array.from({ length: 300_000 }, () => {
      seed ^= seed << 13;
      seed ^= seed >>> 17;
      seed ^= seed << 5;
      return (seed & 15).toString(16);
    });
    const big = `<table:table table:name="Big"><table:table-row><table:table-cell><text:p>${digits.join('')}</text:p></table:table-cell></table:table-row></table:table>`;
    const content = readEntry(odsWith(TABLES + big), 'content.xml');
    for (const level of [0, 6] as const) {
      const pkg = OdsPackage.read(packageWith('﻿' + content, level));
      const chunks = [...pkg.textChunks(OdsXmlNames.contentPart)!];
      expect(chunks.length).toBeGreaterThan(1);
      expect(chunks.join('')).toBe(content);
      expect(pkg.getText(OdsXmlNames.contentPart)).toBe(content);
      const workbook = readOds(packageWith(content, level));
      expect(workbook.sheetCount).toBe(2);
      expect(workbook.worksheets[0]!.getValue(3, 0)).toEqual(CellValue.number(6));
    }
  });

  it('refuses damaged content.xml data as a damaged file', () => {
    const bytes = packageWith(readEntry(odsWith(TABLES), 'content.xml'), 6);
    const pkg = OdsPackage.read(bytes);
    const start = findSubarray(bytes, strToU8(OdsXmlNames.contentPart)) + OdsXmlNames.contentPart.length;
    bytes.fill(0xff, start + 10, start + 60);
    expect(() => readOds(bytes)).toThrow(OdsUnsupportedFileError);
    expect(pkg.contains(OdsXmlNames.contentPart)).toBe(true);
  });

  it('stops inflating content.xml at the package cap, whatever size it declares', () => {
    const bytes = packageWith(`<a>${'x'.repeat(100_000)}</a>`, 6);
    // Declare a tiny uncompressed size in the central directory, as a zip bomb would.
    const name = strToU8(OdsXmlNames.contentPart);
    const entry = findSubarray(bytes, name, findSubarray(bytes, name) + 1) - 46;
    const view = new DataView(bytes.buffer);
    view.setUint32(entry + 24, 100, true);
    const pkg = OdsPackage.read(bytes, 50_000);
    expect(() => [...pkg.textChunks(OdsXmlNames.contentPart)!]).toThrow(OdsPackageTooLargeError);
  });
});

function findSubarray(haystack: Uint8Array, needle: Uint8Array, from = 0): number {
  outer: for (let i = from; i <= haystack.length - needle.length; i++) {
    for (let j = 0; j < needle.length; j++) if (haystack[i + j] !== needle[j]) continue outer;
    return i;
  }
  return -1;
}
