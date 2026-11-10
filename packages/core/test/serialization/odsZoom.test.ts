import { describe, expect, it } from 'vitest';
import { Workbook } from '../../src/model/workbook.ts';
import { OdsSupportXml, writeOds } from '../../src/serialization/index.ts';
import { serializeXml } from '../../src/xml/index.ts';
import { readEntry, roundTrip } from './helpers.ts';

const NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
  'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"';

/** settings.xml as Calc writes it: a view zoom, and a per-sheet zoom for Sheet1. */
function calcSettings(viewZoom: number, sheetZoom: number): string {
  return (
    `<office:document-settings ${NS}><office:settings>` +
    `<config:config-item-set config:name="ooo:view-settings">` +
    `<config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>` +
    `<config:config-item config:name="ViewId" config:type="string">view1</config:config-item>` +
    `<config:config-item-map-named config:name="Tables">` +
    `<config:config-item-map-entry config:name="Sheet1">` +
    `<config:config-item config:name="CursorPositionX" config:type="int">3</config:config-item>` +
    `<config:config-item config:name="ZoomValue" config:type="int">${sheetZoom}</config:config-item>` +
    `</config:config-item-map-entry></config:config-item-map-named>` +
    `<config:config-item config:name="ActiveTable" config:type="string">Sheet1</config:config-item>` +
    `<config:config-item config:name="ZoomValue" config:type="int">${viewZoom}</config:config-item>` +
    `</config:config-item-map-entry></config:config-item-map-indexed>` +
    `</config:config-item-set></office:settings></office:document-settings>`
  );
}

describe('zoom in settings.xml', () => {
  it('round-trips through a save', () => {
    const wb = Workbook.createNew();
    wb.settings.zoom = 150;
    expect(roundTrip(wb).settings.zoom).toBe(150);
    expect(roundTrip(Workbook.createNew()).settings.zoom).toBeNull();
  });

  it("reads the active sheet's zoom over the view's, and writes both, keeping the rest", () => {
    const wb = Workbook.createNew();
    OdsSupportXml.applySettings(calcSettings(90, 120), wb);
    expect(wb.settings.zoom).toBe(120);
    wb.settings.zoom = 175;
    const xml = serializeXml(OdsSupportXml.buildSettings(wb, calcSettings(90, 120)));
    expect(xml.match(/ZoomValue[^>]*>175</g)).toHaveLength(2);
    expect(xml).toContain('CursorPositionX');
    const again = Workbook.createNew();
    OdsSupportXml.applySettings(xml, again);
    expect(again.settings.zoom).toBe(175);
  });

  it('ignores a zoom out of range', () => {
    const wb = Workbook.createNew();
    OdsSupportXml.applySettings(calcSettings(5000, 1), wb);
    expect(wb.settings.zoom).toBeNull();
  });

  it('a file with no zoom gets none written', () => {
    expect(readEntry(writeOds(Workbook.createNew()), 'settings.xml')).not.toContain('ZoomValue');
  });
});
