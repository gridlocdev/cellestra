import { describe, expect, it } from 'vitest';
import { readOdsWithPackage, writeOds } from '../../src/serialization/index.ts';
import { entryNames, odsWith, readEntry } from './helpers.ts';

/** Parts that describe the old content are dropped when a save rewrites it. */
describe('stale parts on save', () => {
  it('drops the document signature and the thumbnail', () => {
    const source = odsWith('<table:table table:name="A"/>', {
      extra: {
        'META-INF/documentsignatures.xml': '<document-signatures/>',
        'Thumbnails/thumbnail.png': 'png bytes',
        'Pictures/logo.png': 'kept',
      },
    });
    const { workbook, package: pkg } = readOdsWithPackage(source);
    const saved = writeOds(workbook, pkg);

    const names = entryNames(saved);
    expect(names).not.toContain('META-INF/documentsignatures.xml');
    expect(names).not.toContain('Thumbnails/thumbnail.png');
    expect(names).toContain('Pictures/logo.png');
    const manifest = readEntry(saved, 'META-INF/manifest.xml');
    expect(manifest).not.toContain('documentsignatures');
    expect(manifest).not.toContain('Thumbnails');
  });
});

describe('manifest on save', () => {
  it('keeps the media type the source manifest gave each part', () => {
    const manifest = `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="manifest.rdf" manifest:media-type="application/rdf+xml"/><manifest:file-entry manifest:full-path="Pictures/logo.svg" manifest:media-type="image/svg+xml"/><manifest:file-entry manifest:full-path="Basic/Standard/Module1.xml" manifest:media-type="text/xml"/></manifest:manifest>`;
    const source = odsWith('<table:table table:name="A"/>', {
      extra: {
        'META-INF/manifest.xml': manifest,
        'manifest.rdf': '<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"/>',
        'Pictures/logo.svg': '<svg/>',
        'Basic/Standard/Module1.xml': '<module/>',
      },
    });
    const { workbook, package: pkg } = readOdsWithPackage(source);
    const saved = readEntry(writeOds(workbook, pkg), 'META-INF/manifest.xml');

    expect(saved).toContain('manifest:full-path="manifest.rdf" manifest:media-type="application/rdf+xml"');
    expect(saved).toContain('manifest:full-path="Pictures/logo.svg" manifest:media-type="image/svg+xml"');
    expect(saved).toContain('manifest:full-path="Basic/Standard/Module1.xml" manifest:media-type="text/xml"');
  });
});
