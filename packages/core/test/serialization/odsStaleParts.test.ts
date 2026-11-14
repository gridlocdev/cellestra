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
