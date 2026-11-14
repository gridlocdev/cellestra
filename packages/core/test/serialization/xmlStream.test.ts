import { describe, expect, it } from 'vitest';
import {
  parseXml,
  serializeXml,
  XmlEventReader,
  XmlParseError,
  XmlTreeBuilder,
  type XmlEventHandler,
} from '../../src/xml/index.ts';

const SAMPLE =
  '<?xml version="1.0" encoding="UTF-8"?><!-- c --><a:r xmlns:a="urn:a" xmlns:b="urn:b">' +
  '<a:c a:k="1">text &amp; more</a:c>\n  <b:d><![CDATA[x<y]]></b:d><?pi body?><a:e>é😀 ü</a:e></a:r>';

function readInChunks(text: string, size: number): string {
  const builder = new XmlTreeBuilder(true);
  const reader = new XmlEventReader(builder);
  for (let i = 0; i < text.length; i += size) reader.write(text.slice(i, i + size));
  reader.close();
  return serializeXml(builder.document());
}

const ignore: XmlEventHandler = {
  declaration() {},
  openElement() {},
  closeElement() {},
  text() {},
  comment() {},
  processingInstruction() {},
};

describe('xml event reader', () => {
  it('builds the same tree whatever size the chunks are', () => {
    const whole = serializeXml(parseXml(SAMPLE));
    for (const size of [1, 2, 3, 7, 64]) expect(readInChunks(SAMPLE, size)).toBe(whole);
  });

  it('throws XmlParseError for malformed or empty input', () => {
    expect(() => readInChunks('<a><b></a>', 3)).toThrow(XmlParseError);
    expect(() => readInChunks('<a>', 1)).toThrow(XmlParseError);
    expect(() => readInChunks('', 1)).toThrow(XmlParseError);
  });

  it("lets the handler's own exception through unchanged", () => {
    class Stop extends Error {}
    const reader = new XmlEventReader({
      ...ignore,
      openElement: () => {
        throw new Stop('stop');
      },
    });
    expect(() => reader.write('<a/>')).toThrow(Stop);
  });
});
