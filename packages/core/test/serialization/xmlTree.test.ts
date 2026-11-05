import { describe, expect, it } from 'vitest';
import {
  parseXml,
  serializeXml,
  serializeXmlElement,
  tryParseXml,
  xmlNamespace,
  XmlElement,
} from '../../src/xml/index.ts';

const A = xmlNamespace('urn:a', 'a');
const B = xmlNamespace('urn:b', 'b');

describe('xml tree', () => {
  it('parses namespaces and finds by uri + local name, whatever the prefix', () => {
    const doc = parseXml('<x:root xmlns:x="urn:a" xmlns:y="urn:b"><y:child x:attr="1">t</y:child></x:root>');
    expect(doc.root.is(A('root'))).toBe(true);
    const child = doc.root.element(B('child'))!;
    expect(child.attr(A('attr'))).toBe('1');
    expect(child.attr('urn:a', 'attr')).toBe('1');
    expect(child.value).toBe('t');
  });

  it('serializes with declarations for the prefixes it uses', () => {
    const root = new XmlElement(A('root'), new XmlElement(B('leaf')).setAttr(A('k'), 'v & "q"'));
    const text = serializeXmlElement(root);
    expect(text).toContain('xmlns:a="urn:a"');
    expect(text).toContain('xmlns:b="urn:b"');
    expect(text).toContain('a:k="v &amp; &quot;q&quot;"');
    const reparsed = parseXml(text).root;
    expect(reparsed.element(B('leaf'))!.attr(A('k'))).toBe('v & "q"');
  });

  it('keeps mixed content and significant whitespace verbatim', () => {
    const source = '<a:p xmlns:a="urn:a">  lead <a:s>x</a:s> trail  </a:p>';
    const doc = parseXml(source, { preserveWhitespace: true });
    expect(serializeXmlElement(doc.root)).toBe(source);
    expect(doc.root.value).toBe('  lead x trail  ');
  });

  it('writes the XML declaration and indents element-only content', () => {
    const text = serializeXml(parseXml('<a:r xmlns:a="urn:a"><a:c/><a:c/></a:r>'));
    expect(text.startsWith('<?xml version="1.0" encoding="UTF-8"?>')).toBe(true);
    expect(text).toContain('\n  <a:c/>');
  });

  it('tryParseXml returns null on malformed input', () => {
    expect(tryParseXml('<a><b></a>')).toBeNull();
  });
});
