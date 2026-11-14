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

  // A million-cell content.xml repeats a handful of names; sharing them keeps the tree small.
  it('shares one name object per distinct name, and parsed elements stay editable', () => {
    const doc = parseXml('<a:r xmlns:a="urn:a"><a:c a:k="1"/><a:c a:k="2"/><a:c/></a:r>');
    const [first, second, third] = doc.root.elements(A('c'));
    expect(first!.name).toBe(second!.name);
    expect(first!.attributes[0]!.name).toBe(second!.attributes[0]!.name);

    third!.setAttr(A('k'), '3').add(new XmlElement(B('leaf')), 'text');
    expect(first!.attributes).toHaveLength(1);
    expect(first!.nodes).toHaveLength(0);
    expect(serializeXmlElement(third!)).toContain('a:k="3"');
    expect(third!.element(B('leaf'))!.parent).toBe(third);
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
