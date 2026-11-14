import { SaxesParser } from 'saxes';
import {
  XmlComment,
  XmlDocument,
  XmlElement,
  XmlProcessingInstruction,
  XmlText,
  xmlName,
  type XmlDeclaration,
} from './xmlTree.ts';

export interface XmlParseOptions {
  /**
   * Keep whitespace-only text nodes. Default true. With false they are dropped; text with any non-blank character is
   * always kept as-is.
   */
  readonly preserveWhitespace?: boolean;
}

export class XmlParseError extends Error {
  override readonly name = 'XmlParseError';
}

/**
 * How deeply elements may nest. Real documents stay in the tens; the parser resolves each
 * element's namespace by walking the open tags, so depth costs quadratic time, and a file of
 * fifty thousand nested spans took most of a minute to open. Past this it is refused instead.
 */
export const MAX_XML_DEPTH = 10_000;

/** Parse an XML document into a namespace-aware tree. Throws `XmlParseError` on malformed input. */
export function parseXml(text: string, options: XmlParseOptions = {}): XmlDocument {
  const preserve = options.preserveWhitespace ?? true;
  const parser = new SaxesParser({ xmlns: true, position: false });

  let declaration: XmlDeclaration | null = null;
  const prolog: Array<XmlComment | XmlProcessingInstruction> = [];
  const stack: XmlElement[] = [];
  let root: XmlElement | null = null;
  let error: Error | null = null;

  const addText = (value: string, cdata: boolean): void => {
    const parent = stack.at(-1);
    if (!parent) return; // whitespace outside the root
    if (!cdata && !preserve && value.trim().length === 0) return;
    const last = parent.nodes.at(-1);
    if (!cdata && last instanceof XmlText && !last.cdata) last.value += value;
    else parent.nodes.push(new XmlText(value, cdata));
  };

  parser.on('xmldecl', (d) => {
    declaration = { version: d.version ?? '1.0', encoding: d.encoding, standalone: d.standalone };
  });
  parser.on('text', (t) => addText(t, false));
  parser.on('cdata', (t) => addText(t, true));
  parser.on('comment', (c) => {
    const node = new XmlComment(c);
    const parent = stack.at(-1);
    if (parent) parent.nodes.push(node);
    else if (!root) prolog.push(node);
  });
  parser.on('processinginstruction', (pi) => {
    const node = new XmlProcessingInstruction(pi.target, pi.body);
    const parent = stack.at(-1);
    if (parent) parent.nodes.push(node);
    else if (!root) prolog.push(node);
  });
  parser.on('opentag', (tag) => {
    if (stack.length >= MAX_XML_DEPTH) {
      throw new XmlParseError(`Elements are nested more than ${MAX_XML_DEPTH} deep.`);
    }
    const element = new XmlElement(xmlName(tag.local, tag.uri, tag.prefix));
    for (const a of Object.values(tag.attributes)) {
      element.attributes.push({ name: xmlName(a.local, a.uri, a.prefix), value: a.value });
    }
    const parent = stack.at(-1);
    if (parent) parent.add(element);
    else root = element;
    stack.push(element);
  });
  parser.on('closetag', () => {
    stack.pop();
  });
  parser.on('error', (e) => {
    error ??= e;
  });

  try {
    parser.write(text).close();
  } catch (e) {
    error ??= e instanceof Error ? e : new Error(String(e));
  }
  if (error) throw new XmlParseError((error as Error).message);
  if (!root) throw new XmlParseError('The document has no root element.');

  const doc = new XmlDocument(root, declaration);
  doc.prolog.push(...prolog);
  return doc;
}

/** Parse, or null when the text is not well-formed XML. */
export function tryParseXml(text: string, options: XmlParseOptions = {}): XmlDocument | null {
  try {
    return parseXml(text, options);
  } catch (e) {
    if (e instanceof XmlParseError) return null;
    throw e;
  }
}

/**
 * Parse a fragment that was serialized standalone (e.g. a preserved `office:spreadsheet` child). The
 * serializer declares every namespace a fragment uses on its root, so it parses as a document.
 */
export function parseXmlElement(text: string, options: XmlParseOptions = {}): XmlElement {
  return parseXml(text, options).root;
}
