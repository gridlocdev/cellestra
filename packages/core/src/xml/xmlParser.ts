import { XmlEventReader, XmlParseError, XmlTreeBuilder } from './xmlStream.ts';
import type { XmlDocument, XmlElement } from './xmlTree.ts';

export interface XmlParseOptions {
  /**
   * Keep whitespace-only text nodes. Default true. With false they are dropped; text with any non-blank character is
   * always kept as-is.
   */
  readonly preserveWhitespace?: boolean;
}

/** Parse an XML document into a namespace-aware tree. Throws `XmlParseError` on malformed input. */
export function parseXml(text: string, options: XmlParseOptions = {}): XmlDocument {
  const builder = new XmlTreeBuilder(options.preserveWhitespace ?? true);
  const reader = new XmlEventReader(builder);
  reader.write(text);
  reader.close();
  return builder.document();
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
