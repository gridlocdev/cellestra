import { SaxesParser } from 'saxes';
import {
  XmlComment,
  XmlDocument,
  XmlElement,
  XmlProcessingInstruction,
  XmlText,
  xmlName,
  type XmlAttribute,
  type XmlDeclaration,
  type XmlName,
} from './xmlTree.ts';

export class XmlParseError extends Error {
  override readonly name = 'XmlParseError';
}

/**
 * How deeply elements may nest. Real documents stay in the tens; the parser resolves each
 * element's namespace by walking the open tags, so depth costs quadratic time, and a file of
 * fifty thousand nested spans took most of a minute to open. Past this it is refused instead.
 */
export const MAX_XML_DEPTH = 10_000;

/** What an `XmlEventReader` reports, in document order. */
export interface XmlEventHandler {
  declaration(declaration: XmlDeclaration): void;
  /** `attributes` is a fresh array the handler may keep. */
  openElement(name: XmlName, attributes: XmlAttribute[]): void;
  closeElement(): void;
  text(value: string, cdata: boolean): void;
  comment(value: string): void;
  processingInstruction(target: string, body: string): void;
}

/**
 * Reads XML text, written in one piece or in chunks, as a stream of namespace-aware events, so a
 * caller can build only the parts of a large document it needs. Names are shared: every element
 * or attribute with the same (uri, prefix, local) gets the same frozen `XmlName`.
 *
 * Malformed input throws `XmlParseError` from `write` or `close` at the first error. An exception
 * thrown by the handler passes through unchanged and stops the reader.
 */
export class XmlEventReader {
  private readonly parser = new SaxesParser({ xmlns: true, position: false });
  private readonly names = new Map<string, XmlName>();
  private depth = 0;
  private sawRoot = false;
  /** The handler's own exception, told apart from the parser's so it is not wrapped. */
  private handlerError: unknown = undefined;

  constructor(handler: XmlEventHandler) {
    const parser = this.parser;
    const call = (f: () => void): void => {
      try {
        f();
      } catch (e) {
        this.handlerError = e;
        throw e;
      }
    };
    parser.on('xmldecl', (d) =>
      call(() =>
        handler.declaration({ version: d.version ?? '1.0', encoding: d.encoding, standalone: d.standalone }),
      ),
    );
    parser.on('text', (t) => call(() => handler.text(t, false)));
    parser.on('cdata', (t) => call(() => handler.text(t, true)));
    parser.on('comment', (c) => call(() => handler.comment(c)));
    parser.on('processinginstruction', (pi) => call(() => handler.processingInstruction(pi.target, pi.body)));
    parser.on('opentag', (tag) => {
      if (this.depth >= MAX_XML_DEPTH) {
        throw new XmlParseError(`Elements are nested more than ${MAX_XML_DEPTH} deep.`);
      }
      this.depth++;
      this.sawRoot = true;
      const attributes = Object.values(tag.attributes).map((a): XmlAttribute => ({
        name: this.nameOf(a.local, a.uri, a.prefix),
        value: a.value,
      }));
      const name = this.nameOf(tag.local, tag.uri, tag.prefix);
      call(() => handler.openElement(name, attributes));
    });
    parser.on('closetag', () => {
      this.depth--;
      call(() => handler.closeElement());
    });
    parser.on('error', (e) => {
      throw new XmlParseError(e.message);
    });
  }

  /** Feed the next piece of the document. */
  write(chunk: string): this {
    this.guard(() => this.parser.write(chunk));
    return this;
  }

  /** The document is complete: throws when it was cut short or had no root element. */
  close(): void {
    this.guard(() => this.parser.close());
    if (!this.sawRoot) throw new XmlParseError('The document has no root element.');
  }

  private guard(f: () => void): void {
    try {
      f();
    } catch (e) {
      if (e === this.handlerError || e instanceof XmlParseError) throw e;
      throw new XmlParseError(e instanceof Error ? e.message : String(e));
    }
  }

  // One shared name per distinct (uri, prefix, local): a million cells repeat a handful of names,
  // and a fresh name object (with its own strings) per element and attribute was a large share of
  // a parsed tree's memory.
  private nameOf(local: string, uri: string, prefix: string): XmlName {
    const key = `${prefix}\u0000${local}\u0000${uri}`;
    let name = this.names.get(key);
    if (!name) {
      name = Object.freeze(xmlName(local, uri, prefix));
      this.names.set(key, name);
    }
    return name;
  }
}

/**
 * Append text to `parent` as the parser keeps it: adjacent plain text joins one node, and with
 * `preserveWhitespace` false a whitespace-only run is dropped (CDATA is always kept).
 */
export function appendXmlText(
  parent: XmlElement,
  value: string,
  cdata: boolean,
  preserveWhitespace: boolean,
): void {
  if (!cdata && !preserveWhitespace && value.trim().length === 0) return;
  const last = parent.nodes.at(-1);
  if (!cdata && last instanceof XmlText && !last.cdata) last.value += value;
  else parent.nodes.push(new XmlText(value, cdata));
}

/** Open an element under `parent` (or detached, with null) with the attributes the reader gave. */
export function openXmlElement(
  parent: XmlElement | null,
  name: XmlName,
  attributes: XmlAttribute[],
): XmlElement {
  const element = new XmlElement(name);
  element.attributes = attributes;
  if (parent) parent.add(element);
  return element;
}

/** Builds the whole document tree from reader events. */
export class XmlTreeBuilder implements XmlEventHandler {
  private readonly preserveWhitespace: boolean;
  private readonly stack: XmlElement[] = [];
  private readonly prolog: Array<XmlComment | XmlProcessingInstruction> = [];
  private xmlDeclaration: XmlDeclaration | null = null;
  private root: XmlElement | null = null;

  constructor(preserveWhitespace: boolean) {
    this.preserveWhitespace = preserveWhitespace;
  }

  /** The finished document; throws when there was no root element. */
  document(): XmlDocument {
    if (!this.root) throw new XmlParseError('The document has no root element.');
    const doc = new XmlDocument(this.root, this.xmlDeclaration);
    doc.prolog.push(...this.prolog);
    return doc;
  }

  declaration(declaration: XmlDeclaration): void {
    this.xmlDeclaration = declaration;
  }

  openElement(name: XmlName, attributes: XmlAttribute[]): void {
    const element = openXmlElement(this.stack.at(-1) ?? null, name, attributes);
    if (this.stack.length === 0) this.root = element;
    this.stack.push(element);
  }

  closeElement(): void {
    this.stack.pop()?.trimToSize();
  }

  text(value: string, cdata: boolean): void {
    const parent = this.stack.at(-1);
    if (parent) appendXmlText(parent, value, cdata, this.preserveWhitespace); // else whitespace outside the root
  }

  comment(value: string): void {
    this.addMisc(new XmlComment(value));
  }

  processingInstruction(target: string, body: string): void {
    this.addMisc(new XmlProcessingInstruction(target, body));
  }

  private addMisc(node: XmlComment | XmlProcessingInstruction): void {
    const parent = this.stack.at(-1);
    if (parent) parent.nodes.push(node);
    else if (!this.root) this.prolog.push(node);
  }
}
