/**
 * A small namespace-aware XML tree that the `.ods` reader and writer are built on. Every element and attribute keeps its namespace URI, its
 * prefix, and its local name; namespace declarations (`xmlns:*`) are ordinary attributes in the
 * reserved xmlns namespace, kept in document order. That is what lets an unmodeled element read from
 * a file be written back verbatim — prefixes, attribute order, whitespace, comments, and all.
 *
 * Queries match on namespace URI + local name; the prefix is presentation only.
 */

export const XML_NS_URI = 'http://www.w3.org/XML/1998/namespace';
export const XMLNS_URI = 'http://www.w3.org/2000/xmlns/';

/** A qualified name: namespace URI + local name, plus the prefix to write it with. */
export interface XmlName {
  readonly uri: string;
  readonly prefix: string;
  readonly local: string;
}

/**
 * A namespace you can call to make names in it: `Table('name')` is `table:name`.
 */
export interface XmlNamespace {
  (local: string): XmlName;
  readonly uri: string;
  readonly prefix: string;
}

export function xmlNamespace(uri: string, prefix: string): XmlNamespace {
  const cache = new Map<string, XmlName>();
  const make = (local: string): XmlName => {
    let name = cache.get(local);
    if (!name) {
      name = Object.freeze({ uri, prefix, local });
      cache.set(local, name);
    }
    return name;
  };
  return Object.assign(make, { uri, prefix }) as XmlNamespace;
}

/** A name in no namespace (an unprefixed attribute). */
export function xmlName(local: string, uri = '', prefix = ''): XmlName {
  return { uri, prefix, local };
}

export function sameName(a: XmlName, b: XmlName): boolean {
  return a.local === b.local && a.uri === b.uri;
}

export function qualifiedName(name: XmlName): string {
  return name.prefix ? `${name.prefix}:${name.local}` : name.local;
}

export interface XmlAttribute {
  readonly name: XmlName;
  value: string;
}

/** True for an `xmlns` / `xmlns:p` declaration attribute. */
export function isNamespaceDeclaration(a: XmlAttribute): boolean {
  return a.name.uri === XMLNS_URI;
}

export class XmlText {
  readonly kind = 'text';
  value: string;
  /** Written back as `<![CDATA[…]]>` when it was one. */
  readonly cdata: boolean;
  constructor(value: string, cdata = false) {
    this.value = value;
    this.cdata = cdata;
  }
}

export class XmlComment {
  readonly kind = 'comment';
  readonly value: string;
  constructor(value: string) {
    this.value = value;
  }
}

export class XmlProcessingInstruction {
  readonly kind = 'pi';
  readonly target: string;
  readonly body: string;
  constructor(target: string, body: string) {
    this.target = target;
    this.body = body;
  }
}

export type XmlNode = XmlElement | XmlText | XmlComment | XmlProcessingInstruction;

/** Something `XmlElement.add` accepts: a node, an attribute, plain text, or nothing (skipped). */
export type XmlContent = XmlNode | XmlAttribute | string | number | null | undefined | false;

export type AttrValue = string | number | boolean;

export class XmlElement {
  readonly kind = 'element';
  readonly name: XmlName;
  readonly attributes: XmlAttribute[] = [];
  readonly nodes: XmlNode[] = [];
  parent: XmlElement | null = null;

  constructor(name: XmlName, ...content: XmlContent[]) {
    this.name = name;
    this.add(...content);
  }

  get uri(): string {
    return this.name.uri;
  }
  get local(): string {
    return this.name.local;
  }

  is(name: XmlName): boolean {
    return sameName(this.name, name);
  }

  // ---- Attributes --------------------------------------------------------

  attribute(name: XmlName): XmlAttribute | undefined {
    return this.attributes.find((a) => sameName(a.name, name));
  }

  /** The attribute's value, or null when absent. Also accepts (uri, local). */
  attr(name: XmlName): string | null;
  attr(uri: string, local: string): string | null;
  attr(nameOrUri: XmlName | string, local?: string): string | null {
    const name = typeof nameOrUri === 'string' ? xmlName(local!, nameOrUri) : nameOrUri;
    return this.attribute(name)?.value ?? null;
  }

  /** Set (or with null, remove) an attribute; new attributes go last, existing ones stay in place. */
  setAttr(name: XmlName, value: AttrValue | null | undefined): this {
    const i = this.attributes.findIndex((a) => sameName(a.name, name));
    if (value === null || value === undefined) {
      if (i >= 0) this.attributes.splice(i, 1);
      return this;
    }
    const text = String(value);
    if (i >= 0) this.attributes[i]!.value = text;
    else this.attributes.push({ name, value: text });
    return this;
  }

  /** Declare a namespace prefix on this element (`xmlns:prefix="uri"`). */
  declare(prefix: string, uri: string): this {
    const name = prefix ? xmlName(prefix, XMLNS_URI, 'xmlns') : xmlName('xmlns', XMLNS_URI);
    return this.setAttr(name, uri);
  }

  /** Attributes other than namespace declarations. */
  get hasAttributes(): boolean {
    return this.attributes.some((a) => !isNamespaceDeclaration(a));
  }

  // ---- Children ----------------------------------------------------------

  add(...content: XmlContent[]): this {
    for (const item of content) {
      if (item === null || item === undefined || item === false) continue;
      if (typeof item === 'string' || typeof item === 'number') {
        const text = String(item);
        const last = this.nodes.at(-1);
        if (last instanceof XmlText && !last.cdata) last.value += text;
        else this.nodes.push(new XmlText(text));
      } else if (item instanceof XmlElement) {
        item.parent = this;
        this.nodes.push(item);
      } else if (
        item instanceof XmlText ||
        item instanceof XmlComment ||
        item instanceof XmlProcessingInstruction
      ) {
        this.nodes.push(item);
      } else {
        this.setAttr(item.name, item.value);
      }
    }
    return this;
  }

  addFirst(node: XmlNode): this {
    if (node instanceof XmlElement) node.parent = this;
    this.nodes.unshift(node);
    return this;
  }

  remove(node: XmlNode): boolean {
    const i = this.nodes.indexOf(node);
    if (i < 0) return false;
    this.nodes.splice(i, 1);
    if (node instanceof XmlElement) node.parent = null;
    return true;
  }

  /** Child elements, optionally only those with `name`. */
  elements(name?: XmlName): XmlElement[] {
    const out: XmlElement[] = [];
    for (const n of this.nodes) if (n instanceof XmlElement && (!name || n.is(name))) out.push(n);
    return out;
  }

  /** The first child element named `name`, or null. */
  element(name: XmlName): XmlElement | null {
    for (const n of this.nodes) if (n instanceof XmlElement && n.is(name)) return n;
    return null;
  }

  /** Alias of `element`, for readers that prefer `child`. */
  child(name: XmlName): XmlElement | null {
    return this.element(name);
  }

  /** Alias of `elements`. */
  children(name?: XmlName): XmlElement[] {
    return this.elements(name);
  }

  get hasElements(): boolean {
    return this.nodes.some((n) => n instanceof XmlElement);
  }

  /**
   * Every descendant element in document order, optionally only those with `name`. Iterative, so a
   * hostile file nested a million levels deep cannot overflow the call stack.
   */
  descendants(name?: XmlName): XmlElement[] {
    const out: XmlElement[] = [];
    const stack: Array<{ e: XmlElement; i: number }> = [{ e: this, i: 0 }];
    while (stack.length > 0) {
      const top = stack[stack.length - 1]!;
      if (top.i >= top.e.nodes.length) {
        stack.pop();
        continue;
      }
      const n = top.e.nodes[top.i++]!;
      if (!(n instanceof XmlElement)) continue;
      if (!name || n.is(name)) out.push(n);
      stack.push({ e: n, i: 0 });
    }
    return out;
  }

  /** The concatenated text of every descendant text node. Iterative. */
  get value(): string {
    let s = '';
    const stack: Array<{ e: XmlElement; i: number }> = [{ e: this, i: 0 }];
    while (stack.length > 0) {
      const top = stack[stack.length - 1]!;
      if (top.i >= top.e.nodes.length) {
        stack.pop();
        continue;
      }
      const n = top.e.nodes[top.i++]!;
      if (n instanceof XmlText) s += n.value;
      else if (n instanceof XmlElement) stack.push({ e: n, i: 0 });
    }
    return s;
  }

  /** Replace all child nodes with a single text node (empty string → no children). */
  set value(text: string) {
    for (const n of this.nodes) if (n instanceof XmlElement) n.parent = null;
    this.nodes.length = 0;
    if (text.length > 0) this.nodes.push(new XmlText(text));
  }

  /** A deep copy, detached from any parent. Iterative, so depth is bounded only by memory. */
  clone(): XmlElement {
    const shallow = (e: XmlElement): XmlElement => {
      const copy = new XmlElement(e.name);
      for (const a of e.attributes) copy.attributes.push({ name: a.name, value: a.value });
      return copy;
    };
    const root = shallow(this);
    const stack: Array<{ from: XmlElement; to: XmlElement; i: number }> = [{ from: this, to: root, i: 0 }];
    while (stack.length > 0) {
      const top = stack[stack.length - 1]!;
      if (top.i >= top.from.nodes.length) {
        stack.pop();
        continue;
      }
      const n = top.from.nodes[top.i++]!;
      if (n instanceof XmlElement) {
        const copy = shallow(n);
        copy.parent = top.to;
        top.to.nodes.push(copy);
        stack.push({ from: n, to: copy, i: 0 });
      } else if (n instanceof XmlText) top.to.nodes.push(new XmlText(n.value, n.cdata));
      else top.to.nodes.push(n);
    }
    return root;
  }
}

export interface XmlDeclaration {
  readonly version: string;
  readonly encoding?: string | undefined;
  readonly standalone?: string | undefined;
}

export class XmlDocument {
  declaration: XmlDeclaration | null;
  /** Comments / processing instructions before the root element. */
  readonly prolog: Array<XmlComment | XmlProcessingInstruction> = [];
  root: XmlElement;

  constructor(root: XmlElement, declaration: XmlDeclaration | null = { version: '1.0', encoding: 'UTF-8' }) {
    this.root = root;
    this.declaration = declaration;
  }
}

/** Shorthand for an attribute value pair, for `new XmlElement(name, attr(…), …)`. */
export function attr(name: XmlName, value: AttrValue): XmlAttribute {
  return { name, value: String(value) };
}
