export {
  attr,
  type AttrValue,
  isNamespaceDeclaration,
  qualifiedName,
  sameName,
  XML_NS_URI,
  type XmlAttribute,
  XmlComment,
  type XmlContent,
  type XmlDeclaration,
  XmlDocument,
  XmlElement,
  xmlName,
  type XmlName,
  xmlNamespace,
  type XmlNamespace,
  type XmlNode,
  XMLNS_URI,
  XmlProcessingInstruction,
  XmlText,
} from './xmlTree.ts';
export { parseXml, parseXmlElement, tryParseXml, type XmlParseOptions } from './xmlParser.ts';
export {
  appendXmlText,
  MAX_XML_DEPTH,
  openXmlElement,
  type XmlEventHandler,
  XmlEventReader,
  XmlParseError,
  XmlTreeBuilder,
} from './xmlStream.ts';
export {
  escapeAttr,
  escapeText,
  serializeXml,
  serializeXmlElement,
  type XmlWriteOptions,
} from './xmlWriter.ts';
