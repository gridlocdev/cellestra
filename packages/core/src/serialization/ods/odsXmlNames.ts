import { xmlNamespace } from '../../xml/index.ts';

/** The ODF XML namespaces and the part names Cellestra reads and writes. */
export const Office = xmlNamespace('urn:oasis:names:tc:opendocument:xmlns:office:1.0', 'office');
export const Table = xmlNamespace('urn:oasis:names:tc:opendocument:xmlns:table:1.0', 'table');
export const Text = xmlNamespace('urn:oasis:names:tc:opendocument:xmlns:text:1.0', 'text');
export const Style = xmlNamespace('urn:oasis:names:tc:opendocument:xmlns:style:1.0', 'style');
export const Fo = xmlNamespace('urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0', 'fo');
export const NumberNs = xmlNamespace('urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0', 'number');
export const Manifest = xmlNamespace('urn:oasis:names:tc:opendocument:xmlns:manifest:1.0', 'manifest');
export const Svg = xmlNamespace('urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0', 'svg');
export const Config = xmlNamespace('urn:oasis:names:tc:opendocument:xmlns:config:1.0', 'config');
export const Xlink = xmlNamespace('http://www.w3.org/1999/xlink', 'xlink');
export const Draw = xmlNamespace('urn:oasis:names:tc:opendocument:xmlns:drawing:1.0', 'draw');
export const Dc = xmlNamespace('http://purl.org/dc/elements/1.1/', 'dc');
/** OpenFormula: the `of:` prefix on table:formula values resolves to this namespace. */
export const Of = xmlNamespace('urn:oasis:names:tc:opendocument:xmlns:of:1.2', 'of');

/**
 * Cellestra's own extension namespace, for the handful of table facts ODF has no attribute for
 * (a header *column*, banded rows). ODF permits foreign attributes; LibreOffice opens the file
 * untroubled and simply ignores them. Written only when a flag is on, so a workbook that uses no
 * Cellestra-specific feature serializes to pure vocabulary ODF.
 */
export const CellestraNs = xmlNamespace('urn:cellestra:names:table:1.0', 'cellestra');

export const OdsXmlNames = {
  /** The ODF media type of an embedded chart sub-document (its object directory). */
  chartMediaType: 'application/vnd.oasis.opendocument.chart',

  // ---- table:database-range (filter + sort round-trip, ODF §9.4.x) ----
  databaseRanges: 'database-ranges',
  databaseRange: 'database-range',
  targetRangeAddress: 'target-range-address',
  displayFilterButtons: 'display-filter-buttons',
  sort: 'sort',
  sortBy: 'sort-by',
  fieldNumber: 'field-number',
  order: 'order',
  filter: 'filter',
  filterAnd: 'filter-and',
  filterOr: 'filter-or',
  filterCondition: 'filter-condition',
  operator: 'operator',
  dataType: 'data-type',
  containsHeader: 'contains-header',

  // ---- cellestra:* extension attributes on table:database-range --------
  headerColumn: 'header-column',
  bandedRows: 'banded-rows',

  // ---- table:content-validation (data validation round-trip, ODF §9.4.5) ----
  contentValidations: 'content-validations',
  contentValidation: 'content-validation',
  contentValidationName: 'content-validation-name',
  condition: 'condition',
  baseCellAddress: 'base-cell-address',
  allowEmptyCell: 'allow-empty-cell',
  displayList: 'display-list',
  errorMessage: 'error-message',
  helpMessage: 'help-message',
  messageType: 'message-type',
  display: 'display',
  title: 'title',

  mimeType: 'application/vnd.oasis.opendocument.spreadsheet',

  mimetypePart: 'mimetype',
  contentPart: 'content.xml',
  stylesPart: 'styles.xml',
  metaPart: 'meta.xml',
  settingsPart: 'settings.xml',
  manifestPart: 'META-INF/manifest.xml',
} as const;

/** The parts Cellestra regenerates from the model on save (everything else is preserved verbatim). */
export const GeneratedParts: ReadonlySet<string> = new Set([
  OdsXmlNames.mimetypePart,
  OdsXmlNames.contentPart,
  OdsXmlNames.stylesPart,
  OdsXmlNames.metaPart,
  OdsXmlNames.settingsPart,
  OdsXmlNames.manifestPart,
]);
