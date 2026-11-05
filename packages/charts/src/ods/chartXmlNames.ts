/** The ODF namespaces used inside an embedded chart object's `content.xml`. */
export const ChartXmlNames = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  chart: 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0',
  table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
  /** Cellestra's private namespace for a literal series name (no ODF equivalent exists). */
  cellestraChart: 'urn:cellestra:chart',
  /** The ODF media type of an embedded chart sub-document. */
  chartMediaType: 'application/vnd.oasis.opendocument.chart',
} as const;
