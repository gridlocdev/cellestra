import { ChartDocument } from '@cellestra/charts';
import type { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import { tryParseXml, XmlElement } from '../../xml/index.ts';
import { formatUpTo3 } from './odsContentShared.ts';
import type { OdsPackage } from './odsPackage.ts';
import { Draw, Office, Svg, Table, Xlink } from './odsXmlNames.ts';

/**
 * The spreadsheet side of chart-sheet serialization: embeds each chart sheet's chart as an ODF
 * embedded object (a `draw:frame`/`draw:object` in the sheet's table that points at an `Object N/`
 * sub-document), and reads it back. The chart object's own `content.xml` is built and parsed by
 * `ChartDocument` in the chart library.
 */
export const OdsCharts = {
  /**
   * Write each chart sheet's chart object into `pkg` (assigning a stable `Object N` directory the
   * first time) and return those directory names so the manifest can register them.
   */
  writeObjects(workbook: Workbook, pkg: OdsPackage): string[] {
    const dirs: string[] = [];
    const used = new Set<string>();
    for (const sheet of workbook.worksheets) {
      if (sheet.chart && sheet.chart.objectName.length > 0) used.add(sheet.chart.objectName);
    }

    let counter = 1;
    for (const sheet of workbook.worksheets) {
      const chart = sheet.chart;
      if (!chart) continue;

      if (!chart.objectName) {
        let name: string;
        for (;;) {
          name = `Object ${counter++}`;
          if (used.has(name) || pkg.contains(`${name}/content.xml`)) continue;
          used.add(name);
          break;
        }
        chart.objectName = name;
      }

      pkg.setText(`${chart.objectName}/content.xml`, ChartDocument.buildContent(chart));
      dirs.push(chart.objectName);
    }
    return dirs;
  },

  /** Embed the chart's `draw:frame` into a chart sheet's table element (after `writeObjects`). */
  injectFrame(table: XmlElement, sheet: Worksheet): void {
    const chart = sheet.chart;
    if (!chart || chart.objectName.length === 0) return;

    let firstRow = table.element(Table('table-row'));
    if (!firstRow) {
      firstRow = new XmlElement(Table('table-row'), new XmlElement(Table('table-cell')));
      table.add(firstRow);
    }

    let firstCell = firstRow.element(Table('table-cell'));
    if (!firstCell) {
      firstCell = new XmlElement(Table('table-cell'));
      firstRow.addFirst(firstCell);
    }

    const frame = new XmlElement(Draw('frame'))
      .setAttr(Draw('name'), `${sheet.name} Chart`)
      .setAttr(Svg('width'), cm(ChartDocument.defaultWidthCm))
      .setAttr(Svg('height'), cm(ChartDocument.defaultHeightCm))
      .setAttr(Svg('x'), '0cm')
      .setAttr(Svg('y'), '0cm')
      .add(
        new XmlElement(Draw('object'))
          .setAttr(Xlink('href'), './' + chart.objectName)
          .setAttr(Xlink('type'), 'simple')
          .setAttr(Xlink('show'), 'embed')
          .setAttr(Xlink('actuate'), 'onLoad'),
      );
    firstCell.add(frame);
  },

  /**
   * If `table` is a chart sheet — it embeds a chart object and holds no data of its own — parse the
   * chart from `pkg`, attach it to `sheet`, and return the object's href. Returns null otherwise:
   * a chart object in a sheet with data is an embedded chart, preserved as a shape, and must not
   * turn the sheet into a chart sheet (which would drop its cells on the next save).
   */
  readChartSheet(table: XmlElement, pkg: OdsPackage | null, sheet: Worksheet): string | null {
    if (!pkg) return null;

    let reference: string | null = null;
    for (const object of table.descendants(Draw('object'))) {
      const href = object.attr(Xlink('href'));
      if (href) {
        reference = href;
        break;
      }
    }
    if (reference === null) return null;

    for (const cell of table.descendants(Table('table-cell'))) {
      if (cell.attr(Office('value-type')) !== null || cell.attr(Table('formula')) !== null) return null;
    }

    const dir = objectDirectory(reference);
    const xml = pkg.getText(`${dir}/content.xml`);
    if (xml === null) return null;

    const doc = tryParseXml(xml, { preserveWhitespace: false });
    const chart = ChartDocument.parse(doc?.root ?? null);
    if (!chart) return null;

    chart.objectName = dir;
    sheet.chart = chart;
    return reference;
  },

  /** True when `element` holds a `draw:object` pointing at `href`. */
  containsObject(element: XmlElement, href: string): boolean {
    return element.descendants(Draw('object')).some((o) => o.attr(Xlink('href')) === href);
  },
} as const;

/** "./Object 1" or "Object 1/" → "Object 1". */
function objectDirectory(reference: string): string {
  return reference.replace(/^\.+/, '').replace(/^\/+|\/+$/g, '');
}

function cm(value: number): string {
  return formatUpTo3(value) + 'cm';
}
