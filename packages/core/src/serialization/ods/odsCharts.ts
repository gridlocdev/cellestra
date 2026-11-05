import { ChartDocument } from '@cellestra/charts';
import type { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import { tryParseXml, XmlElement } from '../../xml/index.ts';
import { formatUpTo3 } from './odsContentShared.ts';
import type { OdsPackage } from './odsPackage.ts';
import { Draw, Svg, Table, Xlink } from './odsXmlNames.ts';

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
   * If `table` embeds a chart object, parse it from `pkg` and attach it to `sheet` as a chart
   * sheet. Does nothing when there is no embedded chart or it does not parse.
   */
  readChart(table: XmlElement, pkg: OdsPackage | null, sheet: Worksheet): void {
    if (!pkg) return;

    let reference: string | null = null;
    for (const object of table.descendants(Draw('object'))) {
      const href = object.attr(Xlink('href'));
      if (href) {
        reference = href;
        break;
      }
    }
    if (reference === null) return;

    // "./Object 1" or "Object 1/" → "Object 1".
    const dir = reference.replace(/^\.+/, '').replace(/^\/+|\/+$/g, '');
    const xml = pkg.getText(`${dir}/content.xml`);
    if (xml === null) return;

    const doc = tryParseXml(xml, { preserveWhitespace: false });
    const chart = ChartDocument.parse(doc?.root ?? null);
    if (!chart) return;

    chart.objectName = dir;
    sheet.chart = chart;
  },
} as const;

function cm(value: number): string {
  return formatUpTo3(value) + 'cm';
}
