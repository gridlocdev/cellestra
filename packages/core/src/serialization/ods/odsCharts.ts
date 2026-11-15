import { ChartDocument } from '@cellestra/charts';
import { ChartTableBinding } from '../../charts/chartTableBinding.ts';
import type { Workbook } from '../../model/workbook.ts';
import type { Worksheet } from '../../model/worksheet.ts';
import { isNamespaceDeclaration, tryParseXml, XmlElement } from '../../xml/index.ts';
import { formatUpTo3 } from './odsContentShared.ts';
import type { OdsPackage } from './odsPackage.ts';
import { CellestraNs, Draw, Office, Svg, Table, Xlink } from './odsXmlNames.ts';

/**
 * Marks the `draw:frame` of a chart sheet Cellestra wrote, so only those sheets are read back as
 * chart sheets: a chart another editor anchored in A1 of an empty sheet is an embedded chart.
 */
const CHART_SHEET_MARK = CellestraNs('chart-sheet');

/**
 * Each chart sheet's chart as it was read or last written, so a save rewrites the chart object
 * only when the chart was edited: an object holds much the model does not (styles, a wall, a
 * local table), and another editor may have added it.
 */
const writtenCharts = new WeakMap<Worksheet, string>();

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
      const chart = sheet.chart;
      if (!chart || chart.objectName.length === 0) continue;
      // Two charts can't share one object (a duplicated sheet copies the name): the first keeps it.
      if (used.has(chart.objectName)) chart.objectName = '';
      else used.add(chart.objectName);
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

      const part = `${chart.objectName}/content.xml`;
      // A table-bound chart is written with the table's current ranges, so other readers plot it.
      const written = ChartTableBinding.materialize(workbook, chart);
      const state = JSON.stringify(written);
      if (writtenCharts.get(sheet) !== state || !pkg.contains(part)) {
        pkg.setText(part, ChartDocument.buildContent(written));
        writtenCharts.set(sheet, state);
      }
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
      .setAttr(CHART_SHEET_MARK, 'true')
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
   * If the table `probe` watched is a chart sheet (see `ChartSheetProbe`) and its frame is one
   * Cellestra wrote, parse the chart from `pkg`, attach it to `sheet`, and return the object's
   * href. Returns null otherwise: a chart in a sheet with anything else on it is an embedded chart,
   * preserved with the rest of the sheet, and must not turn the sheet into a chart sheet (which
   * would drop its cells and shapes on the next save).
   */
  readChartSheet(probe: ChartSheetProbe, pkg: OdsPackage | null, sheet: Worksheet): string | null {
    if (!pkg) return null;
    const frame = probe.frame;
    if (!frame || !(frame.attr(CHART_SHEET_MARK) === 'true' || isUnmarkedCellestraFrame(frame, sheet.name)))
      return null;
    const reference = frame.element(Draw('object'))?.attr(Xlink('href')) ?? null;
    if (reference === null) return null;

    const dir = objectDirectory(reference);
    const xml = pkg.getText(`${dir}/content.xml`);
    if (xml === null) return null;

    const doc = tryParseXml(xml, { preserveWhitespace: false });
    const chart = ChartDocument.parse(doc?.root ?? null);
    if (!chart) return null;

    chart.objectName = dir;
    sheet.chart = chart;
    writtenCharts.set(sheet, JSON.stringify(chart));
    return reference;
  },
} as const;

/**
 * Watches a table as it is read, row element by row element, for the shape of a Cellestra chart
 * sheet: the table's children are its grid only, and no cell holds a value, a formula or any
 * element, except A1 (the first cell of the first row), which holds just a frame. The frame holds
 * one `draw:object` and at most the replacement image LibreOffice adds beside it.
 *
 * The answer is known early for almost every table: a value anywhere, or no frame in A1, settles
 * it. Only a table that still looks like a chart sheet has to be read to its end.
 */
export class ChartSheetProbe {
  private ruledOut = false;
  private sawRow = false;
  private candidate: XmlElement | null = null;

  /** The chart frame, while the table can still be a chart sheet. */
  get frame(): XmlElement | null {
    return this.ruledOut ? null : this.candidate;
  }

  /** Whether the table is known to be a chart sheet or not before its end. */
  get settled(): boolean {
    return this.ruledOut || (this.sawRow && this.candidate === null);
  }

  /** The table has a child outside its grid (shapes, forms, named expressions…). */
  nonGridChild(): void {
    this.ruledOut = true;
  }

  /** The next `table:table-row` of the grid, in document order. */
  row(row: XmlElement): void {
    if (this.settled) return;
    let firstCell = !this.sawRow;
    this.sawRow = true;
    for (const cell of row.elements()) {
      const isFirst = firstCell;
      firstCell = false;
      if (cell.attr(Office('value-type')) !== null || cell.attr(Table('formula')) !== null) {
        this.ruledOut = true;
        return;
      }
      const content = cell.elements();
      if (content.length === 0) continue;
      if (!isFirst || content.length !== 1 || !content[0]!.is(Draw('frame')) || !isChartFrame(content[0]!)) {
        this.ruledOut = true;
        return;
      }
      this.candidate = content[0]!;
    }
  }
}

function isChartFrame(frame: XmlElement): boolean {
  const parts = frame.elements();
  const objects = parts.filter((e) => e.is(Draw('object')));
  return objects.length === 1 && parts.every((e) => e.is(Draw('object')) || e.is(Draw('image')));
}

/**
 * A chart sheet frame as Cellestra wrote it before frames were marked: its name, its fixed size at
 * the origin, and no other attribute. LibreOffice always adds attributes of its own (a z-index, an
 * end cell), so a frame it wrote never matches.
 */
function isUnmarkedCellestraFrame(frame: XmlElement, sheetName: string): boolean {
  const expected = new Map<string, string>([
    [`${Draw.uri} name`, `${sheetName} Chart`],
    [`${Svg.uri} width`, cm(ChartDocument.defaultWidthCm)],
    [`${Svg.uri} height`, cm(ChartDocument.defaultHeightCm)],
    [`${Svg.uri} x`, '0cm'],
    [`${Svg.uri} y`, '0cm'],
  ]);
  const attributes = frame.attributes.filter((a) => !isNamespaceDeclaration(a));
  return (
    attributes.length === expected.size &&
    attributes.every((a) => expected.get(`${a.name.uri} ${a.name.local}`) === a.value)
  );
}

/** "./Object 1" or "Object 1/" → "Object 1". */
function objectDirectory(reference: string): string {
  return reference.replace(/^\.+/, '').replace(/^\/+|\/+$/g, '');
}

function cm(value: number): string {
  return formatUpTo3(value) + 'cm';
}
