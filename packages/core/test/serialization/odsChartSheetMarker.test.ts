import { createChart, createSeries } from '@cellestra/charts';
import { describe, expect, it } from 'vitest';
import { ChartEditCommands } from '../../src/editing/chartEditCommands.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { Workbook } from '../../src/model/workbook.ts';
import { OdsXmlNames, readOdsWithPackage, writeOds } from '../../src/serialization/index.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { odsWith, readEntry } from './helpers.ts';

/** A LibreOffice stock chart: its own size, wall, styles and local table, anchored in A1. */
const STOCK_CHART = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" office:version="1.3"><office:automatic-styles><style:style style:name="ch1" style:family="chart"><style:graphic-properties draw:fill-color="#eeeeee"/></style:style></office:automatic-styles><office:body><office:chart><chart:chart svg:width="9cm" svg:height="5cm" chart:class="chart:stock" chart:style-name="ch1"><chart:plot-area><chart:wall chart:style-name="ch1"/><chart:series chart:values-cell-range-address="local-table.B1:B3"/></chart:plot-area><table:table table:name="local-table"><table:table-rows><table:table-row><table:table-cell office:value-type="float" office:value="1"/></table:table-row></table:table-rows></table:table></chart:chart></office:chart></office:body></office:document-content>`;

const FRAME = (attrs: string) =>
  `<table:table table:name="Sheet1"><table:table-column/><table:table-row><table:table-cell><draw:frame ${attrs}><draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame></table:table-cell></table:table-row></table:table>`;

describe('ODS chart sheets are only the ones Cellestra made', () => {
  it('a LibreOffice chart alone in A1 of an empty sheet stays as it was', () => {
    const bytes = odsWith(
      FRAME('draw:name="Object 1" draw:z-index="0" svg:width="9cm" svg:height="5cm" svg:x="0cm" svg:y="0cm"'),
      {
        extra: { 'Object 1/content.xml': STOCK_CHART },
      },
    );
    const { workbook, package: pkg } = readOdsWithPackage(bytes);
    expect(workbook.activeSheet.isChartSheet).toBe(false);
    const saved = writeOds(workbook, pkg);
    expect(readEntry(saved, 'Object 1/content.xml')).toBe(STOCK_CHART);
    const content = readEntry(saved, OdsXmlNames.contentPart);
    expect(content).toContain('svg:width="9cm"');
    expect(content).not.toContain('24cm');
  });

  it('a chart sheet Cellestra writes is marked and read back as a chart sheet', () => {
    const workbook = Workbook.createNew();
    workbook.activeSheet.setValue(0, 0, CellValue.number(1));
    const chart = createChart({ title: 'T', type: 'line' });
    chart.series.push(createSeries({ valuesReference: 'Sheet1.A1:A1' }));
    workbook.addChartSheet(chart);
    const saved = writeOds(workbook);
    expect(readEntry(saved, OdsXmlNames.contentPart)).toMatch(
      /<draw:frame [^>]*cellestra:chart-sheet="true"/,
    );
    const reopened = readOdsWithPackage(saved);
    const sheet = reopened.workbook.findSheet('Chart1')!;
    expect(sheet.isChartSheet).toBe(true);
    expect(sheet.chart!.type).toBe('line');
  });

  it('a chart sheet an earlier Cellestra wrote, without the mark, is still a chart sheet', () => {
    const chart = `<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" office:version="1.3"><office:body><office:chart><chart:chart chart:class="chart:line"/></office:chart></office:body></office:document-content>`;
    const bytes = odsWith(
      FRAME('draw:name="Sheet1 Chart" svg:width="24cm" svg:height="14cm" svg:x="0cm" svg:y="0cm"'),
      { extra: { 'Object 1/content.xml': chart } },
    );
    const { workbook } = readOdsWithPackage(bytes);
    expect(workbook.activeSheet.isChartSheet).toBe(true);
  });

  it("a chart sheet's object is rewritten only after its chart is edited", () => {
    const workbook = Workbook.createNew();
    const chart = createChart({ title: 'T', type: 'line' });
    chart.series.push(createSeries({ valuesReference: 'Sheet1.A1:A1' }));
    workbook.addChartSheet(chart);
    const first = readOdsWithPackage(writeOds(workbook));
    // Something another editor added to the object, which the model does not hold.
    const objectText = first.package
      .getText('Object 1/content.xml')!
      .replace('</office:document-content>', '<!-- kept --></office:document-content>');
    first.package.setText('Object 1/content.xml', objectText);
    const again = readOdsWithPackage(first.package.write());

    expect(readEntry(writeOds(again.workbook, again.package), 'Object 1/content.xml')).toBe(objectText);

    const edited = { ...again.workbook.findSheet('Chart1')!.chart!, title: 'New' };
    new EditHistory(again.workbook).execute(ChartEditCommands.setChart('Chart1', edited));
    expect(readEntry(writeOds(again.workbook, again.package), 'Object 1/content.xml')).toContain('New');
  });
});
