import { createChart, createSeries } from '@cellestra/charts';
import { describe, expect, it } from 'vitest';
import { Workbook } from '../../src/model/workbook.ts';
import { readOds, writeOds } from '../../src/serialization/index.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { entryNames, readEntry, roundTripWithPackage } from './helpers.ts';

function buildWithChart(): Workbook {
  const workbook = Workbook.createNew();
  const data = workbook.activeSheet; // Sheet1
  data.setValue(0, 0, CellValue.text('Region'));
  data.setValue(0, 1, CellValue.text('Q1'));
  data.setValue(0, 2, CellValue.text('Q2'));
  for (let r = 1; r <= 3; r++) {
    data.setValue(r, 0, CellValue.text(`R${r}`));
    data.setValue(r, 1, CellValue.number(r * 10));
    data.setValue(r, 2, CellValue.number(r * 5));
  }

  const chart = createChart({
    title: 'Quarterly',
    type: 'bar', // horizontal bars → chart:bar with chart:vertical
    stacking: 'percent',
    categoriesReference: 'Sheet1.A2:A4',
    showLegend: true,
    xAxisTitle: 'Region',
    yAxisTitle: 'Sales',
  });
  chart.series.push(createSeries({ nameReference: 'Sheet1.B1', valuesReference: 'Sheet1.B2:B4' }));
  chart.series.push(createSeries({ name: 'Second', valuesReference: 'Sheet1.C2:C4', type: 'line' }));
  workbook.addChartSheet(chart);
  return workbook;
}

describe('ChartSerialization', () => {
  it('chart sheet round trips its definition', () => {
    const { workbook: reopened } = roundTripWithPackage(buildWithChart());

    const chartSheet = reopened.findSheet('Chart1');
    expect(chartSheet).toBeDefined();
    expect(chartSheet!.isChartSheet).toBe(true);

    const chart = chartSheet!.chart!;
    expect(chart.title).toBe('Quarterly');
    expect(chart.type).toBe('bar');
    expect(chart.stacking).toBe('percent');
    expect(chart.categoriesReference).toBe('Sheet1.A2:A4');
    expect(chart.xAxisTitle).toBe('Region');
    expect(chart.yAxisTitle).toBe('Sales');
    expect(chart.showLegend).toBe(true);

    expect(chart.series).toHaveLength(2);
    expect(chart.series[0]!.nameReference).toBe('Sheet1.B1');
    expect(chart.series[0]!.valuesReference).toBe('Sheet1.B2:B4');
    expect(chart.series[0]!.type).toBeNull(); // same as chart type → no override
    expect(chart.series[1]!.name).toBe('Second');
    expect(chart.series[1]!.type).toBe('line'); // combo: per-series override preserved
  });

  it('chart object is a valid ods part with chart class and frame', () => {
    const bytes = writeOds(buildWithChart());

    // The embedded chart object directory exists and declares the chart class.
    const objectEntry = entryNames(bytes).find((n) => n.startsWith('Object') && n.endsWith('content.xml'));
    expect(objectEntry).toBeDefined();
    const chartXml = readEntry(bytes, objectEntry!);
    expect(chartXml).toContain('chart:class="chart:bar"');
    expect(chartXml).toContain('chart:vertical="true"'); // Bar orientation
    expect(chartXml).toContain('chart:percentage="true"'); // 100% stacked
    expect(chartXml).toContain('chart:class="chart:line"'); // combo series

    // The spreadsheet content.xml anchors the chart with a draw frame referencing the object.
    const content = readEntry(bytes, 'content.xml');
    expect(content).toContain('draw:frame');
    expect(content).toContain('draw:object');

    // The manifest declares the object with the chart media type.
    expect(readEntry(bytes, 'META-INF/manifest.xml')).toContain('application/vnd.oasis.opendocument.chart');
  });

  it('a duplicated chart sheet saves its own chart object', () => {
    const { workbook, package: pkg } = roundTripWithPackage(buildWithChart());
    const copy = workbook.duplicateSheet(workbook.findSheet('Chart1')!);
    copy.chart!.title = 'Copy';

    const reopened = readOds(writeOds(workbook, pkg));
    expect(reopened.findSheet('Chart1')!.chart!.title).toBe('Quarterly');
    expect(reopened.findSheet(copy.name)!.chart!.title).toBe('Copy');
  });

  it('two chart sheets that claim the same object are saved to separate objects', () => {
    const workbook = buildWithChart();
    const first = workbook.findSheet('Chart1')!.chart!;
    first.objectName = 'Object 1';
    workbook.addChartSheet(createChart({ title: 'Other', objectName: 'Object 1' }));

    const reopened = readOds(writeOds(workbook));
    const titles = reopened.worksheets.filter((s) => s.isChartSheet).map((s) => s.chart!.title);
    expect(titles).toEqual(['Quarterly', 'Other']);
  });

  it('data sheet without a chart stays a plain sheet', () => {
    const { workbook: reopened } = roundTripWithPackage(buildWithChart());
    expect(reopened.findSheet('Sheet1')!.isChartSheet).toBe(false);
  });
});
