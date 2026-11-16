import { describe, expect, it } from 'vitest';
import { Workbook } from '../../src/model/workbook.ts';
import { readOds, writeOds } from '../../src/serialization/index.ts';
import { CellStyles, NumberFormats } from '../../src/styling/cellStyle.ts';
import { CellError } from '../../src/values/cellError.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { odsWith, readEntry, roundTrip } from './helpers.ts';

/** Cell value types as LibreOffice writes them: currency, percentage, and formula errors. */
function withFormat(code: string): Workbook {
  const workbook = Workbook.createNew();
  const sheet = workbook.activeSheet;
  sheet.setValue(0, 0, CellValue.number(1.5));
  sheet.getOrCreate(0, 0).styleId = workbook.styles.intern({
    ...CellStyles.default,
    numberFormat: NumberFormats.of(code),
  });
  return workbook;
}

function firstCell(workbook: Workbook): string {
  const xml = readEntry(writeOds(workbook), 'content.xml');
  const cell = /<table:table-cell [^>]*>/.exec(xml);
  if (!cell) throw new Error('no cell written');
  return cell[0];
}

describe('ODS value types', () => {
  it('writes a currency-formatted number as a currency value', () => {
    const cell = firstCell(withFormat('[$€-407]#,##0.00'));
    expect(cell).toContain('office:value-type="currency"');
    expect(cell).toContain('office:currency="EUR"');
    expect(cell).toContain('office:value="1.5"');
  });

  it('names the currency of a dollar format', () => {
    const cell = firstCell(withFormat('$#,##0.00'));
    expect(cell).toContain('office:value-type="currency"');
    expect(cell).toContain('office:currency="USD"');
  });

  it('writes a percent-formatted number as a percentage value', () => {
    const cell = firstCell(withFormat('0.00%'));
    expect(cell).toContain('office:value-type="percentage"');
    expect(cell).toContain('office:value="1.5"');
  });

  it('keeps a number too large to be a date when its cell has a date format', () => {
    const workbook = withFormat('yyyy-mm-dd');
    workbook.activeSheet.setValue(0, 0, CellValue.number(1e9));

    expect(firstCell(workbook)).toContain('office:value-type="float"');
    expect(roundTrip(workbook).activeSheet.getValue(0, 0)).toEqual(CellValue.number(1e9));
  });

  it('writes a formula error the way LibreOffice does and reads it back as an error', () => {
    const workbook = Workbook.createNew();
    workbook.settings.calculationMode = 'manual';
    workbook.activeSheet.setFormula(0, 0, '1/0', CellValue.error(CellError.divideByZero()));
    const cell = firstCell(workbook);
    expect(cell).toContain('office:value-type="string"');
    expect(cell).toContain('office:string-value=""');
    expect(cell).toContain('calcext:value-type="error"');

    const value = roundTrip(workbook).activeSheet.getValue(0, 0);
    expect(value.kind === 'error' && value.error.code).toBe('divideByZero');
  });

  it('reads a LibreOffice #N/A cell as an error', () => {
    const bytes = odsWith(
      '<table:table table:name="S"><table:table-row><table:table-cell table:formula="of:=NA()" office:value-type="string" office:string-value="" calcext:value-type="error"><text:p>#N/A</text:p></table:table-cell></table:table-row></table:table>',
    );
    const value = readOds(bytes).activeSheet.getValue(0, 0);
    expect(value.kind === 'error' && value.error.code).toBe('notAvailable');
  });
});

describe('ODS value types from a row or column format', () => {
  function formattedLine(line: 'row' | 'column', code: string): Workbook {
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    const styleId = workbook.styles.intern({ ...CellStyles.default, numberFormat: NumberFormats.of(code) });
    if (line === 'column') sheet.getOrCreateColumn(0).styleId = styleId;
    else sheet.getOrCreateRow(0).styleId = styleId;
    sheet.setValue(0, 0, CellValue.number(46082));
    return workbook;
  }

  it('writes a number under a date column format as a date, as LibreOffice does', () => {
    const workbook = formattedLine('column', 'yyyy-mm-dd');
    expect(workbook.activeSheet.find(0, 0)!.styleId).toBe(0);
    const cell = firstCell(workbook);
    expect(cell).toContain('office:value-type="date"');
    expect(cell).toContain('office:date-value="2026-03-01"');
    expect(readEntry(writeOds(workbook), 'content.xml')).toContain('<text:p>2026-03-01</text:p>');
    const value = roundTrip(workbook).activeSheet.getValue(0, 0);
    expect(value.kind === 'number' && value.value).toBe(46082);
  });

  it('writes a number under a percent row format as a percentage', () => {
    const cell = firstCell(formattedLine('row', '0.00%'));
    expect(cell).toContain('office:value-type="percentage"');
  });
});
