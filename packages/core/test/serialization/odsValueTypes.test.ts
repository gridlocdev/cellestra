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
