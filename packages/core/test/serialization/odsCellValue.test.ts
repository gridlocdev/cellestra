import { describe, expect, it } from 'vitest';
import { readCellValue, writeCellValue } from '../../src/serialization/ods/odsCellValue.ts';
import { Office, Table, Text } from '../../src/serialization/ods/odsXmlNames.ts';
import { NumberFormats } from '../../src/styling/cellStyle.ts';
import { CellError } from '../../src/values/cellError.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { XmlElement, XmlText } from '../../src/xml/index.ts';

type Typed = Parameters<typeof writeCellValue>[1];

/** Write `value` under `format` the way content.xml does, then read it back. */
function roundTrip(value: Typed, format = NumberFormats.general): { back: CellValue; type: string | null } {
  const cell = new XmlElement(Table('table-cell'));
  const lines = writeCellValue(cell, value, format);
  const paragraphs = lines.map((line) => new XmlElement(Text('p')).add(new XmlText(line)));
  return { back: readCellValue(cell, paragraphs), type: cell.attr(Office('value-type')) };
}

describe('ODS cell values', () => {
  it('types a number by its format and reads the same number back', () => {
    const cases: [string, string, number][] = [
      ['General', 'float', 1234.5],
      ['0%', 'percentage', 0.25],
      ['[$€-407] #,##0.00', 'currency', 12.5],
      ['yyyy-mm-dd', 'date', 45000],
      ['yyyy-mm-dd hh:mm:ss', 'date', 45000.5],
      ['[hh]:mm', 'time', 1.25],
    ];
    for (const [code, type, n] of cases) {
      const { back, type: written } = roundTrip(CellValue.number(n), NumberFormats.of(code));
      expect(written, code).toBe(type);
      expect(back, code).toEqual(CellValue.number(n));
    }
  });

  it('keeps infinities, which a plain number text would lose', () => {
    expect(roundTrip(CellValue.number(Infinity)).back).toEqual(CellValue.number(Infinity));
    expect(roundTrip(CellValue.number(-Infinity)).back).toEqual(CellValue.number(-Infinity));
  });

  it('round-trips booleans and errors', () => {
    expect(roundTrip(CellValue.boolean(true)).back).toEqual(CellValue.boolean(true));
    expect(roundTrip(CellValue.boolean(false)).back).toEqual(CellValue.boolean(false));
    const back = roundTrip(CellValue.error(CellError.divideByZero())).back;
    expect(back.kind === 'error' && back.error.token).toBe('#DIV/0!');
  });

  it('reads a string cell from its string-value, else its paragraphs, and an untyped cell as text', () => {
    const typed = new XmlElement(Table('table-cell'))
      .setAttr(Office('value-type'), 'string')
      .setAttr(Office('string-value'), 'value');
    const shown = [new XmlElement(Text('p')).add(new XmlText('display'))];
    expect(readCellValue(typed, shown)).toEqual(CellValue.text('value'));
    expect(readCellValue(new XmlElement(Table('table-cell')), shown)).toEqual(CellValue.text('display'));
    expect(readCellValue(new XmlElement(Table('table-cell')), [])).toEqual(CellValue.empty);
  });

  it('reads a malformed date as empty rather than a wrong day', () => {
    const cell = new XmlElement(Table('table-cell'))
      .setAttr(Office('value-type'), 'date')
      .setAttr(Office('date-value'), '2024-02-30');
    expect(readCellValue(cell, [])).toEqual(CellValue.empty);
  });
});
