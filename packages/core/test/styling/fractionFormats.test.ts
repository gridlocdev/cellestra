import { describe, expect, it } from 'vitest';
import { OdsDataStyles } from '../../src/serialization/ods/mapping/odsDataStyles.ts';
import { classifyNumberFormat } from '../../src/styling/numberFormatParser.ts';
import { NumberFormatter } from '../../src/styling/numberFormatter.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { parseXmlElement } from '../../src/xml/index.ts';
import { NumberNs } from '../../src/serialization/ods/odsXmlNames.ts';
import { NumberFormats } from '../../src/styling/cellStyle.ts';

const show = (n: number, code: string, locale = 'en-US'): string =>
  NumberFormatter.format(CellValue.number(n), code, locale);

const NS =
  'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"';

/** The code a hand-written ODF number style reads back as. */
function codeOf(inner: string): string {
  const root = parseXmlElement(`<number:number-style ${NS} style:name="N1">${inner}</number:number-style>`);
  return OdsDataStyles.toCode(root);
}

describe('fraction number formats', () => {
  it.each([
    [1.5, '# ?/?', '1 1/2'],
    [0.5, '# ?/?', ' 1/2'], // no whole part: the integer placeholder and its digit are blank
    [-1.5, '# ?/?', '-1 1/2'],
    [3, '# ?/?', '3    '], // no fraction left: blanks the width of the ?/? part
    [0, '# ?/?', '0    '],
    [Math.PI, '# ?/?', '3 1/7'],
    [Math.PI, '# ??/??', '3 14/99'],
    [1.5, '# ??/??', '1  1/2 '], // numerator padded on the left, denominator on the right
    [0.999, '# ?/?', '1    '], // rounds up to the next whole number
    [1.25, '?/?', '5/4'], // no whole part: the numerator takes it
    [0, '?/?', '0/1'],
    [0.75, '?/4', '3/4'],
    [1.25, '?/4', '5/4'],
    [1.3, '# ?/4', '1 1/4'],
    [2, '# ?/4', '2    '],
    [0.5, '0 ?/?', '0 1/2'],
    [0.1, '# ?/16', ' 2/16'],
    [1.5, '# ?/?" cups"', '1 1/2 cups'],
  ])('%d under %s is %j', (n, code, expected) => {
    expect(show(n, code)).toBe(expected);
  });

  it.each(['# ?/?', '# ??/??', '?/4', '0 ?/?', '?/?', '# ?/16'])('%s is a fraction format', (code) => {
    expect(classifyNumberFormat(code)).toBe('fraction');
  });

  it('dates with slashes stay dates', () => {
    expect(classifyNumberFormat('d/m/yy')).toBe('date');
  });

  it.each(['# ?/?', '# ??/??', '?/4', '# ?/4', '?/?', '0 ?/?', '# ?/16', '# ???/???'])(
    'writes and reads back %s',
    (code) => {
      const [style] = OdsDataStyles.build('N1', NumberFormats.of(code));
      expect(OdsDataStyles.toCode(style!)).toBe(code);
    },
  );

  it('writes a fixed denominator as number:denominator-value', () => {
    const [style] = OdsDataStyles.build('N1', NumberFormats.of('# ?/8'));
    const fraction = style!.elements()[0]!;
    expect(fraction.local).toBe('fraction');
    expect(fraction.attr(NumberNs('denominator-value'))).toBe('8');
  });

  it.each([
    [
      '<number:fraction number:min-integer-digits="0" number:min-numerator-digits="1" number:min-denominator-digits="1"/>',
      '# ?/?',
    ],
    [
      '<number:fraction number:min-integer-digits="0" number:min-numerator-digits="2" number:min-denominator-digits="2" number:max-denominator-value="99"/>',
      '# ??/??',
    ],
    [
      '<number:fraction number:min-integer-digits="0" number:min-numerator-digits="3" number:min-denominator-digits="1" number:max-denominator-value="999"/>',
      '# ???/???',
    ],
    [
      '<number:fraction number:min-integer-digits="0" number:min-numerator-digits="1" number:min-denominator-digits="1" number:denominator-value="4"/>',
      '# ?/4',
    ],
    ['<number:fraction number:min-numerator-digits="1" number:min-denominator-digits="1"/>', '?/?'],
    [
      '<number:fraction number:min-integer-digits="1" number:min-numerator-digits="1" number:min-denominator-digits="1"/>',
      '0 ?/?',
    ],
  ])('reads LibreOffice fraction style %s as %s', (xml, code) => {
    expect(codeOf(xml)).toBe(code);
  });
});
