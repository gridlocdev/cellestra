import { describe, expect, it } from 'vitest';
import { Workbook } from '../../src/model/workbook.ts';
import {
  Alignments,
  Borders,
  CellStyles,
  Fills,
  Fonts,
  NumberFormats,
  type BorderLineStyle,
  type CellStyle,
  type HorizontalAlignment,
  type NumberFormat,
  type VerticalAlignment,
} from '../../src/styling/cellStyle.ts';
import { Argb } from '../../src/styling/color.ts';
import type { NumberFormatKind } from '../../src/styling/numberFormatParser.ts';
import { CellValue } from '../../src/values/cellValue.ts';
import { roundTrip } from './helpers.ts';

/**
 * A sweep of "does what I formatted come back the way I left it" — every property a CellStyle
 * carries, plus the row/column default styles, through open → save → reopen.
 */
function survive(style: CellStyle, withValue = true): CellStyle {
  const workbook = Workbook.createNew();
  const sheet = workbook.activeSheet;
  if (withValue) sheet.setValue(0, 0, CellValue.number(1234.5));
  sheet.getOrCreate(0, 0).styleId = workbook.styles.intern(style);

  const reopened = roundTrip(workbook);
  const cell = reopened.findSheet('Sheet1')!.find(0, 0);
  return cell ? reopened.styles.get(cell.styleId) : CellStyles.default;
}

function surviveFormat(code: string): NumberFormat {
  return survive({ ...CellStyles.default, numberFormat: NumberFormats.of(code) }).numberFormat;
}

describe('FormattingRoundTrip', () => {
  it.each<[string, NumberFormatKind]>([
    ['General', 'general'],
    ['0', 'number'],
    ['0.00', 'number'],
    ['#,##0', 'number'],
    ['#,##0.00', 'number'],
    ['0%', 'percent'],
    ['0.00%', 'percent'],
    ['#,##0.00%', 'percent'], // grouping used to be dropped
    ['$#,##0.00', 'currency'], // the symbol used to be dropped
    ['#,##0.00 €', 'currency'],
    ['yyyy-mm-dd', 'date'],
    ['dd/mm/yyyy', 'date'],
    ['mmm yyyy', 'date'], // a textual month used to come back numeric
    ['mmmm d, yyyy', 'date'],
    ['hh:mm:ss', 'time'],
    ['0.00E+00', 'scientific'], // used to come back as 0.00
    ['# ?/?', 'fraction'], // used to come back as General
    ['@', 'text'], // used to come back as General
  ])('a number format survives with its kind: %s', (code, kind) => {
    const format = surviveFormat(code);

    expect(format.code).toBe(code);
    expect(NumberFormats.kind(format)).toBe(kind);
  });

  it('a multi-section code keeps its first section and is not read as a date', () => {
    // Known limitation: one section per format, so negative/zero/text sections are dropped. What it
    // must not do is read the "D" of "[RED]" as a date field.
    expect(NumberFormats.kind(NumberFormats.of('0.0;[RED]-0.0'))).toBe('number');
    expect(surviveFormat('0.0;[RED]-0.0').code).toBe('0.0');
  });

  it.each<[string, NumberFormatKind]>([
    ['0" days"', 'number'], // the "d" is text the format prints, not a date field
    ['[h]:mm', 'time'], // …but [h] is the elapsed-hours token, and is a field
    ['[$€-407]#,##0.00', 'currency'],
  ])('a literal does not decide the kind: %s', (code, kind) => {
    expect(NumberFormats.kind(NumberFormats.of(code))).toBe(kind);
  });

  it('font properties survive', () => {
    const style: CellStyle = {
      ...CellStyles.default,
      font: {
        family: 'Georgia',
        sizePoints: 14.5,
        bold: true,
        italic: true,
        underline: true,
        strikethrough: true,
        color: Argb.fromRgb(10, 20, 30),
      },
    };

    expect(survive(style).font).toEqual(style.font);
  });

  it.each<HorizontalAlignment>(['general', 'left', 'center', 'right'])(
    'horizontal alignment survives: %s',
    (horizontal) => {
      const style = { ...CellStyles.default, alignment: { ...Alignments.default, horizontal } };
      expect(survive(style).alignment.horizontal).toBe(horizontal);
    },
  );

  it.each<VerticalAlignment>(['top', 'middle', 'bottom'])('vertical alignment survives: %s', (vertical) => {
    const style = { ...CellStyles.default, alignment: { ...Alignments.default, vertical } };
    expect(survive(style).alignment.vertical).toBe(vertical);
  });

  it('wrap text survives', () => {
    const style = { ...CellStyles.default, alignment: { ...Alignments.default, wrapText: true } };
    expect(survive(style).alignment.wrapText).toBe(true);
  });

  it('fill survives', () => {
    const style = { ...CellStyles.default, fill: Fills.solid(Argb.fromRgb(1, 2, 3)) };
    expect(survive(style).fill).toEqual(style.fill);
  });

  it.each<BorderLineStyle>(['thin', 'medium', 'thick', 'dashed', 'dotted', 'double'])(
    'border style survives: %s',
    (line) => {
      const style = { ...CellStyles.default, border: Borders.all(line) };
      expect(survive(style).border).toEqual(style.border);
    },
  );

  it('a styled but empty cell survives', () => {
    const style = {
      ...CellStyles.default,
      fill: Fills.solid(Argb.fromRgb(200, 0, 0)),
      font: { ...Fonts.default, bold: true },
    };
    const after = survive(style, false);

    expect(after.fill).toEqual(style.fill);
    expect(after.font.bold).toBe(true);
  });

  it('a column default style survives', () => {
    const workbook = Workbook.createNew();
    const styleId = workbook.styles.intern({ ...CellStyles.default, font: { ...Fonts.default, bold: true } });
    workbook.activeSheet.getOrCreateColumn(1).styleId = styleId;

    const reopened = roundTrip(workbook);
    const sheet = reopened.findSheet('Sheet1')!;
    expect(reopened.styles.get(sheet.findColumn(1)!.styleId).font.bold).toBe(true);
  });

  it('a row default style survives', () => {
    // The model resolves cell → row → column → default, so a whole-row format must round-trip.
    const workbook = Workbook.createNew();
    const styleId = workbook.styles.intern({
      ...CellStyles.default,
      font: { ...Fonts.default, italic: true },
    });
    workbook.activeSheet.getOrCreateRow(2).styleId = styleId;

    const reopened = roundTrip(workbook);
    const sheet = reopened.findSheet('Sheet1')!;
    expect(reopened.styles.get(sheet.findRow(2)!.styleId).font.italic).toBe(true);
    expect(reopened.styles.get(sheet.effectiveStyleId(2, 5)).font.italic).toBe(true);
  });
});
