import { describe, expect, it } from 'vitest';
import { SheetDimensions } from '../../src/model/sheetDimensions.ts';
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

  it('a multi-section code keeps every section and is not read as a date', () => {
    // Sections are written as style:map'd sub-styles and read back through the map. What it must
    // not do is read the "D" of "[RED]" as a date field.
    expect(NumberFormats.kind(NumberFormats.of('0.0;[RED]-0.0'))).toBe('number');
    expect(surviveFormat('0.0;[RED]-0.0').code).toBe('0.0;[RED]-0.0');
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

describe('Styled blank runs after a round trip', () => {
  const bold: CellStyle = { ...CellStyles.default, font: { ...Fonts.default, bold: true } };
  const red: CellStyle = {
    ...CellStyles.default,
    font: { ...Fonts.default, color: Argb.fromRgb(255, 0, 0) },
  };

  function styleAt(workbook: Workbook, row: number, column: number): CellStyle {
    return workbook.styles.get(workbook.activeSheet.effectiveStyleId(row, column));
  }

  it('a block that starts past column A stays a block, not a row format', () => {
    // C2:E30000 bold is one styled run per row, last in the row, over 90,000 cells.
    const workbook = Workbook.createNew();
    const boldId = workbook.styles.intern(bold);
    const sheet = workbook.activeSheet;
    for (let r = 1; r < 30000; r++) for (let c = 2; c <= 4; c++) sheet.getOrCreate(r, c).styleId = boldId;

    const reopened = roundTrip(workbook);
    expect(styleAt(reopened, 1, 2)).toMatchObject(bold);
    expect(styleAt(reopened, 29999, 4)).toMatchObject(bold);
    expect(styleAt(reopened, 1, 0)).toMatchObject(CellStyles.default);
    expect(styleAt(reopened, 1, 1)).toMatchObject(CellStyles.default);
    expect(styleAt(reopened, 500, 5)).toMatchObject(CellStyles.default);
    expect(styleAt(reopened, 500, 100)).toMatchObject(CellStyles.default);
  });

  it('a run down to the last row that starts below the header leaves the header alone', () => {
    // A red run down to the last row under a plain header in A1, next to a bold column B. The
    // run is kept to 70,000 rows so the test stays quick.
    const workbook = Workbook.createNew();
    const sheet = workbook.activeSheet;
    sheet.setValue(0, 0, CellValue.text('Name'));
    const top = SheetDimensions.maxRows - 70_000;
    const redId = workbook.styles.intern(red);
    for (let r = top; r < SheetDimensions.maxRows; r++) sheet.getOrCreate(r, 0).styleId = redId;
    sheet.getOrCreateColumn(1).styleId = workbook.styles.intern(bold);

    const reopened = roundTrip(workbook);
    expect(styleAt(reopened, 0, 0)).toMatchObject(CellStyles.default);
    expect(styleAt(reopened, 1, 0)).toMatchObject(CellStyles.default);
    expect(styleAt(reopened, top, 0)).toMatchObject(red);
    expect(styleAt(reopened, SheetDimensions.maxRows - 1, 0)).toMatchObject(red);
    expect(styleAt(reopened, 0, 1)).toMatchObject(bold);
    expect(styleAt(reopened, 1, 2)).toMatchObject(CellStyles.default);
  });
});
