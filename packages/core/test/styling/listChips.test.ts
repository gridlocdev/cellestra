import {
  CellStyles,
  CellValue,
  CHIP_COLORS,
  chipColorsOf,
  chipColorsTransform,
  ConditionalStyles,
  nextChipColor,
  Workbook,
  type CellStyle,
} from '@cellestra/core';
import { describe, expect, it } from 'vitest';

const base = { sheet: 'Sheet1', row: 0, column: 0 };
const red = CHIP_COLORS[0]!;
const green = CHIP_COLORS[3]!;

describe('list chip colours', () => {
  it('writes one equality rule per coloured item and reads the colours back', () => {
    const style = chipColorsTransform(
      [
        { value: 'Yes', color: green },
        { value: 'No - not on Steam', color: red },
        { value: 'Maybe', color: null },
      ],
      [],
      base,
    )(CellStyles.default);
    expect(style.conditionalMaps).toHaveLength(2);
    expect(style.conditionalMaps![0]!.condition).toBe('cell-content()="Yes"');
    const colors = chipColorsOf(style);
    expect(colors.get('YES')).toEqual({ fill: green.fill, text: green.text });
    expect(colors.get('NO - NOT ON STEAM')).toEqual({ fill: red.fill, text: red.text });
    expect(colors.has('MAYBE')).toBe(false);
  });

  it('shows the colour on a cell that holds the item, in any case', () => {
    const wb = Workbook.createNew();
    const sheet = wb.activeSheet;
    sheet.setValue(0, 0, CellValue.text('yes'));
    sheet.setValue(1, 0, CellValue.text('Other'));
    const style = chipColorsTransform([{ value: 'Yes', color: green }], [], base)(CellStyles.default);
    const cs = new ConditionalStyles(wb);
    expect(cs.resolve(style, sheet.name, 0, 0).fill.background).toBe(green.fill);
    expect(cs.resolve(style, sheet.name, 1, 0)).toBe(style);
  });

  it('replaces the colours of items the list had, and keeps other rules', () => {
    const other: CellStyle = {
      ...CellStyles.default,
      conditionalMaps: [
        { condition: 'cell-content()>5', applyStyleName: 'Bad', applyStyle: null, baseCellAddress: null },
      ],
    };
    const first = chipColorsTransform(
      [
        { value: 'A', color: red },
        { value: 'B', color: green },
      ],
      [],
      base,
    )(other);
    const second = chipColorsTransform([{ value: 'A', color: green }], ['A', 'B'], base)(first);
    expect(second.conditionalMaps!.map((m) => m.condition)).toEqual([
      'cell-content()="A"',
      'cell-content()>5',
    ]);
    expect(chipColorsOf(second).get('A')?.fill).toBe(green.fill);
  });

  it('removes every chip rule and leaves a plain style when no item has a colour', () => {
    const coloured = chipColorsTransform([{ value: 'A', color: red }], [], base)(CellStyles.default);
    const cleared = chipColorsTransform([], ['A'], base)(coloured);
    expect(cleared.conditionalMaps).toBeUndefined();
  });

  it('gives a new item the first colour nobody uses', () => {
    expect(nextChipColor([])).toBe(CHIP_COLORS[0]);
    expect(nextChipColor([red, null])).toBe(CHIP_COLORS[1]);
    expect(nextChipColor(CHIP_COLORS)).toBe(CHIP_COLORS[0]);
  });
});
