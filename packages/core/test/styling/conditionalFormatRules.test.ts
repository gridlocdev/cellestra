import {
  CellStyles,
  ConditionalStyles,
  CellValue,
  NumberFormats,
  Workbook,
  type CellStyle,
} from '@cellestra/core';
import { describe, expect, it } from 'vitest';
import {
  applyPreset,
  baseCellAddress,
  editRule,
  newRule,
  parseRuleValue,
  rulesOf,
  rulesTransform,
} from '../../src/styling/conditionalFormatRules.ts';

const base = { sheet: 'Sheet1', row: 0, column: 0 };
const transform = (rules: Parameters<typeof rulesTransform>[0]) => {
  const r = rulesTransform(rules, base);
  if ('error' in r) throw new Error(r.error);
  return r.transform;
};

describe('conditional format rules', () => {
  it('reads typed values as numbers, references, formulas or text', () => {
    expect(parseRuleValue('10', false)).toMatchObject({
      kind: 'literal',
      value: { kind: 'number', value: 10 },
    });
    expect(parseRuleValue('B1', false)).toMatchObject({ kind: 'reference' });
    expect(parseRuleValue('Done', false)).toMatchObject({
      kind: 'literal',
      value: { kind: 'text', value: 'Done' },
    });
    expect(parseRuleValue('In progress', false)).toMatchObject({ value: { value: 'In progress' } });
    expect(parseRuleValue('=B1*2', false)).toMatchObject({ kind: 'binary' });
    expect(parseRuleValue('A1>', true)).toBeNull();
    expect(parseRuleValue('  ', false)).toBeNull();
  });

  it('builds maps over each cell’s own style, and reads them back as the same rules', () => {
    const money: CellStyle = { ...CellStyles.default, numberFormat: NumberFormats.of('#,##0.00') };
    const rules = [
      { ...newRule(), kind: 'greater' as const, value1: '100', format: 'green' as const },
      { ...newRule(), kind: 'between' as const, value1: '1', value2: 'B1', format: 'bold' as const },
      { ...newRule(), kind: 'formula' as const, value1: 'A1=MAX($A$1:$A$9)', format: 'redText' as const },
    ];
    const styled = transform(rules)(money);
    expect(styled.conditionalMaps!.map((m) => m.condition)).toEqual([
      'cell-content()>100',
      'cell-content-is-between(1,[.B1])',
      'is-true-formula([.A1]=MAX([.$A$1:.$A$9]))',
    ]);
    expect(styled.conditionalMaps![0]!.baseCellAddress).toBe('Sheet1.A1');
    expect(styled.conditionalMaps![0]!.applyStyle!.numberFormat.code).toBe('#,##0.00');
    expect(styled.conditionalMaps![0]!.applyStyle).toEqual(applyPreset('green', money));

    const back = rulesOf(styled);
    expect(back.map((r) => [r.kind, r.value1, r.value2, r.format])).toEqual([
      ['greater', '100', '', 'green'],
      ['between', '1', 'B1', 'bold'],
      ['formula', 'A1=MAX($A$1:$A$9)', '', 'redText'],
    ]);
    // Unedited rules keep their maps exactly; no rules removes them.
    expect(transform(back)(styled).conditionalMaps).toEqual(styled.conditionalMaps);
    expect(transform([])(styled)).toEqual(money);
  });

  it('keeps a file’s own format until the rule is given a preset', () => {
    const odd: CellStyle = { ...CellStyles.default, font: { ...CellStyles.default.font, italic: true } };
    const style: CellStyle = {
      ...CellStyles.default,
      conditionalMaps: [
        {
          condition: 'cell-content()<0',
          applyStyleName: 'Bad',
          applyStyle: odd,
          baseCellAddress: 'Sheet1.A1',
        },
      ],
    };
    const [rule] = rulesOf(style);
    expect(rule!.format).toBe('keep');
    const edited = editRule(rule!, { value1: '-5' });
    const maps = transform([edited])(CellStyles.default).conditionalMaps!;
    expect(maps[0]).toMatchObject({ condition: 'cell-content()<-5', applyStyle: odd, applyStyleName: null });
  });

  it('reports a rule with a missing value', () => {
    expect(rulesTransform([newRule()], base)).toEqual({ error: 'Type a value for the rule.' });
    expect(rulesTransform([{ ...newRule(), kind: 'between', value1: '1' }], base)).toEqual({
      error: 'Type both values for the rule.',
    });
  });

  it('quotes sheet names that need it in the base cell', () => {
    expect(baseCellAddress('My Sheet', 4, 27)).toBe("'My Sheet'.AB5");
  });

  it('applies through the engine as the dialog built it', () => {
    const wb = Workbook.createNew();
    const sheet = wb.worksheets[0]!;
    sheet.setValue(0, 0, CellValue.text('Done'));
    sheet.setValue(1, 0, CellValue.text('Open'));
    const style = transform([{ ...newRule(), kind: 'equal', value1: 'done', format: 'lightRed' }])(
      CellStyles.default,
    );
    const cs = new ConditionalStyles(wb);
    expect(cs.resolve(style, 'Sheet1', 0, 0).fill.background).toBe(0xffffc7ce);
    expect(cs.resolve(style, 'Sheet1', 1, 0)).toBe(style);
  });

  describe('a rule the dialog cannot read', () => {
    const map = (condition: string) => ({
      condition,
      applyStyleName: 'Accent',
      applyStyle: applyPreset('bold', CellStyles.default),
      baseCellAddress: 'Sheet1.A1',
    });
    const unread = map('cell-content-is-in-list("a";"b")');
    const style: CellStyle = {
      ...CellStyles.default,
      conditionalMaps: [map('cell-content()>1'), unread, map('cell-content()<0')],
    };

    it('is not shown', () => {
      expect(rulesOf(style).map((r) => r.value1)).toEqual(['1', '0']);
    });

    it('survives OK without changes, in its place', () => {
      expect(transform(rulesOf(style))(style).conditionalMaps).toEqual(style.conditionalMaps);
    });

    it('survives edits to the rules around it', () => {
      const [first, second] = rulesOf(style);
      const maps = transform([editRule(first!, { value1: '2' }), second!])(style).conditionalMaps!;
      expect(maps.map((m) => m.condition)).toEqual([
        'cell-content()>2',
        unread.condition,
        'cell-content()<0',
      ]);
      expect(transform([])(style).conditionalMaps).toEqual([unread]);
    });
  });
});
