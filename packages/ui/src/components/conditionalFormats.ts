import {
  CellValue,
  FormulaWriter,
  Parser,
  SheetDimensions,
  parseMapCondition,
  writeMapCondition,
  type ArgbColor,
  type CellStyle,
  type CellStyleMap,
  type ConditionOperator,
  type Expr,
  type MapCondition,
} from '@cellestra/core';

/** What a rule checks. */
export type RuleKind =
  | 'greater'
  | 'less'
  | 'greaterOrEqual'
  | 'lessOrEqual'
  | 'equal'
  | 'notEqual'
  | 'between'
  | 'notBetween'
  | 'formula';

export const RULE_KINDS: readonly { kind: RuleKind; label: string; operands: 1 | 2 }[] = [
  { kind: 'greater', label: 'Cell value is greater than', operands: 1 },
  { kind: 'less', label: 'Cell value is less than', operands: 1 },
  { kind: 'greaterOrEqual', label: 'Cell value is at least', operands: 1 },
  { kind: 'lessOrEqual', label: 'Cell value is at most', operands: 1 },
  { kind: 'equal', label: 'Cell value is equal to', operands: 1 },
  { kind: 'notEqual', label: 'Cell value is not equal to', operands: 1 },
  { kind: 'between', label: 'Cell value is between', operands: 2 },
  { kind: 'notBetween', label: 'Cell value is not between', operands: 2 },
  { kind: 'formula', label: 'Formula is true', operands: 1 },
];

const OPERATOR_OF: Partial<Record<RuleKind, ConditionOperator>> = {
  greater: '>',
  less: '<',
  greaterOrEqual: '>=',
  lessOrEqual: '<=',
  equal: '=',
  notEqual: '!=',
};

/** The formats a rule can apply. `keep` is a format read from a file that no preset matches. */
export type FormatPresetId = 'lightRed' | 'yellow' | 'green' | 'redText' | 'bold' | 'keep';

interface FormatPreset {
  readonly id: Exclude<FormatPresetId, 'keep'>;
  readonly label: string;
  readonly fill: ArgbColor | null;
  readonly text: ArgbColor | null;
  readonly bold: boolean;
}

export const FORMAT_PRESETS: readonly FormatPreset[] = [
  { id: 'lightRed', label: 'Light red fill, dark red text', fill: 0xffffc7ce, text: 0xff9c0006, bold: false },
  { id: 'yellow', label: 'Yellow fill, dark yellow text', fill: 0xffffeb9c, text: 0xff9c5700, bold: false },
  { id: 'green', label: 'Green fill, dark green text', fill: 0xffc6efce, text: 0xff006100, bold: false },
  { id: 'redText', label: 'Red text', fill: null, text: 0xffc00000, bold: false },
  { id: 'bold', label: 'Bold text', fill: null, text: null, bold: true },
];

/** A preset applied over a cell's own style: its number format, borders and alignment stay. */
export function applyPreset(id: Exclude<FormatPresetId, 'keep'>, style: CellStyle): CellStyle {
  const p = FORMAT_PRESETS.find((f) => f.id === id)!;
  const { conditionalMaps: _maps, ...plain } = style;
  return {
    ...plain,
    fill: p.fill !== null ? { background: p.fill } : plain.fill,
    font: {
      ...plain.font,
      ...(p.text !== null ? { color: p.text } : {}),
      ...(p.bold ? { bold: true } : {}),
    },
  };
}

/** One rule as the dialog edits it. */
export interface RuleDraft {
  readonly kind: RuleKind;
  readonly value1: string;
  readonly value2: string;
  readonly format: FormatPresetId;
  /** The map this rule was read from, reused as it was while the rule is not edited. */
  readonly original?: CellStyleMap;
  /** For `keep`: the style the file's rule applies, which no preset matches. */
  readonly keepStyle?: CellStyle | null;
}

export const newRule = (): RuleDraft => ({ kind: 'greater', value1: '', value2: '', format: 'lightRed' });

/** Edit a rule; it then no longer stands for the map it was read from. */
export function editRule(
  rule: RuleDraft,
  patch: Partial<Pick<RuleDraft, 'kind' | 'value1' | 'value2' | 'format'>>,
): RuleDraft {
  const { original: _original, ...rest } = { ...rule, ...patch };
  return rest;
}

/** The rules a style's conditional maps stand for, in order. */
export function rulesOf(style: CellStyle): RuleDraft[] {
  const { conditionalMaps: maps, ...plain } = style;
  return (maps ?? []).flatMap((map): RuleDraft[] => {
    const condition = parseMapCondition(map.condition);
    if (!condition) return [];
    const format =
      map.applyStyle === null
        ? 'keep'
        : (FORMAT_PRESETS.find((p) => sameLook(applyPreset(p.id, plain), map.applyStyle!))?.id ?? 'keep');
    return [
      {
        ...fromCondition(condition),
        format,
        original: map,
        ...(format === 'keep' ? { keepStyle: map.applyStyle } : {}),
      },
    ];
  });
}

function sameLook(a: CellStyle, b: CellStyle): boolean {
  return (
    (a.fill.background ?? 0) >>> 0 === (b.fill.background ?? 0) >>> 0 &&
    a.font.color >>> 0 === b.font.color >>> 0 &&
    a.font.bold === b.font.bold
  );
}

function fromCondition(c: MapCondition): Pick<RuleDraft, 'kind' | 'value1' | 'value2'> {
  switch (c.kind) {
    case 'compare': {
      const kind = (Object.keys(OPERATOR_OF) as RuleKind[]).find((k) => OPERATOR_OF[k] === c.operator)!;
      return { kind, value1: FormulaWriter.write(c.value), value2: '' };
    }
    case 'between':
    case 'notBetween':
      return { kind: c.kind, value1: FormulaWriter.write(c.low), value2: FormulaWriter.write(c.high) };
    case 'formula':
      return { kind: 'formula', value1: FormulaWriter.write(c.formula), value2: '' };
  }
}

/**
 * A value typed for a rule, as an expression: a number, a reference, a formula (`=B1*2`), or else
 * the text itself (`Done` needs no quotes).
 */
export function parseRuleValue(text: string, formula: boolean): Expr | null {
  const t = text.trim();
  if (!t) return null;
  const source = t.startsWith('=') ? t.slice(1) : t;
  const parsed = Parser.tryParse(source);
  if (formula || t.startsWith('=')) return parsed.ok ? parsed.expr : null;
  if (parsed.ok && !onlyWords(parsed.expr)) return parsed.expr;
  return { kind: 'literal', value: CellValue.text(t) };
}

/** A bare name, or words that only parse as names intersected by spaces (`In progress`). */
function onlyWords(expr: Expr): boolean {
  if (expr.kind === 'name') return true;
  return (
    expr.kind === 'binary' && expr.operator === 'intersect' && onlyWords(expr.left) && onlyWords(expr.right)
  );
}

function conditionOf(rule: RuleDraft): MapCondition | string {
  const info = RULE_KINDS.find((k) => k.kind === rule.kind)!;
  const first = parseRuleValue(rule.value1, rule.kind === 'formula');
  if (!first) return rule.kind === 'formula' ? 'Type a formula for the rule.' : 'Type a value for the rule.';
  if (rule.kind === 'formula') return { kind: 'formula', formula: first };
  if (info.operands === 1) return { kind: 'compare', operator: OPERATOR_OF[rule.kind]!, value: first };
  const second = parseRuleValue(rule.value2, false);
  if (!second) return 'Type both values for the rule.';
  return { kind: rule.kind as 'between' | 'notBetween', low: first, high: second };
}

/** The address a rule's relative references are relative to, as ODF writes it (`Sheet1.A1`). */
export function baseCellAddress(sheet: string, row: number, column: number): string {
  const name = /^[A-Za-z_][A-Za-z0-9_]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`;
  return `${name}.${SheetDimensions.columnToLetters(column)}${row + 1}`;
}

/**
 * The style change that gives every selected cell `rules`, or the first problem with them. Each
 * cell's applied formats are built over its own style, so its number format and borders stay.
 */
export function rulesTransform(
  rules: readonly RuleDraft[],
  base: { readonly sheet: string; readonly row: number; readonly column: number },
): { transform: (style: CellStyle) => CellStyle } | { error: string } {
  const address = baseCellAddress(base.sheet, base.row, base.column);
  const built: (
    { map: CellStyleMap } | { condition: string; format: FormatPresetId; keep: CellStyle | null }
  )[] = [];
  for (const rule of rules) {
    if (rule.original) {
      built.push({ map: rule.original });
      continue;
    }
    const condition = conditionOf(rule);
    if (typeof condition === 'string') return { error: condition };
    built.push({
      condition: writeMapCondition(condition),
      format: rule.format,
      keep: rule.keepStyle ?? null,
    });
  }
  return {
    transform: (style) => {
      const { conditionalMaps: _old, ...plain } = style;
      if (built.length === 0) return plain;
      const conditionalMaps = built.map((b): CellStyleMap =>
        'map' in b
          ? b.map
          : {
              condition: b.condition,
              applyStyleName: null,
              applyStyle: b.format === 'keep' ? b.keep : applyPreset(b.format, plain),
              baseCellAddress: address,
            },
      );
      return { ...plain, conditionalMaps };
    },
  };
}
