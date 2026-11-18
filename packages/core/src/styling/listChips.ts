import { CellValue } from '../values/cellValue.ts';
import type { CellStyle, CellStyleMap } from './cellStyle.ts';
import type { ArgbColor } from './color.ts';
import { parseMapCondition, writeMapCondition } from './conditionalStyles.ts';
import { baseCellAddress } from './conditionalFormatRules.ts';

/**
 * A list item's chip colour (PRD §8.13): a light fill and a dark text colour. The grid draws them
 * as they are on a light theme and swaps them on a dark one.
 */
export interface ChipColor {
  readonly fill: ArgbColor;
  readonly text: ArgbColor;
}

export interface NamedChipColor extends ChipColor {
  readonly name: string;
}

/** The colours the Data Validation dialog offers, in the order new items take them. */
export const CHIP_COLORS: readonly NamedChipColor[] = [
  { name: 'Red', fill: 0xffffc7ce, text: 0xff9c0006 },
  { name: 'Orange', fill: 0xffffe0b2, text: 0xff8a4b00 },
  { name: 'Yellow', fill: 0xfffff2a8, text: 0xff6b5600 },
  { name: 'Green', fill: 0xffc6efce, text: 0xff006100 },
  { name: 'Teal', fill: 0xffc2efea, text: 0xff005c55 },
  { name: 'Blue', fill: 0xffcfe2ff, text: 0xff0b4a9e },
  { name: 'Purple', fill: 0xffe4d7f5, text: 0xff5b2c91 },
  { name: 'Pink', fill: 0xfffad2e6, text: 0xff8a1452 },
  { name: 'Grey', fill: 0xffe3e3e3, text: 0xff3c3c3c },
];

/** Whether two chip colours look the same. */
export function sameChipColor(a: ChipColor | null, b: ChipColor | null): boolean {
  if (a === null || b === null) return a === b;
  return a.fill >>> 0 === b.fill >>> 0 && a.text >>> 0 === b.text >>> 0;
}

/** The text a map's condition tests the cell for equality with, or null when it is any other map. */
function chipValueOf(map: CellStyleMap): string | null {
  if (map.applyStyle === null) return null;
  const c = parseMapCondition(map.condition);
  if (c?.kind !== 'compare' || c.operator !== '=' || c.value.kind !== 'literal') return null;
  return c.value.value.kind === 'text' ? c.value.value.value : null;
}

const key = (text: string) => text.trim().toUpperCase();

/**
 * The chip colour each item has in `style`, keyed by the item in upper case: the fill and text
 * colour of a conditional map that tests the cell for being equal to that text and gives it a fill.
 */
export function chipColorsOf(style: CellStyle): Map<string, ChipColor> {
  const out = new Map<string, ChipColor>();
  for (const map of style.conditionalMaps ?? []) {
    const value = chipValueOf(map);
    const fill = map.applyStyle?.fill.background ?? null;
    if (value === null || fill === null || out.has(key(value))) continue;
    out.set(key(value), { fill, text: map.applyStyle!.font.color });
  }
  return out;
}

/**
 * The style change that gives a list's cells their chip colours. Every map testing for one of
 * `items` (or of `replaced`, the items the list had before) goes; then each item with a colour gets
 * a map that fills the cell with it when the cell holds the item, ahead of the cell's other rules.
 * Calc and LibreOffice read these as ordinary conditional formats.
 */
export function chipColorsTransform(
  items: readonly { readonly value: string; readonly color: ChipColor | null }[],
  replaced: readonly string[],
  base: { readonly sheet: string; readonly row: number; readonly column: number },
): (style: CellStyle) => CellStyle {
  const gone = new Set([...items.map((i) => key(i.value)), ...replaced.map(key)]);
  const address = baseCellAddress(base.sheet, base.row, base.column);
  return (style) => {
    const { conditionalMaps: old, ...plain } = style;
    const kept = (old ?? []).filter((m) => {
      const value = chipValueOf(m);
      return value === null || !gone.has(key(value));
    });
    const added = items.flatMap(({ value, color }): CellStyleMap[] =>
      color === null
        ? []
        : [
            {
              condition: writeMapCondition({
                kind: 'compare',
                operator: '=',
                value: { kind: 'literal', value: CellValue.text(value.trim()) },
              }),
              applyStyleName: null,
              applyStyle: {
                ...plain,
                fill: { ...plain.fill, background: color.fill },
                font: { ...plain.font, color: color.text },
              },
              baseCellAddress: address,
            },
          ],
    );
    const conditionalMaps = [...added, ...kept];
    return conditionalMaps.length > 0 ? { ...plain, conditionalMaps } : plain;
  };
}

/** The palette colour an item takes when it has none: the first one no other item uses. */
export function nextChipColor(used: readonly (ChipColor | null)[]): NamedChipColor {
  const free = CHIP_COLORS.find((c) => !used.some((u) => sameChipColor(u, c)));
  return free ?? CHIP_COLORS[used.length % CHIP_COLORS.length]!;
}
