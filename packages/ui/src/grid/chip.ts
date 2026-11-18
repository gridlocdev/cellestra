import { Argb, type ChipColor } from '@cellestra/core';
import { fitText } from './cellText.ts';
import type { GridPalette } from './palette.ts';
import type { Rect } from './viewport.ts';

/**
 * A list-validated cell draws as a rounded chip with a chevron (PRD §8.13). The grid draws it and
 * the chevron is the button that opens the list, so both take their geometry from here.
 */

/** The chip inside a cell's rect, at `zoom`. */
export function chipRect(cell: Rect, zoom: number): Rect {
  const padX = 4 * zoom;
  const padY = 3 * zoom;
  const h = Math.max(0, Math.min(cell.h - 2 * padY, 22 * zoom));
  return { x: cell.x + padX, y: cell.y + (cell.h - h) / 2, w: Math.max(0, cell.w - 2 * padX), h };
}

/** The chevron's clickable area: the right end of the chip. */
export function chevronRect(cell: Rect, zoom: number): Rect {
  const chip = chipRect(cell, zoom);
  const w = Math.min(chip.w, 22 * zoom);
  return { x: chip.x + chip.w - w, y: chip.y, w, h: chip.h };
}

/** A chip's fill and text as CSS: as stored on a light theme, swapped and darkened on a dark one. */
export function chipColors(color: ChipColor | null, palette: GridPalette): { fill: string; text: string } {
  if (color === null) return { fill: palette.chipFill, text: palette.chipText };
  if (!palette.dark) return { fill: Argb.toCss(color.fill), text: Argb.toCss(color.text) };
  const dim = (c: number) => Math.round(c * 0.72);
  return {
    fill: `rgb(${dim(Argb.r(color.text))},${dim(Argb.g(color.text))},${dim(Argb.b(color.text))})`,
    text: Argb.toCss(color.fill),
  };
}

/**
 * Draw a chip holding `text` in `font` (CSS), or for an empty cell only the chevron, in the
 * header's text colour, so the cell still shows it has a list.
 */
export function drawChip(
  ctx: CanvasRenderingContext2D,
  cell: Rect,
  text: string,
  font: string,
  color: ChipColor | null,
  palette: GridPalette,
  zoom: number,
): void {
  const chip = chipRect(cell, zoom);
  if (chip.w < 8 * zoom || chip.h < 6 * zoom) return;
  const colors = chipColors(color, palette);
  ctx.save();
  ctx.beginPath();
  ctx.rect(cell.x, cell.y, cell.w, cell.h);
  ctx.clip();
  if (text) {
    ctx.fillStyle = colors.fill;
    ctx.beginPath();
    ctx.roundRect(chip.x, chip.y, chip.w, chip.h, chip.h / 2);
    ctx.fill();
  }
  const chevron = chevronRect(cell, zoom);
  ctx.strokeStyle = text ? colors.text : palette.headerText;
  ctx.lineWidth = 1.4 * zoom;
  ctx.beginPath();
  const cx = chevron.x + chevron.w / 2 - 1 * zoom;
  const cy = chevron.y + chevron.h / 2;
  const s = 3.5 * zoom;
  ctx.moveTo(cx - s, cy - s / 2);
  ctx.lineTo(cx, cy + s / 2);
  ctx.lineTo(cx + s, cy - s / 2);
  ctx.stroke();
  if (text) {
    const x = chip.x + 9 * zoom;
    const room = chevron.x - x;
    ctx.font = font;
    ctx.fillStyle = colors.text;
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    ctx.beginPath();
    ctx.rect(x, chip.y, Math.max(0, room), chip.h);
    ctx.clip();
    ctx.fillText(fitText(ctx, text.replace(/\n/g, ' '), room), x, chip.y + chip.h / 2);
  }
  ctx.restore();
}
