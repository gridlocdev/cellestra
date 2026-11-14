import { describe, expect, it } from 'vitest';
import { CellValue, OdsReader, OdsWriter, Workbook } from '../../src/index.ts';

describe('ODS writer on sparse extremes', () => {
  it('writes a sheet whose only far cell is XFD1048576 quickly and reads it back', () => {
    const wb = Workbook.createNew();
    const ws = wb.worksheets[0]!;
    ws.setValue(0, 0, CellValue.text('a'));
    ws.setValue(1_048_575, 16_383, CellValue.text('z'));
    for (let r = 1; r < 20_000; r++) ws.getOrCreateRow(r).manuallyHidden = true;
    const t = performance.now();
    const bytes = OdsWriter.write(wb);
    expect(performance.now() - t).toBeLessThan(3000);
    const back = OdsReader.read(bytes).worksheets[0]!;
    expect(back.getValue(1_048_575, 16_383)).toEqual(CellValue.text('z'));
    expect(back.getValue(0, 0)).toEqual(CellValue.text('a'));
    expect(back.findRow(15_000)?.hidden).toBe(true);
  });

  it('writes a run of rows laid out alike as one repeated row', () => {
    const wb = Workbook.createNew();
    const ws = wb.worksheets[0]!;
    ws.setValue(0, 0, CellValue.text('head'));
    for (let r = 1; r <= 200_000; r++) ws.getOrCreateRow(r).manuallyHidden = true;
    ws.setValue(100_000, 0, CellValue.number(1)); // breaks the run in two
    ws.getOrCreateRow(150_000).height = 40; // and a different layout breaks it again
    const t = performance.now();
    const bytes = OdsWriter.write(wb);
    expect(performance.now() - t).toBeLessThan(300);
    const back = OdsReader.read(bytes).worksheets[0]!;
    expect(back.findRow(1)?.hidden).toBe(true);
    expect(back.findRow(99_999)?.hidden).toBe(true);
    expect(back.findRow(100_000)?.hidden).toBe(true);
    expect(back.getValue(100_000, 0)).toEqual(CellValue.number(1));
    expect(back.findRow(150_000)?.height).toBeCloseTo(40, 0);
    expect(back.findRow(150_001)?.height).not.toBe(40);
    expect(back.findRow(200_000)?.hidden).toBe(true);
    expect(back.findRow(200_001)?.hidden ?? false).toBe(false);
  });
});
