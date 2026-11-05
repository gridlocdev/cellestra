import { describe, expect, it } from 'vitest';
import { CellValue, OdsReader, OdsWriter, Workbook } from '../../src/index.ts';

describe('ODS writer on sparse extremes', () => {
  it('writes a sheet whose only far cell is XFD1048576 quickly and reads it back', () => {
    const wb = Workbook.createNew();
    const ws = wb.worksheets[0]!;
    ws.setValue(0, 0, CellValue.text('a'));
    ws.setValue(1_048_575, 16_383, CellValue.text('z'));
    for (let r = 1; r < 20_000; r++) ws.getOrCreateRow(r).hidden = true;
    const t = performance.now();
    const bytes = OdsWriter.write(wb);
    expect(performance.now() - t).toBeLessThan(3000);
    const back = OdsReader.read(bytes).worksheets[0]!;
    expect(back.getValue(1_048_575, 16_383)).toEqual(CellValue.text('z'));
    expect(back.getValue(0, 0)).toEqual(CellValue.text('a'));
    expect(back.findRow(15_000)?.hidden).toBe(true);
  });
});
