import { describe, expect, it } from 'vitest';
import { CellAddress } from '../../src/model/cellAddress.ts';
import { odfSheetName } from '../../src/serialization/ods/odsContentShared.ts';

describe('odfSheetName', () => {
  it.each(['m²', 'Q1½', 'Sales 2025', 'TRUE', 'AB12', '2024'])(
    'quotes %j so its address parses back',
    (name) => {
      const written = `${odfSheetName(name)}.A1`;
      expect(written.startsWith("'")).toBe(true);
      expect(CellAddress.parse(written).sheet).toBe(name);
    },
  );

  it('leaves a plain name bare', () => {
    expect(odfSheetName('Sheet1')).toBe('Sheet1');
  });
});
