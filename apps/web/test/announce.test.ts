import { describe, expect, it } from 'vitest';
import { cellAnnouncement } from '../src/app/announce.ts';

describe('cell announcements', () => {
  it('says the address and what the cell shows', () => {
    expect(cellAnnouncement('B7', '1,234.50')).toBe('B7, 1,234.50');
  });

  it('calls an empty cell blank', () => {
    expect(cellAnnouncement('A1', '')).toBe('A1, blank');
    expect(cellAnnouncement('A1', '  ')).toBe('A1, blank');
  });

  it('adds the selected range when there is one', () => {
    expect(cellAnnouncement('B2', 'x', 'A1:C9')).toBe('B2, x, A1:C9 selected');
  });
});
