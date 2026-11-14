/**
 * What a screen reader says when the active cell moves: its address and what it shows, then the
 * selected range when more than one cell is selected. `B7, 42` or `B7, blank, A1:C9 selected`.
 */
export function cellAnnouncement(address: string, shown: string, range: string | null = null): string {
  const text = shown.trim() === '' ? 'blank' : shown;
  return range === null ? `${address}, ${text}` : `${address}, ${text}, ${range} selected`;
}
