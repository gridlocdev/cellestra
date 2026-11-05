/** Grid colours per theme (CSS colour strings). Matches the desktop app's GridPalette. */
export interface GridPalette {
  readonly canvas: string;
  readonly gridLine: string;
  readonly headerFill: string;
  readonly activeHeaderFill: string;
  readonly headerText: string;
  readonly defaultText: string;
  readonly freezeDivider: string;
  readonly tableHeaderBand: string;
  readonly tableRowBand: string;
  readonly selectionFill: string;
  readonly activeOutline: string;
  readonly hyperlink: string;
  readonly overflowMarker: string;
}

export const lightPalette: GridPalette = {
  canvas: '#ffffff',
  gridLine: '#e0e0e0',
  headerFill: '#f4f4f4',
  activeHeaderFill: '#e6effb',
  headerText: '#555555',
  defaultText: '#000000',
  freezeDivider: '#9a9a9a',
  tableHeaderBand: '#c6c6c6',
  tableRowBand: 'rgba(198,198,198,0.25)',
  selectionFill: 'rgba(42,125,225,0.13)',
  activeOutline: '#2a7de1',
  hyperlink: '#0b5fcc',
  overflowMarker: '#e53935',
};

export const darkPalette: GridPalette = {
  canvas: '#1e1e1e',
  gridLine: '#3a3a3a',
  headerFill: '#2a2a2a',
  activeHeaderFill: '#2f3e52',
  headerText: '#b0b0b0',
  defaultText: '#e6e6e6',
  freezeDivider: '#5a5a5a',
  tableHeaderBand: '#3a3a3c',
  tableRowBand: 'rgba(58,58,60,0.4)',
  selectionFill: 'rgba(42,125,225,0.2)',
  activeOutline: '#4a90e2',
  hyperlink: '#6bb0ff',
  overflowMarker: '#ff5252',
};
