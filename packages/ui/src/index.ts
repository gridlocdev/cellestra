export {
  ACCEPT,
  AuthError,
  BrowserStore,
  type CellestraNative,
  ChangePoller,
  type ChangePollerOptions,
  clearRecovery,
  CONFLICT,
  ConflictError,
  type DocumentRef,
  type DocumentState,
  download,
  type FilePicker,
  type FileStore,
  forgetRecent,
  keepRecoveryAside,
  latestOnly,
  type LockOwner,
  lockWarning,
  mimeFor,
  NAS_POLL_MS,
  type NasEntry,
  type NasInfo,
  type NasRoot,
  NasStore,
  nativeBridge,
  type NativeFile,
  type NativeMenuItem,
  NativeStore,
  type OpenedDocument,
  orphanedRecoveries,
  parseLockFile,
  type PollClock,
  type PrintPaper,
  type RecentEntry,
  recentFiles,
  recoveryKey,
  type RecoverySnapshot,
  rememberRecent,
  type SavedFile,
  saveRecovery,
  OPENABLE_TYPES,
  SPREADSHEET_TYPES,
  type StoreCapabilities,
  type StoreKind,
} from './platform/index.ts';
export {
  type Command,
  CommandCatalog,
  type CommandGroup,
  type CommandSpec,
  compact,
  formatShortcut,
  inOrder,
  isMac,
  matches,
  matchKind,
  parseShortcut,
  type Shortcut,
  toAccelerator,
} from './commands/index.ts';
export {
  type GridRange,
  MAX_COLUMN,
  MAX_ROW,
  rangeOf,
  Selection,
  type SelectionState,
} from './session/selection.ts';
export { Axis } from './grid/axis.ts';
export { type Band, HEADER_HEIGHT, HEADER_WIDTH, type Hit, type Rect, Viewport } from './grid/viewport.ts';
export { darkPalette, type GridPalette, lightPalette } from './grid/palette.ts';
export { chipColors } from './grid/chip.ts';
export {
  checkboxRect,
  type DisplayText,
  fillHandleRect,
  type GridOverlay,
  MARQUEE_DASH,
  overflowEdge,
  renderGrid,
  type RenderInput,
  type RenderSelection,
  stickyHeaderTitles,
  wrap,
} from './grid/render.ts';
export {
  CELL_PAD_X,
  CELL_PAD_Y,
  cellFont,
  fittedWidth,
  fitText,
  rotatedBox,
  rotatedCellBox,
  textHeight,
} from './grid/cellText.ts';
export {
  canvasMeasure,
  type ColumnFitOptions,
  CSV_FIT_MAX_WIDTH,
  CSV_FIT_ROWS,
  fitCsvColumns,
  fittedColumnWidths,
} from './grid/columnFit.ts';
export { type ChromeButton, chromeButtons, drawChrome, hitChrome } from './grid/chrome.ts';
export { canvasRowMeasure, contentRowHeights, type RowMeasure } from './grid/rowHeights.ts';
export { Grid, type GridHost, type GridController, type GridProps } from './grid/Grid.tsx';
export { Icon } from './components/Icon.tsx';
export { Dialog } from './components/Dialog.tsx';
export { CommandMenu, QuickPick } from './components/CommandMenu.tsx';
export { listStep } from './components/listStep.ts';
export { FormulaBar } from './components/FormulaBar.tsx';
export { SheetTabs } from './components/SheetTabs.tsx';
export { NasBrowser } from './components/NasBrowser.tsx';
export {
  type CellPosition,
  type DocumentFormat,
  formatOf,
  isExcelFile,
  odsNameFor,
  WorkbookSession,
} from './session/workbookSession.ts';
export { csvLosses } from './session/csvLoss.ts';
export {
  calculationBadge,
  clockTime,
  describeCells,
  describeRange,
  narrateEdit,
  narrateHistory,
  pastTense,
  saveStateText,
} from './session/statusText.ts';
export { FormatCellsDialog } from './components/FormatCellsDialog.tsx';
export {
  allSame,
  BORDER_LINE_STYLES,
  BORDER_PRESETS,
  type BorderEdgesOn,
  buildFormatTransform,
  dialogBorder,
  editFormatCells,
  firstVisibleEdge,
  FONT_SIZE_MAX,
  FONT_SIZE_MIN,
  type FormatCellsField,
  type FormatCellsState,
  type FormatCellsValues,
  type FormatCellsVaries,
  formatSample,
  HORIZONTAL_ALIGNMENTS,
  nextTriState,
  NUMBER_FORMAT_PRESETS,
  parseFontSize,
  parseRotation,
  presetIndex,
  sampleValueFor,
  seedFormatCells,
  VARIES,
  VERTICAL_ALIGNMENTS,
} from './components/formatCells.ts';
export { ChartSvg, type ChartSvgProps } from './charts/ChartSvg.tsx';
export {
  CHART_TYPE_LABELS,
  chartFamily,
  type ChartFamily,
  ChartSheetView,
  type ChartSheetViewProps,
  drawAsOptions,
} from './charts/ChartSheetView.tsx';
export { ConditionalFormatDialog } from './components/ConditionalFormatDialog.tsx';
export { printInFrame } from './print/printFrame.ts';
export { draftOf, type PageSetupDraft, setupOf } from './components/pageSetupForm.ts';
export { PageSetupDialog } from './components/PageSetupDialog.tsx';
export { after, debounce, every, nextFrame, wait } from './schedule.ts';
