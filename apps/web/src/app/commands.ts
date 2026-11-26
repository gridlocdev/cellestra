import {
  Borders,
  ChartEditCommands,
  BorderEdges,
  CellRange,
  NumberFormats,
  type CellStyle,
  type HorizontalAlignment,
  type VerticalAlignment,
} from '@cellestra/core';
import {
  CHART_TYPE_ICONS,
  CHART_TYPE_LABELS,
  CommandCatalog,
  focusRegion,
  isMac,
  type CommandSpec,
} from '@cellestra/ui';
import type { AppController } from './controller.ts';
import { MAX_ZOOM, MIN_ZOOM, NAS_BUILD } from './controller.ts';

/**
 * The one command catalog (PRD §9): the native menu, toolbar, ⌘K palette, and key bindings all read
 * it. Ids match the old desktop app's so muscle memory and docs carry over.
 */
export function buildCommands(app: AppController, mac = isMac): CommandCatalog {
  const s = app.session;
  const font =
    (f: Partial<CellStyle['font']>) =>
    (st: CellStyle): CellStyle => ({ ...st, font: { ...st.font, ...f } });
  const toggleFont = (key: 'bold' | 'italic' | 'underline' | 'strikethrough', title: string) => () => {
    const on = !app.activeStyle.font[key];
    app.style(title, font({ [key]: on }));
  };
  const align = (horizontal: HorizontalAlignment) => () =>
    app.style('Align', (st) => ({
      ...st,
      alignment: {
        ...st.alignment,
        horizontal: st.alignment.horizontal === horizontal ? 'general' : horizontal,
      },
    }));
  const valign = (vertical: VerticalAlignment) => () =>
    app.style('Align', (st) => ({ ...st, alignment: { ...st.alignment, vertical } }));
  const numberFormat = (code: string, title: string) => () =>
    app.style(title, (st) => ({ ...st, numberFormat: NumberFormats.of(code) }));
  const decimals = (delta: number) => () =>
    app.style(delta > 0 ? 'Increase Decimals' : 'Decrease Decimals', (st) => ({
      ...st,
      numberFormat: NumberFormats.of(stepDecimals(st.numberFormat.code, delta)),
    }));
  const isActive = (key: 'bold' | 'italic' | 'underline' | 'strikethrough') => () => (
    s.version.value,
    app.selection.state.value,
    app.activeStyle.font[key]
  );
  const edge = BorderEdges.of('thin');
  const setEdges =
    (e: Partial<CellStyle['border']>) =>
    (st: CellStyle): CellStyle => ({ ...st, border: { ...st.border, ...e } });
  const sub = (t: number, l: number, b: number, r: number) => CellRange.fromBounds(t, l, b, r);
  const hasDoc = () => s.document.value !== null;
  /** Re-read on every edit or cursor move, for enabled/active checks. */
  const live =
    <T>(f: () => T) =>
    () => (s.version.value, s.activeSheet.value, app.selection.state.value, f());
  const inTable = live(() => app.activeTable !== null);
  const inList = live(() => {
    const { row, column } = app.selection.active.value;
    return app.sheet.findValidation(row, column)?.kind === 'list';
  });

  const specs: CommandSpec[] = [
    // File
    {
      id: 'file.new',
      title: 'New',
      description: 'Start a blank workbook',
      group: 'File',
      shortcut: 'mod+n',
      icon: 'file-plus',
      run: () => app.newDocument(),
      global: true,
    },
    {
      id: 'file.open',
      title: 'Open…',
      description: 'Open a spreadsheet (.ods, .csv, .xlsx)',
      group: 'File',
      shortcut: 'mod+o',
      icon: 'folder-open',
      run: () => app.open(),
      global: true,
    },
    {
      id: 'file.openLocal',
      title: 'Open from This Computer…',
      description: 'Upload a file from this computer',
      group: 'File',
      icon: 'file-up',
      run: () => app.openLocal(),
      // Only a NAS build has a second place to open from; elsewhere Open already uploads.
      hidden: !NAS_BUILD || !!app.native,
      enabled: () => app.nasAvailable.value,
    },
    {
      id: 'file.close',
      title: 'Close',
      description: app.native ? 'Close the window' : 'Close the workbook and start a blank one',
      group: 'File',
      shortcut: 'mod+w',
      icon: 'x',
      run: () => app.closeDocument(),
      global: true,
    },
    {
      id: 'file.save',
      title: 'Save',
      description: 'Save the workbook',
      group: 'File',
      shortcut: 'mod+s',
      icon: 'save',
      run: () => void app.save(),
      global: true,
    },
    {
      id: 'file.saveAs',
      title: 'Save As…',
      description: 'Save a copy under a new name',
      group: 'File',
      shortcut: 'mod+shift+s',
      run: () => void app.saveAs('ods'),
      global: true,
    },
    {
      id: 'file.download',
      title: 'Download a Copy',
      description: 'Save an .ods copy to this computer',
      group: 'File',
      icon: 'file-down',
      run: () => app.downloadCopy(),
      hidden: !!app.native,
    },
    {
      id: 'file.revert',
      title: 'Revert to Saved',
      description: 'Reload the file and drop unsaved changes',
      group: 'File',
      icon: 'refresh-cw',
      run: () => app.reloadFromStore(),
      enabled: hasDoc,
    },
    {
      id: 'file.importCsv',
      title: 'Import CSV…',
      description: 'Replace the active sheet with a CSV or TSV file’s data',
      group: 'File',
      run: () => app.importCsv(),
    },
    {
      id: 'file.exportCsv',
      title: 'Export Sheet as CSV…',
      description: 'Save the active sheet’s values as CSV',
      group: 'File',
      icon: 'file-text',
      run: () => void app.saveAs('csv'),
    },
    {
      id: 'file.pageSetup',
      title: 'Page Setup…',
      description: 'Paper, margins, scale and what prints for this sheet',
      group: 'File',
      icon: 'page-setup',
      run: () => void (app.dialog.value = { kind: 'pageSetup' }),
    },
    {
      id: 'file.print',
      title: 'Print…',
      description: 'Print the active sheet',
      group: 'File',
      shortcut: 'mod+p',
      icon: 'printer',
      run: () => void app.print(),
      global: true,
    },
    {
      id: 'file.printSelection',
      title: 'Print Selection…',
      description: 'Print only the selected cells',
      group: 'File',
      run: () => void app.print(true),
    },
    {
      id: 'file.exportPdf',
      title: 'Export Sheet as PDF…',
      description: 'Save the active sheet as a PDF, as it prints',
      group: 'File',
      icon: 'file-down',
      run: () => void app.exportPdf(),
    },
    {
      id: 'file.recent',
      title: 'Open Recent…',
      description: 'Reopen a recent file',
      group: 'File',
      icon: 'history',
      run: () => void (app.recentOpen.value = true),
      enabled: () => app.recents.value.length > 0,
    },

    // Edit
    {
      id: 'edit.undo',
      title: 'Undo',
      description: 'Undo the last edit',
      group: 'Edit',
      shortcut: 'mod+z',
      icon: 'undo',
      run: () => s.undo(),
      enabled: () => s.canUndo.value,
    },
    {
      id: 'edit.redo',
      title: 'Redo',
      description: 'Redo the last undone edit',
      group: 'Edit',
      shortcut: 'mod+shift+z',
      alternates: ['mod+y'],
      icon: 'redo',
      run: () => s.redo(),
      enabled: () => s.canRedo.value,
    },
    {
      id: 'edit.cut',
      title: 'Cut',
      description: 'Cut the selection',
      group: 'Edit',
      shortcut: 'mod+x',
      icon: 'scissors',
      run: () => app.copy(true),
    },
    {
      id: 'edit.copy',
      title: 'Copy',
      description: 'Copy the selection',
      group: 'Edit',
      shortcut: 'mod+c',
      icon: 'copy',
      run: () => app.copy(false),
    },
    {
      id: 'edit.paste',
      title: 'Paste',
      description: 'Paste the clipboard',
      group: 'Edit',
      shortcut: 'mod+v',
      icon: 'clipboard-paste',
      run: () => app.pasteFromMenu('all'),
    },
    {
      id: 'edit.pasteSpecial',
      title: 'Paste Values',
      description: 'Paste values only, no formulas or formats',
      group: 'Edit',
      shortcut: 'mod+shift+v',
      icon: 'clipboard',
      run: () => app.pasteFromMenu('values'),
    },
    {
      id: 'edit.delete',
      title: 'Clear Contents',
      description: 'Clear values and formulas',
      group: 'Edit',
      icon: 'eraser',
      run: () => app.clearContents(),
    },
    {
      id: 'edit.clearAll',
      title: 'Clear All',
      description: 'Clear contents and formatting',
      group: 'Edit',
      run: () => app.clearAll(),
    },
    {
      id: 'edit.clearFormats',
      title: 'Clear Formatting',
      description: 'Reset the selection’s formatting',
      group: 'Edit',
      icon: 'remove-formatting',
      run: () => app.clearFormats(),
    },
    {
      id: 'edit.selectAll',
      title: 'Select All',
      description: 'Select the table you are in, then the whole sheet',
      group: 'Edit',
      shortcut: 'mod+a',
      run: () => app.selectAll(),
    },
    {
      id: 'edit.selectRow',
      title: 'Select Row',
      description: 'Select the whole rows of the selection',
      group: 'Edit',
      shortcut: 'shift+space',
      run: () => app.selection.selectLines('rows'),
    },
    {
      id: 'edit.selectColumn',
      title: 'Select Column',
      description: 'Select the whole columns of the selection',
      group: 'Edit',
      shortcut: 'ctrl+space',
      run: () => app.selection.selectLines('columns'),
    },
    {
      id: 'edit.goTo',
      title: 'Go To…',
      description: 'Type a cell, range, sheet or name in the Name Box and go there',
      group: 'Edit',
      // ⌃G on macOS and Ctrl+G elsewhere, as in Excel and Calc. (Not F5: in a browser that reloads.)
      shortcut: 'ctrl+g',
      keywords: 'name box jump cell reference navigate',
      run: () => app.focusNameBox(),
    },
    {
      id: 'edit.find',
      title: 'Find…',
      description: 'Find text in the sheet',
      group: 'Edit',
      shortcut: 'mod+f',
      icon: 'search',
      run: () => void (app.dialog.value = { kind: 'find' }),
      global: true,
    },
    {
      id: 'edit.findReplace',
      title: 'Find and Replace…',
      description: 'Replace text in the sheet',
      group: 'Edit',
      shortcut: 'mod+h',
      icon: 'replace',
      run: () => void (app.dialog.value = { kind: 'find' }),
      global: true,
    },
    {
      id: 'edit.findNext',
      title: 'Find Next',
      description: 'Go to the next match',
      group: 'Edit',
      // ⌘G on macOS. Elsewhere Ctrl+G is Go To, so Find Next takes F3, as browsers and editors do.
      shortcut: mac ? 'mod+g' : 'f3',
      run: () => void (app.findNext(app.findQuery.value) || app.info('No matches.')),
      enabled: () => app.findQuery.value !== '',
    },
    {
      id: 'edit.autoSum',
      title: 'AutoSum',
      description: 'Sum the numbers above or to the left',
      group: 'Edit',
      shortcut: 'alt+=',
      icon: 'sigma',
      run: () => app.autoSum(),
    },
    {
      id: 'fill.down',
      title: 'Fill Down',
      description: 'Copy the top row down the selection',
      group: 'Edit',
      shortcut: 'mod+d',
      icon: 'arrow-down-to-line',
      run: () => app.fill('down'),
    },
    {
      id: 'fill.right',
      title: 'Fill Right',
      description: 'Copy the left column across the selection',
      group: 'Edit',
      shortcut: 'mod+r',
      icon: 'arrow-right-to-line',
      run: () => app.fill('right'),
    },
    {
      id: 'edit.formatPainter',
      title: 'Format Painter',
      description: 'Copy formatting to the next selection',
      group: 'Edit',
      icon: 'paintbrush',
      run: () => app.armFormatPainter(),
      active: () => app.formatPainter.value !== null,
    },
    {
      id: 'edit.deleteRows',
      title: 'Delete Rows',
      description: 'Delete the selected rows',
      group: 'Edit',
      shortcut: 'mod+alt+-',
      run: () => app.deleteRows(),
    },
    {
      id: 'edit.deleteColumns',
      title: 'Delete Columns',
      description: 'Delete the selected columns',
      group: 'Edit',
      run: () => app.deleteColumns(),
    },
    {
      id: 'edit.moveRowsUp',
      title: 'Move Rows Up',
      description: 'Move the selected rows up one row',
      group: 'Edit',
      shortcut: 'alt+up',
      icon: 'move-vertical',
      keywords: 'reorder shift swap',
      run: () => app.moveLines('row', -1),
    },
    {
      id: 'edit.moveRowsDown',
      title: 'Move Rows Down',
      description: 'Move the selected rows down one row',
      group: 'Edit',
      shortcut: 'alt+down',
      icon: 'move-vertical',
      keywords: 'reorder shift swap',
      run: () => app.moveLines('row', 1),
    },
    {
      id: 'edit.moveColumnsLeft',
      title: 'Move Columns Left',
      description: 'Move the selected columns left one column',
      group: 'Edit',
      shortcut: 'alt+left',
      icon: 'move-horizontal',
      keywords: 'reorder shift swap',
      run: () => app.moveLines('column', -1),
    },
    {
      id: 'edit.moveColumnsRight',
      title: 'Move Columns Right',
      description: 'Move the selected columns right one column',
      group: 'Edit',
      shortcut: 'alt+right',
      icon: 'move-horizontal',
      keywords: 'reorder shift swap',
      run: () => app.moveLines('column', 1),
    },

    // View
    {
      id: 'view.commandMenu',
      title: 'Command Menu',
      description: 'Search every command',
      group: 'View',
      shortcut: 'mod+k',
      icon: 'command',
      run: () => void (app.paletteOpen.value = true),
      hidden: true,
      global: true,
    },
    {
      id: 'view.goToSheet',
      title: 'Go to Sheet…',
      description: 'Switch to another sheet by name',
      group: 'View',
      shortcut: 'mod+shift+p',
      icon: 'layers',
      run: () => void (app.sheetPickerOpen.value = true),
      global: true,
    },
    {
      id: 'view.nextRegion',
      title: 'Next Region',
      description: 'Move the focus to the next part of the window: ribbon, grid, sheet tabs',
      group: 'View',
      shortcut: 'f6',
      keywords: 'focus pane area ribbon toolbar sheet tabs grid keyboard',
      run: () => void focusRegion(document, 1),
    },
    {
      id: 'view.previousRegion',
      title: 'Previous Region',
      description: 'Move the focus to the previous part of the window',
      group: 'View',
      shortcut: 'shift+f6',
      keywords: 'focus pane area ribbon toolbar sheet tabs grid keyboard',
      run: () => void focusRegion(document, -1),
    },
    {
      id: 'view.freezePanes',
      title: 'Freeze Panes',
      description: 'Freeze rows above and columns left of the cursor',
      group: 'View',
      icon: 'panel-top',
      run: () => app.toggleFreeze(),
      active: () => (s.version.value, app.sheet.frozenRows > 0 || app.sheet.frozenColumns > 0),
    },
    {
      id: 'view.zoomIn',
      title: 'Zoom In',
      description: 'Make the grid bigger',
      group: 'View',
      shortcut: 'mod+=',
      alternates: ['mod++'],
      icon: 'zoom-in',
      run: () => void (app.zoom.value = Math.min(MAX_ZOOM, round(app.zoom.value + 0.1))),
      global: true,
    },
    {
      id: 'view.zoomOut',
      title: 'Zoom Out',
      description: 'Make the grid smaller',
      group: 'View',
      shortcut: 'mod+-',
      icon: 'zoom-out',
      run: () => void (app.zoom.value = Math.max(MIN_ZOOM, round(app.zoom.value - 0.1))),
      global: true,
    },
    {
      id: 'view.zoomReset',
      title: 'Actual Size',
      description: 'Reset zoom to 100%',
      group: 'View',
      shortcut: 'mod+0',
      run: () => void (app.zoom.value = 1),
      global: true,
    },
    {
      id: 'view.showFormulas',
      title: 'Show Formulas',
      description: 'Show formulas instead of results',
      group: 'View',
      shortcut: 'mod+`',
      icon: 'code',
      run: () => void (app.showFormulas.value = !app.showFormulas.value),
      active: () => app.showFormulas.value,
    },
    {
      id: 'view.theme',
      title: 'Toggle Dark Mode',
      description: 'Switch between light and dark',
      group: 'View',
      icon: 'circle-dot',
      run: () => void (app.theme.value = app.isDark.value ? 'light' : 'dark'),
      active: () => app.isDark.value,
    },
    {
      id: 'view.toggleLayout',
      title: 'Ribbon Layout',
      description: 'Switch between the tabbed ribbon and the compact toolbar',
      group: 'View',
      icon: 'layout-grid',
      keywords: 'toolbar bar compact tabs layout',
      run: () => void (app.layout.value = app.layout.value === 'ribbon' ? 'bar' : 'ribbon'),
      active: () => app.layout.value === 'ribbon',
    },
    {
      id: 'view.toggleRibbon',
      title: 'Collapse Ribbon',
      description: 'Show only the ribbon tabs; click a tab to open it',
      group: 'View',
      icon: 'chevrons-up',
      keywords: 'hide minimize ribbon',
      enabled: () => app.layout.value === 'ribbon',
      run: () => void (app.ribbonCollapsed.value = !app.ribbonCollapsed.value),
      active: () => app.ribbonCollapsed.value,
    },
    {
      id: 'app.preferences',
      title: 'Settings…',
      description: 'Theme and preferences',
      group: 'View',
      shortcut: 'mod+,',
      icon: 'settings-2',
      run: () => void (app.dialog.value = { kind: 'settings' }),
      global: true,
    },

    // Insert
    {
      id: 'insert.rowsAbove',
      title: 'Insert Rows Above',
      description: 'Insert as many rows as are selected',
      group: 'Insert',
      shortcut: 'mod+shift+=',
      icon: 'rows-plus',
      run: () => app.insertRows(),
    },
    {
      id: 'insert.columnsLeft',
      title: 'Insert Columns Left',
      description: 'Insert as many columns as are selected',
      group: 'Insert',
      icon: 'columns-plus',
      run: () => app.insertColumns(),
    },
    {
      id: 'insert.sheet',
      title: 'Insert Sheet',
      description: 'Add a new sheet',
      group: 'Insert',
      shortcut: 'shift+f11',
      icon: 'plus',
      run: () => s.addSheet(),
    },
    {
      id: 'insert.function',
      title: 'Insert Function…',
      description: 'Pick a function and start a formula with it',
      group: 'Insert',
      shortcut: 'ctrl+alt+f',
      icon: 'square-function',
      run: () => (app.dialog.value = { kind: 'functions' }),
    },
    {
      id: 'insert.hyperlink',
      title: 'Hyperlink…',
      description: 'Add, change, or remove the link on the active cell',
      group: 'Insert',
      shortcut: 'mod+alt+k',
      icon: 'link',
      run: () => {
        const { row, column } = app.selection.active.value;
        app.dialog.value = { kind: 'hyperlink', row, column };
      },
    },
    {
      id: 'insert.checkbox',
      title: 'Checkbox',
      description: 'Turn the selected cells into checkboxes, or back into plain cells',
      group: 'Insert',
      icon: 'square-check',
      active: () => (s.version.value, app.selection.active.value, !!app.activeStyle.checkbox),
      run: () => app.toggleCheckboxStyle(),
    },
    {
      id: 'insert.namedRange',
      title: 'Named Ranges…',
      description: 'Name a range so formulas can refer to it',
      group: 'Insert',
      icon: 'tag',
      run: () => (app.dialog.value = { kind: 'namedRanges' }),
    },
    {
      id: 'insert.chartSheet',
      title: 'Chart Sheet',
      description: 'Chart the selected data on a new sheet',
      group: 'Insert',
      shortcut: 'f11',
      icon: 'chart-column',
      run: () => {
        try {
          app.insertChartSheet(
            ChartEditCommands.chartFromRange(s.workbook.value, app.sheetName, app.regionOrSelection),
          );
        } catch (e) {
          app.error(e);
        }
      },
    },
    // Chart: the active chart sheet's type, legend and layout; disabled on sheets of cells.
    ...CHART_TYPE_LABELS.map(([type, label]) => ({
      id: `chart.type.${type}`,
      title: `${label} Chart`,
      description: `Draw this chart as a ${label.toLowerCase()} chart`,
      group: 'Chart' as const,
      icon: CHART_TYPE_ICONS[type],
      active: () => app.activeChart?.type === type,
      enabled: () => (
        s.version.value,
        app.activeChart !== null && app.activeChart.type !== type && app.chartTypeAvailable(type)
      ),
      run: () => app.setChartType(type),
    })),
    {
      id: 'chart.legend',
      title: 'Show Legend',
      description: 'Show or hide the key naming each series',
      group: 'Chart',
      icon: 'list',
      active: () => app.activeChart?.showLegend === true,
      enabled: () => (s.version.value, app.activeChart !== null),
      run: () => app.toggleChartLegend(),
    },
    {
      id: 'chart.editDataRange',
      title: 'Edit Data Range…',
      description: 'Pick one block of cells for the chart to plot',
      group: 'Chart',
      icon: 'grid-2x2',
      enabled: () => (s.version.value, app.activeChart !== null),
      run: () => app.startChartDataPick(),
    },
    {
      id: 'chart.swapRowsColumns',
      title: 'Swap Rows and Columns',
      description: 'Plot the series as categories and the categories as series',
      group: 'Chart',
      icon: 'arrow-left-right',
      enabled: () => (s.version.value, app.activeChart !== null),
      run: () => app.swapChartRowsColumns(),
    },
    {
      id: 'insert.date',
      title: 'Insert Today’s Date',
      description: 'Type today’s date into the cell',
      group: 'Insert',
      shortcut: 'mod+;',
      icon: 'calendar',
      run: () => insertNow(app, 'date'),
    },
    {
      id: 'insert.time',
      title: 'Insert Current Time',
      description: 'Type the current time into the cell',
      group: 'Insert',
      shortcut: 'mod+shift+;',
      run: () => insertNow(app, 'time'),
    },

    // Format
    {
      id: 'format.bold',
      title: 'Bold',
      description: 'Toggle bold',
      group: 'Format',
      shortcut: 'mod+b',
      icon: 'bold',
      run: toggleFont('bold', 'Bold'),
      active: isActive('bold'),
    },
    {
      id: 'format.italic',
      title: 'Italic',
      description: 'Toggle italic',
      group: 'Format',
      shortcut: 'mod+i',
      icon: 'italic',
      run: toggleFont('italic', 'Italic'),
      active: isActive('italic'),
    },
    {
      id: 'format.underline',
      title: 'Underline',
      description: 'Toggle underline',
      group: 'Format',
      shortcut: 'mod+u',
      icon: 'underline',
      run: toggleFont('underline', 'Underline'),
      active: isActive('underline'),
    },
    {
      id: 'format.strikethrough',
      title: 'Strikethrough',
      description: 'Toggle strikethrough',
      group: 'Format',
      shortcut: 'mod+shift+x',
      icon: 'strikethrough',
      run: toggleFont('strikethrough', 'Strikethrough'),
      active: isActive('strikethrough'),
    },
    {
      id: 'format.alignLeft',
      title: 'Align Left',
      description: 'Left-align the text',
      group: 'Format',
      shortcut: 'mod+shift+l',
      icon: 'align-left',
      run: align('left'),
    },
    {
      id: 'format.alignCenter',
      title: 'Align Center',
      description: 'Center the text',
      group: 'Format',
      shortcut: 'mod+shift+e',
      icon: 'align-center',
      run: align('center'),
    },
    {
      id: 'format.alignRight',
      title: 'Align Right',
      description: 'Right-align the text',
      group: 'Format',
      shortcut: 'mod+shift+r',
      icon: 'align-right',
      run: align('right'),
    },
    {
      id: 'format.alignTop',
      title: 'Align Top',
      description: 'Top-align the text',
      group: 'Format',
      icon: 'align-vertical-justify-start',
      run: valign('top'),
    },
    {
      id: 'format.alignMiddle',
      title: 'Align Middle',
      description: 'Center the text vertically',
      group: 'Format',
      icon: 'align-vertical-justify-center',
      run: valign('middle'),
    },
    {
      id: 'format.alignBottom',
      title: 'Align Bottom',
      description: 'Bottom-align the text',
      group: 'Format',
      icon: 'align-vertical-justify-end',
      run: valign('bottom'),
    },
    {
      id: 'format.wrap',
      title: 'Wrap Text',
      description: 'Wrap long text inside the cell',
      group: 'Format',
      icon: 'wrap-text',
      run: () =>
        app.style('Wrap Text', (st) => ({
          ...st,
          alignment: { ...st.alignment, wrapText: !app.activeStyle.alignment.wrapText },
        })),
      active: () => (s.version.value, app.selection.state.value, app.activeStyle.alignment.wrapText),
    },
    {
      id: 'format.merge',
      title: 'Merge Cells',
      description: 'Merge or unmerge the selection',
      group: 'Format',
      icon: 'table-cells-merge',
      run: () => app.merge(),
    },
    {
      id: 'format.asAutomatic',
      title: 'General Format',
      description: 'Show numbers as typed',
      group: 'Format',
      shortcut: 'mod+shift+6',
      run: numberFormat('General', 'General Format'),
      keywords: 'number format automatic',
    },
    {
      id: 'format.asNumber',
      title: 'Number Format',
      description: '1,234.56',
      group: 'Format',
      shortcut: 'mod+shift+1',
      icon: 'binary',
      run: numberFormat('#,##0.00', 'Number Format'),
    },
    {
      id: 'format.asCurrency',
      title: 'Currency Format',
      description: '$1,234.56',
      group: 'Format',
      shortcut: 'mod+shift+4',
      icon: 'dollar-sign',
      run: numberFormat('"$"#,##0.00', 'Currency Format'),
    },
    {
      id: 'format.asPercent',
      title: 'Percent Format',
      description: '12.34%',
      group: 'Format',
      shortcut: 'mod+shift+5',
      icon: 'percent',
      run: numberFormat('0.00%', 'Percent Format'),
    },
    {
      id: 'format.asDate',
      title: 'Date Format',
      description: '2026-11-04',
      group: 'Format',
      shortcut: 'mod+shift+3',
      icon: 'calendar',
      run: numberFormat('yyyy-mm-dd', 'Date Format'),
    },
    {
      id: 'format.asTime',
      title: 'Time Format',
      description: '13:45:00',
      group: 'Format',
      shortcut: 'mod+shift+2',
      run: numberFormat('hh:mm:ss', 'Time Format'),
    },
    {
      id: 'format.asScientific',
      title: 'Scientific Format',
      description: '1.23E+03',
      group: 'Format',
      run: numberFormat('0.00E+00', 'Scientific Format'),
    },
    {
      id: 'format.asText',
      title: 'Text Format',
      description: 'Treat entries as text',
      group: 'Format',
      icon: 'type',
      run: numberFormat('@', 'Text Format'),
    },
    {
      id: 'format.increaseDecimal',
      title: 'Increase Decimals',
      description: 'Show one more decimal place',
      group: 'Format',
      icon: 'move-decimal-right',
      run: decimals(1),
    },
    {
      id: 'format.decreaseDecimal',
      title: 'Decrease Decimals',
      description: 'Show one fewer decimal place',
      group: 'Format',
      icon: 'move-decimal-left',
      run: decimals(-1),
    },
    {
      id: 'format.borderAll',
      title: 'All Borders',
      description: 'Thin borders on every cell edge',
      group: 'Format',
      icon: 'grid-2x2',
      run: () => app.style('Borders', (st) => ({ ...st, border: Borders.all('thin') })),
    },
    {
      id: 'format.borderOutline',
      title: 'Outline Border',
      description: 'Thin border around the selection',
      group: 'Format',
      icon: 'square',
      run: () =>
        app.borders('Borders', (r) => [
          [sub(r.top, r.left, r.top, r.right), setEdges({ top: edge })],
          [sub(r.bottom, r.left, r.bottom, r.right), setEdges({ bottom: edge })],
          [sub(r.top, r.left, r.bottom, r.left), setEdges({ left: edge })],
          [sub(r.top, r.right, r.bottom, r.right), setEdges({ right: edge })],
        ]),
    },
    {
      id: 'format.borderTop',
      title: 'Top Border',
      description: 'Border along the top of the selection',
      group: 'Format',
      run: () =>
        app.borders('Borders', (r) => [[sub(r.top, r.left, r.top, r.right), setEdges({ top: edge })]]),
    },
    {
      id: 'format.borderBottom',
      title: 'Bottom Border',
      description: 'Border along the bottom of the selection',
      group: 'Format',
      run: () =>
        app.borders('Borders', (r) => [
          [sub(r.bottom, r.left, r.bottom, r.right), setEdges({ bottom: edge })],
        ]),
    },
    {
      id: 'format.borderLeft',
      title: 'Left Border',
      description: 'Border along the left of the selection',
      group: 'Format',
      run: () =>
        app.borders('Borders', (r) => [[sub(r.top, r.left, r.bottom, r.left), setEdges({ left: edge })]]),
    },
    {
      id: 'format.borderRight',
      title: 'Right Border',
      description: 'Border along the right of the selection',
      group: 'Format',
      run: () =>
        app.borders('Borders', (r) => [[sub(r.top, r.right, r.bottom, r.right), setEdges({ right: edge })]]),
    },
    {
      id: 'format.borderNone',
      title: 'No Borders',
      description: 'Remove borders',
      group: 'Format',
      icon: 'square-dashed',
      run: () => app.style('Borders', (st) => ({ ...st, border: Borders.none })),
    },
    {
      id: 'format.columnWidth',
      title: 'Column Width…',
      description: 'Set the selected columns’ width',
      group: 'Format',
      icon: 'move-horizontal',
      run: () => app.columnWidth(),
    },
    {
      id: 'format.rowHeight',
      title: 'Row Height…',
      description: 'Set the selected rows’ height',
      group: 'Format',
      icon: 'move-vertical',
      run: () => app.rowHeight(),
    },
    {
      id: 'format.hideRows',
      title: 'Hide Rows',
      description: 'Hide the selected rows',
      group: 'Format',
      shortcut: 'mod+9',
      icon: 'eye-off',
      run: () => app.hide('rows', true),
    },
    {
      id: 'format.unhideRows',
      title: 'Unhide Rows',
      description: 'Show hidden rows in the selection',
      group: 'Format',
      shortcut: 'mod+shift+9',
      icon: 'eye',
      run: () => app.unhideRows(),
    },
    {
      id: 'format.hideColumns',
      title: 'Hide Columns',
      description: 'Hide the selected columns',
      group: 'Format',
      run: () => app.hide('columns', true),
    },
    {
      id: 'format.unhideColumns',
      title: 'Unhide Columns',
      description: 'Show hidden columns in the selection',
      group: 'Format',
      run: () => app.unhideColumns(),
    },
    {
      id: 'format.cleanUpSheet',
      title: 'Clean Up Sheet',
      description: 'Trim stray spaces, fix number-like text, tidy widths',
      group: 'Format',
      icon: 'sparkles',
      run: () => app.cleanUp('sheet'),
    },
    {
      id: 'format.cleanUpWorkbook',
      title: 'Clean Up Workbook',
      description: 'Clean up every sheet',
      group: 'Format',
      run: () => app.cleanUp('workbook'),
    },
    {
      id: 'format.cells',
      title: 'Format Cells…',
      description: 'Number format, font, fill, alignment, and borders',
      group: 'Format',
      shortcut: 'mod+1',
      icon: 'settings-2',
      run: () => (app.dialog.value = { kind: 'formatCells', styles: app.selectionStyles() }),
    },
    {
      id: 'format.conditional',
      title: 'Conditional Formatting…',
      description: 'Format cells by their values: highlight big numbers, matching text, or a formula',
      group: 'Format',
      icon: 'highlighter',
      run: () => (app.dialog.value = { kind: 'conditionalFormats' }),
    },

    // Data
    {
      id: 'data.sortAsc',
      title: 'Sort Ascending',
      description: 'Sort the region by the active column, A→Z',
      group: 'Data',
      icon: 'sort-asc',
      run: () => app.sort(true),
    },
    {
      id: 'data.sortDesc',
      title: 'Sort Descending',
      description: 'Sort the region by the active column, Z→A',
      group: 'Data',
      icon: 'sort-desc',
      run: () => app.sort(false),
    },
    {
      id: 'data.customSort',
      title: 'Custom Sort…',
      description: 'Sort by several columns',
      group: 'Data',
      icon: 'sort',
      run: () => (app.dialog.value = { kind: 'sortFilter', target: app.dataTarget(), sortOnly: true }),
    },
    {
      id: 'data.sortAndFilter',
      title: 'Sort & Filter…',
      description: 'Sort levels and per-column filter conditions in one step',
      group: 'Data',
      shortcut: 'mod+shift+f',
      icon: 'filter',
      run: () => (app.dialog.value = { kind: 'sortFilter', target: app.dataTarget(), sortOnly: false }),
    },
    {
      id: 'data.toggleFilter',
      title: 'AutoFilter',
      description: 'Add or remove filter buttons on the header row',
      group: 'Data',
      icon: 'filter',
      run: () => app.toggleFilter(),
      active: live(() => app.sheet.activeFilter !== null),
    },
    {
      id: 'data.clearFilter',
      title: 'Clear Filter',
      description: 'Show every row again',
      group: 'Data',
      icon: 'filter-x',
      run: () => app.clearFilter(),
      enabled: live(() => {
        const t = app.filterTarget;
        return t !== null && (t.table ? !!t.table.filter : !!app.sheet.activeFilterSpec);
      }),
    },
    {
      id: 'data.createTable',
      title: 'Create Table',
      description: 'Make the region a table with a header row',
      group: 'Data',
      shortcut: 'mod+t',
      icon: 'table',
      run: () => app.createTable(),
      enabled: live(() => app.activeTable === null),
    },
    {
      id: 'data.headerColumn',
      title: 'Table Header Column',
      description: 'Treat the table’s first column as row labels',
      group: 'Data',
      icon: 'panel-left',
      run: () => app.toggleHeaderColumn(),
      enabled: inTable,
      active: live(() => !!app.activeTable?.hasHeaderColumn),
    },
    {
      id: 'data.bandedRows',
      title: 'Table Banded Rows',
      description: 'Stripe alternate table rows',
      group: 'Data',
      icon: 'layers',
      run: () => app.toggleBandedRows(),
      enabled: inTable,
      active: live(() => !!app.activeTable?.hasBandedRows),
    },
    {
      id: 'data.removeTable',
      title: 'Remove Table',
      description: 'Turn the table back into a plain range (cells stay)',
      group: 'Data',
      icon: 'square-x',
      run: () => app.removeTable(),
      enabled: inTable,
    },
    {
      id: 'data.validation',
      title: 'Data Validation…',
      description: 'Limit what can be typed into the selection',
      group: 'Data',
      icon: 'square-check',
      run: () => {
        const { row, column } = app.selection.active.value;
        app.dialog.value = {
          kind: 'validation',
          range: app.range,
          existing: app.sheet.findValidation(row, column) ?? null,
        };
      },
    },
    {
      id: 'data.chooseFromList',
      title: 'Choose from List',
      description: 'Open the selected cell’s dropdown list',
      group: 'Data',
      icon: 'list',
      run: () => {
        app.grid.current?.openList();
      },
      enabled: inList,
    },
    {
      id: 'data.removeDuplicates',
      title: 'Remove Duplicates',
      description: 'Delete repeated rows in the region',
      group: 'Data',
      icon: 'list-x',
      run: () => app.removeDuplicates(),
    },
    {
      id: 'data.textToColumns',
      title: 'Text to Columns…',
      description: 'Split a column on a delimiter',
      group: 'Data',
      icon: 'columns',
      run: () => app.textToColumns(),
    },
    {
      id: 'data.autofitColumns',
      title: 'Autofit Columns',
      description: 'Fit the selected columns to their content',
      group: 'Data',
      icon: 'maximize',
      run: () => app.autofitColumns(),
    },
    {
      id: 'formula.recalc',
      title: 'Recalculate',
      description: 'Recalculate every formula',
      group: 'Data',
      shortcut: 'f9',
      icon: 'refresh-cw',
      run: () => s.recalculateAll(),
    },
    {
      id: 'formula.manualCalc',
      title: 'Manual Calculation',
      description: 'Recalculate only on F9 instead of after every edit',
      group: 'Data',
      icon: 'refresh-cw',
      keywords: 'automatic calculation mode recalc',
      run: () => s.setCalculationMode(s.calculationMode.peek() === 'manual' ? 'automatic' : 'manual'),
      active: () => s.calculationMode.value === 'manual',
    },

    // Sheet
    {
      id: 'sheet.rename',
      title: 'Rename Sheet…',
      description: 'Rename the active sheet',
      group: 'Sheet',
      icon: 'pencil',
      run: () => app.renameSheet(),
    },
    {
      id: 'sheet.duplicate',
      title: 'Duplicate Sheet',
      description: 'Copy the active sheet',
      group: 'Sheet',
      icon: 'copy',
      run: () => s.duplicateSheet(s.activeIndex.value),
    },
    {
      id: 'sheet.delete',
      title: 'Delete Sheet',
      description: 'Delete the active sheet',
      group: 'Sheet',
      icon: 'trash-2',
      run: () => app.deleteSheet(),
      enabled: () => s.sheetNames.value.length > 1,
    },
    {
      id: 'sheet.next',
      title: 'Next Sheet',
      description: 'Go to the next sheet',
      group: 'Sheet',
      shortcut: 'ctrl+pagedown',
      alternates: ['mod+alt+right'],
      run: () => s.setActiveSheet((s.activeIndex.value + 1) % s.sheetNames.value.length),
      global: true,
    },
    {
      id: 'sheet.previous',
      title: 'Previous Sheet',
      description: 'Go to the previous sheet',
      group: 'Sheet',
      shortcut: 'ctrl+pageup',
      alternates: ['mod+alt+left'],
      run: () =>
        s.setActiveSheet((s.activeIndex.value - 1 + s.sheetNames.value.length) % s.sheetNames.value.length),
      global: true,
    },

    {
      id: 'sheet.resetData',
      title: 'Reset Sheet Data',
      description: 'Clear every cell and its formatting on this sheet',
      group: 'Sheet',
      icon: 'eraser',
      run: () => app.resetSheetData(),
    },

    // Help
    {
      id: 'help.shortcuts',
      title: 'Keyboard Shortcuts',
      description: 'Every command and its keys',
      group: 'Help',
      shortcut: 'mod+/',
      icon: 'keyboard',
      run: () => void (app.dialog.value = { kind: 'shortcuts' }),
      global: true,
    },
    {
      id: 'help.about',
      title: 'About Cellestra',
      description: 'Version and license',
      group: 'Help',
      icon: 'info',
      run: () => void (app.dialog.value = { kind: 'about' }),
    },
  ];
  // A chart sheet shows no grid: cell, format, and data commands would write to an invisible A1.
  const onChart = () => s.activeSheet.value.isChartSheet;
  const gridOnly = (c: CommandSpec) =>
    (GRID_GROUPS.has(c.group) && !GRID_EXEMPT.has(c.id)) || GRID_EXTRA.has(c.id);
  const catalog = new CommandCatalog(
    specs.map((c) => {
      if (!gridOnly(c)) return c;
      const enabled = c.enabled;
      return { ...c, enabled: () => !onChart() && (enabled?.() ?? true) };
    }),
  );
  // A command run while a cell is being edited (⌘S, ⌥⌘→, a ribbon button) first commits the
  // edit, so the value lands in the cell it was typed into, on the sheet it was typed on.
  // While a question waits for an answer (save changes?, overwrite?), menu commands stay out of
  // it: running Save or New underneath would leave it unanswered.
  catalog.beforeRun = () => !app.asking && app.commitOpenEdit();
  return catalog;
}

function round(n: number): number {
  return Math.round(n * 100) / 100;
}

/** One more/fewer decimal place in each section of a format code. */
export function stepDecimals(code: string, delta: number): string {
  const base = !code || code.toLowerCase() === 'general' ? '0' : code;
  return base
    .split(';')
    .map((section) => {
      const dot = /\.(0+)/.exec(section);
      if (dot) {
        const n = dot[1]!.length + delta;
        return (
          section.slice(0, dot.index) +
          (n > 0 ? '.' + '0'.repeat(n) : '') +
          section.slice(dot.index + dot[0].length)
        );
      }
      if (delta <= 0) return section;
      const last = Math.max(section.lastIndexOf('0'), section.lastIndexOf('#'));
      return last < 0 ? section : section.slice(0, last + 1) + '.0' + section.slice(last + 1);
    })
    .join(';');
}

function insertNow(app: AppController, what: 'date' | 'time'): void {
  const now = new Date();
  const pad = (n: number) => String(n).padStart(2, '0');
  const text =
    what === 'date'
      ? `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`
      : `${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
  const { row, column } = app.selection.active.value;
  app.commit(row, column, text);
}

const GRID_GROUPS = new Set(['Edit', 'Format', 'Data', 'Insert']);
/** Commands outside those groups that only make sense on a sheet of cells. */
const GRID_EXTRA = new Set(['sheet.resetData', 'file.exportCsv']);
/** Edit/Insert commands that still make sense on a chart sheet. */
const GRID_EXEMPT = new Set([
  'edit.undo',
  'edit.redo',
  'insert.sheet',
  'insert.namedRange',
  'insert.chartSheet',
]);
