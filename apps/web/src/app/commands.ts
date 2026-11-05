import {
  Borders,
  BorderEdges,
  CellRange,
  NumberFormats,
  type CellStyle,
  type HorizontalAlignment,
  type VerticalAlignment,
} from '@cellestra/core';
import { CommandCatalog, type CommandSpec } from '@cellestra/ui';
import type { AppController } from './controller.ts';

/**
 * The one command catalog (PRD §9): the native menu, toolbar, ⌘K palette, and key bindings all read
 * it. Ids match the old desktop app's so muscle memory and docs carry over.
 */
export function buildCommands(app: AppController): CommandCatalog {
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
      description: 'Open a spreadsheet (.ods, .csv)',
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
      hidden: !!app.native,
      enabled: () => app.nasAvailable.value,
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
      description: 'Open a CSV or TSV file',
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
      description: 'Select the whole sheet',
      group: 'Edit',
      shortcut: 'mod+a',
      run: () => app.selection.selectRange({ top: 0, left: 0, bottom: 1_048_575, right: 16_383 }),
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
      shortcut: 'mod+g',
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
      run: () => void (app.zoom.value = Math.min(4, round(app.zoom.value + 0.1))),
      global: true,
    },
    {
      id: 'view.zoomOut',
      title: 'Zoom Out',
      description: 'Make the grid smaller',
      group: 'View',
      shortcut: 'mod+-',
      icon: 'zoom-out',
      run: () => void (app.zoom.value = Math.max(0.25, round(app.zoom.value - 0.1))),
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
      run: numberFormat('YYYY-MM-DD', 'Date Format'),
    },
    {
      id: 'format.asTime',
      title: 'Time Format',
      description: '13:45:00',
      group: 'Format',
      shortcut: 'mod+shift+2',
      run: numberFormat('HH:MM:SS', 'Time Format'),
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
  return new CommandCatalog(specs);
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
