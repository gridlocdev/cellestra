# Features

Cellestra is a spreadsheet editor for OpenDocument spreadsheets (`.ods`). It runs in the browser, on your NAS and on the desktop, all from the same code. Every command is in one catalog, so the ribbon, the menus, the command menu and the shortcuts always agree.

![Cellestra in dark mode with the Home ribbon open: a small monthly budget with bold column headings Month, Rent, Groceries, Transit and Total, amounts in dollars for January to April, and an Average row whose selected total cell shows the formula =AVERAGE(E2:E5) in the formula bar, with the grid zoomed to 160%](images/budget.webp)

## Three ways to run it

- **In the browser.** Open an `.ods`, `.csv` or `.tsv` file from your computer, edit it, and save it back. Everything happens in the page; nothing is uploaded.
- **On your NAS.** A small server shows the folders you choose, so you can open and save spreadsheets on the NAS from any browser on your network.
- **On the desktop.** An Electron app with native menus, file associations and normal Open and Save dialogs.

See [Getting started](getting-started.md) for each one.

## Files

Cellestra reads and writes `.ods` and keeps anything it doesn't understand intact, so LibreOffice files survive a round trip. **Import CSV…** replaces the active sheet with a CSV or TSV file's data (one undo brings the sheet back); opening or importing a CSV fits its columns to the data (up to 400 px wide, measured on the first 2,000 rows), **Export Sheet as CSV…** writes the sheet out as CSV, and **Export Sheet as PDF…** and **Print…** print a sheet with the settings in **Page Setup…**. **Open Recent…** lists the files you used last. **Close** (`⌘W`) asks about unsaved changes, then closes the window in the desktop app or starts a blank workbook in the browser.

When LibreOffice has a file open, it keeps a `.~lock.<name>#` file beside it. Opening that file in the desktop app or from the NAS reads the lock and warns who opened it and when, for example “Opened by Ada on 14.11.2026 10:22 in LibreOffice – your changes may conflict.” The file still opens, and Cellestra never writes or removes lock files. A plain browser can't see the folder around a file, so it can't warn.

While a file from the NAS is open, Cellestra checks every 30 seconds whether someone else saved it (only while the tab is visible). If you have no unsaved changes it offers to reload the file. If you do, it warns you, and saving then asks whether to overwrite their version, save a copy or reload.

## Formulas

Formulas have 47 functions, automatic recalculation and circular-reference detection. The formula syntax matches LibreOffice Calc. **Insert Function…** picks a function, autocomplete suggests names and arguments as you type, and **Recalculate** (F9) recalculates everything. **Manual Calculation** (or a click on the Auto/Manual badge in the status bar) stops formulas recalculating after each edit until you press F9, which helps in a slow workbook; the mode is saved with the file, and switching back to automatic recalculates at once. **Show Formulas** shows the formulas in place of their results, and **Named Ranges…** names a cell or range.

## Formatting

Number formats (general, number, currency, percent, date, time, scientific and text), fonts, fills, borders, alignment, wrapping and merges. **Format Cells…** has every option in one dialog, **Conditional Formatting…** colours cells by their values, and **Format Painter** copies formatting from one cell to others.

**Clean Up Sheet** and **Clean Up Workbook** read what the sheet holds and format it to match: a styled, frozen header row, a number format per column, alignment by type, columns fitted to their contents, and trimmed whitespace and numbers stored as text. It is one undo step.

## Find and Replace

**Find…** (`⌘F`) and **Find and Replace…** (`⌘H`) search the sheet's shown values and formulas, ignoring case. The dialog lists every match with its cell address and text; click one to go there. **Find Next** and **Previous** step through the matches, **Replace** replaces the match at the active cell and moves on to the next, and **Replace All** replaces every match in one undo step. Text stays text when it is replaced, so `007` doesn't turn into `7`.

## Data tools

- **Sort Ascending**, **Sort Descending** and **Custom Sort…** sort a range by one or more columns, and **Sort & Filter…** chains sorts and filters in one dialog.
- **AutoFilter** adds a filter dropdown to each column. Filtered-out rows are hidden, not deleted.
- **Create Table** turns a range into a named table with a header band and filter buttons. The name comes from the text just above the range or its first header cell when that makes a valid name (`Monthly sales` becomes `Monthly_sales`), otherwise `Table1`, `Table2`… The **Name** box in the Data tab's Table group shows the table the cursor is in and renames it, and typing a table's name in the Name Box selects the table.
- **Data Validation…** limits what a cell accepts, for example a dropdown list or a number range.
- **Remove Duplicates** and **Text to Columns…** tidy imported data.

## Charts

**Chart Sheet** adds a sheet with a chart of the selected data. Charts are saved in the `.ods` file, so LibreOffice opens them too.

A chart can plot a table by its column names instead of cell ranges. It then follows the table: rows added at the bottom show up in the chart, and renaming the table or a column header doesn't break it. Removing the table leaves the chart on the cells it last plotted.

## Sheets and the grid

A workbook has any number of sheets: **Insert Sheet**, **Rename Sheet…**, **Duplicate Sheet** and **Delete Sheet** manage them, and **Go to Sheet…** jumps to one by name. **Freeze Panes** keeps header rows and columns in view, rows and columns can be hidden, resized and fitted with **Autofit Columns**, and **Zoom In** and **Zoom Out** change the view. Scrolled down inside a table, the column headers show the table's column titles in place of A, B, C. Text cut off by a filled neighbour (or its own merged cell) shows a small red chevron at the edge where it is cut.

The status bar under the grid narrates the last edit ("Pasted A1:C4", "Inserted rows", "Undid Sort"), sums the selection, and shows the calculation mode (Auto or Manual), the zoom, and when the file was last saved ("Saved 14:05", or "Edited" once it has changed since).

## The keyboard

Every command can be reached from the keyboard. `⌘K` opens the command menu, which searches every command and shows its shortcut, and `⌘⇧P` switches sheets. **Keyboard Shortcuts** (`⌘/`) lists them all in the app. See [Controls](controls.md) for the full list.

## Accessibility

A screen reader such as VoiceOver hears the active cell's address and what it shows each time the selection moves (`B7, 1,234.50`, or `B7, blank`), with the selected range when there is one, and hears each message the app pops up at the bottom. In a dialog, `Tab` and `⇧Tab` cycle through its controls without leaving it, `Esc` closes it, and focus goes back to where it was before the dialog opened. With **Reduce motion** on in the system settings, the dashed border around copied cells stands still and nothing in the app animates or scrolls smoothly.

## Command line

The `cellestra` command reads and edits `.ods` files from scripts and CI with the same engine as the app: convert to and from CSV, read and set cells, recalculate, check a workbook for formula errors, compare two workbooks, and more. See [Development](development.md#the-cellestra-cli) for how to run it.
