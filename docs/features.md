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

**Open…** also opens Excel workbooks (`.xlsx` and `.xlsm`). The workbook opens as a new, unsaved file with the same name ending in `.ods` (`Budget.xlsx` becomes `Budget.ods`), and the first save asks where to put it, so the Excel file is never written. Values, formulas, number formats, fonts, fills, borders, alignment, merges, column widths, row heights, hidden rows and columns, frozen panes, named ranges and dropdown lists come across. Charts, images, comments, conditional formatting, other data validation rules, pivot tables, AutoFilters and macros are left out, and a notice lists the ones the file had. A formula that uses a function Cellestra doesn't have keeps its last result as a plain value, and hidden sheets are shown.

While a file from the NAS is open, Cellestra checks every 30 seconds whether someone else saved it (only while the tab is visible). If you have no unsaved changes it offers to reload the file. If you do, it warns you, and saving then asks whether to overwrite their version, save a copy or reload.

## Formulas

Formulas have 47 functions, automatic recalculation and circular-reference detection. The formula syntax matches LibreOffice Calc. **Insert Function…** picks a function, autocomplete suggests names and arguments as you type, and **Recalculate** (F9) recalculates everything. **Manual Calculation** (or a click on the Auto/Manual badge in the status bar) stops formulas recalculating after each edit until you press F9, which helps in a slow workbook; the mode is saved with the file, and switching back to automatic recalculates at once. **Show Formulas** shows the formulas in place of their results, and **Named Ranges…** names a cell or range, or a value such as `TaxRate` = `0.2` that formulas can use (`=A1*TaxRate`).

## Formatting

Number formats (general, number, currency, percent, date, time, scientific and text), fonts, fills, borders, alignment, wrapping and merges. **Format Cells…** has every option in one dialog, **Conditional Formatting…** colours cells by their values, and **Format Painter** copies formatting from one cell to others.

**Clean Up Sheet** and **Clean Up Workbook** read what the sheet holds and format it to match: a styled, frozen header row, a number format per column, alignment by type, columns fitted to their contents, and trimmed whitespace and numbers stored as text. It is one undo step.

## Find and Replace

**Find…** (`⌘F`) and **Find and Replace…** (`⌘H`) search the sheet's shown values and formulas, ignoring case. The dialog lists every match with its cell address and text; click one to go there. **Find Next** and **Previous** step through the matches, **Replace** replaces the match at the active cell and moves on to the next, and **Replace All** replaces every match in one undo step. Text stays text when it is replaced, so `007` doesn't turn into `7`.

## Data tools

- **Sort Ascending**, **Sort Descending** and **Custom Sort…** sort a range by one or more columns, and **Sort & Filter…** chains sorts and filters in one dialog.
- **AutoFilter** adds a filter dropdown to each column. Filtered-out rows are hidden, not deleted.
- **Create Table** turns a range into a named table with a header band and filter buttons. The name comes from the text just above the range or its first header cell when that makes a valid name (`Monthly sales` becomes `Monthly_sales`), otherwise `Table1`, `Table2`… The **Name** box in the Data tab's Table group shows the table the cursor is in and renames it, and typing a table's name in the Name Box selects the table.
- **Data Validation…** limits what a cell accepts, for example a dropdown list or a number range. A list read from cells skips blank cells and repeats.
- **Dropdown lists show as coloured chips.** Each item can have its own colour, picked in Data Validation. Click a cell's chevron, or use **Choose from List**, to pick an item. LibreOffice shows the colours as cell fills.
- **Remove Duplicates** and **Text to Columns…** tidy imported data.

## Charts

**Chart Sheet** adds a sheet with a chart of the selected data. With the selection in a table, the chart plots the table: its first text column labels the categories and its number columns are the series. Charts are saved in the `.ods` file, so LibreOffice opens them too.

A chart can plot a table by its column names instead of cell ranges. It then follows the table: rows added at the bottom show up in the chart, and renaming the table or a column header doesn't break it. Removing the table leaves the chart on the cells it last plotted.

Rows that repeat a category label, such as several sales for one region, show as one bar or slice. The chart adds them up by default; it can show their average, count, smallest or largest value instead. Empty rows are left out, and rows with a blank label share one "(blank)" category. Scatter and bubble charts plot every row that has data.

Hovering a chart shows what is under the pointer. On a column, bar, line, area or radar chart that is the category nearest the pointer with every series' value there (the hovered bar in bold, a stack's total, and each share on a 100% chart); on a pie or doughnut the slice's value and share; on a scatter or bubble chart the nearest point's X, Y and size. A category that merged repeated rows says so, for example "Sum of 3 rows". On a touch screen, tap to see it.

The chart panel opens with a **Table / Range** switch. In Table mode you pick a table and tick the columns to plot. In Range mode you type cell ranges, or press the pick button beside a field, select the cells on the sheet and press Enter (Esc goes back without changing anything). A whole column or row, such as `Data.A:A` or `Data.2:2`, is written out as cells (`Data.A1:A1048576`) so other spreadsheet apps read it too; the chart plots it only down to the last used row.

A line under the panel's heading says where the chart's data comes from, such as "From table “Sales”: Sales and Cost by Region" or "From cells Data.A2:C7". A range that can't plot says why under its field, for example when it names a sheet that doesn't exist or has fewer values than there are category labels. A range that can plot shows its first few values under the field, such as "January, February", so you can see you picked the right cells. A chart with nothing plotted yet says how to start.

When a range chart's cells could be a table, the panel offers **Make a Table…**. It asks for the table's name, filled in from the labels above the data, and lists the column names the table will get (or says a header row is added when the data has none), then makes the cells a table and plots its columns by name, all as one undo step. Choosing **Table** for cells that aren't a table yet opens the same form. When the ranges sit in a table but cover only some of its rows, **Use Whole Columns** plots the whole columns instead. A chart whose table was removed offers **Keep These Cells** to go on plotting the cells it last showed.

The chart type is a row of pictures; arrow keys move between them, and Space or Enter picks one. A type the table's columns can't fill, such as bubble with fewer than three number columns, is greyed out and its tooltip says why. Each series in Range mode folds to one line with its name and values range, and opens to show its fields; a series with a problem stays open. **Remove** on a series shows a notice with **Undo**. The axis titles and **Show legend** sit in a folded **Labels** group.

**Edit Data Range…** in the panel and the **Chart** menu re-plots the chart from one block of cells: it shows the chart's data selected on its sheet, you select the block (labels in the first column, a series in each other column) and press Enter. The chart keeps its type, title and look, and plots a table by column name when you pick whole columns of it, header included. **Undo** puts the old data back.

The **Chart** menu and the command menu have each chart type, **Show Legend**, **Edit Data Range…** and **Swap Rows and Columns**, which turns the series into categories and the categories into series. `F11` inserts a chart sheet of the selection.

## Sheets and the grid

A workbook has any number of sheets: **Insert Sheet**, **Rename Sheet…**, **Duplicate Sheet** and **Delete Sheet** manage them, and **Go to Sheet…** jumps to one by name. **Freeze Panes** keeps header rows and columns in view, rows and columns can be hidden, resized and fitted with **Autofit Columns**, and **Zoom In** and **Zoom Out** change the view. Scrolled down inside a table, the column headers show the table's column titles in place of A, B, C. Text cut off by a filled neighbour (or its own merged cell) shows a small red chevron at the edge where it is cut.

The status bar under the grid narrates the last edit ("Pasted A1:C4", "Inserted rows", "Undid Sort"), sums the selection, and shows the calculation mode (Auto or Manual), the zoom, and when the file was last saved ("Saved 14:05", or "Edited" once it has changed since).

## The keyboard

Every command can be reached from the keyboard. `⌘K` opens the command menu, which searches every command and shows its shortcut, and `⌘⇧P` switches sheets. **Keyboard Shortcuts** (`⌘/`) lists them all in the app, and **Look Up a Key…** there tells you what any key does. See [Controls](controls.md) for the full list.

## Accessibility

A screen reader such as VoiceOver hears the active cell's address and what it shows each time the selection moves (`B7, 1,234.50`, or `B7, blank`), with the selected range when there is one, and hears each message the app pops up at the bottom. In a dialog, `Tab` and `⇧Tab` cycle through its controls without leaving it, `Esc` closes it, and focus goes back to where it was before the dialog opened. Focus is always drawn in the app's own focus colour, in light and dark. In the sheet tabs, ribbon tabs, toolbar, menus and lists the arrow keys move the current item and the focus with it, and `F6` and `Shift+F6` jump between the ribbon (or toolbar), the grid and the sheet tabs. With **Reduce motion** on in the system settings, the dashed border around copied cells stands still and nothing in the app animates or scrolls smoothly.

## Command line

The `cellestra` command reads and edits `.ods` files from scripts and CI with the same engine as the app: convert to and from CSV, convert `.xlsx` to `.ods`, read and set cells, recalculate, check a workbook for formula errors, compare two workbooks, and more. See [Development](development.md#the-cellestra-cli) for how to run it.
