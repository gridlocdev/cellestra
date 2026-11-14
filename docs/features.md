# Features

Cellestra is a spreadsheet editor for OpenDocument spreadsheets (`.ods`). It runs in the browser, on your NAS and on the desktop, all from the same code. Every command is in one catalog, so the ribbon, the menus, the command menu and the shortcuts always agree.

![Cellestra in dark mode with the Home ribbon open: a small monthly budget with bold column headings Month, Rent, Groceries, Transit and Total, amounts in dollars for January to April, and an Average row whose selected total cell shows the formula =AVERAGE(E2:E5) in the formula bar, with the grid zoomed to 160%](images/budget.webp)

## Three ways to run it

- **In the browser.** Open an `.ods`, `.csv` or `.tsv` file from your computer, edit it, and save it back. Everything happens in the page; nothing is uploaded.
- **On your NAS.** A small server shows the folders you choose, so you can open and save spreadsheets on the NAS from any browser on your network.
- **On the desktop.** An Electron app with native menus, file associations and normal Open and Save dialogs.

See [Getting started](getting-started.md) for each one.

## Files

Cellestra reads and writes `.ods` and keeps anything it doesn't understand intact, so LibreOffice files survive a round trip. **Import CSV…** and **Export Sheet as CSV…** bring in and write out CSV and TSV, and **Export Sheet as PDF…** and **Print…** print a sheet with the settings in **Page Setup…**. **Open Recent…** lists the files you used last.

## Formulas

Formulas have 47 functions, automatic recalculation and circular-reference detection. The formula syntax matches LibreOffice Calc. **Insert Function…** picks a function, autocomplete suggests names and arguments as you type, and **Recalculate** recalculates everything. **Show Formulas** shows the formulas in place of their results, and **Named Ranges…** names a cell or range.

## Formatting

Number formats (general, number, currency, percent, date, time, scientific and text), fonts, fills, borders, alignment, wrapping and merges. **Format Cells…** has every option in one dialog, **Conditional Formatting…** colours cells by their values, and **Format Painter** copies formatting from one cell to others.

**Clean Up Sheet** and **Clean Up Workbook** read what the sheet holds and format it to match: a styled, frozen header row, a number format per column, alignment by type, columns fitted to their contents, and trimmed whitespace and numbers stored as text. It is one undo step.

## Data tools

- **Sort Ascending**, **Sort Descending** and **Custom Sort…** sort a range by one or more columns, and **Sort & Filter…** chains sorts and filters in one dialog.
- **AutoFilter** adds a filter dropdown to each column. Filtered-out rows are hidden, not deleted.
- **Create Table** turns a range into a named table with a header band and filter buttons.
- **Data Validation…** limits what a cell accepts, for example a dropdown list or a number range.
- **Remove Duplicates** and **Text to Columns…** tidy imported data.

## Charts

**Chart Sheet** adds a sheet with a chart of the selected data. Charts are saved in the `.ods` file, so LibreOffice opens them too.

## Sheets and the grid

A workbook has any number of sheets: **Insert Sheet**, **Rename Sheet…**, **Duplicate Sheet** and **Delete Sheet** manage them, and **Go to Sheet…** jumps to one by name. **Freeze Panes** keeps header rows and columns in view, rows and columns can be hidden, resized and fitted with **Autofit Columns**, and **Zoom In** and **Zoom Out** change the view.

The status bar under the grid narrates the last edit ("Pasted A1:C4", "Inserted rows", "Undid Sort"), sums the selection, and shows the calculation mode (Auto or Manual), the zoom, and when the file was last saved ("Saved 14:05", or "Edited" once it has changed since).

## The keyboard

Every command can be reached from the keyboard. `⌘K` opens the command menu, which searches every command and shows its shortcut, and `⌘⇧P` switches sheets. **Keyboard Shortcuts** (`⌘/`) lists them all in the app. See [Controls](controls.md) for the full list.

## Command line

The `cellestra` command reads and edits `.ods` files from scripts and CI with the same engine as the app: convert to and from CSV, read and set cells, recalculate, check a workbook for formula errors, compare two workbooks, and more. See [Development](development.md#the-cellestra-cli) for how to run it.
