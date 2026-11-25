# Changelog

## Unreleased

- **Make a table from the chart panel.** A range chart whose cells could be a table offers **Make a Table…**: name it, and the cells become a table the chart plots by column name, in one undo step. **Use Whole Columns** and **Keep These Cells** give a way forward when switching between Table and Range can't go as asked.
- **The chart panel says what's wrong.** A line under the heading says where the data comes from, and a range that can't plot says why under its field.
- **Chart types as pictures.** Types the table's columns can't fill are greyed out with the reason.
- **Shorter series list.** Each series folds to one line and opens when you need its fields. Removing a series shows a notice with **Undo**.
- **Chart menu.** Each chart type, **Show Legend** and **Swap Rows and Columns** are in a new **Chart** menu and the command menu. `F11` inserts a chart sheet.
- **Edit Data Range…** re-plots a chart from one block you select on the sheet, keeping its type and look. It's in the chart panel and the **Chart** menu.
- **See what a range holds.** Each range field shows its first few values underneath, and an empty chart says how to start.
- **Clearer chart messages.** When a chart can't switch to Table mode, the message names the range or series at fault. **Make a Table…** lists the column names the table will get.
- **Whole columns in charts.** A chart range can be a whole column or row, such as `Data.A:A`; it is saved as cells so other apps read it, and the chart stops at the last used row. The CLI's chart options take them too.
- Removing a series no longer opens or folds the series after it.
- Enter after **Shift+Tab** goes straight down instead of back to the column you started in, when that column is to the right.
- **↓** while typing in a cell goes straight down, even after **Tab**; only **Enter** goes back to where the **Tab**s started. Clicking or jumping to another cell forgets the **Tab**s.
- The axis titles and **Show legend** are together in a folded **Labels** group.
- The **Aggregate** list is now **Combine rows with the same label**, and the name cell field is **Take name from cell**.
- **Names for values.** A name can hold a value or a formula, not only a range: type `0.2` or `=B1*2` in **Named Ranges…** and `=A1*TaxRate` works. Such names from LibreOffice and Excel files now calculate instead of showing `#NAME?`, and they are saved as named expressions. Their references follow inserted rows and renamed sheets. The CLI's `name add` and the `name` apply op take `=0.2` too.
- **3D ranges are outlined while you edit a formula.** A range across sheets, such as `Sheet1.A1:Sheet3.B5`, is outlined on every sheet it covers.
- **Moving a sheet keeps 3D ranges in order.** When a sheet move puts a 3D range's last sheet before its first, the formula swaps them, as LibreOffice does.
- `INDEX` with area number 0 gives `#VALUE!`, as LibreOffice does, instead of `#REF!`.
- The sheet tabs use all the space up to the status text, and the active sheet's tab stays in view when there are more tabs than fit.
- **Look up a key.** **Keyboard Shortcuts** (`⌘/`) has **Look Up a Key…**: press any keys to see which command each one runs, with the command highlighted in the list. Keys only show what they do while it listens, so `⌘Q` and `⌘W` are safe to try in the desktop app. `Esc`, **Stop**, switching windows, closing the window or 30 seconds without a key all end it.
- **Desktop app hardening.** The packaged app can't be run as a plain Node.js program, ignores `NODE_OPTIONS` and `--inspect`, and loads only its own checked app code.
- `just install-desktop` asks to quit a running Cellestra, installs, then opens it again, instead of refusing.

## 1.0.0

Cellestra 1.0.0 is the first stable release. Nothing changed since 0.4.0.

## 0.4.0

### What's new

- **Open Excel workbooks.** **Open…**, drag and drop, Open Recent and the desktop app's file associations accept `.xlsx` and `.xlsm`. The file opens as a new, unsaved workbook with the same name ending in `.ods`, and the first save asks where to put it, so the Excel file is never changed. Values, formulas, number formats, fonts, fills, borders, alignment, merges, column widths, row heights, hidden rows and columns, frozen panes, named ranges and dropdown lists come across. A notice lists what was left out, such as charts, comments or conditional formatting, and formulas Cellestra can't calculate keep their last result as a value.
- **Dropdown lists look like chips.** Every cell with a list rule shows its value in a rounded, coloured chip with a chevron, and an empty one shows the chevron. Pick each item's colour in **Data Validation…**: new items get a colour of their own. Click the chevron, or run **Choose from List**, to pick an item from a list of chips. The colours are saved as conditional formats, so LibreOffice shows them as cell fills.
- **Move rows and columns with the Option key.** **⌥↑** and **⌥↓** move the selected rows one step, and **⌥←** and **⌥→** move the selected columns, trading places with their neighbour. Cells, sizes and formatting go with them, formulas follow the moved cells, and the selection moves too. Nothing moves past the edge of the sheet or splits a merge or a table.
- **CLI reads `.xlsx`.** Every command that reads a workbook reads `.xlsx` and `.xlsm` too, and prints what was left out on stderr. `convert book.xlsx book.ods` converts one. Writing to an `.xlsx` (`--in-place`, `-o`, or `convert` to `.xlsx`) is refused, and the file stays as it was.

### Improvements

- **Shortcuts and typing stay in the dialog that's open.** ⌘K and the other shortcuts no longer open behind Settings, **Go to Sheet** or **Open Recent**. A dialog's first field takes focus as it opens, so what you type straight away goes into it instead of the cell behind. Picking a function in **Insert Function** leaves you typing in the cell.
- **Tab and Shift+Tab move the highlight** in the command menu and **Go to Sheet**, like the arrow keys.
- **Saving is steadier.** Closing the window while a save runs waits for it to finish, then closes or asks. Auto-save waits while the close question is open. ⌘S with nothing unsaved writes nothing. **Revert** stops if you type while it reads the file, as a refresh does.
- **Screen readers can step through list popups.** The command menu, **Go to Sheet**, **Open Recent**, **Insert Function** and the **Open from NAS** file list each have a name, and their rows are options.
- If you run the NAS server yourself, it reads its port from `CELLESTRA_PORT`. `PORT` still works when that isn't set.

### Bugs fixed

- **A dropdown list taken from another sheet offers that sheet's items.** It used to read the same cells on the sheet you were on, in the dropdown, in **Data Validation…** and when checking what you typed.
- A date you type keeps the same format after you save and reopen the file.
- A resized empty row at the bottom of a sheet keeps its height after a save.
- A very large number in a date-formatted cell, such as 1000000000, no longer disappears when the file is saved and reopened.
- A failed save in the desktop app or on the NAS server is reported instead of being passed over for the next one.

## 0.3.0

### Before you update

- **If you use the CLI on a workbook set to manual calculation,** commands that change cells now leave formula results alone, as the app does. Run `cellestra recalc` to update them.

### What's new

- **Charts can work from a table.** A chart sheet's panel has a **Table** / **Range** switch. In **Table** mode you pick the table, the column for the categories and the columns to plot from lists of column names, with no typing. The chart grows when the table grows, and keeps working when you rename the table or a column. **Insert Chart Sheet** with the cursor in a table starts a chart that plots its number columns.
- **Pick a range on the sheet.** In **Range** mode, the ⌖ button next to each range opens the data sheet so you can drag over the cells, then **Enter** brings you back to the chart.
- **Repeated categories are combined.** Column, bar, line, area, pie and radar charts show a label that appears on several rows once, adding up its values. The **Aggregate** list switches to Average, Count, Smallest or Largest. LibreOffice still shows every row.
- **Hover over a chart to read it.** A tooltip shows the category, slice or point under the pointer with its values, the total and share where they apply, and how many rows a combined category holds.
- **Tables have names you can see and change.** A new table is named after the text above it or its first header. Rename it in the **Name** field on the table's ribbon tab, or type a table's name in the Name Box to select it.
- **Opening a CSV fits the columns** to their contents (up to 400 pixels wide), without marking the file as changed.
- **CLI:** `rename-table`, a table name in `chart --from`, `--aggregate` for charts, and `remove-duplicates` and `split` in `apply` files.

### Improvements

- **Autofit matches what the grid shows,** using the grid's own font and padding, so fitted columns are no longer a little too wide.
- **Large clears and sheet deletes ask in Cellestra's own dialog** instead of the browser's box.
- **The desktop menu stays in step** with the layout, the collapsed ribbon, the format painter and the find text.
- **Charts leave out empty rows,** and blank labels share one "(blank)" category. Series without a name are called "Series 1", "Series 2" and so on.
- **List dropdowns skip blank items** and show each item once.
- **The CLI reports a refused edit in one line** instead of a stack trace.

### Bugs fixed

- **Import CSV** fills the current sheet again, as one step you can undo, instead of replacing the open workbook. It asks first when the sheet holds more than 100 cells.
- Duplicating a chart sheet no longer makes the copy overwrite the original's chart when you save.
- Date and time formats that carry a locale tag, such as `[$-409]h:mm AM/PM`, are no longer treated as currency.
- `cellestra open` opens the installed app instead of a blank Electron window.
- `cellestra rename-sheet` updates the formulas that name the sheet.
- Replacing a validation rule from the CLI keeps the old rule on the cells outside the new range.
- `cellestra import` into an existing sheet clears its old column widths, merges and other layout.
- `cellestra chart --from` works with sheet names that contain spaces.
- A CLI save over a hard-linked file no longer empties it first, and a read-only file is refused.

## 0.2.3

### Improvements

- **More of a LibreOffice file survives a save.** Row and column groups (and whether they're collapsed), rows and columns repeated on every printed page, manual page breaks, workbook structure protection, and macros bound to document events are all kept. Groups, breaks and print titles move with their rows and columns when you insert or delete.
- **Filtered rows and rows you hid are kept apart.** Clearing a filter no longer shows rows you hid by hand, **Unhide** shows only the rows you hid, and a save tells LibreOffice which is which. SUBTOTAL 1–11 skips filtered rows, and 101–111 also skips rows you hid, as in Calc.
- **Editing is faster on big sheets.** An edit re-measures only the rows it changed, where it used to re-measure the whole sheet. On a sheet of 50,000 rows of wrapped text, that's under a millisecond instead of about 50.
- **Moving around is smoother in the desktop app.** The menu bar is no longer rebuilt on every arrow key, and it no longer closes an open menu when nothing in it changed.

### Bugs fixed

- Dates in a column formatted as dates are saved as dates, not as plain numbers, and CSV export writes them as the grid shows them.
- ROUND keeps large whole numbers exact, and INT gives the same result as Calc for sums like `INT((0.7+0.1)*10)`.
- `VALUE` with an impossible date, such as `9999-12-40`, gives #VALUE! instead of stopping the recalculation.

## 0.2.2

### Improvements

- **Large edits and undo use less memory.** Clearing, sorting or filling checkboxes over whole columns keeps only the cells that hold something for undo, and a very large paste or fill no longer runs out of stack.
- **Files that only look changed aren't reloaded.** When a sync client or a script touches the open file without changing it, your document and its undo history stay. A real change that arrives while a dialog, a cell edit or a save is open is reloaded once that's over.
- **Print Selection on whole columns or rows** prints only the part that holds something, not thousands of empty pages.

### Bugs fixed

- A styled block, or columns formatted below a header row, no longer spreads to the whole row or column when the file is reopened. A header over several formatted columns opens instantly again instead of being refused as too large.
- Closing the window, quitting, or opening another file while you're typing in a cell keeps what you typed.
- Text pasted from another app without tabs goes in as lines in one column, instead of being split at its commas.
- **Overwrite** after a conflict on the NAS still warns when LibreOffice has the file open.
- In a browser that can only download, **Save** in the unsaved-changes question also keeps a recovery copy, which the next start offers, in case the download wasn't kept.
- Undo and redo no longer lose a step when it fails partway.
- Undo brings back a deleted sheet even when its name came from a file and wouldn't be allowed if you typed it.
- A formula calling a function Cellestra doesn't have keeps the value LibreOffice saved, instead of showing `#NAME?`.
- Cut and paste, and a full copy and paste, keep a cell's comment and its mixed text formatting.
- Applying conditional formatting keeps the rules from the file that the dialog can't show.
- Files that repeat cells or merged cells millions of times are refused with a clear message instead of using up all the memory.

## 0.2.1

### What's new

- **Ranges across sheets.** `=SUM(Sheet1.A1:Sheet3.B2)` adds up the same area on every sheet from Sheet1 to Sheet3. COUNT, AVERAGE, MIN, MAX and the other totals read them too.
- **SUBTOTAL.** It totals a range but leaves out rows a filter hid and other subtotals inside it, so a grand total doesn't count twice. Functions 101–111 also leave out rows you hid. Results update as soon as you hide, show or filter rows.
- **SUMPRODUCT** multiplies matching cells of two or more ranges and adds up the products.
- **INDEX picks an area.** Its fourth argument chooses one range from a list joined with `~`, such as `=INDEX((A1:B4~D1:E4), 2, 1, 2)`.
- **A space between two ranges is their overlap**, as in Calc: `=SUM(A1:C5 B4:D8)`.
- **Autofit several columns at once.** Double-click the edge of a selected column's header to fit every selected column to its contents.
- **A None alert for validation rules.** A rule with the **None** response accepts any entry without a word, like a LibreOffice rule with its error message switched off.

### Improvements

- The status line names whole columns and rows as `B:C` and `2:4`, and says which columns an autofit changed.

### Bugs fixed

- Double-clicking a column header's edge autofits in Chrome, Edge and Safari.
- Validation rules whose error message is switched off in LibreOffice no longer reject entries, and a save keeps the message off.
- TRUE and FALSE cells count as 1 and 0 in SUM, COUNT, AVERAGE, MIN, MAX and PRODUCT, as in Calc.
- Formulas that point at another sheet keep pointing at it when LibreOffice copies the cell to a different sheet.

## 0.2.0

### Before you update

- **The NAS server needs a token on your network.** When the server listens on anything other than localhost, as the Docker image does, it now refuses to start without `CELLESTRA_TOKEN`. Set a token in your container settings, or set `CELLESTRA_NO_AUTH=1` if you really want it open to everyone on the network.
- **Find Next is F3 on Windows and Linux.** **Ctrl+G** is now **Go To**. On a Mac, Find Next stays **⌘G**.

### What's new

- **A status bar.** It shows what you just did ("Pasted A1:C4", "Undid Sort"), when the file was last saved, and whether calculation is automatic or manual. Click the badge to switch.
- **Manual calculation.** Turn on **Manual calculation** in the Formulas tab to stop recalculating after each change, including row and column inserts. Press **F9** to bring every result up to date.
- **Find and Replace, one at a time.** **Replace** changes the current match and moves on, and a list shows every match so you can click straight to it.
- **Faster keyboard selection and moves.** **⇧Space** selects rows, **⌃Space** selects columns, **⌃G** goes to a cell or named range, **Esc** shrinks a selection back to one cell, and after a run of **Tab**, **Enter** goes back to the column you started in.
- **Sticky table headers.** When you scroll down inside a table, the column headers show the table's titles.
- **Close.** **File › Close** (⌘W on desktop) closes the document and asks about unsaved changes first. The desktop app has a **Window** menu, and **Quit** in the File menu on Windows and Linux.
- **LibreOffice warnings.** Opening a file that someone has open in LibreOffice says who and since when, on desktop and on the NAS.
- **Outside changes on the NAS.** If someone else saves the file you have open, Cellestra offers to reload it, or warns you before your next save if you have edits of your own.
- **More formula syntax.** Join ranges with `~`, take their overlap with `!`, and write arrays inline, such as `=SUM({1,2;3,4})`. `CONCATENATE` is new.
- **Fraction formats.** Formats such as `# ?/?` and `?/4` show numbers as fractions, and LibreOffice's fraction formats open and save intact.
- **Accessibility.** Screen readers hear the active cell and its value as you move, plus every message the app shows. Dialogs keep keyboard focus inside them, and **Reduce motion** stops the copy marquee's animation.

### Improvements

- **Big files open with much less memory.** A sheet with a million cells now peaks at about 415 MB instead of about 1.5 GB.
- **Files keep more of what LibreOffice wrote.** Comments and shapes anchored to a cell, rich text inside cells, cell styles, currency and percentage values, and error cells all come back as they were saved.
- **The document's language counts.** Number display, sorting and Text to Columns follow the language saved in the file.
- **CSV works like other spreadsheets.** IDs with leading zeros, long numbers and version numbers such as `1.2.3` stay text on import. Text that starts with `=` is written and read as plain text.
- **Sorting and lookups match Calc more closely.** Text sorts by your language's rules, errors sort after TRUE and FALSE, criteria accept `~` to escape wildcards, and exact lookups stop at the first match.
- **Pinch and ⌘-scroll zoom the sheet**, not the whole page.
- **The Name Box jumps to named ranges.**
- **Deleting a sheet** asks first only when something would be lost, says how many formulas will show `#REF!`, and can be undone.
- **Clearing more than 100 filled cells** asks first.
- **Print** gives each row the height its content needs.
- **Tiny numbers** show in scientific notation in General format instead of as 0.
- **Times round their seconds** instead of cutting them off.
- **Undoing** a row or column insert or delete is faster on big sheets.
- **The NAS server checks an open file for changes** with a quick look at the file instead of reading it all.
- If you use the `cellestra` command line: edits saved to `.csv` or `.tsv` files are written as text, `.tsv` uses tabs, and usage errors exit with code 2.

### Bugs fixed

- A cut in one document no longer moves cells when you paste in another.
- Edits typed while a slow file opens are no longer lost.
- Save As and Export work while an auto-save is running.
- Saving in the browser no longer overwrites a file another app saved since you opened it.
- Saving on desktop notices outside edits even when the file's date didn't change.
- The formula bar writes only what you typed, into the cell you were editing.
- Saving a CSV document saves its own sheet, not whichever sheet is active.
- The cell editor stays open when the window loses focus.
- **⇧⌘=** inserts rows instead of zooming in.
- Desktop Save As keeps the right file extension.
- Pasting past the last row or column is refused with a message, instead of losing cells.
- Formulas Cellestra can't read keep their last value.
- SUMIF recalculates when the cells it sums change.
- Sheet names with spaces or symbols are quoted correctly in formulas and in saved files.
- The sheet tab menu has a solid background.
- The NAS server refuses to replace read-only files and keeps hard links intact.
- A slow folder listing in the NAS browser no longer replaces a newer one.
- **Export PDF** asks before replacing a file.
- **Replace All** keeps text cells as text.
- A failed auto-save is reported once instead of on every try.
- Cached values are kept when you open a file in manual calculation mode.
- Text that spills in from a cell scrolled out of view is drawn.
- The first accented or IME character typed into a cell is kept.
- Hyperlinks come along when you copy, cut, paste or fill.
- **Remove Duplicates** ignores case.
- `ROUND` with negative digits, `MOD`, `SQRT` of a negative number, `INDEX` with 0, and TRUE/FALSE in criteria now give Calc's results.
- Other document types and password-protected files get a clear message instead of a broken open.

## 0.1.1

- **Downloads for macOS, Windows and Linux.** The desktop app is now on the [Releases page](https://github.com/gridlocdev/cellestra/releases/latest): a `.dmg` for Apple silicon or Intel Macs, an installer for Windows and an AppImage for Linux. You no longer need to build it from source. The apps aren't signed yet, so [Getting started](https://github.com/gridlocdev/cellestra/blob/main/docs/getting-started.md#download) explains how to open them the first time.

## 0.1.0

The first release. Cellestra is a spreadsheet editor for OpenDocument spreadsheets (`.ods`) that runs in your browser, on your NAS and on the desktop.

- **Edit LibreOffice files without breaking them.** Cellestra reads and writes `.ods` and keeps anything it doesn't understand intact, so files round-trip to LibreOffice. **Import CSV…** and **Export Sheet as CSV…** handle CSV and TSV, and **Export Sheet as PDF…** prints a sheet.
- **Formulas that match Calc.** 47 functions with LibreOffice Calc's syntax, automatic recalculation, circular-reference detection, **Insert Function…** and autocomplete as you type.
- **Formatting in one place.** **Format Cells…** covers number formats, fonts, fills, borders, alignment, wrapping and rotated text. **Conditional Formatting…** colours cells by value, and **Clean Up Sheet** formats a whole sheet to match what it holds in one undo step.
- **Data tools.** Sort, **AutoFilter**, **Create Table**, **Data Validation…**, **Remove Duplicates**, **Text to Columns…** and cell checkboxes.
- **Charts.** **Chart Sheet** charts the selected data, and the chart is saved in the `.ods` so LibreOffice opens it too.
- **Every command from the keyboard.** `⌘K` searches every command and shows its shortcut, `⌘⇧P` switches sheets, and `⌘/` lists all the shortcuts.
- **Three ways to run it.** In the browser, where nothing is uploaded; on a NAS with Docker, opening and saving files in the folders you choose; or as an Electron desktop app with native menus and file associations.
- **A command line.** The `cellestra` command converts, reads, edits, recalculates and checks `.ods` files from scripts and CI with the same engine as the app.
