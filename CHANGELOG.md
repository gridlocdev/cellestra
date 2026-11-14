# Changelog

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
