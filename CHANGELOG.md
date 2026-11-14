# Changelog

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
