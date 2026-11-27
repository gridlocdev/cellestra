# Cellestra glossary

Cellestra is a spreadsheet editor for OpenDocument spreadsheets (`.ods`) that runs in a browser, on a NAS and as a desktop app. These are the words it uses for its own things.

## Language

### Files

**Workbook**: What you open, edit and save: one `.ods` file with its **Sheets**, **Names** and **Tables**. _Avoid_: spreadsheet (for the file), document, file (for the open workbook). The code's `document` is the workbook's link to its file.

**Sheet**: One page of a **Workbook**: a grid of **Cells**, or a **Chart sheet**. _Avoid_: worksheet, tab (for the sheet itself; the **Sheet tab** is its button)

**Chart sheet**: A **Sheet** that holds one **Chart** in place of a grid.

**Saved**: The **Workbook** matches its file. The title and status bar show "Saved" with the time.

**Edited**: The **Workbook** has a file and has changed since it was last saved. _Avoid_: modified, dirty, unsaved (in the UI)

**Not saved**: The **Workbook** has changes and has never had a file, such as a new workbook or an Excel file opened as `.ods`.

**Auto-save**: Saving an **Edited** workbook in place after a set number of seconds without an edit. A **Not saved** workbook waits for `⌘S`.

**Recovery snapshot**: A copy of an open **Workbook** kept in the browser, which Cellestra offers back after a crash. _Avoid_: backup, autosave file

**Conflict**: A save that finds the file changed by someone else since it was opened. You choose to overwrite, save a copy or reload; nothing merges by itself.

**Lock file**: The `.~lock.<name>#` file LibreOffice keeps beside a file it has open. Cellestra reads it to warn you and never writes it.

### The grid

**Cell**: One box of a **Sheet**, addressed by column and row (`B7`). It holds a value or a **Formula**.

**Active cell**: The one **Cell** that typing goes into, drawn with the cell cursor. _Avoid_: current cell, cursor (for the cell)

**Selection**: The **Cells** that commands act on: the **Active cell** and any **Range** around it.

**Range**: A rectangle of **Cells** written as `A1:C4`, or `Data.A1:C4` on another **Sheet**.

**3D range**: A **Range** that spans several **Sheets**, such as `Sheet1.A1:Sheet3.B5`.

**Name**: A word that stands for a **Range** (`Prices` = `Data.C3:C9`) or a value or formula (`TaxRate` = `0.2`). **Named Ranges…** lists them. _Avoid_: defined name. In the file a value name is a named expression; in the UI both kinds are names.

**Table**: A named **Range** with a header row, a band style and filter buttons, which **Formulas** and **Charts** can use by column name. _Avoid_: list, data range (for a table)

**Frozen panes**: Header rows and columns that stay in view while the rest of the **Sheet** scrolls.

### Rules on cells

**Validation rule**: What a **Cell** accepts, set in **Data Validation…**: a **List**, a number range and so on, with an input message and an error message.

**List**: A **Validation rule** that offers fixed items in a dropdown. Each item can have a colour. _Avoid_: dropdown (for the rule), picklist

**Chip**: How a **Cell** under a **List** draws its value: a rounded label in the item's colour, with a chevron.

**Conditional format**: A rule that colours **Cells** by their values, checked in order; the first one that holds formats the cell.

### Formulas

**Formula**: A **Cell**'s contents starting with `=`, in LibreOffice Calc syntax, whose result the cell shows.

**Function**: A named operation a **Formula** calls, such as `SUM`. **Insert Function…** lists them.

**Calculation mode**: Whether **Formulas** recalculate after each edit (**Auto**) or only on `F9` (**Manual**). It is saved with the **Workbook**.

**Circular reference**: A **Formula** that depends on its own result, directly or through other **Cells**. The status bar says so.

### Charts

**Chart**: A picture of data from a **Table** or from **Ranges**, saved in the `.ods` file.

**Series**: One set of values a **Chart** plots, such as one column of numbers.

**Category**: A label along a **Chart**'s axis. Rows that repeat a label become one category.

**Table mode**: A **Chart** that plots a **Table** by column name and follows it as it grows.

**Range mode**: A **Chart** that plots **Ranges** typed or picked on the **Sheet**.

**Data range**: The one block of **Cells** that **Edit Data Range…** re-plots a **Chart** from.

### The window

**Ribbon**: The tabbed strip of controls above the grid (File, Home, Insert, Formulas, Data, Review, View).

**Compact toolbar**: The one-row layout chosen in **Settings…** in place of the **Ribbon**. _Avoid_: toolbar (for the ribbon)

**Formula bar**: The line above the grid that shows and edits the **Active cell**'s contents.

**Name Box**: The box at the left of the **Formula bar** that shows the **Active cell**'s address and goes to an address, **Name** or **Table** typed into it.

**Sheet tab**: The button for a **Sheet** at the bottom of the window.

**Status bar**: The line under the grid that says what the last edit did, sums the **Selection**, and shows the **Calculation mode**, the zoom and the save state.

**Command**: Anything you can run: one entry in Cellestra's list of commands, with a title, a group and maybe a shortcut. The **Ribbon**, the **Command menu**, the right-click menu and the desktop menu bar all run commands. _Avoid_: action

**Command menu**: The search box that runs any **Command** by name (`⌘K`). _Avoid_: palette, command palette, quick pick

**Notice**: A short message at the bottom of the window, sometimes with a button such as **Undo**. _Avoid_: toast, snackbar, popup

**Dialog**: A window over the app that holds focus until you close it with a button or `Esc`. _Avoid_: modal, sheet (a **Sheet** is part of the workbook)

**Keyboard Shortcuts**: The **Dialog** that lists every shortcut (`⌘/`). **Look Up a Key…** in it says which **Command** a key runs, without running it.

**Current item**: The one item in a group (the **Sheet tabs**, the **Ribbon** tabs, a menu, a list) that the arrow keys move and that has focus when the group does.

### Where it runs

**Browser build**: Cellestra in a browser tab, opening files from your computer.

**NAS server**: Cellestra's own small server, run with Docker, that serves the app and the files in one folder to any browser on your network. _Avoid_: backend

**Desktop app**: Cellestra as an app for macOS, Windows or Linux, opening files on your disk.

**CLI**: The `cellestra` command, which reads and edits **Workbooks** without the window.

## Relationships

- A **Workbook** holds one or more **Sheets**, any number of **Names** and **Tables**, and is in one save state at a time: **Saved**, **Edited** or **Not saved**.
- A **Sheet** is a grid of **Cells** or a **Chart sheet**, never both.
- A **Table** belongs to one **Sheet**; a **Name** belongs to the **Workbook** or to one **Sheet**.
- A **Chart** is in **Table mode** or **Range mode** and plots one or more **Series** against its **Categories**.
- A **List** is one kind of **Validation rule**; its item colours are stored as **Conditional formats**.
- Every button, menu item and shortcut runs a **Command**.

## Flagged ambiguities

- "Named range" was used for every **Name**, though a name can hold a value or formula, not only a range. Resolved: the concept is a **Name**; **Named Ranges…** stays the dialog's title because LibreOffice and Excel users look for it.
- "Edited" was shown for two states. Resolved: **Edited** means the workbook has a file and has changed; **Not saved** means it has changes and no file yet. The status bar and the title bar both follow this.
- "Toast" and "notice" were both used for the message at the bottom. Resolved: it is a **Notice**; "toast" remains only in old code names.
- "Palette" is the code's name for the **Command menu** (`paletteOpen`) and also the fixed set of **List** item colours. Resolved: the menu is the **Command menu**; "palette" means only the colour set.
- "Tab" can mean a **Sheet tab**, a **Ribbon** tab or the `Tab` key. Always say which.
