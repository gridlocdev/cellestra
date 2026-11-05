# Architecture

Cellestra is one TypeScript code base with three hosts. The browser, the NAS server's client and the Electron app all run the same engine and UI. The only thing that differs between them is where files come from.

```
            ┌──────────────── packages/ui (Preact + signals) ────────────────┐
            │  App · Grid (canvas) · FormulaBar · SheetTabs · ⌘K · dialogs   │
            │  WorkbookSession (undo, selection, dirty) · CommandCatalog     │
            │  FileStore ─┬─ BrowserStore  (File System Access / upload)     │
            │             ├─ NasStore      (fetch /api, ETag)                │
            │             └─ NativeStore   (window.cellestraNative IPC)      │
            └──────────────────────────────┬─────────────────────────────────┘
                                           │ imports
            ┌──────────────── packages/core (pure TS) ───────────────────────┐
            │  values · model · formulas · dependencies (recalc) · editing   │
            │  data (sort/filter/tables/validation) · styling · xml          │
            │  serialization (ods via fflate + saxes, csv)                   │
            └──────────────────────────────┬─────────────────────────────────┘
                                           │
                               packages/charts (model + ODF chart XML)

 apps/web      Vite entry; static build with relative URLs (base './')
 apps/server   node:http: static client + /api/{info,roots,list,file}
 apps/desktop  Electron main + sandboxed preload; loads apps/web build via app://
```

## Core (`packages/core`)

- **Values.** `CellValue` is a discriminated union (`empty | number | text | boolean | error`). Spreadsheet errors are values; evaluation never throws for them.
- **Model.** `Workbook` holds `Worksheet[]`, a `StylePool` and `DocumentSettings`.
  - A worksheet keeps its cells in one `Map<number, Cell>` keyed by `row * 16384 + col`, so there is no 2-D array and no object for each empty cell.
  - Rows and columns that differ from the defaults are kept in their own sparse maps.
  - Styles are interned: a cell stores a style id.
- **Formulas.** The pipeline is lexer → parser (AST) → evaluator. `FunctionRegistry` holds every function together with its metadata, which also drives autocomplete and help. `OpenFormulaTranslator` converts between the internal syntax (`SUM(A1:A3)`) and ODF's (`of:=SUM([.A1:.A3])`).
- **Dependencies.** `RecalculationEngine` covers full recalc, recalc after an edit, volatile functions, manual mode, and cycle detection, which shows `#REF!`.
- **Editing.** Every model change is an edit command with apply and revert. `EditHistory` keeps up to 1,000 of them. Structural edits rewrite formula references through `ReferenceAdjuster`.
- **Serialization.**
  - `.ods` is read and written through a small namespace-aware XML tree, so elements Cellestra doesn't model survive a round trip verbatim.
  - The `mimetype` entry is always first in the zip and stored uncompressed.
  - CSV is RFC 4180.
- **Locale.** All number, date and collation behaviour goes through `Intl`. A document's locale is a BCP-47 tag; if it has none, the host's locale is used (set through `setHostLocale` at startup).

## UI (`packages/ui`)

- **Grid.** The grid is a single `<canvas>` over a native scroller.
  - `Axis` stores only the rows and columns with non-default sizes, as prefix sums, so offsets and hit-tests are O(log n) even with a million rows.
  - `Viewport` is the geometry for one frame: frozen panes, zoom and visible bands. Rendering, hit-testing and the inline editor all use it.
  - The renderer reads the `Worksheet` directly.
  - The editor is a `<textarea>` placed over the active cell, so IME works.
- **State.** Signals hold the chrome state: selection, zoom, theme and a document version counter. The grid redraws whenever the version changes.
- **Commands.** `CommandCatalog` is the one list of commands. The toolbar, ⌘K, the shortcut dispatcher and the Electron menu bar (sent over IPC) all read it. Shortcuts are written once as `mod+shift+k` and formatted for each platform.
- **Files.** `FileStore` has three implementations, and the UI only ever sees the interface.
  - Conflict detection uses an ETag (NAS) or the file's mtime (native).
  - IndexedDB keeps the recent files (including browser file handles) and a crash-recovery snapshot.

## Server (`apps/server`)

- **Roots.** `CELLESTRA_ROOTS=Name=/dir;...` names the folders the server exposes.
- **Paths.** Every path is jailed to its root. This includes escapes through `..`, absolute paths, NUL bytes and symlinks.
- **Writes.**
  - Each save is atomic: temp file, then fsync, then rename.
  - `If-Match` gives optimistic concurrency (412 if the file changed). `If-None-Match: *` stops a new file from overwriting an existing one.
  - A LibreOffice `.~lock` file next to the target blocks the save (423).
  - Only `.ods`, `.csv` and `.tsv` files can be written.
- **Auth.** Optional: set `CELLESTRA_TOKEN` to require a bearer token, or put the server behind the NAS's reverse proxy.

## Desktop (`apps/desktop`)

- **Isolation.** The renderer runs with `sandbox`, `contextIsolation` and no Node.
- **Bridge.** The preload exposes exactly one object, `window.cellestraNative`. Its contract is in `bridge.ts` and is mirrored in `packages/ui/src/platform/native.ts`.
- **Loading.** The UI loads from `app://cellestra/`, a privileged standard scheme backed by `out/web`. Navigation away from it is blocked, and links open in the system browser.
- **Main process.** It handles open and save dialogs, atomic saves with an mtime conflict check, the native menu (built from the renderer's catalog), document state (title, the edited dot, represented file), the close guard, the external-change watcher, recent documents, and file associations through `open-file` and argv.
