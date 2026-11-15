# Cellestra — agent guide

A TypeScript spreadsheet editor. The `.ods` engine runs entirely in the browser; it is shared by the web client, the NAS server's client, the Electron app and the `cellestra` CLI.

## Gates

- `just check` (= `pnpm check`) must pass before you commit: Prettier check, ESLint, `tsc -b`, Vitest, and the build.
- When you change behaviour, add or update a test in the package's `test/` folder.

## Layout

| Path              | What it is                                                                                                                                                                                                         |
| ----------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `packages/core`   | The pure engine: values, model, formulas, recalc, editing and undo commands, data tools, styling, conditional format rules, print HTML, and `.ods`/CSV I/O. Must not touch the DOM or Node (ESLint enforces this). |
| `packages/charts` | The chart model and the ODF chart XML.                                                                                                                                                                             |
| `packages/ui`     | Preact and signals building blocks: the canvas grid, formula bar, sheet tabs, the `CommandCatalog` type, ⌘K, dialogs, `WorkbookSession`, and the `FileStore` adapters (browser, NAS, native).                      |
| `apps/web`        | The app shell and the Vite entry point: `App`, the ribbon, the controller, the list of commands, `DocumentStores`, and the static build for the server and the desktop app.                                        |
| `apps/cli`        | The `cellestra` command line (Node, run from source). It depends on `packages/core` only.                                                                                                                          |
| `apps/server`     | The NAS server (`node:http`): serves the client and `/api` for the configured roots. No framework.                                                                                                                 |
| `apps/desktop`    | The Electron shell: main process, sandboxed preload bridge, native menus, file associations.                                                                                                                       |

## Golden rules

1. **The core stays pure.** It has no UI, DOM or Node APIs, and it runs the same in a browser, a worker or Node tests.
2. **No object per cell in the UI.** The grid draws straight from `Worksheet` onto a canvas, and it only draws the cells that are visible.
3. **One command catalog.** The `CommandCatalog` type is in `packages/ui/src/commands`; the commands themselves are in `apps/web/src/app/commands.ts`. The ribbon, the ⌘K menu, keyboard shortcuts and the Electron menu bar all read from it.
4. **One function registry.** Evaluation, autocomplete and help all read from it.
5. **`.ods` correctness.** The `mimetype` entry is first and stored uncompressed. Formulas go through `OpenFormulaTranslator`. Unknown content is kept verbatim.
6. **Formula semantics follow LibreOffice Calc** (PRD §8 and §10). For example, `-2^2 = 4` (the sign binds tighter than `^`), ROUND rounds half away from zero, and serial dates start at 1899-12-30. Dates never use the local time zone.
7. **The UI only talks to files through `FileStore`.** It never knows whether bytes came from the browser, the NAS or Electron. `DocumentStores` (`apps/web/src/app/documentStores.ts`) picks the store for a document and the `FilePicker` for Open and Save As.
8. **Security.** The Electron renderer runs sandboxed with context isolation, and the only native access it has is the `cellestraNative` bridge. The server confines every path to its root, including symlink escapes, writes atomically, honours If-Match, and respects LibreOffice lock files.

## TypeScript conventions

- Strict mode with `noUncheckedIndexedAccess` and `erasableSyntaxOnly`. That means no enums, no parameter properties and no namespaces: use union types and `const` objects instead.
- Relative imports end in `.ts`. Use `import type` for types only.
- Each package's `src/index.ts` names its exports one by one (no `export *`). Other packages and apps import only through `@cellestra/<package>`; ESLint rejects imports into another package's source files.
- Edits that the app and the CLI share are core edit commands, so both make the same change.
- Spreadsheet errors are values (`CellValue` with kind `error`), never thrown exceptions.
- Pin dependencies to exact versions and keep the list short.
