# Cellestra — agent guide

A TypeScript spreadsheet editor. The `.ods` engine runs entirely in the browser; it is shared by the web client, the NAS server's client, and the Electron app.

## Gates

- `just check` (= `pnpm check`) must pass before you commit: Prettier check, ESLint, `tsc -b`, Vitest, and the build.
- When you change behaviour, add or update a test in the package's `test/` folder.

## Layout

| Path              | What it is                                                                                                                                                                   |
| ----------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `packages/core`   | The pure engine: values, model, formulas, recalc, editing and undo commands, data tools, styling, and `.ods`/CSV I/O. Must not touch the DOM or Node (ESLint enforces this). |
| `packages/charts` | The chart model and the ODF chart XML.                                                                                                                                       |
| `packages/ui`     | Preact and signals: the canvas grid, formula bar, sheet tabs, command catalog, ⌘K, dialogs, and the `FileStore` adapters (browser, NAS, native).                             |
| `apps/web`        | The Vite entry point that builds the static client for the server and the desktop app.                                                                                       |
| `apps/server`     | The NAS server (`node:http`): serves the client and `/api` for the configured roots. No framework.                                                                           |
| `apps/desktop`    | The Electron shell: main process, sandboxed preload bridge, native menus, file associations.                                                                                 |

## Golden rules

1. **The core stays pure.** It has no UI, DOM or Node APIs, and it runs the same in a browser, a worker or Node tests.
2. **No object per cell in the UI.** The grid draws straight from `Worksheet` onto a canvas, and it only draws the cells that are visible.
3. **One command catalog** (`packages/ui/src/commands`). The toolbar, the ⌘K menu, keyboard shortcuts and the Electron menu bar all read from it.
4. **One function registry.** Evaluation, autocomplete and help all read from it.
5. **`.ods` correctness.** The `mimetype` entry is first and stored uncompressed. Formulas go through `OpenFormulaTranslator`. Unknown content is kept verbatim.
6. **Formula semantics follow LibreOffice Calc** (PRD §8 and §10). For example, `-2^2 = -4`, ROUND rounds half away from zero, and serial dates start at 1899-12-30. Dates never use the local time zone.
7. **The UI only talks to files through `FileStore`.** It never knows whether bytes came from the browser, the NAS or Electron.
8. **Security.** The Electron renderer runs sandboxed with context isolation, and the only native access it has is the `cellestraNative` bridge. The server confines every path to its root, including symlink escapes, writes atomically, honours If-Match, and respects LibreOffice lock files.

## TypeScript conventions

- Strict mode with `noUncheckedIndexedAccess` and `erasableSyntaxOnly`. That means no enums, no parameter properties and no namespaces: use union types and `const` objects instead.
- Relative imports end in `.ts`. Use `import type` for types only.
- Spreadsheet errors are values (`CellValue` with kind `error`), never thrown exceptions.
- Pin dependencies to exact versions and keep the list short.
