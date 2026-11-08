# Cellestra

A spreadsheet editor for OpenDocument spreadsheets (`.ods`). It runs three ways, all from the same code:

- **In the browser.** Open an `.ods`, `.csv` or `.tsv` file from your computer, edit it, and save it back. Everything happens in the page; nothing is uploaded.
- **On your NAS.** A small server shows the folders you choose, so you can open and save spreadsheets on the NAS from any browser on your network.
- **On the desktop.** An Electron app with native menus, file associations and normal Open/Save dialogs.

## What it does

- Formulas with 47 functions, recalculation and circular-reference detection. The formula syntax matches LibreOffice Calc.
- Reads and writes `.ods` and keeps anything it doesn't understand intact, so LibreOffice files survive a round trip. Also imports and exports CSV and TSV.
- Formatting: number formats, fonts, fills, borders, alignment, wrapping and merges.
- Data tools: sort, filter, tables, validation, remove duplicates and Clean Up.
- Charts.
- Every command is reachable from the keyboard: ⌘K opens the command menu, ⌘⇧P switches sheets.

## Run it on a NAS (Docker)

```bash
docker build -t cellestra .
docker run -d -p 8787:8787 -v /volume1/Documents:/data cellestra
```

Then open `http://<nas>:8787`. To expose several folders, add a password, or run as your NAS user, see [compose.example.yaml](compose.example.yaml).

| Variable           | Default           | Meaning                                                           |
| ------------------ | ----------------- | ----------------------------------------------------------------- |
| `CELLESTRA_ROOTS`  | `Documents=/data` | The folders to show, as `Name=/path;Other=/path2`                 |
| `CELLESTRA_TOKEN`  | unset             | If set, the browser asks for this token before it shows any files |
| `CELLESTRA_HOSTS`  | unset             | Without a token, the extra host names to answer (comma-separated) |
| `PORT`             | `8787`            | The port to listen on                                             |
| `CELLESTRA_MAX_MB` | `64`              | The largest file you can save, in MB                              |

Without `CELLESTRA_TOKEN`, the server only answers requests addressed to an IP address or `localhost`. If you open it by name (for example `http://nas.lan:8787`), list that name in `CELLESTRA_HOSTS=nas.lan`. Other names get `421 Misdirected Request`, which stops a hostile web page from reaching your files through DNS rebinding.

If LibreOffice has a file open (it leaves a `.~lock` file next to it), saving that file is refused. If someone else saved the file after you opened it, Cellestra warns you before overwriting their changes.

## Desktop

```bash
just run-desktop      # run it in dev mode
just package-desktop  # build a .dmg / AppImage / installer in apps/desktop/release
just install-desktop  # macOS: package and install into /Applications
```

## For developers

You need Node 22.18 or later and pnpm (installed via `corepack enable`). Run `just setup`, then `just run-web` for the dev server and `just check` before you commit. Run `just` on its own to list every recipe. [AGENTS.md](AGENTS.md) has the rules and [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) explains how the code is laid out. The product spec is in [docs/PRD.md](docs/PRD.md).

## License

MIT, see [LICENSE](LICENSE).
