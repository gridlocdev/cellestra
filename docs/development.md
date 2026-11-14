# Development

Cellestra is TypeScript in a pnpm workspace: a pure `.ods` and formula engine, Preact and `@preact/signals` for the UI, Vite for the web build, Electron for the desktop app, and a dependency-free Node server and CLI that run straight from source with Node's type stripping. The rules in `AGENTS.md` at the repository root apply to everyone working on it, and the code layout and its layer rules are in [architecture.md](architecture.md). The product spec is [plans/prd.md](plans/prd.md).

## Getting set up

You need Node 22.18 or later, pnpm (through `corepack enable`) and [just](https://github.com/casey/just). Run `just` on its own to list every recipe.

```bash
just setup         # corepack enable, then pnpm install
just run-web       # web client on http://localhost:5173 plus the NAS server on :8787
just run-server    # the NAS server only, serving ./.nas-data, restarting on changes
just run-desktop   # the desktop app in dev mode (builds the web client first)
just preview-web   # the built NAS client served by the NAS server, as on a NAS (:8787)
just cli --help    # the cellestra CLI from source (any arguments)
```

- `just run-web`, `just run-server` and `just preview-web` serve `.nas-data/` in the repository as the root `Sample`. Put spreadsheets in it to try the NAS flow.
- The NAS open and save commands exist only in a NAS build (`VITE_CELLESTRA_NAS=1`), which `just run-web`, `just build-web-nas`, `just preview-web` and the Docker image make. `pnpm --filter @cellestra/web dev` is the browser-only app.

## Checking your work

```bash
just check        # everything CI checks: Prettier, ESLint, tsc -b, Vitest, then the build
just test         # the tests once; narrow with a file or -t "name"
just test-watch   # the tests in watch mode
just lint         # ESLint
just typecheck    # tsc -b over every package
just fmt          # format every file
just fmt-check    # check the formatting
```

`just check` must pass before you commit. When you change behaviour, add or update a test in that package's `test/` folder.

To measure how long reading a large `.ods` takes and how much memory it peaks at, generate a file once and read it in a fresh process each time:

```bash
node scripts/bench-ods-read.mjs generate /tmp/big.ods 100000 10   # 1M cells
node scripts/bench-ods-read.mjs read /tmp/big.ods
```

Two fixtures in `packages/core/test/serialization/fixtures` are files LibreOffice saved itself, and a test checks that saving them without edits keeps every element and attribute they have. The scripts that build them need LibreOffice in `/Applications`:

```bash
python3 scripts/make-libreoffice-sample.py    # values, formulas, formats, merges, validations
python3 scripts/make-libreoffice-outline.py   # groups, print titles, page breaks, macros and events
```

## The cellestra CLI

`cellestra` reads and edits `.ods` files with the same engine as the app, for scripts and CI. It runs from source, from the folder you are in:

```bash
just cli --help                         # every command and option
just cli convert Budget.ods Budget.csv  # convert between .ods and CSV, both directions
just cli cat Budget.ods                 # print a sheet or range
just cli eval "=SUM(1,2,3)"             # evaluate a formula with no file
just cli validate Budget.ods            # exit 1 on formula, #REF! or circular errors
just cli diff Old.ods New.ods           # compare two workbooks; exit 1 if they differ
just cli functions                      # list the supported functions
```

Exit codes: `0` success, `1` a problem was found (validate, diff, find, `lint --strict`) or the input was bad (a missing file, a bad address), `2` a usage error (an unknown command or option, a missing argument). `just cli man` lists them all.

## The NAS server

```bash
just build-docker   # build the image (cellestra:latest)
just run-docker     # run it against ./.nas-data on http://localhost:8787
```

The server is configured only through environment variables, listed in [Getting started](getting-started.md#run-it-on-a-nas-docker) and in [architecture.md](architecture.md#server-appsserver). `compose.example.yaml` is the starting point for a NAS.

## Building and packaging the desktop app

```bash
just build                # the web client and the desktop app
just build-web            # the web client only, into apps/web/dist
just build-web-nas        # the web client with NAS open and save
just build-desktop        # the desktop app only, into apps/desktop/out
just package-desktop      # a .dmg and .zip (macOS), AppImage (Linux) or installer (Windows) in apps/desktop/release
just package-desktop-dir  # an unpacked app only, faster for local testing
just install-desktop      # macOS: package and install into /Applications
```

- The package is signed ad hoc so it runs on Apple silicon. Swap the identity in `apps/desktop/electron-builder.yml` for a Developer ID to distribute it.
- `just install-desktop` moves an existing copy to the Trash first, and refuses while the app is running. It registers the installed copy with Launch Services, so it answers **Open With** for `.ods` and `.csv`.

### The app icon

The icon is drawn by hand in `assets/icon.svg`. After editing it, render the platform files and commit them, so builds don't need Inkscape:

```bash
just icons  # icon.icns (macOS), icon.ico (Windows) and icon.png in apps/desktop/build
```

It needs Inkscape and, for the `.icns`, macOS.

## Releasing

Pushing a `vX.Y.Z` tag runs `.github/workflows/release.yml`. It packages the desktop app on macOS (Apple silicon and Intel), Windows and Linux, and publishes them on a GitHub release. The NAS image isn't published: people build it from the repository ([Getting started](getting-started.md#run-it-on-a-nas-docker)).

1. Set the same version in every `package.json` under `apps/` and `packages/`. The workflow stops if the tag doesn't match `apps/desktop/package.json`.
2. Add a `## X.Y.Z` section to `CHANGELOG.md`. It becomes the release notes.
3. Run `just check`, commit, then tag and push:

   ```bash
   git tag -a vX.Y.Z -m "Cellestra X.Y.Z"
   git push origin main vX.Y.Z
   ```

The apps aren't notarized or code-signed, so [Getting started](getting-started.md#download) tells people how to open them the first time.

## Cleaning up

```bash
just clean        # remove build output
just clean-all    # also remove every node_modules (run just setup after)
```
