# Getting started

You can run Cellestra in your browser, on a NAS with Docker, or as a desktop app. There are no prebuilt downloads yet, so each one starts from this repository.

## What you need

- For the browser and the desktop app: Node 22.18 or later, pnpm (through `corepack enable`) and [just](https://github.com/casey/just).
- For the NAS: Docker.
- To open a file in place and save it back in the browser: a Chromium browser. Other browsers open a file you pick and save a downloaded copy.

Install the dependencies once:

```bash
just setup   # corepack enable, then pnpm install
```

## Run it on the desktop

```bash
just run-desktop      # run it in dev mode
just package-desktop  # build a .dmg / AppImage / installer in apps/desktop/release
just install-desktop  # macOS: package and install into /Applications
```

The desktop app has native menus, file associations for `.ods` and `.csv`, and normal Open and Save dialogs. `just install-desktop` moves an existing copy to the Trash first, and refuses while the app is running.

## Run it in the browser

```bash
pnpm --filter @cellestra/web dev   # the browser-only app on http://localhost:5173
```

Open an `.ods`, `.csv` or `.tsv` file from your computer, edit it, and save it back. Everything happens in the page; nothing is uploaded.

## Run it on a NAS (Docker)

```bash
docker build -t cellestra .
docker run -d -p 8787:8787 -v /volume1/Documents:/data cellestra
```

Then open `http://<nas>:8787`. A small server shows the folders you choose, so you can open and save spreadsheets on the NAS from any browser on your network. To expose several folders, add a password, or run as your NAS user, see [compose.example.yaml](../compose.example.yaml).

| Variable           | Default           | Meaning                                                           |
| ------------------ | ----------------- | ----------------------------------------------------------------- |
| `CELLESTRA_ROOTS`  | `Documents=/data` | The folders to show, as `Name=/path;Other=/path2`                 |
| `CELLESTRA_TOKEN`  | unset             | If set, the browser asks for this token before it shows any files |
| `CELLESTRA_HOSTS`  | unset             | Without a token, the extra host names to answer (comma-separated) |
| `PORT`             | `8787`            | The port to listen on                                             |
| `CELLESTRA_MAX_MB` | `64`              | The largest file you can save, in MB                              |

Without `CELLESTRA_TOKEN`, the server only answers requests addressed to an IP address or `localhost`. If you open it by name (for example `http://nas.lan:8787`), list that name in `CELLESTRA_HOSTS=nas.lan`. Other names get `421 Misdirected Request`, which stops a hostile web page from reaching your files through DNS rebinding.

If LibreOffice has a file open (it leaves a `.~lock` file next to it), saving that file is refused. If someone else saved the file after you opened it, Cellestra warns you before overwriting their changes.

## Where your data lives

Your spreadsheets stay where you opened them: on your computer, or in the NAS folders the server shows. Cellestra keeps the recent files list and a crash-recovery snapshot in the page's IndexedDB storage.
