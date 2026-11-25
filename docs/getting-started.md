# Getting started

You can download the desktop app, use Cellestra in your browser at [gridloc.dev/apps/cellestra](https://gridloc.dev/apps/cellestra/), run it on a NAS with Docker, or run any of them from this repository.

## Download

Get the desktop app from the [Releases page](https://github.com/gridlocdev/cellestra/releases/latest):

| System                | File                        |
| --------------------- | --------------------------- |
| macOS (Apple silicon) | `Cellestra-…-mac-arm64.dmg` |
| macOS (Intel)         | `Cellestra-…-mac-x64.dmg`   |
| Windows               | `Cellestra-…-win-x64.exe`   |
| Linux                 | `Cellestra-….AppImage`      |

The apps aren't signed by Apple or Microsoft yet, so the first launch needs one extra step:

- **macOS** says Cellestra "can't be opened" or "is damaged". Open **System Settings > Privacy & Security**, scroll down, and click **Open Anyway** next to Cellestra. Or run this once after moving the app to Applications:

  ```bash
  xattr -dr com.apple.quarantine /Applications/Cellestra.app
  ```

- **Windows** shows "Windows protected your PC". Click **More info**, then **Run anyway**.
- **Linux**: make the AppImage executable (`chmod +x Cellestra-*.AppImage`), then run it.

Cellestra doesn't check for updates yet. Watch the Releases page for new versions.

## What you need to run it from source

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

The desktop app has native menus, file associations for `.ods` and `.csv`, and normal Open and Save dialogs. `just install-desktop` moves an existing copy to the Trash first. If Cellestra is running, it asks to quit it and opens it again after the install.

## Run it in the browser

```bash
pnpm --filter @cellestra/web dev   # the browser-only app on http://localhost:5173
```

Open an `.ods`, `.csv` or `.tsv` file from your computer, edit it, and save it back. Everything happens in the page; nothing is uploaded.

## Run it on a NAS (Docker)

```bash
docker build -t cellestra .
docker run -d -p 8787:8787 -e CELLESTRA_TOKEN=<a-long-secret> -v /volume1/Documents:/data cellestra
```

Then open `http://<nas>:8787`. A small server shows the folders you choose, so you can open and save spreadsheets on the NAS from any browser on your network. The browser asks for the token once. To expose several folders, use a reverse proxy, or run as your NAS user, see [compose.example.yaml](../compose.example.yaml).

| Variable | Default | Meaning |
| --- | --- | --- |
| `CELLESTRA_ROOTS` | `Documents=/data` | The folders to show, as `Name=/path;Other=/path2` |
| `CELLESTRA_TOKEN` | unset | The browser asks for this token before it shows any files. Required unless `CELLESTRA_HOST` is a loopback address |
| `CELLESTRA_NO_AUTH` | unset | `1` runs without a token on a network address. Only use it when a reverse proxy in front requires a login |
| `CELLESTRA_HOST` | `0.0.0.0` | The address to listen on |
| `CELLESTRA_HOSTS` | unset | Without a token, the extra host names to answer (comma-separated) |
| `CELLESTRA_PORT` | `8787` | The port to listen on. Plain `PORT` also works when this is unset |
| `CELLESTRA_MAX_MB` | `64` | The largest file you can save, in MB |

The server refuses to start on a network address (such as the default `0.0.0.0`) without `CELLESTRA_TOKEN`, because anyone who could reach it could read and overwrite your files. On `CELLESTRA_HOST=127.0.0.1` (or `localhost`, `::1`) it runs without one. If your NAS's reverse proxy already asks for a login and forwards to the container, set `CELLESTRA_NO_AUTH=1` instead.

Without `CELLESTRA_TOKEN`, the server only answers requests addressed to an IP address or `localhost`. If you open it by name (for example `http://nas.lan:8787`), list that name in `CELLESTRA_HOSTS=nas.lan`. Other names get `421 Misdirected Request`, which stops a hostile web page from reaching your files through DNS rebinding.

If LibreOffice has a file open (it leaves a `.~lock` file next to it), saving that file is refused. If someone else saved the file after you opened it, Cellestra warns you before overwriting their changes.

## Where your data lives

Your spreadsheets stay where you opened them: on your computer, or in the NAS folders the server shows. Cellestra keeps the recent files list and a crash-recovery snapshot in the page's IndexedDB storage.
