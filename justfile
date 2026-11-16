# Cellestra — run `just` to list recipes.
set shell := ["bash", "-cu"]

app_name := "Cellestra"
nas_data := justfile_directory() / ".nas-data"
release := "apps/desktop/release"

default:
    @just --list --unsorted

# ── Run (dev) ─────────────────────────────────────────────────────────────

# Web client dev server (http://localhost:5173) plus the NAS file server it talks to. Ctrl-C stops both.
[group('run')]
run-web:
    mkdir -p "{{ nas_data }}"
    VITE_CELLESTRA_NAS=1 pnpm --parallel --filter @cellestra/server --filter @cellestra/web run dev

# NAS file server only (http://localhost:8787), serving ./.nas-data; restarts when its source changes.
[group('run')]
run-server:
    mkdir -p "{{ nas_data }}"
    pnpm --filter @cellestra/server dev

# Desktop app in dev mode (builds the web client first).
[group('run')]
run-desktop: build-web
    pnpm --filter @cellestra/desktop start

# The built web client served by the NAS server, as it runs on a NAS (http://localhost:8787).
[group('run')]
preview-web: build-web-nas
    mkdir -p "{{ nas_data }}"
    CELLESTRA_HOST=127.0.0.1 CELLESTRA_ROOTS="Sample={{ nas_data }}" CELLESTRA_STATIC="{{ justfile_directory() }}/apps/web/dist" node apps/server/src/main.ts

# Run the cellestra CLI from source (paths resolve from where you run it; quoting is kept).
[group('run')]
[no-cd]
[positional-arguments]
cli *args:
    node "{{ justfile_directory() }}/apps/cli/src/main.ts" "$@"

# ── Build and package ─────────────────────────────────────────────────────

# Build everything: the web client and the desktop app.
[group('build')]
build:
    pnpm build

# Build the web client into apps/web/dist.
[group('build')]
build-web:
    pnpm --filter @cellestra/web build

# Build the web client for the NAS server (shows NAS open/save; `build-web` leaves it out).
[group('build')]
build-web-nas:
    VITE_CELLESTRA_NAS=1 pnpm --filter @cellestra/web build

# Regenerate the app icon files in apps/desktop/build from assets/icon.svg (needs Inkscape; .icns needs macOS).
[group('build')]
icons:
    ./scripts/make-icons.sh

# Build the desktop app's main process and copy the web client into apps/desktop/out.
[group('build')]
build-desktop: build-web
    pnpm --filter @cellestra/desktop build

# Package the desktop app for this OS into apps/desktop/release (.dmg and .zip on macOS).
[group('build')]
package-desktop: build-web
    pnpm --filter @cellestra/desktop package

# Package an unpacked app only (faster; for local testing).
[group('build')]
package-desktop-dir: build-web
    pnpm --filter @cellestra/desktop package --dir

# Package and install into /Applications (macOS). The old copy goes to the Trash.
[group('build')]
[macos]
install-desktop: package-desktop-dir
    #!/usr/bin/env bash
    set -euo pipefail
    target="/Applications/{{ app_name }}.app"
    built=$(ls -d "{{ release }}"/mac*/"{{ app_name }}.app" | head -1)
    # Anchored, so a shell command that merely mentions the path does not count.
    if pgrep -f "^$target/Contents/MacOS/" >/dev/null; then
      echo "{{ app_name }} is running — quit it first, then run this again." >&2
      exit 1
    fi
    trashed=""
    if [ -d "$target" ]; then
      trashed="$HOME/.Trash/{{ app_name }} $(date '+%Y-%m-%d %H.%M.%S').app"
      mv "$target" "$trashed"
      echo "Moved the old app to the Trash."
    fi
    ditto "$built" "$target"
    lsregister=/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister
    # Only the installed copy should answer "Open With" for .ods and .csv.
    "$lsregister" -u "$PWD/$built" 2>/dev/null || true
    if [ -n "$trashed" ]; then "$lsregister" -u "$trashed" 2>/dev/null || true; fi
    "$lsregister" -f "$target"
    echo "Installed $target"

# Build the NAS Docker image.
[group('build')]
build-docker tag="cellestra:latest":
    docker build -t {{ tag }} .

# Run the NAS Docker image against ./.nas-data (http://localhost:8787).
[group('build')]
run-docker tag="cellestra:latest":
    mkdir -p "{{ nas_data }}"
    docker run --rm -p 127.0.0.1:8787:8787 -e CELLESTRA_NO_AUTH=1 -v "{{ nas_data }}:/data" {{ tag }}

# ── Quality ───────────────────────────────────────────────────────────────

# Everything CI checks: format, lint, typecheck, tests, build.
[group('quality')]
check:
    pnpm check

# Run the tests once (pass a file or -t "name" to narrow them).
[group('quality')]
test *args:
    pnpm vitest run {{ args }}

# Run the random-operation tests with more seeds; a failure prints FUZZ_SEED=<n> to replay it.
fuzz runs="500":
    FUZZ_RUNS={{ runs }} pnpm vitest run -t random --testTimeout=0

# Run the tests in watch mode.
[group('quality')]
test-watch *args:
    pnpm vitest {{ args }}

# Lint with ESLint.
[group('quality')]
lint:
    pnpm lint

# Typecheck every package.
[group('quality')]
typecheck:
    pnpm typecheck

# Format every file.
[group('quality')]
fmt:
    pnpm format

# Check formatting without changing anything.
[group('quality')]
fmt-check:
    pnpm format:check

# Architecture rules only: layers, entry points, platform APIs, timers, swallowed errors, file size. Part of `check`.
[group('quality')]
check-arch:
    pnpm vitest run test/architecture.test.ts

# Record today's architecture violations and exceptions (once, when adopting the rules; refuses after that).
[group('quality')]
arch-baseline:
    ARCH_BASELINE=write pnpm vitest run test/architecture.test.ts

# ── Housekeeping ──────────────────────────────────────────────────────────

# Install dependencies (pnpm via corepack).
[group('housekeeping')]
setup:
    corepack enable
    pnpm install

# Remove build output.
[group('housekeeping')]
clean:
    rm -rf apps/web/dist apps/desktop/out "{{ release }}" apps/desktop/node_modules/.cache/cellestra-dev .types-test
    find . -name '*.tsbuildinfo' -not -path '*/node_modules/*' -delete

# Remove build output and every node_modules (run `just setup` after).
[group('housekeeping')]
clean-all: clean
    find . -name node_modules -type d -prune -exec rm -rf {} +

# Earlier names, kept so nothing that uses them breaks.
alias web := run-web
alias server := run-server
alias desktop := run-desktop
alias package := package-desktop
alias docker := build-docker
