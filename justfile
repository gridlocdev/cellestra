# Cellestra — run `just` to list recipes.
set shell := ["bash", "-cu"]

default:
    @just --list

# Install dependencies (pnpm via corepack).
setup:
    corepack enable
    pnpm install

# Everything CI checks: format, lint, typecheck, tests, build.
check:
    pnpm check

test *args:
    pnpm vitest run {{args}}

# Web client dev server (http://localhost:5173), with the NAS API proxied to `just server`.
web:
    pnpm --filter @cellestra/web dev

# NAS file server against ./.nas-data (serves apps/web/dist when built).
server:
    mkdir -p .nas-data
    pnpm --filter @cellestra/server dev

# Desktop app (builds the web client first).
desktop:
    pnpm --filter @cellestra/web build
    pnpm --filter @cellestra/desktop start

# Package the desktop app for this OS into apps/desktop/release.
package:
    pnpm --filter @cellestra/web build
    pnpm --filter @cellestra/desktop package

# Build the NAS Docker image.
docker tag="cellestra:latest":
    docker build -t {{tag}} .

fmt:
    pnpm format
