// Desktop smoke test: the built Electron app starts headless (CELLESTRA_HEADLESS=1: offscreen, no window, Dock icon
// or system dialog) on a throwaway profile, so it can't reopen recent files, and opens a fixture workbook passed on
// its command line, the way Finder hands it one. `just e2e-desktop` builds it first.
import { cpSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { _electron as electron } from '@playwright/test';
import { expect, test } from './fixtures.ts';

const DESKTOP = fileURLToPath(new URL('../apps/desktop', import.meta.url));
const FIXTURE = fileURLToPath(new URL('./fixtures/basic', import.meta.url));

test.describe('Three ways to run it', () => {
  test('the desktop app starts on a throwaway profile and opens a workbook it is given', async () => {
    const scratch = mkdtempSync(path.join(tmpdir(), 'cellestra-e2e-'));
    const profile = path.join(scratch, 'profile');
    const files = path.join(scratch, 'files');
    cpSync(FIXTURE, files, { recursive: true });
    const app = await electron.launch({
      cwd: DESKTOP,
      args: ['.', `--user-data-dir=${profile}`, path.join(files, 'budget.ods')],
      env: { ...process.env, CELLESTRA_HEADLESS: '1' },
    });
    try {
      const window = await app.firstWindow();
      await window.getByRole('main').waitFor();
      await expect(window).toHaveTitle('budget.ods — Cellestra');
      await expect(window.getByRole('textbox', { name: 'Cell contents' })).toHaveValue('Month');
      expect(
        await app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().some((w) => w.isVisible())),
      ).toBe(false);
    } finally {
      await app.close();
      rmSync(scratch, { recursive: true, force: true });
    }
  });
});
