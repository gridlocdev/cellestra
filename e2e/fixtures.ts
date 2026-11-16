// Shared scenario setup. `seed('basic')` replaces .e2e-data with a copy of e2e/fixtures/basic,
// so each test starts from known files and nothing one test writes leaks into the next.
// Every page runs on a clock that starts at NOW and keeps going, so dates read the same on every run;
// `page.clock.runFor(ms)` moves it on, which checks a timer (auto-save, the NAS change check) without waiting for it.
//
// The grid draws on a canvas, so a cell is read the way a keyboard user reads it: go to it with the Name Box,
// then read the formula bar.
import { expect, test as base, type Page } from '@playwright/test';
import { NOW, seedData } from './server.ts';

export { expect };
export { DATA, NOW, readCell, seedData, writeCells } from './server.ts';

export const test = base.extend<{ seed: (fixture: string) => void }>({
  page: async ({ page }, use) => {
    await page.clock.install({ time: NOW });
    await use(page);
  },
  // eslint-disable-next-line no-empty-pattern -- Playwright reads fixture names from this pattern.
  seed: async ({}, use) => {
    await use(seedData);
  },
});

/** Open the app the way a user does and wait until it has drawn: a key pressed sooner goes nowhere. */
export async function openApp(page: Page): Promise<void> {
  await page.goto('/');
  await page.getByRole('main').waitFor();
}

/**
 * Run a command from ⌘K by its name. Type the way a person does: wait until focus is in the field, then use the
 * keyboard. `fill()` right after a field appears can land before the component has set itself up, and be lost.
 */
export async function runCommand(page: Page, name: string): Promise<void> {
  await page.keyboard.press('ControlOrMeta+KeyK');
  const menu = page.getByRole('dialog', { name: 'Command menu' });
  await expect(menu.getByRole('combobox')).toBeFocused();
  await page.keyboard.type(name);
  await page.keyboard.press('Enter');
  await expect(menu).toBeHidden();
}

/** Open a workbook from the E2E folder with Open… (⌘O), the way a user picks it from the NAS. */
export async function openFile(page: Page, name: string): Promise<void> {
  await page.keyboard.press('ControlOrMeta+KeyO');
  const dialog = page.getByRole('dialog', { name: 'Open from NAS' });
  await dialog.getByRole('option', { name }).click();
  await dialog.getByRole('button', { name: 'Open', exact: true }).click();
  await expect(dialog).toBeHidden();
  await expect(page).toHaveTitle(`${name} — Cellestra`);
}

/** Select a cell by typing its address in the Name Box, as a keyboard user does. */
export async function goTo(page: Page, cell: string): Promise<void> {
  const nameBox = page.getByRole('textbox', { name: 'Cell reference' });
  await nameBox.click();
  await expect(nameBox).toBeFocused();
  await page.keyboard.press('ControlOrMeta+KeyA');
  await page.keyboard.type(cell);
  await page.keyboard.press('Enter');
  await expect(nameBox).toHaveValue(cell);
}

/** Type into a cell and press Enter. */
export async function enter(page: Page, cell: string, text: string): Promise<void> {
  await goTo(page, cell);
  await page.keyboard.type(text);
  await page.keyboard.press('Enter');
}

/** The formula bar, which shows what the selected cell holds: its text, number or formula. */
export function contents(page: Page) {
  return page.getByRole('textbox', { name: 'Cell contents' });
}

/** Save with ⌘S and wait until the window says nothing is unsaved. */
export async function save(page: Page): Promise<void> {
  await page.keyboard.press('ControlOrMeta+KeyS');
  await expect(page).not.toHaveTitle(/^• /);
}
