// Opening and saving workbooks on the NAS: the disk is checked with the CLI, not just the screen.
import {
  contents,
  enter,
  expect,
  goTo,
  openApp,
  openFile,
  readCell,
  save,
  test,
  writeCells,
} from './fixtures.ts';

/** How often the app looks for a change made elsewhere to the open NAS file, plus a little. */
const CHANGE_CHECK_MS = 31_000;
/** Auto-save's default delay after the last edit, plus a little. */
const AUTO_SAVE_MS = 31_000;

test.describe('Files', () => {
  test.beforeEach(async ({ seed, page }) => {
    seed('basic');
    await openApp(page);
  });

  test('opens a workbook from the NAS with its formulas', async ({ page }) => {
    await openFile(page, 'budget.ods');
    await goTo(page, 'D2');
    await expect(contents(page)).toHaveValue('=B2+C2');
  });

  test('opens a workbook with a Unicode name and contents', async ({ page }) => {
    await openFile(page, 'Café ☕ notes.ods');
    await goTo(page, 'A2');
    await expect(contents(page)).toHaveValue('日本語のテキスト');
  });

  test('Open from NAS says it is loading, then puts focus on the list, with times to the minute', async ({
    page,
  }) => {
    let release = () => {};
    const held = new Promise<void>((r) => (release = r));
    await page.route('**/api/list?**', async (route) => {
      await held;
      await route.continue();
    });
    await page.keyboard.press('ControlOrMeta+KeyO');
    const dialog = page.getByRole('dialog', { name: 'Open from NAS' });
    await expect(dialog.getByText('Loading…')).toBeVisible();
    await expect(dialog.getByText('No spreadsheets here')).toHaveCount(0);
    release();
    const first = dialog.getByRole('option').first();
    await expect(first).toBeFocused();
    await page.keyboard.press('ArrowDown');
    await expect(dialog.getByRole('option').nth(1)).toBeFocused();
    const times = await dialog.locator('.cs-meta').allTextContents();
    expect(times.length).toBeGreaterThan(0);
    for (const t of times) expect(t).not.toMatch(/\d:\d\d:\d\d/);
  });

  test('opens an .xlsx as a new .ods workbook, and the first save asks where to put it', async ({ page }) => {
    await page.keyboard.press('ControlOrMeta+KeyO');
    const dialog = page.getByRole('dialog', { name: 'Open from NAS' });
    await dialog.getByRole('option', { name: 'expenses.xlsx' }).click();
    await dialog.getByRole('button', { name: 'Open', exact: true }).click();
    await expect(dialog).toBeHidden();
    await expect(page).toHaveTitle('expenses.ods — Cellestra');
    await expect(page.getByRole('status')).toContainText(
      'Opened expenses.xlsx as a new workbook, expenses.ods.',
    );
    await goTo(page, 'D2');
    await expect(contents(page)).toHaveValue('=B2+C2');
    await page.keyboard.press('ControlOrMeta+KeyS');
    await expect(page.getByRole('dialog', { name: 'Save to NAS' })).toBeVisible();
  });

  test('an edit is saved to disk with ⌘S and is there after a reload', async ({ page }) => {
    await openFile(page, 'budget.ods');
    await enter(page, 'B2', '1300');
    await expect(page).toHaveTitle(/^• /);
    await save(page);
    expect(readCell('budget.ods', 'B2')).toBe('1300');
    expect(readCell('budget.ods', 'D2'), 'the saved file has the recalculated total').toBe('1610');
    await openApp(page);
    await openFile(page, 'budget.ods');
    await goTo(page, 'B2');
    await expect(contents(page)).toHaveValue('1300');
  });

  test('auto-save writes an edit 30 seconds after the last one', async ({ page }) => {
    await openFile(page, 'budget.ods');
    await enter(page, 'A4', 'March');
    await page.clock.runFor(AUTO_SAVE_MS);
    await expect(page).not.toHaveTitle(/^• /);
    expect(readCell('budget.ods', 'A4')).toBe('March');
  });

  test('a change made elsewhere is offered when nothing is unsaved', async ({ page }) => {
    await openFile(page, 'budget.ods');
    writeCells('budget.ods', 'A1=Changed elsewhere');
    await page.clock.runFor(CHANGE_CHECK_MS);
    const offer = page.getByRole('dialog', { name: 'File changed on the server' });
    await offer.getByRole('button', { name: 'Reload' }).click();
    await expect(offer).toBeHidden();
    await goTo(page, 'A1');
    await expect(contents(page)).toHaveValue('Changed elsewhere');
  });

  test('a change made elsewhere keeps typed text, and ⌘S asks before overwriting it', async ({ page }) => {
    await openFile(page, 'budget.ods');
    await enter(page, 'B2', '999');
    writeCells('budget.ods', 'A1=Changed elsewhere');
    await page.clock.runFor(CHANGE_CHECK_MS);
    await goTo(page, 'B2');
    await expect(contents(page), 'the typed text stays').toHaveValue('999');

    await page.keyboard.press('ControlOrMeta+KeyS');
    const conflict = page.getByRole('dialog', { name: 'File changed', exact: true });
    await expect(conflict).toBeVisible();
    expect(readCell('budget.ods', 'B2'), 'nothing is written before the user picks').toBe('1200');
    await conflict.getByRole('button', { name: 'Overwrite' }).click();
    await expect(page).not.toHaveTitle(/^• /);
    expect(readCell('budget.ods', 'B2')).toBe('999');
    expect(readCell('budget.ods', 'A1'), 'overwrite keeps this window’s version').toBe('Month');
  });

  test('Reload in the save conflict replaces the typed text with the file', async ({ page }) => {
    await openFile(page, 'budget.ods');
    await enter(page, 'B2', '999');
    writeCells('budget.ods', 'A1=Changed elsewhere');
    await page.keyboard.press('ControlOrMeta+KeyS');
    await page
      .getByRole('dialog', { name: 'File changed', exact: true })
      .getByRole('button', { name: 'Reload' })
      .click();
    await expect(page).not.toHaveTitle(/^• /);
    await goTo(page, 'A1');
    await expect(contents(page)).toHaveValue('Changed elsewhere');
    await goTo(page, 'B2');
    await expect(contents(page)).toHaveValue('1200');
    expect(readCell('budget.ods', 'B2'), 'reload writes nothing').toBe('1200');
  });

  test('Close asks about unsaved changes, and Cancel keeps them', async ({ page }) => {
    await openFile(page, 'budget.ods');
    await enter(page, 'B2', '999');
    await page.keyboard.press('ControlOrMeta+KeyW');
    const ask = page.getByRole('dialog', { name: 'Unsaved changes' });
    // Enter must save, not throw the changes away: focus starts on Save.
    await expect(ask.getByRole('button', { name: 'Save', exact: true })).toBeFocused();
    await ask.getByRole('button', { name: 'Cancel' }).click();
    await expect(ask).toBeHidden();
    await expect(page).toHaveTitle('• budget.ods — Cellestra');
    await goTo(page, 'B2');
    await expect(contents(page)).toHaveValue('999');
  });
});
