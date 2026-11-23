// Sheets, the grid and how it looks.
import {
  contents,
  enter,
  expect,
  goTo,
  openApp,
  openFile,
  readCell,
  runCommand,
  save,
  test,
} from './fixtures.ts';

test.describe('Sheets and the grid', () => {
  test.beforeEach(async ({ seed, page }) => {
    seed('basic');
    await openApp(page);
    await openFile(page, 'budget.ods');
  });

  test('a new sheet is saved in the file with what was typed on it', async ({ page }) => {
    await runCommand(page, 'Insert Sheet');
    await expect(page.getByRole('tab', { name: 'Sheet2', selected: true })).toBeVisible();
    await enter(page, 'A1', 'On the second sheet');
    await save(page);
    expect(readCell('budget.ods', 'Sheet2.A1')).toBe('On the second sheet');
    expect(readCell('budget.ods', 'Sheet1.A1'), 'the first sheet is untouched').toBe('Month');
  });

  test('a click on a sheet tab shows that sheet', async ({ page }) => {
    await runCommand(page, 'Insert Sheet');
    await page.getByRole('tab', { name: 'Sheet1' }).click();
    await goTo(page, 'A1');
    await expect(contents(page)).toHaveValue('Month');
  });

  test('the sheet tabs fill the footer up to the status text, and the active tab stays in view', async ({
    page,
  }) => {
    for (let i = 0; i < 16; i++) await runCommand(page, 'Insert Sheet');
    await enter(page, 'A1', 'Quarterly totals');
    const strip = page.getByRole('tablist', { name: 'Sheets' });
    const status = page.getByText('Edited cell A1');
    const [tabs, text] = [await strip.boundingBox(), await status.boundingBox()];
    expect(text!.x - (tabs!.x + tabs!.width), 'no gap between the tabs and the status text').toBeLessThan(24);
    await expect(page.getByRole('tab', { name: 'Sheet17', selected: true })).toBeInViewport({ ratio: 1 });
    const active = await page.getByRole('tab', { name: 'Sheet17' }).boundingBox();
    expect(active!.x + active!.width).toBeLessThanOrEqual(tabs!.x + tabs!.width + 1);
  });

  test('the status bar sums the selection', async ({ page }) => {
    await goTo(page, 'B2:C3');
    await expect(page.getByText('Sum 2,995')).toBeVisible();
  });

  test('looks right in light and dark', async ({ page }) => {
    await goTo(page, 'D2');
    await expect(page).toHaveScreenshot('budget-light.png');
    await page.emulateMedia({ colorScheme: 'dark' });
    await expect(page).toHaveScreenshot('budget-dark.png');
  });
});
