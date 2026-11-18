// Editing a workbook: formulas, formatting, find and replace, data tools and charts, checked on screen and on disk.
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

test.beforeEach(async ({ seed, page }) => {
  seed('basic');
  await openApp(page);
  await openFile(page, 'budget.ods');
});

test.describe('Formulas', () => {
  test('a formula recalculates when a cell it reads changes', async ({ page }) => {
    await enter(page, 'B2', '1000');
    await save(page);
    expect(readCell('budget.ods', 'D2')).toBe('1310');
  });

  test('a new formula shows its source in the formula bar and its result in the file', async ({ page }) => {
    await enter(page, 'B4', '=SUM(B2:B3)');
    await goTo(page, 'B4');
    await expect(contents(page)).toHaveValue('=SUM(B2:B3)');
    await save(page);
    expect(readCell('budget.ods', 'B4')).toBe('2400');
  });
});

test.describe('Formatting', () => {
  test('Bold shows as on in the ribbon for the bold cell only', async ({ page }) => {
    const bold = page.getByRole('button', { name: 'Bold', exact: true });
    await goTo(page, 'A2');
    await page.keyboard.press('ControlOrMeta+KeyB');
    await expect(bold).toHaveAttribute('aria-pressed', 'true');
    await goTo(page, 'A3');
    await expect(bold).toHaveAttribute('aria-pressed', 'false');
  });
});

test.describe('Find and Replace', () => {
  test('Replace All replaces every match and the file has it', async ({ page }) => {
    await page.keyboard.press('ControlOrMeta+KeyH');
    const dialog = page.getByRole('dialog', { name: 'Find and Replace' });
    await expect(dialog.getByRole('textbox', { name: 'Find' })).toBeFocused();
    await page.keyboard.type('uary');
    await expect(dialog).toContainText('2 matches');
    await dialog.getByRole('textbox', { name: 'Replace with' }).click();
    await page.keyboard.type('.');
    await dialog.getByRole('button', { name: 'Replace All' }).click();
    await expect(dialog).toContainText('Replaced 2 cell(s).');
    await page.keyboard.press('Escape');
    await save(page);
    expect(readCell('budget.ods', 'A2')).toBe('Jan.');
    expect(readCell('budget.ods', 'A3')).toBe('Febr.');
  });
});

test.describe('Data tools', () => {
  test('Sort Ascending sorts whole rows by the first column', async ({ page }) => {
    await goTo(page, 'A2:D3');
    await runCommand(page, 'Sort Ascending');
    await save(page);
    expect(readCell('budget.ods', 'A2')).toBe('February');
    expect(readCell('budget.ods', 'C2'), 'the row moves together').toBe('285');
    expect(readCell('budget.ods', 'D2'), 'its total moves with it').toBe('1485');
  });

  test('A dropdown list colours its items and fills a cell from the list', async ({ page }) => {
    await goTo(page, 'F2:F4');
    await runCommand(page, 'Data Validation');
    const dialog = page.getByRole('dialog', { name: 'Data Validation' });
    await dialog.getByLabel('Items (comma-separated)').fill('Yes - native, No - not on Steam');
    await expect(dialog.getByLabel('Colour of No - not on Steam')).toHaveValue('Orange');
    await dialog.getByLabel('Colour of No - not on Steam').selectOption('Red');
    await dialog.getByRole('button', { name: 'OK' }).click();
    await expect(dialog).toBeHidden();

    await goTo(page, 'F3');
    await runCommand(page, 'Choose from List');
    const option = page.getByRole('option', { name: 'No - not on Steam' });
    await expect(option.locator('.cs-chip')).toHaveCSS('background-color', 'rgb(255, 199, 206)');
    await option.click();
    await expect(contents(page)).toHaveValue('No - not on Steam');
    await save(page);
    expect(readCell('budget.ods', 'F3')).toBe('No - not on Steam');
  });
});

test.describe('Charts', () => {
  test('Chart Sheet adds a sheet with a chart of the selection, saved in the file', async ({ page }) => {
    await goTo(page, 'A1:C3');
    await runCommand(page, 'Chart Sheet');
    await expect(page.getByRole('tablist', { name: 'Sheets' }).getByRole('tab')).toHaveCount(2);
    await save(page);
    await openApp(page);
    await openFile(page, 'budget.ods');
    await expect(page.getByRole('tablist', { name: 'Sheets' }).getByRole('tab')).toHaveCount(2);
  });
});
