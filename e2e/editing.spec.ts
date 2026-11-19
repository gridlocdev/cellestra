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
  test('A list from another sheet offers that sheet’s items, and long ones wrap in the dialog', async ({
    page,
  }) => {
    const long =
      'Catalog: every game in the subscription as listed on the day, plus the additions announced since then';
    await runCommand(page, 'Insert Sheet');
    await enter(page, 'A1', 'Unclear');
    await enter(page, 'A2', long);
    await page.getByRole('tab', { name: 'Sheet1' }).click();
    await goTo(page, 'F2:F4');
    await runCommand(page, 'Data Validation');
    const dialog = page.getByRole('dialog', { name: 'Data Validation' });
    await dialog.getByLabel('…or take them from a range').fill('Sheet2.A1:A2');
    await expect(dialog.getByLabel('Colour of Unclear')).toBeVisible();
    await expect(dialog.getByLabel(`Colour of ${long}`)).toBeVisible();
    await expect(dialog.getByLabel('Colour of Month'), 'not the active sheet’s cells').toHaveCount(0);
    const colours = dialog.locator('.cs-chip-colors');
    const fits = await colours.evaluate((el) => ({
      scrolls: el.scrollHeight > el.clientHeight + 1,
      inside:
        el.getBoundingClientRect().right <=
        el.closest('dialog, [role=dialog]')!.getBoundingClientRect().right,
    }));
    expect(fits).toEqual({ scrolls: false, inside: true });
    const widths = await colours.locator('select').evaluateAll((els) => els.map((el) => el.clientWidth));
    expect(new Set(widths).size, 'every colour menu is the same width').toBe(1);
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

  test('Make a Table turns the chart’s cells into a table it plots by name, and Undo puts the cells back', async ({
    page,
  }) => {
    await goTo(page, 'A1:C3');
    await page.keyboard.press('F11');
    const panel = page.getByRole('complementary', { name: 'Chart data' });
    await expect(panel.getByText(/^From cells /)).toBeVisible();
    await panel.getByRole('button', { name: 'Make a Table…' }).click();
    const form = panel.getByRole('form', { name: 'Make a table' });
    await expect(form.getByRole('textbox', { name: 'Table name' })).toBeFocused();
    await page.keyboard.press('ControlOrMeta+KeyA');
    await page.keyboard.type('Costs');
    await form.getByRole('button', { name: 'Create Table' }).click();
    await expect(panel.getByText('From table “Costs”: Rent and Groceries by Month.')).toBeVisible();
    await expect(panel.getByRole('radio', { name: 'Table' })).toBeChecked();
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await expect(panel.getByText(/^From cells /)).toBeVisible();
  });

  test('a range that can’t plot says why under its field', async ({ page }) => {
    await goTo(page, 'A1:C3');
    await runCommand(page, 'Chart Sheet');
    const field = page.getByRole('textbox', { name: 'Category labels' });
    await field.click();
    await page.keyboard.press('ControlOrMeta+KeyA');
    await page.keyboard.type('Nowhere.A2:A3');
    await page.keyboard.press('Enter');
    await expect(field).toHaveAttribute('aria-invalid', 'true');
    await expect(page.getByText('There is no sheet called “Nowhere”.')).toBeVisible();
  });

  test('removing a series offers Undo in the status bar', async ({ page }) => {
    await goTo(page, 'A1:C3');
    await runCommand(page, 'Chart Sheet');
    const series = page.locator('.cs-chart-series');
    await expect(series).toHaveCount(2);
    await page.getByRole('button', { name: 'Remove Rent' }).click();
    await expect(series).toHaveCount(1);
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await expect(series).toHaveCount(2);
  });

  test('a range field shows the first values it holds', async ({ page }) => {
    await goTo(page, 'A1:C3');
    await runCommand(page, 'Chart Sheet');
    const panel = page.getByRole('complementary', { name: 'Chart data' });
    await expect(panel.getByText('January, February', { exact: true })).toBeVisible();
  });

  test('Edit Data Range re-plots the chart from the block picked on the sheet', async ({ page }) => {
    await goTo(page, 'A1:C3');
    await runCommand(page, 'Chart Sheet');
    await expect(page.locator('.cs-chart-series')).toHaveCount(2);
    await page.getByRole('button', { name: 'Edit Data Range…' }).click();
    await expect(page.getByText(/^Select the data for the chart/)).toBeVisible();
    await goTo(page, 'A1:B3');
    await page.getByRole('button', { name: 'Use selection' }).click();
    await expect(page.locator('.cs-chart-series')).toHaveCount(1);
    await expect(page.getByRole('button', { name: 'Remove Rent' })).toBeVisible();
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await expect(page.locator('.cs-chart-series')).toHaveCount(2);
  });

  test('the Chart commands change the type and swap rows and columns', async ({ page }) => {
    await goTo(page, 'A1:C3');
    await runCommand(page, 'Chart Sheet');
    await runCommand(page, 'Bar Chart');
    await expect(page.getByRole('radio', { name: 'Bar' })).toBeChecked();
    await page.getByRole('radio', { name: 'Bar' }).press('ArrowRight');
    await expect(page.getByRole('radio', { name: 'Line' })).toBeChecked();
    await runCommand(page, 'Swap Rows and Columns');
    await expect(page.getByRole('button', { name: 'Remove January' })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Remove February' })).toBeVisible();
  });
});
