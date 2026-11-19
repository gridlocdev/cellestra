// Starting up, the command menu, the shortcuts window and what a screen reader hears.
import { contents, expect, goTo, openApp, openFile, runCommand, test } from './fixtures.ts';

test.describe('The keyboard', () => {
  test.beforeEach(async ({ seed, page }) => {
    seed('basic');
    await openApp(page);
  });

  test('starts with no errors in the console', async ({ page }) => {
    const errors: string[] = [];
    page.on('console', (message) => {
      if (message.type() === 'error') errors.push(message.text());
    });
    page.on('pageerror', (error) => errors.push(error.message));
    await openApp(page);
    await openFile(page, 'budget.ods');
    expect(errors).toEqual([]);
  });

  test('the command menu runs a command', async ({ page }) => {
    await runCommand(page, 'Insert Sheet');
    await expect(page.getByRole('tab', { name: 'Sheet2' })).toBeVisible();
  });

  test('the command menu shows each command’s shortcut', async ({ page }) => {
    await page.keyboard.press('ControlOrMeta+KeyK');
    const menu = page.getByRole('dialog', { name: 'Command menu' });
    await expect(menu.getByRole('combobox')).toBeFocused();
    await page.keyboard.type('Bold');
    await expect(menu.getByRole('option').first()).toContainText('⌘B');
    await page.keyboard.press('Escape');
    await expect(menu).toBeHidden();
  });

  test('⌘K does nothing while another dialog is open', async ({ page }) => {
    await openFile(page, 'budget.ods');
    await runCommand(page, 'Go to Sheet');
    const picker = page.getByRole('dialog', { name: 'Go to sheet' });
    await expect(picker.getByRole('combobox')).toBeFocused();
    await page.keyboard.press('ControlOrMeta+KeyK');
    await expect(page.getByRole('dialog', { name: 'Command menu' })).toHaveCount(0);
    await page.keyboard.press('Escape');
    await expect(picker).toBeHidden();
    await page.getByRole('grid', { name: 'Spreadsheet' }).focus();
    await runCommand(page, 'Settings');
    await expect(page.getByRole('dialog', { name: 'Settings' })).toBeVisible();
    // Focus left on the grid under the dialog: the grid's own shortcuts must not open the menu either.
    await page.getByRole('grid', { name: 'Spreadsheet' }).focus();
    await page.keyboard.press('ControlOrMeta+KeyK');
    await expect(page.getByRole('dialog', { name: 'Command menu' })).toHaveCount(0);
  });

  test('keys typed as soon as a dialog opens go into it, not the grid behind it', async ({ page }) => {
    await openFile(page, 'budget.ods');
    await runCommand(page, 'Insert Function');
    await page.keyboard.type('ABS');
    await expect(page.getByRole('dialog', { name: 'Insert function' }).getByRole('combobox')).toHaveValue(
      'ABS',
    );
    await page.keyboard.press('Enter');
    await page.keyboard.type('-5)');
    await page.keyboard.press('Enter');
    await goTo(page, 'A1');
    await expect(contents(page)).toHaveValue('=ABS(-5)');
  });

  test('⌘Z in a dialog does not put old typing back into a cell editor behind it', async ({ page }) => {
    await openFile(page, 'budget.ods');
    await goTo(page, 'E1');
    await page.keyboard.type('typed');
    await page.keyboard.press('Enter');
    await runCommand(page, 'Settings');
    await expect(page.getByRole('dialog', { name: 'Settings' })).toBeVisible();
    await page.keyboard.press('ControlOrMeta+KeyZ');
    await page.keyboard.press('Escape');
    await goTo(page, 'B1');
    await expect(contents(page)).toHaveValue('Rent');
  });

  test('the Keyboard Shortcuts window lists the shortcuts', async ({ page }) => {
    await page.keyboard.press('ControlOrMeta+Slash');
    const shortcuts = page.getByRole('dialog', { name: 'Keyboard Shortcuts' });
    await expect(shortcuts).toContainText('Bold⌘B');
    await expect(shortcuts.getByRole('button', { name: 'Done' })).toBeFocused();
    await page.keyboard.press('Escape');
    await expect(shortcuts).toBeHidden();
  });

  test('arrow keys move the selection', async ({ page }) => {
    await openFile(page, 'budget.ods');
    await goTo(page, 'A1');
    await page.getByRole('grid', { name: 'Spreadsheet' }).focus();
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowRight');
    await expect(page.getByRole('textbox', { name: 'Cell reference' })).toHaveValue('B2');
    await expect(contents(page)).toHaveValue('1200');
  });

  test('Enter after Shift+Tab goes straight down', async ({ page }) => {
    await openFile(page, 'budget.ods');
    await goTo(page, 'B1');
    await page.getByRole('grid', { name: 'Spreadsheet' }).focus();
    await page.keyboard.press('Enter');
    await page.keyboard.press('Shift+Tab');
    await page.keyboard.press('Enter');
    await expect(page.getByRole('textbox', { name: 'Cell reference' })).toHaveValue('A3');
  });

  test('typing a row with Tab, Enter starts the next row where the typing began', async ({ page }) => {
    await openFile(page, 'budget.ods');
    await goTo(page, 'E1');
    await page.getByRole('grid', { name: 'Spreadsheet' }).focus();
    for (const text of ['a', 'b']) {
      await page.keyboard.type(text);
      await page.keyboard.press('Tab');
    }
    await page.keyboard.type('c');
    await page.keyboard.press('Enter');
    await expect(page.getByRole('textbox', { name: 'Cell reference' })).toHaveValue('E2');
  });

  test('↓ while typing after Tabs goes straight down', async ({ page }) => {
    await openFile(page, 'budget.ods');
    await goTo(page, 'E1');
    await page.getByRole('grid', { name: 'Spreadsheet' }).focus();
    for (const text of ['a', 'b']) {
      await page.keyboard.type(text);
      await page.keyboard.press('Tab');
    }
    await page.keyboard.type('c');
    await page.keyboard.press('ArrowDown');
    await expect(page.getByRole('textbox', { name: 'Cell reference' })).toHaveValue('G2');
  });

  test('Enter goes straight down once the selection has moved away from the Tabs and back', async ({
    page,
  }) => {
    await openFile(page, 'budget.ods');
    await goTo(page, 'E1');
    const grid = page.getByRole('grid', { name: 'Spreadsheet' });
    await grid.focus();
    await page.keyboard.press('Tab');
    await page.keyboard.press('Tab');
    await goTo(page, 'B5');
    await goTo(page, 'G1');
    await grid.focus();
    await page.keyboard.press('Enter');
    await expect(page.getByRole('textbox', { name: 'Cell reference' })).toHaveValue('G2');
  });
});

test.describe('Accessibility', () => {
  test.beforeEach(async ({ seed, page }) => {
    seed('basic');
    await openApp(page);
  });

  test('a screen reader hears the cell address and value as the selection moves', async ({ page }) => {
    await openFile(page, 'budget.ods');
    // The status bar shows the same words, so look for them in the live region a screen reader follows.
    const heard = page.locator('[aria-live="polite"]').first();
    await goTo(page, 'B2');
    await expect(heard).toHaveText('B2, 1200');
    await goTo(page, 'F9');
    await expect(heard).toHaveText('F9, blank');
  });

  test('a dialog keeps focus inside it and gives it back when it closes', async ({ page }) => {
    await page.keyboard.press('ControlOrMeta+KeyH');
    const dialog = page.getByRole('dialog', { name: 'Find and Replace' });
    await expect(dialog.getByRole('textbox', { name: 'Find' })).toBeFocused();
    for (let i = 0; i < 12; i++) await page.keyboard.press('Tab');
    await expect(dialog.locator(':focus')).toHaveCount(1);
    await page.keyboard.press('Escape');
    await expect(dialog).toBeHidden();
    await expect(page.getByRole('grid', { name: 'Spreadsheet' })).toBeFocused();
  });
});
