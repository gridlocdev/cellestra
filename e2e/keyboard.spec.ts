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
    await expect(shortcuts.getByRole('button', { name: 'Look Up a Key…' })).toBeFocused();
    await page.keyboard.press('Escape');
    await expect(shortcuts).toBeHidden();
  });

  test('Look Up a Key says what a key does without running it, and Escape gets out', async ({ page }) => {
    await page.keyboard.press('ControlOrMeta+Slash');
    const shortcuts = page.getByRole('dialog', { name: 'Keyboard Shortcuts' });
    await shortcuts.getByRole('button', { name: 'Look Up a Key…' }).click();
    await expect(shortcuts.getByRole('button', { name: 'Stop' })).toBeVisible();
    // ⌘/ would close this window if it ran; here it is only looked up.
    await page.keyboard.press('ControlOrMeta+Slash');
    await expect(shortcuts).toContainText('Keyboard Shortcuts (Help)');
    await page.keyboard.press('ControlOrMeta+KeyB');
    await expect(shortcuts).toContainText('⌘B: Bold');
    await expect(shortcuts.locator('[data-command="format.bold"]')).toHaveClass(/found/);
    await page.keyboard.press('ArrowDown');
    await expect(shortcuts).toContainText('the grid uses it');
    // The first Escape only stops looking up; the second closes the window as usual.
    await page.keyboard.press('Escape');
    await expect(shortcuts.getByRole('button', { name: 'Look Up a Key…' })).toBeVisible();
    await expect(shortcuts).toBeVisible();
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
  test('no control shows the browser’s own focus ring, and each one shows the app’s', async ({ page }) => {
    await openFile(page, 'budget.ods');
    const rings = () =>
      page.evaluate(() => {
        const found: string[] = [];
        const all = document.querySelectorAll<HTMLInputElement>(
          'button, input, select, textarea, [tabindex], a[href], [contenteditable="true"]',
        );
        for (const el of all) {
          // The search box opens the command menu as it takes focus; the menu's own box is checked in a dialog.
          if (el.closest('[hidden], [inert]') || el.getClientRects().length === 0 || el.readOnly) continue;
          el.focus({ focusVisible: true } as FocusOptions);
          if (document.activeElement !== el) continue;
          const style = getComputedStyle(el);
          if (style.outlineStyle === 'auto') found.push(`${el.tagName}.${el.className}`);
        }
        return found;
      });
    expect(await rings()).toEqual([]);
    await page.getByRole('grid', { name: 'Spreadsheet' }).focus();
    await page.keyboard.press('ControlOrMeta+Slash');
    await expect(page.getByRole('dialog', { name: 'Keyboard Shortcuts' })).toBeVisible();
    expect(await rings()).toEqual([]);
  });

  test('arrow keys move the current sheet tab and the focus together', async ({ page }) => {
    await runCommand(page, 'Insert Sheet');
    const first = page.getByRole('tab', { name: 'Sheet1' });
    const second = page.getByRole('tab', { name: 'Sheet2' });
    await first.click();
    await first.focus();
    await expect(first).toHaveAttribute('data-focused', '');
    await page.keyboard.press('ArrowRight');
    await expect(second).toHaveAttribute('aria-selected', 'true');
    await expect(second).toBeFocused();
    await expect(second).toHaveAttribute('data-focused', '');
    await expect(first).not.toHaveAttribute('data-focused', '');
    await expect(first).toHaveAttribute('tabindex', '-1');
    await page.keyboard.press('Home');
    await expect(first).toBeFocused();
    await expect(first).toHaveAttribute('aria-selected', 'true');
  });

  test('arrow keys move the ribbon tab and the focus together, and focus comes back after a dialog', async ({
    page,
  }) => {
    const tabs = page.locator('.rb-tabs [role="tab"]');
    const start = await tabs.evaluateAll((all) =>
      all.findIndex((t) => t.getAttribute('aria-selected') === 'true'),
    );
    await tabs.nth(start).focus();
    await page.keyboard.press('ArrowRight');
    const next = tabs.nth(start + 1);
    await expect(next).toHaveAttribute('aria-selected', 'true');
    await expect(next).toBeFocused();
    await expect(next).toHaveAttribute('data-focused', '');
    const outline = await next.evaluate((el) => getComputedStyle(el).outlineStyle);
    expect(outline).toBe('solid');
    await page.keyboard.press('ControlOrMeta+Slash');
    const dialog = page.getByRole('dialog', { name: 'Keyboard Shortcuts' });
    await expect(dialog).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(dialog).toBeHidden();
    await expect(next).toBeFocused();
  });

  test('F6 and Shift+F6 cycle the focus through the ribbon, the grid and the sheet tabs, onto each current item', async ({
    page,
  }) => {
    await runCommand(page, 'Insert Sheet');
    const grid = page.getByRole('grid', { name: 'Spreadsheet' });
    const ribbonTab = page.locator('.rb-tabs [aria-selected="true"]');
    const sheetTab = page.getByRole('tab', { name: 'Sheet2' });
    await grid.click();
    await expect(grid).toBeFocused();
    await page.keyboard.press('F6');
    await expect(sheetTab).toBeFocused();
    await expect(sheetTab).toHaveAttribute('data-focused', '');
    await page.keyboard.press('F6');
    await expect(ribbonTab).toBeFocused();
    // An arrow moves the ribbon's current tab, and the next visit lands on it.
    await page.keyboard.press('ArrowRight');
    const moved = page.locator('.rb-tabs [aria-selected="true"]');
    const name = await moved.textContent();
    await page.keyboard.press('F6');
    await expect(grid).toBeFocused();
    await page.keyboard.press('Shift+F6');
    await expect(page.getByRole('tab', { name: name!, exact: true })).toBeFocused();
    await page.keyboard.press('Shift+F6');
    await expect(sheetTab).toBeFocused();
    await page.keyboard.press('Shift+F6');
    await expect(grid).toBeFocused();
  });

  test('F6 reaches the compact toolbar in place of the ribbon', async ({ page }) => {
    await runCommand(page, 'Ribbon Layout');
    const toolbar = page.getByRole('toolbar', { name: 'Commands' });
    await expect(toolbar).toBeVisible();
    const grid = page.getByRole('grid', { name: 'Spreadsheet' });
    await grid.click();
    await page.keyboard.press('Shift+F6');
    await expect(toolbar.locator('[data-focused]')).toBeFocused();
    await page.keyboard.press('F6');
    await expect(grid).toBeFocused();
  });
});
