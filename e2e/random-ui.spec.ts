// Random use of the whole app: seeded runs of commands from the ⌘K menu, typing and keys on a fixture workbook,
// checking after every step that nothing broke. `just e2e` runs a few seeds; `just soak` tries 25 new ones.
// A failure prints its seed: `FUZZ_SEED=<n> just e2e -g random` replays it, with every step. Keep a found bug
// as its own scenario.
import { readFileSync } from 'node:fs';
import type { Page } from '@playwright/test';
import {
  contents,
  DATA,
  expect,
  goTo,
  openApp,
  openFile,
  runCommand,
  test,
  watchErrors,
} from './fixtures.ts';
import { createRandom } from './randomOps.ts';

const STEPS = 40;
const FILE = 'budget.ods';

/**
 * Commands skipped, matched against the title of each ⌘K entry:
 * - ones that show the system's file picker or download (Save As, export, import, open, print);
 * - ones that leave this workbook for another (New, Close, Open Recent), so the end check has nothing to compare;
 * - the clipboard, which in a browser can be the owner's real one.
 */
const SKIP = [
  /^New$/,
  /^Open/,
  /^Close$/,
  /^Save As/,
  /^Download a Copy$/,
  /^Import CSV/,
  /^Export/,
  /^Print/,
  /^(Cut|Copy|Paste|Paste Values)$/,
];

/** Keys worth pressing on their own, besides the letters random typing presses: the grid's own keys among them. */
const KEYS = [
  'Escape',
  'Enter',
  'Tab',
  'Shift+Tab',
  'Space',
  'ArrowUp',
  'ArrowDown',
  'ArrowLeft',
  'ArrowRight',
  'Shift+ArrowDown',
  'Shift+ArrowRight',
  'Home',
  'End',
  'PageDown',
  'Backspace',
  'Delete',
  'F2',
  'F9',
  'ControlOrMeta+KeyZ',
  'ControlOrMeta+Shift+KeyZ',
  'ControlOrMeta+KeyA',
  'ControlOrMeta+KeyB',
  'ControlOrMeta+KeyD',
  'Alt+ArrowDown',
];

const runs = Number(process.env.FUZZ_RUNS ?? 3);
const base = Number(process.env.FUZZ_SEED_BASE ?? 1);
const seeds = process.env.FUZZ_SEED
  ? [Number(process.env.FUZZ_SEED)]
  : Array.from({ length: runs }, (_, i) => base + i);

/**
 * Open ⌘K and click one of its entries. Clicking runs exactly that entry (typing its name may match another one
 * first), and disabled entries are skipped (a click on one waits forever). Options are looked up inside the menu,
 * because other lists have options too. Returns what was run, or why nothing was.
 */
async function runRandomCommand(page: Page, pick: (count: number) => number): Promise<string> {
  await page.keyboard.press('ControlOrMeta+KeyK');
  const menu = page.getByRole('dialog', { name: 'Command menu' });
  const opened = await menu
    .getByRole('combobox')
    .waitFor({ timeout: 1000 })
    .then(
      () => true,
      () => false,
    );
  // ⌘K opens over no dialog but its own kind, so with one open this is expected: the per-step check makes sure
  // that dialog kept the keyboard.
  if (!opened) return 'command menu did not open';
  const options = menu.getByRole('option');
  const entries = await options.evaluateAll((items) =>
    items.map((item) => ({
      // The title is the entry's one span without a class; the others are its icon, group and shortcut.
      label: (item.querySelector(':scope > span:not([class])') ?? item).textContent?.trim() ?? '',
      enabled: item.getAttribute('aria-disabled') !== 'true',
    })),
  );
  const usable = entries
    .map((entry, i) => ({ ...entry, i }))
    .filter(({ label, enabled }) => enabled && !SKIP.some((pattern) => pattern.test(label)));
  const chosen = usable[pick(usable.length)];
  if (!chosen) {
    await page.keyboard.press('Escape');
    return 'no command to run';
  }
  // A dialog a command opens a moment later (a confirmation) can replace the menu before the click lands.
  const clicked = await options
    .nth(chosen.i)
    .click({ timeout: 1000 })
    .then(
      () => true,
      () => false,
    );
  if (!clicked && (await menu.count()) === 0) return `command menu was replaced before ${chosen.label}`;
  if (!clicked) throw new Error(`could not click ${chosen.label} in the command menu`);
  return `command ${chosen.label}`;
}

/** Longer than any animation before a change is made. Auto-save waits 30 seconds, so the end check saves itself. */
const SETTLE_MS = 2000;

/** The pause between random steps, on the page's clock. */
const STEP_PAUSE_MS = 50;

/**
 * `text` as a person could type it where the keyboard is. Playwright types a character that has no key on its
 * keyboard (é, 漢字, and a tab: Tab is a key, but `\t` isn't one of its characters) by inserting it as text,
 * which lands in the last text field, and takes the keyboard there, even when a button has it. No keyboard does
 * that, so outside a text field only the characters with keys are typed.
 */
async function typeable(page: Page, text: string): Promise<string> {
  const inText = await page.evaluate(() => {
    const active = document.activeElement;
    // A radio button or checkbox is an input too, but holds no text.
    const textInput = ['text', 'search', 'email', 'url', 'tel', 'password', 'number'];
    if (active instanceof HTMLInputElement) return textInput.includes(active.type);
    return (
      active instanceof HTMLTextAreaElement || (active instanceof HTMLElement && active.isContentEditable)
    );
  });
  return inText ? text : text.replace(/[^\x20-\x7e\n\r]/g, '');
}

/** Where the keyboard is, when a dialog is open and the keyboard is not in it; null otherwise. */
async function focusOutsideDialog(page: Page): Promise<string | null> {
  return page.evaluate(() => {
    const dialogs = [...document.querySelectorAll('[role=dialog], [role=alertdialog]')];
    const active = document.activeElement;
    if (dialogs.length === 0 || dialogs.some((dialog) => dialog.contains(active))) return null;
    const label = active?.getAttribute('aria-label') ?? '';
    return `${active?.tagName.toLowerCase()} ${label} while ${dialogs.map((d) => d.getAttribute('aria-label')).join(', ')} is open`;
  });
}

/** The cells compared before and after a reload: the fixture's data and a little around it. */
const CELLS = ['A', 'B', 'C', 'D', 'E'].flatMap((column) => [1, 2, 3, 4, 5].map((row) => `${column}${row}`));

/** What the window shows: the sheet tabs, and each cell's contents on the first sheet when it is a grid. */
async function shown(page: Page): Promise<{ tabs: string[]; cells: string[] }> {
  const tabs = await page.getByRole('tablist', { name: 'Sheets' }).getByRole('tab').allTextContents();
  await page.getByRole('tablist', { name: 'Sheets' }).getByRole('tab').first().click();
  const cells: string[] = [];
  if (await page.getByRole('grid', { name: 'Spreadsheet' }).isVisible())
    for (const cell of CELLS) {
      await goTo(page, cell);
      cells.push(`${cell} ${await contents(page).inputValue()}`);
    }
  return { tabs, cells };
}

/** After a run: Save writes every change, and opening the file again shows the same and writes nothing. */
async function checkSettled(page: Page): Promise<void> {
  // A dialog a command opens a moment later (a confirmation) can come up after the first Escape.
  const dialogs = page.locator('[role=dialog], [role=alertdialog]');
  for (let i = 0; i < 6; i++) {
    await page.clock.runFor(STEP_PAUSE_MS);
    if ((await dialogs.count()) === 0) break;
    await page.keyboard.press('Escape');
  }
  await expect(dialogs, 'Escape closes every dialog').toHaveCount(0);
  await page.clock.runFor(SETTLE_MS);
  // Auto-save waits 30 seconds, so save the way a user would.
  await runCommand(page, 'Save');
  await expect(page, 'Save leaves nothing unsaved').toHaveTitle(`${FILE} — Cellestra`);
  const before = await shown(page);
  const file = readFileSync(`${DATA}/${FILE}`);
  await openApp(page);
  await openFile(page, FILE);
  expect(await shown(page), 'the file has what the window showed').toEqual(before);
  await page.clock.runFor(SETTLE_MS);
  expect(readFileSync(`${DATA}/${FILE}`).equals(file), 'opening the file writes nothing').toBe(true);
}

for (const runSeed of seeds) {
  test(`random use, seed ${runSeed}`, async ({ page, seed }) => {
    seed('basic');
    // A step that hangs fails here, with the steps that led to it, instead of at the test's own timeout,
    // which would end the test before the seed and the steps are printed.
    page.setDefaultTimeout(5000);
    const errors = watchErrors(page);
    await openApp(page);
    await openFile(page, FILE);

    const random = createRandom(runSeed);
    const log: string[] = [];
    try {
      for (let step = 0; step < STEPS; step++) {
        const roll = random.int(10);
        if (roll < 4) {
          log.push(await runRandomCommand(page, (count) => random.int(count)));
        } else if (roll < 7) {
          const text = await typeable(page, random.text(8));
          log.push(`type ${JSON.stringify(text)}`);
          await page.keyboard.type(text);
        } else {
          const key = random.pick(KEYS);
          log.push(`key ${key}`);
          await page.keyboard.press(key);
        }
        // A person takes longer than this between keys. Without it, focus the app moves a moment after a command
        // (back to the grid, say) lands after the next key has opened a dialog, which no one could do.
        await page.clock.runFor(STEP_PAUSE_MS);
        expect(errors, 'console errors').toEqual([]);
        expect(await focusOutsideDialog(page), 'an open dialog has the keyboard').toBeNull();
      }
      log.push('settle');
      await checkSettled(page);
      expect(errors, 'console errors').toEqual([]);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      const steps = process.env.FUZZ_SEED ? log : log.slice(-20);
      throw new Error(`FUZZ_SEED=${runSeed} failed after:\n${steps.join('\n')}\n\n${message}`, {
        cause: error,
      });
    }
  });
}
