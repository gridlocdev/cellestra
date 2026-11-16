// Random use of the whole app: seeded runs of commands from the ⌘K menu, typing and keys on a fixture workbook,
// checking after every step that nothing broke. `just e2e` runs a few seeds; FUZZ_RUNS and FUZZ_SEED_BASE try more.
// A failure prints its seed: `FUZZ_SEED=<n> just e2e -g random` replays it, with every step. Keep a found bug
// as its own scenario.
import { readFileSync } from 'node:fs';
import type { Page } from '@playwright/test';
import { contents, DATA, expect, goTo, openApp, openFile, runCommand, test } from './fixtures.ts';
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
  // Another window that is open (a dialog, a menu) can keep ⌘K to itself.
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
  await options.nth(chosen.i).click();
  return `command ${chosen.label}`;
}

/** Longer than any animation before a change is made. Auto-save waits 30 seconds, so the end check saves itself. */
const SETTLE_MS = 2000;

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
  for (let i = 0; i < 4; i++) await page.keyboard.press('Escape');
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
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    page.on('console', (message) => {
      if (message.type() === 'error') errors.push(message.text());
    });
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
          const text = random.text(8);
          log.push(`type ${JSON.stringify(text)}`);
          await page.keyboard.type(text);
        } else {
          const key = random.pick(KEYS);
          log.push(`key ${key}`);
          await page.keyboard.press(key);
        }
        expect(errors, 'console errors').toEqual([]);
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
