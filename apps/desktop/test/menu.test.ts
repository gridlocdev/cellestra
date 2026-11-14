import type { MenuItemConstructorOptions } from 'electron';
import { describe, expect, it } from 'vitest';
import { withStandardMenus } from '../src/menu.ts';

const catalogMenus = (): MenuItemConstructorOptions[] => [
  { label: 'File', submenu: [{ label: 'New' }, { label: 'Close', accelerator: 'CmdOrCtrl+W' }] },
  { label: 'Edit', submenu: [{ role: 'copy' }] },
  { label: 'Help', submenu: [{ label: 'About Cellestra' }] },
];

const labels = (menus: readonly MenuItemConstructorOptions[]) => menus.map((m) => m.label ?? m.role);
const items = (menu: MenuItemConstructorOptions | undefined) =>
  (menu?.submenu as MenuItemConstructorOptions[]).map((i) => i.label ?? i.role ?? i.type);

describe('standard menus', () => {
  it('adds the app menu and the window menu on macOS, and leaves File alone', () => {
    const menus = withStandardMenus(catalogMenus(), 'darwin');
    expect(labels(menus)).toEqual(['appMenu', 'File', 'Edit', 'windowMenu', 'Help']);
    expect(items(menus[1])).toEqual(['New', 'Close']);
  });

  it('puts Quit at the end of File and a Window menu before Help on Windows and Linux', () => {
    for (const platform of ['win32', 'linux'] as const) {
      const menus = withStandardMenus(catalogMenus(), platform);
      expect(labels(menus)).toEqual(['File', 'Edit', 'Window', 'Help']);
      expect(items(menus[0])).toEqual(['New', 'Close', 'separator', 'quit']);
      // No Close in the Window menu: Ctrl+W belongs to File › Close.
      expect(items(menus[2])).toEqual(['minimize', 'Zoom']);
    }
  });

  it('does not change the menus it was given', () => {
    const given = catalogMenus();
    withStandardMenus(given, 'linux');
    expect(labels(given)).toEqual(['File', 'Edit', 'Help']);
    expect(items(given[0])).toEqual(['New', 'Close']);
  });
});
