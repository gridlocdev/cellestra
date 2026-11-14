import type { MenuItemConstructorOptions } from 'electron';

/**
 * The platform's standard menus around the ones the renderer builds from its command catalog:
 * the app menu (macOS, which holds Quit), a Window menu before Help, and on Windows and Linux a
 * Quit item at the end of the File menu. Pure, so it can be tested without Electron.
 */
export function withStandardMenus(
  template: readonly MenuItemConstructorOptions[],
  platform: NodeJS.Platform,
): MenuItemConstructorOptions[] {
  const mac = platform === 'darwin';
  const menus = template.map((menu) =>
    !mac && menu.label === 'File' && Array.isArray(menu.submenu)
      ? { ...menu, submenu: [...menu.submenu, { type: 'separator' as const }, { role: 'quit' as const }] }
      : menu,
  );
  const help = menus.findIndex((m) => m.label === 'Help');
  menus.splice(help < 0 ? menus.length : help, 0, windowMenu(mac));
  if (mac) menus.unshift({ role: 'appMenu' });
  return menus;
}

/**
 * Minimize and Zoom, plus Bring All to Front on macOS. There Electron's `windowMenu` role is
 * exactly that, and makes it the menu macOS lists open windows in. It is not used elsewhere: on
 * Windows and Linux it adds a Close item bound to Ctrl+W, which File › Close already owns.
 */
function windowMenu(mac: boolean): MenuItemConstructorOptions {
  if (mac) return { role: 'windowMenu' };
  return {
    label: 'Window',
    submenu: [{ role: 'minimize' }, { label: 'Zoom', click: (_item, win) => toggleMaximized(win) }],
  };
}

/** What Zoom does outside macOS: maximize the window, or put a maximized one back. */
function toggleMaximized(win: Electron.BaseWindow | undefined): void {
  if (!win) return;
  if (win.isMaximized()) win.unmaximize();
  else win.maximize();
}
