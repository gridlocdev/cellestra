import { spawn } from 'node:child_process';
import { CliError } from './errors.ts';

/** The desktop app's bundle ID (`appId` in apps/desktop/electron-builder.yml). */
const BUNDLE_ID = 'dev.gridloc.cellestra';

/**
 * Arguments for macOS `open`. The app is named by bundle ID: `open -a Cellestra` matches by name,
 * and can pick the development copy of Electron (also called Cellestra), which shows Electron's
 * default page when started without the app's path.
 */
export function launchArgs(path: string | null): string[] {
  return ['-b', BUNDLE_ID, ...(path === null ? [] : [path])];
}

/**
 * Open the Cellestra desktop app (with a document, or none). `open -b` reaches the installed
 * bundle, running or not; the document travels as an Apple Event, which the app handles.
 */
export function launchApp(path: string | null): number {
  if (process.platform !== 'darwin') {
    throw new CliError(
      'Opening a workbook in the app from the command line is only supported on macOS; open it from the desktop app.',
    );
  }
  const args = launchArgs(path);
  try {
    const child = spawn('open', args, { stdio: 'ignore', detached: true });
    child.on('error', () => undefined);
    child.unref();
  } catch (e) {
    throw new CliError(`Couldn't launch Cellestra: ${(e as Error).message}`);
  }
  return 0;
}
