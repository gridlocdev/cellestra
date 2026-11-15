import { describe, expect, it } from 'vitest';
import { launchArgs } from '../src/infra/launcher.ts';

describe('launchArgs', () => {
  it('names the installed app by bundle ID, not by name, so the dev copy of Electron is never picked', () => {
    expect(launchArgs(null)).toEqual(['-b', 'dev.gridloc.cellestra']);
  });

  it('passes the document after the bundle ID', () => {
    expect(launchArgs('/tmp/a b.ods')).toEqual(['-b', 'dev.gridloc.cellestra', '/tmp/a b.ods']);
  });
});
