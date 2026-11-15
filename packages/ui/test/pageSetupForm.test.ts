import { PageSetups, type PageSetup } from '@cellestra/core';
import { describe, expect, it } from 'vitest';
import { draftOf, setupOf } from '../src/components/pageSetupForm.ts';

describe('page setup form', () => {
  it('round-trips a setup through the draft', () => {
    const setup: PageSetup = {
      ...PageSetups.orient(PageSetups.default, 'landscape'),
      scale: { kind: 'fit', width: 1, height: 0 },
      gridlines: true,
    };
    const back = setupOf(draftOf(setup), setup);
    expect('setup' in back && PageSetups.equal(back.setup, setup)).toBe(true);
  });

  it('turns the paper with the orientation', () => {
    const r = setupOf({ ...draftOf(null), paper: 'letter', orientation: 'landscape' }, null);
    expect('setup' in r && [r.setup.paperWidth, r.setup.paperHeight]).toEqual([279.4, 215.9]);
  });

  it('rejects bad values', () => {
    const d = draftOf(null);
    expect(setupOf({ ...d, top: 'x' }, null)).toHaveProperty('error');
    expect(setupOf({ ...d, left: '100', right: '100' }, null)).toHaveProperty('error');
    expect(setupOf({ ...d, percent: '5' }, null)).toHaveProperty('error');
    expect(setupOf({ ...d, scaleKind: 'fit', fitWidth: '', fitHeight: '' }, null)).toHaveProperty('error');
    expect(setupOf({ ...d, scaleKind: 'pages', pages: '0' }, null)).toHaveProperty('error');
  });
});
