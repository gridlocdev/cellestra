import { describe, expect, it } from 'vitest';
import { EditRefusedError, type EditCommand } from '@cellestra/core';
import { WorkbookSession } from '../src/session/workbookSession.ts';

const refusing: EditCommand = {
  label: 'Sort',
  apply() {
    throw new EditRefusedError('Cells that are merged can’t be sorted.');
  },
  undo() {
    throw new Error('never applied');
  },
};

describe('WorkbookSession.execute', () => {
  it('reports a refused edit and leaves the history alone', () => {
    const session = new WorkbookSession();
    const told: string[] = [];
    session.onRefused = (m) => told.push(m);
    session.execute(refusing);
    expect(told).toEqual(['Cells that are merged can’t be sorted.']);
    expect(session.canUndo.value).toBe(false);
    expect(session.modified.value).toBe(false);
  });

  it('still throws other errors', () => {
    const session = new WorkbookSession();
    const broken: EditCommand = {
      ...refusing,
      apply: () => {
        throw new Error('bug');
      },
    };
    expect(() => session.execute(broken)).toThrow('bug');
  });
});
