import { describe, expect, it } from 'vitest';
import { ActionCommand, EditResult } from '../../src/editing/editCommand.ts';
import { EditHistory } from '../../src/editing/editHistory.ts';
import { Workbook } from '../../src/model/workbook.ts';

/** A command whose undo or redo throws while `failing` is set. */
function flaky(): { command: ActionCommand; failing: { undo: boolean; apply: boolean } } {
  const failing = { undo: false, apply: false };
  const command = new ActionCommand(
    'Flaky',
    () => {
      if (failing.apply) throw new Error('apply failed');
      return EditResult.none;
    },
    () => {
      if (failing.undo) throw new Error('undo failed');
      return EditResult.none;
    },
  );
  return { command, failing };
}

describe('EditHistory', () => {
  it('keeps a step on the undo stack when its undo throws', () => {
    const history = new EditHistory(Workbook.createNew());
    const { command, failing } = flaky();
    history.execute(command);

    failing.undo = true;
    expect(() => history.undo()).toThrow('undo failed');
    expect(history.canUndo).toBe(true);
    expect(history.canRedo).toBe(false);

    failing.undo = false;
    history.undo();
    expect(history.canUndo).toBe(false);
    expect(history.canRedo).toBe(true);
  });

  it('keeps a step on the redo stack when its redo throws', () => {
    const history = new EditHistory(Workbook.createNew());
    const { command, failing } = flaky();
    history.execute(command);
    history.undo();

    failing.apply = true;
    expect(() => history.redo()).toThrow('apply failed');
    expect(history.canRedo).toBe(true);
    expect(history.canUndo).toBe(false);

    failing.apply = false;
    history.redo();
    expect(history.canUndo).toBe(true);
    expect(history.canRedo).toBe(false);
  });

  it('names the step Undo would undo next, so a notice can tell its edit is still the last', () => {
    const history = new EditHistory(Workbook.createNew());
    const first = flaky().command;
    const second = flaky().command;
    expect(history.lastCommand).toBeNull();
    history.execute(first);
    history.execute(second);
    expect(history.lastCommand).toBe(second);
    history.undo();
    expect(history.lastCommand).toBe(first);
  });
});
