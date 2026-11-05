import type { Workbook } from '../model/workbook.ts';
import { EditResult, type EditCommand } from './editCommand.ts';

/** Undo/redo stacks over one workbook (PRD: at least 1000 steps). */
export class EditHistory {
  static readonly maxDepth = 1000;
  private readonly undoStack: EditCommand[] = [];
  private readonly redoStack: EditCommand[] = [];
  private readonly workbook: Workbook;

  constructor(workbook: Workbook) {
    this.workbook = workbook;
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }
  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }
  get undoLabel(): string | null {
    return this.undoStack.at(-1)?.label ?? null;
  }
  get redoLabel(): string | null {
    return this.redoStack.at(-1)?.label ?? null;
  }

  /** Apply a command and push it; clears redo, and drops the oldest step past the cap. */
  execute(command: EditCommand): EditResult {
    const result = command.apply(this.workbook);
    this.undoStack.push(command);
    this.redoStack.length = 0;
    if (this.undoStack.length > EditHistory.maxDepth) this.undoStack.shift();
    return result;
  }

  undo(): EditResult {
    const command = this.undoStack.pop();
    if (!command) return EditResult.none;
    const result = command.undo(this.workbook);
    this.redoStack.push(command);
    return result;
  }

  redo(): EditResult {
    const command = this.redoStack.pop();
    if (!command) return EditResult.none;
    const result = command.apply(this.workbook);
    this.undoStack.push(command);
    return result;
  }

  clear(): void {
    this.undoStack.length = 0;
    this.redoStack.length = 0;
  }
}
