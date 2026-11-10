import type { Workbook } from '../model/workbook.ts';
import type { PageSetup } from '../model/pageSetup.ts';
import { EditResult, requireSheet, type EditCommand } from './editCommand.ts';

/** Change how a sheet prints. No cell changes, so nothing recalculates. */
export class SetPageSetupCommand implements EditCommand {
  readonly label = 'Page Setup';
  private readonly sheetName: string;
  private readonly setup: PageSetup | null;
  private before: PageSetup | null = null;

  constructor(sheetName: string, setup: PageSetup | null) {
    this.sheetName = sheetName;
    this.setup = setup;
  }

  apply(workbook: Workbook): EditResult {
    const sheet = requireSheet(workbook, this.sheetName);
    this.before = sheet.pageSetup;
    sheet.pageSetup = this.setup;
    return EditResult.none;
  }

  undo(workbook: Workbook): EditResult {
    requireSheet(workbook, this.sheetName).pageSetup = this.before;
    return EditResult.none;
  }
}
