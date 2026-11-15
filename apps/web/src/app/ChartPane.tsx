import { ChartEditCommands } from '@cellestra/core';
import { ChartSheetView } from '@cellestra/ui';
import type { AppController } from './controller.ts';

/** A chart sheet's page: the chart plus its data and look controls; edits are undoable. */
export function ChartPane({ app }: { readonly app: AppController }) {
  const s = app.session;
  return (
    <ChartSheetView
      workbook={s.workbook.value}
      sheet={s.activeSheet.value}
      version={s.version.value}
      dark={app.isDark.value}
      onChange={(next) => app.run(ChartEditCommands.setChart(app.sheetName, next))}
      onPick={(field) => app.startChartPick(field)}
    />
  );
}
