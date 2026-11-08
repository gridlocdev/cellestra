import { EditSession } from '../infra/editSession.ts';
import { writeReport } from '../infra/report.ts';
import { load } from '../infra/workspace.ts';
import type { CommandSpec, OptionSpec } from '../spec.ts';
import { dispatch } from './apply.ts';
import { addChart, chartSheet, chartTypeName, describeChart } from './chartOps.ts';
import { editTarget, fileArg, inPlaceOpt, jsonOpt, outputOpt } from './options.ts';
import { onOff } from './layout.ts';

/**
 * Chart sheets from the command line — the headless face of the desktop's Insert ▸ Chart Sheet. The
 * result round-trips as an ODF embedded chart object, same as a GUI-made one.
 */

const TYPES = 'column|bar|line|area|pie|doughnut|scatter|bubble|radar|filled-radar';

const greedy = (name: string, description: string): OptionSpec => ({
  name,
  description,
  kind: 'multi',
  greedy: true,
});

export const chart: CommandSpec = {
  name: 'chart',
  description: 'Add a chart sheet plotting data from the workbook',
  args: [fileArg()],
  options: [
    { name: '--type', description: TYPES, default: 'column' },
    { name: '--title', description: 'Chart title' },
    {
      name: '--from',
      description:
        'Seed from a tabular range: 1st column = categories (or X for scatter/bubble), the rest = series; a text header row names them',
    },
    { name: '--categories', description: 'Category-label range (category & pie families)' },
    greedy('--values', "A series' value range; repeat for multiple series"),
    greedy('--names', 'Literal series names, paired with --values in order'),
    { name: '--x-values', description: 'Shared X-value range (scatter & bubble)' },
    { name: '--bubble-sizes', description: 'Shared bubble-size range (bubble only)' },
    { name: '--stacking', description: 'none|stacked|percent (column/bar/area only)' },
    { name: '--no-legend', description: 'Hide the legend', kind: 'flag' },
    { name: '--x-title', description: 'Category/X axis title' },
    { name: '--y-title', description: 'Value/Y axis title' },
    inPlaceOpt,
    outputOpt,
    jsonOpt,
  ],
  run(p, con) {
    const path = p.arg('file');
    const target = editTarget(path, p.flag('--in-place'), p.str('--output'));
    return EditSession.run(con, path, target, null, p.flag('--json'), (session) =>
      addChart(
        session,
        p.str('--type')!,
        {
          title: p.str('--title'),
          xTitle: p.str('--x-title'),
          yTitle: p.str('--y-title'),
          stacking: p.str('--stacking'),
        },
        p.flag('--no-legend'),
        {
          values: p.multi('--values'),
          names: p.multi('--names'),
          categories: p.str('--categories'),
          xValues: p.str('--x-values'),
          bubbleSizes: p.str('--bubble-sizes'),
          from: p.str('--from'),
        },
      ),
    );
  },
};

export const charts: CommandSpec = {
  name: 'charts',
  description: "List, show, update, or remove the workbook's chart sheets",
  subcommands: [
    {
      name: 'list',
      description: 'List chart sheets: name, type, title, series',
      args: [fileArg()],
      options: [jsonOpt],
      run(p, con) {
        const loaded = load(con, p.arg('file'), 'ods', null);
        const sheets = loaded.workbook.worksheets.filter((s) => s.isChartSheet);
        if (p.flag('--json')) {
          writeReport(con, sheets.map(describeChart));
          return 0;
        }
        for (const sheet of sheets) {
          const c = sheet.chart!;
          con.outLine(`${sheet.name}\t${chartTypeName(c.type)}\t${c.title ?? ''}\t${c.series.length} series`);
        }
        return 0;
      },
    },
    {
      name: 'show',
      description: "Print a chart's definition as JSON",
      args: [fileArg(), { name: 'sheet', description: 'The chart sheet' }],
      run(p, con) {
        const loaded = load(con, p.arg('file'), 'ods', null);
        writeReport(con, describeChart(chartSheet(loaded.workbook, p.arg('sheet'))));
        return 0;
      },
    },
    {
      name: 'remove',
      description: 'Remove a chart sheet',
      args: [fileArg(), { name: 'sheet', description: 'The chart sheet to remove' }],
      options: [inPlaceOpt, outputOpt, jsonOpt],
      run: (p, con) => dispatch(p, con, 'chart-remove', { sheet: p.arg('sheet') }),
    },
    {
      name: 'update',
      description: 'Change a chart: any option given replaces that part; data options rebuild the series',
      args: [fileArg(), { name: 'sheet', description: 'The chart sheet to change' }],
      options: [
        { name: '--type', description: TYPES },
        { name: '--title', description: 'Chart title' },
        { name: '--from', description: 'Rebuild the series from a tabular range' },
        { name: '--categories', description: 'Category-label range' },
        greedy('--values', 'Rebuild the series: a value range per series'),
        greedy('--names', 'Series names, paired with --values'),
        { name: '--x-values', description: 'Shared X-value range (scatter & bubble)' },
        { name: '--bubble-sizes', description: 'Shared bubble-size range (bubble only)' },
        { name: '--stacking', description: 'none|stacked|percent' },
        { ...onOff('--legend'), description: 'true | false' },
        { name: '--x-title', description: 'Category/X axis title' },
        { name: '--y-title', description: 'Value/Y axis title' },
        inPlaceOpt,
        outputOpt,
        jsonOpt,
      ],
      run: (p, con) => {
        const values = p.multi('--values');
        const names = p.multi('--names');
        return dispatch(p, con, 'chart-update', {
          sheet: p.arg('sheet'),
          type: p.str('--type'),
          title: p.str('--title'),
          from: p.str('--from'),
          categories: p.str('--categories'),
          values: values.length > 0 ? values : null,
          names: names.length > 0 ? names : null,
          xValues: p.str('--x-values'),
          bubbleSizes: p.str('--bubble-sizes'),
          stacking: p.str('--stacking'),
          legend: p.onOff('--legend'),
          xTitle: p.str('--x-title'),
          yTitle: p.str('--y-title'),
        });
      },
    },
  ],
};
