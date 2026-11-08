import type { ChartDefinition } from '@cellestra/charts';
import { cloneChart } from '@cellestra/charts';
import type { ColumnFilter, DataValidation, FilterSpec, SortSpec, Table } from '../data/types.ts';
import { Expr } from '../formulas/ast/expr.ts';
import { CellAddress } from '../model/cellAddress.ts';
import { CellRange } from '../model/cellRange.ts';
import type { NamedRange } from '../model/namedRange.ts';
import { SheetDimensions } from '../model/sheetDimensions.ts';
import type { Workbook } from '../model/workbook.ts';
import type { Worksheet } from '../model/worksheet.ts';

/**
 * A reference fix-up, the same one formulas get: rewrite `expr` as seen from `homeSheet` (the sheet
 * an unqualified reference means). A reference that no longer exists comes back as `refError`.
 */
export type ReferenceTransform = (expr: Expr, homeSheet: string) => Expr;

/** One sheet's reference-holding state outside its cells. */
interface SheetReferences {
  readonly validations: readonly DataValidation[];
  readonly activeFilter: CellRange | null;
  readonly activeFilterSpec: FilterSpec | null;
  readonly activeSort: SortSpec | null;
  readonly chart: ChartDefinition | null;
}

/** Everything outside cell formulas that names cells, for undo of a structural edit or a rename. */
export interface ReferenceSnapshot {
  readonly names: readonly NamedRange[];
  readonly sheets: ReadonlyMap<string, SheetReferences>;
}

const key = (name: string): string => name.toLowerCase();

export function captureReferences(workbook: Workbook): ReferenceSnapshot {
  const sheets = new Map<string, SheetReferences>();
  for (const s of workbook.worksheets) {
    sheets.set(key(s.name), {
      validations: [...s.validations],
      activeFilter: s.activeFilter,
      activeFilterSpec: s.activeFilterSpec,
      activeSort: s.activeSort,
      chart: s.chart ? cloneChart(s.chart) : null,
    });
  }
  return { names: [...workbook.namedRanges], sheets };
}

/** Put back what `captureReferences` saw; sheets are matched by (current) name. */
export function restoreReferences(workbook: Workbook, snapshot: ReferenceSnapshot): void {
  workbook.replaceNamedRanges(snapshot.names);
  for (const s of workbook.worksheets) {
    const saved = snapshot.sheets.get(key(s.name));
    if (!saved) continue;
    s.replaceValidations(saved.validations);
    s.activeFilter = saved.activeFilter;
    s.activeFilterSpec = saved.activeFilterSpec;
    s.activeSort = saved.activeSort;
    s.chart = saved.chart ? cloneChart(saved.chart) : null;
  }
}

/** A range through the transform; null when it was deleted. */
export function transformRange(
  range: CellRange,
  homeSheet: string,
  transform: ReferenceTransform,
): CellRange | null {
  const out = transform(Expr.range(range), homeSheet);
  return out.kind === 'range' ? out.range : null;
}

/** Where a sheet column lands (a whole-column probe, so row edits never lose it); null if deleted. */
export function transformColumn(
  column: number,
  homeSheet: string,
  transform: ReferenceTransform,
): number | null {
  const probe = new CellRange(
    new CellAddress(column, 0, false, true),
    new CellAddress(column, SheetDimensions.maxRows - 1, false, true),
  );
  return transformRange(probe, homeSheet, transform)?.left ?? null;
}

/** Filter criteria keyed by absolute column, re-keyed; columns that were deleted lose their criteria. */
export function transformFilterSpec(
  spec: FilterSpec,
  homeSheet: string,
  transform: ReferenceTransform,
): FilterSpec {
  const columns = new Map<number, ColumnFilter>();
  for (const [column, criteria] of spec.columns) {
    const mapped = transformColumn(column, homeSheet, transform);
    if (mapped !== null) columns.set(mapped, criteria);
  }
  return { ...spec, columns };
}

/** Sort levels by absolute column, re-pointed; levels on deleted columns are dropped. */
export function transformSortSpec(
  spec: SortSpec,
  homeSheet: string,
  transform: ReferenceTransform,
): SortSpec {
  const levels = [];
  for (const level of spec.levels) {
    const column = transformColumn(level.column, homeSheet, transform);
    if (column !== null) levels.push({ ...level, column });
  }
  return { ...spec, levels };
}

/** A table's filter and sort keys follow their columns (its range is shifted by the caller). */
export function transformTableKeys(t: Table, homeSheet: string, transform: ReferenceTransform): Table {
  return {
    ...t,
    filter: t.filter ? transformFilterSpec(t.filter, homeSheet, transform) : t.filter,
    sort: t.sort ? transformSortSpec(t.sort, homeSheet, transform) : t.sort,
  };
}

/** A chart's stored reference text through the transform; null when its cells were deleted. */
function transformChartReference(
  reference: string | null,
  homeSheet: string,
  transform: ReferenceTransform,
): string | null {
  if (reference === null || reference.trim() === '') return reference;
  const range = CellRange.tryParse(reference.trim());
  if (!range) return reference; // not a plain range: leave it as it was
  const mapped = transformRange(range, homeSheet, transform);
  return mapped ? mapped.toString() : null;
}

function transformChart(
  chart: ChartDefinition,
  homeSheet: string,
  transform: ReferenceTransform,
): ChartDefinition {
  const ref = (r: string | null): string | null => transformChartReference(r, homeSheet, transform);
  const series = [];
  for (const s of chart.series) {
    const values = ref(s.valuesReference);
    if (values === null) continue; // its data is gone: the series goes too
    series.push({
      ...s,
      valuesReference: values,
      nameReference: ref(s.nameReference),
      xValuesReference: ref(s.xValuesReference),
      bubbleSizeReference: ref(s.bubbleSizeReference),
    });
  }
  return { ...chart, categoriesReference: ref(chart.categoriesReference), series };
}

/**
 * Apply a reference fix-up to everything outside cell formulas that names cells: named-range targets,
 * data validations (target and list source), each sheet's AutoFilter range and criteria columns, its
 * sort keys, and chart ranges. A validation whose target was deleted goes, and so does a name whose
 * target was deleted (as a chart series whose values were).
 */
export function transformReferences(
  workbook: Workbook,
  transform: ReferenceTransform,
  chartHome: string | undefined = defaultChartHome(workbook),
): void {
  const names: NamedRange[] = [];
  for (const n of workbook.namedRanges) {
    // A name's unqualified target means its scope's sheet; a global one has no home to resolve.
    const home = n.scope ?? n.target.sheet;
    if (home === null || home === undefined) {
      names.push(n);
      continue;
    }
    const target = transformRange(n.target, home, transform);
    if (target) names.push({ ...n, target });
  }
  workbook.replaceNamedRanges(names);

  for (const s of workbook.worksheets) transformSheetReferences(s, transform, chartHome ?? s.name);
}

/** A chart's unqualified reference means the first data sheet (ChartDataResolver's rule). */
export function defaultChartHome(workbook: Workbook): string | undefined {
  return workbook.worksheets.find((s) => !s.isChartSheet)?.name ?? workbook.worksheets[0]?.name;
}

function transformSheetReferences(s: Worksheet, transform: ReferenceTransform, chartHome: string): void {
  const home = s.name;
  const validations: DataValidation[] = [];
  for (const v of s.validations) {
    const target = transformRange(v.target, home, transform);
    if (!target) continue;
    const listSource = v.listSource ? transformRange(v.listSource, home, transform) : v.listSource;
    validations.push({ ...v, target, listSource });
  }
  s.replaceValidations(validations);

  if (s.activeFilter) {
    const filter = transformRange(s.activeFilter, home, transform);
    s.activeFilter = filter;
    s.activeFilterSpec =
      filter && s.activeFilterSpec ? transformFilterSpec(s.activeFilterSpec, home, transform) : null;
  } else if (s.activeFilterSpec) {
    s.activeFilterSpec = transformFilterSpec(s.activeFilterSpec, home, transform);
  }
  if (s.activeSort) s.activeSort = transformSortSpec(s.activeSort, home, transform);
  if (s.chart) s.chart = transformChart(s.chart, chartHome, transform);
}
