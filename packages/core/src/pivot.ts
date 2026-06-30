/**
 * Pivot — pure data transformation, deliberately decoupled from the Grid.
 *
 * `pivot(data, columns, config)` turns flat rows into a pivot table:
 * the result is `{ columns, rows }` ready to feed straight into a new
 * Grid (or anywhere else). Being a pure function it is trivially
 * testable and imposes zero cost on grids that don't use it.
 *
 *   const p = pivot(grid ? data : rows, columnDefs, {
 *     rows: ['zona'],                 // one or more row dimensions
 *     cols: 'stato',                  // optional column dimension
 *     values: [{ colId: 'importo', aggFunc: 'sum', header: 'Importo' }],
 *     totals: true,                   // row + column totals (default)
 *   });
 *   new Grid(el, { columns: p.columns, data: p.rows, editable: false });
 */
import { getPath } from './infer';
import type { ColumnDef, RowData } from './types';

export type PivotAgg = 'sum' | 'avg' | 'min' | 'max' | 'count';

export interface PivotValueDef {
  colId: string;
  aggFunc: PivotAgg;
  header?: string;
  /** Formatter for the generated value columns. */
  valueFormatter?: (value: unknown, row: RowData) => string;
}

export interface PivotConfig {
  /** Row dimension(s), outermost first. */
  rows: string[];
  /** Optional column dimension. */
  cols?: string;
  values: PivotValueDef[];
  /** Add row totals column(s) and a grand-total row. @default true */
  totals?: boolean;
}

export interface PivotResult {
  columns: ColumnDef[];
  rows: RowData[];
}

const SEP = ' · ';
const TOTAL = 'Totale';

function agg(fn: PivotAgg, values: unknown[]): number | null {
  if (fn === 'count') return values.length;
  const nums = values.filter((v): v is number => typeof v === 'number' && Number.isFinite(v));
  if (!nums.length) return null;
  switch (fn) {
    case 'sum':
      return nums.reduce((s, v) => s + v, 0);
    case 'avg':
      return nums.reduce((s, v) => s + v, 0) / nums.length;
    case 'min':
      return Math.min(...nums);
    case 'max':
      return Math.max(...nums);
  }
}

export function pivot(data: RowData[], columns: ColumnDef[], config: PivotConfig): PivotResult {
  const totals = config.totals ?? true;
  const fieldOf = (colId: string): string => {
    const def = columns.find((c) => (c.id ?? c.field) === colId);
    return def?.field ?? colId;
  };
  const headerOf = (colId: string): string => {
    const def = columns.find((c) => (c.id ?? c.field) === colId);
    return def?.header ?? colId;
  };
  const rowFields = config.rows.map(fieldOf);
  const colField = config.cols ? fieldOf(config.cols) : null;
  const valueDefs = config.values.map((v) => ({
    ...v,
    field: fieldOf(v.colId),
    header: v.header ?? `${v.aggFunc} ${headerOf(v.colId)}`,
  }));

  // Bucket rows by composite row key; collect distinct column-dimension keys.
  const buckets = new Map<string, { keys: string[]; byCol: Map<string, RowData[]>; all: RowData[] }>();
  const colKeys = new Set<string>();
  for (const row of data) {
    if (row === undefined) continue;
    const keys = rowFields.map((f) => String(getPath(row, f) ?? ''));
    const k = keys.join('\u0000');
    let b = buckets.get(k);
    if (!b) buckets.set(k, (b = { keys, byCol: new Map(), all: [] }));
    b.all.push(row);
    if (colField) {
      const ck = String(getPath(row, colField) ?? '');
      colKeys.add(ck);
      let arr = b.byCol.get(ck);
      if (!arr) b.byCol.set(ck, (arr = []));
      arr.push(row);
    }
  }
  const sortedColKeys = [...colKeys].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

  const valueColId = (colKey: string | null, vd: (typeof valueDefs)[number]): string =>
    colKey === null ? vd.header : colKey + SEP + vd.header;

  // ---- Output columns
  const outColumns: ColumnDef[] = config.rows.map((colId, i) => ({
    field: '__r' + i,
    header: headerOf(colId),
    type: 'text',
    width: 160,
    editable: false,
    pinned: i === 0 ? 'left' : null,
  }));
  const dataCols: string[] = [];
  const pushValueCols = (colKey: string | null) => {
    for (const vd of valueDefs) {
      const id = valueColId(colKey, vd);
      dataCols.push(id);
      outColumns.push({
        field: id,
        header: id,
        type: 'number',
        width: 130,
        editable: false,
        aggFunc: vd.aggFunc === 'count' ? 'sum' : vd.aggFunc,
        valueFormatter: vd.valueFormatter,
      });
    }
  };
  if (colField) for (const ck of sortedColKeys) pushValueCols(ck);
  if (!colField || totals) pushValueCols(colField ? TOTAL : null);

  // ---- Output rows
  const outRows: RowData[] = [];
  const sortedBuckets = [...buckets.values()].sort((a, b) =>
    a.keys.join('\u0000').localeCompare(b.keys.join('\u0000'), undefined, { numeric: true }),
  );
  for (const b of sortedBuckets) {
    const row: RowData = {};
    b.keys.forEach((k, i) => (row['__r' + i] = k));
    if (colField) {
      for (const ck of sortedColKeys) {
        const subset = b.byCol.get(ck) ?? [];
        for (const vd of valueDefs)
          row[valueColId(ck, vd)] = subset.length ? agg(vd.aggFunc, subset.map((r) => getPath(r, vd.field))) : null;
      }
      if (totals)
        for (const vd of valueDefs)
          row[valueColId(TOTAL, vd)] = agg(vd.aggFunc, b.all.map((r) => getPath(r, vd.field)));
    } else {
      for (const vd of valueDefs)
        row[valueColId(null, vd)] = agg(vd.aggFunc, b.all.map((r) => getPath(r, vd.field)));
    }
    outRows.push(row);
  }

  // ---- Grand-total row
  if (totals && outRows.length) {
    const totalRow: RowData = { __r0: TOTAL };
    for (let i = 1; i < config.rows.length; i++) totalRow['__r' + i] = '';
    const allRows = data.filter((r) => r !== undefined);
    if (colField) {
      for (const ck of sortedColKeys) {
        const subset = allRows.filter((r) => String(getPath(r, colField) ?? '') === ck);
        for (const vd of valueDefs)
          totalRow[valueColId(ck, vd)] = subset.length ? agg(vd.aggFunc, subset.map((r) => getPath(r, vd.field))) : null;
      }
    }
    for (const vd of valueDefs)
      totalRow[valueColId(colField ? TOTAL : null, vd)] = agg(vd.aggFunc, allRows.map((r) => getPath(r, vd.field)));
    outRows.push(totalRow);
  }

  return { columns: outColumns, rows: outRows };
}
