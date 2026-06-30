/**
 * JSON adapter — makes the grid fit *any* JSON payload.
 *
 *  - `getPath` / `setPath`: dot-path access into nested objects, so a column
 *    can bind to `"cliente.indirizzo.citta"`.
 *  - `inferColumns`: inspects the data and builds sensible ColumnDefs
 *    automatically (type detection, nested-object flattening, nested-array
 *    detection for master/detail), with include / exclude / override control.
 */
import type { CellType, ColumnDef, RowData } from './types';

export function getPath(obj: unknown, path: string): unknown {
  if (obj == null) return undefined;
  if (!path.includes('.')) return (obj as RowData)[path];
  let cur: unknown = obj;
  for (const key of path.split('.')) {
    if (cur == null || typeof cur !== 'object') return undefined;
    cur = (cur as RowData)[key];
  }
  return cur;
}

export function setPath(obj: RowData, path: string, value: unknown): void {
  if (!path.includes('.')) {
    obj[path] = value;
    return;
  }
  const keys = path.split('.');
  let cur: RowData = obj;
  for (let i = 0; i < keys.length - 1; i++) {
    const k = keys[i];
    if (cur[k] == null || typeof cur[k] !== 'object') cur[k] = {};
    cur = cur[k] as RowData;
  }
  cur[keys[keys.length - 1]] = value;
}

export interface InferOptions {
  /** Only these fields/paths (in this order). Wins over `exclude`. */
  include?: string[];
  /** Fields/paths to skip. */
  exclude?: string[];
  /** Per-field overrides merged onto the inferred definition. */
  overrides?: Record<string, Partial<ColumnDef>>;
  /** How deep to flatten nested objects into dot paths. @default 2 */
  maxDepth?: number;
  /** Rows sampled for type detection. @default 50 */
  sample?: number;
  /** Turn `snake_case` / `camelCase` keys into readable headers. @default true */
  prettyHeaders?: boolean;
}

export interface InferResult {
  columns: ColumnDef[];
  /** Fields that contain arrays of objects — candidates for master/detail. */
  detailFields: string[];
}

const ISO_DATE = /^\d{4}-\d{2}-\d{2}(T.*)?$/;

function detectType(values: unknown[]): CellType {
  const nonNull = values.filter((v) => v !== null && v !== undefined && v !== '');
  if (!nonNull.length) return 'text';
  if (nonNull.every((v) => typeof v === 'number')) return 'number';
  if (nonNull.every((v) => typeof v === 'boolean')) return 'boolean';
  if (nonNull.every((v) => typeof v === 'string' && ISO_DATE.test(v))) return 'date';
  return 'text';
}

function prettify(path: string): string {
  const last = path.split('.').pop()!;
  return last
    .replace(/_/g, ' ')
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .replace(/^./, (c) => c.toUpperCase());
}

/** Infer grid columns from arbitrary JSON rows. */
export function inferColumns(data: RowData[], opts: InferOptions = {}): InferResult {
  const { maxDepth = 2, sample = 50, prettyHeaders = true } = opts;
  const rows = data.slice(0, sample);
  const paths: string[] = [];
  const detailFields: string[] = [];
  const seen = new Set<string>();

  const walk = (obj: RowData, prefix: string, depth: number) => {
    for (const key of Object.keys(obj)) {
      const path = prefix ? `${prefix}.${key}` : key;
      if (seen.has(path)) continue;
      const v = obj[key];
      if (Array.isArray(v)) {
        if (v.length && typeof v[0] === 'object' && !seen.has(path)) {
          seen.add(path);
          detailFields.push(path); // nested table candidate
        }
        continue;
      }
      if (v !== null && typeof v === 'object') {
        if (depth < maxDepth) walk(v as RowData, path, depth + 1);
        continue;
      }
      seen.add(path);
      paths.push(path);
    }
  };
  for (const row of rows) if (row && typeof row === 'object') walk(row, '', 0);

  let selected = paths;
  if (opts.include) selected = opts.include.filter((p) => paths.includes(p) || true);
  else if (opts.exclude) {
    const ex = new Set(opts.exclude);
    selected = paths.filter((p) => !ex.has(p) && !opts.exclude!.some((e) => p.startsWith(e + '.')));
  }

  const columns: ColumnDef[] = selected.map((path) => {
    const values = rows.map((r) => getPath(r, path));
    const type = detectType(values);
    const inferred: ColumnDef = {
      field: path,
      header: prettyHeaders ? prettify(path) : path,
      type,
      width: type === 'number' ? 110 : type === 'date' ? 120 : type === 'boolean' ? 95 : 150,
    };
    return { ...inferred, ...opts.overrides?.[path] };
  });

  return { columns, detailFields };
}
