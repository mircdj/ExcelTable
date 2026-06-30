/**
 * Value utilities: locale-aware parsing/formatting, comparators and the
 * Excel-style series detection used by the fill handle.
 */
import type { CellType, ColumnState, RowData } from './types';

export function formatValue(value: unknown, col: ColumnState, locale: string): string {
  if (value === null || value === undefined || value === '') return '';
  if (col.def.valueFormatter) return col.def.valueFormatter(value, {} as RowData);
  switch (col.type) {
    case 'number':
      return typeof value === 'number' && Number.isFinite(value)
        ? value.toLocaleString(locale, { maximumFractionDigits: 6 })
        : String(value);
    case 'date': {
      const d = value instanceof Date ? value : new Date(String(value));
      return isNaN(d.getTime()) ? String(value) : d.toLocaleDateString(locale);
    }
    case 'boolean':
      return value ? '✓' : '';
    default:
      return String(value);
  }
}

/** Parse user input according to column type and locale ("1.234,56" → 1234.56). */
export function parseValue(input: string, col: ColumnState, locale: string): unknown {
  const s = input.trim();
  if (s === '') return null;
  switch (col.type) {
    case 'number': {
      const sep = (1.1).toLocaleString(locale).charAt(1); // decimal separator
      const group = sep === ',' ? '.' : ',';
      const normalized = s
        .replace(new RegExp('\\' + group, 'g'), '')
        .replace(sep, '.')
        .replace('%', '');
      const n = Number(normalized);
      return Number.isFinite(n) ? (s.includes('%') ? n / 100 : n) : s;
    }
    case 'date': {
      // Accept dd/mm/yyyy (it-IT) and ISO.
      const m = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/.exec(s);
      if (m) {
        const year = m[3].length === 2 ? 2000 + Number(m[3]) : Number(m[3]);
        const d = new Date(year, Number(m[2]) - 1, Number(m[1]));
        if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
      }
      const iso = new Date(s);
      return isNaN(iso.getTime()) ? s : iso.toISOString().slice(0, 10);
    }
    case 'boolean':
      return /^(true|1|s[iì]|x|✓|yes|vero)$/i.test(s);
    default:
      return s;
  }
}

export function defaultComparator(type: CellType): (a: unknown, b: unknown) => number {
  return (a, b) => {
    const an = a === null || a === undefined || a === '';
    const bn = b === null || b === undefined || b === '';
    if (an && bn) return 0;
    if (an) return 1; // nulls last
    if (bn) return -1;
    if (type === 'number') return Number(a) - Number(b);
    if (type === 'date') return String(a) < String(b) ? -1 : String(a) > String(b) ? 1 : 0;
    // Natural sort so "CAB2" < "CAB10".
    return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' });
  };
}

/* ------------------------------------------------------------------ */
/* Fill-handle series detection (Excel behaviour)                      */
/* ------------------------------------------------------------------ */

const TEXT_NUM = /^(.*?)(\d+)$/;

/**
 * Given the source values of one column slice, return a generator that
 * produces the value for the i-th filled cell (i starts at 0).
 * Replicates Excel: linear numeric series, "Cabina 1" → "Cabina 2",
 * date series, otherwise pattern repetition.
 */
export function makeSeries(values: unknown[], type: CellType): (i: number) => unknown {
  const n = values.length;
  if (n === 0) return () => null;

  // Numeric linear series (needs ≥2 numbers, constant step; 1 number → step 1).
  if (type === 'number' && values.every((v) => typeof v === 'number' && Number.isFinite(v))) {
    const nums = values as number[];
    const step = n >= 2 ? nums[1] - nums[0] : 1;
    const linear = n < 2 || nums.every((v, i) => i === 0 || Math.abs(v - nums[i - 1] - step) < 1e-9);
    if (linear) return (i) => nums[n - 1] + step * (i + 1);
  }

  // Date series: constant day step.
  if (type === 'date') {
    const dates = values.map((v) => new Date(String(v)).getTime());
    if (dates.every((t) => !isNaN(t))) {
      const day = 86400000;
      const step = n >= 2 ? Math.round((dates[1] - dates[0]) / day) : 1;
      const linear =
        n < 2 || dates.every((t, i) => i === 0 || Math.round((t - dates[i - 1]) / day) === step);
      if (linear)
        return (i) => new Date(dates[n - 1] + step * day * (i + 1)).toISOString().slice(0, 10);
    }
  }

  // Text with numeric suffix: "Cabina 7" → "Cabina 8".
  if (n === 1 && typeof values[0] === 'string') {
    const m = TEXT_NUM.exec(values[0]);
    if (m) {
      const prefix = m[1];
      const num = Number(m[2]);
      const pad = m[2].length;
      return (i) => prefix + String(num + i + 1).padStart(pad, '0');
    }
  }

  // Fallback: repeat the source pattern.
  return (i) => values[i % n];
}
