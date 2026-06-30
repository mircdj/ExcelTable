/**
 * @exceltable/react — React wrapper around @exceltable/core.
 *
 * SSR-safe (Next.js): the Grid is only instantiated inside useEffect, so it
 * never runs on the server. In Next.js App Router mark the consuming
 * component with 'use client'.
 *
 * Usage:
 *   <ExcelTable columns={cols} data={rows} onCellValueChanged={...} apiRef={ref} />
 */
import {
  forwardRef,
  useEffect,
  useImperativeHandle,
  useRef,
  type CSSProperties,
} from 'react';
import {
  Grid,
  type CellChange,
  type CellRange,
  type ChangeSource,
  type ColumnDef,
  type FilterModelItem,
  type GridOptions,
  type RowData,
  type SortModelItem,
} from '@exceltable/core';

export interface ExcelTableProps<T extends RowData = RowData>
  extends Omit<GridOptions<T>, 'columns' | 'data'> {
  columns: ColumnDef<T>[];
  data: T[];
  style?: CSSProperties;
  className?: string;
  /** Receives the Grid instance once mounted. */
  onGridReady?: (api: Grid<T>) => void;
  onCellValueChanged?: (change: CellChange & { source: ChangeSource }) => void;
  onCellsChanged?: (changes: CellChange[], source: ChangeSource) => void;
  onSelectionChanged?: (ranges: CellRange[]) => void;
  onSortChanged?: (model: SortModelItem[]) => void;
  onFilterChanged?: (model: FilterModelItem[]) => void;
  onDirtyStateChanged?: (dirtyRowIds: string[]) => void;
}

export type ExcelTableHandle<T extends RowData = RowData> = { api: Grid<T> | null };

function ExcelTableInner<T extends RowData>(
  props: ExcelTableProps<T>,
  ref: React.Ref<ExcelTableHandle<T>>,
) {
  const containerRef = useRef<HTMLDivElement>(null);
  const gridRef = useRef<Grid<T> | null>(null);
  // Keep latest callbacks without re-creating the grid.
  const cb = useRef(props);
  cb.current = props;

  useImperativeHandle(ref, () => ({ get api() { return gridRef.current; } }), []);

  useEffect(() => {
    const { columns, data, style: _s, className: _c, ...options } = cb.current;
    const grid = new Grid<T>(containerRef.current!, { ...options, columns, data });
    gridRef.current = grid;

    grid.events.on('cellValueChanged', (e) => cb.current.onCellValueChanged?.(e));
    grid.events.on('cellsChanged', (e) => cb.current.onCellsChanged?.(e.changes, e.source));
    grid.events.on('selectionChanged', (e) => cb.current.onSelectionChanged?.(e.ranges));
    grid.events.on('sortChanged', (e) => cb.current.onSortChanged?.(e.sortModel));
    grid.events.on('filterChanged', (e) => cb.current.onFilterChanged?.(e.filterModel));
    grid.events.on('dirtyStateChanged', (e) => cb.current.onDirtyStateChanged?.(e.dirtyRowIds));
    cb.current.onGridReady?.(grid);

    return () => {
      grid.destroy();
      gridRef.current = null;
    };
    // Grid is created once; data/columns updates flow through the effects below.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    gridRef.current?.setRowData(props.data);
  }, [props.data]);

  useEffect(() => {
    if (!gridRef.current) return;
    gridRef.current.setColumns(props.columns);
    gridRef.current.refresh();
  }, [props.columns]);

  return (
    <div
      ref={containerRef}
      className={props.className}
      style={{ width: '100%', height: '100%', ...props.style }}
    />
  );
}

/** Excel-style data grid for React / Next.js. */
export const ExcelTable = forwardRef(ExcelTableInner) as <T extends RowData>(
  props: ExcelTableProps<T> & { ref?: React.Ref<ExcelTableHandle<T>> },
) => React.ReactElement;

/* ------------------------------------------------------------------ */
/* reactCell — mount a React component inside a grid cell              */
/* ------------------------------------------------------------------ */
import { createRoot } from 'react-dom/client';
import type { CellRenderer, CellRendererParams } from '@exceltable/core';

/**
 * Wrap a React component (or render function) as a grid cellRenderer.
 * The component is mounted with `createRoot` and unmounted automatically
 * when the row is recycled by the virtualizer — no leaks.
 *
 *   { field: 'stato', cellRenderer: reactCell((p) => <StatoBadge value={p.value} />) }
 *
 * Combine with strings for multi-line cells (value + component):
 *
 *   cellRenderer: multiLine((p) => [p.displayValue, reactNode(<Bar value={p.row.avanzamento} />)])
 */
export function reactCell<T extends RowData = RowData>(
  render: (params: CellRendererParams<T>) => React.ReactNode,
): CellRenderer<T> {
  return (params) => {
    const el = document.createElement('div');
    el.style.display = 'contents';
    const root = createRoot(el);
    root.render(render(params));
    return { el, destroy: () => queueMicrotask(() => root.unmount()) };
  };
}

/** A single React node as cell content (for use inside multi-line arrays). */
export function reactNode(node: React.ReactNode): { el: Node; destroy: () => void } {
  const el = document.createElement('div');
  el.style.display = 'contents';
  const root = createRoot(el);
  root.render(node);
  return { el, destroy: () => queueMicrotask(() => root.unmount()) };
}

export * from '@exceltable/core';
