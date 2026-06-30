/**
 * Grid — the central orchestrator and public API.
 *
 * Owns: data, column state, sort/filter view pipeline, selection model,
 * transaction log (undo/redo), dirty tracking. Rendering and input are
 * delegated to Renderer and InteractionManager.
 */
import { Emitter } from './events';
import { Renderer } from './renderer';
import { InteractionManager } from './interaction';
import { defaultComparator, formatValue, parseValue } from './values';
import { getPath, setPath, inferColumns } from './infer';
import { FormulaEngine, FormulaError, isFormula } from './formula';
import { IT, interpolate, type Strings } from './i18n';
import { ServerSource } from './serverside';
import type {
  CellChange,
  GroupRow,
  GridState,
  CellRange,
  CellRef,
  ChangeSource,
  ColumnDef,
  ColumnState,
  FilterModelItem,
  GridOptions,
  RowData,
  RowId,
  SortModelItem,
} from './types';

type UndoEntry<T extends RowData> =
  | { kind: 'cells'; changes: CellChange[] }
  | { kind: 'rows'; op: 'insert' | 'remove'; rows: T[]; ids: RowId[]; indices: number[] };

const DEFAULTS = {
  rowHeight: 32,
  headerHeight: 36,
  width: 140,
  minWidth: 48,
  maxWidth: 2000,
};

export class Grid<T extends RowData = RowData> {
  /**
   * Build a grid directly from *any* JSON payload: columns are inferred
   * (types, nested-object flattening), nested arrays of objects become
   * expandable sub-tables automatically.
   *
   *   Grid.fromJson(el, await res.json(), { exclude: ['_meta'] });
   */
  static fromJson(
    container: HTMLElement,
    json: unknown,
    options: Partial<GridOptions> & import('./infer').InferOptions = {},
  ): Grid {
    const data = (Array.isArray(json) ? json : [json]) as RowData[];
    const { include, exclude, overrides, maxDepth, sample, prettyHeaders, ...gridOpts } = options;
    const inferred = inferColumns(data, { include, exclude, overrides, maxDepth, sample, prettyHeaders });
    return new Grid(container, {
      columns: inferred.columns,
      data,
      masterDetail:
        gridOpts.masterDetail ??
        (inferred.detailFields.length ? { field: inferred.detailFields[0] } : undefined),
      ...gridOpts,
    });
  }

  readonly events = new Emitter();
  readonly options: Required<Pick<GridOptions<T>, 'rowHeight' | 'headerHeight' | 'locale' | 'undoLimit'>> &
    GridOptions<T>;

  /** All rows, master order. */
  private data: T[] = [];
  private rowIds: RowId[] = [];
  /** view[i] = index into `data` for visual row i (after sort+filter). */
  view: number[] = [];

  columns: ColumnState<T>[] = [];
  sortModel: SortModelItem[] = [];
  filterModel: FilterModelItem[] = [];
  groupBy: string[] = [];
  /** Group header rows of the current view (referenced by negative view entries). */
  groupRows: GroupRow[] = [];
  private collapsedGroups = new Set<string>();

  /** Selection state (view coordinates). */
  ranges: CellRange[] = [];
  activeCell: CellRef | null = null;

  private undoStack: UndoEntry<T>[] = [];
  private redoStack: UndoEntry<T>[] = [];
  private dirty = new Set<RowId>();
  private idSeq = 0;

  private renderer: Renderer<T>;
  private interaction: InteractionManager<T>;
  private destroyed = false;
  /** Formula engine, present when `formulas: true`. */
  readonly formulaEngine: FormulaEngine | null = null;
  /** Server-side source, present when `serverSide` is configured. */
  readonly serverSource: ServerSource<T> | null = null;
  /** Resolved UI strings (IT default + overrides). */
  readonly strings: Strings;
  /** Tree mode: flattened nodes parallel to `data`. @internal */
  treeMeta: { level: number; hasChildren: boolean; parent: number }[] = [];

  constructor(container: HTMLElement, options: GridOptions<T>) {
    this.options = {
      rowHeight: DEFAULTS.rowHeight,
      headerHeight: DEFAULTS.headerHeight,
      locale: 'it-IT',
      undoLimit: 200,
      showRowNumbers: true,
      editable: true,
      theme: 'excel',
      ...options,
    };
    this.strings = { ...IT, ...options.strings };
    this.setColumns(options.columns);
    this.groupBy = options.treeData ? [] : (options.groupBy ?? []);
    if (options.formulas) {
      this.formulaEngine = new FormulaEngine((col, row) => {
        const c = this.columns[col];
        const r = this.data[row];
        return c && r !== undefined ? getPath(r, c.field) : undefined;
      });
      this.events.on('cellsChanged', () => this.formulaEngine!.invalidate());
      this.events.on('rowDataChanged', () => this.formulaEngine!.invalidate());
      this.events.on('rowsInserted', () => this.formulaEngine!.invalidate());
      this.events.on('rowsRemoved', () => this.formulaEngine!.invalidate());
    }
    if (options.serverSide) {
      this.serverSource = new ServerSource<T>(this, options.serverSide);
      this.events.on('viewportChanged', (e) => this.serverSource!.ensureRange(e.firstRow, e.lastRow));
    } else if (options.data) this.setRowData(options.data, false);

    this.renderer = new Renderer(container, this);
    this.interaction = new InteractionManager(this, this.renderer);
    this.renderer.mount();

    if (this.serverSource) this.serverSource.ensureRange(0, 0); // primo blocco; il viewport richiede il resto
    this.events.emit('gridReady', {});
    requestAnimationFrame(() => {
      if (!this.destroyed) this.events.emit('firstDataRendered', {});
    });
  }

  /** Translate a UI string with optional {param} interpolation. */
  t(key: keyof Strings, params?: Record<string, string | number>): string {
    return interpolate(this.strings[key], params);
  }

  /* ---------------------------------------------------------------- */
  /* Columns                                                           */
  /* ---------------------------------------------------------------- */

  setColumns(defs: ColumnDef<T>[]): void {
    this.columns = defs.map((def) => ({
      def,
      id: def.id ?? def.field,
      field: def.field,
      header: def.header ?? def.field,
      type: def.type ?? 'text',
      width: def.width ?? DEFAULTS.width,
      minWidth: def.minWidth ?? DEFAULTS.minWidth,
      maxWidth: def.maxWidth ?? DEFAULTS.maxWidth,
      pinned: def.pinned ?? null,
      editable: def.editable ?? true,
      sortable: def.sortable ?? true,
      resizable: def.resizable ?? true,
      visible: true,
      left: 0,
    }));
    this.layoutColumns();
  }

  /** Recompute x offsets; pinned columns first, then scrollable. */
  layoutColumns(): void {
    let x = 0;
    for (const c of this.pinnedColumns()) {
      c.left = x;
      x += c.width;
    }
    x = 0;
    for (const c of this.centerColumns()) {
      c.left = x;
      x += c.width;
    }
  }

  visibleColumns(): ColumnState<T>[] {
    return [...this.pinnedColumns(), ...this.centerColumns()];
  }
  pinnedColumns(): ColumnState<T>[] {
    return this.columns.filter((c) => c.visible && c.pinned === 'left');
  }
  centerColumns(): ColumnState<T>[] {
    return this.columns.filter((c) => c.visible && !c.pinned);
  }
  pinnedWidth(): number {
    return this.pinnedColumns().reduce((s, c) => s + c.width, 0);
  }
  centerWidth(): number {
    return this.centerColumns().reduce((s, c) => s + c.width, 0);
  }
  columnByVisibleIndex(i: number): ColumnState<T> | undefined {
    return this.visibleColumns()[i];
  }
  visibleIndexOf(colId: string): number {
    return this.visibleColumns().findIndex((c) => c.id === colId);
  }

  setColumnWidth(colId: string, width: number): void {
    const col = this.columns.find((c) => c.id === colId);
    if (!col) return;
    col.width = Math.max(col.minWidth, Math.min(col.maxWidth, Math.round(width)));
    this.layoutColumns();
    this.renderer.invalidateAll();
    this.events.emit('columnResized', { colId, width: col.width });
  }

  setColumnVisible(colId: string, visible: boolean): void {
    const col = this.columns.find((c) => c.id === colId);
    if (!col || col.visible === visible) return;
    col.visible = visible;
    this.layoutColumns();
    this.renderer.invalidateAll();
    this.events.emit('columnVisibilityChanged', { colId, visible });
  }

  autoSizeColumn(colId: string): void {
    const col = this.columns.find((c) => c.id === colId);
    if (!col) return;
    const probe = document.createElement('span');
    probe.style.cssText =
      'position:absolute;visibility:hidden;white-space:nowrap;font:var(--eg-font, 13px sans-serif);';
    document.body.appendChild(probe);
    let max = 0;
    probe.textContent = col.header;
    max = probe.offsetWidth + 28;
    const sample = Math.min(this.view.length, 500);
    for (let i = 0; i < sample; i++) {
      probe.textContent = this.getDisplayValue(i, col.id);
      if (probe.offsetWidth + 18 > max) max = probe.offsetWidth + 18;
    }
    probe.remove();
    this.setColumnWidth(colId, max);
  }

  /* ---------------------------------------------------------------- */
  /* Data                                                              */
  /* ---------------------------------------------------------------- */

  setRowData(rows: T[], render = true): void {
    if (this.options.treeData) {
      const td = this.options.treeData;
      const flat: T[] = [];
      const meta: { level: number; hasChildren: boolean; parent: number }[] = [];
      const walk = (nodes: T[], level: number, parent: number) => {
        for (const n of nodes) {
          const children = td.getChildren
            ? td.getChildren(n)
            : ((n as RowData)[td.childrenField ?? 'children'] as T[] | undefined);
          const idx = flat.length;
          flat.push(n);
          meta.push({ level, hasChildren: !!children?.length, parent });
          if (children?.length) walk(children, level + 1, idx);
        }
      };
      walk(rows, 0, -1);
      this.data = flat;
      this.treeMeta = meta;
      this.rowIds = flat.map((r) => this.resolveRowId(r));
      // Radici espanse di default
      for (let i = 0; i < flat.length; i++)
        if (meta[i].level === 0 && meta[i].hasChildren) this.expandedIds.add(this.rowIds[i]);
    } else {
      this.data = rows.slice();
      this.treeMeta = [];
      this.rowIds = rows.map((r) => this.resolveRowId(r));
    }
    this.dirty.clear();
    this.undoStack = [];
    this.redoStack = [];
    this.rebuildView();
    if (render) this.renderer?.invalidateAll();
    this.events.emit('rowDataChanged', { rowCount: rows.length });
    this.emitUndoState();
  }

  private resolveRowId(row: T): RowId {
    return this.options.getRowId ? this.options.getRowId(row) : `eg-${++this.idSeq}`;
  }

  get serverMode(): boolean {
    return this.serverSource !== null;
  }

  /** True while a server-side row has not arrived yet. */
  isRowLoading(viewIndex: number): boolean {
    return this.serverMode && !this.isGroupRow(viewIndex) && this.getRowByViewIndex(viewIndex) === undefined;
  }

  /** @internal Store an arrived server block and resize the dataset. */
  __serverStore(startRow: number, rows: T[], totalCount: number): void {
    const totalChanged = totalCount !== this.data.length;
    if (totalChanged) {
      this.data.length = totalCount;
      this.rowIds.length = totalCount;
    }
    for (let k = 0; k < rows.length; k++) {
      this.data[startRow + k] = rows[k];
      this.rowIds[startRow + k] = this.resolveRowId(rows[k]);
    }
    if (totalChanged) {
      this.rebuildView();
      this.renderer.invalidateAll();
      this.events.emit('rowDataChanged', { rowCount: totalCount });
    } else {
      this.renderer.invalidateRows(Array.from({ length: rows.length }, (_, k) => startRow + k));
    }
  }

  /** @internal Evict a block from memory (LRU). */
  __serverEvict(startRow: number, count: number): void {
    for (let k = 0; k < count && startRow + k < this.data.length; k++) {
      delete this.data[startRow + k];
      delete this.rowIds[startRow + k];
    }
  }

  /** @internal Drop loaded blocks (sort/filter delegated to the server). */
  __serverPurge(): void {
    const total = this.data.length;
    this.data = new Array<T>(total) as T[];
    this.rowIds = new Array<RowId>(total);
    this.rebuildView();
    this.renderer.invalidateAll();
  }

  /**
   * Value used for sorting/aggregations: the computed result when the cell
   * holds a formula, the raw value otherwise.
   */
  computedValue(dataIndex: number, col: ColumnState<T>): unknown {
    const raw = getPath(this.data[dataIndex], col.field);
    if (this.formulaEngine && isFormula(raw)) {
      const v = this.formulaEngine.cell(this.columns.indexOf(col), dataIndex);
      return v instanceof FormulaError ? v.code : v;
    }
    return raw;
  }

  /** Height of a view row (uniform unless `getRowHeight` is provided). */
  rowHeightOf(viewIndex: number): number {
    const fn = this.options.getRowHeight;
    if (!fn) return this.options.rowHeight;
    const row = this.getRowByViewIndex(viewIndex);
    return row ? Math.max(20, fn(row, viewIndex)) : this.options.rowHeight;
  }

  get rowCount(): number {
    return this.view.length;
  }
  get totalRowCount(): number {
    return this.data.length;
  }

  getRowByViewIndex(viewIndex: number): T | undefined {
    const v = this.view[viewIndex];
    return v >= 0 ? this.data[v] : undefined;
  }
  getRowIdByViewIndex(viewIndex: number): RowId {
    const v = this.view[viewIndex];
    return v >= 0 ? this.rowIds[v] : '#group:' + (this.groupRows[-v - 1]?.path ?? viewIndex);
  }

  getValue(viewIndex: number, colId: string): unknown {
    const row = this.getRowByViewIndex(viewIndex);
    const col = this.columns.find((c) => c.id === colId);
    return row && col ? getPath(row, col.field) : undefined;
  }

  getDisplayValue(viewIndex: number, colId: string): string {
    const col = this.columns.find((c) => c.id === colId);
    const row = this.getRowByViewIndex(viewIndex);
    if (!col || !row) return '';
    const raw = getPath(row, col.field);
    if (this.formulaEngine && isFormula(raw)) {
      const v = this.formulaEngine.cell(this.columns.indexOf(col as ColumnState<T>), this.view[viewIndex]);
      if (v instanceof FormulaError) return v.code;
      if (col.def.valueFormatter) return col.def.valueFormatter(v, row);
      if (typeof v === 'number')
        return v.toLocaleString(this.options.locale, { maximumFractionDigits: 6 });
      if (typeof v === 'boolean') return v ? 'VERO' : 'FALSO';
      return v === null ? '' : String(v);
    }
    if (col.def.valueFormatter) return col.def.valueFormatter(raw, row);
    return formatValue(raw, col as ColumnState, this.options.locale);
  }

  /** What the inline editor should show: the formula source, not the result. */
  getEditValue(viewIndex: number, colId: string): string {
    const raw = this.getValue(viewIndex, colId);
    return this.formulaEngine && isFormula(raw) ? raw : this.getDisplayValue(viewIndex, colId);
  }

  parseInput(input: string, viewIndex: number, colId: string): unknown {
    if (this.formulaEngine && input.startsWith('=')) return input; // store the formula source
    const col = this.columns.find((c) => c.id === colId)!;
    const row = this.getRowByViewIndex(viewIndex)!;
    if (col.def.valueParser) return col.def.valueParser(input, row);
    return parseValue(input, col as ColumnState, this.options.locale);
  }

  isCellEditable(viewIndex: number, colId: string): boolean {
    if (!this.options.editable) return false;
    const col = this.columns.find((c) => c.id === colId);
    const row = this.getRowByViewIndex(viewIndex);
    if (!col || !row) return false;
    return typeof col.editable === 'function' ? col.editable(row) : col.editable;
  }

  /**
   * Apply a batch of cell changes atomically: validates, emits cancellable
   * `cellValueChanging` per cell, records one undo entry, repaints, and
   * dispatches `cellsChanged` + per-cell `cellValueChanged`.
   */
  applyChanges(changes: CellChange[], source: ChangeSource, recordUndo = true): CellChange[] {
    const applied: CellChange[] = [];
    const errors: { cell: CellRef; message: string }[] = [];

    for (const ch of changes) {
      const col = this.columns.find((c) => c.id === ch.colId);
      const dataIndex = this.rowIds.indexOf(ch.rowId);
      if (!col || dataIndex < 0) continue;
      const row = this.data[dataIndex];

      if (col.def.validator) {
        const msg = col.def.validator(ch.newValue, row);
        if (msg) {
          errors.push({ cell: { rowIndex: ch.rowIndex, colId: ch.colId }, message: msg });
          continue;
        }
      }
      const e = this.events.emit('cellValueChanging', { ...ch, source });
      if (e.defaultPrevented) continue;

      setPath(row as RowData, col.field, ch.newValue);
      this.dirty.add(ch.rowId);
      applied.push(ch);
    }

    if (errors.length) {
      this.events.emit('validationFailed', { errors });
      this.renderer?.flashInvalid(errors);
    }
    if (!applied.length) return applied;

    if (recordUndo) {
      this.pushUndo({ kind: 'cells', changes: applied });
    }

    this.renderer.invalidateRows(applied.map((c) => c.rowIndex));
    this.events.emit('cellsChanged', { changes: applied, source });
    for (const ch of applied) this.events.emit('cellValueChanged', { ...ch, source });
    this.events.emit('dirtyStateChanged', { dirtyRowIds: [...this.dirty] });
    return applied;
  }

  setCellValue(viewIndex: number, colId: string, newValue: unknown, source: ChangeSource = 'api'): void {
    this.applyChanges(
      [
        {
          rowId: this.getRowIdByViewIndex(viewIndex),
          rowIndex: viewIndex,
          colId,
          oldValue: this.getValue(viewIndex, colId),
          newValue,
        },
      ],
      source,
    );
  }

  /* ---------------------------------------------------------------- */
  /* Undo / redo                                                       */
  /* ---------------------------------------------------------------- */

  private pushUndo(entry: UndoEntry<T>): void {
    this.undoStack.push(entry);
    if (this.undoStack.length > this.options.undoLimit) this.undoStack.shift();
    this.redoStack = [];
    this.emitUndoState();
  }

  undo(): void {
    const entry = this.undoStack.pop();
    if (!entry) return;
    if (entry.kind === 'cells') {
      const inverse = entry.changes.map((c) => ({ ...c, oldValue: c.newValue, newValue: c.oldValue }));
      this.applyChanges(inverse, 'undo', false);
      this.events.emit('undoApplied', { changes: inverse });
    } else if (entry.op === 'insert') {
      this.removeByIds(entry.ids);
    } else {
      this.reinsert(entry);
    }
    this.redoStack.push(entry);
    this.emitUndoState();
  }

  redo(): void {
    const entry = this.redoStack.pop();
    if (!entry) return;
    if (entry.kind === 'cells') {
      this.applyChanges(entry.changes, 'redo', false);
      this.events.emit('redoApplied', { changes: entry.changes });
    } else if (entry.op === 'insert') {
      this.reinsert(entry);
    } else {
      this.removeByIds(entry.ids);
    }
    this.undoStack.push(entry);
    this.emitUndoState();
  }

  private removeByIds(ids: RowId[]): void {
    const set = new Set(ids);
    for (let i = this.data.length - 1; i >= 0; i--) {
      if (set.has(this.rowIds[i])) {
        this.data.splice(i, 1);
        this.rowIds.splice(i, 1);
      }
    }
    this.rebuildView();
    this.renderer.invalidateAll();
    this.events.emit('rowsRemoved', { rowIds: ids });
  }

  private reinsert(entry: { rows: T[]; ids: RowId[]; indices: number[] }): void {
    for (let k = 0; k < entry.rows.length; k++) {
      const at = Math.min(entry.indices[k], this.data.length);
      this.data.splice(at, 0, entry.rows[k]);
      this.rowIds.splice(at, 0, entry.ids[k]);
    }
    this.rebuildView();
    this.renderer.invalidateAll();
    this.events.emit('rowsInserted', { rowIds: entry.ids, viewIndex: -1 });
  }

  /* ---------------------------------------------------------------- */
  /* Row insertion / removal (undoable)                                 */
  /* ---------------------------------------------------------------- */

  /** Insert rows at the data position of `atViewIndex` (or append). Undoable. */
  insertRows(rows: T[], atViewIndex?: number): RowId[] {
    if (this.serverMode || this.treeMode) return []; // gerarchia/backend possiedono le righe
    const dataAt =
      atViewIndex !== undefined && this.view[atViewIndex] !== undefined && this.view[atViewIndex] >= 0
        ? this.view[atViewIndex]
        : this.data.length;
    const ids = rows.map((r) => this.resolveRowId(r));
    const indices = rows.map((_, k) => dataAt + k);
    for (let k = 0; k < rows.length; k++) {
      this.data.splice(dataAt + k, 0, rows[k]);
      this.rowIds.splice(dataAt + k, 0, ids[k]);
      this.dirty.add(ids[k]);
    }
    this.pushUndo({ kind: 'rows', op: 'insert', rows: rows.slice(), ids, indices });
    this.rebuildView();
    this.renderer.invalidateAll();
    this.events.emit('rowsInserted', { rowIds: ids, viewIndex: atViewIndex ?? this.rowCount - 1 });
    this.events.emit('dirtyStateChanged', { dirtyRowIds: [...this.dirty] });
    return ids;
  }

  /** Remove the data rows behind these view indices. Undoable. */
  removeRows(viewIndices: number[]): void {
    if (this.serverMode || this.treeMode) return;
    const dataIdx = [...new Set(viewIndices.map((i) => this.view[i]).filter((v) => v >= 0))].sort(
      (a, b) => a - b,
    );
    if (!dataIdx.length) return;
    const rows = dataIdx.map((i) => this.data[i]);
    const ids = dataIdx.map((i) => this.rowIds[i]);
    for (let k = dataIdx.length - 1; k >= 0; k--) {
      this.data.splice(dataIdx[k], 1);
      this.rowIds.splice(dataIdx[k], 1);
    }
    this.pushUndo({ kind: 'rows', op: 'remove', rows, ids, indices: dataIdx });
    this.rebuildView();
    this.renderer.invalidateAll();
    this.events.emit('rowsRemoved', { rowIds: ids });
  }

  private emitUndoState(): void {
    this.events.emit('undoStackChanged', {
      canUndo: this.undoStack.length > 0,
      canRedo: this.redoStack.length > 0,
    });
  }

  getDirtyRows(): { id: RowId; row: T }[] {
    return [...this.dirty]
      .map((id) => ({ id, row: this.data[this.rowIds.indexOf(id)] }))
      .filter((x) => x.row !== undefined);
  }
  /** Mark a row dirty programmatically (used by nested detail grids). */
  markRowDirty(id: RowId): void {
    this.dirty.add(id);
    this.events.emit('dirtyStateChanged', { dirtyRowIds: [...this.dirty] });
  }

  markClean(): void {
    this.dirty.clear();
    this.events.emit('dirtyStateChanged', { dirtyRowIds: [] });
    this.renderer.invalidateAll();
  }
  isRowDirty(viewIndex: number): boolean {
    return this.dirty.has(this.getRowIdByViewIndex(viewIndex));
  }

  /* ---------------------------------------------------------------- */
  /* Sort & filter pipeline                                            */
  /* ---------------------------------------------------------------- */

  setSortModel(model: SortModelItem[]): void {
    this.sortModel = model;
    if (this.serverMode) this.serverSource!.refetch();
    else {
      this.rebuildView();
      this.renderer.invalidateAll();
    }
    this.events.emit('sortChanged', { sortModel: model });
  }

  toggleSort(colId: string, multi: boolean): void {
    const existing = this.sortModel.find((s) => s.colId === colId);
    let next: SortModelItem[] = multi ? this.sortModel.filter((s) => s.colId !== colId) : [];
    if (!existing) next = [...next, { colId, dir: 'asc' }];
    else if (existing.dir === 'asc') next = [...next, { colId, dir: 'desc' }];
    this.setSortModel(next);
  }

  setFilterModel(model: FilterModelItem[]): void {
    this.filterModel = model;
    if (this.serverMode) this.serverSource!.refetch();
    else {
      this.rebuildView();
      this.renderer.invalidateAll();
    }
    this.events.emit('filterChanged', { filterModel: model });
  }

  /** Distinct values of a column (for the Excel-style set filter). */
  getDistinctValues(colId: string, limit = 1000): string[] {
    const col = this.columns.find((c) => c.id === colId);
    if (!col) return [];
    const set = new Set<string>();
    for (const row of this.data) {
      if (row === undefined) continue; // server-side: not yet loaded
      set.add(formatValue(getPath(row, col.field), col as ColumnState, this.options.locale));
      if (set.size >= limit) break;
    }
    return [...set].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  }

  rebuildView(): void {
    if (this.serverMode) {
      this.groupRows = [];
      this.view = Array.from({ length: this.data.length }, (_, i) => i);
      this.clampSelection();
      return;
    }
    let indices = this.data.map((_, i) => i);

    for (const f of this.filterModel) {
      const col = this.columns.find((c) => c.id === f.colId);
      if (!col) continue;
      if (f.values) {
        const allowed = new Set(f.values.map(String));
        indices = indices.filter((i) =>
          allowed.has(formatValue(getPath(this.data[i], col.field), col as ColumnState, this.options.locale)),
        );
      }
      if (f.text) {
        const needle = f.text.toLowerCase();
        indices = indices.filter((i) =>
          formatValue(getPath(this.data[i], col.field), col as ColumnState, this.options.locale)
            .toLowerCase()
            .includes(needle),
        );
      }
    }

    if (this.sortModel.length) {
      const sorters = this.sortModel
        .map((s) => {
          const col = this.columns.find((c) => c.id === s.colId);
          if (!col) return null;
          const cmp = col.def.comparator ?? defaultComparator(col.type);
          return { col, dir: s.dir === 'asc' ? 1 : -1, cmp };
        })
        .filter((x): x is NonNullable<typeof x> => !!x);
      indices.sort((a, b) => {
        for (const s of sorters) {
          const r = s.cmp(this.computedValue(a, s.col), this.computedValue(b, s.col)) * s.dir;
          if (r !== 0) return r;
        }
        return a - b; // stable
      });
    }

    if (this.options.treeData) {
      this.groupRows = [];
      this.view = this.buildTreeView(indices);
    } else if (this.groupBy.length) {
      this.groupRows = [];
      this.view = this.buildGroupedView(indices, 0, '');
    } else {
      this.groupRows = [];
      this.view = indices;
    }
    this.clampSelection();
  }

  /**
   * Tree view: i nodi del filtro restano visibili insieme ai loro antenati;
   * i fratelli sono ordinati col sortModel; i figli compaiono solo se ogni
   * antenato è espanso.
   */
  private buildTreeView(matched: number[]): number[] {
    const matchSet = new Set(matched);
    const visible = new Set<number>();
    for (const i of matched) {
      let cur = i;
      while (cur >= 0) {
        visible.add(cur);
        cur = this.treeMeta[cur].parent;
      }
    }
    const childrenOf = new Map<number, number[]>();
    for (let i = 0; i < this.data.length; i++) {
      if (!visible.has(i)) continue;
      const p = this.treeMeta[i].parent;
      let arr = childrenOf.get(p);
      if (!arr) childrenOf.set(p, (arr = []));
      arr.push(i);
    }
    const sorters = this.sortModel
      .map((s) => {
        const col = this.columns.find((c) => c.id === s.colId);
        if (!col) return null;
        const cmp = col.def.comparator ?? defaultComparator(col.type);
        return { col, dir: s.dir === 'asc' ? 1 : -1, cmp };
      })
      .filter((x): x is NonNullable<typeof x> => !!x);
    const sortSiblings = (arr: number[]) => {
      if (!sorters.length) return arr;
      return arr.slice().sort((a, b) => {
        for (const s of sorters) {
          const r = s.cmp(this.computedValue(a, s.col), this.computedValue(b, s.col)) * s.dir;
          if (r !== 0) return r;
        }
        return a - b;
      });
    };
    const out: number[] = [];
    const emit = (parent: number) => {
      for (const i of sortSiblings(childrenOf.get(parent) ?? [])) {
        out.push(i);
        if (this.expandedIds.has(this.rowIds[i])) emit(i);
      }
    };
    emit(-1);
    void matchSet;
    return out;
  }

  /* ---------------------------------------------------------------- */
  /* Tree API                                                           */
  /* ---------------------------------------------------------------- */

  get treeMode(): boolean {
    return !!this.options.treeData;
  }
  treeLevel(viewIndex: number): number {
    return this.treeMode ? this.treeMeta[this.view[viewIndex]]?.level ?? 0 : 0;
  }
  treeHasChildren(viewIndex: number): boolean {
    return this.treeMode && !!this.treeMeta[this.view[viewIndex]]?.hasChildren;
  }
  toggleTreeNode(viewIndex: number): void {
    if (!this.treeHasChildren(viewIndex)) return;
    const id = this.getRowIdByViewIndex(viewIndex);
    if (this.expandedIds.has(id)) {
      this.expandedIds.delete(id);
      this.events.emit('rowCollapsed', { rowId: id, rowIndex: viewIndex });
    } else {
      this.expandedIds.add(id);
      this.events.emit('rowExpanded', { rowId: id, rowIndex: viewIndex });
    }
    this.rebuildView();
    this.renderer.invalidateAll();
  }

  /** Recursively partition sorted+filtered indices into group rows. */
  private buildGroupedView(indices: number[], level: number, parentPath: string): number[] {
    const colId = this.groupBy[level];
    const col = this.columns.find((c) => c.id === colId);
    if (!col) return indices;
    const buckets = new Map<string, number[]>();
    for (const i of indices) {
      const key = formatValue(getPath(this.data[i], col.field), col as ColumnState, this.options.locale);
      let arr = buckets.get(key);
      if (!arr) buckets.set(key, (arr = []));
      arr.push(i);
    }
    const out: number[] = [];
    const keys = [...buckets.keys()].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    for (const key of keys) {
      const children = buckets.get(key)!;
      const path = parentPath + colId + '=' + key + '/';
      const collapsed = this.collapsedGroups.has(path);
      const leafIndices = children; // children of further levels still reference leaves
      const aggs: Record<string, number | null> = {};
      for (const c of this.columns) {
        if (c.def.aggFunc) aggs[c.id] = this.computeAgg(c, leafIndices);
      }
      this.groupRows.push({ path, colId, key, level, count: children.length, aggs, collapsed });
      out.push(-this.groupRows.length); // -(groupIndex+1)
      if (!collapsed) {
        out.push(
          ...(level + 1 < this.groupBy.length
            ? this.buildGroupedView(children, level + 1, path)
            : children),
        );
      }
    }
    return out;
  }

  private computeAgg(col: ColumnState<T>, dataIndices: number[]): number | null {
    const fn = col.def.aggFunc!;
    if (fn === 'count') return dataIndices.length;
    let sum = 0, n = 0, min = Infinity, max = -Infinity;
    for (const i of dataIndices) {
      const v = this.computedValue(i, col);
      if (typeof v !== 'number' || !Number.isFinite(v)) continue;
      n++; sum += v;
      if (v < min) min = v;
      if (v > max) max = v;
    }
    if (!n) return null;
    return fn === 'sum' ? sum : fn === 'avg' ? sum / n : fn === 'min' ? min : max;
  }

  /* ---------------------------------------------------------------- */
  /* Grouping API                                                       */
  /* ---------------------------------------------------------------- */

  setGroupBy(colIds: string[]): void {
    if (this.serverMode) return; // raggruppamento locale non disponibile in server mode
    this.groupBy = colIds;
    this.collapsedGroups.clear();
    this.rebuildView();
    this.renderer.invalidateAll();
    this.events.emit('groupingChanged', { groupBy: colIds });
  }

  isGroupRow(viewIndex: number): boolean {
    return this.view[viewIndex] < 0;
  }
  getGroupRow(viewIndex: number): GroupRow | undefined {
    const v = this.view[viewIndex];
    return v < 0 ? this.groupRows[-v - 1] : undefined;
  }

  toggleGroup(viewIndex: number): void {
    const g = this.getGroupRow(viewIndex);
    if (!g) return;
    const collapsed = !this.collapsedGroups.has(g.path);
    collapsed ? this.collapsedGroups.add(g.path) : this.collapsedGroups.delete(g.path);
    this.rebuildView();
    this.renderer.invalidateAll();
    this.events.emit('groupToggled', { path: g.path, collapsed });
  }

  /** Grand-total aggregates over the current (filtered) leaf rows. */
  grandTotals(): { count: number; aggs: Record<string, number | null> } {
    const leaves = this.view.filter((v) => v >= 0);
    const aggs: Record<string, number | null> = {};
    for (const c of this.columns) if (c.def.aggFunc) aggs[c.id] = this.computeAgg(c, leaves);
    return { count: leaves.length, aggs };
  }

  showTotalsRow(): boolean {
    return this.options.totalsRow ?? this.columns.some((c) => c.def.aggFunc);
  }

  /* ---------------------------------------------------------------- */
  /* Selection                                                         */
  /* ---------------------------------------------------------------- */

  setSelection(ranges: CellRange[], active: CellRef | null): void {
    this.ranges = ranges.map(normalizeRange);
    this.activeCell = active;
    this.renderer.renderSelection();
    this.events.emit('selectionChanged', { ranges: this.ranges });
    this.events.emit('cellFocused', { cell: active });
  }

  selectCell(rowIndex: number, colIndex: number): void {
    const col = this.columnByVisibleIndex(colIndex);
    if (!col || rowIndex < 0 || rowIndex >= this.rowCount) return;
    this.setSelection(
      [{ startRow: rowIndex, endRow: rowIndex, startCol: colIndex, endCol: colIndex }],
      { rowIndex, colId: col.id },
    );
  }

  private clampSelection(): void {
    const maxRow = Math.max(0, this.rowCount - 1);
    this.ranges = this.ranges.map((r) => ({
      ...r,
      startRow: Math.min(r.startRow, maxRow),
      endRow: Math.min(r.endRow, maxRow),
    }));
    if (this.activeCell && this.activeCell.rowIndex > maxRow) this.activeCell = null;
  }

  /* ---------------------------------------------------------------- */
  /* Master / detail (nested sub-tables)                                */
  /* ---------------------------------------------------------------- */

  /** rowIds currently expanded. Survives sort/filter because it is id-based. */
  readonly expandedIds = new Set<RowId>();

  /** Maximum height of a detail panel (the actual height adapts to content). */
  get detailHeight(): number {
    return this.options.masterDetail?.height ?? 260;
  }



  /** The nested rows for a view row, or undefined when there is no detail. */
  getDetailRows(viewIndex: number): RowData[] | undefined {
    const md = this.options.masterDetail;
    const row = this.getRowByViewIndex(viewIndex);
    if (!md || !row) return undefined;
    const detail = md.getDetail
      ? md.getDetail(row)
      : md.field
        ? (getPath(row, md.field) as RowData[] | undefined)
        : undefined;
    return Array.isArray(detail) && detail.length ? detail : undefined;
  }

  /** Columns for the nested grid: explicit, or inferred from the data. */
  getDetailColumns(rows: RowData[]): ColumnDef[] {
    return this.options.masterDetail?.columns ?? inferColumns(rows).columns;
  }

  hasDetail(viewIndex: number): boolean {
    if (this.options.masterDetail?.detailRenderer) return this.getRowByViewIndex(viewIndex) !== undefined;
    return this.getDetailRows(viewIndex) !== undefined;
  }

  /** Nested array inside the detail rows → the sub-table is itself expandable. */
  getNestedDetailField(rows: RowData[]): string | undefined {
    if (this.options.masterDetail?.columns) return undefined; // colonne esplicite: niente ricorsione implicita
    return inferColumns(rows).detailFields[0];
  }

  /**
   * Deterministic geometry of a detail panel — one source of truth shared
   * by offsets and rendering, so scrollbars never appear "by accident":
   * the panel is sized to the exact content; when capped, scrollbar space
   * is added to the budget; when everything fits, overflow is disabled.
   */
  detailLayout(viewIndex: number): {
    height: number;
    innerWidth: number;
    fits: boolean;
  } {
    const md = this.options.masterDetail!;
    const row = this.getRowByViewIndex(viewIndex);
    const PAD = 14; // padding verticale del pannello (6+8)
    const SCROLLBAR = 18;
    const max = this.detailHeight;

    if (md.detailRenderer) {
      const h = Math.min(960, (row && md.getHeight?.(row)) ?? max);
      return { height: h, innerWidth: -1, fits: true }; // -1 = larghezza libera
    }

    const rows = this.getDetailRows(viewIndex);
    if (!rows) return { height: 0, innerWidth: 0, fits: true };
    const cols = this.getDetailColumns(rows);
    const nestedGutter = this.getNestedDetailField(rows) ? 56 : 0;
    const nestedRowH = Math.max(26, this.options.rowHeight - 4);
    const contentH = 30 /*header*/ + rows.length * nestedRowH + 2 /*bordi root*/;
    const naturalW = nestedGutter + cols.reduce((s, c) => s + (c.width ?? 140), 0) + 2;
    const availableW = Math.max(280, (this.renderer?.viewport.clientWidth ?? 800) - 90);

    let scrollV = contentH + PAD > max;
    const innerW = Math.min(naturalW + (scrollV ? SCROLLBAR : 0), availableW);
    const scrollH = naturalW + (scrollV ? SCROLLBAR : 0) > availableW;
    // La scrollbar orizzontale ruba altezza: se così facendo serve anche la
    // verticale, il budget la include (niente comparsa a catena).
    let height = Math.min(max, contentH + PAD + (scrollH ? SCROLLBAR : 0));
    if (scrollH && contentH + PAD + SCROLLBAR > max) scrollV = true;
    return { height, innerWidth: innerW, fits: !scrollV && !scrollH };
  }

  detailHeightFor(viewIndex: number): number {
    if (!this.hasDetail(viewIndex)) return 0;
    return this.detailLayout(viewIndex).height;
  }
  isExpanded(viewIndex: number): boolean {
    return this.expandedIds.has(this.getRowIdByViewIndex(viewIndex));
  }

  toggleDetail(viewIndex: number): void {
    if (!this.hasDetail(viewIndex)) return;
    const id = this.getRowIdByViewIndex(viewIndex);
    if (this.expandedIds.has(id)) {
      this.expandedIds.delete(id);
      this.events.emit('rowCollapsed', { rowId: id, rowIndex: viewIndex });
    } else {
      this.expandedIds.add(id);
      this.events.emit('rowExpanded', { rowId: id, rowIndex: viewIndex });
    }
    this.renderer.invalidateAll();
  }

  /* ---------------------------------------------------------------- */
  /* Public façade & lifecycle                                         */
  /* ---------------------------------------------------------------- */

  navigateTo(rowIndex: number, colIndex: number): void {
    this.selectCell(rowIndex, colIndex);
    this.renderer.ensureCellVisible(rowIndex, colIndex);
  }

  startEditing(cell?: CellRef): void {
    this.interaction.startEdit(cell ?? this.activeCell ?? undefined);
  }
  stopEditing(cancel = false): void {
    this.interaction.stopEdit(cancel);
  }

  exportTsv(range?: CellRange): string {
    return this.interaction.rangeToTsv(
      range ?? {
        startRow: 0,
        endRow: this.rowCount - 1,
        startCol: 0,
        endCol: this.visibleColumns().length - 1,
      },
    );
  }

  exportCsv(): string {
    const cols = this.visibleColumns();
    const esc = (s: string) => (/[",\n;]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s);
    const lines = [cols.map((c) => esc(c.header)).join(';')];
    for (let r = 0; r < this.rowCount; r++) {
      lines.push(cols.map((c) => esc(this.getDisplayValue(r, c.id))).join(';'));
    }
    return lines.join('\n');
  }

  refresh(): void {
    this.rebuildView();
    this.renderer.invalidateAll();
  }

  /* ---------------------------------------------------------------- */
  /* Find & replace                                                     */
  /* ---------------------------------------------------------------- */

  /** All cells whose display value contains `query`. */
  findAll(query: string, opts: { matchCase?: boolean } = {}): CellRef[] {
    if (!query) return [];
    const needle = opts.matchCase ? query : query.toLowerCase();
    const cols = this.visibleColumns();
    const out: CellRef[] = [];
    for (let r = 0; r < this.rowCount; r++) {
      if (this.isGroupRow(r)) continue;
      for (const c of cols) {
        const hay = this.getDisplayValue(r, c.id);
        if ((opts.matchCase ? hay : hay.toLowerCase()).includes(needle))
          out.push({ rowIndex: r, colId: c.id });
      }
    }
    return out;
  }

  /** Replace `query` with `replacement` in the given cells (undoable, validated). */
  replaceIn(cells: CellRef[], query: string, replacement: string, opts: { matchCase?: boolean } = {}): number {
    const flags = opts.matchCase ? 'g' : 'gi';
    const rx = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags);
    const changes: CellChange[] = [];
    for (const cell of cells) {
      if (!this.isCellEditable(cell.rowIndex, cell.colId)) continue;
      const display = this.getDisplayValue(cell.rowIndex, cell.colId);
      const next = display.replace(rx, replacement);
      if (next === display) continue;
      changes.push({
        rowId: this.getRowIdByViewIndex(cell.rowIndex),
        rowIndex: cell.rowIndex,
        colId: cell.colId,
        oldValue: this.getValue(cell.rowIndex, cell.colId),
        newValue: this.parseInput(next, cell.rowIndex, cell.colId),
      });
    }
    return this.applyChanges(changes, 'user').length;
  }

  /* ---------------------------------------------------------------- */
  /* Saved views                                                        */
  /* ---------------------------------------------------------------- */

  getState(): GridState {
    return {
      sortModel: this.sortModel.map((s) => ({ ...s })),
      filterModel: JSON.parse(JSON.stringify(this.filterModel)),
      groupBy: [...this.groupBy],
      columnWidths: Object.fromEntries(this.columns.map((c) => [c.id, c.width])),
      hiddenColumns: this.columns.filter((c) => !c.visible).map((c) => c.id),
    };
  }

  setState(state: Partial<GridState>): void {
    if (state.columnWidths)
      for (const c of this.columns)
        if (state.columnWidths[c.id]) c.width = state.columnWidths[c.id];
    if (state.hiddenColumns)
      for (const c of this.columns) c.visible = !state.hiddenColumns.includes(c.id);
    if (state.groupBy) this.groupBy = [...state.groupBy];
    if (state.sortModel) this.sortModel = state.sortModel.map((s) => ({ ...s }));
    if (state.filterModel) this.filterModel = JSON.parse(JSON.stringify(state.filterModel));
    this.layoutColumns();
    this.rebuildView();
    this.renderer.invalidateAll();
    this.events.emit('sortChanged', { sortModel: this.sortModel });
    this.events.emit('filterChanged', { filterModel: this.filterModel });
  }

  /** Move a column to a new position among its siblings (pinned stay pinned). */
  moveColumn(colId: string, toIndex: number): void {
    const from = this.columns.findIndex((c) => c.id === colId);
    if (from < 0) return;
    const [col] = this.columns.splice(from, 1);
    this.columns.splice(Math.max(0, Math.min(this.columns.length, toIndex)), 0, col);
    this.layoutColumns();
    this.renderer.invalidateAll();
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    this.interaction.destroy();
    this.renderer.destroy();
    this.events.emit('gridDestroyed', {});
    this.events.clear();
  }
}

export function normalizeRange(r: CellRange): CellRange {
  return {
    startRow: Math.min(r.startRow, r.endRow),
    endRow: Math.max(r.startRow, r.endRow),
    startCol: Math.min(r.startCol, r.endCol),
    endCol: Math.max(r.startCol, r.endCol),
  };
}
