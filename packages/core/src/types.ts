/**
 * @exceltable/core — Public type definitions.
 * Every type here is part of the stable public API (semver-protected).
 */

/** Built-in cell data types. Drives default editor, parser, formatter and comparator. */
export type CellType = 'text' | 'number' | 'date' | 'boolean' | 'select';

export type RowData = Record<string, unknown>;
export type RowId = string;

/** Parameters passed to a custom cell renderer. */
export interface CellRendererParams<T extends RowData = RowData> {
  value: unknown;
  /** Formatted display string (after valueFormatter / locale formatting). */
  displayValue: string;
  row: T;
  rowIndex: number;
  colId: string;
}

/**
 * What a custom renderer may return:
 *  - string            → rendered as text
 *  - Node              → DOM element / framework-mounted component
 *  - { el, destroy }   → element with a cleanup hook (used by the React /
 *                        Angular helpers to unmount components correctly)
 *  - an array of any of the above → stacked vertically inside the cell
 *    (value on one line, custom element on the next — "multi-content" cells)
 */
export type CellContent = string | Node | { el: Node; destroy?: () => void };

/** Parameters passed to a master/detail custom renderer. */
export interface DetailRendererParams<T extends RowData = RowData> {
  row: T;
  rowIndex: number;
  rowId: RowId;
}
export type CellRendererResult = CellContent | CellContent[];
export type CellRenderer<T extends RowData = RowData> = (
  params: CellRendererParams<T>,
) => CellRendererResult;

/** Definition of a single column. */
export interface ColumnDef<T extends RowData = RowData> {
  /** Unique column id. Defaults to `field`. */
  id?: string;
  /**
   * Property of the row object this column reads/writes.
   * Supports dot paths into nested JSON: `"cliente.indirizzo.citta"`.
   */
  field: string;
  /** Header label shown to the user. */
  header?: string;
  /** Data type — selects default editor/parser/comparator. @default 'text' */
  type?: CellType;
  /** Initial width in px. @default 140 */
  width?: number;
  minWidth?: number;
  maxWidth?: number;
  /** Pin the column to the left edge (Excel "freeze"). */
  pinned?: 'left' | null;
  /** Whether cells can be edited. Function form receives the row. @default true */
  editable?: boolean | ((row: T) => boolean);
  /** @default true */
  sortable?: boolean;
  /** @default true */
  resizable?: boolean;
  /** Allowed values for `type: 'select'` columns. */
  options?: string[];
  /** Convert raw value → display string. */
  valueFormatter?: (value: unknown, row: T) => string;
  /** Convert user-typed string → raw value. */
  valueParser?: (input: string, row: T) => unknown;
  /** Return an extra CSS class for a cell, or null. */
  cellClass?: (value: unknown, row: T) => string | null;
  /** Return an error message for an invalid value, or null when valid. */
  validator?: (value: unknown, row: T) => string | null;
  /** Custom comparator for sorting. */
  comparator?: (a: unknown, b: unknown) => number;
  /**
   * Custom cell renderer: return a string, a DOM node, a `{el, destroy}`
   * pair, or an array of them (stacked lines). Use the `reactCell` /
   * `templateCell` helpers from the wrappers to mount framework components.
   */
  cellRenderer?: CellRenderer<T>;
  /** Allow text to wrap on multiple lines inside the cell. */
  wrapText?: boolean;
  /** Aggregation shown on group rows and in the totals row. */
  aggFunc?: 'sum' | 'avg' | 'min' | 'max' | 'count';
}

/** Internal, fully-resolved column (all defaults applied). @internal */
export interface ColumnState<T extends RowData = RowData> extends Required<Pick<ColumnDef<T>, 'field'>> {
  id: string;
  header: string;
  type: CellType;
  width: number;
  minWidth: number;
  maxWidth: number;
  pinned: 'left' | null;
  editable: boolean | ((row: T) => boolean);
  sortable: boolean;
  resizable: boolean;
  visible: boolean;
  /** x offset within its pane, computed. */
  left: number;
  def: ColumnDef<T>;
}

export interface SortModelItem {
  colId: string;
  dir: 'asc' | 'desc';
}

/** Per-column filter. `values` = Excel-style set filter; `text` = contains filter. */
export interface FilterModelItem {
  colId: string;
  values?: unknown[];
  text?: string;
}

export interface CellRef {
  /** Index in the current view (after sort/filter). */
  rowIndex: number;
  colId: string;
}

/** Rectangular selection range, inclusive, in view coordinates. */
export interface CellRange {
  startRow: number;
  endRow: number;
  startCol: number; // visible column index
  endCol: number;
}

export interface GridOptions<T extends RowData = RowData> {
  columns: ColumnDef<T>[];
  data?: T[];
  /** Stable row identity — required for transactions/dirty tracking. Defaults to array identity. */
  getRowId?: (row: T) => RowId;
  /** @default 32 */
  rowHeight?: number;
  /** @default 36 */
  headerHeight?: number;
  /** Show the Excel-style row-number gutter. @default true */
  showRowNumbers?: boolean;
  /** Allow Enter/typing to begin editing. @default true */
  editable?: boolean;
  /** Locale used for number/date parsing and formatting. @default 'it-IT' */
  locale?: string;
  /** Max undo stack depth. @default 200 */
  undoLimit?: number;
  /** Theme class suffix, e.g. 'excel' | 'dark'. @default 'excel' */
  theme?: string;
  /**
   * Nested sub-tables (master/detail). When a row has nested array data,
   * an expand chevron appears in the row-number gutter; expanding shows a
   * nested grid. Columns of the nested grid are inferred from the data
   * unless provided.
   */
  masterDetail?: {
    /** Field containing the nested array, e.g. `"ordini"`. */
    field?: string;
    /** Or compute the detail rows for a row (wins over `field`). */
    getDetail?: (row: T) => RowData[] | undefined;
    /** Max height in px of the expanded panel (adapts to content below this). @default 260 */
    height?: number;
    /** Columns for the nested grid. Inferred from the data when omitted. */
    columns?: ColumnDef[];
    /**
     * Render ANYTHING in the panel instead of the nested grid: a string,
     * a DOM node with your own HTML/CSS, or `{el, destroy}` for components
     * with lifecycle (the destroy hook runs on collapse/recycle).
     * When set, every row is expandable.
     */
    detailRenderer?: (params: DetailRendererParams<T>) => CellContent;
    /** Panel height per row when using `detailRenderer`. @default `height` */
    getHeight?: (row: T) => number;
  };
  /** Accessible label announced by screen readers. @default 'Tabella dati' */
  ariaLabel?: string;
  /** Group rows by these column ids (Excel subtotal-style, collapsible). */
  groupBy?: string[];
  /** Show the pinned grand-totals row at the bottom. @default true when any column has aggFunc */
  totalsRow?: boolean;
  /**
   * Enable the formula engine: cells whose value starts with `=` are
   * computed (es. `=SOMMA(C2:C10)`, funzioni in italiano e inglese).
   */
  formulas?: boolean;
  /** Server-side row model: lazy block loading from a backend. */
  serverSide?: import('./serverside').ServerSideOptions<T>;
  /** Override any UI string (default: italiano; `EN` esportato pronto). */
  strings?: Partial<import('./i18n').Strings>;
  /**
   * Per-row height in px (deterministic — the app decides, e.g. from an
   * estimated line count for `wrapText` columns). Offsets use prefix sums,
   * so virtualization stays exact. Min 20px.
   */
  getRowHeight?: (row: T, viewIndex: number) => number;
  /** Hierarchical rows (WBS-style) with indentation and chevrons. */
  treeData?: {
    /** Children accessor (wins over `childrenField`). */
    getChildren?: (row: T) => T[] | undefined;
    /** @default 'children' */
    childrenField?: string;
    /** Indentation per level in px. @default 18 */
    indent?: number;
  };
  /**
   * Customize the right-click menu: receive the default items and the
   * context, return the items to show. `{ sep: true }` is a separator.
   */
  contextMenuItems?: (
    defaultItems: ContextMenuItem[],
    ctx: { cell?: CellRef; colId?: string; grid: unknown },
  ) => ContextMenuItem[];
}

/** A single cell mutation, as carried by transactions and undo entries. */
export interface CellChange {
  rowId: RowId;
  rowIndex: number;
  colId: string;
  oldValue: unknown;
  newValue: unknown;
}

export type ChangeSource = 'user' | 'api' | 'clipboard' | 'fill' | 'undo' | 'redo' | 'import';

/* ------------------------------------------------------------------ */
/* Events                                                              */
/* ------------------------------------------------------------------ */

export interface GridEventBase {
  type: string;
  timestamp: number;
  /** Call to cancel — only honoured by events documented as cancellable. */
  preventDefault(): void;
  readonly defaultPrevented: boolean;
}

export interface EventMap<T extends RowData = RowData> {
  gridReady: GridEventBase;
  gridDestroyed: GridEventBase;
  firstDataRendered: GridEventBase;
  rowDataChanged: GridEventBase & { rowCount: number };
  /** Batched cell mutations (edit, paste, fill, undo…). */
  cellsChanged: GridEventBase & { changes: CellChange[]; source: ChangeSource };
  /** Single-cell convenience event, fired per change after `cellsChanged`. */
  cellValueChanged: GridEventBase & CellChange & { source: ChangeSource };
  /** Cancellable. Veto or inspect an edit before it is applied. */
  cellValueChanging: GridEventBase & CellChange & { source: ChangeSource };
  cellFocused: GridEventBase & { cell: CellRef | null };
  cellClicked: GridEventBase & { cell: CellRef; originalEvent: MouseEvent };
  cellDoubleClicked: GridEventBase & { cell: CellRef; originalEvent: MouseEvent };
  /** Cancellable. */
  cellEditingStarted: GridEventBase & { cell: CellRef };
  cellEditingStopped: GridEventBase & { cell: CellRef; cancelled: boolean };
  selectionChanged: GridEventBase & { ranges: CellRange[] };
  /** Cancellable. */
  copyStart: GridEventBase & { range: CellRange };
  copyEnd: GridEventBase & { range: CellRange; tsv: string };
  /** Cancellable. `rows` is the parsed clipboard matrix. */
  pasteStart: GridEventBase & { rows: string[][]; target: CellRange };
  pasteEnd: GridEventBase & { changes: CellChange[]; rejected: CellChange[] };
  /** Cancellable. */
  fillStart: GridEventBase & { source: CellRange; target: CellRange };
  fillEnd: GridEventBase & { changes: CellChange[] };
  sortChanged: GridEventBase & { sortModel: SortModelItem[] };
  filterChanged: GridEventBase & { filterModel: FilterModelItem[] };
  columnResized: GridEventBase & { colId: string; width: number };
  columnVisibilityChanged: GridEventBase & { colId: string; visible: boolean };
  undoApplied: GridEventBase & { changes: CellChange[] };
  redoApplied: GridEventBase & { changes: CellChange[] };
  undoStackChanged: GridEventBase & { canUndo: boolean; canRedo: boolean };
  validationFailed: GridEventBase & { errors: { cell: CellRef; message: string }[] };
  viewportChanged: GridEventBase & { firstRow: number; lastRow: number };
  dirtyStateChanged: GridEventBase & { dirtyRowIds: RowId[] };
  rowExpanded: GridEventBase & { rowId: RowId; rowIndex: number };
  rowCollapsed: GridEventBase & { rowId: RowId; rowIndex: number };
  rowsInserted: GridEventBase & { rowIds: RowId[]; viewIndex: number };
  rowsRemoved: GridEventBase & { rowIds: RowId[] };
  groupToggled: GridEventBase & { path: string; collapsed: boolean };
  groupingChanged: GridEventBase & { groupBy: string[] };
  findResultsChanged: GridEventBase & { query: string; total: number; current: number };
}

/** A group header row produced by groupBy. */
export interface GroupRow {
  /** Unique path of the group, e.g. "zona=Nord Milano/stato=In corso". */
  path: string;
  colId: string;
  key: string;
  level: number;
  /** Leaf data rows in this group. */
  count: number;
  /** Aggregates per column id (columns with aggFunc). */
  aggs: Record<string, number | null>;
  collapsed: boolean;
}

/** Serializable grid state for saved views. */
export interface GridState {
  sortModel: SortModelItem[];
  filterModel: FilterModelItem[];
  groupBy: string[];
  columnWidths: Record<string, number>;
  hiddenColumns: string[];
}

export interface ContextMenuItem {
  label?: string;
  key?: string;
  action?: () => void;
  sep?: boolean;
  disabled?: boolean;
}

export type EventType = keyof EventMap;
export type EventHandler<K extends EventType, T extends RowData = RowData> = (
  event: EventMap<T>[K],
) => void;
