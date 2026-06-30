/**
 * Renderer — virtualized DOM rendering.
 *
 *  - Single scroll viewport; header and pinned columns use position:sticky.
 *  - Only visible rows (+ overscan) and visible columns exist in the DOM.
 *  - Master/detail: expanded rows insert a fixed-height detail panel hosting
 *    a nested Grid; row offsets are computed analytically (O(log e) per row
 *    via binary search over the sorted expanded indices), so virtualization
 *    stays exact with thousands of rows and any number of open panels.
 *  - Custom cell renderers may return DOM nodes / framework components;
 *    their `destroy` hooks are tracked per row and invoked on recycle.
 *  - Selection is painted with CSS classes on visible cells (correct with
 *    pinned columns, no overlay math).
 *  - Accessibility: full ARIA grid pattern (grid/row/columnheader/gridcell,
 *    aria-rowcount/colcount/rowindex/colindex/selected/sort,
 *    aria-activedescendant), AA-contrast themes, visible focus.
 */
import type { Grid } from './grid';
import type { CellContent, CellRange, CellRendererResult, ColumnState, RowData } from './types';

const OVERSCAN = 6;
const GUTTER = 56;

const esc = (s: string) =>
  s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);

interface DetailEntry {
  el: HTMLElement;
  grid: Grid | null;
  destroy?: () => void;
}

export class Renderer<T extends RowData> {
  readonly root: HTMLElement;
  readonly viewport: HTMLElement;
  private canvas: HTMLElement;
  private headerRow: HTMLElement;
  private rowPool = new Map<number, HTMLElement>();
  private detailPool = new Map<number, DetailEntry>();
  private destroyers = new Map<HTMLElement, (() => void)[]>();
  private dirtyRows = new Set<number>();
  private rafId = 0;
  private colWindow: [number, number] = [0, -1];
  private lastViewport: [number, number] = [-1, -1];
  private filterPopup: HTMLElement | null = null;
  /** Sorted view indices of expanded rows, rebuilt on invalidateAll. */
  private expanded: number[] = [];
  /** expandedPrefix[k] = total detail height of expanded[0..k-1]. */
  private expandedPrefix: number[] = [0];
  /** Prefix sums of row heights when `getRowHeight` is set (else null). */
  private heightPrefix: Float64Array | null = null;
  private hasCustomRenderers = false;
  fillHandle: HTMLElement;

  private totalsRow: HTMLElement | null = null;

  constructor(private container: HTMLElement, private grid: Grid<T>) {
    this.root = el('div', `eg-root eg-theme-${grid.options.theme}`);
    this.root.setAttribute('role', 'grid');
    this.root.setAttribute('aria-label', grid.options.ariaLabel ?? grid.t('gridLabel'));
    this.root.setAttribute('aria-multiselectable', 'true');
    this.viewport = el('div', 'eg-viewport');
    this.headerRow = el('div', 'eg-header-row');
    this.headerRow.setAttribute('role', 'row');
    this.canvas = el('div', 'eg-canvas');
    this.canvas.setAttribute('role', 'rowgroup');
    this.fillHandle = el('div', 'eg-fill-handle');
    this.fillHandle.setAttribute('aria-hidden', 'true');
    this.viewport.append(this.headerRow, this.canvas);
    this.root.append(this.viewport);
  }

  mount(): void {
    this.container.appendChild(this.root);
    this.viewport.addEventListener('scroll', () => this.schedule(), { passive: true });
    if (typeof ResizeObserver !== 'undefined')
      new ResizeObserver(() => this.schedule()).observe(this.viewport);
    this.invalidateAll();
  }

  destroy(): void {
    this.closeFilterPopup();
    this.rowPool.forEach((r) => this.disposeRow(r));
    this.detailPool.forEach((d) => this.disposeDetail(d));
    this.root.remove();
    this.rowPool.clear();
    this.detailPool.clear();
  }

  /* ---------------------------------------------------------------- */
  /* Geometry — analytic offsets with expanded detail panels           */
  /* ---------------------------------------------------------------- */

  get gutter(): number {
    return this.grid.options.showRowNumbers || this.grid.options.masterDetail || this.grid.treeMode
      ? GUTTER
      : 0;
  }
  totalWidth(): number {
    return this.gutter + this.grid.pinnedWidth() + this.grid.centerWidth();
  }

  /** Number of expanded rows strictly before view index i (binary search). */
  private expandedBefore(i: number): number {
    let lo = 0,
      hi = this.expanded.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      this.expanded[mid] < i ? (lo = mid + 1) : (hi = mid);
    }
    return lo;
  }

  rowTop(i: number): number {
    const base = this.heightPrefix ? this.heightPrefix[i] : i * this.grid.options.rowHeight;
    return base + this.expandedPrefix[this.expandedBefore(i)];
  }

  totalHeight(): number {
    const base = this.heightPrefix
      ? this.heightPrefix[this.grid.rowCount]
      : this.grid.rowCount * this.grid.options.rowHeight;
    return base + this.expandedPrefix[this.expanded.length];
  }

  /** First row whose bottom edge (incl. its detail) is below y. */
  private rowAtY(y: number): number {
    let lo = 0,
      hi = Math.max(0, this.grid.rowCount - 1);
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      const bottom =
        this.rowTop(mid) +
        this.grid.rowHeightOf(mid) +
        (this.grid.isExpanded(mid) && this.grid.hasDetail(mid) ? this.grid.detailHeightFor(mid) : 0);
      bottom <= y ? (lo = mid + 1) : (hi = mid);
    }
    return lo;
  }

  cellAt(target: EventTarget | null): { row: number; col: number } | null {
    const cell = (target as HTMLElement | null)?.closest?.('[data-row][data-col]');
    if (!cell || (cell as HTMLElement).closest('.eg-detail')) return null;
    return {
      row: Number((cell as HTMLElement).dataset.row),
      col: Number((cell as HTMLElement).dataset.col),
    };
  }

  /* ---------------------------------------------------------------- */
  /* Render cycle                                                      */
  /* ---------------------------------------------------------------- */

  schedule(): void {
    if (this.rafId) return;
    this.rafId = requestAnimationFrame(() => {
      this.rafId = 0;
      this.render();
    });
  }

  invalidateAll(): void {
    this.rowPool.forEach((r) => this.disposeRow(r));
    this.rowPool.clear();
    this.detailPool.forEach((d) => this.disposeDetail(d));
    this.detailPool.clear();
    this.lastViewport = [-1, -1];
    this.expanded = [];
    this.expandedPrefix = [0];
    for (let i = 0; i < this.grid.rowCount; i++) {
      if (this.grid.isExpanded(i) && this.grid.hasDetail(i)) {
        this.expanded.push(i);
        this.expandedPrefix.push(
          this.expandedPrefix[this.expandedPrefix.length - 1] + this.grid.detailHeightFor(i),
        );
      }
    }
    // Altezze variabili: somme prefisse O(n) una volta, lookup O(log n).
    if (this.grid.options.getRowHeight) {
      const n = this.grid.rowCount;
      this.heightPrefix = new Float64Array(n + 1);
      for (let i = 0; i < n; i++)
        this.heightPrefix[i + 1] = this.heightPrefix[i] + this.grid.rowHeightOf(i);
    } else this.heightPrefix = null;
    this.hasCustomRenderers = this.grid.visibleColumns().some((c) => c.def.cellRenderer);
    this.buildHeader();
    this.root.setAttribute('aria-rowcount', String(this.grid.rowCount + 1));
    this.root.setAttribute('aria-colcount', String(this.grid.visibleColumns().length));
    this.canvas.style.height = this.totalHeight() + 'px';
    this.canvas.style.width = this.totalWidth() + 'px';
    this.headerRow.style.width = this.totalWidth() + 'px';
    this.buildTotalsRow();
    this.schedule();
  }

  /** Pinned grand-totals row (sticky bottom). */
  private buildTotalsRow(): void {
    this.totalsRow?.remove();
    this.totalsRow = null;
    if (!this.grid.showTotalsRow()) return;
    const g = this.grid;
    const { count, aggs } = g.grandTotals();
    const row = el('div', 'eg-totals-row');
    row.setAttribute('role', 'row');
    row.setAttribute('aria-label', g.t('totalsRow'));
    row.style.width = this.totalWidth() + 'px';
    const parts: string[] = [];
    if (this.gutter)
      parts.push(`<div class="eg-cell eg-gutter eg-totals-label" style="left:0;width:${this.gutter}px">Σ</div>`);
    let stickyLeft = this.gutter;
    const base = this.gutter + g.pinnedWidth();
    for (const c of g.visibleColumns()) {
      const sticky = c.pinned === 'left';
      const left = sticky ? stickyLeft : base + c.left;
      if (sticky) stickyLeft += c.width;
      const v = aggs[c.id];
      const fnLabel = c.def.aggFunc ? c.def.aggFunc + ': ' : '';
      const display =
        v == null
          ? ''
          : `<span class="eg-agg-fn">${fnLabel}</span>` +
            v.toLocaleString(g.options.locale, { maximumFractionDigits: 2 });
      parts.push(
        `<div class="eg-cell eg-type-number${sticky ? ' eg-pinned' : ''}" role="gridcell" style="${sticky ? 'position:sticky;' : ''}left:${left}px;width:${c.width}px" title="${c.def.aggFunc ? esc(g.t('totalsTitle', { fn: c.def.aggFunc, name: c.header, n: count })) : ''}">${display}</div>`,
      );
    }
    row.innerHTML = parts.join('');
    this.viewport.appendChild(row);
    this.totalsRow = row;
  }

  invalidateRows(indices: number[]): void {
    indices.forEach((i) => this.dirtyRows.add(i));
    this.schedule();
  }

  private visibleRowRange(): [number, number] {
    const top = this.viewport.scrollTop;
    const first = Math.max(0, this.rowAtY(top) - OVERSCAN);
    const last = Math.min(
      this.grid.rowCount - 1,
      this.rowAtY(top + this.viewport.clientHeight) + OVERSCAN,
    );
    return [first, last];
  }

  private visibleColRange(): [number, number] {
    const cols = this.grid.centerColumns();
    const left = this.viewport.scrollLeft;
    const width = this.viewport.clientWidth - this.gutter - this.grid.pinnedWidth();
    let first = 0;
    while (first < cols.length - 1 && cols[first].left + cols[first].width < left) first++;
    let last = first;
    while (last < cols.length - 1 && cols[last].left < left + width) last++;
    return [Math.max(0, first - 2), Math.min(cols.length - 1, last + 2)];
  }

  private render(): void {
    const [first, last] = this.visibleRowRange();
    const colWindow = this.visibleColRange();
    const colsChanged = colWindow[0] !== this.colWindow[0] || colWindow[1] !== this.colWindow[1];
    this.colWindow = colWindow;

    for (const [i, rowEl] of this.rowPool) {
      if (i < first || i > last) {
        this.disposeRow(rowEl);
        this.rowPool.delete(i);
      }
    }
    for (const [i, d] of this.detailPool) {
      if (i < first || i > last || !this.grid.isExpanded(i)) {
        this.disposeDetail(d);
        this.detailPool.delete(i);
      }
    }

    for (let i = first; i <= last; i++) {
      let rowEl = this.rowPool.get(i);
      if (!rowEl) {
        rowEl = el('div', 'eg-row');
        rowEl.setAttribute('role', 'row');
        this.canvas.appendChild(rowEl);
        this.rowPool.set(i, rowEl);
        this.renderRow(i, rowEl);
      } else if (colsChanged || this.dirtyRows.has(i)) {
        this.renderRow(i, rowEl);
      }
      if (this.grid.isExpanded(i) && this.grid.hasDetail(i) && !this.detailPool.has(i)) this.renderDetail(i);
    }
    this.dirtyRows.clear();

    if (first !== this.lastViewport[0] || last !== this.lastViewport[1]) {
      this.lastViewport = [first, last];
      this.grid.events.emit('viewportChanged', { firstRow: first, lastRow: last });
    }
    this.renderSelection();
  }

  /* ---------------------------------------------------------------- */
  /* Row rendering — fast innerHTML path / DOM path for custom cells   */
  /* ---------------------------------------------------------------- */

  private disposeRow(rowEl: HTMLElement): void {
    this.destroyers.get(rowEl)?.forEach((d) => d());
    this.destroyers.delete(rowEl);
    rowEl.remove();
  }

  private renderRow(i: number, rowEl: HTMLElement): void {
    const g = this.grid;
    const h = g.rowHeightOf(i);
    this.destroyers.get(rowEl)?.forEach((d) => d());
    this.destroyers.delete(rowEl);
    rowEl.style.cssText = `top:${this.rowTop(i)}px;height:${h}px;width:${this.totalWidth()}px;`;
    rowEl.dataset.rowIndex = String(i);
    rowEl.setAttribute('aria-rowindex', String(i + 2));
    if (g.treeMode) rowEl.setAttribute('aria-level', String(g.treeLevel(i) + 1));
    rowEl.classList.toggle('eg-row--dirty', g.isRowDirty(i));
    rowEl.classList.remove('eg-group-row');
    rowEl.classList.toggle('eg-row--loading', g.isRowLoading(i));

    const group = g.getGroupRow(i);
    if (group) {
      rowEl.classList.add('eg-group-row');
      const cols = g.visibleColumns();
      const aggCells = cols
        .map((c, vi) => {
          const v = group.aggs[c.id];
          if (v == null) return '';
          const left = this.gutter + g.pinnedWidth() + (c.pinned ? 0 : c.left);
          const display = g.options.locale
            ? v.toLocaleString(g.options.locale, { maximumFractionDigits: 2 })
            : String(v);
          return `<div class="eg-cell eg-type-number eg-agg" role="gridcell" style="left:${c.pinned ? c.left + this.gutter : left}px;width:${c.width}px" data-row="${i}" data-col="${vi}">${display}</div>`;
        })
        .join('');
      rowEl.innerHTML =
        `<div class="eg-cell eg-gutter" role="rowheader" style="left:0;width:${this.gutter}px">` +
        `<button class="eg-expand eg-group-toggle" aria-expanded="${!group.collapsed}" aria-label="${esc(g.t('group', { name: group.key }))}" tabindex="-1">${group.collapsed ? '▸' : '▾'}</button></div>` +
        `<div class="eg-group-label" style="left:${this.gutter}px" data-row="${i}" data-col="0">` +
        `${'&nbsp;'.repeat(group.level * 4)}<b>${esc(group.key)}</b>&nbsp;<span class="eg-group-count">(${group.count})</span></div>` +
        aggCells;
      return;
    }
    if (this.hasCustomRenderers) this.renderRowDom(i, rowEl);
    else this.renderRowHtml(i, rowEl);
  }

  private gutterHtml(i: number): string {
    const g = this.grid;
    const tree = g.treeMode && g.treeHasChildren(i);
    const expandable = tree || g.hasDetail(i);
    const chevron = expandable
      ? `<button class="eg-expand${tree ? ' eg-tree-toggle' : ''}" aria-expanded="${g.isExpanded(i)}" aria-label="${esc(g.t('expandRow', { n: i + 1 }))}" tabindex="-1">${g.isExpanded(i) ? '▾' : '▸'}</button>`
      : '';
    return `<div class="eg-cell eg-gutter" role="rowheader" style="left:0;width:${this.gutter}px">${chevron}<span>${i + 1}</span></div>`;
  }

  private renderRowHtml(i: number, rowEl: HTMLElement): void {
    const g = this.grid;
    const pinned = g.pinnedColumns();
    const center = g.centerColumns();
    const [c0, c1] = this.colWindow;
    const parts: string[] = [];
    if (this.gutter) parts.push(this.gutterHtml(i));
    let stickyLeft = this.gutter;
    for (const c of pinned) {
      parts.push(this.cellHtml(i, g.visibleIndexOf(c.id), c, stickyLeft, true));
      stickyLeft += c.width;
    }
    const base = this.gutter + g.pinnedWidth();
    for (let ci = c0; ci <= c1 && ci < center.length; ci++) {
      const c = center[ci];
      parts.push(this.cellHtml(i, pinned.length + ci, c, base + c.left, false));
    }
    rowEl.innerHTML = parts.join('');
  }

  private cellMeta(row: number, col: ColumnState<T>) {
    const g = this.grid;
    const dataRow = g.getRowByViewIndex(row);
    const extra = (dataRow && col.def.cellClass?.(g.getValue(row, col.id), dataRow)) || '';
    return {
      cls:
        `eg-cell eg-type-${col.type}` +
        (col.pinned ? ' eg-pinned' : '') +
        (col.def.wrapText ? ' eg-wrap' : '') +
        (extra ? ' ' + extra : '') +
        (g.isCellEditable(row, col.id) ? '' : ' eg-readonly'),
      dataRow,
    };
  }

  private cellHtml(
    row: number,
    visibleCol: number,
    col: ColumnState<T>,
    left: number,
    sticky: boolean,
  ): string {
    const { cls } = this.cellMeta(row, col);
    const g = this.grid;
    let indent = '';
    if (g.treeMode && visibleCol === 0) {
      const px = g.treeLevel(row) * (g.options.treeData?.indent ?? 18);
      if (px) indent = `padding-left:${8 + px}px;`;
    }
    const style =
      (sticky ? `position:sticky;left:${left}px;` : `left:${left}px;`) + `width:${col.width}px;${indent}`;
    const content = g.isRowLoading(row)
      ? esc(g.t('loadingCell'))
      : esc(g.getDisplayValue(row, col.id));
    return `<div class="${cls}" role="gridcell" aria-colindex="${visibleCol + 1}" style="${style}" data-row="${row}" data-col="${visibleCol}">${content}</div>`;
  }

  /** DOM path: used when any column has a custom renderer. */
  private renderRowDom(i: number, rowEl: HTMLElement): void {
    const g = this.grid;
    rowEl.innerHTML = '';
    const frag = document.createDocumentFragment();
    if (this.gutter) {
      const tmp = document.createElement('div');
      tmp.innerHTML = this.gutterHtml(i);
      frag.appendChild(tmp.firstElementChild!);
    }
    let stickyLeft = this.gutter;
    const base = this.gutter + g.pinnedWidth();
    const center = g.centerColumns();
    const [c0, c1] = this.colWindow;
    const make = (col: ColumnState<T>, visibleCol: number, left: number, sticky: boolean) => {
      const cell = document.createElement('div');
      const { cls, dataRow } = this.cellMeta(i, col);
      cell.className = cls;
      cell.setAttribute('role', 'gridcell');
      cell.setAttribute('aria-colindex', String(visibleCol + 1));
      cell.style.cssText = (sticky ? 'position:sticky;' : '') + `left:${left}px;width:${col.width}px`;
      cell.dataset.row = String(i);
      cell.dataset.col = String(visibleCol);
      const display = g.getDisplayValue(i, col.id);
      if (col.def.cellRenderer && dataRow) {
        this.applyRendered(
          cell,
          rowEl,
          col.def.cellRenderer({
            value: g.getValue(i, col.id),
            displayValue: display,
            row: dataRow,
            rowIndex: i,
            colId: col.id,
          }),
        );
      } else cell.textContent = display;
      frag.appendChild(cell);
    };
    for (const c of g.pinnedColumns()) {
      make(c, g.visibleIndexOf(c.id), stickyLeft, true);
      stickyLeft += c.width;
    }
    for (let ci = c0; ci <= c1 && ci < center.length; ci++) {
      make(center[ci], g.pinnedColumns().length + ci, base + center[ci].left, false);
    }
    rowEl.appendChild(frag);
  }

  /** Mount a renderer result: string | Node | {el,destroy} | array (lines). */
  private applyRendered(cell: HTMLElement, rowEl: HTMLElement, result: CellRendererResult): void {
    const parts: CellContent[] = Array.isArray(result) ? result : [result];
    const multi = parts.length > 1;
    if (multi) cell.classList.add('eg-multi');
    for (const part of parts) {
      const line = multi ? cell.appendChild(el('div', 'eg-line')) : cell;
      if (typeof part === 'string') line.appendChild(document.createTextNode(part));
      else if (part instanceof Node) line.appendChild(part);
      else {
        line.appendChild(part.el);
        if (part.destroy) {
          let list = this.destroyers.get(rowEl);
          if (!list) this.destroyers.set(rowEl, (list = []));
          list.push(part.destroy);
        }
      }
    }
  }

  /* ---------------------------------------------------------------- */
  /* Detail panels (nested sub-tables)                                 */
  /* ---------------------------------------------------------------- */

  private renderDetail(i: number): void {
    const g = this.grid;
    const md = g.options.masterDetail!;
    const layout = g.detailLayout(i);
    const panel = el('div', 'eg-detail');
    panel.style.cssText = `top:${this.rowTop(i) + g.rowHeightOf(i)}px;height:${layout.height}px;width:${this.totalWidth()}px;`;
    panel.setAttribute('role', 'row');
    const inner = el('div', 'eg-detail-inner');
    panel.appendChild(inner);
    this.canvas.appendChild(panel);

    // ---- Pannello custom: qualsiasi HTML/CSS/elemento, con lifecycle ----
    if (md.detailRenderer) {
      const row = g.getRowByViewIndex(i)!;
      const host = el('div', 'eg-detail-custom');
      inner.classList.add('eg-detail-inner--custom');
      inner.appendChild(host);
      const result = md.detailRenderer({ row, rowIndex: i, rowId: g.getRowIdByViewIndex(i) });
      let destroy: (() => void) | undefined;
      if (typeof result === 'string') host.textContent = result;
      else if (result instanceof Node) host.appendChild(result);
      else {
        host.appendChild(result.el);
        destroy = result.destroy;
      }
      this.detailPool.set(i, { el: panel, grid: null, destroy });
      return;
    }

    // ---- Griglia annidata (con ricorsione se il JSON la contiene) ----
    const rows = g.getDetailRows(i);
    if (!rows) return;
    if (layout.innerWidth > 0) inner.style.width = layout.innerWidth + 'px';
    const nestedField = g.getNestedDetailField(rows);
    const GridCtor = this.grid.constructor as new (c: HTMLElement, o: unknown) => Grid;
    const nested = new GridCtor(inner, {
      columns: g.getDetailColumns(rows),
      data: rows,
      theme: g.options.theme,
      locale: g.options.locale,
      rowHeight: Math.max(26, g.options.rowHeight - 4),
      headerHeight: 30,
      showRowNumbers: false,
      masterDetail: nestedField ? { field: nestedField, height: g.detailHeight } : undefined,
      ariaLabel: g.t('detailOf', { n: i + 1 }),
    });
    // Quando tutto entra nel pannello, l'overflow è spento: nessuna
    // scrollbar può comparire per arrotondamenti di un pixel.
    if (layout.fits) nested['renderer'].root.classList.add('eg-fit');
    // Bubble nested edits into the host grid's dirty tracking.
    nested.events.on('cellsChanged', () => g.markRowDirty(g.getRowIdByViewIndex(i)));
    this.detailPool.set(i, { el: panel, grid: nested });
  }

  private disposeDetail(d: DetailEntry): void {
    d.destroy?.();
    d.grid?.destroy();
    d.el.remove();
  }

  /* ---------------------------------------------------------------- */
  /* Header                                                            */
  /* ---------------------------------------------------------------- */

  private buildHeader(): void {
    const g = this.grid;
    this.headerRow.style.height = g.options.headerHeight + 'px';
    const parts: string[] = [];
    if (this.gutter) {
      parts.push(
        `<div class="eg-hcell eg-gutter eg-corner" role="columnheader" aria-label="${esc(g.t('rowNumber'))}" style="left:0;width:${this.gutter}px"></div>`,
      );
    }
    let stickyLeft = this.gutter;
    const base = this.gutter + g.pinnedWidth();
    for (const c of g.visibleColumns()) {
      const sticky = c.pinned === 'left';
      const left = sticky ? stickyLeft : base + c.left;
      if (sticky) stickyLeft += c.width;
      const sort = g.sortModel.find((s) => s.colId === c.id);
      const ariaSort = sort ? (sort.dir === 'asc' ? 'ascending' : 'descending') : 'none';
      const sortIcon = sort ? (sort.dir === 'asc' ? '▲' : '▼') : '';
      const filtered = g.filterModel.some((f) => f.colId === c.id);
      parts.push(
        `<div class="eg-hcell${sticky ? ' eg-pinned' : ''}" role="columnheader" aria-sort="${ariaSort}" data-colid="${esc(c.id)}"
           style="${sticky ? 'position:sticky;' : ''}left:${left}px;width:${c.width}px">
           <span class="eg-htext">${esc(c.header)}</span>
           <span class="eg-sort" aria-hidden="true">${sortIcon}</span>
           <button class="eg-filter-btn${filtered ? ' eg-filter-on' : ''}" data-filter="${esc(c.id)}" aria-label="${esc(g.t('filterLabel', { name: c.header }))}" aria-haspopup="dialog">⏷</button>
           ${c.resizable ? `<span class="eg-resize" data-resize="${esc(c.id)}" aria-hidden="true"></span>` : ''}
         </div>`,
      );
    }
    this.headerRow.innerHTML = parts.join('');
  }

  /* ---------------------------------------------------------------- */
  /* Selection painting (class-based, visible cells only)              */
  /* ---------------------------------------------------------------- */

  copyRange: CellRange | null = null;
  private invalidTimers = new Map<string, ReturnType<typeof setTimeout>>();

  /** Visual inline validation: red outline + tooltip for a few seconds. */
  flashInvalid(errors: { cell: { rowIndex: number; colId: string }; message: string }[]): void {
    for (const e of errors) {
      const colIdx = this.grid.visibleIndexOf(e.cell.colId);
      const host = this.canvas.querySelector<HTMLElement>(
        `.eg-cell[data-row="${e.cell.rowIndex}"][data-col="${colIdx}"]`,
      );
      if (!host) continue;
      host.classList.add('eg-invalid');
      host.title = e.message;
      const key = e.cell.rowIndex + ':' + e.cell.colId;
      clearTimeout(this.invalidTimers.get(key));
      this.invalidTimers.set(
        key,
        setTimeout(() => {
          host.classList.remove('eg-invalid');
          host.removeAttribute('title');
          this.invalidTimers.delete(key);
        }, 2600),
      );
    }
  }

  renderSelection(): void {
    const g = this.grid;
    const active = g.activeCell;
    const activeCol = active ? g.visibleIndexOf(active.colId) : -1;
    const lastRange = g.ranges[g.ranges.length - 1];

    this.canvas.querySelectorAll<HTMLElement>('.eg-cell[data-row]').forEach((cell) => {
      if (cell.closest('.eg-detail')) return;
      const r = Number(cell.dataset.row);
      const c = Number(cell.dataset.col);
      let sel = false;
      for (const rg of g.ranges) {
        if (r >= rg.startRow && r <= rg.endRow && c >= rg.startCol && c <= rg.endCol) {
          sel = true;
          break;
        }
      }
      const isActive = !!active && r === active.rowIndex && c === activeCol;
      cell.classList.toggle('eg-sel', sel);
      cell.classList.toggle('eg-active', isActive);
      cell.setAttribute('aria-selected', String(sel || isActive));
      if (isActive) {
        cell.id = 'eg-active-cell';
        this.root.setAttribute('aria-activedescendant', 'eg-active-cell');
      } else if (cell.id === 'eg-active-cell') cell.removeAttribute('id');
      const cp = this.copyRange;
      cell.classList.toggle('eg-copy-t', !!cp && r === cp.startRow && c >= cp.startCol && c <= cp.endCol);
      cell.classList.toggle('eg-copy-b', !!cp && r === cp.endRow && c >= cp.startCol && c <= cp.endCol);
      cell.classList.toggle('eg-copy-l', !!cp && c === cp.startCol && r >= cp.startRow && r <= cp.endRow);
      cell.classList.toggle('eg-copy-r', !!cp && c === cp.endCol && r >= cp.startRow && r <= cp.endRow);
    });

    this.fillHandle.remove();
    if (lastRange && this.grid.options.editable) {
      const host = this.canvas.querySelector<HTMLElement>(
        `.eg-cell[data-row="${lastRange.endRow}"][data-col="${lastRange.endCol}"]`,
      );
      if (host && !host.closest('.eg-detail')) host.appendChild(this.fillHandle);
    }

    this.canvas.querySelectorAll<HTMLElement>('.eg-gutter').forEach((gut) => {
      const ri = Number(gut.parentElement?.dataset.rowIndex);
      gut.classList.toggle('eg-hl', g.ranges.some((rg) => ri >= rg.startRow && ri <= rg.endRow));
    });
  }

  ensureCellVisible(rowIndex: number, colIndex: number): void {
    const g = this.grid;
    const vp = this.viewport;
    const headerH = g.options.headerHeight;
    const top = this.rowTop(rowIndex);
    const h = g.rowHeightOf(rowIndex);
    if (top < vp.scrollTop) vp.scrollTop = top;
    else if (top + h > vp.scrollTop + vp.clientHeight - headerH)
      vp.scrollTop = top + h - vp.clientHeight + headerH;

    const col = g.columnByVisibleIndex(colIndex);
    if (!col || col.pinned) return;
    const fixed = this.gutter + g.pinnedWidth();
    if (col.left < vp.scrollLeft) vp.scrollLeft = col.left;
    else if (col.left + col.width > vp.scrollLeft + vp.clientWidth - fixed)
      vp.scrollLeft = col.left + col.width - vp.clientWidth + fixed;
  }

  /* ---------------------------------------------------------------- */
  /* Excel-style filter popup                                          */
  /* ---------------------------------------------------------------- */

  openFilterPopup(colId: string, anchor: HTMLElement): void {
    this.closeFilterPopup();
    const g = this.grid;
    const values = g.getDistinctValues(colId);
    const current = g.filterModel.find((f) => f.colId === colId);
    const selected = new Set(current?.values?.map(String) ?? values);

    const pop = el('div', 'eg-filter-pop');
    pop.setAttribute('role', 'dialog');
    pop.setAttribute('aria-label', g.t('filterDialog'));
    pop.innerHTML = `
      <input class="eg-filter-search" placeholder="${esc(g.t('filterSearch'))}" aria-label="${esc(g.t('filterSearchLabel'))}" />
      <label class="eg-filter-item"><input type="checkbox" class="eg-filter-all" checked> <b>${esc(g.t('filterSelectAll'))}</b></label>
      <div class="eg-filter-list" role="listbox"></div>
      <div class="eg-filter-actions">
        <button class="eg-btn eg-filter-clear">${esc(g.t('filterClear'))}</button>
        <button class="eg-btn eg-btn-primary eg-filter-apply">${esc(g.t('filterOk'))}</button>
      </div>`;
    const list = pop.querySelector('.eg-filter-list')!;
    const renderList = (needle = '') => {
      list.innerHTML = values
        .filter((v) => v.toLowerCase().includes(needle.toLowerCase()))
        .slice(0, 400)
        .map(
          (v) =>
            `<label class="eg-filter-item"><input type="checkbox" data-v="${esc(v)}" ${
              selected.has(v) ? 'checked' : ''
            }> ${esc(v) || `<i>${esc(g.t('filterEmpty'))}</i>`}</label>`,
        )
        .join('');
    };
    renderList();

    pop.addEventListener('change', (e) => {
      const t = e.target as HTMLInputElement;
      if (t.classList.contains('eg-filter-all')) {
        selected.clear();
        if (t.checked) values.forEach((v) => selected.add(v));
        renderList((pop.querySelector('.eg-filter-search') as HTMLInputElement).value);
      } else if (t.dataset.v !== undefined) {
        t.checked ? selected.add(t.dataset.v) : selected.delete(t.dataset.v);
      }
    });
    pop.querySelector('.eg-filter-search')!.addEventListener('input', (e) =>
      renderList((e.target as HTMLInputElement).value),
    );
    pop.querySelector('.eg-filter-apply')!.addEventListener('click', () => {
      const rest = g.filterModel.filter((f) => f.colId !== colId);
      const model =
        selected.size === values.length ? rest : [...rest, { colId, values: [...selected] }];
      g.setFilterModel(model);
      this.closeFilterPopup();
    });
    pop.querySelector('.eg-filter-clear')!.addEventListener('click', () => {
      g.setFilterModel(g.filterModel.filter((f) => f.colId !== colId));
      this.closeFilterPopup();
    });

    const rect = anchor.getBoundingClientRect();
    const rootRect = this.root.getBoundingClientRect();
    pop.style.left = Math.max(4, rect.right - rootRect.left - 230) + 'px';
    pop.style.top = rect.bottom - rootRect.top + 2 + 'px';
    this.root.appendChild(pop);
    (pop.querySelector('.eg-filter-search') as HTMLElement).focus();
    this.filterPopup = pop;
    setTimeout(() => document.addEventListener('mousedown', this.outsideClose, { capture: true }));
  }

  private outsideClose = (e: MouseEvent): void => {
    if (this.filterPopup && !this.filterPopup.contains(e.target as Node)) this.closeFilterPopup();
  };

  closeFilterPopup(): void {
    this.filterPopup?.remove();
    this.filterPopup = null;
    document.removeEventListener('mousedown', this.outsideClose, { capture: true });
  }
}

function el(tag: string, className: string): HTMLElement {
  const e = document.createElement(tag);
  e.className = className;
  return e;
}
