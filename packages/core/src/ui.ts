/**
 * UI overlays: the Find & Replace panel (Ctrl+F / Ctrl+H) and the
 * right-click context menu. Both operate exclusively through the public
 * Grid API, so every action they perform is observable and undoable.
 */
import type { Grid } from './grid';
import type { Renderer } from './renderer';
import type { CellRef, ContextMenuItem, RowData } from './types';

const esc = (s: string) =>
  s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);

/* ================================================================== */
/* Find & Replace                                                      */
/* ================================================================== */

export class FindPanel<T extends RowData> {
  private el: HTMLElement | null = null;
  private results: CellRef[] = [];
  private index = -1;
  private query = '';
  private matchCase = false;

  constructor(private grid: Grid<T>, private renderer: Renderer<T>) {}

  get isOpen(): boolean {
    return !!this.el;
  }

  open(withReplace: boolean): void {
    if (this.el) {
      this.el.classList.toggle('eg-find--replace', withReplace);
      (this.el.querySelector('.eg-find-q') as HTMLInputElement).focus();
      return;
    }
    const panel = document.createElement('div');
    panel.className = 'eg-find' + (withReplace ? ' eg-find--replace' : '');
    panel.setAttribute('role', 'dialog');
    panel.setAttribute('aria-label', this.grid.t('findTitle'));
    panel.innerHTML = `
      <div class="eg-find-row">
        <input class="eg-find-q" placeholder="${this.grid.t('findPlaceholder')}" aria-label="${this.grid.t('findPlaceholder')}">
        <span class="eg-find-count" aria-live="polite">0 / 0</span>
        <button class="eg-btn" data-f="prev" aria-label="${this.grid.t('findPrev')}">↑</button>
        <button class="eg-btn" data-f="next" aria-label="${this.grid.t('findNext')}">↓</button>
        <label class="eg-find-case" title="${this.grid.t('findCase')}"><input type="checkbox" data-f="case"> Aa</label>
        <button class="eg-btn" data-f="close" aria-label="${this.grid.t('findClose')}">✕</button>
      </div>
      <div class="eg-find-row eg-find-replace-row">
        <input class="eg-find-r" placeholder="${this.grid.t('findReplacePlaceholder')}" aria-label="${this.grid.t('findReplacePlaceholder')}">
        <button class="eg-btn" data-f="replace">${this.grid.t('findReplace')}</button>
        <button class="eg-btn eg-btn-primary" data-f="replaceAll">${this.grid.t('findReplaceAll')}</button>
      </div>`;
    this.renderer.root.appendChild(panel);
    this.el = panel;

    const q = panel.querySelector('.eg-find-q') as HTMLInputElement;
    const r = panel.querySelector('.eg-find-r') as HTMLInputElement;
    q.focus();

    q.addEventListener('input', () => this.search(q.value));
    panel.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Escape') this.close();
      else if (e.key === 'Enter' && e.target === q) this.step(e.shiftKey ? -1 : 1);
      else if (e.key === 'Enter' && e.target === r) this.replaceCurrent(r.value);
    });
    panel.addEventListener('click', (e) => {
      const f = (e.target as HTMLElement).closest('[data-f]') as HTMLElement | null;
      if (!f) return;
      switch (f.dataset.f) {
        case 'prev': return this.step(-1);
        case 'next': return this.step(1);
        case 'close': return this.close();
        case 'replace': return this.replaceCurrent(r.value);
        case 'replaceAll': return this.replaceAll(r.value);
      }
    });
    panel.addEventListener('change', (e) => {
      const t = e.target as HTMLInputElement;
      if (t.dataset.f === 'case') {
        this.matchCase = t.checked;
        this.search(q.value);
      }
    });
    // Re-run on data changes so counts stay honest.
    this.offData = this.grid.events.on('cellsChanged', () => this.search(this.query, true));
  }

  private offData: (() => void) | null = null;

  private search(query: string, keepIndex = false): void {
    this.query = query;
    this.results = this.grid.findAll(query, { matchCase: this.matchCase });
    this.index = this.results.length ? (keepIndex ? Math.min(this.index, this.results.length - 1) : 0) : -1;
    this.updateCount();
    if (this.index >= 0 && !keepIndex) this.goTo(this.index);
    this.grid.events.emit('findResultsChanged', {
      query,
      total: this.results.length,
      current: this.index + 1,
    });
  }

  private step(dir: 1 | -1): void {
    if (!this.results.length) return;
    this.index = (this.index + dir + this.results.length) % this.results.length;
    this.goTo(this.index);
    this.updateCount();
  }

  private goTo(i: number): void {
    const cell = this.results[i];
    const colIdx = this.grid.visibleIndexOf(cell.colId);
    this.grid.navigateTo(cell.rowIndex, colIdx);
  }

  private replaceCurrent(replacement: string): void {
    if (this.index < 0) return;
    const n = this.grid.replaceIn([this.results[this.index]], this.query, replacement, {
      matchCase: this.matchCase,
    });
    if (n) this.search(this.query, true);
  }

  private replaceAll(replacement: string): void {
    if (!this.results.length) return;
    this.grid.replaceIn(this.results, this.query, replacement, { matchCase: this.matchCase });
    this.search(this.query);
  }

  private updateCount(): void {
    const c = this.el?.querySelector('.eg-find-count');
    if (c) c.textContent = `${this.results.length ? this.index + 1 : 0} / ${this.results.length}`;
  }

  close(): void {
    this.offData?.();
    this.offData = null;
    this.el?.remove();
    this.el = null;
    this.results = [];
    this.renderer.root.focus({ preventScroll: true });
  }
}

/* ================================================================== */
/* Context menu                                                        */
/* ================================================================== */

export interface MenuActions {
  copy(): void;
  cut(): void;
  paste(): Promise<void>;
}

export class ContextMenu<T extends RowData> {
  private el: HTMLElement | null = null;

  constructor(
    private grid: Grid<T>,
    private renderer: Renderer<T>,
    private actions: MenuActions,
  ) {}

  open(e: MouseEvent, targetOverride?: HTMLElement): void {
    this.close();
    const g = this.grid;
    const target = targetOverride ?? (e.target as HTMLElement);
    const cellHit = this.renderer.cellAt(target);
    const header = target.closest('.eg-hcell[data-colid]') as HTMLElement | null;
    if (!cellHit && !header) return;
    e.preventDefault();

    let items: ContextMenuItem[] = [];
    if (cellHit) {
      const sel = g.ranges.length ? g.ranges : [];
      const selRows = new Set<number>();
      for (const r of sel) for (let i = r.startRow; i <= r.endRow; i++) selRows.add(i);
      items.push(
        { label: g.t('menuCopy'), key: 'Ctrl+C', action: () => this.actions.copy() },
        { label: g.t('menuCut'), key: 'Ctrl+X', action: () => this.actions.cut() },
        { label: g.t('menuPaste'), key: 'Ctrl+V', action: () => void this.actions.paste() },
        { label: g.t('menuClear'), key: g.t('keyDelete'), action: () => this.clearCells() },
        { sep: true },
        {
          label: g.t('menuInsertAbove'),
          action: () => g.insertRows([{} as T], cellHit.row),
          disabled: g.isGroupRow(cellHit.row),
        },
        { label: g.t('menuInsertBelow'), action: () => g.insertRows([{} as T], cellHit.row + 1) },
        {
          label: selRows.size > 1 ? g.t('menuDeleteRows', { n: selRows.size }) : g.t('menuDeleteRow'),
          action: () => g.removeRows([...selRows]),
          disabled: g.isGroupRow(cellHit.row),
        },
        { sep: true },
      );
      const col = g.columnByVisibleIndex(cellHit.col);
      if (col) {
        items.push(
          { label: g.t('menuAutofit', { name: col.header }), action: () => g.autoSizeColumn(col.id) },
          { label: g.t('menuHide', { name: col.header }), action: () => g.setColumnVisible(col.id, false) },
        );
      }
    } else if (header) {
      const colId = header.dataset.colid!;
      const col = g.columns.find((c) => c.id === colId)!;
      const grouped = g.groupBy.includes(colId);
      items.push(
        { label: g.t('menuSortAsc'), action: () => g.setSortModel([{ colId, dir: 'asc' }]) },
        { label: g.t('menuSortDesc'), action: () => g.setSortModel([{ colId, dir: 'desc' }]) },
        { sep: true },
        {
          label: grouped ? g.t('menuUngroup') : g.t('menuGroupBy', { name: col.header }),
          action: () =>
            g.setGroupBy(grouped ? g.groupBy.filter((c) => c !== colId) : [...g.groupBy, colId]),
        },
        { sep: true },
        { label: g.t('menuAutofit', { name: col.header }), action: () => g.autoSizeColumn(colId) },
        { label: g.t('menuHide', { name: col.header }), action: () => g.setColumnVisible(colId, false) },
      );
      const hidden = g.columns.filter((c) => !c.visible);
      if (hidden.length)
        items.push({
          label: g.t('menuShowHidden', { n: hidden.length }),
          action: () => hidden.forEach((c) => g.setColumnVisible(c.id, true)),
        });
    }

    // Estensibilità: l'app può aggiungere/togliere/riordinare voci.
    if (g.options.contextMenuItems) {
      items = g.options.contextMenuItems(items, {
        cell: cellHit ? { rowIndex: cellHit.row, colId: g.columnByVisibleIndex(cellHit.col)?.id ?? '' } : undefined,
        colId: header?.dataset.colid,
        grid: g,
      });
    }

    const menu = document.createElement('div');
    menu.className = 'eg-menu';
    menu.setAttribute('role', 'menu');
    menu.innerHTML = items
      .map((it, i) =>
        it.sep
          ? '<div class="eg-menu-sep" role="separator"></div>'
          : `<button class="eg-menu-item" role="menuitem" data-i="${i}" ${it.disabled ? 'disabled' : ''}>
               <span>${esc(it.label ?? '')}</span>${it.key ? `<kbd>${esc(it.key)}</kbd>` : ''}
             </button>`,
      )
      .join('');
    menu.addEventListener('click', (ev) => {
      const b = (ev.target as HTMLElement).closest('[data-i]') as HTMLElement | null;
      if (!b) return;
      items[Number(b.dataset.i)].action?.();
      this.close();
    });

    const rootRect = this.renderer.root.getBoundingClientRect();
    menu.style.left = Math.min(e.clientX - rootRect.left, rootRect.width - 240) + 'px';
    menu.style.top = Math.min(e.clientY - rootRect.top, rootRect.height - 40) + 'px';
    this.renderer.root.appendChild(menu);
    // Keep the menu inside the grid vertically.
    const mh = menu.offsetHeight;
    if (e.clientY - rootRect.top + mh > rootRect.height)
      menu.style.top = Math.max(4, rootRect.height - mh - 4) + 'px';
    this.el = menu;
    (menu.querySelector('.eg-menu-item:not([disabled])') as HTMLElement | null)?.focus();
    setTimeout(() => document.addEventListener('mousedown', this.outside, { capture: true }));
    menu.addEventListener('keydown', (ev) => {
      if (ev.key === 'Escape') this.close();
      ev.stopPropagation();
    });
  }

  private clearCells(): void {
    const g = this.grid;
    const changes = [];
    for (const range of g.ranges)
      for (let r = range.startRow; r <= range.endRow; r++)
        for (let c = range.startCol; c <= range.endCol; c++) {
          const col = g.columnByVisibleIndex(c);
          if (!col || !g.isCellEditable(r, col.id)) continue;
          changes.push({
            rowId: g.getRowIdByViewIndex(r),
            rowIndex: r,
            colId: col.id,
            oldValue: g.getValue(r, col.id),
            newValue: null,
          });
        }
    g.applyChanges(changes, 'user');
  }

  private outside = (e: MouseEvent): void => {
    if (this.el && !this.el.contains(e.target as Node)) this.close();
  };

  close(): void {
    this.el?.remove();
    this.el = null;
    document.removeEventListener('mousedown', this.outside, { capture: true });
  }
}
