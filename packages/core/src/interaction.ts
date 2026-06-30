/**
 * InteractionManager — every mouse/keyboard behaviour, tuned to match
 * Excel exactly (zero learning curve is the product requirement).
 *
 * Implements: range selection (drag / Shift / Ctrl), full keyboard
 * navigation (arrows, Ctrl+arrows, Tab/Enter cycles, Home/End, PageUp/Down,
 * Ctrl+A), editing (F2 / double-click / type-to-edit, Esc, Ctrl+Enter
 * multi-fill, Alt+Enter newline), clipboard TSV interop with Excel
 * (copy / cut / paste with pattern repetition + marching ants), the fill
 * handle with series detection, column resize and sort.
 */
import { normalizeRange, type Grid } from './grid';
import type { Renderer } from './renderer';
import { makeSeries } from './values';
import { ContextMenu, FindPanel } from './ui';
import type { CellChange, CellRange, CellRef, RowData } from './types';

export class InteractionManager<T extends RowData> {
  private editor: HTMLTextAreaElement | HTMLInputElement | HTMLSelectElement | null = null;
  private editorDatalist: HTMLElement | null = null;
  private editorOptions: string[] | null = null;
  private editCell: CellRef | null = null;
  private anchor: { row: number; col: number } | null = null;
  private dragMode: 'select' | 'fill' | 'resize' | 'header' | null = null;
  private headerDrag: {
    colId: string;
    startX: number;
    active: boolean;
    ghost: HTMLElement | null;
    marker: HTMLElement | null;
    shiftKey: boolean;
  } | null = null;
  private resizeCol: { id: string; startX: number; startW: number } | null = null;
  private fillTarget: CellRange | null = null;
  private cutPending = false;
  private disposers: (() => void)[] = [];
  readonly find: FindPanel<T>;
  readonly menu: ContextMenu<T>;

  constructor(private grid: Grid<T>, private renderer: Renderer<T>) {
    this.find = new FindPanel(grid, renderer);
    this.menu = new ContextMenu(grid, renderer, {
      copy: () => this.copyToClipboard(false),
      cut: () => this.copyToClipboard(true),
      paste: async () => {
        try {
          const text = await navigator.clipboard.readText();
          if (text) this.pasteText(text);
        } catch {
          /* permesso negato: l'utente può usare Ctrl+V */
        }
      },
    });
    const vp = renderer.viewport;
    const listen = <K extends keyof HTMLElementEventMap>(
      target: HTMLElement | Document,
      type: K,
      fn: (e: HTMLElementEventMap[K]) => void,
      opts?: AddEventListenerOptions,
    ) => {
      target.addEventListener(type, fn as EventListener, opts);
      this.disposers.push(() => target.removeEventListener(type, fn as EventListener, opts));
    };

    renderer.root.tabIndex = 0;
    listen(vp, 'pointerdown' as keyof HTMLElementEventMap, (e) => this.onPointerDown(e as PointerEvent));
    listen(document, 'pointermove' as keyof HTMLElementEventMap, (e) => this.onPointerMove(e as PointerEvent));
    listen(document, 'pointerup' as keyof HTMLElementEventMap, (e) => this.onPointerUp(e as PointerEvent));
    listen(document, 'pointercancel' as keyof HTMLElementEventMap, () => this.cancelTouchGesture());
    listen(vp, 'dblclick', (e) => this.onDoubleClick(e));
    listen(renderer.root, 'keydown', (e) => this.onKeyDown(e));
    listen(renderer.root, 'copy' as keyof HTMLElementEventMap, (e) => this.onCopy(e as ClipboardEvent));
    listen(renderer.root, 'cut' as keyof HTMLElementEventMap, (e) => this.onCut(e as ClipboardEvent));
    listen(renderer.root, 'paste' as keyof HTMLElementEventMap, (e) => this.onPaste(e as ClipboardEvent));
    listen(renderer.root, 'contextmenu' as keyof HTMLElementEventMap, (e) => {
      if ((e.target as HTMLElement).closest('.eg-detail')) return;
      this.menu.open(e as MouseEvent);
    });
  }

  destroy(): void {
    this.find.close();
    this.menu.close();
    this.disposers.forEach((d) => d());
    this.disposers = [];
  }

  /* ---------------------------------------------------------------- */
  /* Mouse                                                             */
  /* ---------------------------------------------------------------- */

  /* ---- Touch gesture state ---- */
  private touch: {
    id: number;
    x: number;
    y: number;
    moved: boolean;
    timer: ReturnType<typeof setTimeout> | null;
    longPressed: boolean;
    selecting: boolean;
    lastTapTime: number;
    lastTapCell: string;
  } = { id: -1, x: 0, y: 0, moved: false, timer: null, longPressed: false, selecting: false, lastTapTime: 0, lastTapCell: '' };

  private onPointerDown(e: PointerEvent): void {
    if (e.pointerType === 'touch' || e.pointerType === 'pen') return this.onTouchStart(e);
    if (e.button === 2) return; // il tasto destro è gestito da contextmenu
    this.onMouseDown(e);
  }
  private onPointerMove(e: PointerEvent): void {
    if (e.pointerType === 'touch' || e.pointerType === 'pen') return this.onTouchMove(e);
    this.onMouseMove(e);
  }
  private onPointerUp(e: PointerEvent): void {
    if (e.pointerType === 'touch' || e.pointerType === 'pen') return this.onTouchEnd(e);
    this.onMouseUp(e);
  }

  /**
   * Modello touch (tablet di cantiere):
   *  - tap: seleziona la cella · doppio tap: modifica
   *  - long-press fermo: menu contestuale
   *  - long-press poi trascina: estende la selezione (lo scroll nativo
   *    resta sul semplice trascinamento)
   *  - il quadratino di riempimento è trascinabile col dito
   */
  private onTouchStart(e: PointerEvent): void {
    const target = e.target as HTMLElement;
    if (target.closest('.eg-detail')) return;
    if (target === this.renderer.fillHandle) {
      e.preventDefault();
      this.dragMode = 'fill';
      this.fillTarget = null;
      this.fillSource = this.grid.ranges[this.grid.ranges.length - 1] ?? null;
      this.renderer.fillHandle.setPointerCapture?.(e.pointerId);
      return;
    }
    const t = this.touch;
    t.id = e.pointerId;
    t.x = e.clientX;
    t.y = e.clientY;
    t.moved = false;
    t.longPressed = false;
    t.selecting = false;
    const hit = this.renderer.cellAt(target);
    t.timer = setTimeout(() => {
      t.longPressed = true;
      if (hit) {
        // ancora qui dopo 450ms fermo: prepara estensione o menu
        this.anchor = { row: hit.row, col: hit.col };
        this.grid.selectCell(hit.row, hit.col);
      }
    }, 450);
  }

  private onTouchMove(e: PointerEvent): void {
    const t = this.touch;
    if (this.dragMode === 'fill') {
      e.preventDefault();
      this.onMouseMove(e);
      return;
    }
    if (e.pointerId !== t.id) return;
    const dist = Math.hypot(e.clientX - t.x, e.clientY - t.y);
    if (!t.longPressed) {
      if (dist > 10) this.cancelTouchGesture(); // è uno scroll nativo
      return;
    }
    // long-press già scattato: il trascinamento estende la selezione
    if (dist > 6 || t.selecting) {
      t.selecting = true;
      e.preventDefault();
      const hit = this.renderer.cellAt(document.elementFromPoint(e.clientX, e.clientY));
      if (hit && this.anchor) this.extendTo(hit.row, hit.col, false, true);
    }
  }

  private onTouchEnd(e: PointerEvent): void {
    if (this.dragMode === 'fill') {
      this.onMouseUp();
      return;
    }
    const t = this.touch;
    if (e.pointerId !== t.id) return;
    if (t.timer) clearTimeout(t.timer);
    const target = e.target as HTMLElement;

    if (t.longPressed && !t.selecting) {
      // long-press fermo → menu contestuale
      this.menu.open(
        new MouseEvent('contextmenu', { clientX: e.clientX, clientY: e.clientY, bubbles: true }),
        target,
      );
    } else if (!t.longPressed && !t.moved) {
      // tap semplice: gestisci expand/gruppi/celle come un click
      const expand = target.closest('.eg-expand') as HTMLElement | null;
      if (expand) {
        const row = Number((expand.closest('.eg-row') as HTMLElement).dataset.rowIndex);
        this.grid.treeMode && this.grid.treeHasChildren(row)
          ? this.grid.toggleTreeNode(row)
          : this.grid.isGroupRow(row)
            ? this.grid.toggleGroup(row)
            : this.grid.toggleDetail(row);
      } else {
        const hit = this.renderer.cellAt(target);
        if (hit) {
          const cellKey = hit.row + ':' + hit.col;
          const now = Date.now();
          if (now - t.lastTapTime < 350 && t.lastTapCell === cellKey) {
            // doppio tap → modifica
            const col = this.grid.columnByVisibleIndex(hit.col)!;
            this.startEdit({ rowIndex: hit.row, colId: col.id }, undefined, false);
            t.lastTapTime = 0;
          } else {
            this.anchor = { row: hit.row, col: hit.col };
            this.grid.selectCell(hit.row, hit.col);
            t.lastTapTime = now;
            t.lastTapCell = cellKey;
          }
        }
      }
    }
    t.id = -1;
    t.longPressed = false;
    t.selecting = false;
  }

  private cancelTouchGesture(): void {
    const t = this.touch;
    if (t.timer) clearTimeout(t.timer);
    t.id = -1;
    t.longPressed = false;
    t.selecting = false;
  }

  private onMouseDown(e: MouseEvent): void {
    const target = e.target as HTMLElement;
    if (target.closest('.eg-detail')) return; // nested grid handles itself

    const expand = target.closest('.eg-expand') as HTMLElement | null;
    if (expand) {
      e.preventDefault();
      const row = Number((expand.closest('.eg-row') as HTMLElement).dataset.rowIndex);
      if (this.grid.treeMode && this.grid.treeHasChildren(row)) this.grid.toggleTreeNode(row);
      else this.grid.isGroupRow(row) ? this.grid.toggleGroup(row) : this.grid.toggleDetail(row);
      return;
    }
    const groupLabel = target.closest('.eg-group-label') as HTMLElement | null;
    if (groupLabel) {
      e.preventDefault();
      this.grid.toggleGroup(Number(groupLabel.dataset.row));
      return;
    }

    if (target === this.renderer.fillHandle) {
      e.preventDefault();
      this.dragMode = 'fill';
      this.fillTarget = null;
      this.fillSource = this.grid.ranges[this.grid.ranges.length - 1] ?? null;
      return;
    }
    const resize = target.closest('[data-resize]') as HTMLElement | null;
    if (resize) {
      e.preventDefault();
      const id = resize.dataset.resize!;
      const col = this.grid.columns.find((c) => c.id === id)!;
      this.resizeCol = { id, startX: e.clientX, startW: col.width };
      this.dragMode = 'resize';
      return;
    }
    const filterBtn = target.closest('[data-filter]') as HTMLElement | null;
    if (filterBtn) {
      e.preventDefault();
      this.renderer.openFilterPopup(filterBtn.dataset.filter!, filterBtn);
      return;
    }
    const hcell = target.closest('.eg-hcell[data-colid]') as HTMLElement | null;
    if (hcell) {
      // Drag & drop dell'header: parte solo oltre 5px di movimento,
      // altrimenti al rilascio resta il click di ordinamento.
      this.headerDrag = {
        colId: hcell.dataset.colid!,
        startX: e.clientX,
        active: false,
        ghost: null,
        marker: null,
        shiftKey: e.shiftKey,
      };
      this.dragMode = 'header';
      e.preventDefault();
      return;
    }

    const hit = this.renderer.cellAt(target);
    if (!hit) {
      // Row-number gutter: select the whole row.
      const gutterRow = target.closest('.eg-row') as HTMLElement | null;
      if (target.classList.contains('eg-gutter') && gutterRow) {
        const row = Number(gutterRow.dataset.rowIndex);
        const lastCol = this.grid.visibleColumns().length - 1;
        this.grid.setSelection(
          [{ startRow: row, endRow: row, startCol: 0, endCol: lastCol }],
          { rowIndex: row, colId: this.grid.visibleColumns()[0].id },
        );
      }
      return;
    }

    this.stopEdit(false);
    this.renderer.root.focus({ preventScroll: true });
    e.preventDefault();

    if (e.shiftKey && this.anchor) {
      this.extendTo(hit.row, hit.col, e.ctrlKey || e.metaKey);
    } else {
      this.anchor = { row: hit.row, col: hit.col };
      const range = { startRow: hit.row, endRow: hit.row, startCol: hit.col, endCol: hit.col };
      const ranges = e.ctrlKey || e.metaKey ? [...this.grid.ranges, range] : [range];
      const col = this.grid.columnByVisibleIndex(hit.col)!;
      this.grid.setSelection(ranges, { rowIndex: hit.row, colId: col.id });
      this.grid.events.emit('cellClicked', {
        cell: { rowIndex: hit.row, colId: col.id },
        originalEvent: e,
      });
    }
    this.dragMode = 'select';
  }

  private onMouseMove(e: MouseEvent): void {
    if (this.dragMode === 'header' && this.headerDrag) {
      const hd = this.headerDrag;
      if (!hd.active && Math.abs(e.clientX - hd.startX) > 5) {
        hd.active = true;
        const col = this.grid.columns.find((c) => c.id === hd.colId)!;
        hd.ghost = document.createElement('div');
        hd.ghost.className = 'eg-drag-ghost';
        hd.ghost.textContent = col.header;
        hd.marker = document.createElement('div');
        hd.marker.className = 'eg-drop-marker';
        this.renderer.root.append(hd.ghost, hd.marker);
      }
      if (hd.active && hd.ghost && hd.marker) {
        const rootRect = this.renderer.root.getBoundingClientRect();
        hd.ghost.style.left = e.clientX - rootRect.left + 10 + 'px';
        hd.ghost.style.top = e.clientY - rootRect.top + 6 + 'px';
        const drop = this.headerDropTarget(e.clientX);
        if (drop) {
          hd.marker.style.left = drop.x - rootRect.left + 'px';
          hd.marker.style.display = 'block';
        } else hd.marker.style.display = 'none';
      }
      return;
    }
    if (this.dragMode === 'resize' && this.resizeCol) {
      this.grid.setColumnWidth(
        this.resizeCol.id,
        this.resizeCol.startW + (e.clientX - this.resizeCol.startX),
      );
      return;
    }
    if (!this.dragMode) return;
    const hit = this.renderer.cellAt(document.elementFromPoint(e.clientX, e.clientY));
    if (!hit) return;

    if (this.dragMode === 'select' && this.anchor) {
      this.extendTo(hit.row, hit.col, false, true);
    } else if (this.dragMode === 'fill') {
      const src = this.fillSource;
      if (!src) return;
      // Excel fills along the dominant axis only.
      const vert = Math.abs(hit.row - src.endRow) >= Math.abs(hit.col - src.endCol);
      this.fillTarget = normalizeRange(
        vert
          ? { startRow: src.startRow, endRow: hit.row, startCol: src.startCol, endCol: src.endCol }
          : { startRow: src.startRow, endRow: src.endRow, startCol: src.startCol, endCol: hit.col },
      );
      this.grid.setSelection([this.fillTarget], this.grid.activeCell);
    }
  }

  private onMouseUp(e?: MouseEvent): void {
    if (this.dragMode === 'header' && this.headerDrag) {
      const hd = this.headerDrag;
      hd.ghost?.remove();
      hd.marker?.remove();
      if (!hd.active) {
        // nessun trascinamento: è un click → ordina
        const col = this.grid.columns.find((c) => c.id === hd.colId);
        if (col?.sortable) this.grid.toggleSort(hd.colId, hd.shiftKey);
      } else if (e) {
        const drop = this.headerDropTarget(e.clientX);
        if (drop) this.grid.moveColumn(hd.colId, drop.index);
      }
      this.headerDrag = null;
      this.dragMode = null;
      return;
    }
    if (this.dragMode === 'fill' && this.fillTarget) this.executeFill(this.fillTarget);
    this.dragMode = null;
    this.resizeCol = null;
    this.fillTarget = null;
    this.fillSource = null;
  }

  private onDoubleClick(e: MouseEvent): void {
    const target = e.target as HTMLElement;
    const resize = target.closest('[data-resize]') as HTMLElement | null;
    if (resize) {
      this.grid.autoSizeColumn(resize.dataset.resize!); // Excel: dblclick border = autofit
      return;
    }
    const hit = this.renderer.cellAt(target);
    if (!hit) return;
    const col = this.grid.columnByVisibleIndex(hit.col)!;
    this.grid.events.emit('cellDoubleClicked', {
      cell: { rowIndex: hit.row, colId: col.id },
      originalEvent: e,
    });
    this.startEdit({ rowIndex: hit.row, colId: col.id }, undefined, false);
  }

  private extendTo(row: number, col: number, keepOthers: boolean, replaceLast = false): void {
    const a = this.anchor!;
    const range = normalizeRange({ startRow: a.row, endRow: row, startCol: a.col, endCol: col });
    const base = keepOthers || replaceLast ? this.grid.ranges.slice(0, -1) : [];
    this.grid.setSelection([...base, range], this.grid.activeCell);
  }

  /* ---------------------------------------------------------------- */
  /* Keyboard                                                          */
  /* ---------------------------------------------------------------- */

  /** Where would the dragged header land? → {index in columns, marker x}. */
  private headerDropTarget(clientX: number): { index: number; x: number } | null {
    const headers = [...this.renderer.root.querySelectorAll<HTMLElement>('.eg-hcell[data-colid]')];
    for (const h of headers) {
      const r = h.getBoundingClientRect();
      if (clientX >= r.left && clientX <= r.right) {
        const colId = h.dataset.colid!;
        const idx = this.grid.columns.findIndex((c) => c.id === colId);
        const before = clientX < (r.left + r.right) / 2;
        return { index: before ? idx : idx + 1, x: before ? r.left : r.right };
      }
    }
    return null;
  }

  private onKeyDown(e: KeyboardEvent): void {
    if ((e.target as HTMLElement).closest?.('.eg-detail')) return; // nested grid
    if (this.editor) {
      this.onEditorKey(e);
      return;
    }
    const g = this.grid;
    const active = g.activeCell;
    const ctrl = e.ctrlKey || e.metaKey;

    if (ctrl && e.key.toLowerCase() === 'z') {
      e.preventDefault();
      e.shiftKey ? g.redo() : g.undo();
      return;
    }
    if (ctrl && e.key.toLowerCase() === 'y') {
      e.preventDefault();
      g.redo();
      return;
    }
    if (ctrl && (e.key.toLowerCase() === 'f' || e.key.toLowerCase() === 'h')) {
      e.preventDefault();
      this.find.open(e.key.toLowerCase() === 'h');
      return;
    }
    if (ctrl && e.key.toLowerCase() === 'a') {
      e.preventDefault();
      g.setSelection(
        [{ startRow: 0, endRow: g.rowCount - 1, startCol: 0, endCol: g.visibleColumns().length - 1 }],
        active,
      );
      return;
    }
    if (!active) return;
    const colIdx = g.visibleIndexOf(active.colId);
    const lastRow = g.rowCount - 1;
    const lastCol = g.visibleColumns().length - 1;

    const move = (row: number, col: number, extend: boolean) => {
      e.preventDefault();
      row = Math.max(0, Math.min(lastRow, row));
      col = Math.max(0, Math.min(lastCol, col));
      if (extend && this.anchor) {
        this.extendTo(row, col, false, true);
        const c = g.columnByVisibleIndex(col)!;
        g.activeCell = { rowIndex: row, colId: c.id };
        this.renderer.renderSelection();
      } else {
        this.anchor = { row, col };
        g.selectCell(row, col);
      }
      this.renderer.ensureCellVisible(row, col);
    };

    switch (e.key) {
      case 'ArrowUp':
        return move(ctrl ? this.jump(active.rowIndex, colIdx, -1, 0) : active.rowIndex - 1, colIdx, e.shiftKey);
      case 'ArrowDown':
        return move(ctrl ? this.jump(active.rowIndex, colIdx, 1, 0) : active.rowIndex + 1, colIdx, e.shiftKey);
      case 'ArrowLeft':
        return move(active.rowIndex, ctrl ? this.jump(active.rowIndex, colIdx, 0, -1) : colIdx - 1, e.shiftKey);
      case 'ArrowRight':
        return move(active.rowIndex, ctrl ? this.jump(active.rowIndex, colIdx, 0, 1) : colIdx + 1, e.shiftKey);
      case 'Home':
        return move(ctrl ? 0 : active.rowIndex, 0, e.shiftKey);
      case 'End':
        return move(ctrl ? lastRow : active.rowIndex, lastCol, e.shiftKey);
      case 'PageDown':
        return move(active.rowIndex + this.pageSize(), colIdx, e.shiftKey);
      case 'PageUp':
        return move(active.rowIndex - this.pageSize(), colIdx, e.shiftKey);
      case 'Tab':
        return this.stepInRange(e, e.shiftKey ? -1 : 1, 'h');
      case 'Enter':
        return this.stepInRange(e, e.shiftKey ? -1 : 1, 'v');
      case 'F2':
        e.preventDefault();
        return void this.startEdit(active, undefined, false);
      case 'Delete':
      case 'Backspace':
        e.preventDefault();
        return this.clearSelection();
      case 'Escape':
        this.renderer.copyRange = null;
        this.cutPending = false;
        this.renderer.renderSelection();
        return;
    }

    // Type-to-edit: any printable character starts overwrite editing.
    if (e.key.length === 1 && !ctrl && !e.altKey) {
      e.preventDefault();
      this.startEdit(active, e.key, true);
    }
  }

  /** Ctrl+arrow: jump to the edge of the data block, like Excel. */
  private jump(row: number, col: number, dr: number, dc: number): number {
    const g = this.grid;
    const colId = g.columnByVisibleIndex(col)!.id;
    const valAt = (r: number, c: number) => {
      const id = g.columnByVisibleIndex(c)?.id ?? colId;
      const v = g.getValue(r, id);
      return v !== null && v !== undefined && v !== '';
    };
    const maxR = g.rowCount - 1;
    const maxC = g.visibleColumns().length - 1;
    let r = row;
    let c = col;
    const inBounds = () => r + dr >= 0 && r + dr <= maxR && c + dc >= 0 && c + dc <= maxC;
    if (!inBounds()) return dr ? r : c;
    const startFilled = valAt(r, c);
    const nextFilled = valAt(r + dr, c + dc);
    if (startFilled && nextFilled) {
      while (inBounds() && valAt(r + dr, c + dc)) (r += dr), (c += dc);
    } else {
      while (inBounds() && !valAt(r + dr, c + dc)) (r += dr), (c += dc);
      if (inBounds()) (r += dr), (c += dc);
    }
    return dr ? r : c;
  }

  private pageSize(): number {
    return Math.max(1, Math.floor(this.renderer.viewport.clientHeight / this.grid.options.rowHeight) - 1);
  }

  /** Tab/Enter cycle inside the selected range when it is multi-cell (Excel). */
  private stepInRange(e: KeyboardEvent, dir: 1 | -1, axis: 'h' | 'v'): void {
    e.preventDefault();
    const g = this.grid;
    const active = g.activeCell;
    if (!active) return;
    const range = g.ranges[g.ranges.length - 1];
    const colIdx = g.visibleIndexOf(active.colId);
    const multi =
      range && (range.endRow > range.startRow || range.endCol > range.startCol) && g.ranges.length === 1;

    if (multi) {
      const rows = range.endRow - range.startRow + 1;
      const cols = range.endCol - range.startCol + 1;
      let r = active.rowIndex - range.startRow;
      let c = colIdx - range.startCol;
      const total = rows * cols;
      const pos = axis === 'h' ? r * cols + c : c * rows + r;
      const next = (pos + dir + total) % total;
      r = axis === 'h' ? Math.floor(next / cols) : next % rows;
      c = axis === 'h' ? next % cols : Math.floor(next / rows);
      const col = g.columnByVisibleIndex(range.startCol + c)!;
      g.activeCell = { rowIndex: range.startRow + r, colId: col.id };
      this.renderer.renderSelection();
      this.renderer.ensureCellVisible(range.startRow + r, range.startCol + c);
    } else {
      const r = axis === 'v' ? active.rowIndex + dir : active.rowIndex;
      const c = axis === 'h' ? colIdx + dir : colIdx;
      this.anchor = { row: r, col: c };
      this.grid.selectCell(
        Math.max(0, Math.min(g.rowCount - 1, r)),
        Math.max(0, Math.min(g.visibleColumns().length - 1, c)),
      );
      this.renderer.ensureCellVisible(r, c);
    }
  }

  private clearSelection(): void {
    const changes: CellChange[] = [];
    this.forEachSelectedCell((row, colId) => {
      if (!this.grid.isCellEditable(row, colId)) return;
      changes.push({
        rowId: this.grid.getRowIdByViewIndex(row),
        rowIndex: row,
        colId,
        oldValue: this.grid.getValue(row, colId),
        newValue: null,
      });
    });
    this.grid.applyChanges(changes, 'user');
  }

  private forEachSelectedCell(fn: (row: number, colId: string) => void): void {
    for (const range of this.grid.ranges) {
      for (let r = range.startRow; r <= range.endRow; r++) {
        for (let c = range.startCol; c <= range.endCol; c++) {
          const col = this.grid.columnByVisibleIndex(c);
          if (col) fn(r, col.id);
        }
      }
    }
  }

  /* ---------------------------------------------------------------- */
  /* Editing                                                           */
  /* ---------------------------------------------------------------- */

  startEdit(cell?: CellRef, initialText?: string, overwrite = false): void {
    const g = this.grid;
    if (!cell || !g.isCellEditable(cell.rowIndex, cell.colId)) return;
    if (g.events.emit('cellEditingStarted', { cell }).defaultPrevented) return;
    this.stopEdit(false);

    const colIdx = g.visibleIndexOf(cell.colId);
    const col = g.columnByVisibleIndex(colIdx)!;
    this.renderer.ensureCellVisible(cell.rowIndex, colIdx);
    this.renderer.viewport.dispatchEvent(new Event('scroll')); // flush render
    const host = this.renderer.viewport.querySelector<HTMLElement>(
      `.eg-cell[data-row="${cell.rowIndex}"][data-col="${colIdx}"]`,
    );
    if (!host) return;

    let editor: HTMLTextAreaElement | HTMLInputElement | HTMLSelectElement;
    if (col.type === 'date') {
      // Editor ricco: date picker nativo (il valore resta ISO yyyy-mm-dd)
      editor = document.createElement('input');
      (editor as HTMLInputElement).type = 'date';
      const raw = g.getValue(cell.rowIndex, cell.colId);
      (editor as HTMLInputElement).value = typeof raw === 'string' ? raw.slice(0, 10) : '';
    } else if (col.type === 'select' && col.def.options && col.def.options.length > 8) {
      // Molte voci → autocomplete con datalist nativa
      editor = document.createElement('input');
      const listId = 'eg-dl-' + Math.random().toString(36).slice(2);
      (editor as HTMLInputElement).setAttribute('list', listId);
      (editor as HTMLInputElement).value = initialText ?? String(g.getValue(cell.rowIndex, cell.colId) ?? '');
      const dl = document.createElement('datalist');
      dl.id = listId;
      dl.innerHTML = col.def.options.map((o) => `<option value="${o}">`).join('');
      this.renderer.root.appendChild(dl);
      this.editorDatalist = dl;
      this.editorOptions = col.def.options;
    } else if (col.type === 'select' && col.def.options) {
      editor = document.createElement('select');
      editor.innerHTML = col.def.options
        .map((o) => `<option ${String(g.getValue(cell.rowIndex, cell.colId)) === o ? 'selected' : ''}>${o}</option>`)
        .join('');
    } else if (col.type === 'boolean') {
      // Boolean: toggle immediately, no editor.
      g.setCellValue(cell.rowIndex, cell.colId, !g.getValue(cell.rowIndex, cell.colId), 'user');
      return;
    } else {
      editor = document.createElement('textarea');
      (editor as HTMLTextAreaElement).value =
        initialText ?? g.getEditValue(cell.rowIndex, cell.colId);
    }
    editor.className = 'eg-editor';
    host.appendChild(editor);
    editor.focus();
    if (!overwrite && editor instanceof HTMLTextAreaElement) {
      editor.selectionStart = editor.selectionEnd = editor.value.length; // F2: caret at end
    }
    this.editor = editor;
    this.editCell = cell;
  }

  private onEditorKey(e: KeyboardEvent): void {
    const ctrl = e.ctrlKey || e.metaKey;
    if (e.key === 'Escape') {
      e.preventDefault();
      this.stopEdit(true);
    } else if (e.key === 'Enter' && e.altKey && this.editor instanceof HTMLTextAreaElement) {
      // Alt+Enter = newline inside the cell, like Excel.
      e.preventDefault();
      const t = this.editor;
      const pos = t.selectionStart;
      t.value = t.value.slice(0, pos) + '\n' + t.value.slice(pos);
      t.selectionStart = t.selectionEnd = pos + 1;
    } else if (e.key === 'Enter' && ctrl) {
      // Ctrl+Enter: write the value into every selected cell.
      e.preventDefault();
      this.commitToSelection();
    } else if (e.key === 'Enter') {
      e.preventDefault();
      this.stopEdit(false);
      this.stepInRange(e, e.shiftKey ? -1 : 1, 'v');
    } else if (e.key === 'Tab') {
      e.preventDefault();
      this.stopEdit(false);
      this.stepInRange(e, e.shiftKey ? -1 : 1, 'h');
    }
    e.stopPropagation();
  }

  stopEdit(cancel: boolean): void {
    if (!this.editor || !this.editCell) return;
    const cell = this.editCell;
    const text = (this.editor as HTMLInputElement).value;
    const options = this.editorOptions;
    this.editor.remove();
    this.editorDatalist?.remove();
    this.editor = null;
    this.editorDatalist = null;
    this.editorOptions = null;
    this.editCell = null;
    if (!cancel) {
      // Autocomplete select: il valore deve appartenere all'elenco
      if (options && text !== '' && !options.includes(text)) {
        const errors = [{ cell, message: this.grid.t('selectNotInList') }];
        this.grid.events.emit('validationFailed', { errors });
        this.renderer.flashInvalid(errors);
      } else {
        const value = this.grid.parseInput(text, cell.rowIndex, cell.colId);
        this.grid.setCellValue(cell.rowIndex, cell.colId, value, 'user');
      }
    }
    this.grid.events.emit('cellEditingStopped', { cell, cancelled: cancel });
    this.renderer.root.focus({ preventScroll: true });
  }

  private commitToSelection(): void {
    if (!this.editor || !this.editCell) return;
    const text = (this.editor as HTMLInputElement).value;
    const cell = this.editCell;
    this.editor.remove();
    this.editor = null;
    this.editCell = null;
    const changes: CellChange[] = [];
    this.forEachSelectedCell((row, colId) => {
      if (!this.grid.isCellEditable(row, colId)) return;
      changes.push({
        rowId: this.grid.getRowIdByViewIndex(row),
        rowIndex: row,
        colId,
        oldValue: this.grid.getValue(row, colId),
        newValue: this.grid.parseInput(text, row, colId),
      });
    });
    this.grid.applyChanges(changes, 'user');
    this.grid.events.emit('cellEditingStopped', { cell, cancelled: false });
    this.renderer.root.focus({ preventScroll: true });
  }

  /* ---------------------------------------------------------------- */
  /* Clipboard — TSV interop with Excel                                */
  /* ---------------------------------------------------------------- */

  /** Copy the active range to the system clipboard (context menu path). */
  copyToClipboard(cut: boolean): void {
    const range = this.grid.ranges[this.grid.ranges.length - 1];
    if (!range) return;
    const tsv = this.rangeToTsv(range);
    const html = this.rangeToHtml(range);
    if (typeof ClipboardItem !== 'undefined' && navigator.clipboard?.write) {
      void navigator.clipboard
        .write([
          new ClipboardItem({
            'text/plain': new Blob([tsv], { type: 'text/plain' }),
            'text/html': new Blob([html], { type: 'text/html' }),
          }),
        ])
        .catch(() => void navigator.clipboard?.writeText(tsv).catch(() => {}));
    } else void navigator.clipboard?.writeText(tsv).catch(() => {});
    this.renderer.copyRange = range;
    this.cutPending = cut;
    this.renderer.renderSelection();
    this.grid.events.emit('copyEnd', { range, tsv });
  }

  /** Range → tabella HTML (per incollare formattato in Outlook/Word/Excel). */
  rangeToHtml(range: CellRange): string {
    const g = this.grid;
    const rows: string[] = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      const cells: string[] = [];
      for (let c = range.startCol; c <= range.endCol; c++) {
        const col = g.columnByVisibleIndex(c);
        const v = col ? g.getDisplayValue(r, col.id) : '';
        const align = col?.type === 'number' ? ' style="text-align:right"' : '';
        cells.push(`<td${align}>${v.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/\n/g, '<br>')}</td>`);
      }
      rows.push('<tr>' + cells.join('') + '</tr>');
    }
    return '<table>' + rows.join('') + '</table>';
  }

  rangeToTsv(range: CellRange): string {
    const lines: string[] = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      const cells: string[] = [];
      for (let c = range.startCol; c <= range.endCol; c++) {
        const col = this.grid.columnByVisibleIndex(c);
        cells.push(col ? this.grid.getDisplayValue(r, col.id) : '');
      }
      lines.push(cells.join('\t'));
    }
    return lines.join('\n');
  }

  private onCopy(e: ClipboardEvent): void {
    if (this.editor) return;
    if ((e.target as HTMLElement)?.closest?.('.eg-detail')) return;
    const range = this.grid.ranges[this.grid.ranges.length - 1];
    if (!range) return;
    if (this.grid.events.emit('copyStart', { range }).defaultPrevented) return;
    e.preventDefault();
    const tsv = this.rangeToTsv(range);
    e.clipboardData?.setData('text/plain', tsv);
    e.clipboardData?.setData('text/html', this.rangeToHtml(range));
    this.renderer.copyRange = range; // marching ants
    this.cutPending = false;
    this.renderer.renderSelection();
    this.grid.events.emit('copyEnd', { range, tsv });
  }

  private onCut(e: ClipboardEvent): void {
    if (this.editor) return;
    if ((e.target as HTMLElement)?.closest?.('.eg-detail')) return;
    const range = this.grid.ranges[this.grid.ranges.length - 1];
    if (!range) return;
    if (this.grid.events.emit('cutStart' as 'copyStart', { range }).defaultPrevented) return;
    this.onCopy(e);
    this.cutPending = true;
  }

  private onPaste(e: ClipboardEvent): void {
    if (this.editor) return;
    if ((e.target as HTMLElement)?.closest?.('.eg-detail')) return;
    e.preventDefault();
    const text = e.clipboardData?.getData('text/plain') ?? '';
    if (text) this.pasteText(text);
  }

  /** Apply clipboard text (TSV) at the current selection — used by Ctrl+V and the context menu. */
  pasteText(text: string): void {
    const g = this.grid;
    const active = g.activeCell;
    if (!active) return;
    const rows = text.replace(/\r/g, '').replace(/\n$/, '').split('\n').map((l) => l.split('\t'));

    const startCol = g.visibleIndexOf(active.colId);
    const sel = g.ranges[g.ranges.length - 1];
    // Excel rule: repeat the pattern to fill the selection when it is a
    // multiple of the copied block; otherwise paste once from the anchor.
    const selRows = sel ? sel.endRow - sel.startRow + 1 : 1;
    const selCols = sel ? sel.endCol - sel.startCol + 1 : 1;
    const repeatR = selRows % rows.length === 0 && selRows > rows.length ? selRows : rows.length;
    const repeatC =
      selCols % rows[0].length === 0 && selCols > rows[0].length ? selCols : rows[0].length;
    const originRow = sel ? sel.startRow : active.rowIndex;
    const originCol = sel ? sel.startCol : startCol;

    const target: CellRange = {
      startRow: originRow,
      endRow: Math.min(g.rowCount - 1, originRow + repeatR - 1),
      startCol: originCol,
      endCol: Math.min(g.visibleColumns().length - 1, originCol + repeatC - 1),
    };
    if (g.events.emit('pasteStart', { rows, target }).defaultPrevented) return;

    const changes: CellChange[] = [];
    const rejected: CellChange[] = [];
    for (let r = 0; r < repeatR; r++) {
      const rowIndex = originRow + r;
      if (rowIndex >= g.rowCount) break;
      for (let c = 0; c < repeatC; c++) {
        const col = g.columnByVisibleIndex(originCol + c);
        if (!col) break;
        const raw = rows[r % rows.length][c % rows[0].length] ?? '';
        const change: CellChange = {
          rowId: g.getRowIdByViewIndex(rowIndex),
          rowIndex,
          colId: col.id,
          oldValue: g.getValue(rowIndex, col.id),
          newValue: g.parseInput(raw, rowIndex, col.id),
        };
        g.isCellEditable(rowIndex, col.id) ? changes.push(change) : rejected.push(change);
      }
    }

    if (this.cutPending && this.renderer.copyRange) {
      const cut = this.renderer.copyRange;
      this.forEachCellInRange(cut, (row, colId) => {
        if (!g.isCellEditable(row, colId)) return;
        changes.push({
          rowId: g.getRowIdByViewIndex(row),
          rowIndex: row,
          colId,
          oldValue: g.getValue(row, colId),
          newValue: null,
        });
      });
      this.cutPending = false;
    }

    const applied = g.applyChanges(changes, 'clipboard');
    this.renderer.copyRange = null;
    g.setSelection([target], active);
    g.events.emit('pasteEnd', { changes: applied, rejected });
  }

  private forEachCellInRange(range: CellRange, fn: (row: number, colId: string) => void): void {
    for (let r = range.startRow; r <= range.endRow; r++)
      for (let c = range.startCol; c <= range.endCol; c++) {
        const col = this.grid.columnByVisibleIndex(c);
        if (col) fn(r, col.id);
      }
  }

  /* ---------------------------------------------------------------- */
  /* Fill handle                                                       */
  /* ---------------------------------------------------------------- */

  private executeFill(target: CellRange): void {
    const g = this.grid;
    const src = this.sourceRangeForFill(target);
    if (!src) return;
    if (g.events.emit('fillStart', { source: src, target }).defaultPrevented) return;

    const changes: CellChange[] = [];
    const vertical = target.endRow > src.endRow || target.startRow < src.startRow;

    for (let c = target.startCol; c <= target.endCol; c++) {
      const col = g.columnByVisibleIndex(c);
      if (!col) continue;
      if (vertical) {
        const sourceValues: unknown[] = [];
        for (let r = src.startRow; r <= src.endRow; r++) sourceValues.push(g.getValue(r, col.id));
        const series = makeSeries(sourceValues, col.type);
        const down = target.endRow > src.endRow;
        const from = down ? src.endRow + 1 : src.startRow - 1;
        const to = down ? target.endRow : target.startRow;
        for (let r = from, i = 0; down ? r <= to : r >= to; r += down ? 1 : -1, i++) {
          if (!g.isCellEditable(r, col.id)) continue;
          changes.push({
            rowId: g.getRowIdByViewIndex(r),
            rowIndex: r,
            colId: col.id,
            oldValue: g.getValue(r, col.id),
            newValue: series(i),
          });
        }
      }
    }
    if (!vertical) {
      // Horizontal fill: repeat each row's source slice across new columns.
      for (let r = target.startRow; r <= target.endRow; r++) {
        const sourceValues: unknown[] = [];
        for (let c = src.startCol; c <= src.endCol; c++) {
          const col = g.columnByVisibleIndex(c);
          sourceValues.push(col ? g.getValue(r, col.id) : null);
        }
        const right = target.endCol > src.endCol;
        const from = right ? src.endCol + 1 : src.startCol - 1;
        const to = right ? target.endCol : target.startCol;
        for (let c = from, i = 0; right ? c <= to : c >= to; c += right ? 1 : -1, i++) {
          const col = g.columnByVisibleIndex(c);
          if (!col || !g.isCellEditable(r, col.id)) continue;
          changes.push({
            rowId: g.getRowIdByViewIndex(r),
            rowIndex: r,
            colId: col.id,
            oldValue: g.getValue(r, col.id),
            newValue: makeSeries(sourceValues, col.type)(i),
          });
        }
      }
    }

    const applied = g.applyChanges(changes, 'fill');
    g.events.emit('fillEnd', { changes: applied });
  }

  /** The original selection captured when the fill drag started. */
  private fillSource: CellRange | null = null;
  private sourceRangeForFill(_target: CellRange): CellRange | null {
    return this.fillSource;
  }
}
