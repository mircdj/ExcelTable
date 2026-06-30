/**
 * @exceltable/angular — Angular wrapper around @exceltable/core.
 *
 * Standalone component (Angular 15+). The grid runs outside Angular's zone
 * for performance — only @Output emissions re-enter the zone, so scrolling
 * a 500k-row grid never triggers change detection.
 *
 * Usage:
 *   <excel-table [columns]="cols" [data]="rows"
 *              (cellValueChanged)="onChange($event)" />
 */
import {
  Component,
  ElementRef,
  EventEmitter,
  Input,
  NgZone,
  OnChanges,
  OnDestroy,
  OnInit,
  Output,
  SimpleChanges,
} from '@angular/core';
import {
  Grid,
  type CellChange,
  type CellRange,
  type ChangeSource,
  type ColumnDef,
  type FilterModelItem,
  type RowData,
  type SortModelItem,
} from '@exceltable/core';

@Component({
  selector: 'excel-table',
  standalone: true,
  template: '',
  styles: [':host { display: block; width: 100%; height: 100%; }'],
})
export class ExcelTableComponent<T extends RowData = RowData>
  implements OnInit, OnChanges, OnDestroy
{
  @Input({ required: true }) columns: ColumnDef<T>[] = [];
  @Input() data: T[] = [];
  @Input() rowHeight = 32;
  @Input() headerHeight = 36;
  @Input() theme = 'excel';
  @Input() locale = 'it-IT';
  @Input() showRowNumbers = true;
  @Input() editable = true;
  @Input() getRowId?: (row: T) => string;
  /**
   * Passthrough for every other GridOption (masterDetail, formulas,
   * treeData, serverSide, strings, contextMenuItems, getRowHeight, …).
   * Applied at construction time.
   */
  @Input() options?: Partial<import('@exceltable/core').GridOptions<T>>;

  @Output() gridReady = new EventEmitter<Grid<T>>();
  @Output() cellValueChanged = new EventEmitter<CellChange & { source: ChangeSource }>();
  @Output() cellsChanged = new EventEmitter<{ changes: CellChange[]; source: ChangeSource }>();
  @Output() selectionChanged = new EventEmitter<CellRange[]>();
  @Output() sortChanged = new EventEmitter<SortModelItem[]>();
  @Output() filterChanged = new EventEmitter<FilterModelItem[]>();
  @Output() dirtyStateChanged = new EventEmitter<string[]>();

  /** The underlying Grid instance — full imperative API. */
  api: Grid<T> | null = null;

  constructor(private host: ElementRef<HTMLElement>, private zone: NgZone) {}

  ngOnInit(): void {
    this.zone.runOutsideAngular(() => {
      const grid = new Grid<T>(this.host.nativeElement, {
        ...this.options,
        columns: this.columns,
        data: this.data,
        rowHeight: this.rowHeight,
        headerHeight: this.headerHeight,
        theme: this.theme,
        locale: this.locale,
        showRowNumbers: this.showRowNumbers,
        editable: this.editable,
        getRowId: this.getRowId,
      });
      this.api = grid;

      const inZone = <E>(emitter: EventEmitter<E>) => (payload: E) =>
        this.zone.run(() => emitter.emit(payload));

      grid.events.on('cellValueChanged', inZone(this.cellValueChanged));
      grid.events.on('cellsChanged', (e) =>
        this.zone.run(() => this.cellsChanged.emit({ changes: e.changes, source: e.source })),
      );
      grid.events.on('selectionChanged', (e) =>
        this.zone.run(() => this.selectionChanged.emit(e.ranges)),
      );
      grid.events.on('sortChanged', (e) => this.zone.run(() => this.sortChanged.emit(e.sortModel)));
      grid.events.on('filterChanged', (e) =>
        this.zone.run(() => this.filterChanged.emit(e.filterModel)),
      );
      grid.events.on('dirtyStateChanged', (e) =>
        this.zone.run(() => this.dirtyStateChanged.emit(e.dirtyRowIds)),
      );
      this.zone.run(() => this.gridReady.emit(grid));
    });
  }

  ngOnChanges(changes: SimpleChanges): void {
    if (!this.api) return;
    if (changes['data'] && !changes['data'].firstChange) this.api.setRowData(this.data);
    if (changes['columns'] && !changes['columns'].firstChange) {
      this.api.setColumns(this.columns);
      this.api.refresh();
    }
  }

  ngOnDestroy(): void {
    this.api?.destroy();
    this.api = null;
  }
}

/* ------------------------------------------------------------------ */
/* templateCell — render an Angular TemplateRef inside a grid cell      */
/* ------------------------------------------------------------------ */
import { TemplateRef, ViewContainerRef } from '@angular/core';
import type { CellRenderer, CellRendererParams } from '@exceltable/core';

/**
 * Wrap an Angular `<ng-template>` as a grid cellRenderer. The embedded view
 * is destroyed automatically when the virtualizer recycles the row.
 *
 *   <ng-template #statoTpl let-p>
 *     <app-stato-badge [value]="p.value" />
 *   </ng-template>
 *
 *   cols = [{ field: 'stato', cellRenderer: templateCell(this.statoTpl, this.vcr) }];
 *
 * The template context exposes the CellRendererParams as `$implicit` (let-p).
 */
export function templateCell<T extends RowData = RowData>(
  tpl: TemplateRef<{ $implicit: CellRendererParams<T> }>,
  vcr: ViewContainerRef,
): CellRenderer<T> {
  return (params) => {
    const view = vcr.createEmbeddedView(tpl, { $implicit: params });
    view.detectChanges();
    const el = document.createElement('div');
    el.style.display = 'contents';
    for (const node of view.rootNodes) el.appendChild(node);
    return { el, destroy: () => view.destroy() };
  };
}
