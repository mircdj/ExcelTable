/**
 * Server-side row model — lazy block loading for datasets too large for
 * the client.
 *
 * The grid renders `totalCount` virtual rows immediately; blocks of rows
 * are requested from `getRows` as they scroll into view and cached with
 * LRU eviction. Sorting and filtering are delegated: changing them purges
 * the cache and refetches with the current models, and a request sequence
 * number guarantees that stale responses (slow network, rapid sort
 * changes) can never overwrite fresh data.
 *
 * In server mode the grid keeps full editing, selection, clipboard and
 * dirty tracking on loaded rows; local grouping, row insert/remove and
 * the fill handle across unloaded rows are not available (the backend
 * owns the dataset).
 */
import type { Grid } from './grid';
import type { FilterModelItem, RowData, SortModelItem } from './types';

export interface ServerSideParams {
  startRow: number;
  /** Exclusive. */
  endRow: number;
  sortModel: SortModelItem[];
  filterModel: FilterModelItem[];
}

export interface ServerSideResult<T extends RowData = RowData> {
  rows: T[];
  totalCount: number;
}

export interface ServerSideOptions<T extends RowData = RowData> {
  getRows(params: ServerSideParams): Promise<ServerSideResult<T>>;
  /** Rows per request. @default 100 */
  blockSize?: number;
  /** Blocks kept in memory before LRU eviction. @default 30 */
  maxBlocks?: number;
}

type BlockState = 'loading' | 'loaded';

export class ServerSource<T extends RowData> {
  private blocks = new Map<number, BlockState>();
  private lru: number[] = [];
  private seq = 0;
  readonly blockSize: number;
  readonly maxBlocks: number;
  /** Requests in flight (observable for tests/metrics). */
  pending = 0;

  constructor(private grid: Grid<T>, private opts: ServerSideOptions<T>) {
    this.blockSize = Math.max(1, opts.blockSize ?? 100);
    this.maxBlocks = Math.max(2, opts.maxBlocks ?? 30);
  }

  /** Make sure every block covering [firstRow, lastRow] is loaded/loading. */
  ensureRange(firstRow: number, lastRow: number): void {
    const first = Math.max(0, Math.floor(firstRow / this.blockSize));
    const last = Math.max(first, Math.floor(Math.max(0, lastRow) / this.blockSize));
    for (let b = first; b <= last; b++) {
      if (!this.blocks.has(b)) void this.fetchBlock(b);
      else this.touch(b);
    }
  }

  /** Purge everything and reload the visible range (sort/filter changed). */
  refetch(): void {
    this.seq++;
    this.blocks.clear();
    this.lru = [];
    this.grid.__serverPurge();
    // __serverPurge re-renders, which emits viewportChanged → ensureRange.
    // Guarantee at least the first block even with a zero-height viewport:
    void this.fetchBlock(0);
  }

  private touch(b: number): void {
    const i = this.lru.indexOf(b);
    if (i >= 0) this.lru.splice(i, 1);
    this.lru.push(b);
  }

  private async fetchBlock(b: number): Promise<void> {
    this.blocks.set(b, 'loading');
    this.touch(b);
    const mySeq = this.seq;
    this.pending++;
    try {
      const res = await this.opts.getRows({
        startRow: b * this.blockSize,
        endRow: (b + 1) * this.blockSize,
        sortModel: this.grid.sortModel,
        filterModel: this.grid.filterModel,
      });
      if (mySeq !== this.seq) return; // stale: a refetch happened meanwhile
      this.blocks.set(b, 'loaded');
      this.grid.__serverStore(b * this.blockSize, res.rows, res.totalCount);
      this.evict();
    } catch {
      if (mySeq === this.seq) this.blocks.delete(b); // retry on next scroll
    } finally {
      this.pending--;
    }
  }

  private evict(): void {
    while (this.lru.length > this.maxBlocks) {
      const b = this.lru.shift()!;
      if (this.blocks.get(b) === 'loaded') {
        this.blocks.delete(b);
        this.grid.__serverEvict(b * this.blockSize, this.blockSize);
      }
    }
  }
}
