/**
 * Audit trail — registro delle modifiche per riga (chi, quando, cosa).
 *
 * Costruito interamente sul flusso eventi pubblico: ogni `cellValueChanged`
 * diventa una voce {ts, user, colId, oldValue, newValue, source}. La
 * tracciabilità è un requisito, non un extra: il modulo è pensato per
 * essere agganciato sempre, con costo trascurabile (append su Map).
 *
 *   const audit = attachAuditTrail(grid, { user: () => session.userName });
 *   const storia = audit.forRow('INT-00042');
 */
import type { Grid } from './grid';
import type { ChangeSource, RowData, RowId } from './types';

export interface AuditEntry {
  ts: number;
  user: string;
  rowId: RowId;
  colId: string;
  oldValue: unknown;
  newValue: unknown;
  source: ChangeSource;
}

export interface AuditOptions {
  /** Chi sta modificando (es. dal contesto di sessione). @default 'utente' */
  user?: () => string;
  /** Voci massime per riga (le più vecchie vengono scartate). @default 200 */
  limitPerRow?: number;
  /** Sorgenti registrate. @default tutte tranne 'import' */
  sources?: ChangeSource[];
}

export interface AuditHandle {
  forRow(rowId: RowId): AuditEntry[];
  all(): AuditEntry[];
  clear(rowId?: RowId): void;
  detach(): void;
}

export function attachAuditTrail<T extends RowData>(
  grid: Grid<T>,
  options: AuditOptions = {},
): AuditHandle {
  const limit = options.limitPerRow ?? 200;
  const sources = new Set<ChangeSource>(
    options.sources ?? ['user', 'api', 'clipboard', 'fill', 'undo', 'redo'],
  );
  const byRow = new Map<RowId, AuditEntry[]>();

  const off = grid.events.on('cellValueChanged', (e) => {
    if (!sources.has(e.source)) return;
    let list = byRow.get(e.rowId);
    if (!list) byRow.set(e.rowId, (list = []));
    list.push({
      ts: Date.now(),
      user: options.user?.() ?? 'utente',
      rowId: e.rowId,
      colId: e.colId,
      oldValue: e.oldValue,
      newValue: e.newValue,
      source: e.source,
    });
    if (list.length > limit) list.splice(0, list.length - limit);
  });

  return {
    forRow: (rowId) => (byRow.get(rowId) ?? []).slice(),
    all: () => [...byRow.values()].flat().sort((a, b) => a.ts - b.ts),
    clear: (rowId) => (rowId ? byRow.delete(rowId) : byRow.clear()),
    detach: off,
  };
}
