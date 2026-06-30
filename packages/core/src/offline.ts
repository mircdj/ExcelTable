/**
 * Offline — persistenza locale delle modifiche non salvate + coda di sync.
 *
 * In cantiere la rete va e viene: questo modulo intercetta ogni modifica
 * utente, la persiste subito in uno storage locale (localStorage di
 * default, con fallback in-memory se non disponibile, o qualsiasi adapter
 * iniettato), la ripristina al riavvio e la invia al backend quando torna
 * la linea — con semantica "tutto o niente per batch": se l'invio fallisce
 * la coda resta intatta e si ritenta.
 *
 *   const offline = attachOfflinePersistence(grid, { key: 'registro' });
 *   offline.restore();                       // al bootstrap, dopo setRowData
 *   button.onclick = () => offline.flush(async (changes) => api.patch(changes));
 */
import type { Grid } from './grid';
import type { CellChange, ChangeSource, RowData, RowId } from './types';

export interface PendingChange {
  rowId: RowId;
  colId: string;
  /** Primo valore visto prima delle modifiche offline. */
  oldValue: unknown;
  newValue: unknown;
  /** Ultimo aggiornamento (epoch ms). */
  ts: number;
  source: ChangeSource;
}

export interface OfflineStorage {
  get(key: string): string | null;
  set(key: string, value: string): void;
  remove(key: string): void;
}

export interface OfflineOptions {
  /** Chiave dello storage (namespace per più griglie). */
  key: string;
  /** Adapter custom (test, IndexedDB wrapper, …). Default: localStorage con fallback in-memory. */
  storage?: OfflineStorage;
  /** Sorgenti da persistere. @default user, clipboard, fill, redo */
  sources?: ChangeSource[];
}

export interface OfflineHandle {
  /** Modifiche in coda (ultimo valore per cella). */
  pending(): PendingChange[];
  /** Riapplica la coda alla griglia (dopo un reload) e marca dirty le righe. */
  restore(): number;
  /**
   * Invia la coda: `sender` riceve i batch; se risolve, la coda è svuotata;
   * se lancia/rigetta, la coda resta intatta per il prossimo tentativo.
   */
  flush(sender: (changes: PendingChange[]) => Promise<void>): Promise<boolean>;
  /** Svuota la coda senza inviare (es. dopo un salvataggio fatto altrove). */
  clear(): void;
  /** Stacca i listener. */
  detach(): void;
}

function defaultStorage(): OfflineStorage {
  try {
    const probe = '__eg_probe__';
    localStorage.setItem(probe, '1');
    localStorage.removeItem(probe);
    return {
      get: (k) => localStorage.getItem(k),
      set: (k, v) => localStorage.setItem(k, v),
      remove: (k) => localStorage.removeItem(k),
    };
  } catch {
    const mem = new Map<string, string>();
    return {
      get: (k) => mem.get(k) ?? null,
      set: (k, v) => void mem.set(k, v),
      remove: (k) => void mem.delete(k),
    };
  }
}

export function attachOfflinePersistence<T extends RowData>(
  grid: Grid<T>,
  options: OfflineOptions,
): OfflineHandle {
  const storage = options.storage ?? defaultStorage();
  const storageKey = 'exceltable:offline:' + options.key;
  const sources = new Set<ChangeSource>(options.sources ?? ['user', 'clipboard', 'fill', 'redo']);

  const load = (): Map<string, PendingChange> => {
    try {
      const raw = storage.get(storageKey);
      if (!raw) return new Map();
      return new Map(Object.entries(JSON.parse(raw) as Record<string, PendingChange>));
    } catch {
      return new Map();
    }
  };
  const save = (queue: Map<string, PendingChange>): void => {
    try {
      if (queue.size === 0) storage.remove(storageKey);
      else storage.set(storageKey, JSON.stringify(Object.fromEntries(queue)));
    } catch {
      /* storage pieno o negato: la coda resta in memoria nel Map */
    }
  };

  let queue = load();

  const onChanges = (e: { changes: CellChange[]; source: ChangeSource }) => {
    if (!sources.has(e.source)) return;
    for (const ch of e.changes) {
      const key = ch.rowId + '\u0000' + ch.colId;
      const existing = queue.get(key);
      queue.set(key, {
        rowId: ch.rowId,
        colId: ch.colId,
        oldValue: existing ? existing.oldValue : ch.oldValue, // primo old visto
        newValue: ch.newValue,
        ts: Date.now(),
        source: e.source,
      });
      // Tornati al valore di partenza → la cella non è più pendente
      const cur = queue.get(key)!;
      if (JSON.stringify(cur.newValue) === JSON.stringify(cur.oldValue)) queue.delete(key);
    }
    save(queue);
  };
  const off = grid.events.on('cellsChanged', onChanges);

  return {
    pending: () => [...queue.values()].sort((a, b) => a.ts - b.ts),

    restore(): number {
      queue = load();
      const changes: CellChange[] = [];
      for (const p of queue.values()) {
        // ritrova la riga per id nella vista corrente
        for (let i = 0; i < grid.rowCount; i++) {
          if (grid.getRowIdByViewIndex(i) === p.rowId) {
            changes.push({
              rowId: p.rowId,
              rowIndex: i,
              colId: p.colId,
              oldValue: grid.getValue(i, p.colId),
              newValue: p.newValue,
            });
            break;
          }
        }
      }
      // 'import': non finisce di nuovo in coda, niente undo del ripristino
      const applied = grid.applyChanges(changes, 'import', false);
      return applied.length;
    },

    async flush(sender): Promise<boolean> {
      const batch = [...queue.values()].sort((a, b) => a.ts - b.ts);
      if (!batch.length) return true;
      try {
        await sender(batch);
        for (const p of batch) queue.delete(p.rowId + '\u0000' + p.colId);
        save(queue);
        return true;
      } catch {
        return false; // coda intatta, si ritenta
      }
    },

    clear(): void {
      queue.clear();
      save(queue);
    },

    detach(): void {
      off();
    },
  };
}
