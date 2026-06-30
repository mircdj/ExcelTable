# ExcelTable — Riferimento API

Tutto ciò che è esportato da `@exceltable/core` è tipizzato in TypeScript strict; i wrapper riesportano gli stessi tipi.

---

## 1. Costruzione

```ts
import { Grid } from '@exceltable/core';
const grid = new Grid(container: HTMLElement, options: GridOptions);
```

### `GridOptions`

| Proprietà | Tipo | Default | Descrizione |
|---|---|---|---|
| `columns` | `ColumnDef[]` | — | Definizione colonne (obbligatoria) |
| `data` | `T[]` | `[]` | Righe iniziali |
| `getRowId` | `(row) => string` | identità interna | **Consigliato**: identità stabile per dirty tracking e undo |
| `rowHeight` | `number` | `32` | Altezza riga in px |
| `headerHeight` | `number` | `36` | Altezza header |
| `showRowNumbers` | `boolean` | `true` | Gutter con numeri riga stile Excel |
| `editable` | `boolean` | `true` | Abilita l'editing globale |
| `locale` | `string` | `'it-IT'` | Locale per parsing/formattazione numeri e date |
| `undoLimit` | `number` | `200` | Profondità massima dello stack undo |
| `theme` | `string` | `'excel'` | Suffisso classe tema: `excel` \| `dark` \| `blue` |
| `masterDetail` | `{ field?, getDetail?, height?, columns? }` | — | Sottotabelle annidate — vedi §7 |
| `groupBy` | `string[]` | `[]` | Raggruppa per queste colonne — vedi §10 |
| `totalsRow` | `boolean` | auto | Riga totali Σ in basso (auto se c'è un `aggFunc`) |
| `formulas` | `boolean` | `false` | Motore formule — vedi §14 |
| `serverSide` | `{ getRows, blockSize?, maxBlocks? }` | — | Row model lazy dal backend — vedi §16 |
| `ariaLabel` | `string` | `'Tabella dati'` | Etichetta annunciata dagli screen reader |

### `ColumnDef`

| Proprietà | Tipo | Descrizione |
|---|---|---|
| `field` | `string` | Proprietà della riga letta/scritta (obbligatoria) |
| `id` | `string` | Id colonna; default = `field` |
| `header` | `string` | Etichetta header |
| `type` | `'text' \| 'number' \| 'date' \| 'boolean' \| 'select'` | Determina editor, parser, comparatore e allineamento |
| `width` / `minWidth` / `maxWidth` | `number` | Larghezze (default 140 / 48 / 2000) |
| `pinned` | `'left' \| null` | Blocca la colonna a sinistra (freeze) |
| `editable` | `boolean \| (row) => boolean` | Editabilità per cella |
| `sortable` / `resizable` | `boolean` | Default `true` |
| `options` | `string[]` | Valori ammessi per `type: 'select'` |
| `valueFormatter` | `(value, row) => string` | Valore → testo mostrato (es. formato valuta) |
| `valueParser` | `(input, row) => unknown` | Testo digitato → valore |
| `cellClass` | `(value, row) => string \| null` | Classe CSS condizionale per cella |
| `validator` | `(value, row) => string \| null` | Messaggio d'errore, o `null` se valido |
| `comparator` | `(a, b) => number` | Comparatore custom per l'ordinamento |
| `cellRenderer` | `CellRenderer` | Contenuto custom della cella — vedi §6 |
| `wrapText` | `boolean` | Testo su più righe nella cella |
| `aggFunc` | `'sum'\|'avg'\|'min'\|'max'\|'count'` | Aggregato su righe gruppo e riga totali |

---

## 2. API dell'istanza

### Dati

```ts
grid.setRowData(rows)                 // sostituisce i dati (azzera undo e dirty)
grid.rowCount                         // righe nella vista corrente (post filtro)
grid.totalRowCount                    // righe totali
grid.getRowByViewIndex(i)             // riga alla posizione visuale i
grid.getValue(viewIndex, colId)
grid.setCellValue(viewIndex, colId, value, source?)
grid.applyChanges(changes, source, recordUndo?)  // batch transazionale
grid.refresh()                        // ricostruisce vista e re-renderizza
```

### Selezione e navigazione

```ts
grid.ranges                           // CellRange[] correnti (coordinate vista)
grid.activeCell                       // { rowIndex, colId } | null
grid.setSelection(ranges, activeCell)
grid.selectCell(rowIndex, colVisibleIndex)
grid.navigateTo(rowIndex, colVisibleIndex)   // seleziona e scrolla
```

### Editing

```ts
grid.startEditing(cell?)              // F2 programmatica
grid.stopEditing(cancel?)
grid.isCellEditable(viewIndex, colId)
```

### Colonne

```ts
grid.setColumns(defs)
grid.setColumnWidth(colId, px)
grid.setColumnVisible(colId, visible)
grid.autoSizeColumn(colId)            // = doppio click sul bordo
grid.visibleColumns() / grid.pinnedColumns() / grid.centerColumns()
```

### Ordinamento e filtri

```ts
grid.setSortModel([{ colId: 'zona', dir: 'asc' }, { colId: 'ore', dir: 'desc' }])
grid.toggleSort(colId, multi)         // = click header (multi = Shift)
grid.setFilterModel([{ colId: 'stato', values: ['In corso'] },
                     { colId: 'note', text: 'permesso' }])
grid.getDistinctValues(colId)         // per costruire UI filtro custom
```

### Righe (inserimento/eliminazione undoabili)

```ts
grid.insertRows([{...}, {...}], atViewIndex?)  // → RowId[]; Ctrl+Z le rimuove
grid.removeRows([0, 4, 7])                      // indici vista; Ctrl+Z le ripristina
```

### Undo / dirty / export

```ts
grid.undo() / grid.redo()             // celle E operazioni di riga
grid.getDirtyRows()                   // [{ id, row }] modificate dall'ultimo markClean
grid.markClean()
grid.exportTsv(range?) / grid.exportCsv()
grid.getState() / grid.setState(s)    // viste salvabili (sort, filtri, larghezze, gruppi)
grid.moveColumn(colId, toIndex)
grid.destroy()
```

---

## 3. Eventi

Sottoscrizione tipizzata; `on` restituisce la funzione di unsubscribe.

```ts
const off = grid.events.on('cellValueChanged', (e) => { ... });
grid.events.once('gridReady', ...);
off();
```

Gli eventi **cancellabili** (⛔) supportano `e.preventDefault()` per bloccare l'operazione prima che venga applicata.

| Evento | Payload principale | Quando |
|---|---|---|
| `gridReady` | — | Istanza montata |
| `firstDataRendered` | — | Primo frame renderizzato |
| `gridDestroyed` | — | `destroy()` |
| `rowDataChanged` | `rowCount` | `setRowData` |
| ⛔ `cellValueChanging` | `CellChange`, `source` | Prima di applicare ogni modifica (veto) |
| `cellsChanged` | `changes: CellChange[]`, `source` | **Batch unico** per edit/paste/fill/undo |
| `cellValueChanged` | `CellChange`, `source` | Per singola cella, dopo `cellsChanged` |
| `cellFocused` | `cell` | Cambio cella attiva |
| `cellClicked` / `cellDoubleClicked` | `cell`, `originalEvent` | Mouse |
| ⛔ `cellEditingStarted` | `cell` | Apertura editor |
| `cellEditingStopped` | `cell`, `cancelled` | Chiusura editor |
| `selectionChanged` | `ranges` | Ogni variazione di selezione |
| ⛔ `copyStart` / `copyEnd` | `range`, `tsv` | Ctrl+C |
| ⛔ `pasteStart` | `rows: string[][]`, `target` | Prima di applicare l'incolla |
| `pasteEnd` | `changes`, `rejected` | Dopo l'incolla (rejected = celle readonly) |
| ⛔ `fillStart` / `fillEnd` | `source`, `target` / `changes` | Quadratino di riempimento |
| `sortChanged` | `sortModel` | Ordinamento |
| `filterChanged` | `filterModel` | Filtri |
| `columnResized` | `colId`, `width` | Resize/autofit |
| `columnVisibilityChanged` | `colId`, `visible` | Show/hide |
| `undoApplied` / `redoApplied` | `changes` | Ctrl+Z / Ctrl+Y |
| `undoStackChanged` | `canUndo`, `canRedo` | Per abilitare i pulsanti |
| `validationFailed` | `errors: { cell, message }[]` | Valore rifiutato dal validator |
| `viewportChanged` | `firstRow`, `lastRow` | Righe visibili (lazy loading) |
| `dirtyStateChanged` | `dirtyRowIds` | Righe modificate non salvate |

`CellChange = { rowId, rowIndex, colId, oldValue, newValue }`
`source = 'user' | 'api' | 'clipboard' | 'fill' | 'undo' | 'redo' | 'import'`

### Pattern tipici

**Salvataggio incrementale verso il backend**
```ts
grid.events.on('cellsChanged', async (e) => {
  if (e.source === 'undo' || e.source === 'redo') return;
  await api.patch('/interventi', e.changes);
});
```

**Veto su una regola di business**
```ts
grid.events.on('cellValueChanging', (e) => {
  if (e.colId === 'stato' && e.newValue === 'Completato' && !isCollaudato(e.rowId))
    e.preventDefault();
});
```

**Conferma prima di un incolla massivo**
```ts
grid.events.on('pasteStart', (e) => {
  const cells = e.rows.length * e.rows[0].length;
  if (cells > 1000 && !confirm(`Incollare ${cells} celle?`)) e.preventDefault();
});
```

---

## 4. Theming

Ogni token è una CSS custom property su `.eg-root` — basta sovrascriverla:

```css
.mio-grid .eg-root {
  --eg-accent: #0047bb;
  --eg-font: 14px/1.4 'Roboto', sans-serif;
  --eg-row-hover: #f0f6ff;
  --eg-cell-padding: 0 12px;
}
```

| Token | Default (tema excel) | Ruolo |
|---|---|---|
| `--eg-font` | `13px/1.35 'Segoe UI'…` | Font globale |
| `--eg-bg` / `--eg-fg` | `#fff` / `#1f2328` | Sfondo / testo celle |
| `--eg-border` | `#d4d4d8` | Bordo esterno |
| `--eg-grid-line` | `#e4e4e7` | Linee della griglia |
| `--eg-header-bg` / `--eg-header-fg` | `#f4f5f7` / `#3f3f46` | Header e gutter |
| `--eg-row-hover` | `#f8fafc` | Hover riga |
| `--eg-row-dirty` | `#fffbe6` | Riga modificata |
| `--eg-readonly-fg` | `#8a8f98` | Testo celle readonly |
| `--eg-accent` | `#107c41` | Colore primario (selezione, sort, fill handle) |
| `--eg-selection-bg` | `rgba(16,124,65,.10)` | Riempimento selezione |
| `--eg-active-border` | `var(--eg-accent)` | Bordo cella attiva |
| `--eg-cell-padding` | `0 8px` | Padding celle |
| `--eg-radius` | `6px` | Raggio bordi del contenitore |

Classi di stato utili per CSS custom: `.eg-cell.eg-sel`, `.eg-active`, `.eg-readonly`, `.eg-row--dirty`, `.eg-pinned`, più qualsiasi classe restituita da `cellClass`.

Temi inclusi: `eg-theme-excel`, `eg-theme-dark`, `eg-theme-blue` (si cambiano sostituendo la classe su `.eg-root`).

---

## 5. Scorciatoie implementate (matrice Excel)

| Gesto | Comportamento |
|---|---|
| Frecce | Sposta cella attiva |
| `Ctrl+Frecce` | Salto al bordo del blocco dati |
| `Shift+Frecce` / `Shift+Click` | Estende la selezione |
| `Ctrl+Click/Drag` | Range multipli |
| `Tab` / `Shift+Tab`, `Invio` / `Shift+Invio` | Avanza; in un range multi-cella **cicla nel range** |
| `Home`/`End`, `Ctrl+Home`/`Ctrl+End`, `PagSu`/`PagGiù` | Navigazione estesa |
| `Ctrl+A` | Seleziona tutto |
| `F2` / doppio click / digitazione | Editing (digitare sovrascrive, F2 mette il cursore a fine) |
| `Esc` | Annulla editing / annulla copia |
| `Alt+Invio` | A capo nella cella |
| `Ctrl+Invio` | Scrive il valore su tutta la selezione |
| `Canc`/`Backspace` | Svuota le celle selezionate |
| `Ctrl+C/X/V` | Clipboard TSV interoperabile con Excel |
| `Ctrl+Z` / `Ctrl+Y` | Undo / redo |
| Click header (+`Shift`) | Ordina (multi-colonna) |
| Doppio click sul bordo header | Autofit |
| Drag del quadratino | Riempimento con serie |
| Click sul numero di riga | Seleziona la riga |


---

## 6. Celle custom e multi-contenuto

`cellRenderer` riceve `{ value, displayValue, row, rowIndex, colId }` e può restituire:

| Ritorno | Risultato |
|---|---|
| `string` | Testo semplice |
| `Node` | Elemento DOM / componente montato |
| `{ el, destroy }` | Elemento con hook di distruzione (chiamato quando il virtualizzatore ricicla la riga) |
| **array** dei precedenti | **Righe impilate nella cella**: un valore e, a capo, un elemento custom o altro testo |

```ts
// JS puro: valore + barra di avanzamento a capo
{
  field: 'avanzamento', type: 'number', width: 220,
  cellRenderer: (p) => {
    const bar = document.createElement('div');
    bar.className = 'bar';
    bar.innerHTML = `<i style="width:${p.value}%"></i>`;
    return [p.value + '% completato', bar];   // ← multi-contenuto
  },
}
```

```tsx
// React (@exceltable/react): componenti veri, unmount automatico
import { reactCell, reactNode } from '@exceltable/react';
{ field: 'stato', cellRenderer: reactCell((p) => <StatoBadge value={p.value} />) }
{ field: 'avanzamento',
  cellRenderer: (p) => [p.displayValue, reactNode(<ProgressBar value={p.row.avanzamento} />)] }
```

```ts
// Angular (@exceltable/angular): ng-template con contesto $implicit
// <ng-template #statoTpl let-p><app-badge [value]="p.value" /></ng-template>
import { templateCell } from '@exceltable/angular';
cols = [{ field: 'stato', cellRenderer: templateCell(this.statoTpl, this.vcr) }];
```

Suggerimento: con i renderer custom alza `rowHeight` (es. 52) per ospitare più righe. Le colonne senza renderer mantengono il percorso di rendering veloce.

---

## 7. Qualsiasi JSON + sottotabelle annidate

### `Grid.fromJson(container, json, options)`

Costruisce la griglia da un payload JSON arbitrario:

```ts
const grid = Grid.fromJson(el, await res.json(), {
  exclude: ['meta', 'cliente.cap'],     // campi/sottoalberi da ignorare
  // include: ['contratto', 'cliente.nome'],  // oppure: solo questi, in quest'ordine
  overrides: { valore: { header: 'Valore €', width: 140 } },
});
```

- I tipi (`number`/`boolean`/`date`/`text`) sono inferiti dai dati
- Gli oggetti annidati diventano colonne dot-path (`cliente.citta`) leggibili **e scrivibili**
- Gli array di oggetti diventano **sottotabelle espandibili** automaticamente

`inferColumns(data, opts)`, `getPath(obj, 'a.b.c')` e `setPath(obj, 'a.b.c', v)` sono esportati per usi custom.

### Master/detail esplicito

```ts
new Grid(el, {
  columns, data,
  masterDetail: {
    field: 'materiali',          // o getDetail: (row) => row.materiali
    height: 240,                  // altezza del pannello (default 260)
    columns: [...],               // opzionale: inferite se omesse
  },
});
grid.toggleDetail(viewIndex);     // anche dal chevron ▸ nel numero di riga
grid.events.on('rowExpanded', (e) => prefetch(e.rowId));
```

Il pannello ospita una griglia ExcelTable completa (editing, selezione, clipboard). Le modifiche nella sottotabella marcano dirty la riga padre. La virtualizzazione resta esatta con qualsiasi numero di pannelli aperti, anche di altezze diverse (offset a somme prefisse).

**Dimensionamento deterministico**: il pannello si adatta al contenuto (niente vuoto sotto poche righe) fino al massimo `height`; il budget include lo spazio delle scrollbar quando servono e, quando tutto entra, l'overflow è disattivato — nessuna scrollbar può comparire per un pixel di arrotondamento.

**Ricorsione**: se le righe del detail contengono a loro volta array di oggetti, la sottotabella ha il proprio chevron (sottotabelle dentro sottotabelle), automaticamente quando le colonne sono inferite.

**Pannello custom — qualsiasi contenuto**: con `detailRenderer` il pannello mostra ciò che vuoi al posto della griglia annidata — testo, HTML/CSS tuo, un componente con lifecycle:

```ts
masterDetail: {
  getDetail: (r) => r.ordini,
  detailRenderer: ({ row, rowIndex, rowId }) => {
    const card = document.createElement('div');
    card.className = 'mia-card';            // stila come vuoi nel TUO css
    card.innerHTML = `<h4>${row.cliente.nome}</h4>…`;
    return { el: card, destroy: () => pulizia() };  // o solo un Node / stringa
  },
  getHeight: (row) => 86 + row.ordini.length * 46,  // altezza per riga (opz.)
}
```

Con `detailRenderer` ogni riga è espandibile; `destroy` viene chiamato al collapse o al riciclo del virtualizzatore. Funziona anche con i wrapper: dentro `detailRenderer` puoi montare componenti React (`reactNode(<Card/>)`) o template Angular con gli stessi helper delle celle.

---

## 8. Accessibilità (WCAG 2.1 AA)

- Pattern **ARIA grid** completo: `role="grid|row|columnheader|gridcell|rowheader"`, `aria-rowcount/colcount/rowindex/colindex`
- `aria-activedescendant` sulla cella attiva, `aria-selected` sulle selezionate, `aria-sort` sugli header
- Pulsanti filtro/espandi con `aria-label`; il popup filtro è un `role="dialog"` con focus iniziale
- Tutta l'interazione è possibile da tastiera (matrice Excel, §5)
- Contrasti testo ≥ 4.5:1 in tutti i temi inclusi; focus visibile su ogni controllo
- `prefers-reduced-motion` disattiva le animazioni

---

## 9. Test

`node tests/run.mjs` esegue 33 test (jsdom): serie del fill handle, parsing it-IT, inferenza JSON (tipi, dot-path, include/exclude/overrides, detail), pipeline dati (eventi cancellabili, validazione, undo/redo, sort+filtro, CSV), master/detail (espansione, offset, griglia annidata, `fromJson`), renderer custom (lifecycle `destroy`, multi-contenuto) e contratto ARIA.


---

## 10. Raggruppamento, aggregazioni e riga totali

```ts
new Grid(el, {
  columns: [
    { field: 'zona' },
    { field: 'importo', type: 'number', aggFunc: 'sum' },
    { field: 'ore', type: 'number', aggFunc: 'avg' },
  ],
  data,
  groupBy: ['zona'],          // anche multi-livello: ['zona', 'stato']
});

grid.setGroupBy(['zona', 'stato']);  // runtime (anche dal context menu sull'header)
grid.toggleGroup(viewIndex);          // o click sulla riga gruppo
grid.grandTotals();                   // { count, aggs } sul filtrato
```

Le righe gruppo mostrano chiave, conteggio e gli aggregati delle colonne con `aggFunc`; sono comprimibili (▸) e non editabili. La **riga totali Σ** resta fissata in basso e si aggiorna con filtri e modifiche. `isGroupRow(i)` / `getGroupRow(i)` distinguono le righe gruppo nelle integrazioni.

---

## 11. Trova & Sostituisci

`Ctrl+F` apre il pannello (conteggio "n / totale", ↑/↓, Aa per maiuscole), `Ctrl+H` aggiunge la riga di sostituzione ("Sostituisci" / "Tutto"). Tutte le sostituzioni passano da `applyChanges`: validate, vetabili e **undoabili**. Via API:

```ts
const hits = grid.findAll('permesso', { matchCase: false });  // CellRef[]
grid.replaceIn(hits, 'permesso', 'autorizzazione');           // → n sostituzioni
```

---

## 12. Context menu

Tasto destro su una cella: Copia/Taglia/Incolla (clipboard di sistema), Svuota, **Inserisci riga sopra/sotto**, **Elimina righe selezionate**, Autofit, Nascondi colonna. Tasto destro sull'header: Ordina A→Z / Z→A, **Raggruppa per la colonna** (o rimuovi), Autofit, Nascondi, Mostra colonne nascoste. Ogni voce invoca l'API pubblica: tutto resta osservabile dagli eventi e undoabile.

---

## 13. Excel nativo (.xlsx)

```ts
import { exportXlsx, importXlsx } from '@exceltable/core';

// Export: bytes pronti per il download (vista corrente, header inclusi)
const bytes = exportXlsx(grid, 'Registro');
download(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));

// Import: prima sheet → righe keyed per header
const { headers, rows } = await importXlsx(file);   // File | Blob | ArrayBuffer
grid.setRowData(mappaSuiCampi(rows));
```

Zero dipendenze: l'export scrive un OOXML store-only (apribile da Excel/LibreOffice/Google Sheets); l'import legge shared strings e inline strings, celle store e deflate (via `DecompressionStream`, disponibile in tutti i browser moderni e Node 18+). Numeri e booleani mantengono il tipo. Fuori scope di questo modulo: stili, formule, fogli multipli.


---

## 14. Formule

Attiva con `formulas: true`. Una cella il cui valore inizia con `=` è una formula: la griglia mostra il risultato, **l'editor (F2) mostra la sorgente**, undo/redo e dirty tracking funzionano invariati (il valore salvato È la formula). Ordinamento e aggregazioni usano i valori calcolati; `valueFormatter` si applica al risultato.

**Sintassi** (convenzioni Excel italiano): separatore argomenti `;`, decimali `,` o `.`, operatori `+ - * / ^ &` e confronti `= <> < > <= >=`, riferimenti `A1` (lettera = posizione colonna, numero = riga dei dati 1-based), range `A1:B10`, `$` assoluti accettati.

**Funzioni** (nomi IT/EN): `SOMMA/SUM`, `MEDIA/AVERAGE`, `MIN`, `MAX`, `CONTA/COUNT`, `CONTA.VALORI/COUNTA`, `SE/IF`, `E/AND`, `O/OR`, `NON/NOT`, `ARROTONDA/ROUND`, `ASS/ABS`, `CONCATENA/CONCAT`, `LUNGHEZZA/LEN`, `MAIUSC/UPPER`, `MINUSC/LOWER`, `OGGI/TODAY`.

**Errori**: `#NOME?` funzione sconosciuta, `#DIV/0!`, `#CIRC!` riferimenti circolari, `#ERRORE!` sintassi/valore.

**Semantica dei riferimenti — deliberata e documentata**: i riferimenti puntano alle coordinate dei *dati* (riga 1 = prima riga dell'ordine originale), quindi ordinare o filtrare non cambia mai il significato di una formula. Inserire/eliminare righe ricalcola tutto ma non riscrive i riferimenti.

```ts
new Grid(el, { formulas: true, columns, data: [
  { voce: 'Cavo', q: 2, prezzo: 10, tot: '=B1*C1' },
  { voce: 'TOT',  q: null, prezzo: null, tot: '=SOMMA(D1:D1)' },
]});
```

Il motore (`parseFormula`, `evaluate`, `FormulaEngine`) è esportato anche per usi standalone.

---

## 15. Pivot

`pivot(data, columns, config)` è una **funzione pura** — nessun accoppiamento con la griglia, banale da testare:

```ts
import { pivot } from '@exceltable/core';

const p = pivot(rows, columnDefs, {
  rows: ['zona', 'ditta'],          // dimensioni riga (multi-livello)
  cols: 'stato',                     // dimensione colonna (opzionale)
  values: [
    { colId: 'importo', aggFunc: 'sum', header: 'Importo', valueFormatter: euro },
    { colId: 'id', aggFunc: 'count', header: 'N' },
  ],
  totals: true,                      // colonna "Totale" + riga "Totale" (default)
});
new Grid(el, { columns: p.columns, data: p.rows, editable: false });
```

Il risultato è `{ columns, rows }`: la prima dimensione riga è bloccata a sinistra, le colonne valore hanno `aggFunc` (quindi la riga Σ compare gratis), le celle senza dati sono `null`.

---

## 16. Server-side row model

Per dataset troppo grandi per il client:

```ts
new Grid(el, {
  columns,
  getRowId: (r) => r.id,
  serverSide: {
    blockSize: 100,         // righe per richiesta (default 100)
    maxBlocks: 30,          // blocchi in cache LRU (default 30)
    async getRows({ startRow, endRow, sortModel, filterModel }) {
      const res = await fetch(`/api/pod?from=${startRow}&to=${endRow}&sort=...`);
      const { rows, totalCount } = await res.json();
      return { rows, totalCount };
    },
  },
});
```

Comportamento garantito (e testato): la griglia mostra subito `totalCount` righe virtuali; i blocchi vengono richiesti quando entrano nel viewport; le righe non arrivate mostrano «…» (`isRowLoading`); **ordinamento e filtri sono delegati** (cache svuotata, refetch col modello corrente); un numero di sequenza **scarta le risposte stantie** così un sort rapido su rete lenta non può sovrascrivere dati freschi; cache LRU con eviction. Editing, selezione, clipboard e dirty tracking funzionano sulle righe caricate. Non disponibili in server mode (per design): raggruppamento locale, `insertRows`/`removeRows`, fill su righe non caricate; il filtro per valori elenca solo i valori caricati.


---

## 17. i18n — stringhe UI

Tutte le stringhe della UI (filtri, menu, trova/sostituisci, ARIA) vivono in una tabella tipizzata `Strings`. Italiano è il default, `EN` è esportato pronto, qualsiasi override parziale:

```ts
import { EN } from '@exceltable/core';
new Grid(el, { columns, data, strings: EN });                      // inglese
new Grid(el, { columns, data, strings: { menuCopy: 'Copia!' } }); // override puntuale
grid.t('menuDeleteRows', { n: 3 })  // → "Elimina 3 righe"
```

---

## 18. Altezza righe variabile

```ts
new Grid(el, {
  columns, data,
  getRowHeight: (row) => 32 + estimateLines(row.note) * 16,  // deterministica, decide l'app
});
```

Gli offset usano somme prefisse (`Float64Array`, O(n) una volta, O(log n) il lookup): la virtualizzazione resta esatta con 500k righe di altezze tutte diverse. Minimo 20px. Combinabile con `wrapText` sulle colonne.

---

## 19. Tree data — gerarchie WBS

```ts
new Grid(el, {
  columns,
  data: commesse,                                  // nodi con children annidati
  getRowId: (r) => r.id,
  treeData: { childrenField: 'children', indent: 22 },   // o getChildren: (r) => ...
});
grid.treeLevel(i); grid.treeHasChildren(i); grid.toggleTreeNode(i);
```

Semantica: radici espanse all'avvio; **l'ordinamento agisce tra fratelli** (la gerarchia non si rompe); **il filtro mantiene visibili gli antenati** dei match; editing su un nodo marca dirty quel nodo; `aria-level` corretto per gli screen reader. In tree mode `insertRows`/`removeRows`/`groupBy` sono disabilitati per design.

---

## 20. Touch, drag & drop, editor ricchi, copia HTML

**Touch** (pointer events, stessi handler del mouse): tap = seleziona · doppio tap = modifica · long-press fermo = menu contestuale · long-press + trascina = estende la selezione (lo scroll nativo resta sul trascinamento semplice) · fill handle trascinabile col dito (`touch-action: none` solo sul quadratino).

**Drag & drop colonne**: trascina un header oltre 5px → ghost + indicatore di inserimento → rilascia per riordinare; sotto i 5px resta il click di ordinamento. Via API resta `moveColumn(colId, index)`.

**Editor ricchi**: `type: 'date'` apre il date picker nativo (valore ISO); `type: 'select'` con più di 8 opzioni diventa input con datalist (autocomplete) e i valori fuori elenco vengono rifiutati con flash visivo; ogni `validationFailed` mostra bordo rosso + tooltip sulla cella per ~2,5s oltre a emettere l'evento.

**Copia come HTML**: Ctrl+C e la voce Copia scrivono negli appunti sia TSV sia `<table>` (numeri allineati a destra): incollato in Outlook/Word/Excel arriva come tabella formattata.

---

## 21. Offline — coda di sincronizzazione

```ts
import { attachOfflinePersistence } from '@exceltable/core';

const offline = attachOfflinePersistence(grid, { key: 'registro-cantiere' });
offline.restore();                      // al bootstrap, dopo setRowData: riapplica la coda
offline.pending();                      // [{rowId, colId, oldValue, newValue, ts, source}]
await offline.flush(async (changes) => api.patch(changes));  // ok → coda svuotata
                                                             // throw → coda intatta, si ritenta
```

Ogni modifica utente è persistita subito (localStorage di default, fallback in-memory, adapter iniettabile per test/IndexedDB). Una cella riportata al valore iniziale esce dalla coda da sola. `oldValue` è il **primo** valore visto prima delle modifiche offline — quello che il backend si aspetta per l'optimistic locking.

---

## 22. Audit trail

```ts
import { attachAuditTrail } from '@exceltable/core';
const audit = attachAuditTrail(grid, { user: () => session.user, limitPerRow: 200 });
audit.forRow('INT-00042');  // [{ts, user, colId, oldValue, newValue, source}]
```

Registra ogni `cellValueChanged` (incluse undo/redo, escluso `import` di default). Con `contextMenuItems` aggiungi la voce "Storico riga" al menu:

```ts
new Grid(el, { ..., contextMenuItems: (items, ctx) => ctx.cell
  ? [...items, { sep: true }, { label: 'Storico riga', action: () => apriPannello(ctx.cell) }]
  : items });
```

---

## 23. Suite E2E Playwright

In `tests/e2e/`: regressioni di **layout** (allineamento colonne bloccate — il test che avrebbe intercettato il bug pinned —, sticky su scroll, totali, pannelli senza scrollbar superflue), **tastiera reale** (navigazione, editing, undo, trova), **clipboard** (TSV + HTML letti dagli appunti veri), **mouse** (drag selezione, fill handle, drag&drop header, context menu), **touch** su profilo iPad (tap, doppio tap, long-press via CDP), **screenshot** (3 temi + gruppi + detail).

```bash
npx playwright install chromium   # una volta
npm run e2e                       # esegue (22 test)
npm run e2e:update                # genera/aggiorna le baseline screenshot
```
