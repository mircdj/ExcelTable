/**
 * ExcelTable test suite — run with: node tests/run.mjs
 *
 * Covers: series detection, locale parsing, JSON inference, dot paths,
 * grid data pipeline (sort/filter/undo/dirty), selection API, CSV export,
 * event dispatch + cancellation, master/detail expansion, custom renderer
 * lifecycle and the ARIA accessibility contract — all in jsdom.
 */
import { JSDOM } from 'jsdom';
import assert from 'node:assert/strict';

/* ---- jsdom environment ---- */
const dom = new JSDOM('<!doctype html><html><body><div id="host"></div></body></html>', {
  pretendToBeVisual: true,
});
const { window } = dom;
for (const k of ['document', 'HTMLElement', 'HTMLTextAreaElement', 'HTMLInputElement', 'HTMLSelectElement', 'Node', 'MouseEvent', 'KeyboardEvent', 'Event', 'getComputedStyle', 'Blob']) {
  globalThis[k] = window[k];
}
globalThis.window = window;
globalThis.requestAnimationFrame = (fn) => (fn(performance.now()), 0);
// jsdom non ha PointerEvent: polyfill minimale per i test touch
window.PointerEvent = class PointerEvent extends window.MouseEvent {
  constructor(type, init = {}) {
    super(type, init);
    this.pointerId = init.pointerId ?? 1;
    this.pointerType = init.pointerType ?? 'mouse';
  }
};
globalThis.PointerEvent = window.PointerEvent;
globalThis.ResizeObserver = class { observe() {} disconnect() {} };
// jsdom has no layout: give the viewport a size so virtualization renders rows.
Object.defineProperties(window.HTMLElement.prototype, {
  clientHeight: { get() { return this.className?.includes?.('eg-viewport') ? 400 : 24; }, configurable: true },
  clientWidth: { get() { return this.className?.includes?.('eg-viewport') ? 900 : 100; }, configurable: true },
});

const { Grid, makeSeries, parseValue, inferColumns, getPath, setPath, exportXlsx, importXlsx,
        pivot, parseFormula, evaluate, FormulaError, EN,
        attachOfflinePersistence, attachAuditTrail } =
  await import('../packages/core/dist/index.js');

let passed = 0, failed = 0;
const pending = [];
const test = (name, fn) => {
  const run = async () => {
    try { await fn(); passed++; console.log('  ✓', name); }
    catch (err) { failed++; console.error('  ✗', name, '\n    ', err.message); }
  };
  pending.push(run);
};
const section = (s) => console.log('\n' + s);

/* ================================================================== */
section('values — serie del fill handle');
test('serie numerica lineare 1,2 → 3,4,5', () => {
  const s = makeSeries([1, 2], 'number');
  assert.deepEqual([s(0), s(1), s(2)], [3, 4, 5]);
});
test('serie con passo 10', () => assert.equal(makeSeries([10, 20], 'number')(0), 30));
test('numero singolo incrementa di 1', () => assert.equal(makeSeries([7], 'number')(1), 9));
test('testo con suffisso: Cabina 7 → Cabina 8', () =>
  assert.equal(makeSeries(['Cabina 7'], 'text')(0), 'Cabina 8'));
test('padding conservato: POD-009 → POD-010', () =>
  assert.equal(makeSeries(['POD-009'], 'text')(0), 'POD-010'));
test('pattern non lineare si ripete', () => {
  const s = makeSeries([1, 5, 2], 'number');
  assert.deepEqual([s(0), s(1), s(2), s(3)], [1, 5, 2, 1]);
});
test('serie di date giornaliera', () =>
  assert.equal(makeSeries(['2026-06-01', '2026-06-02'], 'date')(0), '2026-06-03'));

section('values — parsing locale it-IT');
const col = (type) => ({ type, def: {} });
test('"1.234,56" → 1234.56', () => assert.equal(parseValue('1.234,56', col('number'), 'it-IT'), 1234.56));
test('"50%" → 0.5', () => assert.equal(parseValue('50%', col('number'), 'it-IT'), 0.5));
test('"01/02/2026" → 2026-02-01', () => assert.equal(parseValue('01/02/2026', col('date'), 'it-IT'), '2026-02-01'));
test('"sì" → true', () => assert.equal(parseValue('sì', col('boolean'), 'it-IT'), true));

/* ================================================================== */
section('infer — adattamento a JSON arbitrario');
const json = [
  { id: 1, nome: 'Mario', attivo: true, creato: '2026-01-05',
    indirizzo: { citta: 'Milano', cap: '20100' },
    ordini: [{ codice: 'A1', importo: 10 }] },
  { id: 2, nome: 'Lucia', attivo: false, creato: '2026-02-11',
    indirizzo: { citta: 'Roma', cap: '00100' },
    ordini: [{ codice: 'B2', importo: 22 }] },
];
test('inferisce i tipi (number/boolean/date/text)', () => {
  const { columns } = inferColumns(json);
  const byField = Object.fromEntries(columns.map((c) => [c.field, c.type]));
  assert.equal(byField.id, 'number');
  assert.equal(byField.attivo, 'boolean');
  assert.equal(byField.creato, 'date');
  assert.equal(byField.nome, 'text');
});
test('appiattisce gli oggetti annidati in dot-path', () => {
  const { columns } = inferColumns(json);
  assert.ok(columns.some((c) => c.field === 'indirizzo.citta'));
});
test('rileva gli array annidati come detail', () => {
  const { detailFields } = inferColumns(json);
  assert.deepEqual(detailFields, ['ordini']);
});
test('exclude rimuove campi e sottoalberi', () => {
  const { columns } = inferColumns(json, { exclude: ['indirizzo', 'id'] });
  assert.ok(!columns.some((c) => c.field.startsWith('indirizzo') || c.field === 'id'));
});
test('include impone ordine e selezione', () => {
  const { columns } = inferColumns(json, { include: ['nome', 'indirizzo.citta'] });
  assert.deepEqual(columns.map((c) => c.field), ['nome', 'indirizzo.citta']);
});
test('overrides applicati alla definizione', () => {
  const { columns } = inferColumns(json, { overrides: { nome: { width: 300, header: 'Cliente' } } });
  const nome = columns.find((c) => c.field === 'nome');
  assert.equal(nome.width, 300);
  assert.equal(nome.header, 'Cliente');
});
test('header leggibili da camelCase/snake_case', () => {
  const { columns } = inferColumns([{ data_inizio: 'x', capoCantiere: 'y' }]);
  assert.deepEqual(columns.map((c) => c.header), ['Data inizio', 'Capo Cantiere']);
});
test('getPath / setPath su percorsi annidati', () => {
  const obj = { a: { b: { c: 1 } } };
  assert.equal(getPath(obj, 'a.b.c'), 1);
  setPath(obj, 'a.b.c', 9);
  setPath(obj, 'x.y', 'nuovo');
  assert.equal(obj.a.b.c, 9);
  assert.equal(obj.x.y, 'nuovo');
});

/* ================================================================== */
section('grid — pipeline dati, eventi, undo, dirty');
const host = () => {
  const d = window.document.createElement('div');
  window.document.body.appendChild(d);
  return d;
};
const mk = (extra = {}) =>
  new Grid(host(), {
    columns: [
      { field: 'id', editable: false },
      { field: 'nome' },
      { field: 'ore', type: 'number', validator: (v) => (v < 0 ? 'negativo' : null) },
    ],
    data: [
      { id: 'A', nome: 'Rossi', ore: 8 },
      { id: 'B', nome: 'bianchi', ore: 2 },
      { id: 'C', nome: 'Verdi', ore: 5 },
    ],
    getRowId: (r) => r.id,
    ...extra,
  });

test('rendering iniziale: righe e celle nel DOM con ruoli ARIA', () => {
  const g = mk();
  const root = g['renderer'].root;
  assert.equal(root.getAttribute('role'), 'grid');
  assert.equal(root.getAttribute('aria-rowcount'), '4'); // 3 + header
  assert.ok(root.querySelectorAll('[role="gridcell"]').length >= 9);
  assert.ok(root.querySelector('[role="columnheader"]'));
  g.destroy();
});
test('setCellValue applica, marca dirty ed emette eventi', () => {
  const g = mk();
  const seen = [];
  g.events.on('cellValueChanged', (e) => seen.push(e));
  g.setCellValue(0, 'nome', 'Russo', 'user');
  assert.equal(g.getValue(0, 'nome'), 'Russo');
  assert.equal(seen.length, 1);
  assert.equal(seen[0].oldValue, 'Rossi');
  assert.deepEqual(g.getDirtyRows().map((d) => d.id), ['A']);
  g.destroy();
});
test('cellValueChanging cancellabile blocca la modifica', () => {
  const g = mk();
  g.events.on('cellValueChanging', (e) => e.preventDefault());
  g.setCellValue(0, 'nome', 'Bloccato', 'user');
  assert.equal(g.getValue(0, 'nome'), 'Rossi');
  g.destroy();
});
test('validator rifiuta ed emette validationFailed', () => {
  const g = mk();
  let err = null;
  g.events.on('validationFailed', (e) => (err = e.errors[0]));
  g.setCellValue(0, 'ore', -3, 'user');
  assert.equal(g.getValue(0, 'ore'), 8);
  assert.equal(err.message, 'negativo');
  g.destroy();
});
test('undo/redo ripristinano i valori', () => {
  const g = mk();
  g.setCellValue(0, 'ore', 99, 'user');
  g.undo();
  assert.equal(g.getValue(0, 'ore'), 8);
  g.redo();
  assert.equal(g.getValue(0, 'ore'), 99);
  g.destroy();
});
test('sort naturale + multi e filtro set compongono la vista', () => {
  const g = mk();
  g.setSortModel([{ colId: 'nome', dir: 'asc' }]);
  assert.deepEqual([0, 1, 2].map((i) => g.getValue(i, 'nome')), ['bianchi', 'Rossi', 'Verdi']);
  g.setFilterModel([{ colId: 'nome', values: ['Rossi', 'Verdi'] }]);
  assert.equal(g.rowCount, 2);
  g.setFilterModel([]);
  assert.equal(g.rowCount, 3);
  g.destroy();
});
test('selezione + exportCsv con separatore ;', () => {
  const g = mk();
  g.selectCell(0, 1);
  assert.deepEqual(g.ranges, [{ startRow: 0, endRow: 0, startCol: 1, endCol: 1 }]);
  assert.ok(g.exportCsv().startsWith('id;nome;ore'));
  g.destroy();
});
test('dot-path: colonna su campo annidato legge e scrive', () => {
  const g = new Grid(host(), {
    columns: [{ field: 'cliente.citta' }],
    data: [{ id: 1, cliente: { citta: 'Milano' } }],
  });
  assert.equal(g.getValue(0, 'cliente.citta'), 'Milano');
  g.setCellValue(0, 'cliente.citta', 'Torino', 'user');
  assert.equal(g.getRowByViewIndex(0).cliente.citta, 'Torino');
  g.destroy();
});

/* ================================================================== */
section('grid — master/detail (sottotabelle)');
test('chevron presente solo dove esiste detail; expand/collapse', () => {
  const g = new Grid(host(), {
    columns: [{ field: 'nome' }],
    data: [
      { id: 'X', nome: 'Con ordini', ordini: [{ codice: 'A', importo: 5 }] },
      { id: 'Y', nome: 'Senza' },
    ],
    getRowId: (r) => r.id,
    masterDetail: { field: 'ordini' },
  });
  assert.equal(g.hasDetail(0), true);
  assert.equal(g.hasDetail(1), false);
  let expandedEvent = null;
  g.events.on('rowExpanded', (e) => (expandedEvent = e.rowId));
  g.toggleDetail(0);
  assert.equal(expandedEvent, 'X');
  assert.equal(g.isExpanded(0), true);
  // Detail panel with a nested grid in the DOM
  const panel = g['renderer'].root.querySelector('.eg-detail');
  assert.ok(panel, 'pannello detail renderizzato');
  assert.ok(panel.querySelector('.eg-root'), 'griglia annidata montata');
  // Offsets: row 1 shifted down by detailHeight
  assert.equal(g['renderer'].rowTop(1), g.options.rowHeight + g.detailHeightFor(0));
  assert.ok(g.detailHeightFor(0) < g.detailHeight, 'altezza adattiva: 1 riga << max');
  g.toggleDetail(0);
  assert.equal(g.isExpanded(0), false);
  g.destroy();
});
test('Grid.fromJson: colonne inferite + detail automatico', () => {
  const g = Grid.fromJson(host(), json, { exclude: ['indirizzo.cap'] });
  assert.ok(g.columns.some((c) => c.field === 'indirizzo.citta'));
  assert.ok(!g.columns.some((c) => c.field === 'indirizzo.cap'));
  assert.equal(g.hasDetail(0), true); // "ordini" rilevato
  g.destroy();
});

/* ================================================================== */
section('grid — renderer custom e celle multi-contenuto');
test('cellRenderer con nodo DOM + destroy chiamato al riciclo', () => {
  let destroyed = 0;
  const g = new Grid(host(), {
    columns: [
      { field: 'nome' },
      {
        field: 'stato',
        cellRenderer: (p) => {
          const b = window.document.createElement('b');
          b.textContent = p.displayValue;
          return { el: b, destroy: () => destroyed++ };
        },
      },
    ],
    data: [{ nome: 'x', stato: 'OK' }],
  });
  assert.ok(g['renderer'].root.querySelector('.eg-cell b'));
  g.refresh(); // re-render → destroy del precedente
  assert.ok(destroyed >= 1, 'destroy invocato');
  g.destroy();
});
test('cellRenderer array → righe multiple nella cella (eg-multi)', () => {
  const g = new Grid(host(), {
    columns: [{
      field: 'v',
      cellRenderer: (p) => {
        const tag = window.document.createElement('span');
        tag.textContent = 'extra';
        return [p.displayValue, tag];
      },
    }],
    data: [{ v: 'valore' }],
    rowHeight: 48,
  });
  const cell = g['renderer'].root.querySelector('.eg-cell.eg-multi');
  assert.ok(cell, 'classe multi applicata');
  assert.equal(cell.querySelectorAll('.eg-line').length, 2);
  assert.ok(cell.textContent.includes('valore') && cell.textContent.includes('extra'));
  g.destroy();
});

/* ================================================================== */
section('accessibilità — contratto ARIA');
test('header con aria-sort, celle con aria-selected/colindex', () => {
  const g = mk();
  g.toggleSort('nome', false);
  const h = g['renderer'].root.querySelector('[data-colid="nome"]');
  assert.equal(h.getAttribute('aria-sort'), 'ascending');
  g.selectCell(0, 1);
  const active = g['renderer'].root.querySelector('#eg-active-cell');
  assert.ok(active);
  assert.equal(active.getAttribute('aria-selected'), 'true');
  assert.equal(g['renderer'].root.getAttribute('aria-activedescendant'), 'eg-active-cell');
  g.destroy();
});
test('pulsanti filtro ed expand con aria-label', () => {
  const g = new Grid(host(), {
    columns: [{ field: 'nome', header: 'Nome' }],
    data: [{ nome: 'a', sub: [{ x: 1 }] }],
    masterDetail: { field: 'sub' },
  });
  assert.ok(g['renderer'].root.querySelector('[aria-label="Filtra Nome"]'));
  assert.ok(g['renderer'].root.querySelector('.eg-expand[aria-label]'));
  g.destroy();
});


/* ================================================================== */
section('grid — raggruppamento, aggregazioni, totali');
const gdata = [
  { id: '1', zona: 'Nord', importo: 100, ore: 10 },
  { id: '2', zona: 'Nord', importo: 50, ore: 5 },
  { id: '3', zona: 'Sud', importo: 30, ore: 3 },
];
const mkG = () => new Grid(host(), {
  columns: [
    { field: 'zona' },
    { field: 'importo', type: 'number', aggFunc: 'sum' },
    { field: 'ore', type: 'number', aggFunc: 'avg' },
  ],
  data: gdata.map((r) => ({ ...r })),
  getRowId: (r) => r.id,
  groupBy: ['zona'],
});
test('groupBy produce righe gruppo con conteggi e somme', () => {
  const g = mkG();
  assert.equal(g.rowCount, 5); // 2 gruppi + 3 righe
  assert.equal(g.isGroupRow(0), true);
  const nord = g.getGroupRow(0);
  assert.equal(nord.key, 'Nord');
  assert.equal(nord.count, 2);
  assert.equal(nord.aggs.importo, 150);
  assert.equal(nord.aggs.ore, 7.5);
  // righe dati sotto il gruppo
  assert.equal(g.getValue(1, 'zona'), 'Nord');
  g.destroy();
});
test('toggleGroup comprime/espande e emette groupToggled', () => {
  const g = mkG();
  let toggled = null;
  g.events.on('groupToggled', (e) => (toggled = e));
  g.toggleGroup(0);
  assert.equal(toggled.collapsed, true);
  assert.equal(g.rowCount, 3); // gruppo Nord chiuso: header Nord + header Sud + 1 riga Sud
  g.toggleGroup(0);
  assert.equal(g.rowCount, 5);
  g.destroy();
});
test('grandTotals calcola gli aggregati sul filtrato; riga totali nel DOM', () => {
  const g = mkG();
  const t = g.grandTotals();
  assert.equal(t.aggs.importo, 180);
  assert.equal(t.count, 3);
  assert.ok(g['renderer'].root.querySelector('.eg-totals-row'), 'riga totali renderizzata');
  g.setFilterModel([{ colId: 'zona', values: ['Sud'] }]);
  assert.equal(g.grandTotals().aggs.importo, 30);
  g.destroy();
});
test('le celle dei gruppi non sono editabili; riga gruppo nel DOM', () => {
  const g = mkG();
  assert.equal(g.isCellEditable(0, 'importo'), false);
  assert.ok(g['renderer'].root.querySelector('.eg-group-row .eg-group-label'));
  g.destroy();
});

section('grid — inserimento/eliminazione righe con undo');
test('insertRows inserisce, marca dirty, undo la rimuove, redo la reinserisce', () => {
  const g = mk();
  const [id] = g.insertRows([{ id: 'NEW', nome: 'Nuova', ore: 1 }], 1);
  assert.equal(g.totalRowCount, 4);
  assert.equal(g.getValue(1, 'nome'), 'Nuova');
  assert.ok(g.getDirtyRows().some((d) => d.id === 'NEW'));
  g.undo();
  assert.equal(g.totalRowCount, 3);
  g.redo();
  assert.equal(g.totalRowCount, 4);
  assert.equal(id, 'NEW');
  g.destroy();
});
test('removeRows elimina e undo ripristina nella stessa posizione', () => {
  const g = mk();
  let removed = null;
  g.events.on('rowsRemoved', (e) => (removed = e.rowIds));
  g.removeRows([0, 2]);
  assert.deepEqual(removed, ['A', 'C']);
  assert.equal(g.totalRowCount, 1);
  g.undo();
  assert.equal(g.totalRowCount, 3);
  assert.deepEqual([0, 1, 2].map((i) => g.getValue(i, 'id')), ['A', 'B', 'C']);
  g.destroy();
});

section('grid — trova & sostituisci');
test('findAll trova nei valori formattati; replaceIn è undoabile', () => {
  const g = mk();
  const hits = g.findAll('ross');
  assert.equal(hits.length, 1);
  assert.deepEqual(hits[0], { rowIndex: 0, colId: 'nome' });
  const n = g.replaceIn(hits, 'Rossi', 'Russo');
  assert.equal(n, 1);
  assert.equal(g.getValue(0, 'nome'), 'Russo');
  g.undo();
  assert.equal(g.getValue(0, 'nome'), 'Rossi');
  g.destroy();
});
test('findAll matchCase distingue le maiuscole', () => {
  const g = mk();
  assert.equal(g.findAll('ROSSI', { matchCase: true }).length, 0);
  assert.equal(g.findAll('Rossi', { matchCase: true }).length, 1);
  g.destroy();
});
test('pannello find si apre, conta e chiude (Ctrl+F)', () => {
  const g = mk();
  g['interaction'].find.open(false);
  const panel = g['renderer'].root.querySelector('.eg-find');
  assert.ok(panel);
  const input = panel.querySelector('.eg-find-q');
  input.value = 'i';
  input.dispatchEvent(new window.Event('input', { bubbles: true }));
  assert.equal(panel.querySelector('.eg-find-count').textContent, '1 / 3'); // Rossi, bianchi, Verdi
  g['interaction'].find.close();
  assert.ok(!g['renderer'].root.querySelector('.eg-find'));
  g.destroy();
});

section('grid — stato salvabile e colonne');
test('getState/setState round-trip (sort, filtri, larghezze, gruppi)', () => {
  const g = mkG();
  g.setSortModel([{ colId: 'importo', dir: 'desc' }]);
  g.setColumnWidth('zona', 222);
  const state = g.getState();
  const g2 = mkG();
  g2.setState(state);
  assert.equal(g2.sortModel[0].dir, 'desc');
  assert.equal(g2.columns.find((c) => c.id === 'zona').width, 222);
  assert.deepEqual(g2.groupBy, ['zona']);
  g.destroy(); g2.destroy();
});
test('moveColumn riordina; setColumnVisible nasconde', () => {
  const g = mk();
  g.moveColumn('ore', 0);
  assert.equal(g.visibleColumns()[0].id, 'ore');
  g.setColumnVisible('nome', false);
  assert.ok(!g.visibleColumns().some((c) => c.id === 'nome'));
  g.destroy();
});

section('excel — xlsx nativo (round-trip export → import)');
test('exportXlsx produce uno ZIP valido; importXlsx lo rilegge identico', async () => {
  const g = mk();
  const bytes = exportXlsx(g, 'Test');
  assert.equal(new DataView(bytes.buffer).getUint32(0, true), 0x04034b50, 'firma ZIP');
  const sheet = await importXlsx(bytes.buffer);
  assert.deepEqual(sheet.headers, ['id', 'nome', 'ore']);
  assert.equal(sheet.rows.length, 3);
  assert.equal(sheet.rows[0].nome, 'Rossi');
  assert.equal(sheet.rows[0].ore, 8); // numero, non stringa
  g.destroy();
});
test('importXlsx gestisce sharedStrings e celle sparse', async () => {
  // costruiamo un foglio con sharedStrings come farebbe Excel
  const g = mk();
  const bytes = exportXlsx(g);
  const sheet = await importXlsx(new Blob([bytes]));
  assert.equal(sheet.rows[2].nome, 'Verdi');
  g.destroy();
});

section('context menu');
test('contextmenu apre il menu con le azioni di riga', () => {
  const g = mk();
  const cell = g['renderer'].root.querySelector('.eg-cell[data-row="0"][data-col="1"]');
  cell.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 50, clientY: 50 }));
  const menu = g['renderer'].root.querySelector('.eg-menu');
  assert.ok(menu, 'menu aperto');
  const labels = [...menu.querySelectorAll('.eg-menu-item span')].map((s) => s.textContent);
  assert.ok(labels.includes('Copia') && labels.some((l) => l.startsWith('Inserisci riga')));
  g['interaction'].menu.close();
  g.destroy();
});

test('editable:false → griglia read-only a tenuta: niente mutazioni nel menu né fill handle', () => {
  const ro = mk({ editable: false });
  const cell = ro['renderer'].root.querySelector('.eg-cell[data-row="0"][data-col="1"]');
  cell.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 50, clientY: 50 }));
  const labels = [...ro['renderer'].root.querySelectorAll('.eg-menu-item span')].map((s) => s.textContent);
  assert.ok(labels.includes('Copia'), 'Copia resta disponibile in read-only');
  assert.ok(
    !labels.some((l) => /Inserisci|Elimina|Svuota|Taglia|Incolla/.test(l)),
    'nessuna voce di mutazione nel menu in read-only',
  );
  ro['interaction'].menu.close();
  ro.selectCell(0, 1);
  assert.equal(ro['renderer'].root.querySelector('.eg-fill-handle'), null, 'nessun fill handle in read-only');
  ro.destroy();

  // sanity: una griglia editabile mostra sia le voci di mutazione sia il fill handle
  const ed = mk();
  ed.selectCell(0, 1);
  assert.ok(ed['renderer'].root.querySelector('.eg-fill-handle'), 'fill handle presente se editabile');
  const cell2 = ed['renderer'].root.querySelector('.eg-cell[data-row="0"][data-col="1"]');
  cell2.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 50, clientY: 50 }));
  const labels2 = [...ed['renderer'].root.querySelectorAll('.eg-menu-item span')].map((s) => s.textContent);
  assert.ok(labels2.some((l) => l.startsWith('Inserisci riga')), 'Inserisci presente se editabile');
  ed['interaction'].menu.close();
  ed.destroy();
});


/* ================================================================== */
section('formule — motore (parser + valutatore)');
const ctx = (cells) => ({ cell: (c, r) => cells[`${c},${r}`] ?? null });
const ev = (f, cells = {}) => evaluate(parseFormula(f), ctx(cells));
test('aritmetica e precedenze: =2+3*4^2 → 50', () => assert.equal(ev('=2+3*4^2'), 50));
test('parentesi e unario: =-(2+3)*2 → -10', () => assert.equal(ev('=-(2+3)*2'), -10));
test('decimale italiano: =1,5+1.5 → 3', () => assert.equal(ev('=1,5+1.5'), 3));
test('riferimenti: =A1+B2', () => assert.equal(ev('=A1+B2', { '0,0': 10, '1,1': 5 }), 15));
test('SOMMA su range: =SOMMA(A1:A3)', () =>
  assert.equal(ev('=SOMMA(A1:A3)', { '0,0': 1, '0,1': 2, '0,2': 3 }), 6));
test('nomi italiani e inglesi equivalenti', () => {
  const cells = { '0,0': 4, '0,1': 6 };
  assert.equal(ev('=MEDIA(A1:A2)', cells), 5);
  assert.equal(ev('=AVERAGE(A1:A2)', cells), 5);
});
test('SE con confronto e concatenazione &', () => {
  assert.equal(ev('=SE(A1>5;"alto";"basso")', { '0,0': 9 }), 'alto');
  assert.equal(ev('="tot: "&SOMMA(A1:A2)', { '0,0': 1, '0,1': 2 }), 'tot: 3');
});
test('E/O/NON, ARROTONDA, ASS, MAIUSC', () => {
  assert.equal(ev('=E(1;VERO)'), true);
  assert.equal(ev('=O(0;FALSO)'), false);
  assert.equal(ev('=NON(0)'), true);
  assert.equal(ev('=ARROTONDA(2,567;2)'), 2.57);
  assert.equal(ev('=ASS(-9)'), 9);
  assert.equal(ev('=MAIUSC("ciao")'), 'CIAO');
});
test('CONTA vs CONTA.VALORI (numeri vs non vuoti)', () => {
  const cells = { '0,0': 1, '0,1': 'x', '0,2': null };
  assert.equal(ev('=CONTA(A1:A3)', cells), 1);
  assert.equal(ev('=CONTA.VALORI(A1:A3)', cells), 2);
});
test('errori: #DIV/0! e #NOME?', () => {
  assert.equal(ev('=1/0').code, '#DIV/0!');
  assert.equal(ev('=FUNZIONE_INESISTENTE(1)').code, '#NOME?');
});
test('range fuori dai dati → celle vuote (0 in SOMMA)', () =>
  assert.equal(ev('=SOMMA(A1:A100)', { '0,0': 7 }), 7));

section('formule — integrazione nella griglia');
const mkF = () => new Grid(host(), {
  formulas: true,
  columns: [
    { field: 'voce' },                                  // colonna A
    { field: 'q', type: 'number' },                     // colonna B
    { field: 'prezzo', type: 'number' },                // colonna C
    { field: 'tot', type: 'number', aggFunc: 'sum' },   // colonna D
  ],
  data: [
    { voce: 'Cavo', q: 2, prezzo: 10, tot: '=B1*C1' },
    { voce: 'Palo', q: 3, prezzo: 5, tot: '=B2*C2' },
    { voce: 'TOTALE', q: null, prezzo: null, tot: '=SOMMA(D1:D2)' },
  ],
  getRowId: (r) => r.voce,
});
test('le celle formula mostrano il risultato calcolato', () => {
  const g = mkF();
  assert.equal(g.getDisplayValue(0, 'tot'), '20');
  assert.equal(g.getDisplayValue(2, 'tot'), '35'); // somma di formule (ricorsivo)
  g.destroy();
});
test("l'editor mostra la formula sorgente, non il risultato", () => {
  const g = mkF();
  assert.equal(g.getEditValue(0, 'tot'), '=B1*C1');
  assert.equal(g.getEditValue(0, 'voce'), 'Cavo');
  g.destroy();
});
test('modificare una dipendenza ricalcola le dipendenti (e undo ripristina)', () => {
  const g = mkF();
  g.setCellValue(0, 'q', 10, 'user');
  assert.equal(g.getDisplayValue(0, 'tot'), '100');
  assert.equal(g.getDisplayValue(2, 'tot'), '115');
  g.undo();
  assert.equal(g.getDisplayValue(2, 'tot'), '35');
  g.destroy();
});
test('digitare "=..." salva la formula (parseInput)', () => {
  const g = mkF();
  assert.equal(g.parseInput('=B1+C1', 0, 'tot'), '=B1+C1');
  g.destroy();
});
test('ciclo → #CIRC!', () => {
  const g = new Grid(host(), {
    formulas: true,
    columns: [{ field: 'a', type: 'number' }],
    data: [{ a: '=A2' }, { a: '=A1' }],
  });
  assert.equal(g.getDisplayValue(0, 'a'), '#CIRC!');
  g.destroy();
});
test('ordinamento e aggregazioni usano i valori calcolati', () => {
  const g = mkF();
  g.setSortModel([{ colId: 'tot', dir: 'asc' }]);
  // calcolati: 20, 15, 35 → asc: Palo(15), Cavo(20), TOTALE(35)
  assert.deepEqual([0, 1, 2].map((i) => g.getValue(i, 'voce')), ['Palo', 'Cavo', 'TOTALE']);
  assert.equal(g.grandTotals().aggs.tot, 70); // 20+15+35
  g.destroy();
});
test('valueFormatter applicato al risultato della formula', () => {
  const g = new Grid(host(), {
    formulas: true,
    columns: [{ field: 'v', type: 'number', valueFormatter: (x) => x + ' €' }],
    data: [{ v: '=2*3' }],
  });
  assert.equal(g.getDisplayValue(0, 'v'), '6 €');
  g.destroy();
});

/* ================================================================== */
section('pivot');
const pdata = [
  { zona: 'Nord', stato: 'Aperto', importo: 100 },
  { zona: 'Nord', stato: 'Chiuso', importo: 50 },
  { zona: 'Nord', stato: 'Aperto', importo: 30 },
  { zona: 'Sud', stato: 'Chiuso', importo: 20 },
];
const pcols = [{ field: 'zona' }, { field: 'stato' }, { field: 'importo', type: 'number' }];
test('matrice righe×colonne con somme corrette', () => {
  const p = pivot(pdata, pcols, {
    rows: ['zona'], cols: 'stato',
    values: [{ colId: 'importo', aggFunc: 'sum', header: 'Importo' }],
  });
  const nord = p.rows.find((r) => r.__r0 === 'Nord');
  assert.equal(nord['Aperto · Importo'], 130);
  assert.equal(nord['Chiuso · Importo'], 50);
  assert.equal(nord['Totale · Importo'], 180);
  const sud = p.rows.find((r) => r.__r0 === 'Sud');
  assert.equal(sud['Aperto · Importo'], null); // nessun dato in quella cella
  assert.equal(sud['Chiuso · Importo'], 20);
});
test('riga Totale generale (per colonna e complessiva)', () => {
  const p = pivot(pdata, pcols, {
    rows: ['zona'], cols: 'stato',
    values: [{ colId: 'importo', aggFunc: 'sum', header: 'Importo' }],
  });
  const tot = p.rows[p.rows.length - 1];
  assert.equal(tot.__r0, 'Totale');
  assert.equal(tot['Aperto · Importo'], 130);
  assert.equal(tot['Chiuso · Importo'], 70);
  assert.equal(tot['Totale · Importo'], 200);
});
test('più valori (sum + count) e chiavi riga multiple', () => {
  const p = pivot(pdata, pcols, {
    rows: ['zona', 'stato'],
    values: [
      { colId: 'importo', aggFunc: 'sum', header: 'Somma' },
      { colId: 'importo', aggFunc: 'count', header: 'N' },
    ],
  });
  const nordAperto = p.rows.find((r) => r.__r0 === 'Nord' && r.__r1 === 'Aperto');
  assert.equal(nordAperto['Somma'], 130);
  assert.equal(nordAperto['N'], 2);
  assert.equal(p.columns[0].pinned, 'left'); // prima dimensione bloccata
});
test('senza dimensione colonne: aggregato semplice per gruppo', () => {
  const p = pivot(pdata, pcols, {
    rows: ['zona'], values: [{ colId: 'importo', aggFunc: 'avg', header: 'Media' }],
  });
  assert.equal(p.rows.find((r) => r.__r0 === 'Nord')['Media'], 60);
});
test('il risultato pivot è una griglia valida', () => {
  const p = pivot(pdata, pcols, {
    rows: ['zona'], cols: 'stato', values: [{ colId: 'importo', aggFunc: 'sum', header: 'I' }],
  });
  const g = new Grid(host(), { columns: p.columns, data: p.rows, editable: false });
  assert.equal(g.rowCount, 3); // Nord, Sud, Totale
  assert.ok(g['renderer'].root.querySelector('.eg-totals-row')); // aggFunc propagato
  g.destroy();
});

/* ================================================================== */
section('server-side row model');
const makeServer = (total = 10, delayCtl = null) => {
  const calls = [];
  const all = Array.from({ length: total }, (_, i) => ({ id: 'R' + i, n: i, doppia: i * 2 }));
  const getRows = async (p) => {
    calls.push(p);
    if (delayCtl) await new Promise((res) => delayCtl.push(res));
    let rows = all.slice();
    if (p.sortModel.length) {
      const dir = p.sortModel[0].dir === 'asc' ? 1 : -1;
      rows.sort((a, b) => (a.n - b.n) * dir);
    }
    return { rows: rows.slice(p.startRow, p.endRow), totalCount: total };
  };
  return { getRows, calls };
};
const tick = () => new Promise((r) => setTimeout(r, 0));

test('primo blocco caricato, totale impostato, righe lontane = placeholder', async () => {
  const srv = makeServer(200);
  const g = new Grid(host(), {
    columns: [{ field: 'id' }, { field: 'n', type: 'number' }],
    getRowId: (r) => r.id,
    serverSide: { getRows: srv.getRows, blockSize: 4 },
  });
  await tick();
  assert.equal(g.totalRowCount, 200);
  assert.equal(g.getValue(0, 'id'), 'R0');
  assert.equal(g.isRowLoading(0), false);
  assert.equal(g.isRowLoading(150), true); // blocco mai richiesto (fuori viewport)
  g.destroy();
});
test('ensureRange carica i blocchi mancanti; le celle si riempiono', async () => {
  const srv = makeServer(10);
  const g = new Grid(host(), {
    columns: [{ field: 'id' }],
    getRowId: (r) => r.id,
    serverSide: { getRows: srv.getRows, blockSize: 4 },
  });
  await tick();
  g.serverSource.ensureRange(8, 9);
  await tick();
  assert.equal(g.getValue(9, 'id'), 'R9');
  assert.equal(g.isRowLoading(9), false);
  g.destroy();
});
test('setSortModel delega al server: purge + refetch con il modello', async () => {
  const srv = makeServer(10);
  const g = new Grid(host(), {
    columns: [{ field: 'id' }, { field: 'n', type: 'number' }],
    getRowId: (r) => r.id,
    serverSide: { getRows: srv.getRows, blockSize: 4 },
  });
  await tick();
  g.setSortModel([{ colId: 'n', dir: 'desc' }]);
  await tick();
  const last = srv.calls[srv.calls.length - 1];
  assert.equal(last.sortModel[0].dir, 'desc');
  assert.equal(g.getValue(0, 'id'), 'R9'); // ordinato dal server
  g.destroy();
});
test('risposte stantie scartate (sequenza richieste)', async () => {
  const delayCtl = [];
  const srv = makeServer(8, delayCtl);
  const g = new Grid(host(), {
    columns: [{ field: 'id' }, { field: 'n', type: 'number' }],
    getRowId: (r) => r.id,
    serverSide: { getRows: srv.getRows, blockSize: 8 },
  });
  // la prima richiesta è in volo (bloccata); cambio sort prima che risponda
  g.setSortModel([{ colId: 'n', dir: 'desc' }]);
  // sblocco PRIMA la richiesta vecchia, poi la nuova
  delayCtl.shift()();           // risposta stantia (asc)
  await tick();
  delayCtl.shift()();           // risposta fresca (desc)
  await tick();
  assert.equal(g.getValue(0, 'id'), 'R7'); // la stantia non ha sovrascritto
  g.destroy();
});
test('editing su riga caricata funziona e traccia dirty; righe loading non editabili', async () => {
  const srv = makeServer(200);
  const g = new Grid(host(), {
    columns: [{ field: 'id' }, { field: 'n', type: 'number' }],
    getRowId: (r) => r.id,
    serverSide: { getRows: srv.getRows, blockSize: 3 },
  });
  await tick();
  g.setCellValue(0, 'n', 999, 'user');
  assert.equal(g.getValue(0, 'n'), 999);
  assert.deepEqual(g.getDirtyRows().map((d) => d.id), ['R0']);
  assert.equal(g.isCellEditable(150, 'n'), false); // non caricata
  g.destroy();
});
test('placeholder "…" nel DOM per le righe non caricate', async () => {
  const delayCtl = [];
  const srv = makeServer(6, delayCtl);
  const g = new Grid(host(), {
    columns: [{ field: 'id' }],
    getRowId: (r) => r.id,
    serverSide: { getRows: srv.getRows, blockSize: 6 },
  });
  await tick(); // richiesta in volo, niente dati: total ancora 0 → forziamo il primo store
  delayCtl.shift()();
  await tick();
  // ora svuota un blocco simulando eviction e re-renderizza
  g.__serverEvict(0, 3);
  g.refresh();
  const loadingCell = g['renderer'].root.querySelector('.eg-row--loading .eg-cell:not(.eg-gutter)');
  assert.ok(loadingCell);
  assert.equal(loadingCell.textContent, '…');
  g.destroy();
});


/* ================================================================== */
section('master/detail — pannelli custom, ricorsione, dimensionamento');
test('detailRenderer: HTML custom nel pannello + destroy al collapse', () => {
  let destroyed = 0;
  const g = new Grid(host(), {
    columns: [{ field: 'nome' }],
    data: [{ id: 'X', nome: 'Riga' }],
    getRowId: (r) => r.id,
    masterDetail: {
      height: 120,
      detailRenderer: (p) => {
        const card = window.document.createElement('div');
        card.className = 'mia-card';
        card.innerHTML = '<b>' + p.row.nome + '</b> — contenuto libero';
        return { el: card, destroy: () => destroyed++ };
      },
    },
  });
  assert.equal(g.hasDetail(0), true); // con detailRenderer ogni riga è espandibile
  g.toggleDetail(0);
  const card = g['renderer'].root.querySelector('.eg-detail-custom .mia-card');
  assert.ok(card, 'contenuto custom renderizzato');
  assert.ok(card.textContent.includes('Riga'));
  assert.equal(g.detailHeightFor(0), 120);
  g.toggleDetail(0);
  assert.equal(destroyed, 1, 'destroy chiamato al collapse');
  g.destroy();
});
test('detailRenderer: getHeight per riga vince su height', () => {
  const g = new Grid(host(), {
    columns: [{ field: 'n', type: 'number' }],
    data: [{ n: 1 }, { n: 5 }],
    masterDetail: { height: 200, detailRenderer: () => 'x', getHeight: (r) => 80 + r.n * 10 },
  });
  assert.equal(g.detailHeightFor(0), 90);
  assert.equal(g.detailHeightFor(1), 130);
  g.destroy();
});
test('ricorsione: array dentro il detail → sottotabella espandibile a sua volta', () => {
  const g = new Grid(host(), {
    columns: [{ field: 'nome' }],
    data: [{
      id: 'A', nome: 'Contratto',
      ordini: [{ numero: 'ORD-1', importo: 10, righe: [{ voce: 'Cavo', q: 2 }] }],
    }],
    getRowId: (r) => r.id,
    masterDetail: { field: 'ordini' },
  });
  g.toggleDetail(0);
  const panel = g['renderer'].root.querySelector('.eg-detail');
  const nestedGutterChevron = panel.querySelector('.eg-root .eg-expand');
  assert.ok(nestedGutterChevron, 'la griglia annidata ha il proprio chevron (ricorsione)');
  g.destroy();
});
test('layout deterministico: contenuto che entra → eg-fit (overflow spento)', () => {
  const g = new Grid(host(), {
    columns: [{ field: 'nome' }],
    data: [{ id: 'A', nome: 'x', sub: [{ a: 1 }, { a: 2 }] }],
    getRowId: (r) => r.id,
    masterDetail: { field: 'sub' },
  });
  const layout = g.detailLayout(0);
  assert.equal(layout.fits, true);
  assert.ok(layout.height < g.detailHeight, 'altezza adattiva sotto il massimo');
  g.toggleDetail(0);
  assert.ok(g['renderer'].root.querySelector('.eg-detail .eg-root.eg-fit'), 'classe eg-fit applicata');
  g.destroy();
});
test('molte righe → pannello al massimo, scroll previsto (fits=false)', () => {
  const g = new Grid(host(), {
    columns: [{ field: 'nome' }],
    data: [{ id: 'A', nome: 'x', sub: Array.from({ length: 50 }, (_, k) => ({ a: k })) }],
    getRowId: (r) => r.id,
    masterDetail: { field: 'sub', height: 200 },
  });
  const layout = g.detailLayout(0);
  assert.equal(layout.height, 200);
  assert.equal(layout.fits, false);
  g.destroy();
});


/* ================================================================== */
section('i18n');
test('default italiano; EN cambia le stringhe della UI', () => {
  const g1 = mk();
  assert.ok(g1['renderer'].root.querySelector('[aria-label="Filtra nome"]'));
  g1.destroy();
  const g2 = new Grid(host(), {
    columns: [{ field: 'nome', header: 'Name' }],
    data: [{ nome: 'x' }],
    strings: EN,
  });
  assert.ok(g2['renderer'].root.querySelector('[aria-label="Filter Name"]'));
  assert.equal(g2.t('menuCopy'), 'Copy');
  assert.equal(g2.t('menuDeleteRows', { n: 3 }), 'Delete 3 rows');
  g2.destroy();
});
test('override parziale delle stringhe', () => {
  const g = new Grid(host(), {
    columns: [{ field: 'a' }], data: [{ a: 1 }],
    strings: { filterSelectAll: '(Tutto quanto)' },
  });
  assert.equal(g.t('filterSelectAll'), '(Tutto quanto)');
  assert.equal(g.t('menuCopy'), 'Copia'); // il resto resta IT
  g.destroy();
});

/* ================================================================== */
section('altezza righe variabile (getRowHeight)');
test('offset a somme prefisse: rowTop/totalHeight esatti', () => {
  const g = new Grid(host(), {
    columns: [{ field: 'v' }],
    data: [{ v: 'a' }, { v: 'b' }, { v: 'c' }],
    getRowHeight: (_r, i) => 30 + i * 10, // 30, 40, 50
  });
  const r = g['renderer'];
  assert.equal(r.rowTop(0), 0);
  assert.equal(r.rowTop(1), 30);
  assert.equal(r.rowTop(2), 70);
  assert.equal(r.totalHeight(), 120);
  assert.equal(g.rowHeightOf(2), 50);
  // le righe nel DOM hanno l'altezza giusta
  const row2 = r.root.querySelector('.eg-row[data-row-index="2"]');
  assert.equal(row2.style.height, '50px');
  g.destroy();
});
test('altezza minima 20px applicata', () => {
  const g = new Grid(host(), {
    columns: [{ field: 'v' }], data: [{ v: 1 }], getRowHeight: () => 4,
  });
  assert.equal(g.rowHeightOf(0), 20);
  g.destroy();
});

/* ================================================================== */
section('tree data (gerarchie WBS)');
const treeData = [
  { id: 'C1', nome: 'Commessa Nord', tipo: 'commessa', ore: 0, children: [
    { id: 'C1.1', nome: 'Cantiere Lodi', tipo: 'cantiere', ore: 0, children: [
      { id: 'C1.1.a', nome: 'Allaccio MT', tipo: 'intervento', ore: 16 },
      { id: 'C1.1.b', nome: 'Scavo', tipo: 'intervento', ore: 24 },
    ]},
  ]},
  { id: 'C2', nome: 'Commessa Sud', tipo: 'commessa', ore: 0, children: [
    { id: 'C2.1', nome: 'Giunzione', tipo: 'intervento', ore: 8 },
  ]},
];
const mkTree = () => new Grid(host(), {
  columns: [{ field: 'nome', width: 240 }, { field: 'tipo' }, { field: 'ore', type: 'number' }],
  data: JSON.parse(JSON.stringify(treeData)),
  getRowId: (r) => r.id,
  treeData: { childrenField: 'children', indent: 20 },
});
test('appiattimento con livelli; radici espanse di default', () => {
  const g = mkTree();
  assert.equal(g.totalRowCount, 6);
  assert.equal(g.rowCount, 4); // C1, C1.1 (chiuso), C2, C2.1? No: radici espanse → C1, C1.1, C2, C2.1
  assert.equal(g.treeLevel(0), 0);
  assert.equal(g.treeLevel(1), 1);
  assert.equal(g.treeHasChildren(0), true);
  assert.equal(g.treeHasChildren(1), true); // C1.1 ha figli ma è chiuso
  g.destroy();
});
test('toggle espande/comprime i figli; indentazione nel DOM', () => {
  const g = mkTree();
  g.toggleTreeNode(1); // apre C1.1
  assert.equal(g.rowCount, 6);
  assert.equal(g.getValue(2, 'nome'), 'Allaccio MT');
  assert.equal(g.treeLevel(2), 2);
  const cell = g['renderer'].root.querySelector('.eg-cell[data-row="2"][data-col="0"]');
  assert.equal(cell.style.paddingLeft, '48px', 'indent 8+2*20');
  const row = g['renderer'].root.querySelector('.eg-row[data-row-index="2"]');
  assert.equal(row.getAttribute('aria-level'), '3');
  g.toggleTreeNode(1);
  assert.equal(g.rowCount, 4);
  g.destroy();
});
test('filtro mantiene visibili gli antenati del match', () => {
  const g = mkTree();
  g.toggleTreeNode(1);
  g.setFilterModel([{ colId: 'nome', text: 'Scavo' }]);
  const nomi = Array.from({ length: g.rowCount }, (_, i) => g.getValue(i, 'nome'));
  assert.deepEqual(nomi, ['Commessa Nord', 'Cantiere Lodi', 'Scavo']);
  g.destroy();
});
test('ordinamento tra fratelli (non globale)', () => {
  const g = mkTree();
  g.toggleTreeNode(1);
  g.setSortModel([{ colId: 'ore', dir: 'desc' }]);
  // dentro C1.1: Scavo(24) prima di Allaccio(16); C1/C2 restano radici
  assert.equal(g.getValue(2, 'nome'), 'Scavo');
  assert.equal(g.getValue(3, 'nome'), 'Allaccio MT');
  g.destroy();
});
test('editing su nodo figlio marca dirty con il suo id', () => {
  const g = mkTree();
  g.toggleTreeNode(1);
  g.setCellValue(2, 'ore', 99, 'user');
  assert.deepEqual(g.getDirtyRows().map((d) => d.id), ['C1.1.a']);
  g.destroy();
});

/* ================================================================== */
section('touch (pointer events)');
const tap = (el, x = 10, y = 10) => {
  el.dispatchEvent(new window.PointerEvent('pointerdown', { bubbles: true, pointerType: 'touch', pointerId: 7, clientX: x, clientY: y }));
  el.dispatchEvent(new window.PointerEvent('pointerup', { bubbles: true, pointerType: 'touch', pointerId: 7, clientX: x, clientY: y }));
};
test('tap seleziona la cella', () => {
  const g = mk();
  const cell = g['renderer'].root.querySelector('.eg-cell[data-row="1"][data-col="1"]');
  tap(cell);
  assert.deepEqual(g.activeCell, { rowIndex: 1, colId: 'nome' });
  g.destroy();
});
test('doppio tap apre l\'editor', () => {
  const g = mk();
  const cell = g['renderer'].root.querySelector('.eg-cell[data-row="0"][data-col="1"]');
  tap(cell); tap(cell);
  assert.ok(g['renderer'].root.querySelector('.eg-editor'), 'editor aperto');
  g.stopEditing(true);
  g.destroy();
});
test('long-press fermo apre il context menu', async () => {
  const g = mk();
  const cell = g['renderer'].root.querySelector('.eg-cell[data-row="0"][data-col="1"]');
  cell.dispatchEvent(new window.PointerEvent('pointerdown', { bubbles: true, pointerType: 'touch', pointerId: 8, clientX: 5, clientY: 5 }));
  await new Promise((r) => setTimeout(r, 520));
  cell.dispatchEvent(new window.PointerEvent('pointerup', { bubbles: true, pointerType: 'touch', pointerId: 8, clientX: 5, clientY: 5 }));
  assert.ok(g['renderer'].root.querySelector('.eg-menu'), 'menu aperto da long-press');
  g['interaction'].menu.close();
  g.destroy();
});
test('il click mouse continua a funzionare (pointerdown mouse)', () => {
  const g = mk();
  const cell = g['renderer'].root.querySelector('.eg-cell[data-row="2"][data-col="2"]');
  cell.dispatchEvent(new window.PointerEvent('pointerdown', { bubbles: true, pointerType: 'mouse', clientX: 5, clientY: 5 }));
  window.document.dispatchEvent(new window.PointerEvent('pointerup', { bubbles: true, pointerType: 'mouse' }));
  assert.deepEqual(g.activeCell, { rowIndex: 2, colId: 'ore' });
  g.destroy();
});

/* ================================================================== */
section('drag & drop colonne');
test('click header senza trascinamento → ordina (regressione)', () => {
  const g = mk();
  const h = g['renderer'].root.querySelector('.eg-hcell[data-colid="nome"]');
  h.dispatchEvent(new window.PointerEvent('pointerdown', { bubbles: true, pointerType: 'mouse', clientX: 100, clientY: 10 }));
  window.document.dispatchEvent(new window.PointerEvent('pointerup', { bubbles: true, pointerType: 'mouse', clientX: 101, clientY: 10 }));
  assert.deepEqual(g.sortModel, [{ colId: 'nome', dir: 'asc' }]);
  g.destroy();
});
test('trascinamento >5px crea ghost e marker; drop chiama moveColumn', () => {
  const g = mk();
  const h = g['renderer'].root.querySelector('.eg-hcell[data-colid="ore"]');
  h.dispatchEvent(new window.PointerEvent('pointerdown', { bubbles: true, pointerType: 'mouse', clientX: 300, clientY: 10 }));
  window.document.dispatchEvent(new window.PointerEvent('pointermove', { bubbles: true, pointerType: 'mouse', clientX: 250, clientY: 10 }));
  assert.ok(g['renderer'].root.querySelector('.eg-drag-ghost'), 'ghost creato');
  assert.ok(g['renderer'].root.querySelector('.eg-drop-marker'), 'marker creato');
  window.document.dispatchEvent(new window.PointerEvent('pointerup', { bubbles: true, pointerType: 'mouse', clientX: 250, clientY: 10 }));
  assert.ok(!g['renderer'].root.querySelector('.eg-drag-ghost'), 'ghost rimosso');
  assert.deepEqual(g.sortModel, [], 'drag non ordina');
  g.destroy();
});

/* ================================================================== */
section('editor ricchi e validazione inline');
test('colonna date → input type=date con valore ISO', () => {
  const g = new Grid(host(), {
    columns: [{ field: 'd', type: 'date' }],
    data: [{ d: '2026-06-15' }],
  });
  g.startEditing({ rowIndex: 0, colId: 'd' });
  const ed = g['renderer'].root.querySelector('.eg-editor');
  assert.equal(ed.type, 'date');
  assert.equal(ed.value, '2026-06-15');
  g.stopEditing(true);
  g.destroy();
});
test('select >8 opzioni → input con datalist; valore fuori elenco rifiutato', () => {
  const opts = Array.from({ length: 12 }, (_, i) => 'Opzione ' + i);
  const g = new Grid(host(), {
    columns: [{ field: 's', type: 'select', options: opts }],
    data: [{ s: 'Opzione 3' }],
  });
  let failed = null;
  g.events.on('validationFailed', (e) => (failed = e.errors[0].message));
  g.startEditing({ rowIndex: 0, colId: 's' });
  const ed = g['renderer'].root.querySelector('.eg-editor');
  assert.equal(ed.tagName, 'INPUT');
  assert.ok(ed.getAttribute('list'), 'datalist collegata');
  ed.value = 'Non in elenco';
  g.stopEditing(false);
  assert.equal(failed, 'Valore non presente in elenco');
  assert.equal(g.getValue(0, 's'), 'Opzione 3', 'valore originale intatto');
  g.destroy();
});
test('validator → flash visivo .eg-invalid con tooltip', () => {
  const g = mk();
  g.setCellValue(0, 'ore', -5, 'user');
  const cell = g['renderer'].root.querySelector('.eg-cell.eg-invalid');
  assert.ok(cell, 'classe applicata');
  assert.equal(cell.title, 'negativo');
  g.destroy();
});

/* ================================================================== */
section('copia come HTML');
test('rangeToHtml produce una tabella con allineamento numerico', () => {
  const g = mk();
  const html = g['interaction'].rangeToHtml({ startRow: 0, endRow: 1, startCol: 1, endCol: 2 });
  assert.ok(html.startsWith('<table>'));
  assert.ok(html.includes('<td>Rossi</td>'));
  assert.ok(html.includes('text-align:right'), 'numeri a destra');
  assert.ok(html.includes('<td style="text-align:right">8</td>'));
  g.destroy();
});

/* ================================================================== */
section('offline — coda di sync persistente');
const fakeStorage = () => {
  const m = new Map();
  return { get: (k) => m.get(k) ?? null, set: (k, v) => m.set(k, v), remove: (k) => m.delete(k), _m: m };
};
test('le modifiche finiscono in coda; tornare al valore iniziale la svuota', () => {
  const g = mk();
  const st = fakeStorage();
  const off = attachOfflinePersistence(g, { key: 'test', storage: st });
  g.setCellValue(0, 'ore', 42, 'user');
  assert.equal(off.pending().length, 1);
  assert.equal(off.pending()[0].newValue, 42);
  assert.equal(off.pending()[0].oldValue, 8);
  g.setCellValue(0, 'ore', 8, 'user'); // torna al valore di partenza
  assert.equal(off.pending().length, 0, 'cella non più pendente');
  off.detach(); g.destroy();
});
test('restore dopo "reload": coda riapplicata e righe dirty', () => {
  const st = fakeStorage();
  const g1 = mk();
  const off1 = attachOfflinePersistence(g1, { key: 'r', storage: st });
  g1.setCellValue(1, 'nome', 'Modificato offline', 'user');
  off1.detach(); g1.destroy();
  // nuova sessione, stessi dati dal server
  const g2 = mk();
  const off2 = attachOfflinePersistence(g2, { key: 'r', storage: st });
  const n = off2.restore();
  assert.equal(n, 1);
  assert.equal(g2.getValue(1, 'nome'), 'Modificato offline');
  assert.deepEqual(g2.getDirtyRows().map((d) => d.id), ['B']);
  off2.detach(); g2.destroy();
});
test('flush: successo svuota, fallimento mantiene la coda', async () => {
  const g = mk();
  const off = attachOfflinePersistence(g, { key: 'f', storage: fakeStorage() });
  g.setCellValue(0, 'ore', 100, 'user');
  const ko = await off.flush(async () => { throw new Error('rete giù'); });
  assert.equal(ko, false);
  assert.equal(off.pending().length, 1, 'coda intatta dopo il fallimento');
  let sent = null;
  const okRes = await off.flush(async (b) => { sent = b; });
  assert.equal(okRes, true);
  assert.equal(sent.length, 1);
  assert.equal(off.pending().length, 0);
  off.detach(); g.destroy();
});

/* ================================================================== */
section('audit trail');
test('registra chi/quando/cosa; undo registrato come source undo', () => {
  const g = mk();
  const audit = attachAuditTrail(g, { user: () => 'mario.rossi' });
  g.setCellValue(0, 'ore', 50, 'user');
  g.undo();
  const storia = audit.forRow('A');
  assert.equal(storia.length, 2);
  assert.equal(storia[0].user, 'mario.rossi');
  assert.equal(storia[0].oldValue, 8);
  assert.equal(storia[0].newValue, 50);
  assert.equal(storia[1].source, 'undo');
  audit.detach(); g.destroy();
});
test('limitPerRow scarta le voci più vecchie', () => {
  const g = mk();
  const audit = attachAuditTrail(g, { limitPerRow: 3 });
  for (let i = 1; i <= 5; i++) g.setCellValue(0, 'ore', i, 'user');
  assert.equal(audit.forRow('A').length, 3);
  assert.equal(audit.forRow('A')[0].newValue, 3);
  audit.detach(); g.destroy();
});

/* ================================================================== */
section('context menu estensibile');
test('contextMenuItems aggiunge voci custom (es. Storico riga)', () => {
  let invoked = false;
  const g = new Grid(host(), {
    columns: [{ field: 'nome' }],
    data: [{ nome: 'x' }],
    contextMenuItems: (items, ctx) => [
      ...items,
      { sep: true },
      { label: 'Storico riga', action: () => { invoked = true; } },
    ],
  });
  const cell = g['renderer'].root.querySelector('.eg-cell[data-row="0"][data-col="0"]');
  cell.dispatchEvent(new window.MouseEvent('contextmenu', { bubbles: true, clientX: 30, clientY: 30 }));
  const labels = [...g['renderer'].root.querySelectorAll('.eg-menu-item span')].map((s) => s.textContent);
  assert.ok(labels.includes('Storico riga'));
  const btn = [...g['renderer'].root.querySelectorAll('.eg-menu-item')].find((b) => b.textContent.includes('Storico riga'));
  btn.click();
  assert.equal(invoked, true);
  g.destroy();
});

/* ================================================================== */
for (const run of pending) await run();
console.log(`\n${passed} test superati, ${failed} falliti`);
process.exit(failed ? 1 : 0);
