/**
 * i18n — every user-facing string in one typed table.
 * Italian is the default; `EN` ships ready; any partial override via
 * `GridOptions.strings`. `{name}`-style placeholders are interpolated.
 */

export interface Strings {
  gridLabel: string;
  rowNumber: string;
  expandRow: string; // {n}
  detailOf: string; // {n}
  group: string; // {name}
  totalsRow: string;
  totalsTitle: string; // {fn} {name} {n}
  filterLabel: string; // {name}
  filterDialog: string;
  filterSearch: string;
  filterSearchLabel: string;
  filterSelectAll: string;
  filterEmpty: string;
  filterClear: string;
  filterOk: string;
  findTitle: string;
  findPlaceholder: string;
  findReplacePlaceholder: string;
  findPrev: string;
  findNext: string;
  findClose: string;
  findCase: string;
  findReplace: string;
  findReplaceAll: string;
  menuCopy: string;
  menuCut: string;
  menuPaste: string;
  menuClear: string;
  menuInsertAbove: string;
  menuInsertBelow: string;
  menuDeleteRow: string;
  menuDeleteRows: string; // {n}
  menuAutofit: string; // {name}
  menuHide: string; // {name}
  menuSortAsc: string;
  menuSortDesc: string;
  menuGroupBy: string; // {name}
  menuUngroup: string;
  menuShowHidden: string; // {n}
  keyDelete: string;
  selectNotInList: string;
  loadingCell: string;
}

export const IT: Strings = {
  gridLabel: 'Tabella dati',
  rowNumber: 'Numero riga',
  expandRow: 'Espandi dettaglio riga {n}',
  detailOf: 'Dettaglio riga {n}',
  group: 'Gruppo {name}',
  totalsRow: 'Riga dei totali',
  totalsTitle: '{fn} di {name} su {n} righe',
  filterLabel: 'Filtra {name}',
  filterDialog: 'Filtro per valori',
  filterSearch: 'Cerca…',
  filterSearchLabel: 'Cerca valori',
  filterSelectAll: '(Seleziona tutto)',
  filterEmpty: '(vuote)',
  filterClear: 'Cancella filtro',
  filterOk: 'OK',
  findTitle: 'Trova e sostituisci',
  findPlaceholder: 'Trova…',
  findReplacePlaceholder: 'Sostituisci con…',
  findPrev: 'Precedente',
  findNext: 'Successivo',
  findClose: 'Chiudi',
  findCase: 'Maiuscole/minuscole',
  findReplace: 'Sostituisci',
  findReplaceAll: 'Tutto',
  menuCopy: 'Copia',
  menuCut: 'Taglia',
  menuPaste: 'Incolla',
  menuClear: 'Svuota celle',
  menuInsertAbove: 'Inserisci riga sopra',
  menuInsertBelow: 'Inserisci riga sotto',
  menuDeleteRow: 'Elimina riga',
  menuDeleteRows: 'Elimina {n} righe',
  menuAutofit: 'Autofit "{name}"',
  menuHide: 'Nascondi "{name}"',
  menuSortAsc: 'Ordina A→Z',
  menuSortDesc: 'Ordina Z→A',
  menuGroupBy: 'Raggruppa per "{name}"',
  menuUngroup: 'Rimuovi raggruppamento',
  menuShowHidden: 'Mostra colonne nascoste ({n})',
  keyDelete: 'Canc',
  selectNotInList: 'Valore non presente in elenco',
  loadingCell: '…',
};

export const EN: Strings = {
  gridLabel: 'Data grid',
  rowNumber: 'Row number',
  expandRow: 'Expand row {n} details',
  detailOf: 'Row {n} details',
  group: 'Group {name}',
  totalsRow: 'Totals row',
  totalsTitle: '{fn} of {name} over {n} rows',
  filterLabel: 'Filter {name}',
  filterDialog: 'Filter by values',
  filterSearch: 'Search…',
  filterSearchLabel: 'Search values',
  filterSelectAll: '(Select all)',
  filterEmpty: '(blanks)',
  filterClear: 'Clear filter',
  filterOk: 'OK',
  findTitle: 'Find and replace',
  findPlaceholder: 'Find…',
  findReplacePlaceholder: 'Replace with…',
  findPrev: 'Previous',
  findNext: 'Next',
  findClose: 'Close',
  findCase: 'Match case',
  findReplace: 'Replace',
  findReplaceAll: 'All',
  menuCopy: 'Copy',
  menuCut: 'Cut',
  menuPaste: 'Paste',
  menuClear: 'Clear cells',
  menuInsertAbove: 'Insert row above',
  menuInsertBelow: 'Insert row below',
  menuDeleteRow: 'Delete row',
  menuDeleteRows: 'Delete {n} rows',
  menuAutofit: 'Autofit "{name}"',
  menuHide: 'Hide "{name}"',
  menuSortAsc: 'Sort A→Z',
  menuSortDesc: 'Sort Z→A',
  menuGroupBy: 'Group by "{name}"',
  menuUngroup: 'Remove grouping',
  menuShowHidden: 'Show hidden columns ({n})',
  keyDelete: 'Del',
  selectNotInList: 'Value not in list',
  loadingCell: '…',
};

export function interpolate(template: string, params?: Record<string, string | number>): string {
  if (!params) return template;
  return template.replace(/\{(\w+)\}/g, (_, k) => String(params[k] ?? ''));
}
