<div align="center">

# ExcelTable

**An Excel-grade data grid for the web — zero learning curve.**

A pure-TypeScript core with **zero dependencies** and a virtualized DOM renderer, plus official wrappers for **Angular**, **React** and **Next.js**. Every gesture — selection, keyboard shortcuts, copy/paste, the fill handle — mirrors Excel exactly. If you know Excel, you already know ExcelTable.

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.x-3178c6.svg)](https://www.typescriptlang.org/)
[![Dependencies](https://img.shields.io/badge/dependencies-0-brightgreen.svg)](#)
[![Core size](https://img.shields.io/badge/core-~10kb%20gzip-blue.svg)](#)

[**Live demo**](https://mircdj.github.io/ExcelTable/demo/) · [**Documentation**](https://mircdj.github.io/ExcelTable/) · [**API reference**](docs/API.md)

</div>

---

## Table of contents

- [Why ExcelTable](#why-exceltable)
- [Packages](#packages)
- [Installation](#installation)
- [Quickstart](#quickstart)
  - [Vanilla JS / any framework](#vanilla-js--any-framework)
  - [React](#react)
  - [Next.js (App Router)](#nextjs-app-router)
  - [Angular (15+)](#angular-15-standalone-component)
- [Theming](#theming)
- [Features](#features)
- [Demo](#demo)
- [Documentation](#documentation)
- [Development](#development)
- [Documented limits](#documented-limits)
- [Contributing](#contributing)
- [License](#license)

## Why ExcelTable

- **Excel fidelity, not approximation.** Range selection, `Ctrl+Arrow` edge jumps, `Tab`/`Enter` cycling inside a range, fill-handle series, marching ants, Excel-interoperable TSV/HTML clipboard, `Ctrl+Enter` write-to-selection — the behaviors users expect are already there.
- **Fast by default.** Row *and* column virtualization renders 10,000+ rows at 60 fps with a minimal, recycled DOM.
- **Zero dependencies in the core.** ~10 kb gzipped, framework-agnostic, SSR-safe.
- **Typed end to end.** Every mutation emits a typed, often cancelable event.
- **Accessible.** A complete WCAG 2.1 AA ARIA grid pattern, visible focus, and `prefers-reduced-motion` support.

## Packages

| Package | Description |
| --- | --- |
| [`@exceltable/core`](packages/core) | The engine — pure TypeScript, 0 dependencies, ~10 kb gzip |
| [`@exceltable/react`](packages/react) | React / Next.js wrapper (SSR-safe) |
| [`@exceltable/angular`](packages/angular) | Angular 15+ wrapper (standalone component, zone-less rendering) |

## Installation

> **Note:** the packages are not published to npm yet. Use ExcelTable from source as shown below. The `@exceltable/*` import paths in the examples are the intended public API once the packages are published.

Clone the repository and build the core:

```bash
git clone https://github.com/mircdj/ExcelTable.git
cd ExcelTable
npm install
cd packages/core
npx esbuild src/index.ts --bundle --format=esm --outfile=dist/index.js
```

Then reference the built bundle and stylesheet from your app:

```ts
import { Grid } from '@exceltable/core';   // resolves to packages/core/dist/index.js
import '@exceltable/core/theme.css';        // packages/core/src/theme.css
```

The fastest way to see it running is the [self-contained demo](demo/index.html) — open it directly in a browser, no build required.

## Quickstart

### Vanilla JS / any framework

```ts
import { Grid } from '@exceltable/core';
import '@exceltable/core/theme.css';

const grid = new Grid(document.getElementById('grid')!, {
  columns: [
    { field: 'id', header: 'ID', pinned: 'left', editable: false },
    { field: 'zone', header: 'Zone', type: 'select', options: ['North', 'South'] },
    { field: 'hours', header: 'Hours', type: 'number',
      validator: (v) => (v as number) < 0 ? 'Hours cannot be negative' : null },
    { field: 'startDate', header: 'Start date', type: 'date' },
    { field: 'tested', header: 'Tested', type: 'boolean' },
  ],
  data: rows,
  getRowId: (r) => r.id,
  theme: 'excel',          // 'excel' | 'dark' | 'blue'
  locale: 'it-IT',
});

grid.events.on('cellsChanged', (e) => console.log(e.changes, e.source));
grid.events.on('dirtyStateChanged', (e) => updateSaveButton(e.dirtyRowIds.length));
```

### React

```tsx
import { ExcelTable, type ExcelTableHandle } from '@exceltable/react';
import '@exceltable/core/theme.css';

function Register() {
  const ref = useRef<ExcelTableHandle>(null);
  return (
    <ExcelTable
      columns={columns}
      data={rows}
      getRowId={(r) => r.id}
      onCellsChanged={(changes, source) => save(changes)}
      onDirtyStateChanged={(ids) => setDirtyCount(ids.length)}
      ref={ref}                      // ref.current.api → full GridApi
    />
  );
}
```

### Next.js (App Router)

The wrapper is SSR-safe: the engine is instantiated only inside `useEffect`, so you just need to mark the component as a client component.

```tsx
'use client';
import { ExcelTable } from '@exceltable/react';
import '@exceltable/core/theme.css';

export default function GridPage() {
  return <div style={{ height: '80vh' }}><ExcelTable columns={cols} data={rows} /></div>;
}
```

### Angular (15+, standalone component)

```ts
import { Component } from '@angular/core';
import { ExcelTableComponent } from '@exceltable/angular';

@Component({
  standalone: true,
  imports: [ExcelTableComponent],
  template: `
    <excel-table [columns]="cols" [data]="rows" theme="blue"
                 (cellsChanged)="onChanges($event)"
                 (dirtyStateChanged)="dirty = $event.length" />
  `,
  styles: [':host { display:block; height: 80vh; }'],
})
export class RegisterComponent { /* … */ }
```

The grid runs **outside the Angular zone**: scrolling never triggers change detection; only `@Output`s re-enter the zone.

> Import the CSS once in `styles.css`: `@import '@exceltable/core/theme.css';`

## Theming

Three built-in themes — `excel` (default, green accent), `dark`, and `blue` — plus a complete CSS custom-property surface. Override a handful of `--eg-*` variables to brand the grid, or define a whole new theme class and switch it at runtime by swapping the root class. See the [theming guide](docs/API.md#theming).

## Features

<details>
<summary><strong>Selection, navigation &amp; editing</strong></summary>

- **Excel selection**: drag ranges, Shift+click, multiple ranges with Ctrl, whole row from the row number, `Ctrl+A`
- **Excel navigation**: arrows, `Ctrl+Arrow` (jump to data edge), `Shift` to extend, `Tab`/`Enter` cycling inside the selected range, Home/End, PageUp/Down
- **Excel editing**: `F2`, double-click, type-to-overwrite, `Esc`, `Alt+Enter` newline, **`Ctrl+Enter` writes across the whole selection**, single-key boolean toggle
- **Context menu** (right click): copy/cut/paste, clear, insert/delete rows, group, autofit, hide/show columns
</details>

<details>
<summary><strong>Clipboard &amp; fill</strong></summary>

- **Excel-interoperable clipboard**: bidirectional TSV copy/cut/paste, pattern repetition over the selection, marching ants, cancelable `pasteStart` / `copyStart` events
- **Copy as HTML**: the clipboard carries TSV *and* a `<table>` — pasting into Outlook/Word arrives formatted
- **Fill handle**: numeric series with step, date series, text suffixes (`Cabin 7` → `Cabin 8`, `POD-009` → `POD-010`), pattern repetition
</details>

<details>
<summary><strong>Columns, sorting &amp; filtering</strong></summary>

- **Columns**: freeze left (pinned), drag-resize, **double-click-border autofit**, show/hide, min/max widths, `moveColumn` API
- **Drag &amp; drop columns**: reorder by dragging headers (ghost + drop indicator); click remains sort
- **Multi-column sort** (Shift+click) with natural ordering (`CAB2` < `CAB10`)
- **Excel-style value filter**: checkbox dropdown with search and "Select all"
</details>

<details>
<summary><strong>Data model</strong></summary>

- **Undo/Redo**: unlimited transactional undo (`Ctrl+Z` / `Ctrl+Y`)
- **Per-column validation** with a `validationFailed` event
- **Dirty tracking**: changed rows are flagged, `getDirtyRows()` for saving, `markClean()`
- **Undoable row insert/delete** (`insertRows` / `removeRows`) with `rowsInserted` / `rowsRemoved` events
- **Find &amp; Replace**: `Ctrl+F` / `Ctrl+H`, match count, navigate, replace one/all — undoable; `findAll` / `replaceIn` API
- **Saveable views**: `getState()` / `setState()` serialize sort, filters, widths, hidden columns and grouping
- **Locale-aware parsing** (`it-IT` default): `1.234,56`, `50%`, `01/02/2026`, `yes`
- Typed events on every mutation, many cancelable — see [docs/API.md](docs/API.md)
</details>

<details>
<summary><strong>Advanced</strong></summary>

- **Adapt to any JSON**: `Grid.fromJson(el, json, { include, exclude, overrides })` infers columns and types, flattens nested objects to dot-paths (`customer.city`) and binds nested fields for read *and* write
- **Nested sub-tables (master/detail)**: arrays of objects become expandable panels (▸ on the row number) holding a full ExcelTable; exact virtualization even with panels open, `rowExpanded` / `rowCollapsed` events, nested edits tracked in the parent's dirty set
- **Row grouping** (`groupBy` / `setGroupBy`, multi-level), collapsible groups, per-column **aggregations** (`aggFunc: sum|avg|min|max|count`) on group rows and a pinned **Σ totals row**
- **Pivot**: `pivot(data, columns, config)` — pure rows×columns transform with multiple row dimensions, a column dimension, multiple aggregated values and row/column totals; the result renders in a normal ExcelTable
- **Formulas** (`formulas: true`): `=SUM(C2:C10)` cells, A1 references and ranges, operators and comparisons, 16 functions (Italian and English names), cascading recalculation with caching and cycle detection (`#CIRC!`); the editor shows the source, sort and aggregations use computed values, all undoable
- **Server-side row model** (`serverSide: { getRows }`): lazy block loading with an LRU cache, stale-response rejection, sort/filter delegated to the backend, `…` placeholders, editing and dirty tracking on loaded rows
- **Tree data** (`treeData`): WBS hierarchies with indentation and chevrons; sibling sort, ancestor-preserving filter
- **Offline** (`attachOfflinePersistence`): local queue of unsaved changes, restore on restart, `flush()` to the backend with all-or-nothing semantics
- **Audit trail** (`attachAuditTrail`): who/when/what per row, configurable limit, extendable context menu
- **Native `.xlsx` import/export, zero dependencies**: `exportXlsx(grid)` writes a real Excel file (ZIP+XML); `importXlsx(file)` reads Excel sheets back (shared &amp; inline strings, store and deflate via `DecompressionStream`)
- **Custom cells**: `cellRenderer` returns text, DOM nodes or framework components — `reactCell()` for React, `templateCell()` for Angular, with mount/unmount handled by the virtualizer
- **Multi-content cells**: a renderer may return an array — a value on one line and, on the next, a custom element or more text (`['75% complete', progressBar]`); `wrapText` for wrapping
- **Variable row height** (`getRowHeight`) with prefix-sum offsets and exact virtualization
</details>

<details>
<summary><strong>Platform &amp; accessibility</strong></summary>

- **Accessibility WCAG 2.1 AA**: complete ARIA grid pattern, `aria-activedescendant`, `aria-sort`, labels on every control, ≥ 4.5:1 contrast, visible focus, `prefers-reduced-motion`
- **Full touch** (pointer events): tap to select, double-tap to edit, long-press for the menu, long-press+drag to extend selection, draggable fill handle
- **Responsive**: the grid fills its container with smooth vertical and horizontal scrolling; enlarged touch targets on mobile
- **i18n** (`strings`): every UI string in a typed table, Italian by default, `EN` included, partial overrides
- **CSV/TSV export**
</details>

## Demo

A self-contained demo lives in [`demo/index.html`](demo/index.html) — open it directly in a browser, no server required. It showcases four scenarios:

1. **Work register** — 10,000 rows with nested material sub-tables (▸ on the row number), undo/redo toolbar, CSV import/export, themes, save-changed-rows, a sum/average status bar
2. **Custom & multi-content cells** — status badges, a "value + inline bar" column, action buttons
3. **Any JSON** — paste arbitrary JSON, choose include/exclude, `Grid.fromJson` builds the grid with automatic sub-tables
4. **Events & dispatch** — a live event console plus buttons that call the public API, including the `pasteStart` veto

▶️ **Try it online: [mircdj.github.io/ExcelTable/demo](https://mircdj.github.io/ExcelTable/demo/)**

## Documentation

- 📖 **Full docs site:** [mircdj.github.io/ExcelTable](https://mircdj.github.io/ExcelTable/)
- 📑 **API reference:** [docs/API.md](docs/API.md) — complete API, events and theming

The documentation site is published automatically to GitHub Pages on every push to `main` (see [`.github/workflows/deploy-docs.yml`](.github/workflows/deploy-docs.yml)).

## Development

```bash
npm install
```

**Build the core:**

```bash
cd packages/core
npx tsc                                                          # type-check + declarations
npx esbuild src/index.ts --bundle --format=esm --outfile=dist/index.js
```

**Run tests:**

```bash
node tests/run.mjs    # 106 tests: series, parsing, JSON inference, data pipeline,
                      # undo, master/detail, custom renderers, ARIA contract
npm run e2e           # Playwright end-to-end (mouse, keyboard, touch, clipboard, layout)
```

## Documented limits

Deliberate semantic choices worth knowing:

- **Formula references** are positional on data coordinates (sorting does not change them; inserting/deleting rows recalculates but does not rewrite references — no Excel-style relative adjustment).
- In **tree mode**, row insert/delete and grouping are disabled (the hierarchy owns the structure).
- In **server mode**, local grouping, row insert/delete and fill on unloaded rows are disabled (the dataset belongs to the backend), and the value filter lists only already-loaded values.
- The **xlsx** support covers a single sheet's values and headers (no styles/formulas/multi-sheet).

## Contributing

Contributions are welcome. Please open an issue to discuss substantial changes first, keep the core dependency-free, and make sure `node tests/run.mjs` and the Playwright suite pass before opening a pull request.

## License

[MIT](LICENSE) © Mirco Sabatino
