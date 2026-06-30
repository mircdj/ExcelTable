/**
 * Native .xlsx import/export — zero dependencies.
 *
 * An .xlsx file is a ZIP of XML parts. Export writes a store-only ZIP
 * (no compression: maximum compatibility, still small for typical sheets)
 * with inline strings. Import reads both stored and deflated entries —
 * deflated ones are inflated with the browser's built-in
 * `DecompressionStream('deflate-raw')` (all evergreen browsers, Node 18+).
 *
 * Scope: one worksheet, values + headers, types number/boolean/date/text.
 * Styles, formulas and multiple sheets are intentionally out of scope here.
 */
import type { Grid } from './grid';
import type { RowData } from './types';

/* ================================================================== */
/* CRC32 (required by the ZIP format)                                  */
/* ================================================================== */

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(data: Uint8Array): number {
  let c = 0xffffffff;
  for (let i = 0; i < data.length; i++) c = CRC_TABLE[(c ^ data[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/* ================================================================== */
/* Minimal ZIP writer (store only)                                     */
/* ================================================================== */

const te = new TextEncoder();
const td = new TextDecoder();

function zipStore(files: { name: string; data: Uint8Array }[]): Uint8Array {
  const chunks: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;

  for (const f of files) {
    const name = te.encode(f.name);
    const crc = crc32(f.data);
    const local = new Uint8Array(30 + name.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true); // version
    lv.setUint16(8, 0, true); // method: store
    lv.setUint32(14, crc, true);
    lv.setUint32(18, f.data.length, true);
    lv.setUint32(22, f.data.length, true);
    lv.setUint16(26, name.length, true);
    local.set(name, 30);
    chunks.push(local, f.data);

    const cd = new Uint8Array(46 + name.length);
    const cv = new DataView(cd.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true);
    cv.setUint16(6, 20, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, f.data.length, true);
    cv.setUint32(24, f.data.length, true);
    cv.setUint16(28, name.length, true);
    cv.setUint32(42, offset, true);
    cd.set(name, 46);
    central.push(cd);
    offset += local.length + f.data.length;
  }

  const cdSize = central.reduce((s, c) => s + c.length, 0);
  const end = new Uint8Array(22);
  const ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, files.length, true);
  ev.setUint16(10, files.length, true);
  ev.setUint32(12, cdSize, true);
  ev.setUint32(16, offset, true);

  const total = offset + cdSize + 22;
  const out = new Uint8Array(total);
  let p = 0;
  for (const c of [...chunks, ...central, end]) {
    out.set(c, p);
    p += c.length;
  }
  return out;
}

/* ================================================================== */
/* Minimal ZIP reader (store + deflate via DecompressionStream)        */
/* ================================================================== */

async function zipRead(buf: ArrayBuffer): Promise<Map<string, Uint8Array>> {
  const u8 = new Uint8Array(buf);
  const dv = new DataView(buf);
  // Find End Of Central Directory.
  let eocd = -1;
  for (let i = u8.length - 22; i >= 0; i--) {
    if (dv.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error('File ZIP non valido');
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const files = new Map<string, Uint8Array>();

  for (let n = 0; n < count; n++) {
    if (dv.getUint32(p, true) !== 0x02014b50) break;
    const method = dv.getUint16(p + 10, true);
    const csize = dv.getUint32(p + 20, true);
    const nameLen = dv.getUint16(p + 28, true);
    const extraLen = dv.getUint16(p + 30, true);
    const commentLen = dv.getUint16(p + 32, true);
    const localOff = dv.getUint32(p + 42, true);
    const name = td.decode(u8.subarray(p + 46, p + 46 + nameLen));

    const lNameLen = dv.getUint16(localOff + 26, true);
    const lExtraLen = dv.getUint16(localOff + 28, true);
    const dataStart = localOff + 30 + lNameLen + lExtraLen;
    const raw = u8.slice(dataStart, dataStart + csize);

    if (method === 0) files.set(name, raw);
    else if (method === 8) files.set(name, await inflateRaw(raw));
    else throw new Error('Metodo di compressione ZIP non supportato: ' + method);

    p += 46 + nameLen + extraLen + commentLen;
  }
  return files;
}

async function inflateRaw(data: Uint8Array): Promise<Uint8Array> {
  if (typeof DecompressionStream === 'undefined')
    throw new Error('DecompressionStream non disponibile in questo ambiente');
  const ds = new DecompressionStream('deflate-raw');
  const stream = new Blob([data as BlobPart]).stream().pipeThrough(ds);
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/* ================================================================== */
/* xlsx export                                                         */
/* ================================================================== */

const xmlEsc = (s: string) =>
  s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c]!);

function colLetter(n: number): string {
  let s = '';
  n++;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = (n - m - 1) / 26;
  }
  return s;
}

/** Build an .xlsx (as bytes) from the grid's current view. */
export function exportXlsx<T extends RowData>(grid: Grid<T>, sheetName = 'Dati'): Uint8Array {
  const cols = grid.visibleColumns();
  const rowsXml: string[] = [];

  const headerCells = cols
    .map((c, i) => `<c r="${colLetter(i)}1" t="inlineStr"><is><t>${xmlEsc(c.header)}</t></is></c>`)
    .join('');
  rowsXml.push(`<row r="1">${headerCells}</row>`);

  let r = 2;
  for (let v = 0; v < grid.rowCount; v++) {
    if (grid.isGroupRow(v)) continue;
    const cells = cols
      .map((c, i) => {
        const raw = grid.getValue(v, c.id);
        const ref = `${colLetter(i)}${r}`;
        if (typeof raw === 'number' && Number.isFinite(raw)) return `<c r="${ref}"><v>${raw}</v></c>`;
        if (typeof raw === 'boolean') return `<c r="${ref}" t="b"><v>${raw ? 1 : 0}</v></c>`;
        const s = grid.getDisplayValue(v, c.id);
        return s === '' ? '' : `<c r="${ref}" t="inlineStr"><is><t>${xmlEsc(s)}</t></is></c>`;
      })
      .join('');
    rowsXml.push(`<row r="${r}">${cells}</row>`);
    r++;
  }

  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rowsXml.join('')}</sheetData></worksheet>`;

  const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${xmlEsc(sheetName)}" sheetId="1" r:id="rId1"/></sheets></workbook>`;

  const wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`;

  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`;

  return zipStore([
    { name: '[Content_Types].xml', data: te.encode(contentTypes) },
    { name: '_rels/.rels', data: te.encode(rels) },
    { name: 'xl/workbook.xml', data: te.encode(workbook) },
    { name: 'xl/_rels/workbook.xml.rels', data: te.encode(wbRels) },
    { name: 'xl/worksheets/sheet1.xml', data: te.encode(sheet) },
  ]);
}

/* ================================================================== */
/* xlsx import                                                         */
/* ================================================================== */

export interface ImportedSheet {
  headers: string[];
  /** Row objects keyed by header. */
  rows: RowData[];
}

/** Parse the first worksheet of an .xlsx file into rows keyed by header. */
export async function importXlsx(file: ArrayBuffer | Blob): Promise<ImportedSheet> {
  const buf = file instanceof Blob ? await file.arrayBuffer() : file;
  const files = await zipRead(buf);

  // Shared strings (Excel writes them; our export uses inline strings).
  const shared: string[] = [];
  const ss = files.get('xl/sharedStrings.xml');
  if (ss) {
    const xml = td.decode(ss);
    for (const m of xml.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
      const texts = [...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => unesc(t[1]));
      shared.push(texts.join(''));
    }
  }

  const sheetName =
    [...files.keys()].find((n) => /^xl\/worksheets\/sheet1\.xml$/.test(n)) ??
    [...files.keys()].find((n) => n.startsWith('xl/worksheets/'));
  if (!sheetName) throw new Error('Nessun foglio trovato nel file');
  const xml = td.decode(files.get(sheetName)!);

  // cell matrix: ref → value
  const matrix = new Map<string, unknown>();
  let maxRow = 0;
  let maxCol = 0;
  for (const m of xml.matchAll(/<c r="([A-Z]+)(\d+)"([^>]*)>([\s\S]*?)<\/c>/g)) {
    const colIdx = letterToIndex(m[1]);
    const rowIdx = Number(m[2]);
    const attrs = m[3];
    const body = m[4];
    const tMatch = /t="([^"]+)"/.exec(attrs);
    const t = tMatch?.[1];
    let value: unknown = null;
    const vMatch = /<v>([\s\S]*?)<\/v>/.exec(body);
    const isMatch = /<is>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>[\s\S]*?<\/is>/.exec(body);
    if (t === 'inlineStr' && isMatch) value = unesc(isMatch[1]);
    else if (t === 's' && vMatch) value = shared[Number(vMatch[1])] ?? '';
    else if (t === 'b' && vMatch) value = vMatch[1] === '1';
    else if (t === 'str' && vMatch) value = unesc(vMatch[1]);
    else if (vMatch) value = Number(vMatch[1]);
    matrix.set(`${colIdx},${rowIdx}`, value);
    if (rowIdx > maxRow) maxRow = rowIdx;
    if (colIdx > maxCol) maxCol = colIdx;
  }

  const headers: string[] = [];
  for (let c = 0; c <= maxCol; c++) headers.push(String(matrix.get(`${c},1`) ?? colLetter(c)));
  const rows: RowData[] = [];
  for (let r = 2; r <= maxRow; r++) {
    const row: RowData = {};
    let any = false;
    for (let c = 0; c <= maxCol; c++) {
      const v = matrix.get(`${c},${r}`);
      if (v !== undefined) any = true;
      row[headers[c]] = v ?? null;
    }
    if (any) rows.push(row);
  }
  return { headers, rows };
}

function letterToIndex(s: string): number {
  let n = 0;
  for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

const unesc = (s: string) =>
  s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
