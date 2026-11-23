/* =============================================================
   Rellenado CRM · Complemento de Excel + modo web local
   Todo ocurre en el navegador: sin servidores ni bases de datos.
   Trabaja sobre VARIAS hojas a la vez (BD-BBPE y Seguimiento semanal)
   con las mismas reglas.
   ============================================================= */
'use strict';

/* ---------------- Configuración ---------------- */
const RESPONSABLES = ['Cinthya Sandoval', 'Fiorella Viena', 'Harold Sovero', 'Lorena Suncion'];
const DEFAULTS = {
  sheet: 'BD-BBPE',
  idTarget: 'Customer ID',
  idCrm: 'Código de Cliente',
  dateCol: 'FECHA CONTACTO',
  respCol: 'RESPONSABLE',
  dateFmt: 'mm/dd/yyyy',
};
// Hojas que se proponen por defecto: BD-BBPE y cualquier hoja de "seguimiento"
const DEFAULT_SHEET_PATTERNS = [/^bdbbpe$/, /seguimiento/];
// Nombres alternativos que también se aceptan como columna identificadora
const ID_ALIASES = ['customerid', 'codigodecliente', 'codigocliente', 'codcliente'];
const HEADER_SCAN_ROWS = 10;
const LS_KEY = 'rellenado_crm_v1';
// Acción 2: ÚNICAS columnas que se rellenan desde el archivo CRM
const A2_COLUMNS = ['ESTADO CONTACTO', 'PRÓXIMA ACCIÓN', 'OBSERVACIÓN', 'NOMBRE DEL CONTACTO', 'CORREO', 'TELÉFONO'];
// Columnas que JAMÁS se escriben en filas existentes
const PROTECTED_KEYS = ['customerid', 'codigodecliente', 'codigocliente', 'codcliente', 'razonsocial', 'numerodecontacto', 'numerocontacto', 'numero', 'nrodecontacto', 'celular'];
// Solo en FILAS NUEVAS (códigos agregados) se permite escribir además estas columnas
const NEW_ROW_RAZON = ['razonsocial'];
const NEW_ROW_NUMERO = ['numerodecontacto', 'numerocontacto', 'numero', 'nrodecontacto', 'celular'];
const DEFAULT_EXCLUDED = [];

/* ---------------- Utilidades ---------------- */
const $ = (s) => document.querySelector(s);
const $$ = (s) => [...document.querySelectorAll(s)];
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const pad = (n) => String(n).padStart(2, '0');
const normKey = (h) => String(h ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
const isEmpty = (v) => v === null || v === undefined || (typeof v === 'string' && v.trim() === '');

function codeKey(v) {
  if (v === null || v === undefined) return '';
  let s = String(v).trim().replace(/^'/, '').replace(/\s+/g, '').toUpperCase();
  if (/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, '');
  return s;
}
/** Valor a escribir como Customer ID en una fila nueva (número si es numérico sin ceros a la izquierda) */
function codeCellValue(raw) {
  const t = String(raw).trim();
  return /^[1-9]\d{0,14}$/.test(t) ? Number(t) : t;
}
function todayISO() { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }
// Número de serie de Excel (días desde 30/12/1899), sin desfases de zona horaria
function excelSerial(iso) { const [y, m, d] = iso.split('-').map(Number); return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000; }
function fmtDate(iso, fmt) {
  const [y, m, d] = iso.split('-');
  return fmt.toLowerCase().replace('yyyy', y).replace('mm', m).replace('dd', d);
}
function toast(msg, type = 'ok', ms = 2400) {
  const c = { ok: 'bg-slate-900', success: 'bg-emerald-600', error: 'bg-rose-600', warn: 'bg-amber-500' }[type] || 'bg-slate-900';
  const el = document.createElement('div');
  el.className = `toast pointer-events-auto ${c} text-white text-[12.5px] font-medium px-3 py-2 rounded-lg shadow-lg max-w-[380px]`;
  el.textContent = msg; $('#toasts').appendChild(el);
  setTimeout(() => { el.style.transition = 'opacity .2s'; el.style.opacity = '0'; setTimeout(() => el.remove(), 220); }, ms);
}
async function copyText(t) {
  try { await navigator.clipboard.writeText(t); return true; } catch (e) {
    const ta = document.createElement('textarea'); ta.value = t; ta.style.position = 'fixed'; ta.style.opacity = '0';
    document.body.appendChild(ta); ta.select(); let ok = false; try { ok = document.execCommand('copy'); } catch (e2) {} ta.remove(); return ok;
  }
}
function colLetter(c) { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }

/* ---------------- Estado persistente (solo preferencias) ---------------- */
let P = (() => {
  const base = { ...DEFAULTS, resp: '', codes: '', tab: 'a1', excluded: DEFAULT_EXCLUDED.slice(), sheets: [], addRazon: true };
  try { const p = { ...base, ...JSON.parse(localStorage.getItem(LS_KEY) || '{}') }; if (!Array.isArray(p.sheets)) p.sheets = []; return p; } catch (e) { return base; }
})();
function persist() { try { localStorage.setItem(LS_KEY, JSON.stringify(P)); } catch (e) {} }

/* ---------------- Estado de sesión ---------------- */
const APP = {
  mode: 'detecting',   // 'office' | 'local'
  master: null,        // modo local: { wb, fileName, changes }
  crm: null,           // { fileName, headers[], idIdx, rows[][] }
  undo: null,          // { label, writes[] }   (cada write lleva su hoja)
  sheetNames: [],      // hojas del libro
  missing: { 1: null, 2: null }, // códigos no encontrados pendientes de agregar, por acción
  busy: false,
};
const selectedSheets = () => P.sheets.filter((s) => !APP.sheetNames.length || APP.sheetNames.includes(s));

/* =============================================================
   DETECCIÓN DEL ENCABEZADO
   ============================================================= */
function idAliases(custom) { return [normKey(custom), ...ID_ALIASES].filter(Boolean); }
/** rows: matriz con las primeras filas. Devuelve {rowOffset, idCol, cols: Map(normKey -> {col, name})} */
function detectHeader(rows, aliasesList) {
  for (let r = 0; r < rows.length; r++) {
    const keys = rows[r].map(normKey);
    let idCol = -1;
    for (const a of aliasesList) { idCol = keys.indexOf(a); if (idCol >= 0) break; }
    if (idCol < 0) continue;
    const cols = new Map();
    keys.forEach((k, c) => { if (k && !cols.has(k)) cols.set(k, { col: c, name: String(rows[r][c]).trim() }); });
    return { rowOffset: r, idCol, cols };
  }
  return null;
}

/* =============================================================
   DESTINO: EXCEL ONLINE (Office.js)
   ============================================================= */
const OfficeTarget = {
  async sheetNames() {
    return Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets; sheets.load('items/name');
      const active = ctx.workbook.worksheets.getActiveWorksheet(); active.load('name');
      await ctx.sync();
      return { names: sheets.items.map((s) => s.name), active: active.name };
    });
  },
  async getSheet(ctx, name) {
    const ws = ctx.workbook.worksheets.getItemOrNullObject(name);
    await ctx.sync();
    if (ws.isNullObject) throw new Error(`No existe la hoja "${name}" en este libro. Revisa ⚙ Configuración.`);
    return ws;
  },
  async scanHeader(ctx, ws, name) {
    const used = ws.getUsedRange(true); // solo celdas con valores: la última fila escrita
    used.load('rowIndex,columnIndex,rowCount,columnCount');
    await ctx.sync();
    const scan = ws.getRangeByIndexes(used.rowIndex, used.columnIndex, Math.min(HEADER_SCAN_ROWS, used.rowCount), used.columnCount);
    scan.load('values');
    await ctx.sync();
    const h = detectHeader(scan.values, idAliases(P.idTarget));
    if (!h) throw new Error(`No se encontró la columna "${P.idTarget}" en las ${HEADER_SCAN_ROWS} primeras filas de "${name}".`);
    return { used, h };
  },
  async headers(name) {
    return Excel.run(async (ctx) => {
      const ws = await this.getSheet(ctx, name);
      const { h } = await this.scanHeader(ctx, ws, name);
      return h;
    });
  },
  /** Carga la columna ID y las columnas pedidas (fórmulas = valor original exacto). */
  async buildIndex(neededKeys, name) {
    return Excel.run(async (ctx) => {
      const ws = await this.getSheet(ctx, name);
      const { used, h } = await this.scanHeader(ctx, ws, name);
      const headerRow = used.rowIndex + h.rowOffset;
      const first = headerRow + 1;
      const last = used.rowIndex + used.rowCount - 1;
      const n = last - first + 1;
      const idx = newIndex(name, h, used.columnIndex, first, Math.max(last, headerRow));
      if (n <= 0) return idx;
      const idRange = ws.getRangeByIndexes(first, used.columnIndex + h.idCol, n, 1); idRange.load('values');
      const loaded = [];
      for (const k of neededKeys) {
        const c = h.cols.get(k); if (!c) continue;
        const r = ws.getRangeByIndexes(first, used.columnIndex + c.col, n, 1); r.load('formulas');
        loaded.push([k, r]);
      }
      await ctx.sync();
      idRange.values.forEach((row, i) => addRow(idx, row[0], first + i));
      for (const [k, r] of loaded) idx.existing.set(k, r.formulas.map((x) => x[0]));
      return idx;
    });
  },
  /** writes: [{row, col, value, fmt?, restore?}] – filas/columnas absolutas de la hoja */
  async write(writes, name, onProgress) {
    const runs = groupRuns(writes);
    await Excel.run(async (ctx) => {
      const ws = await this.getSheet(ctx, name);
      let done = 0;
      for (const run of runs) {
        const rng = ws.getRangeByIndexes(run.row, run.col, run.items.length, 1);
        if (run.items.some((w) => w.fmt)) rng.numberFormat = run.items.map((w) => [w.fmt || 'General']);
        rng.formulas = run.items.map((w) => [officeSafe(w)]);
        done += run.items.length;
        if (done % 400 < run.items.length) { await ctx.sync(); onProgress && onProgress(done / writes.length); }
      }
      await ctx.sync();
      onProgress && onProgress(1);
    });
  },
};
// Evita que un texto que empiece por "=" se convierta en fórmula (salvo al deshacer)
function officeSafe(w) {
  const v = w.value;
  if (v === null || v === undefined) return '';
  if (!w.restore && typeof v === 'string' && /^[=+\-@]/.test(v) && isNaN(Number(v))) return "'" + v;
  return v;
}
function groupRuns(writes) {
  const byCol = new Map();
  for (const w of writes) { if (!byCol.has(w.col)) byCol.set(w.col, []); byCol.get(w.col).push(w); }
  const runs = [];
  for (const [col, list] of byCol) {
    list.sort((a, b) => a.row - b.row);
    let cur = null;
    for (const w of list) {
      if (cur && w.row === cur.row + cur.items.length && cur.items.length < 2000) cur.items.push(w);
      else { cur = { col, row: w.row, items: [w] }; runs.push(cur); }
    }
  }
  return runs;
}

/* =============================================================
   DESTINO: ARCHIVO LOCAL (SheetJS, en memoria)
   ============================================================= */
const LocalTarget = {
  ws(name) {
    if (!APP.master) throw new Error('Primero carga el Excel maestro (arriba).');
    const ws = APP.master.wb.Sheets[name];
    if (!ws) throw new Error(`El archivo no tiene la hoja "${name}". Revisa ⚙ Configuración.`);
    return ws;
  },
  scan(ws, name) {
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const rows = [];
    for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + HEADER_SCAN_ROWS - 1); r++) {
      const row = [];
      for (let c = range.s.c; c <= range.e.c; c++) { const cell = ws[XLSX.utils.encode_cell({ r, c })]; row.push(cell ? (cell.w ?? cell.v ?? '') : ''); }
      rows.push(row);
    }
    const h = detectHeader(rows, idAliases(P.idTarget));
    if (!h) throw new Error(`No se encontró la columna "${P.idTarget}" en las ${HEADER_SCAN_ROWS} primeras filas de "${name}".`);
    return { range, h };
  },
  /** Última fila que tiene algún valor escrito (ignora celdas solo con formato) */
  lastValueRow(ws) {
    let last = -1;
    for (const k of Object.keys(ws)) {
      if (k[0] === '!') continue;
      const cell = ws[k]; if (!cell || isEmpty(cell.v)) continue;
      const r = XLSX.utils.decode_cell(k).r; if (r > last) last = r;
    }
    return last;
  },
  async headers(name) { const ws = this.ws(name); return this.scan(ws, name).h; },
  async buildIndex(neededKeys, name) {
    const ws = this.ws(name);
    const { range, h } = this.scan(ws, name);
    const first = range.s.r + h.rowOffset + 1;
    const idx = newIndex(name, h, range.s.c, first, Math.max(this.lastValueRow(ws), first - 1));
    const cols = neededKeys.map((k) => [k, h.cols.get(k)]).filter(([, c]) => c);
    for (const [k] of cols) idx.existing.set(k, []);
    for (let r = first; r <= range.e.r; r++) {
      const idCell = ws[XLSX.utils.encode_cell({ r, c: range.s.c + h.idCol })];
      addRow(idx, idCell ? idCell.v : '', r);
      for (const [k, c] of cols) { const cell = ws[XLSX.utils.encode_cell({ r, c: range.s.c + c.col })]; idx.existing.get(k).push(cell ? cell.v : ''); }
    }
    return idx;
  },
  async write(writes, name, onProgress) {
    const ws = this.ws(name);
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    writes.forEach((w, i) => {
      const addr = XLSX.utils.encode_cell({ r: w.row, c: w.col });
      if (w.restore) { if (w.value === undefined || w.value === '') delete ws[addr]; else ws[addr] = JSON.parse(JSON.stringify(w.value)); }
      else {
        const v = w.value;
        if (typeof v === 'number') ws[addr] = { t: 'n', v, ...(w.fmt ? { z: w.fmt } : (ws[addr]?.z ? { z: ws[addr].z } : {})) };
        else if (typeof v === 'boolean') ws[addr] = { t: 'b', v };
        else ws[addr] = { t: 's', v: String(v) };
        if (w.row > range.e.r) range.e.r = w.row;   // filas nuevas agrandan el rango de la hoja
        if (w.col > range.e.c) range.e.c = w.col;
      }
      if (i % 500 === 0 && onProgress) onProgress(i / writes.length);
    });
    ws['!ref'] = XLSX.utils.encode_range(range);
    APP.master.changes += writes.length;
    onProgress && onProgress(1);
    renderFooter();
  },
  snapshot(name, row, col) {
    const cell = this.ws(name)[XLSX.utils.encode_cell({ r: row, c: col })];
    return cell ? JSON.parse(JSON.stringify(cell)) : undefined;
  },
};

/* ---------------- Índice común ---------------- */
function newIndex(sheet, h, colOffset, firstRow, lastRow) {
  return { sheet, h, colOffset, firstRow, lastRow, rows: new Map(), existing: new Map() };
}
function addRow(idx, rawId, absRow) {
  const k = codeKey(rawId); if (!k) return;
  if (!idx.rows.has(k)) idx.rows.set(k, []);
  idx.rows.get(k).push(absRow);
}
const absCol = (idx, key) => idx.colOffset + idx.h.cols.get(key).col;
const existingAt = (idx, key, absRow) => (idx.existing.get(key) || [])[absRow - idx.firstRow];
const T = () => (APP.mode === 'office' ? OfficeTarget : LocalTarget);
const findKey = (idx, keys) => keys.find((k) => idx.h.cols.has(k));

/** Lee todas las hojas seleccionadas; las que fallan se informan y se omiten */
async function buildIndexes(neededKeys) {
  const sheets = selectedSheets();
  if (!sheets.length) throw new Error('Marca al menos una hoja en ⚙ Configuración.');
  const idxs = [], errores = [];
  for (const s of sheets) {
    try { idxs.push(await T().buildIndex(neededKeys, s)); } catch (e) { errores.push(e.message); }
  }
  if (!idxs.length) throw new Error(errores.join(' '));
  return { idxs, errores };
}

/** Columnas en las que cada acción tiene permiso de escribir (filas existentes) */
function allowedKeys(action) {
  const keys = action === 1 ? [P.dateCol, P.respCol] : A2_COLUMNS;
  const blocked = new Set([...PROTECTED_KEYS, ...idAliases(P.idTarget), ...idAliases(P.idCrm)]);
  return new Set(keys.map(normKey).filter((k) => !blocked.has(k)));
}
/** Aplica escrituras (agrupadas por hoja) y prepara el "deshacer".
 *  Candado final: solo columnas permitidas y solo celdas vacías.
 *  En filas NUEVAS se permite además el ID (y Razón Social / Número si se pidió). */
async function applyWrites(writes, label, progEl, action) {
  const allowed = allowedKeys(action);
  const extraNueva = new Set([...idAliases(P.idTarget), ...NEW_ROW_RAZON, ...NEW_ROW_NUMERO]);
  for (const w of writes) {
    const ok = allowed.has(w.key) || (w.nueva && extraNueva.has(w.key));
    if (!ok) throw new Error(`Bloqueado: la columna "${w.key}" no está permitida. No se escribió nada.`);
    if (!isEmpty(w.old)) throw new Error('Bloqueado: se intentó escribir sobre una celda con datos. No se escribió nada.');
  }
  const porHoja = new Map();
  for (const w of writes) { if (!porHoja.has(w.sheet)) porHoja.set(w.sheet, []); porHoja.get(w.sheet).push(w); }
  const undoWrites = writes.map((w) => ({
    sheet: w.sheet, row: w.row, col: w.col, restore: true,
    value: APP.mode === 'office' ? (isEmpty(w.old) ? '' : w.old) : LocalTarget.snapshot(w.sheet, w.row, w.col),
    fmt: null,
  }));
  let hecho = 0;
  for (const [hoja, lista] of porHoja) {
    await T().write(lista, hoja, (p) => setProgress(progEl, (hecho + p * lista.length) / writes.length, `Escribiendo en ${hoja}…`));
    hecho += lista.length;
  }
  APP.undo = { label, writes: undoWrites };
  renderFooter();
}

/* =============================================================
   ACCIÓN 1 · Códigos + fecha + responsable
   ============================================================= */
function parseCodes(text) {
  const lines = String(text || '').split(/\r?\n/);
  const seen = new Set(), list = []; let empty = 0, dup = 0;
  for (const line of lines) {
    const raw = line.split('\t')[0];
    const k = codeKey(raw);
    if (!k) { empty++; continue; }
    if (seen.has(k)) { dup++; continue; }
    seen.add(k); list.push({ key: k, raw: raw.trim() });
  }
  return { list, empty, dup };
}
function renderCodesCounter() {
  const r = parseCodes($('#codes').value);
  $('#codesCount').textContent = r.list.length;
  const parts = [];
  if (r.dup) parts.push(`${r.dup} duplicado(s) quitados`);
  if (r.empty && r.list.length) parts.push(`${r.empty} línea(s) vacía(s)`);
  $('#codesExtra').textContent = parts.join(' · ');
}
function renderRespChips() {
  $('#respChips').innerHTML = RESPONSABLES.map((n) => `<button class="chip ${P.resp === n ? 'on' : ''}" data-resp="${esc(n)}">${esc(n)}</button>`).join('');
}
function renderDatePreview() {
  const v = $('#date').value;
  $('#datePreview').textContent = v ? `${fmtDate(v, P.dateFmt)}  (${P.dateFmt})` : '—';
}

async function runAction1() {
  const { list } = parseCodes($('#codes').value);
  const date = $('#date').value;
  if (!list.length) return toast('Pega al menos un código de cliente.', 'warn');
  if (!date) return toast('Elige una fecha.', 'warn');
  if (!P.resp) return toast('Marca un responsable.', 'warn');
  const kDate = normKey(P.dateCol), kResp = normKey(P.respCol);
  const allowed = allowedKeys(1);
  if (!allowed.has(kDate) || !allowed.has(kResp)) return toast('La columna de fecha o de responsable apunta a una columna protegida. Revisa ⚙ Configuración.', 'error', 4500);

  await busy('#btnRun1', async () => {
    const prog = $('#prog1');
    setProgress(prog, 0.05, 'Leyendo las hojas…');
    const { idxs, errores } = await buildIndexes([kDate, kResp]);
    const serial = excelSerial(date);
    const writes = [], porHoja = {}, notas = [...errores];
    for (const idx of idxs) {
      const faltan = [[kDate, P.dateCol], [kResp, P.respCol]].filter(([k]) => !idx.h.cols.has(k)).map(([, n]) => n);
      porHoja[idx.sheet] = { encontrados: 0, filas: 0, celdas: 0, conDatos: 0, repetidos: 0, faltan };
      if (faltan.length) notas.push(`"${idx.sheet}" no tiene la(s) columna(s) ${faltan.join(', ')}: no se escribió ahí.`);
    }
    const notFound = [];
    for (const { key, raw } of list) {
      const donde = idxs.filter((idx) => idx.rows.has(key));
      if (!donde.length) { notFound.push(raw); continue; }
      for (const idx of donde) {
        const st = porHoja[idx.sheet]; st.encontrados++;
        const rows = idx.rows.get(key);
        if (rows.length > 1) st.repetidos++;
        if (st.faltan.length) continue;
        for (const row of rows) {
          let tocada = false;
          for (const [k, value, fmt] of [[kDate, serial, P.dateFmt], [kResp, P.resp, null]]) {
            const old = existingAt(idx, k, row);
            if (!isEmpty(old)) { st.conDatos++; continue; } // nunca se sobrescribe
            writes.push({ sheet: idx.sheet, row, col: absCol(idx, k), key: k, value, fmt, old });
            st.celdas++; tocada = true;
          }
          if (tocada) st.filas++;
        }
      }
    }
    if (writes.length) await applyWrites(writes, `Acción 1 (${writes.length} celdas)`, prog, 1);
    setProgress(prog, 1, 'Listo');
    APP.missing[1] = notFound.length ? { codes: notFound, serial, resp: P.resp } : null;
    renderResult('#result1', 1, { total: list.length, porHoja, notas, notFound });
    toast(writes.length ? `Acción 1 aplicada: ${writes.length} celdas` : 'No hubo cambios que aplicar', writes.length ? 'success' : 'warn');
  });
}

/* =============================================================
   ACCIÓN 2 · Archivo CRM → hojas destino
   ============================================================= */
function readCrmFile(file) {
  if (!file) return;
  if (typeof XLSX === 'undefined') return toast('No cargó la librería de Excel (revisa la conexión).', 'error');
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      let found = null;
      for (const name of wb.SheetNames) {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '', raw: true, blankrows: false });
        const h = detectHeader(rows.slice(0, HEADER_SCAN_ROWS), idAliases(P.idCrm));
        if (h) { found = { sheet: name, rows, h }; break; }
      }
      if (!found) throw new Error(`El archivo no tiene la columna "${P.idCrm}".`);
      const headers = found.rows[found.h.rowOffset].map((x) => String(x).trim());
      const data = found.rows.slice(found.h.rowOffset + 1).filter((r) => codeKey(r[found.h.idCol]));
      APP.crm = { fileName: file.name, sheet: found.sheet, headers, idIdx: found.h.idCol, rows: data };
      const uniq = new Set(data.map((r) => codeKey(r[found.h.idCol]))).size;
      $('#crmName').textContent = '📄 ' + file.name;
      $('#crmInfo').classList.remove('hidden');
      $('#crmInfo').innerHTML = `Hoja <b>${esc(found.sheet)}</b> · <b>${data.length}</b> filas · <b>${uniq}</b> códigos únicos`;
      $('#result2').innerHTML = '';
      refreshColumnMatch();
    } catch (err) {
      APP.crm = null;
      $('#crmInfo').classList.remove('hidden');
      $('#crmInfo').innerHTML = `<span class="text-rose-600">❌ ${esc(err.message)}</span>`;
      $('#colsCard').classList.add('hidden');
    }
  };
  reader.readAsArrayBuffer(file);
}
function crmColumns() {
  if (!APP.crm) return [];
  const allowed = allowedKeys(2);
  const seen = new Set();
  return APP.crm.headers.map((name, i) => ({ name, i, key: normKey(name) }))
    .filter((c) => allowed.has(c.key) && !seen.has(c.key) && seen.add(c.key));
}
/** Índice de una columna del archivo CRM por nombres posibles */
function crmIndexOf(keys) { if (!APP.crm) return -1; const ks = APP.crm.headers.map(normKey); for (const k of keys) { const i = ks.indexOf(k); if (i >= 0) return i; } return -1; }
async function refreshColumnMatch() {
  if (!APP.crm) return;
  const card = $('#colsCard'), list = $('#colsList');
  card.classList.remove('hidden');
  const enHoja = new Map(); let err = '';
  for (const s of selectedSheets()) {
    try { const h = await T().headers(s); for (const k of h.cols.keys()) enHoja.set(k, (enHoja.get(k) || []).concat(s)); } catch (e) { err += e.message + ' '; }
  }
  const cols = crmColumns(), total = selectedSheets().length;
  list.innerHTML = (err ? `<div class="text-[11.5px] text-amber-700 bg-amber-50 rounded-md p-2 mb-1">⚠️ ${esc(err)}</div>` : '') +
    cols.map((c) => {
      const hojas = enHoja.get(c.key) || [];
      const exists = hojas.length > 0;
      const checked = exists && !P.excluded.includes(c.key);
      const tag = !exists ? '<span class="text-[10px] text-slate-400">no existe</span>' : `<span class="text-[10px] font-bold text-emerald-700">✓ ${hojas.length === total ? 'EN TODAS' : esc(hojas.join(', '))}</span>`;
      return `<label class="flex items-center gap-2 px-2 py-1.5 rounded-md ${!exists ? 'opacity-50' : 'hover:bg-slate-50 cursor-pointer'}">
        <input type="checkbox" class="w-4 h-4 accent-emerald-600" data-colkey="${esc(c.key)}" ${checked ? 'checked' : ''} ${!exists ? 'disabled' : ''}>
        <span class="flex-1 truncate text-[12.5px] font-medium">${esc(c.name)}</span>${tag}
      </label>`;
    }).join('');
}
/** Fusiona filas repetidas del CRM por código: los valores no vacíos posteriores tienen prioridad */
function mergeCrm(cols) {
  const merged = new Map(), order = [];
  const iRaz = crmIndexOf(['razonsocial']), iNum = crmIndexOf(['numerodecontacto', 'numerocontacto', 'numero', 'celular']);
  for (const r of APP.crm.rows) {
    const raw = String(r[APP.crm.idIdx]).trim(), key = codeKey(raw);
    if (!merged.has(key)) { merged.set(key, { raw, vals: {}, razon: '', numero: '' }); order.push(key); }
    const m = merged.get(key);
    for (const c of cols) {
      let v = r[c.i];
      if (typeof v === 'string') v = v.trim();
      if (!isEmpty(v)) m.vals[c.key] = v;
    }
    if (iRaz >= 0 && !isEmpty(r[iRaz])) m.razon = String(r[iRaz]).trim();
    if (iNum >= 0 && !isEmpty(r[iNum])) m.numero = String(r[iNum]).trim();
  }
  return { merged, order };
}

async function runAction2() {
  if (!APP.crm) return toast('Primero elige el archivo CRM.', 'warn');
  const cols = crmColumns().filter((c) => !P.excluded.includes(c.key));
  if (!cols.length) return toast('El archivo CRM no trae ninguna de las 6 columnas a rellenar.', 'warn');

  await busy('#btnRun2', async () => {
    const prog = $('#prog2');
    setProgress(prog, 0.05, 'Leyendo las hojas…');
    const { idxs, errores } = await buildIndexes(cols.map((c) => c.key));
    if (!idxs.some((idx) => cols.some((c) => idx.h.cols.has(c.key)))) throw new Error('Ninguna columna del archivo CRM coincide con los encabezados de las hojas.');
    const { merged, order } = mergeCrm(cols);
    setProgress(prog, 0.3, 'Cruzando códigos…');
    const writes = [], porHoja = {}, notas = [...errores], notFound = [];
    for (const idx of idxs) porHoja[idx.sheet] = { encontrados: 0, filas: 0, celdas: 0, conDatos: 0, repetidos: 0, faltan: [] };
    let sinDatos = 0;
    for (const key of order) {
      const m = merged.get(key);
      const donde = idxs.filter((idx) => idx.rows.has(key));
      if (!donde.length) { notFound.push(m.raw); continue; }
      const entries = Object.entries(m.vals);
      if (!entries.length) sinDatos++;
      for (const idx of donde) {
        const st = porHoja[idx.sheet]; st.encontrados++;
        const rows = idx.rows.get(key);
        if (rows.length > 1) st.repetidos++;
        for (const row of rows) {
          let tocada = false;
          for (const [k, value] of entries) {
            if (!idx.h.cols.has(k)) continue;
            const old = existingAt(idx, k, row);
            if (!isEmpty(old)) { st.conDatos++; continue; } // nunca se sobrescribe
            writes.push({ sheet: idx.sheet, row, col: absCol(idx, k), key: k, value, old });
            st.celdas++; tocada = true;
          }
          if (tocada) st.filas++;
        }
      }
    }
    if (sinDatos) notas.push(`${sinDatos} registro(s) del CRM sin datos nuevos (celdas vacías).`);
    if (writes.length) await applyWrites(writes, `Acción 2 (${writes.length} celdas)`, prog, 2);
    setProgress(prog, 1, 'Listo');
    APP.missing[2] = notFound.length ? { codes: notFound, cols } : null;
    renderResult('#result2', 2, { total: order.length, porHoja, notas, notFound });
    toast(writes.length ? `Acción 2 aplicada: ${writes.length} celdas` : 'No hubo cambios que aplicar', writes.length ? 'success' : 'warn');
  });
}

/* =============================================================
   AGREGAR CÓDIGOS QUE NO EXISTEN (al final de la hoja elegida)
   ============================================================= */
async function addMissing(action) {
  const pend = APP.missing[action];
  if (!pend || !pend.codes.length || APP.busy) return;
  const hoja = $(`#addSheet${action}`).value;
  const conRazon = action === 2 && $('#addRazon2') && $('#addRazon2').checked;
  if (!hoja) return toast('Elige la hoja donde agregarlos.', 'warn');
  if (!confirm(`¿Agregar ${pend.codes.length} código(s) al final de la hoja "${hoja}"?`)) return;
  const prog = $(`#prog${action}`);
  await busy(`#btnRun${action}`, async () => {
    setProgress(prog, 0.1, `Revisando "${hoja}"…`);
    const kDate = normKey(P.dateCol), kResp = normKey(P.respCol);
    const needed = action === 1 ? [kDate, kResp] : pend.cols.map((c) => c.key);
    // Se vuelve a leer TODAS las hojas marcadas justo antes de escribir: así nunca se duplica un código
    const { idxs } = await buildIndexes(needed);
    const idx = idxs.find((x) => x.sheet === hoja) || await T().buildIndex(needed, hoja);
    const yaExisten = [], aAgregar = [];
    for (const raw of pend.codes) {
      const k = codeKey(raw);
      const en = idxs.filter((x) => x.rows.has(k)).map((x) => x.sheet);
      if (idx.rows.has(k) || en.length) yaExisten.push(`${raw} (${(en.length ? en : [hoja]).join(', ')})`);
      else if (!aAgregar.some((x) => codeKey(x) === k)) aAgregar.push(raw);
    }
    if (!aAgregar.length) { toast('Esos códigos ya existen: no se agregó nada.', 'warn', 4000); APP.missing[action] = null; return; }
    const kId = idx.h.idCol;
    const colId = idx.colOffset + kId;
    const keyId = normKey(idx.h.cols.size ? [...idx.h.cols.entries()].find(([, v]) => v.col === kId)[0] : P.idTarget);
    const { merged } = action === 2 ? mergeCrm(pend.cols) : { merged: null };
    const kRaz = findKey(idx, NEW_ROW_RAZON), kNum = findKey(idx, NEW_ROW_NUMERO);
    let fila = idx.lastRow + 1; // primera fila sin nada escrito al final
    const writes = [], filas = [];
    for (const raw of aAgregar) {
      const w = (key, col, value, fmt) => writes.push({ sheet: hoja, row: fila, col, key, value, fmt: fmt || null, old: '', nueva: true });
      w(keyId, colId, codeCellValue(raw));
      if (action === 1) {
        if (idx.h.cols.has(kDate)) w(kDate, absCol(idx, kDate), pend.serial, P.dateFmt);
        if (idx.h.cols.has(kResp)) w(kResp, absCol(idx, kResp), pend.resp);
      } else {
        const m = merged.get(codeKey(raw));
        if (m) {
          for (const [k, v] of Object.entries(m.vals)) if (idx.h.cols.has(k)) w(k, absCol(idx, k), v);
          if (conRazon && kRaz && m.razon) w(kRaz, absCol(idx, kRaz), m.razon);
          if (conRazon && kNum && m.numero) w(kNum, absCol(idx, kNum), m.numero);
        }
      }
      filas.push(fila + 1); fila++;
    }
    await applyWrites(writes, `Agregar ${aAgregar.length} en ${hoja}`, prog, action);
    setProgress(prog, 1, 'Listo');
    APP.missing[action] = null;
    const box = $(`#addBox${action}`);
    if (box) box.innerHTML = `<div class="text-[12px] text-emerald-700 font-semibold">✓ ${aAgregar.length} código(s) agregados en "${esc(hoja)}" (filas ${filas[0]}–${filas[filas.length - 1]}).</div>` +
      (yaExisten.length ? `<div class="text-[11.5px] text-slate-500 mt-1">No se duplicaron ${yaExisten.length}: ${esc(yaExisten.join(', '))}</div>` : '');
    toast(`${aAgregar.length} código(s) agregados en ${hoja}`, 'success');
  });
}

/* =============================================================
   UI COMPARTIDA
   ============================================================= */
function setProgress(el, p, label) {
  el.classList.remove('hidden');
  const pct = Math.round(Math.max(0, Math.min(1, p)) * 100);
  el.innerHTML = `<div class="flex justify-between text-[11px] text-slate-500 mb-1"><span>${esc(label)}</span><span class="tabular-nums">${pct}%</span></div>
    <div class="h-2 bg-slate-200 rounded-full overflow-hidden"><div class="h-full ${pct === 100 ? 'bg-emerald-500' : 'bg-indigo-500'} transition-all duration-200" style="width:${pct}%"></div></div>`;
}
function renderResult(sel, action, { total, porHoja, notas, notFound }) {
  const hojas = Object.entries(porHoja);
  const celdas = hojas.reduce((s, [, x]) => s + x.celdas, 0);
  const encontrados = total - notFound.length;
  const nf = notFound.length;
  const opciones = selectedSheets().map((s) => `<option>${esc(s)}</option>`).join('');
  $(sel).innerHTML = `<div class="card p-3 space-y-2">
    <div class="grid grid-cols-3 gap-1.5 text-center">
      <div class="rounded-lg bg-emerald-50 py-2"><div class="text-[16px] font-bold text-emerald-700 tabular-nums">${encontrados}/${total}</div><div class="text-[10.5px] text-emerald-800/80">Encontrados</div></div>
      <div class="rounded-lg bg-emerald-50 py-2"><div class="text-[16px] font-bold text-emerald-700 tabular-nums">${celdas}</div><div class="text-[10.5px] text-emerald-800/80">Celdas escritas</div></div>
      <div class="rounded-lg ${nf ? 'bg-amber-50' : 'bg-emerald-50'} py-2"><div class="text-[16px] font-bold ${nf ? 'text-amber-700' : 'text-emerald-700'} tabular-nums">${nf}</div><div class="text-[10.5px] ${nf ? 'text-amber-800/80' : 'text-emerald-800/80'}">No están</div></div>
    </div>
    <div class="space-y-1">
      ${hojas.map(([h, x]) => `<div class="flex items-center gap-2 text-[11.5px] rounded-md bg-slate-50 px-2 py-1.5">
        <span class="font-semibold truncate flex-1">📄 ${esc(h)}</span>
        <span class="tabular-nums text-slate-600">${x.encontrados} cód. · ${x.filas} filas · ${x.celdas} celdas</span>
      </div>${x.conDatos || x.repetidos ? `<div class="text-[11px] text-slate-500 pl-2">${x.conDatos ? `${x.conDatos} celda(s) ya tenían datos (intactas)` : ''}${x.conDatos && x.repetidos ? ' · ' : ''}${x.repetidos ? `⚠️ ${x.repetidos} código(s) repetidos en esta hoja` : ''}</div>` : ''}`).join('')}
    </div>
    ${notas.filter(Boolean).map((n) => `<div class="text-[11.5px] text-amber-700">• ${esc(n)}</div>`).join('')}
    ${nf ? `<div class="rounded-lg border border-amber-200 bg-amber-50 p-2 space-y-2">
      <div class="flex items-center justify-between"><span class="text-[12px] font-semibold text-amber-800">⚠️ ${nf} código(s) no están en ninguna hoja</span>
      <button class="text-[11px] font-semibold text-amber-800 underline" data-copy-nf="${esc(notFound.join('\n'))}">Copiar</button></div>
      <div class="max-h-24 overflow-y-auto scroll-thin font-mono text-[11px] text-amber-900 break-all">${notFound.map(esc).join(', ')}</div>
      <div id="addBox${action}" class="border-t border-amber-200 pt-2 space-y-1.5">
        <div class="flex items-center gap-1.5">
          <span class="text-[11.5px] text-amber-900 shrink-0">Agregar en</span>
          <select id="addSheet${action}" class="inp !py-1 !text-[12px]">${opciones}</select>
        </div>
        ${action === 2 ? `<label class="flex items-center gap-1.5 text-[11.5px] text-amber-900"><input id="addRazon2" type="checkbox" class="accent-indigo-600" ${P.addRazon ? 'checked' : ''}> Incluir Razón Social y Número del CRM (solo filas nuevas)</label>` : ''}
        <button class="btn btn-primary w-full !py-1.5 !text-[12px]" data-add-missing="${action}">➕ Rellenar y agregar ${nf} al final</button>
        <div class="text-[10.5px] text-amber-800/80">Se escriben en la primera fila vacía del final. Antes se vuelve a revisar para no duplicar.</div>
      </div></div>`
    : '<div class="text-[11.5px] text-emerald-700 font-semibold">✓ Todos los códigos están en las hojas.</div>'}
  </div>`;
}
async function busy(btnSel, fn) {
  if (APP.busy) return;
  APP.busy = true;
  const btn = $(btnSel), html = btn.innerHTML;
  btn.disabled = true; btn.innerHTML = '⏳ Procesando…';
  try { await fn(); } catch (err) {
    console.error(err);
    toast(err.message || String(err), 'error', 4500);
    $(btnSel.replace('btnRun', 'prog')).classList.add('hidden');
  } finally { APP.busy = false; btn.disabled = false; btn.innerHTML = html; }
}
function renderFooter() {
  const u = $('#btnUndo');
  u.disabled = !APP.undo;
  u.textContent = APP.undo ? `↶ Deshacer ${APP.undo.label}` : '↶ Deshacer última acción';
  const d = $('#btnDownload');
  d.classList.toggle('hidden', APP.mode !== 'local' || !APP.master);
  if (APP.master) { d.disabled = !APP.master.changes; d.textContent = APP.master.changes ? `⬇ Descargar Excel (${APP.master.changes})` : '⬇ Descargar Excel'; }
}
async function undoLast() {
  if (!APP.undo || APP.busy) return;
  if (!confirm(`¿Deshacer ${APP.undo.label}? Se restaurarán los valores anteriores.`)) return;
  APP.busy = true;
  try {
    const porHoja = new Map();
    for (const w of APP.undo.writes) { if (!porHoja.has(w.sheet)) porHoja.set(w.sheet, []); porHoja.get(w.sheet).push(w); }
    for (const [hoja, lista] of porHoja) await T().write(lista, hoja);
    toast('Cambios deshechos', 'success');
    APP.undo = null; renderFooter();
  } catch (e) { toast(e.message, 'error', 4000); } finally { APP.busy = false; }
}
function setTab(t) {
  P.tab = t; persist();
  $$('.tab').forEach((b) => b.classList.toggle('on', b.dataset.tab === t));
  $('#tab-a1').classList.toggle('hidden', t !== 'a1');
  $('#tab-a2').classList.toggle('hidden', t !== 'a2');
}

/* ---------------- Hojas disponibles ---------------- */
async function loadSheetOptions() {
  let names = [];
  try {
    if (APP.mode === 'office') ({ names } = await OfficeTarget.sheetNames());
    else if (APP.master) names = APP.master.wb.SheetNames;
  } catch (e) { console.error(e); }
  APP.sheetNames = names;
  if (names.length) {
    // Ajusta nombres guardados (mayúsculas/espacios) y, si no hay ninguna marcada, propone BD-BBPE + Seguimiento
    let sel = P.sheets.map((s) => names.find((n) => normKey(n) === normKey(s))).filter(Boolean);
    if (!sel.length && P.sheet) { const m = names.find((n) => normKey(n) === normKey(P.sheet)); if (m) sel = [m]; }
    const sugeridas = names.filter((n) => DEFAULT_SHEET_PATTERNS.some((re) => re.test(normKey(n))));
    if (!P.sheetsTouched) sel = [...new Set([...sel, ...sugeridas])];
    if (!sel.length) sel = [names[0]];
    P.sheets = sel; persist();
  } else if (!P.sheets.length) P.sheets = [P.sheet || DEFAULTS.sheet];
  renderSheetPicker();
  renderMode();
}
function renderSheetPicker() {
  const names = APP.sheetNames.length ? APP.sheetNames : P.sheets;
  // Chips: marcadas en verde (se usan), las demás en gris. Un clic marca o desmarca.
  $('#cfgSheets').innerHTML = names.map((n) => { const on = P.sheets.includes(n);
    return `<label class="flex items-center gap-1.5 px-2.5 py-1.5 rounded-full border text-[12px] font-semibold cursor-pointer select-none ${on ? 'bg-emerald-600 border-emerald-600 text-white' : 'bg-white border-slate-200 text-slate-500 hover:bg-slate-50'}">
    <input type="checkbox" class="hidden" data-sheet="${esc(n)}" ${on ? 'checked' : ''}>${on ? '✓ ' : ''}${esc(n)}</label>`; }).join('');
}
function renderMode() {
  const b = $('#modeBadge'), hojas = selectedSheets();
  const txt = hojas.length ? hojas.map(esc).join(' + ') : 'ninguna hoja';
  if (APP.mode === 'office') b.innerHTML = `<span class="text-emerald-600 font-semibold">● Excel Online</span> · <b>${txt}</b>`;
  else if (APP.mode === 'local') b.innerHTML = `<span class="text-indigo-600 font-semibold">● Modo local</span>${APP.master ? ` · ${txt}` : ' · carga tu Excel maestro'}`;
}

/* ---------------- Excel maestro local ---------------- */
function readMasterFile(file) {
  if (!file) return;
  if (typeof XLSX === 'undefined') return toast('No cargó la librería de Excel (revisa la conexión).', 'error');
  const reader = new FileReader();
  reader.onload = async (e) => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellStyles: true, cellNF: true, cellDates: false });
      APP.master = { wb, fileName: file.name, changes: 0 };
      APP.undo = null;
      $('#masterName').textContent = '📗 ' + file.name;
      await loadSheetOptions();
      const partes = [];
      for (const s of selectedSheets()) {
        try { const h = await LocalTarget.headers(s); partes.push(`<b>${esc(s)}</b>: ${h.cols.size} columnas`); }
        catch (err) { partes.push(`<span class="text-rose-600">${esc(err.message)}</span>`); }
      }
      $('#masterInfo').classList.remove('hidden');
      $('#masterInfo').innerHTML = partes.join('<br>') + '<br><span class="text-[10.5px] text-slate-400">La descarga conserva datos y formatos de número; colores y estilos pueden perderse.</span>';
      renderFooter(); refreshColumnMatch();
      toast('Excel maestro cargado', 'success');
    } catch (err) { toast('No se pudo leer el archivo: ' + err.message, 'error', 4000); }
  };
  reader.readAsArrayBuffer(file);
}
function downloadMaster() {
  if (!APP.master) return;
  const base = APP.master.fileName.replace(/\.(xlsx|xlsm|xls)$/i, '');
  XLSX.writeFile(APP.master.wb, `${base}_actualizado.xlsx`, { bookType: 'xlsx', compression: true, cellStyles: true });
  toast('Archivo actualizado descargado', 'success');
}

/* =============================================================
   ARRANQUE
   ============================================================= */
function bindEvents() {
  $$('.tab').forEach((b) => b.addEventListener('click', () => setTab(b.dataset.tab)));
  $('#btnSettings').addEventListener('click', () => $('#settings').classList.toggle('hidden'));
  $('#btnReloadSheets').addEventListener('click', async () => { await loadSheetOptions(); refreshColumnMatch(); toast('Hojas actualizadas'); });
  $('#btnReloadSheets2').addEventListener('click', async () => { await loadSheetOptions(); refreshColumnMatch(); toast('Hojas actualizadas'); });
  $('#cfgSheets').addEventListener('change', (e) => {
    const cb = e.target.closest('[data-sheet]'); if (!cb) return;
    const n = cb.dataset.sheet;
    P.sheets = cb.checked ? [...new Set([...P.sheets, n])] : P.sheets.filter((x) => x !== n);
    P.sheetsTouched = true; persist(); renderSheetPicker(); renderMode(); refreshColumnMatch();
  });
  for (const [id, key] of [['#cfgIdTarget', 'idTarget'], ['#cfgIdCrm', 'idCrm'], ['#cfgDateCol', 'dateCol'], ['#cfgRespCol', 'respCol'], ['#cfgDateFmt', 'dateFmt']]) {
    $(id).addEventListener('change', (e) => { P[key] = e.target.value.trim() || DEFAULTS[key]; e.target.value = P[key]; persist(); renderDatePreview(); refreshColumnMatch(); });
  }
  // Acción 1
  $('#codes').addEventListener('input', () => { P.codes = $('#codes').value; persist(); renderCodesCounter(); });
  $('#codes').addEventListener('blur', () => {
    const r = parseCodes($('#codes').value);
    if (r.dup || r.empty) { $('#codes').value = r.list.map((x) => x.raw).join('\n'); P.codes = $('#codes').value; persist(); renderCodesCounter(); }
  });
  $('#btnClearCodes').addEventListener('click', () => { $('#codes').value = ''; P.codes = ''; persist(); renderCodesCounter(); $('#result1').innerHTML = ''; $('#prog1').classList.add('hidden'); APP.missing[1] = null; });
  $('#date').addEventListener('input', renderDatePreview);
  $('#btnToday').addEventListener('click', () => { $('#date').value = todayISO(); renderDatePreview(); });
  $('#respChips').addEventListener('click', (e) => { const b = e.target.closest('[data-resp]'); if (!b) return; P.resp = b.dataset.resp; persist(); renderRespChips(); });
  $('#btnRun1').addEventListener('click', runAction1);
  // Acción 2
  $('#crmFile').addEventListener('change', (e) => { readCrmFile(e.target.files[0]); e.target.value = ''; });
  $('#btnRefreshCols').addEventListener('click', refreshColumnMatch);
  $('#colsList').addEventListener('change', (e) => {
    const cb = e.target.closest('[data-colkey]'); if (!cb) return;
    const k = cb.dataset.colkey;
    P.excluded = cb.checked ? P.excluded.filter((x) => x !== k) : [...new Set([...P.excluded, k])];
    persist();
  });
  $('#btnRun2').addEventListener('click', runAction2);
  // Agregar los que no existen
  document.addEventListener('click', (e) => { const b = e.target.closest('[data-add-missing]'); if (b) addMissing(Number(b.dataset.addMissing)); });
  document.addEventListener('change', (e) => { if (e.target.id === 'addRazon2') { P.addRazon = e.target.checked; persist(); } });
  // Local
  $('#masterFile').addEventListener('change', (e) => { readMasterFile(e.target.files[0]); e.target.value = ''; });
  $('#btnDownload').addEventListener('click', downloadMaster);
  $('#btnUndo').addEventListener('click', undoLast);
  // Copiar códigos no encontrados
  document.addEventListener('click', async (e) => { const b = e.target.closest('[data-copy-nf]'); if (b && await copyText(b.dataset.copyNf)) toast('Códigos copiados'); });
  // Arrastrar y soltar
  for (const [zone, handler] of [['#masterDrop', readMasterFile], ['#crmDrop', readCrmFile]]) {
    const z = $(zone);
    ['dragenter', 'dragover'].forEach((ev) => z.addEventListener(ev, (e) => { e.preventDefault(); z.classList.add('border-indigo-400', 'bg-indigo-50'); }));
    ['dragleave', 'drop'].forEach((ev) => z.addEventListener(ev, (e) => { e.preventDefault(); z.classList.remove('border-indigo-400', 'bg-indigo-50'); }));
    z.addEventListener('drop', (e) => handler(e.dataTransfer.files[0]));
  }
  window.addEventListener('beforeunload', (e) => { if (APP.mode === 'local' && APP.master?.changes) { e.preventDefault(); e.returnValue = ''; } });
}

function initUI() {
  $('#cfgIdTarget').value = P.idTarget; $('#cfgIdCrm').value = P.idCrm;
  $('#cfgDateCol').value = P.dateCol; $('#cfgRespCol').value = P.respCol; $('#cfgDateFmt').value = P.dateFmt;
  $('#codes').value = P.codes || '';
  $('#date').value = todayISO();
  renderCodesCounter(); renderRespChips(); renderDatePreview(); setTab(P.tab || 'a1'); renderFooter();
}

async function start(mode) {
  // Solo se permite pasar de "detectando" a un modo, o de "local" a "office" si Excel respondió tarde
  if (APP.mode === mode || (APP.mode === 'office' && mode === 'local')) return;
  APP.mode = mode;
  $('#localCard').classList.toggle('hidden', mode !== 'local');
  await loadSheetOptions();
  renderFooter();
}

initUI();
bindEvents();
// Modo dual: dentro de Excel → Office.js; en el navegador → archivo local
if (window.Office && typeof Office.onReady === 'function') {
  Office.onReady((info) => start(info && info.host === Office.HostType.Excel ? 'office' : 'local'));
  setTimeout(() => start('local'), 4000); // por si Office.js nunca responde
} else {
  start('local');
}
