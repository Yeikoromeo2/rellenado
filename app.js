/* =============================================================
   Rellenado CRM · Complemento de Excel + modo web local
   Todo ocurre en el navegador: sin servidores ni bases de datos.
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
// Nombres alternativos que también se aceptan como columna identificadora
const ID_ALIASES = ['customerid', 'codigodecliente', 'codigocliente', 'codcliente'];
const HEADER_SCAN_ROWS = 10;
const LS_KEY = 'rellenado_crm_v1';
// Acción 2: ÚNICAS columnas que se rellenan desde el archivo CRM
const A2_COLUMNS = ['ESTADO CONTACTO', 'PRÓXIMA ACCIÓN', 'OBSERVACIÓN', 'NOMBRE DEL CONTACTO', 'CORREO', 'TELÉFONO'];
// Columnas que JAMÁS se escriben (solo se usan para buscar o se dejan intactas)
const PROTECTED_KEYS = ['customerid', 'codigodecliente', 'codigocliente', 'codcliente', 'razonsocial', 'numerodecontacto', 'numerocontacto', 'numero', 'nrodecontacto', 'celular'];
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
function sameValue(a, b) {
  if (isEmpty(a) && isEmpty(b)) return true;
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) < 1e-9;
  return String(a ?? '').trim() === String(b ?? '').trim();
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

/* ---------------- Estado persistente (solo preferencias) ---------------- */
let P = (() => { try { return { ...DEFAULTS, resp: '', codes: '', tab: 'a1', excluded: DEFAULT_EXCLUDED.slice(), ...JSON.parse(localStorage.getItem(LS_KEY) || '{}') }; } catch (e) { return { ...DEFAULTS, resp: '', codes: '', tab: 'a1', excluded: DEFAULT_EXCLUDED.slice() }; } })();
function persist() { try { localStorage.setItem(LS_KEY, JSON.stringify(P)); } catch (e) {} }

/* ---------------- Estado de sesión ---------------- */
const APP = {
  mode: 'detecting',   // 'office' | 'local'
  master: null,        // modo local: { wb, fileName, changes }
  crm: null,           // { fileName, headers[], idIdx, rows[][] }
  undo: null,          // { label, writes[] }
  busy: false,
};

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
  async getSheet(ctx) {
    const ws = ctx.workbook.worksheets.getItemOrNullObject(P.sheet);
    await ctx.sync();
    if (ws.isNullObject) throw new Error(`No existe la hoja "${P.sheet}" en este libro. Revisa ⚙ Configuración.`);
    return ws;
  },
  async scanHeader(ctx, ws) {
    const used = ws.getUsedRange(true);
    used.load('rowIndex,columnIndex,rowCount,columnCount');
    await ctx.sync();
    const scan = ws.getRangeByIndexes(used.rowIndex, used.columnIndex, Math.min(HEADER_SCAN_ROWS, used.rowCount), used.columnCount);
    scan.load('values');
    await ctx.sync();
    const h = detectHeader(scan.values, idAliases(P.idTarget));
    if (!h) throw new Error(`No se encontró la columna "${P.idTarget}" en las ${HEADER_SCAN_ROWS} primeras filas de "${P.sheet}".`);
    return { used, h };
  },
  async headers() {
    return Excel.run(async (ctx) => {
      const ws = await this.getSheet(ctx);
      const { h } = await this.scanHeader(ctx, ws);
      return h;
    });
  },
  /** Carga la columna ID y las columnas pedidas (fórmulas = valor original exacto). */
  async buildIndex(neededKeys) {
    return Excel.run(async (ctx) => {
      const ws = await this.getSheet(ctx);
      const { used, h } = await this.scanHeader(ctx, ws);
      const headerRow = used.rowIndex + h.rowOffset;
      const first = headerRow + 1;
      const last = used.rowIndex + used.rowCount - 1;
      const n = last - first + 1;
      const idx = newIndex(h, used.columnIndex, first);
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
  async write(writes, onProgress) {
    const runs = groupRuns(writes);
    await Excel.run(async (ctx) => {
      const ws = await this.getSheet(ctx);
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
  ws() {
    if (!APP.master) throw new Error('Primero carga el Excel maestro (arriba).');
    const ws = APP.master.wb.Sheets[P.sheet];
    if (!ws) throw new Error(`El archivo no tiene la hoja "${P.sheet}". Revisa ⚙ Configuración.`);
    return ws;
  },
  scan(ws) {
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const rows = [];
    for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + HEADER_SCAN_ROWS - 1); r++) {
      const row = [];
      for (let c = range.s.c; c <= range.e.c; c++) { const cell = ws[XLSX.utils.encode_cell({ r, c })]; row.push(cell ? (cell.w ?? cell.v ?? '') : ''); }
      rows.push(row);
    }
    const h = detectHeader(rows, idAliases(P.idTarget));
    if (!h) throw new Error(`No se encontró la columna "${P.idTarget}" en las ${HEADER_SCAN_ROWS} primeras filas de "${P.sheet}".`);
    return { range, h };
  },
  async headers() { return this.scan(this.ws()).h; },
  async buildIndex(neededKeys) {
    const ws = this.ws();
    const { range, h } = this.scan(ws);
    const first = range.s.r + h.rowOffset + 1;
    const idx = newIndex(h, range.s.c, first);
    const cols = neededKeys.map((k) => [k, h.cols.get(k)]).filter(([, c]) => c);
    for (const [k] of cols) idx.existing.set(k, []);
    for (let r = first; r <= range.e.r; r++) {
      const idCell = ws[XLSX.utils.encode_cell({ r, c: range.s.c + h.idCol })];
      addRow(idx, idCell ? idCell.v : '', r);
      for (const [k, c] of cols) { const cell = ws[XLSX.utils.encode_cell({ r, c: range.s.c + c.col })]; idx.existing.get(k).push(cell ? cell.v : ''); }
    }
    return idx;
  },
  async write(writes, onProgress) {
    const ws = this.ws();
    writes.forEach((w, i) => {
      const addr = XLSX.utils.encode_cell({ r: w.row, c: w.col });
      if (w.restore) { if (w.value === undefined) delete ws[addr]; else ws[addr] = JSON.parse(JSON.stringify(w.value)); }
      else {
        const v = w.value;
        if (typeof v === 'number') ws[addr] = { t: 'n', v, ...(w.fmt ? { z: w.fmt } : (ws[addr]?.z ? { z: ws[addr].z } : {})) };
        else if (typeof v === 'boolean') ws[addr] = { t: 'b', v };
        else ws[addr] = { t: 's', v: String(v) };
      }
      if (i % 500 === 0 && onProgress) onProgress(i / writes.length);
    });
    APP.master.changes += writes.length;
    onProgress && onProgress(1);
    renderFooter();
  },
  snapshot(row, col) {
    const cell = this.ws()[XLSX.utils.encode_cell({ r: row, c: col })];
    return cell ? JSON.parse(JSON.stringify(cell)) : undefined;
  },
};

/* ---------------- Índice común ---------------- */
function newIndex(h, colOffset, firstRow) {
  return { h, colOffset, firstRow, rows: new Map(), existing: new Map() };
}
function addRow(idx, rawId, absRow) {
  const k = codeKey(rawId); if (!k) return;
  if (!idx.rows.has(k)) idx.rows.set(k, []);
  idx.rows.get(k).push(absRow);
}
const absCol = (idx, key) => idx.colOffset + idx.h.cols.get(key).col;
const existingAt = (idx, key, absRow) => (idx.existing.get(key) || [])[absRow - idx.firstRow];
const T = () => (APP.mode === 'office' ? OfficeTarget : LocalTarget);

/** Columnas en las que cada acción tiene permiso de escribir */
function allowedKeys(action) {
  const keys = action === 1 ? [P.dateCol, P.respCol] : A2_COLUMNS;
  const blocked = new Set([...PROTECTED_KEYS, ...idAliases(P.idTarget), ...idAliases(P.idCrm)]);
  return new Set(keys.map(normKey).filter((k) => !blocked.has(k)));
}
/** Aplica escrituras y prepara el "deshacer".
 *  Candado final: solo columnas permitidas y solo celdas que estaban vacías. */
async function applyWrites(writes, label, progEl, action) {
  const allowed = allowedKeys(action);
  for (const w of writes) {
    if (!allowed.has(w.key)) throw new Error(`Bloqueado: la columna "${w.key}" no está permitida. No se escribió nada.`);
    if (!isEmpty(w.old)) throw new Error('Bloqueado: se intentó escribir sobre una celda con datos. No se escribió nada.');
  }
  const undoWrites = writes.map((w) => ({
    row: w.row, col: w.col, restore: true,
    value: APP.mode === 'office' ? (isEmpty(w.old) ? '' : w.old) : LocalTarget.snapshot(w.row, w.col),
    fmt: null,
  }));
  await T().write(writes, (p) => setProgress(progEl, p, 'Escribiendo en Excel…'));
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

  await busy('#btnRun1', async () => {
    const prog = $('#prog1');
    setProgress(prog, 0.05, 'Leyendo la hoja…');
    const idx = await T().buildIndex([kDate, kResp]);
    const allowed = allowedKeys(1);
    if (!allowed.has(kDate) || !allowed.has(kResp)) throw new Error('La columna de fecha o de responsable apunta a una columna protegida. Revisa ⚙ Configuración.');
    const missing = [[kDate, P.dateCol], [kResp, P.respCol]].filter(([k]) => !idx.h.cols.has(k)).map(([, n]) => n);
    if (missing.length) throw new Error(`Faltan columnas en "${P.sheet}": ${missing.join(', ')}.`);

    const serial = excelSerial(date);
    const writes = [], notFound = [];
    let rowsFilled = 0, codesFound = 0, protectedCells = 0;
    for (const { key, raw } of list) {
      const rows = idx.rows.get(key);
      if (!rows) { notFound.push(raw); continue; }
      codesFound++;
      for (const row of rows) {
        let touched = false;
        for (const [k, value, fmt] of [[kDate, serial, P.dateFmt], [kResp, P.resp, null]]) {
          const old = existingAt(idx, k, row);
          if (!isEmpty(old)) { protectedCells++; continue; } // nunca se sobrescribe
          writes.push({ row, col: absCol(idx, k), key: k, value, fmt, old });
          touched = true;
        }
        if (touched) rowsFilled++;
      }
    }
    if (writes.length) await applyWrites(writes, `Acción 1 (${writes.length} celdas)`, prog, 1);
    setProgress(prog, 1, 'Listo');
    renderResult('#result1', {
      ok: [
        ['Filas rellenadas', rowsFilled],
        ['Códigos encontrados', `${codesFound} de ${list.length}`],
        ['Celdas escritas', writes.length],
      ],
      notes: [
        protectedCells ? `${protectedCells} celda(s) ya tenían datos y se dejaron intactas.` : '',
      ],
      notFound,
    });
    toast(writes.length ? `Acción 1 aplicada: ${rowsFilled} filas` : 'No hubo cambios que aplicar', writes.length ? 'success' : 'warn');
  });
}

/* =============================================================
   ACCIÓN 2 · Archivo CRM → hoja destino
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
async function refreshColumnMatch() {
  if (!APP.crm) return;
  const card = $('#colsCard'), list = $('#colsList');
  card.classList.remove('hidden');
  let targetCols = null, err = '';
  try { targetCols = (await T().headers()).cols; } catch (e) { err = e.message; }
  const cols = crmColumns();
  list.innerHTML = (err ? `<div class="text-[11.5px] text-amber-700 bg-amber-50 rounded-md p-2 mb-1">⚠️ ${esc(err)}</div>` : '') +
    cols.map((c) => {
      const exists = targetCols ? targetCols.has(c.key) : null;
      const checked = exists !== false && !P.excluded.includes(c.key);
      return `<label class="flex items-center gap-2 px-2 py-1.5 rounded-md ${exists === false ? 'opacity-50' : 'hover:bg-slate-50 cursor-pointer'}">
        <input type="checkbox" class="w-4 h-4 accent-emerald-600" data-colkey="${esc(c.key)}" ${checked ? 'checked' : ''} ${exists === false ? 'disabled' : ''}>
        <span class="flex-1 truncate text-[12.5px] font-medium">${esc(c.name)}</span>
        ${exists === true ? '<span class="text-[10px] font-bold text-emerald-700">✓ EN HOJA</span>' : exists === false ? '<span class="text-[10px] text-slate-400">no existe</span>' : ''}
      </label>`;
    }).join('');
}

async function runAction2() {
  if (!APP.crm) return toast('Primero elige el archivo CRM.', 'warn');
  const cols = crmColumns().filter((c) => !P.excluded.includes(c.key));
  if (!cols.length) return toast('El archivo CRM no trae ninguna de las 6 columnas a rellenar.', 'warn');

  await busy('#btnRun2', async () => {
    const prog = $('#prog2');
    setProgress(prog, 0.05, 'Leyendo la hoja…');
    const idx = await T().buildIndex(cols.map((c) => c.key));
    const usable = cols.filter((c) => idx.h.cols.has(c.key));
    if (!usable.length) throw new Error('Ninguna columna del archivo CRM coincide con los encabezados de la hoja.');

    // Fusiona filas repetidas por código: los valores no vacíos posteriores tienen prioridad
    const merged = new Map(), order = [];
    for (const r of APP.crm.rows) {
      const raw = String(r[APP.crm.idIdx]).trim(), key = codeKey(raw);
      if (!merged.has(key)) { merged.set(key, { raw, vals: {} }); order.push(key); }
      const m = merged.get(key);
      for (const c of usable) {
        let v = r[c.i];
        if (typeof v === 'string') v = v.trim();
        if (!isEmpty(v)) m.vals[c.key] = v;
      }
    }
    setProgress(prog, 0.3, 'Cruzando códigos…');
    const writes = [], notFound = [];
    let updated = 0, unchangedCodes = 0, emptyCodes = 0, protectedCells = 0;
    for (const key of order) {
      const m = merged.get(key);
      const rows = idx.rows.get(key);
      if (!rows) { notFound.push(m.raw); continue; }
      const entries = Object.entries(m.vals);
      if (!entries.length) { emptyCodes++; continue; }
      let any = false;
      for (const row of rows) {
        for (const [k, value] of entries) {
          const old = existingAt(idx, k, row);
          if (!isEmpty(old)) { protectedCells++; continue; } // nunca se sobrescribe
          writes.push({ row, col: absCol(idx, k), key: k, value, old });
          any = true;
        }
      }
      if (any) updated++; else unchangedCodes++;
    }
    if (writes.length) await applyWrites(writes, `Acción 2 (${writes.length} celdas)`, prog, 2);
    setProgress(prog, 1, 'Listo');
    renderResult('#result2', {
      ok: [
        ['Registros actualizados', updated],
        ['Celdas escritas', writes.length],
        ['Columnas cruzadas', usable.length],
      ],
      notes: [
        unchangedCodes ? `${unchangedCodes} registro(s) sin celdas vacías por rellenar.` : '',
        protectedCells ? `${protectedCells} celda(s) ya tenían datos y se dejaron intactas.` : '',
        emptyCodes ? `${emptyCodes} registro(s) sin datos nuevos (celdas vacías, se conservó lo existente).` : '',
      ],
      notFound,
    });
    toast(writes.length ? `Acción 2 aplicada: ${updated} registros` : 'No hubo cambios que aplicar', writes.length ? 'success' : 'warn');
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
function renderResult(sel, { ok, notes, notFound }) {
  const nf = notFound.length;
  $(sel).innerHTML = `<div class="card p-3 space-y-2">
    <div class="grid grid-cols-3 gap-1.5 text-center">
      ${ok.map(([l, v]) => `<div class="rounded-lg bg-emerald-50 py-2"><div class="text-[16px] font-bold text-emerald-700 tabular-nums">${esc(v)}</div><div class="text-[10.5px] text-emerald-800/80 leading-tight">${esc(l)}</div></div>`).join('')}
    </div>
    ${notes.filter(Boolean).map((n) => `<div class="text-[11.5px] text-slate-500">• ${esc(n)}</div>`).join('')}
    ${nf ? `<div class="rounded-lg border border-amber-200 bg-amber-50 p-2">
      <div class="flex items-center justify-between"><span class="text-[12px] font-semibold text-amber-800">⚠️ ${nf} código(s) no existen en la hoja</span>
      <button class="text-[11px] font-semibold text-amber-800 underline" data-copy-nf="${esc(notFound.join('\n'))}">Copiar</button></div>
      <div class="mt-1 max-h-24 overflow-y-auto scroll-thin font-mono text-[11px] text-amber-900 break-all">${notFound.map(esc).join(', ')}</div></div>`
    : '<div class="text-[11.5px] text-emerald-700 font-semibold">✓ Todos los códigos se encontraron en la hoja.</div>'}
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
    await T().write(APP.undo.writes);
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
  const sel = $('#cfgSheet');
  let names = [], active = '';
  try {
    if (APP.mode === 'office') ({ names, active } = await OfficeTarget.sheetNames());
    else if (APP.master) names = APP.master.wb.SheetNames;
  } catch (e) { console.error(e); }
  if (names.length && !names.includes(P.sheet)) {
    const match = names.find((n) => normKey(n) === normKey(P.sheet));
    P.sheet = match || (names.includes(DEFAULTS.sheet) ? DEFAULTS.sheet : (active || names[0]));
    persist();
  }
  const opts = names.length ? names : [P.sheet];
  sel.innerHTML = opts.map((n) => `<option ${n === P.sheet ? 'selected' : ''}>${esc(n)}</option>`).join('');
  renderMode();
}
function renderMode() {
  const b = $('#modeBadge');
  if (APP.mode === 'office') b.innerHTML = `<span class="text-emerald-600 font-semibold">● Excel Online</span> · hoja <b>${esc(P.sheet)}</b>`;
  else if (APP.mode === 'local') b.innerHTML = `<span class="text-indigo-600 font-semibold">● Modo local</span>${APP.master ? ` · ${esc(P.sheet)}` : ' · carga tu Excel maestro'}`;
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
      let info = `${wb.SheetNames.length} hoja(s) · usando <b>${esc(P.sheet)}</b>`;
      try { const h = await LocalTarget.headers(); info += ` · ${h.cols.size} columnas detectadas`; } catch (err) { info += `<br><span class="text-rose-600">${esc(err.message)}</span>`; }
      $('#masterInfo').classList.remove('hidden');
      $('#masterInfo').innerHTML = info + '<br><span class="text-[10.5px] text-slate-400">La descarga conserva datos y formatos de número; colores, estilos y formatos condicionales pueden perderse (limitación de SheetJS). Para el libro compartido usa el modo complemento.</span>';
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
  $('#cfgSheet').addEventListener('change', (e) => { P.sheet = e.target.value; persist(); renderMode(); refreshColumnMatch(); });
  for (const [id, key] of [['#cfgIdTarget', 'idTarget'], ['#cfgIdCrm', 'idCrm'], ['#cfgDateCol', 'dateCol'], ['#cfgRespCol', 'respCol'], ['#cfgDateFmt', 'dateFmt']]) {
    $(id).addEventListener('change', (e) => { P[key] = e.target.value.trim() || DEFAULTS[key]; e.target.value = P[key]; persist(); renderDatePreview(); refreshColumnMatch(); });
  }
  // Acción 1
  $('#codes').addEventListener('input', () => { P.codes = $('#codes').value; persist(); renderCodesCounter(); });
  $('#codes').addEventListener('blur', () => {
    const r = parseCodes($('#codes').value);
    if (r.dup || r.empty) { $('#codes').value = r.list.map((x) => x.raw).join('\n'); P.codes = $('#codes').value; persist(); renderCodesCounter(); }
  });
  $('#btnClearCodes').addEventListener('click', () => { $('#codes').value = ''; P.codes = ''; persist(); renderCodesCounter(); $('#result1').innerHTML = ''; $('#prog1').classList.add('hidden'); });
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
