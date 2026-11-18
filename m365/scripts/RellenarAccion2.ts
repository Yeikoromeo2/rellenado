/**
 * Rellenado CRM · Acción 2 — Office Script
 * ------------------------------------------------------------------
 * Traslada la gestión del CRM (lista CRM_Gestiones de SharePoint o el Excel
 * descargado de la herramienta de difusión) a la hoja BD-BBPE.
 * Mismas reglas que el complemento "Rellenado CRM":
 *  - Cruza "Código de Cliente" del CRM con "Customer ID" de la hoja.
 *  - Solo escribe en: ESTADO CONTACTO, PRÓXIMA ACCIÓN, OBSERVACIÓN,
 *    NOMBRE DEL CONTACTO, CORREO, TELÉFONO.
 *  - Customer ID, Razón Social y Número NUNCA se tocan.
 *  - Solo rellena celdas VACÍAS: nunca sobrescribe.
 *  - Una celda vacía del CRM no borra nada.
 *  - Si un código viene repetido, gana el último dato no vacío.
 *  - Devuelve la lista de cambios para poder deshacerlos.
 *
 * Parámetro filasJson: arreglo JSON de objetos. Ejemplo:
 * [{"Código de Cliente":"C001","ESTADO CONTACTO":"Cotizado – seguimiento","CORREO":"compras@x.com"}]
 * (los nombres de columna se reconocen sin importar mayúsculas, tildes ni espacios)
 */

const HOJA = "BD-BBPE";
const COLUMNA_ID = "Customer ID";
const ALIAS_ID = ["customerid", "codigodecliente", "codigocliente", "codcliente"];
const ALIAS_ID_CRM = ["codigodecliente", "customerid", "codigocliente", "codcliente", "codigo", "title"];
const COLUMNAS_PERMITIDAS = ["ESTADO CONTACTO", "PRÓXIMA ACCIÓN", "OBSERVACIÓN", "NOMBRE DEL CONTACTO", "CORREO", "TELÉFONO"];
const PROTEGIDAS = ["customerid", "codigodecliente", "codigocliente", "codcliente", "razonsocial", "numerodecontacto", "numerocontacto", "numero", "nrodecontacto", "celular"];
const FILAS_ENCABEZADO = 10;

interface Cambio { fila: number; columna: number; direccion: string; valor: string | number; }
interface FilaCRM { [columna: string]: string | number | boolean | null; }
interface ResultadoAccion2 {
  ok: boolean;
  mensaje: string;
  registrosRecibidos: number;
  registrosActualizados: number;
  registrosSinCambios: number;
  celdasEscritas: number;
  celdasConDatos: number;
  columnasCruzadas: string[];
  noEncontrados: string[];
  cambios: Cambio[];
}

function main(workbook: ExcelScript.Workbook, filasJson: string): ResultadoAccion2 {
  const res: ResultadoAccion2 = { ok: false, mensaje: "", registrosRecibidos: 0, registrosActualizados: 0, registrosSinCambios: 0, celdasEscritas: 0, celdasConDatos: 0, columnasCruzadas: [], noEncontrados: [], cambios: [] };

  // ---- Entrada ----
  let filas: FilaCRM[] = [];
  try { filas = JSON.parse(filasJson || "[]") as FilaCRM[]; } catch (e) { res.mensaje = "filasJson no es un JSON válido."; return res; }
  if (!Array.isArray(filas) || !filas.length) { res.mensaje = "No se recibieron registros."; return res; }
  res.registrosRecibidos = filas.length;

  // ---- Hoja destino ----
  const ws = workbook.getWorksheet(HOJA);
  if (!ws) { res.mensaje = `No existe la hoja "${HOJA}".`; return res; }
  const permitidas = COLUMNAS_PERMITIDAS.map(normKey).filter(k => PROTEGIDAS.indexOf(k) < 0);
  const hoja = leerHoja(ws, COLUMNAS_PERMITIDAS);
  if (!hoja.ok) { res.mensaje = hoja.error; return res; }
  const usables = permitidas.filter(k => hoja.columnas.has(k));
  if (!usables.length) { res.mensaje = "La hoja no tiene ninguna de las 6 columnas de gestión."; return res; }
  res.columnasCruzadas = COLUMNAS_PERMITIDAS.filter(c => usables.indexOf(normKey(c)) >= 0);

  // ---- Fusionar registros por código (gana el último dato no vacío) ----
  const fusion = new Map<string, { original: string; valores: Map<string, string | number> }>();
  const orden: string[] = [];
  for (const fila of filas) {
    let idOriginal = "";
    const valores = new Map<string, string | number>();
    for (const nombre of Object.keys(fila)) {
      const k = normKey(nombre); const v = fila[nombre];
      if (ALIAS_ID_CRM.indexOf(k) >= 0 && !idOriginal && !vacio(v as Celda)) { idOriginal = String(v).trim(); continue; }
      if (usables.indexOf(k) < 0) continue;            // columnas no permitidas se ignoran
      if (v === null || v === undefined || typeof v === "boolean") continue;
      const limpio = typeof v === "string" ? v.trim() : v;
      if (vacio(limpio)) continue;                     // vacío en el CRM: no borra nada
      valores.set(k, limpio);
    }
    const clave = claveCodigo(idOriginal);
    if (!clave) continue;
    if (!fusion.has(clave)) { fusion.set(clave, { original: idOriginal, valores: new Map<string, string | number>() }); orden.push(clave); }
    const destino = fusion.get(clave);
    if (destino) valores.forEach((v, k) => destino.valores.set(k, v));
  }

  // ---- Calcular escrituras (solo celdas vacías) ----
  const escrituras: Escritura[] = [];
  for (const clave of orden) {
    const reg = fusion.get(clave); if (!reg) continue;
    const filasHoja = hoja.filasPorCodigo.get(clave);
    if (!filasHoja) { res.noEncontrados.push(reg.original); continue; }
    let algo = false;
    for (const fila of filasHoja) {
      reg.valores.forEach((valor, k) => {
        const actual = hoja.valorActual(k, fila);
        if (!vacio(actual)) { res.celdasConDatos++; return; }
        escrituras.push({ fila, columna: hoja.columnas.get(k) as number, valor, formato: "" });
        algo = true;
      });
    }
    if (algo) res.registrosActualizados++; else res.registrosSinCambios++;
  }

  // ---- Candado final y escritura ----
  const colsPermitidas = new Set(usables.map(k => hoja.columnas.get(k) as number));
  if (escrituras.some(e => !colsPermitidas.has(e.columna))) { res.mensaje = "Bloqueado: se intentó escribir fuera de las columnas permitidas. No se escribió nada."; return res; }
  escribir(ws, escrituras);
  res.cambios = escrituras.map(e => ({ fila: e.fila, columna: e.columna, direccion: direccion(e.fila, e.columna), valor: e.valor }));
  res.celdasEscritas = escrituras.length;
  res.ok = true;
  res.mensaje = `Acción 2: ${res.registrosActualizados} registro(s) actualizados, ${res.celdasEscritas} celda(s) escritas` +
    (res.celdasConDatos ? `, ${res.celdasConDatos} ya tenían datos (intactas)` : "") +
    (res.noEncontrados.length ? `, ${res.noEncontrados.length} código(s) no existen en la hoja` : "") + ".";
  console.log(res.mensaje);
  return res;
}

/* ================= Utilidades compartidas (idénticas en los 3 scripts) ================= */
type Celda = string | number | boolean;
interface Escritura { fila: number; columna: number; valor: string | number; formato: string; }
interface HojaLeida {
  ok: boolean; error: string;
  columnas: Map<string, number>;
  filasPorCodigo: Map<string, number[]>;
  valorActual: (clave: string, fila: number) => Celda;
}

function normKey(h: Celda): string {
  return String(h === null || h === undefined ? "" : h).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]/g, "");
}
function claveCodigo(v: Celda): string {
  let s = String(v === null || v === undefined ? "" : v).trim().replace(/^'/, "").replace(/\s+/g, "").toUpperCase();
  if (/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, "");
  return s;
}
function vacio(v: Celda): boolean { return v === null || v === undefined || String(v).trim() === ""; }
function letraColumna(c: number): string { let s = ""; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
function direccion(fila: number, columna: number): string { return `${letraColumna(columna)}${fila + 1}`; }

function leerHoja(ws: ExcelScript.Worksheet, columnasNecesarias: string[]): HojaLeida {
  const vacia: HojaLeida = { ok: false, error: "", columnas: new Map<string, number>(), filasPorCodigo: new Map<string, number[]>(), valorActual: () => "" };
  const usado = ws.getUsedRange(true);
  if (!usado) { vacia.error = `La hoja "${ws.getName()}" está vacía.`; return vacia; }
  const r0 = usado.getRowIndex(), c0 = usado.getColumnIndex(), nf = usado.getRowCount(), nc = usado.getColumnCount();
  const cabecera = ws.getRangeByIndexes(r0, c0, Math.min(FILAS_ENCABEZADO, nf), nc).getValues();
  let filaEnc = -1, colId = -1;
  for (let r = 0; r < cabecera.length && filaEnc < 0; r++) {
    const claves = cabecera[r].map(normKey);
    const alias = [normKey(COLUMNA_ID)].concat(ALIAS_ID);
    for (const a of alias) { const i = claves.indexOf(a); if (i >= 0) { filaEnc = r; colId = i; break; } }
  }
  if (filaEnc < 0) { vacia.error = `No se encontró la columna "${COLUMNA_ID}" en las ${FILAS_ENCABEZADO} primeras filas de "${ws.getName()}".`; return vacia; }
  const columnas = new Map<string, number>();
  cabecera[filaEnc].forEach((h, i) => { const k = normKey(h); if (k && !columnas.has(k)) columnas.set(k, c0 + i); });
  const primera = r0 + filaEnc + 1, n = r0 + nf - primera;
  const filasPorCodigo = new Map<string, number[]>();
  const existentes = new Map<string, Celda[]>();
  if (n > 0) {
    const ids = ws.getRangeByIndexes(primera, c0 + colId, n, 1).getValues();
    ids.forEach((fila, i) => { const k = claveCodigo(fila[0]); if (!k) return; const lst = filasPorCodigo.get(k) || []; lst.push(primera + i); filasPorCodigo.set(k, lst); });
    for (const nombre of columnasNecesarias) {
      const k = normKey(nombre), col = columnas.get(k);
      if (col === undefined) continue;
      existentes.set(k, ws.getRangeByIndexes(primera, col, n, 1).getFormulas().map(f => f[0]));
    }
  }
  return {
    ok: true, error: "", columnas, filasPorCodigo,
    valorActual: (clave: string, fila: number): Celda => { const arr = existentes.get(clave); return arr ? arr[fila - primera] : ""; },
  };
}
function escribir(ws: ExcelScript.Worksheet, escrituras: Escritura[]): void {
  const porColumna = new Map<number, Escritura[]>();
  for (const e of escrituras) { const l = porColumna.get(e.columna) || []; l.push(e); porColumna.set(e.columna, l); }
  porColumna.forEach((lista, col) => {
    lista.sort((a, b) => a.fila - b.fila);
    let i = 0;
    while (i < lista.length) {
      let j = i;
      while (j + 1 < lista.length && lista[j + 1].fila === lista[j].fila + 1 && lista[j + 1].formato === lista[i].formato) j++;
      const tramo = lista.slice(i, j + 1);
      const rango = ws.getRangeByIndexes(tramo[0].fila, col, tramo.length, 1);
      if (tramo[0].formato) rango.setNumberFormat(tramo[0].formato);
      rango.setValues(tramo.map(e => [textoSeguro(e.valor)]));
      i = j + 1;
    }
  });
}
function textoSeguro(v: string | number): string | number {
  if (typeof v === "string" && /^[=+\-@]/.test(v) && isNaN(Number(v))) return "'" + v;
  return v;
}
