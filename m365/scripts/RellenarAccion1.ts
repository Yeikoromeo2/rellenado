/**
 * Rellenado CRM · Acción 1 — Office Script
 * ------------------------------------------------------------------
 * Rellena FECHA CONTACTO y RESPONSABLE en la hoja BD-BBPE para una lista
 * de códigos de cliente. Mismas reglas que el complemento "Rellenado CRM":
 *  - Busca por Customer ID (también acepta "Código de Cliente").
 *  - Encabezados detectados en las 10 primeras filas (sin importar mayúsculas,
 *    tildes ni espacios).
 *  - Solo escribe en FECHA CONTACTO y RESPONSABLE.
 *  - Customer ID, Razón Social y Número NUNCA se tocan.
 *  - Solo rellena celdas VACÍAS: nunca sobrescribe.
 *  - Si un código aparece en varias filas, rellena todas.
 *  - Si falta una columna, avisa y no crea nada.
 *  - Devuelve la lista de cambios para poder deshacerlos.
 *
 * Uso manual: Excel para la web → Automatizar → este script → Ejecutar.
 * Uso automático: Power Automate → "Ejecutar script" (Excel Online Business).
 */

const HOJA = "BD-BBPE";
const COLUMNA_ID = "Customer ID";
const ALIAS_ID = ["customerid", "codigodecliente", "codigocliente", "codcliente"];
const COLUMNA_FECHA = "FECHA CONTACTO";
const COLUMNA_RESPONSABLE = "RESPONSABLE";
const FORMATO_FECHA = "mm/dd/yyyy";
const RESPONSABLES = ["Cinthya Sandoval", "Fiorella Viena", "Harold Sovero", "Lorena Suncion"];
const PROTEGIDAS = ["customerid", "codigodecliente", "codigocliente", "codcliente", "razonsocial", "numerodecontacto", "numerocontacto", "numero", "nrodecontacto", "celular"];
const FILAS_ENCABEZADO = 10;
const ZONA_HORARIA = -5; // Lima (UTC-5): se usa solo si no se envía fecha

interface Cambio { fila: number; columna: number; direccion: string; valor: string | number; }
interface ResultadoAccion1 {
  ok: boolean;
  mensaje: string;
  fecha: string;
  responsable: string;
  codigosSolicitados: number;
  codigosEncontrados: number;
  filasRellenadas: number;
  celdasEscritas: number;
  celdasConDatos: number;
  noEncontrados: string[];
  cambios: Cambio[];
}

/**
 * @param codigos Códigos separados por salto de línea, coma o punto y coma.
 * @param responsable Uno de: Cinthya Sandoval, Fiorella Viena, Harold Sovero, Lorena Suncion.
 * @param fecha (opcional) Fecha en formato AAAA-MM-DD. Si se deja vacía, usa la fecha de hoy en Lima.
 */
function main(workbook: ExcelScript.Workbook, codigos: string, responsable: string, fecha?: string): ResultadoAccion1 {
  const res: ResultadoAccion1 = { ok: false, mensaje: "", fecha: "", responsable: "", codigosSolicitados: 0, codigosEncontrados: 0, filasRellenadas: 0, celdasEscritas: 0, celdasConDatos: 0, noEncontrados: [], cambios: [] };

  // ---- Validar entradas ----
  const lista = limpiarCodigos(codigos);
  res.codigosSolicitados = lista.length;
  if (!lista.length) { res.mensaje = "No se recibió ningún código."; return res; }
  const resp = RESPONSABLES.find(r => normKey(r) === normKey(responsable));
  if (!resp) { res.mensaje = `Responsable no válido: "${responsable}". Usa: ${RESPONSABLES.join(", ")}.`; return res; }
  res.responsable = resp;
  const iso = (fecha || "").trim() || hoyLima();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) { res.mensaje = `Fecha no válida: "${fecha}". Usa AAAA-MM-DD.`; return res; }
  res.fecha = iso;

  // ---- Leer hoja ----
  const ws = workbook.getWorksheet(HOJA);
  if (!ws) { res.mensaje = `No existe la hoja "${HOJA}".`; return res; }
  const hoja = leerHoja(ws, [COLUMNA_FECHA, COLUMNA_RESPONSABLE]);
  if (!hoja.ok) { res.mensaje = hoja.error; return res; }
  const faltan = [COLUMNA_FECHA, COLUMNA_RESPONSABLE].filter(c => !hoja.columnas.has(normKey(c)));
  if (faltan.length) { res.mensaje = `Faltan columnas en "${HOJA}": ${faltan.join(", ")}. No se escribió nada.`; return res; }
  const kFecha = normKey(COLUMNA_FECHA), kResp = normKey(COLUMNA_RESPONSABLE);
  if (PROTEGIDAS.indexOf(kFecha) >= 0 || PROTEGIDAS.indexOf(kResp) >= 0) { res.mensaje = "Configuración inválida: columna protegida."; return res; }

  // ---- Calcular escrituras (solo celdas vacías) ----
  const serial = serialExcel(iso);
  const escrituras: Escritura[] = [];
  for (const cod of lista) {
    const filas = hoja.filasPorCodigo.get(cod.clave);
    if (!filas) { res.noEncontrados.push(cod.original); continue; }
    res.codigosEncontrados++;
    for (const fila of filas) {
      let tocada = false;
      const objetivos: [string, string | number, boolean][] = [[kFecha, serial, true], [kResp, resp, false]];
      for (const [clave, valor, esFecha] of objetivos) {
        const actual = hoja.valorActual(clave, fila);
        if (!vacio(actual)) { res.celdasConDatos++; continue; }
        escrituras.push({ fila, columna: hoja.columnas.get(clave) as number, valor, formato: esFecha ? FORMATO_FECHA : "" });
        tocada = true;
      }
      if (tocada) res.filasRellenadas++;
    }
  }

  // ---- Candado final y escritura ----
  const permitidas = new Set([hoja.columnas.get(kFecha) as number, hoja.columnas.get(kResp) as number]);
  if (escrituras.some(e => !permitidas.has(e.columna))) { res.mensaje = "Bloqueado: se intentó escribir fuera de las columnas permitidas. No se escribió nada."; return res; }
  escribir(ws, escrituras);
  res.cambios = escrituras.map(e => ({ fila: e.fila, columna: e.columna, direccion: direccion(e.fila, e.columna), valor: e.valor }));
  res.celdasEscritas = escrituras.length;
  res.ok = true;
  res.mensaje = `Acción 1: ${res.filasRellenadas} fila(s) rellenadas, ${res.celdasEscritas} celda(s) escritas` +
    (res.celdasConDatos ? `, ${res.celdasConDatos} ya tenían datos (intactas)` : "") +
    (res.noEncontrados.length ? `, ${res.noEncontrados.length} código(s) no existen en la hoja` : "") + ".";
  console.log(res.mensaje);
  return res;
}

/* ================= Utilidades compartidas (idénticas en los 3 scripts) ================= */
type Celda = string | number | boolean;
interface Escritura { fila: number; columna: number; valor: string | number; formato: string; }
interface CodigoLimpio { clave: string; original: string; }
interface HojaLeida {
  ok: boolean; error: string;
  columnas: Map<string, number>;          // clave normalizada -> índice absoluto de columna
  filasPorCodigo: Map<string, number[]>;  // código normalizado -> filas absolutas
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
function limpiarCodigos(texto: string): CodigoLimpio[] {
  const vistos = new Set<string>(); const out: CodigoLimpio[] = [];
  for (const parte of String(texto || "").split(/[\r\n,;]+/)) {
    const original = parte.split("\t")[0].trim(); const clave = claveCodigo(original);
    if (!clave || vistos.has(clave)) continue;
    vistos.add(clave); out.push({ clave, original });
  }
  return out;
}
function hoyLima(): string {
  const d = new Date(Date.now() + ZONA_HORARIA * 3600000);
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
}
function serialExcel(iso: string): number {
  const p = iso.split("-").map(Number);
  return (Date.UTC(p[0], p[1] - 1, p[2]) - Date.UTC(1899, 11, 30)) / 86400000;
}
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
      // Fórmulas: una celda con fórmula NO se considera vacía aunque muestre ""
      existentes.set(k, ws.getRangeByIndexes(primera, col, n, 1).getFormulas().map(f => f[0]));
    }
  }
  return {
    ok: true, error: "", columnas, filasPorCodigo,
    valorActual: (clave: string, fila: number): Celda => { const arr = existentes.get(clave); return arr ? arr[fila - primera] : ""; },
  };
}
/** Escribe agrupando celdas contiguas de una misma columna (menos llamadas a Excel). */
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
/** Un texto que empieza con "=" se guarda como texto, nunca como fórmula. */
function textoSeguro(v: string | number): string | number {
  if (typeof v === "string" && /^[=+\-@]/.test(v) && isNaN(Number(v))) return "'" + v;
  return v;
}
