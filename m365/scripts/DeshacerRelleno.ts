/**
 * Rellenado CRM · Deshacer — Office Script
 * ------------------------------------------------------------------
 * Deshace un relleno anterior usando la lista de "cambios" que devolvieron
 * RellenarAccion1 o RellenarAccion2 (guardada en la lista CRM_RegistroRelleno).
 *
 * Como esos scripts SOLO escriben en celdas vacías, deshacer = volver a dejar
 * la celda vacía. Por seguridad, solo se limpia una celda si todavía contiene
 * exactamente el valor que escribió el script; si alguien la modificó después,
 * se respeta y se informa como "omitida".
 *
 * Parámetro cambiosJson: el arreglo "cambios" devuelto por la acción, en JSON.
 */

const HOJA = "BD-BBPE";

interface Cambio { fila: number; columna: number; direccion: string; valor: string | number; }
interface ResultadoDeshacer { ok: boolean; mensaje: string; restauradas: number; omitidas: number; direccionesOmitidas: string[]; }

function main(workbook: ExcelScript.Workbook, cambiosJson: string): ResultadoDeshacer {
  const res: ResultadoDeshacer = { ok: false, mensaje: "", restauradas: 0, omitidas: 0, direccionesOmitidas: [] };
  let cambios: Cambio[] = [];
  try { cambios = JSON.parse(cambiosJson || "[]") as Cambio[]; } catch (e) { res.mensaje = "cambiosJson no es un JSON válido."; return res; }
  if (!Array.isArray(cambios) || !cambios.length) { res.mensaje = "No hay cambios que deshacer."; return res; }
  const ws = workbook.getWorksheet(HOJA);
  if (!ws) { res.mensaje = `No existe la hoja "${HOJA}".`; return res; }

  for (const c of cambios) {
    if (typeof c.fila !== "number" || typeof c.columna !== "number") continue;
    const celda = ws.getRangeByIndexes(c.fila, c.columna, 1, 1);
    const actual = celda.getValues()[0][0];
    if (igual(actual, c.valor)) { celda.setValues([[""]]); res.restauradas++; }
    else { res.omitidas++; res.direccionesOmitidas.push(c.direccion || `${c.fila},${c.columna}`); }
  }
  res.ok = true;
  res.mensaje = `Deshacer: ${res.restauradas} celda(s) vuelven a quedar vacías` + (res.omitidas ? `; ${res.omitidas} se respetaron porque alguien las cambió después.` : ".");
  console.log(res.mensaje);
  return res;
}

function igual(a: string | number | boolean, b: string | number): boolean {
  if (typeof a === "number" && typeof b === "number") return Math.abs(a - b) < 1e-9;
  // El apóstrofo inicial que protege textos como "=..." no forma parte del valor
  const sa = String(a).trim().replace(/^'/, ""), sb = String(b).trim().replace(/^'/, "");
  return sa === sb;
}
