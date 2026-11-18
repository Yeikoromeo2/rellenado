/**
 * CRM Difusión · Generar lote — Office Script (se ejecuta sobre CRM_Motor.xlsx)
 * ------------------------------------------------------------------
 * Motor de la herramienta de difusión llevado a Microsoft 365:
 *  - Elige N contactos al azar que nunca fueron asignados.
 *  - Respeta No contactar, consentimiento, frecuencia por número y teléfonos repetidos.
 *  - Genera un mensaje ÚNICO por contacto (plantillas + Spintax + bloques):
 *      saludo a la RAZÓN SOCIAL · presentación del EJECUTIVO · cuerpo ·
 *      SOLICITUD DE DATOS (siempre) · pregunta · despedida · opción de salida.
 *  - Los códigos de cliente son de uso interno: NUNCA aparecen en el mensaje.
 *  - Asigna una hora programada dentro de la jornada (09:00–18:00), con
 *    intervalos de distribución normal, pausas por fatiga y tope diario.
 *
 * Toda la configuración se edita en las hojas de CRM_Motor.xlsx:
 *   Plantillas · Bloques · Ejecutivos · Ajustes · Respaldos
 *
 * Modo automático (Power Automate): se pasa entradaJson y devuelve el lote.
 * Modo prueba (manual en Excel): deja entradaJson vacío; usa la hoja "Prueba"
 * y escribe el resultado en la hoja "Resultado".
 */

type Celda = string | number | boolean;
interface Mapa { [clave: string]: string; }
interface Bloque { nombre: string; prob: number; opciones: string[]; }
interface Ejecutivo { nombre: string; cargo: string; correo: string; }
interface Ajustes {
  inicio: string; fin: string; sabado: boolean;
  gapMin: number; gapMax: number; escMin: number; escMax: number;
  fatMin: number; fatMax: number; desMin: number; desMax: number;
  modo: string; piso: number; frecuenciaDias: number; topeDiario: number;
  soloConsentimiento: boolean; incluirSalida: boolean; miEmpresa: string;
  zona: number; tamano: number;
}
interface Config { plantillas: string[]; bloques: Map<string, Bloque>; ejecutivos: Ejecutivo[]; ajustes: Ajustes; respaldos: Mapa; }
interface ClienteIn { codigo: string; razon: string; telefono: string; consentimiento?: string; nombre?: string; ciudad?: string; }
interface EnvioPrevio { telefono: string; fecha: string; }
interface Entrada {
  ahoraLocal?: string; ejecutivo?: string; tamano?: number; loteNumero?: number;
  clientes: ClienteIn[]; asignados?: string[]; noContactar?: string[];
  ultimosEnvios?: EnvioPrevio[]; enviadosHoy?: number;
}
interface Generado { texto: string; plantilla: number; combo: string; bytes: number; caracteres: number; }
interface ItemLote {
  codigo: string; razonSocial: string; telefono: string;
  mensaje: string; bytes: number; caracteres: number; plantilla: number;
  horaProgramada: string; horaLocal: string; estado: string; descansoAntes: boolean;
  enlaceWhatsApp: string; ejecutivo: string; correoEjecutivo: string;
}
interface Omitidos { noContactar: number; sinConsentimiento: number; frecuencia: number; telefonoRepetido: number; sinTelefono: number; yaAsignados: number; }
interface Resumen { contactos: number; hoy: number; reprogramados: number; inicio: string; cierre: string; cierreHoy: string; ritmoPorHora: number; factorAjuste: number; topeAlcanzado: boolean; }
interface SalidaLote { ok: boolean; mensaje: string; loteNombre: string; ejecutivo: string; correoEjecutivo: string; items: ItemLote[]; omitidos: Omitidos; resumen: Resumen; }

const ALIAS: Mapa = {
  empresa: "razon", razon_social: "razon", razonsocial: "razon", cliente: "razon",
  ejecutivo: "ejecutivo", asesor: "ejecutivo", asesora: "ejecutivo", firma: "ejecutivo",
  cargo: "cargo", mi_empresa: "mi_empresa",
  nombre: "nombre", nombre_del_contacto: "nombre", contacto: "nombre", ciudad: "ciudad",
  codigo: "interno", codigo_cliente: "interno", codigo_de_cliente: "interno", customer_id: "interno", cod_cliente: "interno",
  telefono: "interno", numero: "interno", numero_de_contacto: "interno", celular: "interno",
};
const ORDEN_BLOQUES = ["saludo", "presentacion", "datos", "pregunta", "despedida", "salida"];
const DIA = 86400000;

function main(workbook: ExcelScript.Workbook, entradaJson?: string): SalidaLote {
  const cfg = leerConfig(workbook);
  const manual = !entradaJson || !entradaJson.trim() || entradaJson.trim() === "{}";
  let entrada: Entrada;
  if (manual) entrada = entradaDePrueba(workbook, cfg);
  else {
    try { entrada = JSON.parse(entradaJson as string) as Entrada; }
    catch (e) { return salidaError("entradaJson no es un JSON válido."); }
  }
  const salida = generar(entrada, cfg);
  if (manual) escribirResultado(workbook, salida);
  console.log(salida.mensaje);
  return salida;
}

/* =================== Generación del lote =================== */
function generar(e: Entrada, cfg: Config): SalidaLote {
  const a = cfg.ajustes;
  const ejec = cfg.ejecutivos.find(x => normKey(x.nombre) === normKey(e.ejecutivo || "")) || cfg.ejecutivos[0];
  if (!ejec) return salidaError("No hay ejecutivos en la hoja Ejecutivos.");
  if (!cfg.plantillas.length) return salidaError("No hay plantillas activas en la hoja Plantillas.");
  const tamano = Math.max(1, Math.round(e.tamano || a.tamano || 30));
  const ahora = parseLocal(e.ahoraLocal) || Date.now() + a.zona * 3600000;
  const hoy = claveDia(ahora);

  // --- Filtrar candidatos con las reglas de protección ---
  const om: Omitidos = { noContactar: 0, sinConsentimiento: 0, frecuencia: 0, telefonoRepetido: 0, sinTelefono: 0, yaAsignados: 0 };
  const asignados = new Set((e.asignados || []).map(claveCodigo));
  const bloqueados = new Set<string>();
  (e.noContactar || []).forEach(x => { const t = String(x || "").trim(); if (!t) return; bloqueados.add(claveCodigo(t)); const ph = normalizarTelefono(t); if (telefonoValido(ph)) bloqueados.add(ph); });
  const ultimo = new Map<string, number>();
  (e.ultimosEnvios || []).forEach(x => { const ph = normalizarTelefono(x.telefono); const t = parseLocal(x.fecha); if (ph && t !== null && (ultimo.get(ph) || 0) < t) ultimo.set(ph, t); });
  const telefonosVistos = new Set<string>();
  const candidatos: ClienteIn[] = [];
  for (const c0 of e.clientes || []) {
    const codigo = String(c0.codigo || "").trim(); if (!codigo) continue;
    if (asignados.has(claveCodigo(codigo))) { om.yaAsignados++; continue; }
    const tel = normalizarTelefono(c0.telefono);
    if (!telefonoValido(tel)) { om.sinTelefono++; continue; }
    if (bloqueados.has(claveCodigo(codigo)) || bloqueados.has(tel)) { om.noContactar++; continue; }
    if (a.soloConsentimiento && !verdadero(c0.consentimiento)) { om.sinConsentimiento++; continue; }
    const t = ultimo.get(tel);
    if (a.frecuenciaDias > 0 && t !== undefined && ahora - t < a.frecuenciaDias * DIA) { om.frecuencia++; continue; }
    candidatos.push({ codigo, razon: String(c0.razon || "").trim().replace(/\s+/g, " "), telefono: tel, consentimiento: c0.consentimiento, nombre: c0.nombre, ciudad: c0.ciudad });
  }
  mezclar(candidatos);
  const elegidos: ClienteIn[] = [];
  for (const c of candidatos) {
    if (elegidos.length >= tamano) break;
    if (telefonosVistos.has(c.telefono)) { om.telefonoRepetido++; continue; }
    telefonosVistos.add(c.telefono); elegidos.push(c);
  }
  if (!elegidos.length) {
    const s = salidaError("No hay contactos disponibles que cumplan las reglas (No contactar, consentimiento, frecuencia).");
    s.omitidos = om; return s;
  }

  // --- Mensajes únicos (sin repetir plantilla/combinación con el anterior) ---
  let previo: Generado | null = null;
  const generados: Generado[] = [];
  for (const c of elegidos) { const g = componer(c, ejec, cfg, previo); generados.push(g); previo = g; }

  // --- Programación dentro de la jornada ---
  const plan = planificar(generados, cfg, ahora, Math.max(0, e.enviadosHoy || 0));

  const items: ItemLote[] = elegidos.map((c, i) => ({
    codigo: c.codigo, razonSocial: c.razon, telefono: c.telefono,
    mensaje: generados[i].texto, bytes: generados[i].bytes, caracteres: generados[i].caracteres, plantilla: generados[i].plantilla + 1,
    horaProgramada: isoConZona(plan.horas[i], a.zona), horaLocal: fmtLocal(plan.horas[i]),
    estado: claveDia(plan.horas[i]) > hoy ? "Reprogramado" : "Pendiente",
    descansoAntes: plan.descanso[i],
    enlaceWhatsApp: `https://wa.me/${c.telefono}?text=${encodeURIComponent(generados[i].texto)}`,
    ejecutivo: ejec.nombre, correoEjecutivo: ejec.correo,
  }));
  const numero = e.loteNumero && e.loteNumero > 0 ? e.loteNumero : 1;
  const loteNombre = `Lote #${numero} - ${fmtFecha(ahora)}`;
  const hoyItems = items.filter(x => x.estado === "Pendiente");
  const resumen: Resumen = {
    contactos: items.length, hoy: hoyItems.length, reprogramados: items.length - hoyItems.length,
    inicio: fmtLocal(plan.inicio), cierre: items.length ? items[items.length - 1].horaLocal : "",
    cierreHoy: hoyItems.length ? hoyItems[hoyItems.length - 1].horaLocal : "",
    ritmoPorHora: plan.ritmo, factorAjuste: Math.round(plan.factor * 100) / 100, topeAlcanzado: plan.tope,
  };
  const omTxt = om.noContactar + om.sinConsentimiento + om.frecuencia + om.telefonoRepetido;
  return {
    ok: true, loteNombre, ejecutivo: ejec.nombre, correoEjecutivo: ejec.correo, items, omitidos: om, resumen,
    mensaje: `${loteNombre}: ${items.length} contactos firmados por ${ejec.nombre}. ` +
      (resumen.hoy ? `${resumen.hoy} hoy (cierre ${resumen.cierreHoy})` : "Ninguno hoy") +
      (resumen.reprogramados ? `, ${resumen.reprogramados} reprogramados al siguiente día hábil` : "") +
      (omTxt ? `. Omitidos por protección: ${omTxt}` : "") + ".",
  };
}

/* =================== Motor de texto =================== */
function valorVariable(nombre: string, c: ClienteIn, ej: Ejecutivo, cfg: Config): string {
  const k = slug(nombre), tipo = ALIAS[k];
  if (tipo === "interno") return "";                       // códigos y teléfonos: uso interno
  if (tipo === "ejecutivo") return ej.nombre;
  if (tipo === "cargo") return ej.cargo || "asesor comercial";
  if (tipo === "mi_empresa") return cfg.ajustes.miEmpresa;
  let v = "";
  if (tipo === "razon") v = c.razon;
  else if (tipo === "nombre") v = c.nombre || "";
  else if (tipo === "ciudad") v = c.ciudad || "";
  v = String(v || "").replace(/\s+/g, " ").trim();
  if (!v) v = cfg.respaldos[k] || "";
  return v;
}
function esVariableConocida(x: string, cfg: Config): boolean {
  const k = slug(x);
  return /^[\w\sáéíóúñ]+$/i.test(x) && !!k && (ALIAS[k] !== undefined || cfg.respaldos[k] !== undefined);
}
interface Render { texto: string; ruta: string; }
function spintax(plantilla: string, c: ClienteIn, ej: Ejecutivo, cfg: Config): Render {
  const vals: string[] = [], lits: string[] = [], ruta: number[] = [];
  let s = String(plantilla).replace(/\{\{\s*([^{}]+?)\s*\}\}/g, (_m: string, n: string) => { vals.push(valorVariable(n, c, ej, cfg)); return `\u0001${vals.length - 1}\u0002`; });
  const re = /\{([^{}]*)\}/;
  let m = re.exec(s), guarda = 0;
  while (m && guarda++ < 5000) {
    const dentro = m[1]; let rep: string;
    if (dentro.indexOf("|") >= 0) { const ops = dentro.split("|"); const i = Math.floor(Math.random() * ops.length); ruta.push(i); rep = ops[i]; }
    else if (esVariableConocida(dentro, cfg)) { vals.push(valorVariable(dentro, c, ej, cfg)); rep = `\u0001${vals.length - 1}\u0002`; }
    else { lits.push(m[0]); rep = `\u0003${lits.length - 1}\u0004`; }
    s = s.slice(0, m.index) + rep + s.slice(m.index + m[0].length);
    m = re.exec(s);
  }
  s = s.replace(/\u0003(\d+)\u0004/g, (_m: string, i: string) => lits[Number(i)]).replace(/\u0001(\d+)\u0002/g, (_m: string, i: string) => vals[Number(i)]);
  return { texto: s, ruta: ruta.join(".") };
}
function limpiar(t: string): string {
  return t.split("\n").map(l => l.replace(/[ \t]+/g, " ")
    .replace(/\s+(de|del|para|con)\s*([.,;:!?])/gi, "$2").replace(/\(\s*\)/g, "")
    .replace(/\s+([,.;:!?)])/g, "$1").replace(/([¡¿(])\s+/g, "$1").replace(/,\s*,+/g, ",")
    .replace(/^\s*[,;:]\s*/, "").replace(/¡\s*!/g, "").trim()).join("\n").replace(/\n{3,}/g, "\n\n").trim();
}
function componer(c: ClienteIn, ej: Ejecutivo, cfg: Config, evitar: Generado | null): Generado {
  const indices = cfg.plantillas.map((_p, i) => i);
  let res: Generado = { texto: "", plantilla: -1, combo: "", bytes: 0, caracteres: 0 };
  for (let intento = 0; intento < 30; intento++) {
    const pool = evitar && indices.length > 1 ? indices.filter(i => i !== evitar.plantilla) : indices;
    const ti = pool[Math.floor(Math.random() * pool.length)];
    const claves: string[] = [`T${ti}`];
    const bloque = (nombre: string): string => {
      const b = cfg.bloques.get(nombre);
      if (!b || !b.opciones.length || Math.random() * 100 >= b.prob) { claves.push(nombre + "-"); return ""; }
      if (nombre === "salida" && !cfg.ajustes.incluirSalida) { claves.push(nombre + "-"); return ""; }
      const i = Math.floor(Math.random() * b.opciones.length); const r = spintax(b.opciones[i], c, ej, cfg);
      claves.push(`${nombre}${i}:${r.ruta}`); return r.texto;
    };
    const sal = bloque("saludo"), pres = bloque("presentacion");
    // Si el saludo termina en letra (p. ej. "Saludos, Empresa X") se cierra con punto antes de la presentación
    const salCerrado = sal && pres && /[A-Za-z\u00C0-\u00FF0-9]$/.test(sal.trim()) ? sal.trim() + "." : sal;
    const cabeza = [salCerrado, pres].filter(x => !!x).join(Math.random() < 0.5 ? " " : "\n");
    const cuerpo = spintax(cfg.plantillas[ti], c, ej, cfg); claves.push("B" + cuerpo.ruta);
    const datos = bloque("datos"), preg = bloque("pregunta"), desp = bloque("despedida"), salida = bloque("salida");
    const sep = Math.random() < 0.5 ? "\n\n" : "\n"; claves.push(String(sep.length));
    const texto = limpiar([cabeza, cuerpo.texto, datos, preg, desp].filter(x => !!x && !!x.trim()).join(sep) + (salida ? "\n\n" + salida : ""));
    res = { texto, plantilla: ti, combo: claves.join("|"), bytes: bytesUtf8(texto), caracteres: Array.from(texto).length };
    if (!evitar || (res.combo !== evitar.combo && res.texto !== evitar.texto && (indices.length < 2 || res.plantilla !== evitar.plantilla))) break;
  }
  return res;
}
function bytesUtf8(t: string): number {
  let n = 0;
  for (const ch of Array.from(t)) { const cp = ch.codePointAt(0) || 0; n += cp < 0x80 ? 1 : cp < 0x800 ? 2 : cp < 0x10000 ? 3 : 4; }
  return n;
}

/* =================== Motor de tiempo =================== */
interface Plan { horas: number[]; descanso: boolean[]; inicio: number; factor: number; ritmo: number; tope: boolean; }
function gaussiana(): number { let u = 0, v = 0; while (!u) u = Math.random(); while (!v) v = Math.random(); return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v); }
function azar(a: number, b: number): number { return a + Math.random() * (b - a); }
function azarEntero(a: number, b: number): number { return Math.floor(azar(Math.min(a, b), Math.max(a, b) + 1)); }
function intervaloMs(a: Ajustes): number {
  const min = Math.max(5, a.gapMin), max = Math.max(min, a.gapMax), media = (min + max) / 2, sd = (max - min) / 5 || 1;
  let x = media, n = 0;
  do { x = media + gaussiana() * sd; } while ((x < min || x > max) && ++n < 25);
  return Math.round(Math.min(max, Math.max(min, x)) * 1000 + azar(0, 999));
}
function minutos(hhmm: string): number { const p = String(hhmm).split(":").map(Number); return (p[0] || 0) * 60 + (p[1] || 0); }
function inicioDia(t: number): number { return Math.floor(t / DIA) * DIA; }
function aLas(t: number, hhmm: string): number { return inicioDia(t) + minutos(hhmm) * 60000; }
function habil(t: number, a: Ajustes): boolean { const d = new Date(t).getUTCDay(); return (d >= 1 && d <= 5) || (d === 6 && a.sabado); }
function siguienteApertura(t: number, a: Ajustes): number {
  if (habil(t, a) && t < aLas(t, a.inicio)) return aLas(t, a.inicio);
  let d = inicioDia(t);
  for (let i = 0; i < 14; i++) { d += DIA; if (habil(d, a)) return aLas(d, a.inicio); }
  return aLas(d, a.inicio);
}
function planificar(gen: Generado[], cfg: Config, ahora: number, enviadosHoy: number): Plan {
  const a = cfg.ajustes;
  const piso = Math.max(5, a.piso) * 1000;
  // Duraciones base: intervalo humano + escritura + descanso por fatiga
  const dur = gen.map(() => ({ gap: 0, esc: 0, des: 0 }));
  let desde = 0, cada = azarEntero(a.fatMin, a.fatMax);
  gen.forEach((g, i) => {
    let des = 0;
    if (desde >= cada) { des = Math.round(azar(a.desMin, Math.max(a.desMin, a.desMax)) * 60000); desde = 0; cada = azarEntero(a.fatMin, a.fatMax); }
    dur[i] = { gap: intervaloMs(a), esc: Math.round(g.caracteres * azar(a.escMin, Math.max(a.escMin, a.escMax))), des };
    desde++;
  });
  let t = ahora;
  if (!(habil(t, a) && t >= aLas(t, a.inicio) && t < aLas(t, a.fin))) t = siguienteApertura(t, a);
  const conteo = new Map<string, number>(); conteo.set(claveDia(ahora), enviadosHoy);
  let tope = false;
  while (a.topeDiario > 0 && (conteo.get(claveDia(t)) || 0) >= a.topeDiario) { t = siguienteApertura(aLas(t, a.fin), a); tope = true; }
  const inicio = t, primerDia = claveDia(t);
  let fin = aLas(t, a.fin);
  // Ajuste proporcional: mayor factor ≤1 con el que todo cabe hoy sin bajar del piso
  let factor = 1;
  if (a.modo === "ajuste") {
    const disponible = fin - t - 15000;
    const total = (f: number): number => dur.reduce((s, d) => s + d.esc + Math.max(piso, d.gap * f) + (d.des ? Math.max(60000, d.des * f) : 0), 0);
    const cupoHoy = a.topeDiario > 0 ? a.topeDiario - (conteo.get(primerDia) || 0) : dur.length;
    if (cupoHoy >= dur.length && total(1) > disponible) {
      let lo = 0, hi = 1;
      for (let i = 0; i < 30; i++) { const mid = (lo + hi) / 2; if (total(mid) <= disponible) lo = mid; else hi = mid; }
      factor = lo;
    }
  }
  const horas: number[] = [], descanso: boolean[] = [];
  let hoyCuenta = 0, hoyFin = t;
  dur.forEach(d => {
    const f = claveDia(t) === primerDia ? factor : 1;
    const gap = Math.max(piso, Math.round(d.gap * f));
    const des = d.des ? Math.max(60000, Math.round(d.des * f)) : 0;
    t += des + gap + d.esc;
    if (t > fin) { const ap = siguienteApertura(fin, a); t = ap + Math.min(gap, 90000) + d.esc; fin = aLas(ap, a.fin); }
    while (a.topeDiario > 0 && (conteo.get(claveDia(t)) || 0) >= a.topeDiario) {
      const ap = siguienteApertura(aLas(t, a.fin), a); t = ap + Math.min(gap, 90000) + d.esc; fin = aLas(ap, a.fin); tope = true;
    }
    conteo.set(claveDia(t), (conteo.get(claveDia(t)) || 0) + 1);
    horas.push(t); descanso.push(des > 0);
    if (claveDia(t) === primerDia) { hoyCuenta++; hoyFin = t; }
  });
  const horasHoy = (hoyFin - inicio) / 3600000;
  return { horas, descanso, inicio, factor, ritmo: horasHoy > 0 ? Math.round((hoyCuenta / horasHoy) * 10) / 10 : 0, tope };
}

/* =================== Configuración (hojas de CRM_Motor.xlsx) =================== */
function leerTabla(wb: ExcelScript.Workbook, nombre: string): Mapa[] {
  const ws = wb.getWorksheet(nombre); if (!ws) return [];
  const usado = ws.getUsedRange(true); if (!usado) return [];
  const v = usado.getValues(); if (v.length < 2) return [];
  const cab = v[0].map(h => slug(String(h)));
  const out: Mapa[] = [];
  for (let r = 1; r < v.length; r++) {
    const fila: Mapa = {}; let alguno = false;
    cab.forEach((h, i) => { if (!h) return; const x = celdaTexto(v[r][i], h); fila[h] = x; if (x) alguno = true; });
    if (alguno) out.push(fila);
  }
  return out;
}
function celdaTexto(v: Celda, columna: string): string {
  if (typeof v === "number" && v > 0 && v < 1 && /valor|hora/.test(columna)) {   // horas escritas como 09:00 que Excel guardó como fracción
    const m = Math.round(v * 1440); return `${String(Math.floor(m / 60)).padStart(2, "0")}:${String(m % 60).padStart(2, "0")}`;
  }
  return String(v === null || v === undefined ? "" : v).trim();
}
function leerConfig(wb: ExcelScript.Workbook): Config {
  const plantillas = leerTabla(wb, "Plantillas").filter(f => f.texto && !/^(no|n|0|false)$/i.test(f.activa || "si")).map(f => f.texto);
  const bloques = new Map<string, Bloque>();
  leerTabla(wb, "Bloques").forEach(f => {
    const nombre = slug(f.bloque); if (ORDEN_BLOQUES.indexOf(nombre) < 0 || !f.texto) return;
    const b = bloques.get(nombre) || { nombre, prob: 0, opciones: [] };
    const p = Number(String(f.probabilidad || "").replace("%", ""));
    if (!isNaN(p)) b.prob = Math.max(b.prob, p <= 1 && /\./.test(String(f.probabilidad)) ? p * 100 : p);
    b.opciones.push(f.texto); bloques.set(nombre, b);
  });
  const ejecutivos = leerTabla(wb, "Ejecutivos").filter(f => f.nombre).map(f => ({ nombre: f.nombre, cargo: f.cargo || "asesor comercial", correo: f.correo || f.correo_teams || "" }));
  const aj: Mapa = {}; leerTabla(wb, "Ajustes").forEach(f => { if (f.clave) aj[slug(f.clave)] = f.valor || ""; });
  const respaldos: Mapa = {}; leerTabla(wb, "Respaldos").forEach(f => { if (f.variable) respaldos[slug(f.variable)] = f.valor || ""; });
  const num = (k: string, d: number): number => { const x = Number(String(aj[k] === undefined ? "" : aj[k]).replace(",", ".")); return aj[k] !== undefined && aj[k] !== "" && !isNaN(x) ? x : d; };
  const bool = (k: string, d: boolean): boolean => aj[k] === undefined || aj[k] === "" ? d : verdadero(aj[k]);
  const ajustes: Ajustes = {
    inicio: aj.inicio || "09:00", fin: aj.fin || "18:00", sabado: bool("incluir_sabados", false),
    gapMin: num("intervalo_min_seg", 45), gapMax: num("intervalo_max_seg", 180),
    escMin: num("escritura_min_ms", 40), escMax: num("escritura_max_ms", 90),
    fatMin: num("fatiga_cada_min", 6), fatMax: num("fatiga_cada_max", 10),
    desMin: num("descanso_min_minutos", 4), desMax: num("descanso_max_minutos", 10),
    modo: /ajuste/i.test(aj.modo_si_no_cabe || "") ? "ajuste" : "desborde", piso: num("piso_minimo_seg", 30),
    frecuenciaDias: num("frecuencia_dias", 7), topeDiario: num("tope_diario", 40),
    soloConsentimiento: bool("solo_con_consentimiento", false), incluirSalida: bool("incluir_opcion_salida", true),
    miEmpresa: aj.mi_empresa || "", zona: num("zona_horaria", -5), tamano: num("tamano_lote", 30),
  };
  return { plantillas, bloques, ejecutivos, ajustes, respaldos };
}

/* =================== Modo prueba (manual dentro de CRM_Motor.xlsx) =================== */
function entradaDePrueba(wb: ExcelScript.Workbook, cfg: Config): Entrada {
  const clientes = leerTabla(wb, "Prueba").map(f => ({
    codigo: f.codigo_de_cliente || f.codigo || "", razon: f.razon_social || f.razon || "",
    telefono: f.numero_de_contacto || f.telefono || "", consentimiento: f.consentimiento || "", nombre: f.nombre || "", ciudad: f.ciudad || "",
  }));
  return { clientes, ejecutivo: cfg.ejecutivos.length ? cfg.ejecutivos[0].nombre : "", tamano: cfg.ajustes.tamano, loteNumero: 1 };
}
function escribirResultado(wb: ExcelScript.Workbook, s: SalidaLote): void {
  let ws = wb.getWorksheet("Resultado");
  if (!ws) ws = wb.addWorksheet("Resultado");
  const usado = ws.getUsedRange(); if (usado) usado.clear();
  const filas: Celda[][] = [["#", "Código (interno)", "Razón Social", "Teléfono", "Hora programada", "Estado", "Bytes", "Caracteres", "Plantilla", "Mensaje generado", "Enlace WhatsApp"]];
  s.items.forEach((x, i) => filas.push([i + 1, x.codigo, x.razonSocial, x.telefono, x.horaLocal, x.estado + (x.descansoAntes ? " (tras descanso)" : ""), x.bytes, x.caracteres, x.plantilla, x.mensaje, x.enlaceWhatsApp]));
  filas.push(["", "", "", "", "", "", "", "", "", s.mensaje, ""]);
  ws.getRangeByIndexes(0, 0, filas.length, filas[0].length).setValues(filas);
  ws.getRangeByIndexes(0, 0, 1, filas[0].length).getFormat().getFont().setBold(true);
  ws.activate();
}

/* =================== Utilidades =================== */
function salidaError(msg: string): SalidaLote {
  return { ok: false, mensaje: msg, loteNombre: "", ejecutivo: "", correoEjecutivo: "", items: [],
    omitidos: { noContactar: 0, sinConsentimiento: 0, frecuencia: 0, telefonoRepetido: 0, sinTelefono: 0, yaAsignados: 0 },
    resumen: { contactos: 0, hoy: 0, reprogramados: 0, inicio: "", cierre: "", cierreHoy: "", ritmoPorHora: 0, factorAjuste: 1, topeAlcanzado: false } };
}
function normKey(h: string): string { return String(h || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]/g, ""); }
function slug(h: string): string { return String(h || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, ""); }
function claveCodigo(v: string): string { let s = String(v || "").trim().replace(/^'/, "").replace(/\s+/g, "").toUpperCase(); if (/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, ""); return s; }
function verdadero(v: string | undefined): boolean { return /^(si|sí|s|yes|y|true|verdadero|1|x|acepta|autoriza|ok)$/i.test(String(v || "").trim()); }
function normalizarTelefono(v: string): string {
  let d = String(v === null || v === undefined ? "" : v).trim();
  if (/e\+/i.test(d) && !isNaN(Number(d))) d = Number(d).toFixed(0);
  d = d.replace(/\D/g, "");
  if (d.indexOf("00") === 0) d = d.slice(2);
  if (d.length === 9 && d[0] === "9") d = "51" + d;
  return d;
}
function telefonoValido(p: string): boolean { return /^\d{9,15}$/.test(p || ""); }
function mezclar<T>(a: T[]): T[] { for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); const t = a[i]; a[i] = a[j]; a[j] = t; } return a; }
function p2(n: number): string { return String(n).padStart(2, "0"); }
/** "2026-11-18 08:50", "2026-11-18T08:50:00", "2026-11-18" o "18/11/2026" → ms en hora local (tratada como UTC) */
function parseLocal(s?: string): number | null {
  if (!s) return null;
  const t = String(s).trim();
  let m = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?/.exec(t);
  if (m) return Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
  m = /^(\d{2})\/(\d{2})\/(\d{4})(?: (\d{2}):(\d{2}))?/.exec(t);
  if (m) return Date.UTC(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0));
  return null;
}
function claveDia(t: number): string { const d = new Date(t); return `${d.getUTCFullYear()}-${p2(d.getUTCMonth() + 1)}-${p2(d.getUTCDate())}`; }
function fmtFecha(t: number): string { const d = new Date(t); return `${p2(d.getUTCDate())}/${p2(d.getUTCMonth() + 1)}/${d.getUTCFullYear()}`; }
function fmtLocal(t: number): string { const d = new Date(t); return `${fmtFecha(t)} ${p2(d.getUTCHours())}:${p2(d.getUTCMinutes())}:${p2(d.getUTCSeconds())}`; }
function isoConZona(t: number, zona: number): string {
  const d = new Date(t), signo = zona < 0 ? "-" : "+", h = Math.abs(zona);
  return `${claveDia(t)}T${p2(d.getUTCHours())}:${p2(d.getUTCMinutes())}:${p2(d.getUTCSeconds())}${signo}${p2(Math.floor(h))}:${p2(Math.round((h % 1) * 60))}`;
}
