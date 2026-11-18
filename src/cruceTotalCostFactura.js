// ============================================================================
// CRUCE: TOTAL EMPLOYEE COST (nómina / facturación EOR)  vs
//        EO Third parties (Movimiento por cuenta contable de Siigo, cuenta 28)
// ============================================================================
// Entradas (todas con el formato de filas que ya usa la app: rows[r][c] = { v, f }):
//   nominaFiles:   [{ name, rows }]  -> hoja INVOICING de cada archivo de nómina
//   movimientoFiles: [{ name, rows }] -> hoja "Mov-cuenta-contable" de cada Siigo
//
// Cómo cruza:
//   1) Nómina: suma TOTAL EMPLOYEE COST por periodo (mes + tipo: MONTHLY / OFFCYCLE),
//      usando las columnas "Payroll Month" y "Service Type / Invoice Type" de cada fila.
//      Funciona con N archivos: los consolida todos antes de cruzar.
//   2) Siigo: toma las facturas "EO Third parties service ..." (crédito) y saca el
//      periodo del TEXTO de la descripción ("..., June", "..., Offcycle June",
//      "..., April"), NO de la fecha de la factura. Una factura de abril hecha el
//      9 de junio cruza con abril, no con junio.
//   3) Cruza periodo contra periodo. Cruce ok si |nómina - factura| <= tolerancia
//      (10.000 por defecto, opción `tolerance`).
// ============================================================================

export const TOLERANCE_DEFAULT = 10000;

export const CRUCE_COLORS = {
  ok: "92D050",       // verde: cruza dentro de la tolerancia
  missing: "FFFF00",  // amarillo: está de un solo lado
  diff: "FF0000"      // rojo: diferencia mayor a la tolerancia
};

const MONTHS = [
  ["ENERO", "JANUARY", 1], ["FEBRERO", "FEBRUARY", 2], ["MARZO", "MARCH", 3],
  ["ABRIL", "APRIL", 4], ["MAYO", "MAY", 5], ["JUNIO", "JUNE", 6],
  ["JULIO", "JULY", 7], ["AGOSTO", "AUGUST", 8], ["SEPTIEMBRE", "SEPTEMBER", 9],
  ["SETIEMBRE", "SEPTEMBER", 9], ["OCTUBRE", "OCTOBER", 10],
  ["NOVIEMBRE", "NOVEMBER", 11], ["DICIEMBRE", "DECEMBER", 12]
];

// ---------------------------------------------------------------------------
// utilidades
// ---------------------------------------------------------------------------
const norm = (v) =>
  v === null || v === undefined
    ? ""
    : String(v).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/\s+/g, " ").trim();

const val = (cell) => (cell ? cell.v : null);
const round2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const fmt = (n) => Math.round(Number(n) || 0).toLocaleString("es-CO");

function toNumber(v) {
  if (v === null || v === undefined || v === "") return NaN;
  if (typeof v === "number") return Number.isFinite(v) ? v : NaN;
  let t = String(v).trim().replace(/\$/g, "").replace(/\s/g, "");
  if (!t) return NaN;
  if (t.includes(".") && t.includes(",") && t.lastIndexOf(",") > t.lastIndexOf(".")) t = t.replace(/\./g, "").replace(",", ".");
  else if (t.includes(",") && !t.includes(".")) t = t.replace(",", ".");
  else t = t.replace(/,/g, "");
  const n = Number(t);
  return Number.isFinite(n) ? n : NaN;
}

function toDate(v) {
  if (v instanceof Date) return v;
  if (typeof v === "number" && Number.isFinite(v)) return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
  if (typeof v === "string" && v.trim()) {
    const m = /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/.exec(v.trim());
    if (m) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
    const d = new Date(v);
    if (!Number.isNaN(d.getTime())) return d;
  }
  return null;
}

// "June" / "Offcycle June" / "Junio 2026" -> { month, year|null, offcycle }
function parsePeriod(text) {
  const t = norm(text);
  if (!t) return null;
  const words = ` ${t.replace(/[^A-Z0-9]+/g, " ")} `;
  let month = null;
  for (const [es, en, n] of MONTHS) {
    if (words.includes(` ${es} `) || words.includes(` ${en} `)) { month = n; break; }
  }
  if (!month) return null;
  const y = /\b(20\d{2})\b/.exec(t);
  const offcycle = /OFF[\s-]?CYCLE|FUERA DE CICLO/.test(t);
  return { month, year: y ? Number(y[1]) : null, offcycle };
}

const periodKey = (year, month, offcycle) =>
  `${year}-${String(month).padStart(2, "0")}|${offcycle ? "OFFCYCLE" : "MONTHLY"}`;

function periodLabel(key) {
  const [ym, tipo] = key.split("|");
  const [y, m] = ym.split("-").map(Number);
  const name = MONTHS.find((x) => x[2] === m)[0];
  const nice = name.charAt(0) + name.slice(1).toLowerCase();
  return `${nice} ${y}${tipo === "OFFCYCLE" ? " (offcycle)" : ""}`;
}

// ---------------------------------------------------------------------------
// 1) NÓMINA -> total por periodo
// ---------------------------------------------------------------------------
function findNominaHeader(row) {
  if (!Array.isArray(row)) return null;
  const map = {};
  row.forEach((c, i) => {
    const t = norm(val(c));
    if (t && map[t] === undefined) map[t] = i;
  });
  const codeCol = map["EMPLOYEE CODE"] ?? map["CODIGO EMPLEADO"] ?? map["CEDULA"] ?? map["DOCUMENTO"];
  const tecCol = map["TOTAL EMPLOYEE COST"];
  if (codeCol === undefined || tecCol === undefined) return null;
  return {
    codeCol,
    nameCol: map["NAME"] ?? map["NOMBRE"] ?? null,
    monthCol: map["PAYROLL MONTH"] ?? null,
    typeCol: map["SERVICE TYPE / INVOICE TYPE"] ?? null,
    customerCol: map["CUSTOMER NAME"] ?? null,
    tecCol
  };
}

export function readNominaTotals(files, options = {}) {
  const usable = (files || []).filter((f) => f && Array.isArray(f.rows) && f.rows.length);
  const notes = [];
  const byPeriod = new Map(); // key -> { total, employees, files:Set }
  const byFile = new Map();   // `${archivo}|${key}` -> { file, key, total, employees }
  const noMonthFiles = [];
  let noValue = 0;

  // año de referencia: opción > año en el nombre del archivo > año en curso
  const fileYear = (name) => {
    const m = /(20\d{2})/.exec(name || "");
    return m ? Number(m[1]) : null;
  };

  for (const f of usable) {
    let header = null;
    let count = 0;
    const yearRef = options.year || fileYear(f.name) || options.defaultYear || new Date().getFullYear();
    const filePeriodFromName = parsePeriod(f.name);

    for (const row of f.rows) {
      const h = findNominaHeader(row);
      if (h) { header = h; continue; }
      if (!header) continue;

      const code = String(val(row[header.codeCol]) ?? "").trim();
      if (!code) continue; // fila de totales / vacía

      const tecRaw = val(row[header.tecCol]);
      const tec = toNumber(tecRaw);
      if (Number.isNaN(tec)) { noValue += 1; continue; }

      const monthText = header.monthCol !== null ? val(row[header.monthCol]) : null;
      const p = parsePeriod(monthText) || filePeriodFromName;
      if (!p) { if (!noMonthFiles.includes(f.name)) noMonthFiles.push(f.name); continue; }

      const typeText = header.typeCol !== null ? val(row[header.typeCol]) : "";
      const offcycle = p.offcycle || /OFF[\s-]?CYCLE/.test(norm(typeText));
      const key = periodKey(p.year || yearRef, p.month, offcycle);

      if (!byPeriod.has(key)) byPeriod.set(key, { total: 0, employees: 0, files: new Set() });
      const g = byPeriod.get(key);
      g.total = round2(g.total + tec);
      g.employees += 1;
      g.files.add(f.name || "archivo");

      const fk = `${f.name || "archivo"}||${key}`;
      if (!byFile.has(fk)) byFile.set(fk, { file: f.name || "archivo", key, total: 0, employees: 0 });
      const bf = byFile.get(fk);
      bf.total = round2(bf.total + tec);
      bf.employees += 1;
      count += 1;
    }
    notes.push({ type: "info", text: `Nómina "${f.name || "archivo"}": ${count} empleado(s) leídos.` });
  }

  if (noValue > 0)
    notes.push({
      type: "warn",
      text: `${noValue} fila(s) de nómina sin valor en TOTAL EMPLOYEE COST (fórmula sin calcular). Abre y guarda el Excel en Excel antes de subirlo.`
    });
  if (noMonthFiles.length)
    notes.push({ type: "warn", text: `No se pudo detectar el mes en: ${noMonthFiles.join(", ")}. Revisa la columna Payroll Month.` });

  return { byPeriod, byFile, notes };
}

// ---------------------------------------------------------------------------
// 2) SIIGO -> facturas EO Third parties (una por comprobante)
// ---------------------------------------------------------------------------
function findMovHeader(row) {
  if (!Array.isArray(row)) return null;
  const map = {};
  row.forEach((c, i) => {
    const t = norm(val(c));
    if (t && map[t] === undefined) map[t] = i;
  });
  const need = ["COMPROBANTE", "FECHA ELABORACION", "DESCRIPCION", "DEBITO", "CREDITO"];
  if (need.some((k) => map[k] === undefined)) return null;
  return {
    comp: map["COMPROBANTE"], fecha: map["FECHA ELABORACION"], desc: map["DESCRIPCION"],
    deb: map["DEBITO"], cred: map["CREDITO"],
    tercero: map["NOMBRE TERCERO"] ?? null, doc: map["IDENTIFICACION TERCERO"] ?? null
  };
}

const isFactura = (descNorm) => descNorm.includes("EO THIRD PARTIES") || descNorm.includes("INGRESOS RECIBIDOS");

export function readFacturas(files) {
  const usable = (files || []).filter((f) => f && Array.isArray(f.rows) && f.rows.length);
  const notes = [];
  const facturas = []; // { comprobante, fecha, periodKey, valor, descripcion, tercero, file }
  let sinPeriodo = 0;

  for (const f of usable) {
    let header = null;
    let n = 0;
    for (const row of f.rows) {
      const h = findMovHeader(row);
      if (h) { header = h; continue; }
      if (!header) continue;

      const desc = String(val(row[header.desc]) ?? "");
      const dn = norm(desc);
      if (!dn || !isFactura(dn)) continue;

      const fecha = toDate(val(row[header.fecha]));
      const deb = toNumber(val(row[header.deb])) || 0;
      const cred = toNumber(val(row[header.cred])) || 0;
      const valor = round2(cred - deb); // la factura es crédito
      if (valor === 0) continue;

      // Periodo desde la 1ª línea de la descripción: "EO Third parties service X, June"
      const firstLine = desc.split(/\r?\n/)[0];
      const afterComma = firstLine.includes(",") ? firstLine.slice(firstLine.lastIndexOf(",") + 1) : firstLine;
      const p = parsePeriod(afterComma) || parsePeriod(desc);

      if (!p || !fecha) { sinPeriodo += 1; continue; }

      // Año: el del texto, o el más cercano hacia atrás desde la fecha de la factura.
      let year = p.year;
      if (!year) {
        year = fecha.getUTCFullYear();
        if (p.month > fecha.getUTCMonth() + 1) year -= 1;
      }

      facturas.push({
        comprobante: String(val(row[header.comp]) ?? "").trim(),
        fecha,
        periodKey: periodKey(year, p.month, p.offcycle),
        valor,
        descripcion: firstLine.trim(),
        tercero: header.tercero !== null ? String(val(row[header.tercero]) ?? "").trim() : "",
        file: f.name || "archivo"
      });
      n += 1;
    }
    notes.push({ type: "info", text: `Siigo "${f.name || "archivo"}": ${n} factura(s) EO Third parties encontradas.` });
  }

  if (sinPeriodo > 0)
    notes.push({
      type: "warn",
      text: `${sinPeriodo} factura(s) EO Third parties sin mes reconocible en la descripción; no se cruzaron.`
    });

  return { facturas, notes };
}

// ---------------------------------------------------------------------------
// 3) CRUCE
// ---------------------------------------------------------------------------
export function cruzarTotalCostConFactura(nominaFiles, movimientoFiles, options = {}) {
  const tolerance = Number.isFinite(options.tolerance) ? options.tolerance : TOLERANCE_DEFAULT;

  const nom = readNominaTotals(nominaFiles, options);
  const fac = readFacturas(movimientoFiles);
  const notes = [...nom.notes, ...fac.notes];

  if (nom.byPeriod.size === 0 && fac.facturas.length === 0) return null;

  const facByPeriod = new Map();
  for (const f of fac.facturas) {
    if (!facByPeriod.has(f.periodKey)) facByPeriod.set(f.periodKey, { total: 0, items: [] });
    const g = facByPeriod.get(f.periodKey);
    g.total = round2(g.total + f.valor);
    g.items.push(f);
  }

  const keys = [...new Set([...nom.byPeriod.keys(), ...facByPeriod.keys()])].sort();
  const rows = [];

  for (const key of keys) {
    const n = nom.byPeriod.get(key);
    const f = facByPeriod.get(key);
    const nomina = n ? n.total : 0;
    const factura = f ? f.total : 0;
    const diff = round2(nomina - factura);

    let status;
    if (!n || !f) status = "No está en el otro lado";
    else if (Math.abs(diff) <= tolerance) status = "Cruce ok";
    else status = "Diferencias";

    const fill = status === "Cruce ok" ? CRUCE_COLORS.ok : status === "Diferencias" ? CRUCE_COLORS.diff : CRUCE_COLORS.missing;

    rows.push({
      Periodo: periodLabel(key),
      "Total Employee Cost (nómina)": nomina,
      "EO Third parties (Siigo)": factura,
      Diferencia: diff,
      Estado: status,
      Empleados: n ? n.employees : 0,
      "Archivos de nómina": n ? [...n.files].join(", ") : "",
      Facturas: f ? f.items.map((i) => `${i.comprobante} (${fmt(i.valor)})`).join("; ") : "",
      _fill: fill
    });
  }

  // ---- Cruce archivo de nómina <-> factura(s) (sirve cuando faltan archivos) ----
  // Cada archivo de nómina busca una factura del mismo periodo que cuadre dentro de la
  // tolerancia; si ninguna sola cuadra, prueba con la suma de dos facturas.
  const used = new Set();
  const detail = [];
  for (const bf of nom.byFile.values()) {
    const f = facByPeriod.get(bf.key);
    const cand = f ? f.items.filter((i) => !used.has(i)) : [];
    let best = null;
    for (const i of cand) {
      const d = Math.abs(bf.total - i.valor);
      if (d <= tolerance && (!best || d < best.d)) best = { d, items: [i] };
    }
    if (!best) {
      for (let a = 0; a < cand.length; a++)
        for (let b = a + 1; b < cand.length; b++) {
          const d = Math.abs(bf.total - (cand[a].valor + cand[b].valor));
          if (d <= tolerance && (!best || d < best.d)) best = { d, items: [cand[a], cand[b]] };
        }
    }
    // Sin cruce dentro de la tolerancia: si hay una factura muy cercana (<= 5%), es una diferencia (rojo).
    let near = null;
    if (!best) {
      for (const i of cand) {
        const d = Math.abs(bf.total - i.valor);
        if (d <= Math.abs(bf.total) * 0.05 && (!near || d < near.d)) near = { d, items: [i] };
      }
    }
    const pick = best || near;
    if (pick) pick.items.forEach((i) => used.add(i));
    const factura = pick ? round2(pick.items.reduce((acc, i) => acc + i.valor, 0)) : 0;
    detail.push({
      Periodo: periodLabel(bf.key),
      "Archivo de nómina": bf.file,
      "Total Employee Cost (nómina)": bf.total,
      "EO Third parties (Siigo)": factura,
      Diferencia: round2(bf.total - factura),
      Estado: best ? "Cruce ok" : near ? "Diferencias" : "No está en el otro lado",
      Facturas: pick ? pick.items.map((i) => `${i.comprobante} (${fmt(i.valor)})`).join(" + ") : "",
      _fill: best ? CRUCE_COLORS.ok : near ? CRUCE_COLORS.diff : CRUCE_COLORS.missing
    });
  }
  for (const [key, f] of facByPeriod) {
    for (const i of f.items) {
      if (used.has(i)) continue;
      detail.push({
        Periodo: periodLabel(key),
        "Archivo de nómina": "",
        "Total Employee Cost (nómina)": 0,
        "EO Third parties (Siigo)": i.valor,
        Diferencia: round2(-i.valor),
        Estado: "No está en el otro lado",
        Facturas: `${i.comprobante} (${fmt(i.valor)})`,
        _fill: CRUCE_COLORS.missing
      });
    }
  }
  detail.sort((a, b) => a.Periodo.localeCompare(b.Periodo));

  const okFiles = detail.filter((d) => d.Estado === "Cruce ok").length;
  notes.unshift({
    type: "info",
    text: `Cruce por archivo/factura: ${okFiles} cruce(s) ok de ${detail.length} línea(s). Lo amarillo es una nómina sin factura o una factura sin nómina (puede faltar subir un archivo).`
  });

  const ok = rows.filter((r) => r.Estado === "Cruce ok").length;
  notes.unshift({
    type: ok === rows.length ? "info" : "warn",
    text: `Cruce con tolerancia de ${fmt(tolerance)}: ${ok} de ${rows.length} periodo(s) cruzan.`
  });

  return { rows, detail, notes, tolerance };
}