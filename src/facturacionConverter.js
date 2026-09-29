// ============================================================================
// facturacionConverter.js
// Facturación EOR (formato ancho, una fila por empleado) -> formato largo
// ============================================================================
//
// Entrada (igual que siigoConverter.js): filas de Excel como
//   [ [ { v: valor, f: "RRGGBB" }, ... ], ... ]
//
// Salida (mismo esquema que convertNominaRows / convertMovimientoRows):
//   { records, notes }
//   records[i] = {
//     "Mes elaboración", Concepto, Empleado, "Valor Concepto", "Valor Totales", _fill
//   }
//
// Estructura que entiende (hoja INVOICING del archivo de facturación):
//   - Encabezado con "EMPLOYEE CODE", "NAME" y "PAYROLL MONTH" (no depende de
//     posiciones fijas). PAYROLL MONTH + SERVICE TYPE / CUSTOMER NAME es lo que
//     distingue la facturación de una nómina normal.
//   - Pueden existir VARIOS bloques con su propio encabezado en la misma hoja
//     (p. ej. "Monthly Payroll" y luego "Monthly Adjustment").
//   - Mes = columna "Payroll Month" (texto en inglés/español, o fecha).
//     El año se toma del nombre del archivo ("... Agosto 2026.xlsx") o de
//     options.year. Los meses posteriores al mes de facturación se asumen del
//     año anterior (ajustes de meses pasados).
//   - Conceptos = columnas entre las columnas de datos generales
//     (Country, Customer, Payroll Month, ER SS rate %...) y "TOTAL EMPLOYEE COST".
//     Se excluyen los subtotales "TOTAL", "PAYMENTS", "TOTAL COST AND LEGAL
//     BENEFITS" y todo lo que está después de TOTAL EMPLOYEE COST
//     (FEE, BANKING TAX, IVA, USD...).
// ============================================================================

export const SIN_EMPLEADO_FACT = "(sin asignar)";

const TOTAL_LABEL = "TOTAL EMPLOYEE COST";

const CODE_HEADERS = new Set([
  "EMPLOYEE CODE", "EMPLOYEE ID", "CODE", "CODIGO", "CODIGO EMPLEADO",
  "COD EMPLEADO", "ID EMPLEADO", "CEDULA", "DOCUMENTO", "IDENTIFICACION"
]);

const NAME_HEADERS = new Set([
  "NAME", "EMPLOYEE NAME", "FULL NAME", "NOMBRE", "NOMBRE EMPLEADO",
  "NOMBRE DEL EMPLEADO", "EMPLEADO", "NOMBRE COMPLETO", "NOMBRES Y APELLIDOS"
]);

// Columnas de datos generales que NO son conceptos de costo.
const META_HEADERS = new Set([
  "EE RF WID", "ONBOARDING DATE", "OFFBOARDING DATE", "COUNTRY",
  "CUSTOMER NAME", "CUSTOMER ID", "PAYROLL MONTH",
  "SERVICE TYPE / INVOICE TYPE", "ER SS RATE %"
]);

const MONTHS = [
  ["ENERO", 1], ["JANUARY", 1],
  ["FEBRERO", 2], ["FEBRUARY", 2],
  ["MARZO", 3], ["MARCH", 3],
  ["ABRIL", 4], ["APRIL", 4],
  ["MAYO", 5], ["MAY", 5],
  ["JUNIO", 6], ["JUNE", 6],
  ["JULIO", 7], ["JULY", 7],
  ["AGOSTO", 8], ["AUGUST", 8],
  ["SEPTIEMBRE", 9], ["SETIEMBRE", 9], ["SEPTEMBER", 9],
  ["OCTUBRE", 10], ["OCTOBER", 10],
  ["NOVIEMBRE", 11], ["NOVEMBER", 11],
  ["DICIEMBRE", 12], ["DECEMBER", 12]
];

// ---------------------------------------------------------------------------
// Utilidades
// ---------------------------------------------------------------------------

function norm(v) {
  if (v === null || v === undefined) return "";
  return String(v)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}

function clean(v) {
  return String(v ?? "").replace(/\s+/g, " ").trim();
}

function round2(n) {
  return Math.round((Number(n) || 0) * 100) / 100;
}

function fmt(n) {
  return Math.round(Number(n) || 0).toLocaleString("es-CO");
}

function cellValue(cell) {
  return cell ? cell.v : null;
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;

  let t = String(value).replace(/\$/g, "").replace(/\s/g, "");
  if (!t) return 0;

  if (t.includes(".") && t.includes(",")) {
    t = t.lastIndexOf(",") > t.lastIndexOf(".")
      ? t.replace(/\./g, "").replace(",", ".")
      : t.replace(/,/g, "");
  } else if (t.includes(",")) {
    t = t.replace(",", ".");
  }

  const n = Number(t);
  return Number.isFinite(n) ? n : 0;
}

// Mes a partir de un texto corto de celda ("August", "julio", "Jul 2026"...).
function monthFromText(text) {
  const t = norm(text);
  if (!t || t.length > 40) return null;
  for (const [name, num] of MONTHS) {
    if (t.includes(name)) return num;
  }
  return null;
}

// Mes a partir del nombre del archivo (largo, con códigos y guiones bajos).
// Si aparecen varios meses toma el que está más a la derecha.
function monthFromFileName(fileName) {
  const t = norm(fileName).replace(/[_\-.]+/g, " ");
  let best = null;
  let bestPos = -1;
  for (const [name, num] of MONTHS) {
    const re = new RegExp(`(^|[^A-Z])${name}(?![A-Z])`, "g");
    let m;
    while ((m = re.exec(t)) !== null) {
      const pos = m.index + m[1].length;
      if (pos > bestPos) {
        bestPos = pos;
        best = num;
      }
    }
  }
  return best;
}

// Celda de Payroll Month: Date, serial de Excel (número) o texto.
function monthFromCell(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.getUTCMonth() + 1;
  }
  if (typeof value === "number") {
    if (value > 20000 && value < 80000) {
      // serial de Excel (días desde 1899-12-30)
      const d = new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 86400000);
      return d.getUTCMonth() + 1;
    }
    if (Number.isInteger(value) && value >= 1 && value <= 12) return value;
    return null;
  }
  return monthFromText(value);
}

// Último año de 4 dígitos (20xx) que no esté pegado a otros dígitos, así
// "C2045_..._Agosto_2026" devuelve 2026 y no 2045.
function yearFromText(text) {
  const matches = [...String(text ?? "").matchAll(/(?<!\d)(20\d{2})(?!\d)/g)];
  return matches.length ? Number(matches[matches.length - 1][1]) : null;
}

function looksLikeEmployeeName(value) {
  const name = clean(value);
  if (name.length < 2) return false;
  const up = norm(name);
  // Filas de totales (prefijo) o repetición del encabezado (texto exacto).
  // Ojo: "NAME PARADA SOFIA" es un empleado real, por eso NAME/NOMBRE van exactos.
  if (/^(TOTAL|SUBTOTAL|GRAN TOTAL)\b/.test(up)) return false;
  if (/^(NAME|NOMBRE|EMPLOYEE|EMPLEADO)$/.test(up)) return false;
  if (/^[\d.,\s\-+$()%\/]+$/.test(name)) return false;
  return true;
}

// ---------------------------------------------------------------------------
// Encabezado
// ---------------------------------------------------------------------------

function readHeader(row, opts) {
  if (!Array.isArray(row)) return null;

  const extraMeta = opts.extraMetaHeaders;
  let codeCol = -1, nameCol = -1;
  let monthCol = -1, customerCol = -1, serviceCol = -1;
  let tecCol = null, paymentsCol = null, bankingCol = null;
  let metaEnd = -1;

  for (let c = 0; c < row.length; c++) {
    const t = norm(cellValue(row[c]));
    if (!t) continue;

    if (codeCol < 0 && CODE_HEADERS.has(t)) { codeCol = c; continue; }
    if (nameCol < 0 && NAME_HEADERS.has(t)) { nameCol = c; continue; }

    if (t === "PAYROLL MONTH" && monthCol < 0) monthCol = c;
    if (t === "CUSTOMER NAME" && customerCol < 0) customerCol = c;
    if (t === "SERVICE TYPE / INVOICE TYPE" && serviceCol < 0) serviceCol = c;
    if (t === TOTAL_LABEL && tecCol === null) tecCol = c;
    if (t === "PAYMENTS" && paymentsCol === null) paymentsCol = c;
    if (t === "BANKING TAX" && bankingCol === null) bankingCol = c;

    if (
      nameCol >= 0 && c > nameCol &&
      (META_HEADERS.has(t) || t.startsWith("EE STATUS") || extraMeta.has(t))
    ) {
      metaEnd = Math.max(metaEnd, c);
    }
  }

  // Sin PAYROLL MONTH (y Service Type o Customer Name) no es facturación EOR:
  // así no se "roba" las nóminas normales que también tienen EMPLOYEE CODE y NAME.
  if (codeCol < 0 || nameCol < 0 || monthCol < 0) return null;
  if (serviceCol < 0 && customerCol < 0) return null;

  const start = Math.max(nameCol, metaEnd) + 1;
  const end = tecCol !== null ? tecCol : row.length;
  const concepts = [];

  for (let c = start; c < end; c++) {
    const v = cellValue(row[c]);
    if (typeof v !== "string" || !clean(v)) continue;

    const t = norm(v);
    if (t.startsWith("TOTAL") || t.startsWith("SUBTOTAL")) continue; // subtotales de sección
    if (t === "PAYMENTS") continue;
    if (META_HEADERS.has(t) || t.startsWith("EE STATUS") || extraMeta.has(t)) continue;

    concepts.push({ idx: c, label: clean(v) });
  }

  if (concepts.length === 0) return null;

  return {
    codeCol, nameCol, monthCol, customerCol, serviceCol,
    tecCol, paymentsCol, bankingCol, concepts
  };
}

// ---------------------------------------------------------------------------
// Conversión de UN archivo
// ---------------------------------------------------------------------------

export function convertFacturacionRows(rows, options = {}) {
  if (!Array.isArray(rows) || rows.length === 0) return null;

  const opts = {
    extraMetaHeaders: new Set((options.extraMetaHeaders || []).map(norm))
  };

  const includeExtras = options.includeExtras === true;
  const fileName = options.fileName || "";

  // ---- Pasada 1: encontrar empleados -------------------------------------
  const entries = [];
  let header = null;

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];

    const h = readHeader(row, opts);
    if (h) { header = h; continue; }
    if (!header) continue;

    const name = clean(cellValue(row[header.nameCol]));
    const code = clean(cellValue(row[header.codeCol]));

    if (!looksLikeEmployeeName(name)) continue;

    const hasNumbers = header.concepts.some(
      (c) => toNumber(cellValue(row[c.idx])) !== 0
    );
    if (!code && !hasNumbers) continue;

    entries.push({
      header,
      row,
      rowNum: r + 1,
      name,
      code,
      month: header.monthCol >= 0 ? monthFromCell(cellValue(row[header.monthCol])) : null,
      customer: header.customerCol >= 0 ? clean(cellValue(row[header.customerCol])) : "",
      service: header.serviceCol >= 0 ? clean(cellValue(row[header.serviceCol])) : ""
    });
  }

  if (entries.length === 0) return null;

  // ---- Mes y año de facturación ------------------------------------------
  // Mes de facturación = el más frecuente entre las filas "Payroll" (o entre todas).
  const countMonths = (list) => {
    const m = new Map();
    for (const e of list) if (e.month) m.set(e.month, (m.get(e.month) || 0) + 1);
    return [...m.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? null;
  };

  const payrollRows = entries.filter((e) => /PAYROLL/.test(norm(e.service)));
  const invoiceMonth =
    countMonths(payrollRows.length ? payrollRows : entries) ||
    monthFromFileName(fileName);

  let invoiceYear = options.year || yearFromText(fileName);
  let yearAssumed = false;
  if (!invoiceYear) {
    invoiceYear = options.defaultYear || new Date().getFullYear();
    yearAssumed = true;
  }

  // ---- Pasada 2: formato largo -------------------------------------------
  const records = [];
  const months = new Set();
  const seenEmpMonth = new Map();

  let paymentMismatch = 0;
  const totalMismatch = [];
  let zeroRows = 0;
  let noMonthRows = 0;
  let adjustmentRows = 0;
  let duplicates = 0;

  for (const e of entries) {
    const { concepts, tecCol, paymentsCol, bankingCol } = e.header;

    let month = e.month;
    if (!month) {
      month = invoiceMonth;
      noMonthRows += 1;
    }
    if (!month) continue;

    // Un mes posterior al de facturación es del año anterior (ajustes).
    const year = invoiceMonth && month > invoiceMonth ? invoiceYear - 1 : invoiceYear;
    const monthDate = new Date(Date.UTC(year, month - 1, 1));
    const monthId = `${year}-${String(month).padStart(2, "0")}`;

    if (/ADJUST/.test(norm(e.service))) adjustmentRows += 1;

    let total = 0;
    let payrollSum = 0;
    const out = [];

    for (const c of concepts) {
      const value = round2(toNumber(cellValue(e.row[c.idx])));

      if (paymentsCol !== null && c.idx < paymentsCol) {
        payrollSum = round2(payrollSum + value);
      }

      if (value === 0) continue; // los ceros no salen al formato largo

      total = round2(total + value);
      out.push({ label: c.label, value });
    }

    // Control: PAYMENTS de la hoja vs suma de conceptos de nómina.
    if (paymentsCol !== null) {
      const paid = round2(toNumber(cellValue(e.row[paymentsCol])));
      if (Math.abs(paid - payrollSum) > 1) paymentMismatch += 1;
    }

    // Control: TOTAL EMPLOYEE COST de la hoja (sin BANKING TAX, que no es concepto).
    if (tecCol !== null) {
      const tec = round2(toNumber(cellValue(e.row[tecCol])));
      const bank = bankingCol !== null
        ? round2(toNumber(cellValue(e.row[bankingCol])))
        : 0;
      const expected = round2(tec - bank);
      if (tec !== 0 && Math.abs(total - expected) > 1) {
        totalMismatch.push({ month: monthId, name: e.name, diff: round2(total - expected) });
      }
    }

    if (out.length === 0) {
      zeroRows += 1;
      continue;
    }

    months.add(monthId);

    const empMonthKey = `${monthId}|${e.code || e.name}|${norm(e.service)}`;
    seenEmpMonth.set(empMonthKey, (seenEmpMonth.get(empMonthKey) || 0) + 1);
    if (seenEmpMonth.get(empMonthKey) === 2) duplicates += 1;

    const extras = includeExtras
      ? { Documento: e.code, Cliente: e.customer, "Tipo de servicio": e.service }
      : {};

    for (const o of out) {
      records.push({
        "Mes elaboración": monthDate,
        Concepto: o.label,
        Empleado: e.name || SIN_EMPLEADO_FACT,
        "Valor Concepto": o.value,
        "Valor Totales": 0,
        _fill: null,
        ...extras
      });
    }

    records.push({
      "Mes elaboración": monthDate,
      Concepto: TOTAL_LABEL,
      Empleado: e.name || SIN_EMPLEADO_FACT,
      "Valor Concepto": 0,
      "Valor Totales": total,
      _fill: null,
      ...extras
    });
  }

  if (records.length === 0) return null;

  // ---- Avisos ------------------------------------------------------------
  const notes = [];
  const sortedMonths = [...months].sort();

  notes.push({
    type: "info",
    text:
      `Facturación: se leyeron ${entries.length} fila(s) de empleado en ${sortedMonths.length} mes(es)` +
      (sortedMonths.length
        ? ` (${sortedMonths[0]} a ${sortedMonths[sortedMonths.length - 1]})`
        : "") +
      (adjustmentRows ? `, de las cuales ${adjustmentRows} son ajustes de meses anteriores` : "") +
      "."
  });

  notes.push({
    type: "info",
    text:
      "El TOTAL EMPLOYEE COST del resultado es la suma de los conceptos (nómina + seguridad social + prestaciones + otros costos del empleador). " +
      "No incluye FEE, BANKING TAX, IVA ni columnas en USD."
  });

  if (yearAssumed) {
    notes.push({
      type: "warn",
      text: `No se encontró el año en el nombre del archivo ni en options.year; se usó ${invoiceYear}. Revisa la columna Mes elaboración.`
    });
  }

  if (noMonthRows > 0) {
    notes.push({
      type: "warn",
      text: `${noMonthRows} fila(s) no tenían un Payroll Month reconocible; se les asignó el mes de facturación.`
    });
  }

  if (zeroRows > 0) {
    notes.push({
      type: "info",
      text: `${zeroRows} fila(s) con todos los conceptos en 0 no generaron registros.`
    });
  }

  if (duplicates > 0) {
    notes.push({
      type: "info",
      text: `${duplicates} empleado(s) aparecen más de una vez en el mismo mes y tipo de servicio; sus filas se dejaron separadas.`
    });
  }

  if (paymentMismatch > 0) {
    notes.push({
      type: "warn",
      text: `${paymentMismatch} fila(s) donde los conceptos de nómina no suman el PAYMENTS de la hoja.`
    });
  }

  if (totalMismatch.length > 0) {
    const ex = totalMismatch
      .slice(0, 4)
      .map((m) => `${m.month} ${m.name} (${m.diff > 0 ? "+" : ""}${fmt(m.diff)})`)
      .join("; ");
    notes.push({
      type: "warn",
      text: `${totalMismatch.length} fila(s) donde el total calculado no coincide con TOTAL EMPLOYEE COST - BANKING TAX de la hoja. Ejemplos: ${ex}.`
    });
  }

  return { records, notes };
}

// ---------------------------------------------------------------------------
// Conversión de VARIOS archivos
// ---------------------------------------------------------------------------
// files = [{ name, rows }]  (mismo formato que convertNominaFiles)

export function convertFacturacionFiles(files, options = {}) {
  const usable = (files || []).filter(
    (f) => f && Array.isArray(f.rows) && f.rows.length > 0
  );
  if (usable.length === 0) return null;

  const records = [];
  const notes = [];

  for (const [i, f] of usable.entries()) {
    const name = f.name || `archivo ${i + 1}`;
    const res = convertFacturacionRows(f.rows, { ...options, fileName: name });

    if (!res) {
      notes.push({
        type: "warn",
        text: `"${name}": no se encontró un encabezado con EMPLOYEE CODE, NAME y Payroll Month, o no hay valores.`
      });
      continue;
    }

    records.push(...res.records);
    for (const n of res.notes) {
      notes.push(usable.length > 1 ? { ...n, text: `[${name}] ${n.text}` } : n);
    }
  }

  if (records.length === 0) return { records, notes };

  records.sort((a, b) => a["Mes elaboración"] - b["Mes elaboración"]);

  notes.unshift({
    type: "info",
    text: `Se consolidaron ${usable.length} archivo(s) de facturación en un solo resultado.`
  });

  return { records, notes };
}