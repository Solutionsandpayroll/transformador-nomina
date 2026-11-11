// ============================================================================
// CONVERTIDOR: Nómina (formato ancho, un bloque por mes) y Movimiento CC de
// Siigo -> formato largo
// ============================================================================
// Salida de ambos conversores:
//   { records, notes }
//   records[i] = {
//     "Mes elaboración", Concepto, Empleado, Documento,
//     "Valor Concepto", "Valor Totales", _fill
//   }
// Documento = cédula del empleado (EMPLOYEE CODE en la nómina, "Identificación
// tercero" en el Siigo cuando el tercero es el empleado). Se usa para cruzar.
// ============================================================================

// ============================================================================
// CONFIGURACIÓN
// ============================================================================

export const STATUS_COLORS = {
  "92D050": "Cruce ok",
  "00B0F0": "Cruce ok (azul)",
  "FFFF00": "No está en el otro lado",
  "FF0000": "Diferencias",
  "00FF00": "Cruza entre meses",
  "7030A0": "Débito - Crédito se anulan",
  "00FFFF": "Otro (sin leyenda)"
};

const STRUCTURAL_FILLS = new Set(["DCFAFA", "D6E3BC", "1F4763", "FFFFFF", "000000"]);

const NON_CONCEPT_EXACT = new Set([
  "PAYMENTS", "EE RF WID", "ONBOARDING DATE", "OFFBOARDING DATE", "COUNTRY",
  "PAYROLL MONTH", "SERVICE TYPE / INVOICE TYPE", "ER SS RATE %", "EMPLOYEE CODE",
  "CODE", "CODIGO", "CODIGO EMPLEADO", "COD EMPLEADO", "ID EMPLEADO", "CEDULA",
  "DOCUMENTO", "IDENTIFICACION", "NAME", "NOMBRE", "NOMBRE EMPLEADO", "EMPLEADO",
  "NOMBRE COMPLETO", "NOMBRE DEL EMPLEADO", "NOMBRES Y APELLIDOS"
]);

const NOT_AN_EMPLOYEE =
  /^(TOTAL|NOMINA|NÓMINA|CONTABILIDAD|NOVEDAD|NOVADADES|DIFERENCIA|CRUCE|NO ESTA|NO ESTÁ|SUBTOTAL|TOTAL GENERAL|GRAN TOTAL|RESUMEN|SUMMARY|CONTROL|OBSERVACION|OBSERVACIONES)/;

const SUMMARY_LABEL =
  /^(CRUCE|NO ESTA EN EL OTRO|NO ESTÁ EN EL OTRO|DIFERENCIAS|CRUZA ENTRE|DEBITO - CREDITO|DÉBITO - CRÉDITO)/;

const FILE_BREAK = "__FILE_BREAK__";

const DEFAULT_CODE_HEADER_ALIASES = [
  "EMPLOYEE CODE", "EMPLOYEE ID", "EMPLOYEE NUMBER", "CODE", "CODIGO", "CÓDIGO",
  "CODIGO EMPLEADO", "CÓDIGO EMPLEADO", "COD EMPLEADO", "COD. EMPLEADO", "ID EMPLEADO",
  "ID DEL EMPLEADO", "CEDULA", "CÉDULA", "DOCUMENTO", "DOCUMENTO IDENTIDAD",
  "DOCUMENTO DE IDENTIDAD", "NO. IDENTIFICACION", "NO IDENTIFICACION",
  "NO. IDENTIFICACIÓN", "IDENTIFICACION", "IDENTIFICACIÓN"
];

const DEFAULT_NAME_HEADER_ALIASES = [
  "NAME", "EMPLOYEE NAME", "NOMBRE", "NOMBRE EMPLEADO", "NOMBRE DEL EMPLEADO", "EMPLEADO",
  "NOMBRE COMPLETO", "NOMBRES Y APELLIDOS", "FULL NAME"
];

const MONTHS = [
  ["ENERO", 1], ["FEBRERO", 2], ["MARZO", 3], ["ABRIL", 4], ["MAYO", 5], ["JUNIO", 6],
  ["JULIO", 7], ["AGOSTO", 8], ["SEPTIEMBRE", 9], ["SETIEMBRE", 9], ["OCTUBRE", 10],
  ["NOVIEMBRE", 11], ["DICIEMBRE", 12]
];

const SIN_EMPLEADO = "(sin asignar)";

// ============================================================================
// UTILIDADES GENERALES
// ============================================================================

function norm(value) {
  if (value === null || value === undefined) return "";
  return String(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}

function clean(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;

  let text = String(value).trim().replace(/\$/g, "").replace(/\s/g, "");
  if (!text) return 0;

  if (text.includes(".") && text.includes(",") && text.lastIndexOf(",") > text.lastIndexOf(".")) {
    text = text.replace(/\./g, "").replace(",", "."); // 1.234.567,89
  } else if (text.includes(",") && !text.includes(".")) {
    text = text.replace(",", "."); // 123456,78
  } else {
    text = text.replace(/,/g, ""); // 1,234,567
  }

  const n = Number(text);
  return Number.isFinite(n) ? n : 0;
}

function round2(n) {
  return Math.round((Number(n) || 0) * 100) / 100;
}

function formatMoney(n) {
  return Math.round(Number(n) || 0).toLocaleString("es-CO");
}

function cellValue(cell) {
  return cell ? cell.v : null;
}

function cellFill(cell) {
  return cell ? cell.f || null : null;
}

function isHexFill(value) {
  return typeof value === "string" && /^[0-9A-F]{6}$/i.test(value);
}

function statusFill(fill) {
  if (!fill) return null;
  const normalized = String(fill).replace("#", "").toUpperCase();
  return STATUS_COLORS[normalized] ? normalized : null;
}

// ============================================================================
// MESES
// ============================================================================

function parseMonthLabel(text) {
  if (text === null || text === undefined) return null;
  const t = norm(text);
  if (!t || t.length > 50) return null;

  for (const [name, num] of MONTHS) {
    if (t.includes(name)) {
      const y = /(20\d{2})/.exec(t);
      return { month: num, year: y ? Number(y[1]) : null };
    }
  }
  return null;
}

// Busca un mes en toda la fila (primero la columna A).
function findMonthLabelInRow(row) {
  if (!Array.isArray(row)) return null;

  const first = parseMonthLabel(cellValue(row[0]));
  if (first) return first;

  for (let c = 1; c < row.length; c++) {
    const value = cellValue(row[c]);
    if (typeof value !== "string" || !value.trim()) continue;
    const parsed = parseMonthLabel(value);
    if (parsed) return parsed;
  }
  return null;
}

// ============================================================================
// DETECCIÓN DE ENCABEZADOS DE NÓMINA
// ============================================================================

function isCodeHeader(text, aliases) {
  const t = norm(text);
  if (!t) return false;
  if (aliases.includes(t)) return true;

  return (
    /^(EMPLOYEE\s+)?CODE$/.test(t) ||
    /^EMPLOYEE\s+(ID|NUMBER|NO)$/.test(t) ||
    /^COD(IGO)?(\s+DEL?)?\s*EMPLEADO$/.test(t) ||
    /^ID(\s+DEL?)?\s*EMPLEADO$/.test(t) ||
    /^NO\.?\s*(DE\s+)?IDENTIFICACION$/.test(t) ||
    /^DOCUMENTO(\s+DE)?\s+IDENTIDAD$/.test(t)
  );
}

function isNameHeader(text, aliases) {
  const t = norm(text);
  if (!t) return false;
  if (aliases.includes(t)) return true;

  return (
    /^EMPLOYEE\s+NAME$/.test(t) ||
    /^FULL\s+NAME$/.test(t) ||
    /^NOMBRE.*EMPLEADO$/.test(t) ||
    /^NOMBRE.*COMPLETO$/.test(t) ||
    /^NOMBRES?\s+Y\s+APELLIDOS$/.test(t)
  );
}

function looksLikeStructuralHeader(text) {
  const t = norm(text);
  if (!t) return true;
  if (NON_CONCEPT_EXACT.has(t)) return true;
  if (t.startsWith("EE STATUS") || t.startsWith("TOTAL") || t.startsWith("SUBTOTAL")) return true;
  if (t.startsWith("FEE") || t.includes("EXCHANGE") || t.includes("USD") || t.includes("RATE %")) return true;
  return false;
}

// Detecta la fila de encabezado de un bloque.
function readHeader(row, codeAliases, nameAliases) {
  let codeCol = -1;
  let nameCol = -1;

  if (!Array.isArray(row)) return null;

  for (let c = 0; c < row.length; c++) {
    const t = norm(cellValue(row[c]));
    if (!t) continue;

    if (codeCol < 0 && isCodeHeader(t, codeAliases)) {
      codeCol = c;
      continue;
    }
    if (nameCol < 0 && isNameHeader(t, nameAliases)) {
      nameCol = c;
    }
  }

  if (codeCol < 0 || nameCol < 0) return null;

  const leaves = [];
  for (let c = nameCol + 1; c < row.length; c++) {
    const v = cellValue(row[c]);
    if (typeof v !== "string" || !clean(v)) continue;
    leaves.push({ idx: c, label: clean(v), upper: norm(v) });
  }

  const tec = leaves.find((l) => l.upper === "TOTAL EMPLOYEE COST");
  const payments = leaves.find((l) => l.upper === "PAYMENTS");
  const limit = tec ? tec.idx : Infinity;

  let concepts = leaves.filter((l) => l.idx < limit && !looksLikeStructuralHeader(l.upper));

  // Si existe SALARY junto con INTEGRATED / ORDINARY SALARY, se evita duplicar SALARY.
  const hasBreakdown = concepts.some(
    (l) => l.upper === "INTEGRATED SALARY" || l.upper === "ORDINARY SALARY"
  );
  if (hasBreakdown) concepts = concepts.filter((l) => l.upper !== "SALARY");

  if (concepts.length === 0) return null;

  return {
    codeCol,
    nameCol,
    concepts,
    tecCol: tec ? tec.idx : null,
    paymentsCol: payments ? payments.idx : null
  };
}

// ============================================================================
// DETECCIÓN DE EMPLEADOS
// ============================================================================

function looksLikeEmployeeName(value) {
  const name = clean(value);
  if (!name) return false;

  const upper = norm(name);

  if (upper.startsWith("#")) return false;
  if (NOT_AN_EMPLOYEE.test(upper)) return false;
  if (/^(NAME|NOMBRE|EMPLOYEE|EMPLEADO)$/i.test(name)) return false;
  if (/^\d{1,4}[\/-]\d{1,2}[\/-]\d{1,4}$/.test(name)) return false; // fechas
  if (/^[\d.,\s\-+$()%]+$/.test(name)) return false; // números
  if (name.length < 2) return false;

  return true;
}

function rowHasNumericConcept(row, concepts) {
  return concepts.some((c) => {
    const value = cellValue(row[c.idx]);
    if (typeof value === "number") return Number.isFinite(value);
    if (typeof value === "string" && value.trim()) return toNumber(value) !== 0;
    return false;
  });
}

// ============================================================================
// NÓMINA -> FORMATO LARGO
// ============================================================================

export function convertNominaRows(rows, options = {}) {
  if (!Array.isArray(rows) || rows.length === 0) return null;

  const unifyNames = options.unifyNamesByCode !== false;
  const defaultYear = options.defaultYear || new Date().getFullYear();

  const codeAliases = (
    options.codeHeaderAliases || [...DEFAULT_CODE_HEADER_ALIASES, ...(options.extraCodeHeaderAliases || [])]
  ).map(norm);

  const nameAliases = (
    options.nameHeaderAliases || [...DEFAULT_NAME_HEADER_ALIASES, ...(options.extraNameHeaderAliases || [])]
  ).map(norm);

  const labelEvents = [];
  const entries = [];
  const blockInfo = new Map();

  let header = null;
  let blockId = 0;
  let curLabel = -1;

  // --------------------------------------------------------------------------
  // PASADA 1: encontrar bloques y empleados
  // --------------------------------------------------------------------------

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];

    // Separador entre archivos.
    if (norm(cellValue(row[0])) === FILE_BREAK) {
      header = null;
      curLabel = -1;
      continue;
    }

    // Resumen final.
    if (SUMMARY_LABEL.test(norm(cellValue(row[0])))) {
      header = null;
      continue;
    }

    // Detectar mes.
    const label = findMonthLabelInRow(row);
    if (label) {
      labelEvents.push({ ...label });
      curLabel = labelEvents.length - 1;
    }

    // Detectar encabezado.
    const h = readHeader(row, codeAliases, nameAliases);
    if (h) {
      header = h;
      blockId += 1;
      blockInfo.set(blockId, { employees: 0, checkFill: null });
      continue;
    }

    if (!header) continue;

    const name = clean(cellValue(row[header.nameCol]));
    const code = clean(cellValue(row[header.codeCol]));
    const hasNumbers = rowHasNumericConcept(row, header.concepts);
    const validName = looksLikeEmployeeName(name);

    // Fila de empleado: nombre válido + (código o algún valor en los conceptos).
    const isEmployee = validName && (Boolean(code) || hasNumbers);

    if (!isEmployee) {
      const info = blockInfo.get(blockId);

      if (info && header.tecCol !== null && !info.checkFill) {
        const totalCell = row[header.tecCol];
        if (typeof cellValue(totalCell) === "number" || toNumber(cellValue(totalCell)) !== 0) {
          info.checkFill = statusFill(cellFill(totalCell));
        }
      }
      continue;
    }

    const info = blockInfo.get(blockId);
    if (info) info.employees += 1;

    entries.push({ blockId, labelIdx: curLabel, header, rowNum: r + 1, row, code, name });
  }

  if (entries.length === 0) return null;

  // --------------------------------------------------------------------------
  // DEDUCIR AÑOS DE LOS MESES
  // --------------------------------------------------------------------------

  // Pasada hacia atrás.
  let refYear = null;
  let refMonth = null;

  for (let i = labelEvents.length - 1; i >= 0; i--) {
    const ev = labelEvents[i];
    if (ev.year) {
      refYear = ev.year;
      refMonth = ev.month;
    } else if (refYear !== null) {
      ev.year = ev.month <= refMonth ? refYear : refYear - 1;
      ev.inferred = true;
      refYear = ev.year;
      refMonth = ev.month;
    }
  }

  // Pasada hacia adelante.
  refYear = null;
  refMonth = null;

  for (const ev of labelEvents) {
    if (ev.year) {
      refYear = ev.year;
      refMonth = ev.month;
    } else if (refYear !== null) {
      ev.year = ev.month >= refMonth ? refYear : refYear + 1;
      ev.inferred = true;
      refYear = ev.year;
      refMonth = ev.month;
    }
  }

  let noYearAtAll = false;
  for (const ev of labelEvents) {
    if (!ev.year) {
      ev.year = defaultYear;
      ev.inferred = true;
      noYearAtAll = true;
    }
  }

  // --------------------------------------------------------------------------
  // UNIFICAR NOMBRES POR CÓDIGO
  // --------------------------------------------------------------------------

  const latestName = new Map();
  const latestRank = new Map();

  for (const e of entries) {
    if (!e.code || e.labelIdx < 0) continue;
    const ev = labelEvents[e.labelIdx];
    if (!ev) continue;

    const rank = ev.year * 12 + ev.month;
    if (!latestRank.has(e.code) || rank >= latestRank.get(e.code)) {
      latestRank.set(e.code, rank);
      latestName.set(e.code, e.name);
    }
  }

  // --------------------------------------------------------------------------
  // ARMAR FORMATO LARGO
  // --------------------------------------------------------------------------

  const records = [];
  const notes = [];
  const colorCount = new Map();
  const unknownFills = new Map();

  let uncolored = 0;
  let paymentMismatch = 0;
  const totalMismatch = [];
  let noLabelRows = 0;
  let unifiedNames = 0;
  const months = new Set();

  for (const e of entries) {
    const ev = e.labelIdx >= 0 ? labelEvents[e.labelIdx] : null;

    if (!ev) {
      noLabelRows += 1;
      continue;
    }

    const monthDate = new Date(Date.UTC(ev.year, ev.month - 1, 1));
    const monthId = `${ev.year}-${String(ev.month).padStart(2, "0")}`;
    months.add(monthId);

    let shownName = e.name;
    if (unifyNames && e.code && latestName.get(e.code) && latestName.get(e.code) !== e.name) {
      shownName = latestName.get(e.code);
      unifiedNames += 1;
    }

    const documento = e.code;
    const { concepts, tecCol, paymentsCol } = e.header;

    let total = 0;
    let payrollSum = 0;
    const rowsOut = [];

    // ---- Conceptos ----
    for (const c of concepts) {
      const cell = e.row[c.idx];
      const value = round2(toNumber(cellValue(cell)));

      if (paymentsCol !== null && c.idx < paymentsCol) {
        payrollSum = round2(payrollSum + value);
      }

      if (value === 0) continue; // los ceros no salen al formato largo

      total = round2(total + value);

      const rawFill = cellFill(cell);
      const fill = statusFill(rawFill);

      if (!fill && rawFill && isHexFill(rawFill) && !STRUCTURAL_FILLS.has(String(rawFill).toUpperCase())) {
        const hex = String(rawFill).toUpperCase();
        unknownFills.set(hex, (unknownFills.get(hex) || 0) + 1);
      }

      rowsOut.push({ label: c.label, value, fill });
    }

    // ---- PAYMENTS ----
    if (paymentsCol !== null) {
      const paid = round2(toNumber(cellValue(e.row[paymentsCol])));
      if (Math.abs(paid - payrollSum) > 1) paymentMismatch += 1;
    }

    total = round2(total);

    // ---- TOTAL EMPLOYEE COST de la hoja (control) ----
    if (tecCol !== null) {
      const sheetTotal = round2(toNumber(cellValue(e.row[tecCol])));
      if (sheetTotal !== 0 && Math.abs(sheetTotal - total) > 1) {
        totalMismatch.push({ month: monthId, name: shownName, diff: round2(total - sheetTotal) });
      }
    }

    // ---- Registros de conceptos ----
    for (const o of rowsOut) {
      records.push({
        "Mes elaboración": monthDate,
        Concepto: o.label,
        Empleado: shownName,
        Documento: documento,
        "Valor Concepto": o.value,
        "Valor Totales": 0,
        _fill: o.fill
      });

      if (o.fill) colorCount.set(o.fill, (colorCount.get(o.fill) || 0) + 1);
      else uncolored += 1;
    }

    // ---- TOTAL EMPLOYEE COST ----
    if (total !== 0) {
      let totalFill = tecCol !== null ? statusFill(cellFill(e.row[tecCol])) : null;

      const info = blockInfo.get(e.blockId);
      if (!totalFill && info && info.employees === 1) totalFill = info.checkFill;

      records.push({
        "Mes elaboración": monthDate,
        Concepto: "TOTAL EMPLOYEE COST",
        Empleado: shownName,
        Documento: documento,
        "Valor Concepto": 0,
        "Valor Totales": total,
        _fill: totalFill
      });

      if (totalFill) colorCount.set(totalFill, (colorCount.get(totalFill) || 0) + 1);
      else uncolored += 1;
    }
  }

  // --------------------------------------------------------------------------
  // AVISOS
  // --------------------------------------------------------------------------

  const sortedMonths = [...months].sort();

  notes.push({
    type: "info",
    text:
      `Se leyeron ${entries.length} empleado(s)-mes de ${sortedMonths.length} mes(es)` +
      (sortedMonths.length ? ` (${sortedMonths[0]} a ${sortedMonths[sortedMonths.length - 1]})` : "") +
      "."
  });

  if (colorCount.size > 0 || uncolored > 0) {
    const parts = [...colorCount.entries()]
      .sort((a, b) => b[1] - a[1])
      .map(([hex, n]) => `${STATUS_COLORS[hex] || `#${hex}`}: ${n}`);

    notes.push({
      type: "info",
      text:
        `Colores tomados de la nómina — ${parts.length ? parts.join(", ") : "ninguno"}` +
        (uncolored > 0 ? `; sin color: ${uncolored}` : "") +
        "."
    });
  }

  if (unifiedNames > 0) {
    notes.push({
      type: "info",
      text: `Se unificó el nombre en ${unifiedNames} fila(s) usando el nombre más reciente de cada código de empleado.`
    });
  }

  if (labelEvents.some((ev) => ev.inferred)) {
    notes.push({
      type: noYearAtAll ? "warn" : "info",
      text: noYearAtAll
        ? `Ningún rótulo de mes trae el año; se usó ${defaultYear}. Revisa la columna Mes elaboración.`
        : "Algunos rótulos de mes no traen el año; se dedujo por el orden de los bloques."
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
      .map((m) => `${m.month} ${m.name} (${m.diff > 0 ? "+" : ""}${formatMoney(m.diff)})`)
      .join("; ");

    notes.push({
      type: "warn",
      text: `${totalMismatch.length} fila(s) donde el TOTAL EMPLOYEE COST calculado no coincide con el de la hoja. Ejemplos: ${ex}.`
    });
  }

  if (unknownFills.size > 0) {
    const list = [...unknownFills.entries()].map(([hex, n]) => `#${hex} (${n})`).join(", ");
    notes.push({
      type: "warn",
      text: `Colores de la nómina que no están en la leyenda del cruce y se dejaron sin color: ${list}.`
    });
  }

  if (noLabelRows > 0) {
    notes.push({
      type: "warn",
      text: `${noLabelRows} empleado(s) aparecen antes de cualquier rótulo de mes y se omitieron.`
    });
  }

  return { records, notes };
}

// ============================================================================
// CONSOLIDAR VARIAS NÓMINAS
// ============================================================================

export function convertNominaFiles(files, options = {}) {
  const usable = (files || []).filter((f) => f && Array.isArray(f.rows) && f.rows.length > 0);
  if (usable.length === 0) return null;

  const allRows = [];
  for (const f of usable) {
    allRows.push([{ v: FILE_BREAK, f: null }]);
    allRows.push(...f.rows);
  }

  const result = convertNominaRows(allRows, options);
  if (!result) return null;

  result.records = [...result.records].sort((a, b) => a["Mes elaboración"] - b["Mes elaboración"]);

  const names = usable.map((f, i) => f.name || `archivo ${i + 1}`).join(", ");
  result.notes.unshift({
    type: "info",
    text: `Se consolidaron ${usable.length} archivo(s) de entrada en un solo resultado (${names}).`
  });

  return result;
}

// ============================================================================
// MOVIMIENTO CC
// ============================================================================
// group:
//   "empleado" -> el Tercero es el empleado (salario, auxilios, gastos...)
//   "aporte"   -> aportes patronales: el Tercero es la entidad (Compensar, ARL...);
//                 se asignan al empleado de la misma fecha
//   "seguro"   -> seguros de vida: el Tercero es la aseguradora; el empleado
//                 viene al final de la Descripción ("Seguros de vida - Ramon Cardona")
//   "excluir"  -> facturación / ingresos, no es costo de un empleado
// El orden importa: gana la primera palabra clave que coincida.

export const CONCEPT_KEYWORDS = [
  { keyword: "EO THIRD PARTIES", concepto: null, group: "excluir" },
  { keyword: "INGRESOS RECIBIDOS", concepto: null, group: "excluir" },

  { keyword: "001050 - SALARIO", concepto: "SALARY", group: "empleado" },
  { keyword: "SALARIO", concepto: "SALARY", group: "empleado" },
  { keyword: "SUBSIDIO DE TRANSPORTE", concepto: "Transport allowance", group: "empleado" },
  { keyword: "AUXILIO EXTRALEGAL", concepto: "Allowance (No salarial)", group: "empleado" },
  { keyword: "AUXILIO MOVILIZACION", concepto: "Allowance (Car)", group: "empleado" },
  { keyword: "AUXILIO TELECOMUNICACIONES", concepto: "Allowance (Mobile & Internet)", group: "empleado" },
  { keyword: "GROSS UP", concepto: "Allowance (Gross up)", group: "empleado" },
  { keyword: "DIAS HABILES EN VACACIONES", concepto: "Vacation (días hábiles)", group: "empleado" },
  { keyword: "DIAS NO HABILES EN VACACIONES", concepto: "Vacation (días no hábiles)", group: "empleado" },
  { keyword: "LICENCIA REMUNERADA", concepto: "Paid leave", group: "empleado" },
  { keyword: "PRIMA DE SERVICIOS", concepto: "13TH SALARY", group: "empleado" },
  { keyword: "INTERESES CESANTIAS", concepto: "INTEREST ON 14TH SALARY", group: "empleado" },
  { keyword: "CONSIGNACION CESANTIAS", concepto: "14TH SALARY", group: "empleado" },
  { keyword: "GASTOS", concepto: "EXPENSES", group: "empleado" },
  { keyword: "FPP", concepto: "Other (People Pass / dotación)", group: "empleado", wholeWord: true },

  { keyword: "SEGUROS DE VIDA", concepto: "HEALTH INSURANCE", group: "seguro" },
  { keyword: "SEGURO DE VIDA", concepto: "HEALTH INSURANCE", group: "seguro" },

  { keyword: "APORTES A FONDOS DE", concepto: "PENSION COST", group: "aporte" },
  { keyword: "APORTES A ENTIDADES PROMOTORAS DE SALUD", concepto: "HEALTH COST", group: "aporte" },
  { keyword: "APORTES A CAJAS DE COMPENSACION", concepto: "FAMILY FUND COST", group: "aporte" },
  { keyword: "APORTES A ADMINISTRADORAS DE RIESGOS", concepto: "LABOR RISK COST", group: "aporte" },
  { keyword: "SENA", concepto: "SENA COST", group: "aporte", wholeWord: true },
  { keyword: "ICBF", concepto: "ICBF COST", group: "aporte", wholeWord: true }
];

const MOVIMIENTO_REQUIRED_HEADERS = ["COMPROBANTE", "FECHA ELABORACION", "DESCRIPCION", "DEBITO", "CREDITO"];

const MOVIMIENTO_STOP_LABEL =
  /^(CRUCE OK|NO ESTA EN EL OTRO LADO|DIFERENCIAS|CRUZA ENTRE MESES|CRUZA ENTRE ANO|DEBITO - CREDITO SE ANULAN)/;

// ============================================================================
// FECHAS MOVIMIENTO CC
// ============================================================================

function dateFromExcelSerial(value) {
  if (value instanceof Date) return value;

  if (typeof value === "number" && Number.isFinite(value)) {
    return new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 86400000);
  }

  if (typeof value === "string" && value.trim()) {
    const text = value.trim();

    // dd/mm/yyyy  o  dd-mm-yyyy
    const m = /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/.exec(text);
    if (m) {
      const d = new Date(Date.UTC(Number(m[3]), Number(m[2]) - 1, Number(m[1])));
      if (!Number.isNaN(d.getTime())) return d;
    }

    const parsed = new Date(text);
    if (!Number.isNaN(parsed.getTime())) return parsed;
  }

  return null;
}

function firstOfMonth(date) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), 1));
}

function monthKey(date) {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`;
}

function dateKey(date) {
  return date.toISOString().slice(0, 10);
}

// ============================================================================
// ENCABEZADO MOVIMIENTO CC
// ============================================================================

function readMovimientoHeader(row) {
  const map = {};

  for (let c = 0; c < row.length; c++) {
    const t = norm(cellValue(row[c]));
    if (t && map[t] === undefined) map[t] = c;
  }

  for (const required of MOVIMIENTO_REQUIRED_HEADERS) {
    if (map[required] === undefined) return null;
  }

  return {
    comprobanteCol: map["COMPROBANTE"],
    fechaCol: map["FECHA ELABORACION"],
    descripcionCol: map["DESCRIPCION"],
    // "Nombre tercero" (o "Tercero")
    terceroCol: map["NOMBRE TERCERO"] ?? map["TERCERO"] ?? null,
    // "Identificación tercero"
    docCol: map["IDENTIFICACION TERCERO"] ?? map["NIT"] ?? null,
    debitoCol: map["DEBITO"],
    creditoCol: map["CREDITO"]
  };
}

// ============================================================================
// BUSCAR CONCEPTO EN DESCRIPCIÓN
// ============================================================================

function findMovimientoConcept(descripcionNorm, conceptKeywords) {
  const words = ` ${descripcionNorm.replace(/[^A-Z0-9]+/g, " ")} `;

  for (const entry of conceptKeywords) {
    const k = norm(entry.keyword);
    if (!k) continue;

    const hit = entry.wholeWord ? words.includes(` ${k} `) : descripcionNorm.includes(k);
    if (hit) return entry;
  }

  return null;
}

// ============================================================================
// EXTRAER EMPLEADO DESDE DESCRIPCIÓN (cuando el Tercero viene vacío)
// ============================================================================
// Ej.: "001050 - Salario OFIR ELIZABETH ESPAÑA LOPEZ" -> "OFIR ELIZABETH ESPAÑA LOPEZ"

function extractEmployeeFromDescription(descripcion, match) {
  if (!descripcion || !match || match.group !== "empleado") return "";

  let text = clean(descripcion);
  const normalized = norm(text);
  const keyword = norm(match.keyword);
  const index = normalized.indexOf(keyword);

  if (index >= 0) {
    text = text.slice(index + keyword.length);
  } else {
    text = text.replace(/^\s*\d+\s*[-:]\s*/, "");
  }

  text = text
    .replace(/^\s*[-:.;]+\s*/, "")
    .replace(/^\s*\.+\s*/, "")
    .replace(/\s+/g, " ")
    .trim();

  if (!text || text.length < 3) return "";

  const upper = norm(text);

  if (/^(SALARIO|SUBSIDIO|AUXILIO|APORTES|APORTE|PENSION|SALUD|RIESGO|PRIMA|VACACIONES|LICENCIA)/.test(upper)) {
    return "";
  }

  // "001050 - Salario Integral": "Integral" no es un nombre
  if (/^(INTEGRAL|ORDINARIO)$/.test(upper)) return "";

  return clean(text);
}

// ============================================================================
// MOVIMIENTO CC -> FORMATO LARGO
// ============================================================================

export function convertMovimientoRows(rows, options = {}) {
  const conceptKeywords = options.conceptKeywords || CONCEPT_KEYWORDS;
  const matchConceptWith = (descripcionNorm) => findMovimientoConcept(descripcionNorm, conceptKeywords);

  let header = null;
  let started = false;

  const parsedRows = [];
  const employeesByDate = new Map();
  const employeesByMonth = new Map();
  const nameToDoc = new Map(); // empleado -> cédula (solo filas donde el tercero es el empleado)
  const unknownFills = new Map();

  let excludedRows = 0;
  let skippedNoDate = 0;
  let dataRows = 0;
  let unclassifiedCount = 0;
  const unclassifiedExamples = new Map();

  // --------------------------------------------------------------------------
  // PASADA 1
  // --------------------------------------------------------------------------

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];

    if (!started) {
      const h = readMovimientoHeader(row);
      if (h) {
        header = h;
        started = true;
      }
      continue;
    }

    // Otro encabezado dentro del archivo.
    if (readMovimientoHeader(row)) continue;

    const comprobanteVal = clean(cellValue(row[header.comprobanteCol]));
    const descripcionVal = clean(cellValue(row[header.descripcionCol]));

    if (!comprobanteVal && MOVIMIENTO_STOP_LABEL.test(norm(descripcionVal))) break;
    if (!comprobanteVal && !descripcionVal) continue;

    const fecha = dateFromExcelSerial(cellValue(row[header.fechaCol]));
    if (!fecha) {
      skippedNoDate += 1;
      continue;
    }

    const monthDate = firstOfMonth(fecha);
    const dKey = dateKey(fecha);

    const terceroVal = header.terceroCol !== null ? clean(cellValue(row[header.terceroCol])) : "";
    const docVal =
      header.docCol !== null ? clean(cellValue(row[header.docCol])).replace(/\D/g, "") : "";

    const debitoCell = row[header.debitoCol];
    const creditoCell = row[header.creditoCol];
    const debito = toNumber(cellValue(debitoCell));
    const credito = toNumber(cellValue(creditoCell));
    const value = round2(debito - credito);

    if (value === 0) continue;

    const descripcionNorm = norm(descripcionVal);
    const match = matchConceptWith(descripcionNorm);

    if (match && match.group === "excluir") {
      excludedRows += 1;
      continue;
    }

    const group = match ? match.group : "otro";
    const concepto = match ? match.concepto : descripcionVal;

    if (!match) {
      unclassifiedCount += 1;
      unclassifiedExamples.set(descripcionVal, (unclassifiedExamples.get(descripcionVal) || 0) + 1);
    }

    // ---- Empleado y documento según el tipo de fila ----
    let resolvedEmployee = "";
    let documento = "";

    if (group === "empleado") {
      // El tercero es el empleado; si viene vacío se intenta sacar de la descripción.
      resolvedEmployee = terceroVal || extractEmployeeFromDescription(descripcionVal, match);
      if (terceroVal) documento = docVal;
    } else if (group === "seguro") {
      // El tercero es la aseguradora: el empleado va al final de la descripción.
      resolvedEmployee = clean(descripcionVal.split("-").pop());
    }
    // "aporte" y "otro": no se toma el tercero (es una entidad); los aportes se asignan por fecha.

    if (group === "empleado" && resolvedEmployee) {
      if (documento) nameToDoc.set(resolvedEmployee, documento);

      if (!employeesByDate.has(dKey)) employeesByDate.set(dKey, new Set());
      employeesByDate.get(dKey).add(resolvedEmployee);

      const mKey = monthKey(monthDate);
      if (!employeesByMonth.has(mKey)) employeesByMonth.set(mKey, new Set());
      employeesByMonth.get(mKey).add(resolvedEmployee);
    }

    // ---- Color ----
    const rawFill =
      cellFill(debito !== 0 ? debitoCell : creditoCell) || cellFill(debitoCell) || cellFill(creditoCell);
    const fill = statusFill(rawFill);

    if (!fill && rawFill && isHexFill(rawFill) && !STRUCTURAL_FILLS.has(String(rawFill).toUpperCase())) {
      const hex = String(rawFill).toUpperCase();
      unknownFills.set(hex, (unknownFills.get(hex) || 0) + 1);
    }

    dataRows += 1;

    parsedRows.push({
      monthDate,
      dKey,
      comprobante: comprobanteVal,
      concepto,
      group,
      empleadoDirecto: resolvedEmployee,
      value,
      fill
    });
  }

  if (!header || dataRows === 0) return null;

  // --------------------------------------------------------------------------
  // PASADA 2: asignar empleado (nombres incompletos y aportes patronales)
  // --------------------------------------------------------------------------

  const knownEmployees = [...nameToDoc.keys()];

  // "Ramon Cardona" -> "RAMON ESTEBAN CARDONA SALAZAR" si hay un único nombre que lo contenga
  const fullNameOf = (name) => {
    const t = norm(name).split(" ").filter(Boolean);
    if (t.length < 2) return name;
    const hits = knownEmployees.filter((k) => {
      const kt = norm(k).split(" ");
      return t.every((x) => kt.includes(x));
    });
    return hits.length === 1 ? hits[0] : name;
  };

  let ambiguousAporte = 0;
  let unresolvedAporte = 0;
  let monthFallback = 0;

  for (const pr of parsedRows) {
    if (pr.group !== "aporte") {
      pr.empleado = fullNameOf(pr.empleadoDirecto || "");
      continue;
    }

    const byDate = employeesByDate.get(pr.dKey);
    const byMonth = employeesByMonth.get(monthKey(pr.monthDate));
    const pool = byDate && byDate.size > 0 ? byDate : byMonth;

    if (pool && pool.size === 1) {
      pr.empleado = [...pool][0];
      if (pool === byMonth && !(byDate && byDate.size > 0)) monthFallback += 1;
    } else {
      pr.empleado = "";
      if (pool && pool.size > 1) ambiguousAporte += 1;
      else unresolvedAporte += 1;
    }
  }

  // --------------------------------------------------------------------------
  // AGRUPAR
  // --------------------------------------------------------------------------

  const grouped = new Map();
  const totalsByMonthEmployee = new Map();
  const colorCount = new Map();
  let uncolored = 0;
  const months = new Set();

  for (const pr of parsedRows) {
    months.add(monthKey(pr.monthDate));

    const key = `${monthKey(pr.monthDate)}|${pr.concepto}|${pr.empleado}`;

    if (!grouped.has(key)) {
      grouped.set(key, {
        monthDate: pr.monthDate,
        concepto: pr.concepto,
        empleado: pr.empleado,
        documento: nameToDoc.get(pr.empleado) || "",
        value: 0,
        fill: pr.fill
      });
    }

    const g = grouped.get(key);
    g.value = round2(g.value + pr.value);
    if (!g.fill && pr.fill) g.fill = pr.fill;

    if (pr.empleado) {
      const teKey = `${monthKey(pr.monthDate)}|${pr.empleado}`;
      totalsByMonthEmployee.set(teKey, round2((totalsByMonthEmployee.get(teKey) || 0) + pr.value));
    }
  }

  // --------------------------------------------------------------------------
  // RECORDS
  // --------------------------------------------------------------------------

  const records = [];

  for (const g of grouped.values()) {
    if (g.value === 0) continue;

    records.push({
      "Mes elaboración": g.monthDate,
      Concepto: g.concepto,
      Empleado: g.empleado || SIN_EMPLEADO,
      Documento: g.documento,
      "Valor Concepto": g.value,
      "Valor Totales": 0,
      _fill: g.fill
    });

    if (g.fill) colorCount.set(g.fill, (colorCount.get(g.fill) || 0) + 1);
    else uncolored += 1;
  }

  // ---- TOTAL EMPLOYEE COST ----
  for (const [key, total] of totalsByMonthEmployee) {
    if (total === 0) continue;

    const parts = key.split("|");
    const mKey = parts[0];
    const empleado = parts.slice(1).join("|");
    const [y, m] = mKey.split("-").map(Number);

    records.push({
      "Mes elaboración": new Date(Date.UTC(y, m - 1, 1)),
      Concepto: "TOTAL EMPLOYEE COST",
      Empleado: empleado || SIN_EMPLEADO,
      Documento: nameToDoc.get(empleado) || "",
      "Valor Concepto": 0,
      "Valor Totales": total,
      _fill: null
    });
  }

  // --------------------------------------------------------------------------
  // ORDEN
  // --------------------------------------------------------------------------

  const firstSeen = new Map();
  parsedRows.forEach((pr, idx) => {
    const key = `${monthKey(pr.monthDate)}|${pr.empleado || SIN_EMPLEADO}|${pr.concepto}`;
    if (!firstSeen.has(key)) firstSeen.set(key, idx);
  });

  const sortKey = (rec) => {
    const d = rec["Mes elaboración"];
    const monthNum = d.getUTCFullYear() * 12 + d.getUTCMonth();
    const isTotal = rec.Concepto === "TOTAL EMPLOYEE COST" ? 1 : 0;
    const orderKey = `${monthKey(d)}|${rec.Empleado}|${rec.Concepto}`;
    const order = firstSeen.has(orderKey) ? firstSeen.get(orderKey) : Number.MAX_SAFE_INTEGER;
    return [monthNum, rec.Empleado, isTotal, order];
  };

  records.sort((a, b) => {
    const ka = sortKey(a);
    const kb = sortKey(b);
    for (let i = 0; i < ka.length; i++) {
      if (ka[i] < kb[i]) return -1;
      if (ka[i] > kb[i]) return 1;
    }
    return 0;
  });

  // --------------------------------------------------------------------------
  // AVISOS
  // --------------------------------------------------------------------------

  const notes = [];
  const sortedMonths = [...months].sort();

  notes.push({
    type: "info",
    text:
      `Se leyeron ${dataRows} movimiento(s) en ${sortedMonths.length} mes(es)` +
      (sortedMonths.length ? ` (${sortedMonths[0]} a ${sortedMonths[sortedMonths.length - 1]})` : "") +
      ` y se agruparon en ${grouped.size} fila(s) de concepto + ${totalsByMonthEmployee.size} de TOTAL EMPLOYEE COST.`
  });

  if (excludedRows > 0) {
    notes.push({
      type: "info",
      text: `${excludedRows} movimiento(s) de facturación/ingresos se excluyeron por no ser costo de un empleado.`
    });
  }

  if (header.docCol === null) {
    notes.push({
      type: "warn",
      text: 'No se encontró la columna "Identificación tercero": el cruce con la nómina se hará solo por nombre.'
    });
  }

  if (colorCount.size > 0 || uncolored > 0) {
    const parts = [...colorCount.entries()]
      .sort((a, b) => b[1] - a[1])
      .map(([hex, n]) => `${STATUS_COLORS[hex] || `#${hex}`}: ${n}`);

    notes.push({
      type: "info",
      text:
        `Colores tomados del movimiento — ${parts.length ? parts.join(", ") : "ninguno"}` +
        (uncolored > 0 ? `; sin color: ${uncolored}` : "") +
        "."
    });
  }

  if (monthFallback > 0) {
    notes.push({
      type: "info",
      text: `${monthFallback} movimiento(s) de aporte patronal no tenían un salario en su fecha exacta y se asignaron al único empleado de ese mes.`
    });
  }

  if (unresolvedAporte > 0) {
    notes.push({
      type: "warn",
      text: `${unresolvedAporte} fila(s) de aporte patronal no se pudieron asignar a ningún empleado porque no hay ninguna fila de salario en esa fecha ni en ese mes; quedaron marcadas como "${SIN_EMPLEADO}".`
    });
  }

  if (ambiguousAporte > 0) {
    notes.push({
      type: "warn",
      text: `${ambiguousAporte} fila(s) de aporte patronal quedaron marcadas como "${SIN_EMPLEADO}" porque hay más de un empleado candidato.`
    });
  }

  if (unclassifiedCount > 0) {
    const ex = [...unclassifiedExamples.entries()]
      .slice(0, 5)
      .map(([desc, n]) => `"${desc}" (${n})`)
      .join("; ");

    notes.push({
      type: "warn",
      text: `${unclassifiedCount} movimiento(s) no coincidieron con ninguna palabra clave de CONCEPT_KEYWORDS y quedaron con el texto tal cual de Descripción: ${ex}.`
    });
  }

  if (skippedNoDate > 0) {
    notes.push({
      type: "warn",
      text: `${skippedNoDate} fila(s) se omitieron por no tener una fecha válida en "Fecha elaboración".`
    });
  }

  if (unknownFills.size > 0) {
    const list = [...unknownFills.entries()].map(([hex, n]) => `#${hex} (${n})`).join(", ");
    notes.push({
      type: "warn",
      text: `Colores del movimiento que no están en la leyenda del cruce y se dejaron sin color: ${list}.`
    });
  }

  notes.push({
    type: "info",
    text:
      "Movimiento CC: el empleado se toma de Nombre tercero en salarios, auxilios y gastos (y su cédula de Identificación tercero); en los seguros de vida, de la Descripción; y los aportes patronales se asignan por fecha y, cuando solo existe un candidato, por mes."
  });

  return { records, notes };
}

// ============================================================================
// CONSOLIDAR VARIOS MOVIMIENTO CC
// ============================================================================

function trimMovimientoStopSection(rows) {
  let header = null;
  let headerFound = false;

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];

    if (!headerFound) {
      const h = readMovimientoHeader(row);
      if (h) {
        header = h;
        headerFound = true;
      }
      continue;
    }

    if (readMovimientoHeader(row)) continue;

    const comprobanteVal = clean(cellValue(row[header.comprobanteCol]));
    const descripcionVal = clean(cellValue(row[header.descripcionCol]));

    if (!comprobanteVal && MOVIMIENTO_STOP_LABEL.test(norm(descripcionVal))) {
      return rows.slice(0, r);
    }
  }

  return rows;
}

export function convertMovimientoFiles(files, options = {}) {
  const usable = (files || []).filter((f) => f && Array.isArray(f.rows) && f.rows.length > 0);
  if (usable.length === 0) return null;

  const allRows = [];
  for (const f of usable) allRows.push(...trimMovimientoStopSection(f.rows));

  const result = convertMovimientoRows(allRows, options);
  if (!result) return null;

  result.records = [...result.records].sort((a, b) => a["Mes elaboración"] - b["Mes elaboración"]);

  const names = usable.map((f, i) => f.name || `archivo ${i + 1}`).join(", ");
  result.notes.unshift({
    type: "info",
    text: `Se consolidaron ${usable.length} archivo(s) de Movimiento CC en un solo resultado (${names}).`
  });

  return result;
}