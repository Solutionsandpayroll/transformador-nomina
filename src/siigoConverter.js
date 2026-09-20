// ============================================================================
// CONVERTIDOR: Movimiento CC de Siigo -> formato largo de nómina
// ============================================================================
// Entrada: la hoja del "Movimiento CC" (columnas Comprobante, Fecha elaboración,
// Descripción, Tercero, Débito, Crédito), como array de arrays (la misma forma
// que produce parseSheetXmlToRows en App.jsx).
//
// Salida: filas con las columnas del ejemplo real (Hoja2 de BUBBLE - Nómina):
//   Mes elaboración | Concepto | Empleado | Valor Concepto | Valor Totales
// más una marca interna `_fill` ('green' = viene de Siigo, 'yellow' = calculado)
// que se usa solo para colorear la vista previa y el Excel exportado.
//
// Qué se toma de Siigo (solo comprobantes CC-*, las FV/NC de facturación se ignoran):
//   Salario, Subsidio de transporte, Auxilio extralegal, Vacaciones, Licencia
//   remunerada, y los aportes de pensión, salud, cajas y ARL.
// Qué se CALCULA (no existe en Siigo, va en amarillo):
//   13TH SALARY, 14TH SALARY e INTEREST ON 14TH SALARY.
// Qué se omite a propósito (ya lo cubren las provisiones):
//   Prima de servicios, intereses de cesantías, consignación de cesantías.
// Qué NO se puede generar desde Siigo:
//   SENA QUOTE u otros conceptos manuales. Lo que no se reconoce queda listado
//   en las notas para que lo revises; nunca se descarta en silencio.

// --- Configuración editable ---------------------------------------------------

// Nombre con el que se rotula el "Auxilio Extralegal" de Siigo. En el ejemplo
// real cambia de un mes a otro (junio usa "Alloawance 2", julio "Alloawance 4"),
// así que se puede fijar por mes con la clave 'YYYY-MM'.
export const DEFAULT_ALLOWANCE_NAME = 'Alloawance 4 (Other allowances)';
export const ALLOWANCE_NAME_BY_MONTH = {
  '2026-06': 'Alloawance 2 (Mobile & Internet Allowance)'
};

// Cambios manuales de nombre: { 'NOMBRE COMO VIENE EN SIIGO': 'NOMBRE A MOSTRAR' }
export const NAME_OVERRIDES = {};

// Orden en que salen los conceptos dentro de cada mes (igual que la Hoja2).
const CONCEPT_ORDER = [
  'SALARY',
  'ANUAL LEAVE',
  'Paid Leave',
  'ALLOWANCE', // se reemplaza por el nombre configurado arriba
  'Transport allowance',
  'PENSION COST',
  'HEALTH COST',
  'LABOR RISK COST',
  'FAMILY FUND COST',
  '13TH SALARY',
  '14TH SALARY',
  'INTEREST ON 14TH SALARY'
];

const SS_CONCEPTS = new Set(['PENSION COST', 'HEALTH COST', 'LABOR RISK COST', 'FAMILY FUND COST']);
const PAYROLL_BASE_CONCEPTS = ['SALARY', 'ANUAL LEAVE', 'Paid Leave', 'Transport allowance'];

// --- Utilidades ---------------------------------------------------------------

function norm(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function toNumber(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const n = Number(String(value).replace(/,/g, ''));
  return Number.isFinite(n) ? n : 0;
}

// Acepta serial de Excel, "dd/mm/yyyy" (texto, como en abril 2026) y "yyyy-mm-dd".
export function parseSiigoDate(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') {
    if (value < 20000 || value > 80000) return null;
    return new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000);
  }
  const s = String(value).trim();
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
  if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  m = /^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})/.exec(s);
  if (m) {
    let day = +m[1];
    let month = +m[2];
    if (month > 12 && day <= 12) [day, month] = [month, day]; // venía como mm/dd
    return new Date(Date.UTC(+m[3], month - 1, day));
  }
  const asNumber = Number(s);
  if (Number.isFinite(asNumber)) return parseSiigoDate(asNumber);
  return null;
}

function monthKeyOf(date) {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}`;
}

function formatMoney(n) {
  return Math.round(n).toLocaleString('es-CO');
}

// Encuentra la fila de encabezado: debe traer las 5 columnas clave a la vez.
// (Así se ignoran las hojas auxiliares con tablas dinámicas.)
export function findSiigoHeader(rows) {
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    const cols = {};
    for (let c = 0; c < row.length; c++) {
      const h = norm(row[c]);
      if (!h) continue;
      if (cols.comprobante === undefined && h === 'COMPROBANTE') cols.comprobante = c;
      else if (cols.fecha === undefined && h.startsWith('FECHA')) cols.fecha = c;
      else if (cols.descripcion === undefined && h === 'DESCRIPCION') cols.descripcion = c;
      else if (cols.tercero === undefined && h === 'TERCERO') cols.tercero = c;
      else if (cols.debito === undefined && h === 'DEBITO') cols.debito = c;
      else if (cols.credito === undefined && h === 'CREDITO') cols.credito = c;
    }
    if (
      cols.comprobante !== undefined &&
      cols.fecha !== undefined &&
      cols.descripcion !== undefined &&
      cols.debito !== undefined &&
      cols.credito !== undefined
    ) {
      return { rowIndex: r, cols };
    }
  }
  return null;
}

// Clasifica la descripción de Siigo. Tolera los errores de digitación reales del
// archivo ("pesiones", "pensione", "ARL" pegado, con/sin tilde).
function classify(description) {
  const d = norm(description);
  if (/PRIMA DE SERVICIOS/.test(d)) return { excluded: 'Prima de servicios' };
  if (/INTERESES CESANTIAS/.test(d)) return { excluded: 'Intereses de cesantías' };
  if (/CONSIGNACION CESANTIAS/.test(d)) return { excluded: 'Consignación de cesantías' };
  if (/VACACIONES/.test(d)) return { concept: 'ANUAL LEAVE' };
  if (/LICENCIA REMUNERADA/.test(d)) return { concept: 'Paid Leave' };
  if (/SUBSIDIO DE TRANSPORTE/.test(d)) return { concept: 'Transport allowance' };
  if (/AUXILIO EXTRALEGAL/.test(d)) return { concept: 'ALLOWANCE' };
  if (/SALARIO/.test(d)) return { concept: 'SALARY' };
  if (/FONDOS? DE PE/.test(d)) return { concept: 'PENSION COST' };
  if (/PROMOTORAS DE SALUD|\bEPS\b/.test(d)) return { concept: 'HEALTH COST' };
  if (/CAJAS? DE COMPENSACION/.test(d)) return { concept: 'FAMILY FUND COST' };
  if (/RIESGOS LABORALES|\bARL\b/.test(d)) return { concept: 'LABOR RISK COST' };
  return null;
}

// "OFIR ELIZABETH ESPAÑA LOPEZ" -> "ESPAÑA LOPEZ OFIR ELIZABETH" (solo nombres de
// 4 palabras; con otra cantidad es ambiguo y se deja como viene).
function displayName(rawName, reorder) {
  if (NAME_OVERRIDES[rawName]) return NAME_OVERRIDES[rawName];
  if (!reorder) return rawName;
  const t = rawName.split(/\s+/).filter(Boolean);
  if (t.length === 4) return [t[2], t[3], t[0], t[1]].join(' ');
  return rawName;
}

// --- Conversión principal -----------------------------------------------------

export function convertSiigoRows(rows, options = {}) {
  const reorderNames = options.reorderNames !== false;
  const header = findSiigoHeader(rows);
  if (!header) return null; // no es un Movimiento CC de Siigo

  const { cols } = header;
  const notes = [];
  const entries = [];
  const unmapped = new Map(); // descripción -> { total, count, months:Set }
  const excluded = new Map(); // motivo -> { total, count }
  let skippedNoDate = 0;

  for (let r = header.rowIndex + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const comprobante = String(row[cols.comprobante] ?? '').trim().toUpperCase();
    if (!comprobante.startsWith('CC-')) continue; // FV, NC, totales, filas vacías

    const date = parseSiigoDate(row[cols.fecha]);
    const amount = toNumber(row[cols.debito]) - toNumber(row[cols.credito]);
    if (amount === 0) continue;
    if (!date) {
      skippedNoDate += 1;
      continue;
    }

    const description = String(row[cols.descripcion] ?? '').trim();
    const tercero = cols.tercero !== undefined ? String(row[cols.tercero] ?? '').trim() : '';
    const month = monthKeyOf(date);
    const kind = classify(description);

    if (!kind) {
      const key = description.replace(/\d{5,}/g, '#');
      const u = unmapped.get(key) || { total: 0, count: 0, months: new Set() };
      u.total += amount;
      u.count += 1;
      u.months.add(month);
      unmapped.set(key, u);
      continue;
    }
    if (kind.excluded) {
      const e = excluded.get(kind.excluded) || { total: 0, count: 0 };
      e.total += amount;
      e.count += 1;
      excluded.set(kind.excluded, e);
      continue;
    }
    entries.push({ month, concept: kind.concept, amount, tercero });
  }

  // Empleados: los aportes (pensión, salud...) vienen a nombre del fondo, no de la
  // persona, así que se asignan al empleado que tenga nómina ese mes.
  const personsByMonth = new Map();
  const allPersons = new Set();
  for (const e of entries) {
    if (SS_CONCEPTS.has(e.concept) || !e.tercero) continue;
    allPersons.add(e.tercero);
    if (!personsByMonth.has(e.month)) personsByMonth.set(e.month, new Set());
    personsByMonth.get(e.month).add(e.tercero);
  }
  // Meses sin tercero (en 2025 Siigo no lo trae): si en todo el archivo hay una
  // sola persona, se usa esa.
  const fallbackName = allPersons.size === 1 ? [...allPersons][0] : null;
  const ambiguousMonths = new Set();

  const buckets = new Map(); // `${month}|${empleado}` -> Map(concepto -> valor)
  for (const e of entries) {
    let employee;
    if (SS_CONCEPTS.has(e.concept)) {
      const persons = personsByMonth.get(e.month);
      if (persons && persons.size === 1) employee = [...persons][0];
      else if (persons && persons.size > 1) {
        employee = 'SIN ASIGNAR (aportes)';
        ambiguousMonths.add(e.month);
      } else employee = fallbackName || 'SIN TERCERO';
    } else {
      employee = e.tercero || fallbackName || 'SIN TERCERO';
    }
    const key = `${e.month}|${employee}`;
    if (!buckets.has(key)) buckets.set(key, new Map());
    const m = buckets.get(key);
    m.set(e.concept, (m.get(e.concept) || 0) + e.amount);
  }

  // Primer mes con datos de cada empleado en cada año (para el interés acumulado).
  const firstMonthOfYear = new Map(); // `${empleado}|${año}` -> número de mes mínimo
  for (const key of buckets.keys()) {
    const [month, employee] = key.split('|');
    const [year, mm] = month.split('-').map(Number);
    const k = `${employee}|${year}`;
    if (!firstMonthOfYear.has(k) || mm < firstMonthOfYear.get(k)) firstMonthOfYear.set(k, mm);
  }

  const records = [];
  const sortedKeys = [...buckets.keys()].sort();
  for (const key of sortedKeys) {
    const [month, employee] = key.split('|');
    const [year, mm] = month.split('-').map(Number);
    const concepts = new Map(buckets.get(key));
    const monthDate = new Date(Date.UTC(year, mm - 1, 1));
    const shownName = displayName(employee, reorderNames);

    // Provisiones (calculadas). Base = salario + transporte + vacaciones + licencia;
    // el auxilio extralegal no salarial no entra.
    const base = PAYROLL_BASE_CONCEPTS.reduce((s, c) => s + (concepts.get(c) || 0), 0);
    const provisionKeys = new Set();
    if (base > 0) {
      const monthly = Math.round(base / 12);
      const monthsSinceStart = mm - firstMonthOfYear.get(`${employee}|${year}`) + 1;
      concepts.set('13TH SALARY', monthly);
      concepts.set('14TH SALARY', monthly);
      concepts.set('INTEREST ON 14TH SALARY', Math.round((monthly * (2 * monthsSinceStart - 1)) / 100));
      ['13TH SALARY', '14TH SALARY', 'INTEREST ON 14TH SALARY'].forEach((c) => provisionKeys.add(c));
    }

    let total = 0;
    for (const concept of CONCEPT_ORDER) {
      const value = concepts.get(concept);
      if (!value) continue;
      total += value;
      const label =
        concept === 'ALLOWANCE'
          ? ALLOWANCE_NAME_BY_MONTH[month] || DEFAULT_ALLOWANCE_NAME
          : concept;
      records.push({
        'Mes elaboración': monthDate,
        Concepto: label,
        Empleado: shownName,
        'Valor Concepto': value,
        'Valor Totales': 0,
        _fill: provisionKeys.has(concept) ? 'yellow' : 'green'
      });
    }
    if (total !== 0) {
      records.push({
        'Mes elaboración': monthDate,
        Concepto: 'TOTAL EMPLOYEE COST',
        Empleado: shownName,
        'Valor Concepto': 0,
        'Valor Totales': total,
        _fill: 'green'
      });
    }
  }

  // Notas para el usuario
  if (records.length > 0) {
    notes.push({
      type: 'info',
      text: '13TH SALARY, 14TH SALARY e INTEREST ON 14TH SALARY no existen en Siigo: se calculan (filas en amarillo). Revísalos contra tu criterio.'
    });
  }
  for (const [reason, e] of excluded) {
    notes.push({
      type: 'info',
      text: `Se omitieron ${e.count} línea(s) de "${reason}" (${formatMoney(e.total)}): ya las cubren las provisiones.`
    });
  }
  for (const [desc, u] of unmapped) {
    notes.push({
      type: 'warn',
      text: `Sin mapear: "${desc.slice(0, 70)}" — ${u.count} línea(s), ${formatMoney(u.total)} (${[...u.months].sort().join(', ')}). No está en el resultado.`
    });
  }
  if (ambiguousMonths.size > 0) {
    notes.push({
      type: 'warn',
      text: `Hay más de un empleado en ${[...ambiguousMonths].sort().join(', ')}: los aportes no se pueden repartir desde Siigo y quedaron como "SIN ASIGNAR (aportes)".`
    });
  }
  if (skippedNoDate > 0) {
    notes.push({ type: 'warn', text: `${skippedNoDate} línea(s) con fecha ilegible se omitieron.` });
  }
  if (allPersons.size > 1 && [...buckets.keys()].some((k) => k.endsWith('|SIN TERCERO'))) {
    notes.push({
      type: 'warn',
      text: 'Hay líneas de nómina sin tercero y varios empleados en el archivo: quedaron como "SIN TERCERO".'
    });
  }

  return { records, notes };
}