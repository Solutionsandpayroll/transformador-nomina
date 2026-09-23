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
//   remunerada, los aportes (pensión, salud, cajas, ARL, SENA, ICBF) y los extra
//   configurados en EXTRA_PAYROLL_PATTERNS (auxilios, gross up, bonificaciones).
// Qué se CALCULA (no existe en Siigo, va en amarillo):
//   13TH SALARY, 14TH SALARY e INTEREST ON 14TH SALARY.
// Qué se omite a propósito (ya lo cubren las provisiones):
//   Prima de servicios / prima legal, cesantías e intereses de cesantías.
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

// El salario integral ya incluye las prestaciones (prima, cesantías, intereses), así que
// por defecto NO se calculan 13TH SALARY, 14TH SALARY ni INTEREST ON 14TH SALARY sobre él.
// Ponlo en true si tu criterio contable es provisionarlos igual.
export const PROVISIONS_ON_INTEGRAL_SALARY = false;

// Cambios manuales de nombre: { 'NOMBRE COMO VIENE EN SIIGO': 'NOMBRE A MOSTRAR' }
export const NAME_OVERRIDES = {};

// Conceptos extra que SÍ son costo del empleado pero no tienen nombre fijo
// (auxilios, gross up, bonificaciones...). Se toman de Siigo con su nombre
// original limpio (sin códigos ni consecutivos) y salen en azul. Para incluir
// otros (por ejemplo pólizas o seguros de vida) agrega su patrón aquí, sin tildes
// y en mayúsculas, por ejemplo: /POLIZA|SEGUROS? DE VIDA/
export const EXTRA_PAYROLL_PATTERNS = [/AUXILIO|GROSS UP|BONIFICACION/];

// Orden en que salen los conceptos dentro de cada mes (igual que la Hoja2):
// primero los de nómina, luego los extra (alfabético) y al final aportes y provisiones.
const HEAD_ORDER = [
  'SALARY',
  'ANUAL LEAVE',
  'Paid Leave',
  'ALLOWANCE', // se reemplaza por el nombre configurado arriba
  'Transport allowance'
];
const TAIL_ORDER = [
  'PENSION COST',
  'HEALTH COST',
  'LABOR RISK COST',
  'FAMILY FUND COST',
  'SENA COST',
  'ICBF COST',
  '13TH SALARY',
  '14TH SALARY',
  'INTEREST ON 14TH SALARY'
];

// Aportes que Siigo registra a nombre de la entidad (fondo, EPS, caja, SENA...), no del
// empleado: se asignan al empleado que tenga nómina ese mes.
const SS_CONCEPTS = new Set([
  'PENSION COST',
  'HEALTH COST',
  'LABOR RISK COST',
  'FAMILY FUND COST',
  'SENA COST',
  'ICBF COST'
]);
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

// Quita códigos y consecutivos del inicio: "D016-AUXILIO X", "12345 - Auxilio X", "# Auxilio X".
function stripPrefix(text) {
  return String(text)
    .replace(/^\s*\d{4,}\s*-?\s*/, '')
    .replace(/^\s*[A-Za-z]\d{3}\s*-\s*/, '')
    .replace(/^\s*#\s*-?\s*/, '')
    .trim();
}

function titleCase(text) {
  return text
    .toLowerCase()
    .replace(/(^|[\s(/-])(\p{L})/gu, (m, sep, ch) => sep + ch.toUpperCase());
}

// Clasifica la descripción de Siigo. Tolera los errores de digitación reales del
// archivo ("pesiones", "pensione", "ARL" pegado, con/sin tilde).
function classify(description) {
  const d = norm(description);
  // Los aportes de pensión mencionan "cesantías"; se revisan antes que las cesantías.
  if (/FONDOS? DE PE/.test(d)) return { concept: 'PENSION COST' };
  if (/PRIMA DE SERVICIOS|PRIMA LEGAL|AJUSTE PRIMA/.test(d)) return { excluded: 'Prima de servicios' };
  if (/INTERESES (SOBRE )?CESANTIAS/.test(d)) return { excluded: 'Intereses de cesantías' };
  if (/CONSIGNACION CESANTIAS|(^|[^A-Z])CESANTIAS$/.test(d)) return { excluded: 'Cesantías' };
  if (/VACACIONES/.test(d)) return { concept: 'ANUAL LEAVE' };
  if (/LICENCIA REMUNERADA/.test(d)) return { concept: 'Paid Leave' };
  if (/SUBSIDIO DE TRANSPORTE/.test(d)) return { concept: 'Transport allowance' };
  if (/AUXILIO EXTRALEGAL/.test(d)) return { concept: 'ALLOWANCE' };
  if (/SALARIO/.test(d)) return { concept: 'SALARY' };
  if (/PROMOTORAS DE SALUD|\bEPS\b/.test(d)) return { concept: 'HEALTH COST' };
  if (/CAJAS? DE COMPENSACION/.test(d)) return { concept: 'FAMILY FUND COST' };
  if (/RIESGOS LABORALES|\bARL\b/.test(d)) return { concept: 'LABOR RISK COST' };

  const core = norm(stripPrefix(description));
  if (/^SENA_?$|^APORTES? (AL )?SENA$/.test(core)) return { concept: 'SENA COST' };
  if (/^(ICBF|IBCF)$|^APORTES? (AL )?ICBF$/.test(core)) return { concept: 'ICBF COST' };

  if (EXTRA_PAYROLL_PATTERNS.some((p) => p.test(d))) {
    return { extra: titleCase(stripPrefix(description)) };
  }
  return null;
}

// Agrupa lo que no se incluye para que el aviso sea corto en vez de una lista de cientos.
const UNMAPPED_GROUPS = [
  {
    label: 'Gastos, legalizaciones y reembolsos (no son costo de nómina)',
    test: /GASTOS|LEGALIZACION|REEMBOLSO|TIQUETE|CASINO|RESTAURANTE|UBER|TAXI|LICOR|GRAVAMEN|CUOTA DE MANEJO|INTERESES CORRIENTES|COMPRA|EXAMENES|VISA|SIN SOPORTES/
  },
  {
    label: 'Pólizas y seguros (para incluirlos, agrega /POLIZA|SEGUROS? DE VIDA/ en EXTRA_PAYROLL_PATTERNS)',
    test: /POLIZA|SEGUROS?/
  },
  { label: 'Indemnizaciones, liquidaciones y descuentos', test: /INDEMNIZACION|LIQUIDACION|BONUS|DESCUENTO AUTORIZADO/ },
  { label: 'Reclasificaciones de anticipos', test: /RECLASIFICACION/ }
];

// "OFIR ELIZABETH ESPAÑA LOPEZ" -> "ESPAÑA LOPEZ OFIR ELIZABETH" (solo nombres de
// 4 palabras; con otra cantidad es ambiguo y se deja como viene).
function displayName(rawName, reorder) {
  if (NAME_OVERRIDES[rawName]) return NAME_OVERRIDES[rawName];
  if (!reorder || /^(Sin nombre|SIN )/.test(rawName)) return rawName;
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
  const extraNames = new Set();
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

    const d = norm(description);
    const concept = kind.extra || kind.concept;
    if (kind.extra) extraNames.add(kind.extra);
    entries.push({
      month,
      comprobante,
      concept,
      amount,
      tercero,
      type: kind.extra ? 'extra' : SS_CONCEPTS.has(concept) ? 'ss' : 'core',
      // Una línea de salario (que no sea retroactivo) abre el bloque de un empleado.
      isSalaryStart: concept === 'SALARY' && !/RETROACTIV/.test(d),
      isIntegral: concept === 'SALARY' && /INTEGRAL/.test(d),
      person: null,
      block: null,
      employee: null
    });
  }

  // --- Quién es el empleado de cada línea ---------------------------------
  // 1) Líneas con Tercero de persona: ese es el empleado.
  // 2) Líneas sin Tercero (Siigo no lo trae en algunos meses): dentro de cada
  //    comprobante, cada empleado viene en un bloque seguido que empieza en su línea
  //    de salario. Se separan por bloque para no sumar a varias personas juntas.
  const namedPersonsByMonth = new Map();
  const allPersons = new Set();
  for (const e of entries) {
    if (e.type !== 'core' || !e.tercero) continue;
    allPersons.add(e.tercero);
    if (!namedPersonsByMonth.has(e.month)) namedPersonsByMonth.set(e.month, new Set());
    namedPersonsByMonth.get(e.month).add(e.tercero);
  }
  const fallbackName = allPersons.size === 1 ? [...allPersons][0] : null;

  const blocks = [];
  let lastComprobante = null;
  let current = null;
  for (const e of entries) {
    if (e.type === 'ss') continue;
    if (e.comprobante !== lastComprobante) {
      lastComprobante = e.comprobante;
      current = null;
    }
    const persons = namedPersonsByMonth.get(e.month);
    const isNamed = e.tercero && (e.type === 'core' || (persons && persons.has(e.tercero)));
    if (isNamed) {
      e.person = e.tercero;
      current = { named: e.tercero };
      continue;
    }
    if (e.isSalaryStart || !current) {
      current = { month: e.month, salary: e.isSalaryStart ? e.amount : null, name: null, label: null };
      blocks.push(current);
    }
    if (current.named) e.person = current.named;
    else e.block = current;
  }

  const blocksByMonth = new Map();
  for (const b of blocks) {
    if (!blocksByMonth.has(b.month)) blocksByMonth.set(b.month, []);
    blocksByMonth.get(b.month).push(b);
  }

  // ¿El archivo trae varios empleados a la vez? Si nunca los hay (caso de un solo
  // empleado), todo lo que no trae tercero es de la única persona con nombre.
  let multiEmployeeFile = false;
  for (const m of new Set([...blocksByMonth.keys(), ...namedPersonsByMonth.keys()])) {
    const n = (blocksByMonth.get(m)?.length || 0) + (namedPersonsByMonth.get(m)?.size || 0);
    if (n > 1) multiEmployeeFile = true;
  }

  // Si el salario de un bloque sin nombre es idéntico al de una persona con nombre
  // en otro mes, es evidencia suficiente para darle su nombre.
  const salaryToPeople = new Map();
  for (const e of entries) {
    if (e.type !== 'core' || !e.isSalaryStart || !e.person) continue;
    if (!salaryToPeople.has(e.amount)) salaryToPeople.set(e.amount, new Set());
    salaryToPeople.get(e.amount).add(e.person);
  }

  let unnamedBlocks = 0;
  for (const [month, list] of blocksByMonth) {
    const usedNames = new Set(namedPersonsByMonth.get(month) || []);
    for (const b of list) {
      if (!multiEmployeeFile) {
        b.name = fallbackName || 'SIN TERCERO';
        continue;
      }
      const people = b.salary != null ? salaryToPeople.get(b.salary) : null;
      if (people && people.size === 1 && !usedNames.has([...people][0])) {
        b.name = [...people][0];
        usedNames.add(b.name);
      } else {
        b.label = b.salary != null ? `Sin nombre - salario ${formatMoney(b.salary)}` : 'Sin nombre - sin salario';
      }
    }
    // Etiquetas repetidas en el mismo mes (dos personas con el mismo salario): numerarlas.
    const counts = new Map();
    list.forEach((b) => b.label && counts.set(b.label, (counts.get(b.label) || 0) + 1));
    const seen = new Map();
    for (const b of list) {
      if (!b.label) continue;
      unnamedBlocks += 1;
      if (counts.get(b.label) > 1) {
        const n = (seen.get(b.label) || 0) + 1;
        seen.set(b.label, n);
        b.name = `${b.label} (#${n})`;
      } else {
        b.name = b.label;
      }
    }
  }

  const employeesByMonth = new Map(); // mes -> Set(empleados con nómina ese mes)
  for (const e of entries) {
    if (e.type === 'ss') continue;
    e.employee = e.person || e.block.name;
    if (!employeesByMonth.has(e.month)) employeesByMonth.set(e.month, new Set());
    employeesByMonth.get(e.month).add(e.employee);
  }

  // 3) Aportes (pensión, salud, cajas, ARL, SENA, ICBF): Siigo los registra a nombre
  //    del fondo, no de la persona. Con un solo empleado ese mes se le asignan; con
  //    varios no se pueden repartir desde Siigo y quedan aparte.
  const ambiguousMonths = new Set();
  for (const e of entries) {
    if (e.type !== 'ss') continue;
    const set = employeesByMonth.get(e.month);
    if (set && set.size === 1) e.employee = [...set][0];
    else if (set && set.size > 1) {
      e.employee = 'SIN ASIGNAR (aportes)';
      ambiguousMonths.add(e.month);
    } else e.employee = fallbackName || 'SIN TERCERO';
  }

  // --- Agrupar por mes y empleado ---------------------------------------------
  const buckets = new Map(); // `${mes}|${empleado}` -> Map(concepto -> valor)
  const integralByKey = new Map();
  for (const e of entries) {
    const key = `${e.month}|${e.employee}`;
    if (!buckets.has(key)) buckets.set(key, new Map());
    const m = buckets.get(key);
    m.set(e.concept, (m.get(e.concept) || 0) + e.amount);
    if (e.isIntegral) integralByKey.set(key, (integralByKey.get(key) || 0) + e.amount);
  }

  // Primer mes con datos de cada empleado en cada año (para el interés acumulado).
  const firstMonthOfYear = new Map(); // `${empleado}|${año}` -> número de mes mínimo
  for (const key of buckets.keys()) {
    const [month, employee] = key.split('|');
    const [year, mm] = month.split('-').map(Number);
    const k = `${employee}|${year}`;
    if (!firstMonthOfYear.has(k) || mm < firstMonthOfYear.get(k)) firstMonthOfYear.set(k, mm);
  }

  const extraSorted = [...extraNames].sort((a, b) => a.localeCompare(b, 'es'));
  const conceptOrder = [...HEAD_ORDER, ...extraSorted, ...TAIL_ORDER];
  const provisionKeys = new Set(['13TH SALARY', '14TH SALARY', 'INTEREST ON 14TH SALARY']);
  let integralWithoutProvisions = false;

  const records = [];
  const sortedKeys = [...buckets.keys()].sort();
  for (const key of sortedKeys) {
    const [month, employee] = key.split('|');
    const [year, mm] = month.split('-').map(Number);
    const concepts = new Map(buckets.get(key));
    const monthDate = new Date(Date.UTC(year, mm - 1, 1));
    const shownName = displayName(employee, reorderNames);

    // Provisiones (calculadas). Base = salario + transporte + vacaciones + licencia;
    // el auxilio extralegal no salarial y los extra no entran. El salario integral
    // tampoco entra (ya incluye las prestaciones), salvo que se configure lo contrario.
    const integral = integralByKey.get(key) || 0;
    let base = PAYROLL_BASE_CONCEPTS.reduce((s, c) => s + (concepts.get(c) || 0), 0);
    if (!PROVISIONS_ON_INTEGRAL_SALARY) {
      base -= integral;
      if (integral > 0) integralWithoutProvisions = true;
    }
    // Sin salario ordinario ese mes (p. ej. solo vacaciones de una liquidación) no hay provisión.
    const hasOrdinarySalary = (concepts.get('SALARY') || 0) - (PROVISIONS_ON_INTEGRAL_SALARY ? 0 : integral) > 0;
    if (base > 0 && hasOrdinarySalary) {
      const monthly = Math.round(base / 12);
      const monthsSinceStart = mm - firstMonthOfYear.get(`${employee}|${year}`) + 1;
      concepts.set('13TH SALARY', monthly);
      concepts.set('14TH SALARY', monthly);
      concepts.set('INTEREST ON 14TH SALARY', Math.round((monthly * (2 * monthsSinceStart - 1)) / 100));
    }

    let total = 0;
    for (const concept of conceptOrder) {
      const value = concepts.get(concept);
      if (!value) continue;
      total += value;
      const label =
        concept === 'ALLOWANCE'
          ? ALLOWANCE_NAME_BY_MONTH[month] || DEFAULT_ALLOWANCE_NAME
          : concept;
      let fill = 'green';
      if (provisionKeys.has(concept)) fill = 'yellow';
      else if (extraNames.has(concept)) fill = 'blue';
      records.push({
        'Mes elaboración': monthDate,
        Concepto: label,
        Empleado: shownName,
        'Valor Concepto': value,
        'Valor Totales': 0,
        _fill: fill
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

  // --- Notas para el usuario (cortas: lo no incluido va agrupado por tipo) ------
  if (records.length > 0) {
    notes.push({
      type: 'info',
      text: '13TH SALARY, 14TH SALARY e INTEREST ON 14TH SALARY no existen en Siigo: se calculan (amarillo). Revísalos contra tu criterio.'
    });
  }
  if (integralWithoutProvisions) {
    notes.push({
      type: 'info',
      text: 'Salario integral: no se calculan 13TH/14TH/interés sobre él porque el integral ya incluye las prestaciones. Para calcularlos igual, pon PROVISIONS_ON_INTEGRAL_SALARY = true en siigoConverter.js.'
    });
  }
  if (unnamedBlocks > 0) {
    notes.push({
      type: 'warn',
      text: `En varios meses Siigo no trae el nombre del empleado (Tercero vacío) y hay más de uno a la vez: se separaron por bloque de salario y salen como "Sin nombre - salario X". Ponles nombre en el panel "Empleados sin nombre" (abajo, junto al resultado).`
    });
  }
  if (extraNames.size > 0) {
    notes.push({
      type: 'info',
      text: `Conceptos tomados de Siigo con su nombre original (azul): ${extraSorted.join(', ')}.`
    });
  }
  for (const [reason, e] of excluded) {
    notes.push({
      type: 'info',
      text: `Se omitieron ${e.count} línea(s) de "${reason}" (${formatMoney(e.total)}): ya las cubren las provisiones.`
    });
  }

  const groupTotals = new Map(); // etiqueta -> { total, count }
  const otherLines = [];
  for (const [desc, u] of unmapped) {
    const d = norm(desc);
    const group = UNMAPPED_GROUPS.find((g) => g.test.test(d));
    if (group) {
      const g = groupTotals.get(group.label) || { total: 0, count: 0 };
      g.total += u.total;
      g.count += u.count;
      groupTotals.set(group.label, g);
    } else {
      otherLines.push({ desc, ...u });
    }
  }
  for (const [label, g] of groupTotals) {
    notes.push({
      type: 'warn',
      text: `No incluido — ${label}: ${g.count} línea(s), ${formatMoney(g.total)}.`
    });
  }
  otherLines.sort((a, b) => Math.abs(b.total) - Math.abs(a.total));
  for (const o of otherLines.slice(0, 8)) {
    notes.push({
      type: 'warn',
      text: `Sin mapear: "${o.desc.slice(0, 70)}" — ${o.count} línea(s), ${formatMoney(o.total)}. No está en el resultado.`
    });
  }
  if (otherLines.length > 8) {
    const rest = otherLines.slice(8);
    notes.push({
      type: 'warn',
      text: `Sin mapear: otros ${rest.length} conceptos más (${formatMoney(rest.reduce((s, o) => s + o.total, 0))}).`
    });
  }
  if (ambiguousMonths.size > 0) {
    notes.push({
      type: 'warn',
      text: `En ${ambiguousMonths.size} mes(es) hay varios empleados a la vez (${[...ambiguousMonths].sort()[0]} y otros): Siigo registra los aportes (pensión, salud, cajas, ARL, SENA, ICBF) a nombre del fondo y no se pueden repartir por persona. Salen juntos como "SIN ASIGNAR (aportes)".`
    });
  }
  if (skippedNoDate > 0) {
    notes.push({ type: 'warn', text: `${skippedNoDate} línea(s) con fecha ilegible se omitieron.` });
  }

  return { records, notes };
}