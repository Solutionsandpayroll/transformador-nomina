// ============================================================================
// CONVERTIDOR: Nómina (formato ancho, un bloque por mes) -> formato largo
// ============================================================================
// Entrada: las filas de la hoja de la nómina, como array de arrays donde cada
// celda es { v: valor, f: color de relleno 'RRGGBB' | 'theme:N' | null } (o null
// si la celda no existe). Lo produce readWorkbookSheetsWithJSZip en App.jsx.
//
// Salida: una fila por (mes, empleado, concepto) con las columnas de la Hoja2:
//   Mes elaboración | Concepto | Empleado | Valor Concepto | Valor Totales
// más una marca interna `_fill` ('RRGGBB' o null) con el color que ya traía
// la celda en la nómina (el resultado del cruce con Siigo que hace el equipo).
//
// Cómo se lee la nómina (todas las empresas la arman distinto, así que NO hay
// columnas fijas):
//   - Cada bloque mensual empieza con una fila de encabezado que trae un
//     encabezado de "código de empleado" y uno de "nombre" (ver
//     CODE_HEADER_ALIASES / NAME_HEADER_ALIASES más abajo — RemoFirst usa
//     "EMPLOYEE CODE" / "NAME", pero cada empresa nueva puede traer otra
//     redacción; agrega el alias ahí en vez de tocar la lógica). Los
//     conceptos son los encabezados de las columnas que están a la derecha
//     de la columna de nombre, en el mismo orden en que aparecen.
//   - El mes sale del rótulo de la columna A ("Junio 2026", "MAYO"...). Si el
//     rótulo no trae año se deduce por la secuencia de bloques.
//   - Se ignoran los subtotales (PAYMENTS, TOTAL...), las columnas en USD, FEE y
//     tasa de cambio, y las filas de control que Excel repite debajo de cada empleado.
//   - Cada empleado de cada bloque lleva su propio TOTAL EMPLOYEE COST, así dos
//     corridas de nómina del mismo mes (p. ej. una con bonificación aparte) salen
//     separadas, igual que en el ejemplo.
//   - Los valores en cero no se listan. El color de cada valor es el relleno que
//     tenía la celda en la nómina.

// --- Configuración editable ---------------------------------------------------

// Colores que el equipo usa para marcar el cruce con Siigo (leyenda de las hojas).
export const STATUS_COLORS = {
  '92D050': 'Cruce ok',
  '00B0F0': 'Cruce ok (azul)',
  FFFF00: 'No está en el otro lado',
  FF0000: 'Diferencias',
  '00FF00': 'Cruza entre meses',
  '7030A0': 'Débito - Crédito se anulan',
  '00FFFF': 'Otro (sin leyenda)'
};

// Rellenos que son solo formato de la hoja (encabezados, filas de control, USD).
const STRUCTURAL_FILLS = new Set(['DCFAFA', 'D6E3BC', '1F4763', 'FFFFFF', '000000']);

// Encabezados que NO son un concepto de costo.
const NON_CONCEPT_EXACT = new Set([
  'PAYMENTS',
  'EE RF WID',
  'ONBOARDING DATE',
  'OFFBOARDING DATE',
  'COUNTRY',
  'PAYROLL MONTH',
  'SERVICE TYPE / INVOICE TYPE',
  'ER SS RATE %'
]);

// Filas que traen algo en la columna NAME pero no son empleados.
const NOT_AN_EMPLOYEE = /^(TOTAL|NOMINA|CONTABILIDAD|NOVEDAD|NOVADADES|DIFERENCIA|CRUCE|NO ESTA)/;

// Rótulos del resumen del cruce que hay al final de la hoja (columna A).
const SUMMARY_LABEL = /^(CRUCE|NO ESTA EN EL OTRO|DIFERENCIAS|CRUZA ENTRE|DEBITO - CREDITO)/;

// Encabezados que identifican la columna de "código de empleado" y la de
// "nombre" en la fila de encabezado de cada bloque mensual. RemoFirst usa el
// texto en inglés ("EMPLOYEE CODE" / "NAME"); a medida que llegan archivos de
// otras empresas, agrega aquí la variante exacta que traigan (en mayúsculas y
// sin tildes, que es como los deja `norm()`) — no hace falta tocar el resto
// del código. Si al convertir un archivo nuevo el resultado sale vacío, lo
// primero que hay que revisar es si su encabezado real está en estas listas.
const DEFAULT_CODE_HEADER_ALIASES = [
  'EMPLOYEE CODE',
  'CODIGO',
  'CODIGO EMPLEADO',
  'COD EMPLEADO',
  'ID EMPLEADO',
  'CEDULA',
  'DOCUMENTO',
  'NO. IDENTIFICACION',
  'IDENTIFICACION'
];
const DEFAULT_NAME_HEADER_ALIASES = [
  'NAME',
  'NOMBRE',
  'NOMBRE EMPLEADO',
  'EMPLEADO',
  'NOMBRE COMPLETO',
  'NOMBRE DEL EMPLEADO',
  'NOMBRES Y APELLIDOS'
];

const MONTHS = [
  ['ENERO', 1],
  ['FEBRERO', 2],
  ['MARZO', 3],
  ['ABRIL', 4],
  ['MAYO', 5],
  ['JUNIO', 6],
  ['JULIO', 7],
  ['AGOSTO', 8],
  ['SEPTIEMBRE', 9],
  ['SETIEMBRE', 9],
  ['OCTUBRE', 10],
  ['NOVIEMBRE', 11],
  ['DICIEMBRE', 12]
];

// Cuando un empleado no se pudo determinar (p. ej. un aporte patronal sin
// ningún salario en la misma fecha), se muestra con esta etiqueta en vez de
// dejar la celda vacía, para que no se pierda de vista en la Hoja2.
const SIN_EMPLEADO = '(sin asignar)';

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

function clean(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function toNumber(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const n = Number(String(value).replace(/,/g, ''));
  return Number.isFinite(n) ? n : 0;
}

function round2(n) {
  return Math.round(n * 100) / 100;
}

function formatMoney(n) {
  return Math.round(n).toLocaleString('es-CO');
}

function cellValue(cell) {
  return cell ? cell.v : null;
}

function cellFill(cell) {
  return cell ? cell.f || null : null;
}

// "Junio 2026" -> {month: 6, year: 2026}; "MAYO " -> {month: 5, year: null}.
function parseMonthLabel(text) {
  if (typeof text !== 'string') return null;
  const t = norm(text);
  if (!t || t.length > 30) return null;
  for (const [name, num] of MONTHS) {
    if (t.includes(name)) {
      const y = /(20\d{2})/.exec(t);
      return { month: num, year: y ? Number(y[1]) : null };
    }
  }
  return null;
}

// ¿Esta fila es el encabezado de un bloque? Debe traer una columna de código
// de empleado y una de nombre (ver codeAliases / nameAliases).
function readHeader(row, codeAliases, nameAliases) {
  let codeCol = -1;
  let nameCol = -1;
  for (let c = 0; c < row.length; c++) {
    const t = norm(cellValue(row[c]));
    if (codeCol < 0 && codeAliases.includes(t)) codeCol = c;
    else if (nameCol < 0 && nameAliases.includes(t)) nameCol = c;
  }
  if (codeCol < 0 || nameCol < 0) return null;

  // Encabezados a la derecha de NAME.
  const leaves = [];
  for (let c = nameCol + 1; c < row.length; c++) {
    const v = cellValue(row[c]);
    if (typeof v !== 'string' || !clean(v)) continue;
    leaves.push({ idx: c, label: clean(v), upper: norm(v) });
  }

  const tec = leaves.find((l) => l.upper === 'TOTAL EMPLOYEE COST');
  const payments = leaves.find((l) => l.upper === 'PAYMENTS');
  const limit = tec ? tec.idx : Infinity; // lo que va después (FEE, USD...) no es costo

  let concepts = leaves.filter((l) => {
    if (l.idx >= limit) return false;
    if (l.upper.startsWith('TOTAL')) return false; // subtotales
    if (NON_CONCEPT_EXACT.has(l.upper)) return false;
    if (l.upper.startsWith('EE STATUS')) return false;
    if (l.upper.startsWith('FEE') || l.upper.includes('EXCHANGE') || l.upper.includes('USD')) return false;
    return true;
  });

  // Algunas nóminas viejas traen SALARY (informativo) junto a INTEGRATED SALARY y
  // ORDINARY SALARY, que son las que sí suman en PAYMENTS: no se cuenta dos veces.
  const hasBreakdown = concepts.some((l) => l.upper === 'INTEGRATED SALARY' || l.upper === 'ORDINARY SALARY');
  if (hasBreakdown) concepts = concepts.filter((l) => l.upper !== 'SALARY');

  if (concepts.length === 0) return null;
  return {
    codeCol,
    nameCol,
    concepts,
    tecCol: tec ? tec.idx : null,
    paymentsCol: payments ? payments.idx : null
  };
}

function statusFill(fill) {
  if (!fill) return null;
  return STATUS_COLORS[fill] ? fill : null;
}

// --- Conversión principal -----------------------------------------------------

export function convertNominaRows(rows, options = {}) {
  const unifyNames = options.unifyNamesByCode !== false;
  const defaultYear = options.defaultYear || new Date().getFullYear();

  // Alias de encabezado: los de por defecto + los que se agreguen por empresa
  // (options.extraCodeHeaderAliases / extraNameHeaderAliases), o una lista
  // completa propia si se pasa options.codeHeaderAliases / nameHeaderAliases.
  const codeAliases = (
    options.codeHeaderAliases || [...DEFAULT_CODE_HEADER_ALIASES, ...(options.extraCodeHeaderAliases || [])]
  ).map(norm);
  const nameAliases = (
    options.nameHeaderAliases || [...DEFAULT_NAME_HEADER_ALIASES, ...(options.extraNameHeaderAliases || [])]
  ).map(norm);

  const labelEvents = []; // rótulos de mes encontrados, en orden
  const entries = []; // empleado-bloque
  const blockInfo = new Map(); // blockId -> { employees: n, checkFill }
  let header = null;
  let blockId = 0;
  let curLabel = -1;

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];

    // El resumen del cruce (Cruce ok, Diferencias...) va al final de la hoja: ahí terminan los bloques.
    if (SUMMARY_LABEL.test(norm(cellValue(row[0])))) {
      header = null;
      continue;
    }

    const label = parseMonthLabel(cellValue(row[0]));
    if (label) {
      labelEvents.push({ ...label });
      curLabel = labelEvents.length - 1;
    }

    const h = readHeader(row, codeAliases, nameAliases);
    if (h) {
      header = h;
      blockId += 1;
      blockInfo.set(blockId, { employees: 0, checkFill: null });
      continue;
    }
    if (!header) continue;

    const nameRaw = cellValue(row[header.nameCol]);
    const name = typeof nameRaw === 'string' ? clean(nameRaw) : '';
    const code = clean(cellValue(row[header.codeCol]));
    const hasNumbers = header.concepts.some((c) => typeof cellValue(row[c.idx]) === 'number');
    const isEmployee = name && !name.startsWith('#') && !/^[\d.,\s-]+$/.test(name) && !NOT_AN_EMPLOYEE.test(norm(name)) && (code || hasNumbers);

    if (!isEmployee) {
      // Fila de control (sin nombre) con el total del bloque: guarda su color.
      const info = blockInfo.get(blockId);
      if (info && header.tecCol !== null && !info.checkFill && typeof cellValue(row[header.tecCol]) === 'number') {
        info.checkFill = statusFill(cellFill(row[header.tecCol]));
      }
      continue;
    }

    blockInfo.get(blockId).employees += 1;
    entries.push({ blockId, labelIdx: curLabel, header, rowNum: r + 1, row, code, name });
  }

  if (entries.length === 0) return null;

  // --- Año de los rótulos sin año -----------------------------------------------
  // Los bloques van en orden cronológico: un rótulo sin año toma el año del bloque
  // con año más cercano (hacia atrás si viene antes, hacia adelante si viene después).
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

  // --- Nombre único por código de empleado ---------------------------------------
  // La misma persona aparece con el nombre en distinto orden según el mes
  // ("RAMON ESTEBAN CARDONA SALAZAR" / "CARDONA SALAZAR RAMON ESTEBAN"). Se usa el
  // de su bloque más reciente, que es el formato de la Hoja2.
  const latestName = new Map();
  const latestRank = new Map();
  for (const e of entries) {
    if (!e.code || e.labelIdx < 0) continue;
    const ev = labelEvents[e.labelIdx];
    const rank = ev.year * 12 + ev.month;
    if (!latestRank.has(e.code) || rank >= latestRank.get(e.code)) {
      latestRank.set(e.code, rank);
      latestName.set(e.code, e.name);
    }
  }

  // --- Armar las filas largas ------------------------------------------------------
  const records = [];
  const notes = [];
  const colorCount = new Map();
  const unknownFills = new Map();
  let uncolored = 0;
  let paymentMismatch = 0;
  let totalMismatch = [];
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
    months.add(`${ev.year}-${String(ev.month).padStart(2, '0')}`);

    let shownName = e.name;
    if (unifyNames && e.code && latestName.get(e.code) && latestName.get(e.code) !== e.name) {
      shownName = latestName.get(e.code);
      unifiedNames += 1;
    }

    const { concepts, tecCol, paymentsCol } = e.header;
    let total = 0;
    let payrollSum = 0;
    const rowsOut = [];
    for (const c of concepts) {
      const cell = e.row[c.idx];
      const value = toNumber(cellValue(cell));
      if (paymentsCol !== null && c.idx < paymentsCol) payrollSum += value;
      if (value === 0) continue;
      total += value;
      const rawFill = cellFill(cell);
      let fill = statusFill(rawFill);
      if (!fill && rawFill && /^[0-9A-F]{6}$/.test(rawFill) && !STRUCTURAL_FILLS.has(rawFill)) {
        unknownFills.set(rawFill, (unknownFills.get(rawFill) || 0) + 1);
      }
      rowsOut.push({ label: c.label, value, fill });
    }

    // Comprobaciones contra los totales que trae la propia nómina.
    if (paymentsCol !== null) {
      const paid = toNumber(cellValue(e.row[paymentsCol]));
      if (Math.abs(paid - payrollSum) > 1) paymentMismatch += 1;
    }
    total = round2(total);
    if (tecCol !== null) {
      const sheetTotal = toNumber(cellValue(e.row[tecCol]));
      if (sheetTotal !== 0 && Math.abs(sheetTotal - total) > 1) {
        totalMismatch.push({ month: `${ev.year}-${String(ev.month).padStart(2, '0')}`, name: shownName, diff: total - sheetTotal });
      }
    }

    for (const o of rowsOut) {
      records.push({
        'Mes elaboración': monthDate,
        Concepto: o.label,
        Empleado: shownName,
        'Valor Concepto': o.value,
        'Valor Totales': 0,
        _fill: o.fill
      });
      if (o.fill) colorCount.set(o.fill, (colorCount.get(o.fill) || 0) + 1);
      else uncolored += 1;
    }

    if (total !== 0) {
      // El color del total: el de su propia celda o, si el bloque tiene un solo
      // empleado, el de la fila de control (ahí es donde el equipo lo pinta).
      let totalFill = tecCol !== null ? statusFill(cellFill(e.row[tecCol])) : null;
      const info = blockInfo.get(e.blockId);
      if (!totalFill && info && info.employees === 1) totalFill = info.checkFill;
      records.push({
        'Mes elaboración': monthDate,
        Concepto: 'TOTAL EMPLOYEE COST',
        Empleado: shownName,
        'Valor Concepto': 0,
        'Valor Totales': total,
        _fill: totalFill
      });
      if (totalFill) colorCount.set(totalFill, (colorCount.get(totalFill) || 0) + 1);
      else uncolored += 1;
    }
  }

  // --- Avisos ---------------------------------------------------------------------
  const sortedMonths = [...months].sort();
  notes.push({
    type: 'info',
    text: `Se leyeron ${entries.length} empleado(s)-mes de ${sortedMonths.length} mes(es)${
      sortedMonths.length ? ` (${sortedMonths[0]} a ${sortedMonths[sortedMonths.length - 1]})` : ''
    }.`
  });
  if (colorCount.size > 0 || uncolored > 0) {
    const parts = [...colorCount.entries()]
      .sort((a, b) => b[1] - a[1])
      .map(([hex, n]) => `${STATUS_COLORS[hex]}: ${n}`);
    notes.push({
      type: 'info',
      text: `Colores tomados de la nómina — ${parts.length ? parts.join(', ') : 'ninguno'}${
        uncolored > 0 ? `; sin color (la celda no estaba pintada en la nómina): ${uncolored}` : ''
      }.`
    });
  }
  if (unifiedNames > 0) {
    notes.push({
      type: 'info',
      text: `Se unificó el nombre en ${unifiedNames} fila(s) usando el nombre más reciente de cada código de empleado (en la nómina el orden de nombre y apellido cambia de un mes a otro).`
    });
  }
  if (labelEvents.some((ev) => ev.inferred)) {
    notes.push({
      type: noYearAtAll ? 'warn' : 'info',
      text: noYearAtAll
        ? `Ningún rótulo de mes trae el año; se usó ${defaultYear}. Revisa la columna Mes elaboración.`
        : 'Algunos rótulos de mes no traen el año (p. ej. "MAYO"); se dedujo por el orden de los bloques.'
    });
  }
  if (paymentMismatch > 0) {
    notes.push({
      type: 'warn',
      text: `${paymentMismatch} fila(s) donde los conceptos de nómina no suman el PAYMENTS de la hoja: revisa que no falte o sobre una columna en ese bloque.`
    });
  }
  if (totalMismatch.length > 0) {
    const ex = totalMismatch
      .slice(0, 4)
      .map((m) => `${m.month} ${m.name} (${m.diff > 0 ? '+' : ''}${formatMoney(m.diff)})`)
      .join('; ');
    notes.push({
      type: 'warn',
      text: `${totalMismatch.length} fila(s) donde el TOTAL EMPLOYEE COST calculado no coincide con el de la hoja. Ejemplos: ${ex}.`
    });
  }
  if (unknownFills.size > 0) {
    const list = [...unknownFills.entries()].map(([hex, n]) => `#${hex} (${n})`).join(', ');
    notes.push({
      type: 'warn',
      text: `Colores de la nómina que no están en la leyenda del cruce y se dejaron sin color: ${list}.`
    });
  }
  if (noLabelRows > 0) {
    notes.push({
      type: 'warn',
      text: `${noLabelRows} empleado(s) aparecen antes de cualquier rótulo de mes y se omitieron.`
    });
  }

  return { records, notes };
}

// --- Consolidar varias nóminas (varios archivos del mismo mes o de varios
// meses) en un solo resultado ---------------------------------------------------
// RemoFirst, por ejemplo, no entrega un único archivo general por mes sino
// 4, 5 o 6 archivos distintos según el mes. Esta función junta las filas de
// todos los archivos de una misma empresa (en el orden en que los subas —
// idealmente cronológico) y llama a convertNominaRows una sola vez sobre el
// conjunto completo, para que el resultado quede consolidado en un solo
// archivo en vez de uno por archivo de entrada.
//
// Uso desde App.jsx: en vez de llamar convertNominaRows por cada hoja leída,
// junta las hojas de una misma empresa en `files` y llama a esta función una
// sola vez; el resultado (records + notes) es el que se exporta a Excel.
export function convertNominaFiles(files, options = {}) {
  // files: [{ rows, name? }, ...] — `rows` es el mismo array de arrays que
  // recibe convertNominaRows; `name` es opcional, solo para los avisos.
  const usable = (files || []).filter((f) => f && Array.isArray(f.rows) && f.rows.length > 0);
  if (usable.length === 0) return null;

  const allRows = [];
  for (const f of usable) {
    allRows.push(...f.rows);
  }

  const result = convertNominaRows(allRows, options);
  if (!result) return null;

  const names = usable.map((f, i) => f.name || `archivo ${i + 1}`).join(', ');
  result.notes.unshift({
    type: 'info',
    text: `Se consolidaron ${usable.length} archivo(s) de entrada en un solo resultado (${names}).`
  });
  return result;
}

// ============================================================================
// CONVERTIDOR: Movimiento CC (extracto contable de Siigo) -> resumen tipo nómina
// ============================================================================
// Entrada: la hoja del movimiento de cuenta contable, con las columnas típicas
// de Siigo: Comprobante | Fecha elaboración | Descripción | Tercero | Débito |
// Crédito | Saldo Movimiento (el orden de columnas puede variar entre
// empresas; se detectan por el nombre del encabezado, no por posición fija).
//
// A diferencia de la versión anterior (una fila cruda por movimiento), esta
// versión reproduce el mismo resumen "tipo nómina" que se arma a mano en
// Excel con Tabla dinámica / SUMIFS:
//
//   1. Cada Descripción se clasifica a un concepto tipo nómina (SALARY,
//      Transport allowance, PENSION COST, HEALTH COST...) usando la tabla
//      CONCEPT_KEYWORDS de abajo — el equivalente en código a la hoja
//      "Mapeo" del Excel. Es la primera palabra clave que aparece como
//      substring de la Descripción (sin tildes ni mayúsculas/minúsculas).
//      Si no coincide con ninguna, el concepto queda tal cual venía en
//      Descripción (igual que antes) y se marca como "sin clasificar" en
//      los avisos, para que edites CONCEPT_KEYWORDS y no quede escondido.
//
//   2. Empleado: en las filas de salario/prestaciones (grupo 'empleado')
//      Tercero normalmente YA es el nombre del empleado. Pero cuando ese
//      Movimiento CC viene sin Tercero diligenciado en esas filas (pasa en
//      algunos meses/archivos), y en los aportes patronales (pensión, EPS,
//      caja de compensación, ARL — grupo 'aporte'), donde Tercero es la
//      entidad (Porvenir, Sanitas...) y nunca el empleado, el empleado se
//      infiere igual en ambos casos: se busca, en la MISMA fecha de
//      elaboración, qué empleado aparece en alguna fila del grupo
//      'empleado' que sí trajo Tercero (normalmente el comprobante de
//      nómina de ese mismo cierre). Si en esa fecha hay un solo empleado
//      candidato, se le asigna. Si hay varios (empresa con más de un
//      empleado pagado el mismo día) o ninguno, la fila queda marcada como
//      "(sin asignar)" en vez de adivinar.
//
//   3. Las filas de ingreso/facturación al cliente (grupo 'excluir': "EO
//      Third parties service...", "Ingresos recibidos...") no son costo de
//      un empleado, así que no entran al resumen.
//
//   4. Se agrupa por (Mes, Concepto, Empleado) sumando "Valor Concepto"
//      (Débito - Crédito) — el equivalente a SUMIFS. Y se agrega una fila
//      TOTAL EMPLOYEE COST por (Mes, Empleado), sumando todo lo que sí se
//      pudo atribuir a ese empleado ese mes — el mismo patrón que usa
//      convertNominaRows.
//
//   5. Las filas de salida quedan ordenadas por Mes, luego Empleado y,
//      dentro de cada mes-empleado, en el orden en que los conceptos
//      aparecieron en el extracto contable, con TOTAL EMPLOYEE COST al
//      final del bloque — igual que en la Hoja2 de referencia.
//
//   6. Cuando se juntan varios archivos de Movimiento CC (ver
//      convertMovimientoFiles más abajo), cada archivo trae su propio
//      resumen del cruce al final ("CRUCE OK", "DIFERENCIAS"...); al
//      llegar a ese resumen la lectura no se detiene, solo cierra el
//      bloque actual y sigue buscando el encabezado del siguiente archivo,
//      para que ninguno de los archivos consolidados se pierda.
//
// Esta función sigue devolviendo las mismas columnas de siempre (Mes
// elaboración | Concepto | Empleado | Valor Concepto | Valor Totales), así
// que App.jsx no necesita ningún cambio.

// Tabla de clasificación — el equivalente en código a la hoja "Mapeo" del
// Excel. Se evalúa en orden, con la primera palabra clave (normalizada, sin
// tildes/mayúsculas) que aparezca dentro de la Descripción. Agrega aquí una
// fila nueva si aparece un concepto que todavía no se reconoce (los avisos
// del resultado te dicen cuáles quedaron "sin clasificar").
//
//   group: 'empleado' -> Tercero, cuando viene diligenciado, ya es el nombre
//                        del empleado; si viene vacío se infiere por fecha.
//   group: 'aporte'   -> Tercero es la entidad; el empleado se infiere por fecha.
//   group: 'excluir'  -> no es costo de nómina, se descarta (ingresos/facturación).
export const CONCEPT_KEYWORDS = [
  { keyword: 'EO THIRD PARTIES', concepto: null, group: 'excluir' },
  { keyword: 'INGRESOS RECIBIDOS', concepto: null, group: 'excluir' },

  { keyword: '001050 - SALARIO', concepto: 'SALARY', group: 'empleado' },
  { keyword: 'SUBSIDIO DE TRANSPORTE', concepto: 'Transport allowance', group: 'empleado' },
  { keyword: 'AUXILIO EXTRALEGAL', concepto: 'Allowance (No salarial)', group: 'empleado' },
  { keyword: 'DIAS HABILES EN VACACIONES', concepto: 'Vacation (días hábiles)', group: 'empleado' },
  { keyword: 'DIAS NO HABILES EN VACACIONES', concepto: 'Vacation (días no hábiles)', group: 'empleado' },
  { keyword: 'LICENCIA REMUNERADA', concepto: 'Paid leave', group: 'empleado' },
  { keyword: 'PRIMA DE SERVICIOS', concepto: '13TH SALARY', group: 'empleado' },
  { keyword: 'INTERESES CESANTIAS', concepto: 'INTEREST ON 14TH SALARY', group: 'empleado' },
  { keyword: 'CONSIGNACION CESANTIAS', concepto: '14TH SALARY', group: 'empleado' },
  { keyword: 'FPP', concepto: 'Other (People Pass / dotación)', group: 'empleado' },

  { keyword: 'APORTES A FONDOS DE', concepto: 'PENSION COST', group: 'aporte' },
  { keyword: 'APORTES A ENTIDADES PROMOTORAS DE SALUD', concepto: 'HEALTH COST', group: 'aporte' },
  { keyword: 'APORTES A CAJAS DE COMPENSACION', concepto: 'FAMILY FUND COST', group: 'aporte' },
  { keyword: 'APORTES A ADMINISTRADORAS DE RIESGOS', concepto: 'LABOR RISK COST', group: 'aporte' }
];

const MOVIMIENTO_REQUIRED_HEADERS = ['COMPROBANTE', 'FECHA ELABORACION', 'DESCRIPCION', 'DEBITO', 'CREDITO'];

const MOVIMIENTO_STOP_LABEL = /^(CRUCE OK|NO ESTA EN EL OTRO LADO|DIFERENCIAS|CRUZA ENTRE MESES|CRUZA ENTRE ANO|DEBITO - CREDITO SE ANULAN)/;

// Excel guarda las fechas como número de serie (días desde 1899-12-30). El
// lector de App.jsx no las convierte a Date (solo lee el valor crudo de la
// celda), así que se decodifican aquí antes de usarlas. También hay filas
// con la fecha como texto "dd/mm/aaaa" (capturadas a mano) — se parsean sin
// depender del locale del navegador.
function dateFromExcelSerial(value) {
  if (value instanceof Date) return value;
  if (typeof value === 'number' && Number.isFinite(value)) {
    const ms = Date.UTC(1899, 11, 30) + Math.round(value) * 86400000;
    return new Date(ms);
  }
  if (typeof value === 'string' && value.trim()) {
    const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(value.trim());
    if (m) {
      const day = Number(m[1]);
      const month = Number(m[2]);
      const year = Number(m[3]);
      const d = new Date(Date.UTC(year, month - 1, day));
      if (!Number.isNaN(d.getTime())) return d;
    }
    const parsed = new Date(value);
    if (!Number.isNaN(parsed.getTime())) return parsed;
  }
  return null;
}

function firstOfMonth(date) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), 1));
}

function monthKey(date) {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}`;
}

function dateKey(date) {
  return date.toISOString().slice(0, 10);
}

// Busca la fila de encabezado (Comprobante / Fecha elaboración / Descripción /
// Débito / Crédito), sin asumir en qué columna queda cada una.
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
    comprobanteCol: map['COMPROBANTE'],
    fechaCol: map['FECHA ELABORACION'],
    descripcionCol: map['DESCRIPCION'],
    terceroCol: map['TERCERO'] !== undefined ? map['TERCERO'] : null,
    debitoCol: map['DEBITO'],
    creditoCol: map['CREDITO']
  };
}

export function convertMovimientoRows(rows, options = {}) {
  const conceptKeywords = options.conceptKeywords || CONCEPT_KEYWORDS;
  const matchConceptWith = (descripcionNorm) => {
    for (const entry of conceptKeywords) {
      if (descripcionNorm.includes(norm(entry.keyword))) return entry;
    }
    return null;
  };

  let header = null;
  let started = false;

  // --- Pasada 1: leer cada movimiento y clasificarlo -----------------------------
  const parsedRows = [];
  const employeesByDate = new Map(); // dateKey -> Set(nombre de empleado)
  const unknownFills = new Map();
  let excludedRows = 0;
  let skippedNoDate = 0;
  let dataRows = 0;
  let unclassifiedCount = 0;
  const unclassifiedExamples = new Map();

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

    if (readMovimientoHeader(row)) continue; // encabezado repetido (varias hojas/archivos pegados)

    const comprobanteVal = clean(cellValue(row[header.comprobanteCol]));
    const descripcionVal = clean(cellValue(row[header.descripcionCol]));

    if (!comprobanteVal && MOVIMIENTO_STOP_LABEL.test(norm(descripcionVal))) {
      // Resumen del cruce al final de ESTE archivo. No se corta la lectura
      // por completo (antes hacía `break` y, al consolidar varios archivos
      // de Movimiento CC en un solo array, eso dejaba sin procesar todo lo
      // que venía después del primer archivo). En vez de eso, se cierra el
      // bloque actual y se vuelve a buscar un encabezado válido: si hay más
      // filas, serán las del siguiente archivo concatenado.
      started = false;
      header = null;
      continue;
    }
    if (!comprobanteVal && !descripcionVal) continue; // fila en blanco

    const fecha = dateFromExcelSerial(cellValue(row[header.fechaCol]));
    if (!fecha) {
      skippedNoDate += 1;
      continue;
    }
    const monthDate = firstOfMonth(fecha);
    const dKey = dateKey(fecha);

    const terceroVal = header.terceroCol !== null ? clean(cellValue(row[header.terceroCol])) : '';
    const debitoCell = row[header.debitoCol];
    const creditoCell = row[header.creditoCol];
    const debito = toNumber(cellValue(debitoCell));
    const credito = toNumber(cellValue(creditoCell));
    const value = round2(debito - credito);
    if (value === 0) continue;

    const descripcionNorm = norm(descripcionVal);
    const match = matchConceptWith(descripcionNorm);

    if (match && match.group === 'excluir') {
      excludedRows += 1;
      continue;
    }

    const group = match ? match.group : 'otro';
    const concepto = match ? match.concepto : descripcionVal;
    if (!match) {
      unclassifiedCount += 1;
      unclassifiedExamples.set(descripcionVal, (unclassifiedExamples.get(descripcionVal) || 0) + 1);
    }

    if (group === 'empleado' && terceroVal) {
      if (!employeesByDate.has(dKey)) employeesByDate.set(dKey, new Set());
      employeesByDate.get(dKey).add(terceroVal);
    }

    const rawFill = cellFill(debito !== 0 ? debitoCell : creditoCell) || cellFill(debitoCell) || cellFill(creditoCell);
    const fill = statusFill(rawFill);
    if (!fill && rawFill && /^[0-9A-F]{6}$/.test(rawFill) && !STRUCTURAL_FILLS.has(rawFill)) {
      unknownFills.set(rawFill, (unknownFills.get(rawFill) || 0) + 1);
    }

    dataRows += 1;
    parsedRows.push({ monthDate, dKey, comprobante: comprobanteVal, concepto, group, tercero: terceroVal, value, fill });
  }

  if (!header || dataRows === 0) return null;

  // --- Pasada 2: resolver el empleado ---------------------------------------------
  // Grupo 'empleado' con Tercero diligenciado: Tercero ya es el nombre, se usa tal
  // cual. Cualquier otra fila sin nombre propio (aportes patronales, o filas de
  // salario/prestación que vinieron sin Tercero en el Movimiento CC) se infiere
  // igual: se busca qué empleado aparece, en esa misma fecha, en alguna fila del
  // grupo 'empleado' que sí trajo Tercero. Si hay un único candidato, se le asigna
  // también esa fila; si hay varios o ninguno, queda "(sin asignar)".
  let ambiguousSinTercero = 0;
  let unresolvedSinTercero = 0;
  for (const pr of parsedRows) {
    if (pr.group === 'empleado' && pr.tercero) {
      pr.empleado = pr.tercero;
      continue;
    }
    const candidates = employeesByDate.get(pr.dKey);
    if (candidates && candidates.size === 1) {
      pr.empleado = [...candidates][0];
    } else if (candidates && candidates.size > 1) {
      pr.empleado = '';
      ambiguousSinTercero += 1;
    } else {
      pr.empleado = '';
      unresolvedSinTercero += 1;
    }
  }

  // --- Agrupar por Mes + Concepto + Empleado (equivalente a SUMIFS) --------------
  const grouped = new Map(); // key -> { monthDate, concepto, empleado, value, fill }
  const totalsByMonthEmployee = new Map(); // "mes|empleado" -> total
  const colorCount = new Map();
  let uncolored = 0;
  const months = new Set();

  for (const pr of parsedRows) {
    months.add(monthKey(pr.monthDate));
    const key = `${monthKey(pr.monthDate)}|${pr.concepto}|${pr.empleado}`;
    if (!grouped.has(key)) {
      grouped.set(key, { monthDate: pr.monthDate, concepto: pr.concepto, empleado: pr.empleado, value: 0, fill: pr.fill });
    }
    const g = grouped.get(key);
    g.value = round2(g.value + pr.value);
    if (!g.fill && pr.fill) g.fill = pr.fill; // conserva el primer color no nulo del grupo

    if (pr.empleado) {
      const teKey = `${monthKey(pr.monthDate)}|${pr.empleado}`;
      totalsByMonthEmployee.set(teKey, round2((totalsByMonthEmployee.get(teKey) || 0) + pr.value));
    }
  }

  const records = [];
  for (const g of grouped.values()) {
    if (g.value === 0) continue;
    records.push({
      'Mes elaboración': g.monthDate,
      Concepto: g.concepto,
      Empleado: g.empleado || SIN_EMPLEADO,
      'Valor Concepto': g.value,
      'Valor Totales': 0,
      _fill: g.fill
    });
    if (g.fill) colorCount.set(g.fill, (colorCount.get(g.fill) || 0) + 1);
    else uncolored += 1;
  }

  for (const [key, total] of totalsByMonthEmployee) {
    if (total === 0) continue;
    const [mKey, empleado] = key.split('|');
    const [y, m] = mKey.split('-').map(Number);
    records.push({
      'Mes elaboración': new Date(Date.UTC(y, m - 1, 1)),
      Concepto: 'TOTAL EMPLOYEE COST',
      Empleado: empleado || SIN_EMPLEADO,
      'Valor Concepto': 0,
      'Valor Totales': total,
      _fill: null
    });
  }

  // --- Orden de salida -------------------------------------------------------------
  // Mes -> Empleado -> conceptos en el orden en que aparecieron en el extracto,
  // con TOTAL EMPLOYEE COST al final de cada bloque mes-empleado. Así queda
  // igual a como se arma la Hoja2 a mano.
  const firstSeen = new Map(); // "mes|empleado|concepto" -> índice de aparición
  parsedRows.forEach((pr, idx) => {
    const key = `${monthKey(pr.monthDate)}|${pr.empleado || SIN_EMPLEADO}|${pr.concepto}`;
    if (!firstSeen.has(key)) firstSeen.set(key, idx);
  });

  const sortKey = (rec) => {
    const d = rec['Mes elaboración'];
    const monthNum = d.getUTCFullYear() * 12 + d.getUTCMonth();
    const isTotal = rec.Concepto === 'TOTAL EMPLOYEE COST' ? 1 : 0;
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

  // --- Avisos ---------------------------------------------------------------------
  const notes = [];
  const sortedMonths = [...months].sort();
  notes.push({
    type: 'info',
    text: `Se leyeron ${dataRows} movimiento(s) en ${sortedMonths.length} mes(es)${
      sortedMonths.length ? ` (${sortedMonths[0]} a ${sortedMonths[sortedMonths.length - 1]})` : ''
    } y se agruparon en ${grouped.size} fila(s) de concepto + ${totalsByMonthEmployee.size} de TOTAL EMPLOYEE COST.`
  });
  if (excludedRows > 0) {
    notes.push({
      type: 'info',
      text: `${excludedRows} movimiento(s) de facturación/ingresos (p. ej. "EO Third parties service...") se excluyeron por no ser costo de un empleado.`
    });
  }
  if (colorCount.size > 0 || uncolored > 0) {
    const parts = [...colorCount.entries()]
      .sort((a, b) => b[1] - a[1])
      .map(([hex, n]) => `${STATUS_COLORS[hex]}: ${n}`);
    notes.push({
      type: 'info',
      text: `Colores tomados del movimiento — ${parts.length ? parts.join(', ') : 'ninguno'}${
        uncolored > 0 ? `; sin color: ${uncolored}` : ''
      }.`
    });
  }
  if (unresolvedSinTercero > 0) {
    notes.push({
      type: 'warn',
      text: `${unresolvedSinTercero} fila(s) sin Tercero (aporte patronal, o salario/prestación que vino sin Tercero en el Movimiento CC) no se pudieron asignar a ningún empleado porque no hay ninguna fila con Tercero diligenciado en esa misma fecha; quedaron marcadas como "${SIN_EMPLEADO}".`
    });
  }
  if (ambiguousSinTercero > 0) {
    notes.push({
      type: 'warn',
      text: `${ambiguousSinTercero} fila(s) sin Tercero quedaron marcadas como "${SIN_EMPLEADO}" porque ese día hay más de un empleado candidato (empresa con varios empleados pagados la misma fecha) — revísalas a mano.`
    });
  }
  if (unclassifiedCount > 0) {
    const ex = [...unclassifiedExamples.entries()]
      .slice(0, 5)
      .map(([desc, n]) => `"${desc}" (${n})`)
      .join('; ');
    notes.push({
      type: 'warn',
      text: `${unclassifiedCount} movimiento(s) no coincidieron con ninguna palabra clave de CONCEPT_KEYWORDS y quedaron con el texto tal cual de Descripción: ${ex}. Agrégalos a CONCEPT_KEYWORDS si deben salir con el nombre tipo nómina.`
    });
  }
  if (skippedNoDate > 0) {
    notes.push({
      type: 'warn',
      text: `${skippedNoDate} fila(s) se omitieron por no tener una fecha válida en "Fecha elaboración".`
    });
  }
  if (unknownFills.size > 0) {
    const list = [...unknownFills.entries()].map(([hex, n]) => `#${hex} (${n})`).join(', ');
    notes.push({
      type: 'warn',
      text: `Colores del movimiento que no están en la leyenda del cruce y se dejaron sin color: ${list}.`
    });
  }
  notes.push({
    type: 'info',
    text: 'Este archivo se leyó como Movimiento CC (extracto contable) y se resumió como nómina: Concepto sale de CONCEPT_KEYWORDS, Empleado se toma de Tercero (o se infiere por fecha cuando falta, tanto en aportes patronales como en salario/prestaciones sin Tercero), y se agregó TOTAL EMPLOYEE COST por mes y empleado. Las filas quedan ordenadas por mes, empleado y el orden de aparición de cada concepto en el extracto.'
  });

  return { records, notes };
}

// --- Consolidar varios Movimiento CC (varios archivos, p. ej. uno por rango de
// fechas o por corrida contable) en un solo resultado -----------------------------
// Igual que convertNominaFiles: junta las filas de todos los archivos antes de
// convertir, en el orden en que los subas (idealmente cronológico), y llama a
// convertMovimientoRows una sola vez sobre el conjunto completo. Esto importa
// especialmente aquí porque la inferencia de empleado por fecha (aportes
// patronales, o filas de salario/prestación sin Tercero) necesita ver, para una
// fecha dada, todas las filas de esa fecha aunque hayan llegado en archivos
// distintos.
//
// Uso desde App.jsx: en vez de llamar convertMovimientoRows por cada archivo de
// Movimiento CC leído, junta las hojas en `files` y llama a esta función una
// sola vez; el resultado (records + notes) es el que se exporta a la Hoja2.
export function convertMovimientoFiles(files, options = {}) {
  // files: [{ rows, name? }, ...] — mismo formato de `rows` que espera
  // convertMovimientoRows; `name` es opcional, solo para el aviso de consolidación.
  const usable = (files || []).filter((f) => f && Array.isArray(f.rows) && f.rows.length > 0);
  if (usable.length === 0) return null;

  const allRows = [];
  for (const f of usable) {
    allRows.push(...f.rows);
  }

  const result = convertMovimientoRows(allRows, options);
  if (!result) return null;

  const names = usable.map((f, i) => f.name || `archivo ${i + 1}`).join(', ');
  result.notes.unshift({
    type: 'info',
    text: `Se consolidaron ${usable.length} archivo(s) de Movimiento CC en un solo resultado (${names}).`
  });
  return result;
}