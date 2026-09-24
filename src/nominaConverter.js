// ============================================================================
// CONVERTIDOR: Nómina (formato ancho, un bloque por mes) -> formato largo
// ============================================================================
//
// Entrada:
//   Filas de Excel como:
//   [
//     { v: valor, f: "RRGGBB" },
//     ...
//   ]
//
// Salida:
//   Mes elaboración | Concepto | Empleado | Valor Concepto | Valor Totales
//
// Además:
//   _fill = color original de la celda
//
// El objetivo de este archivo es NO depender de posiciones fijas de columnas.
// Cada empresa puede tener una estructura diferente.
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

const STRUCTURAL_FILLS = new Set([
  "DCFAFA",
  "D6E3BC",
  "1F4763",
  "FFFFFF",
  "000000"
]);

const NON_CONCEPT_EXACT = new Set([
  "PAYMENTS",
  "EE RF WID",
  "ONBOARDING DATE",
  "OFFBOARDING DATE",
  "COUNTRY",
  "PAYROLL MONTH",
  "SERVICE TYPE / INVOICE TYPE",
  "ER SS RATE %",
  "EMPLOYEE CODE",
  "CODE",
  "CODIGO",
  "CODIGO EMPLEADO",
  "COD EMPLEADO",
  "ID EMPLEADO",
  "CEDULA",
  "DOCUMENTO",
  "IDENTIFICACION",
  "NAME",
  "NOMBRE",
  "NOMBRE EMPLEADO",
  "EMPLEADO",
  "NOMBRE COMPLETO",
  "NOMBRE DEL EMPLEADO",
  "NOMBRES Y APELLIDOS"
]);

const NOT_AN_EMPLOYEE =
  /^(TOTAL|NOMINA|NÓMINA|CONTABILIDAD|NOVEDAD|NOVADADES|DIFERENCIA|CRUCE|NO ESTA|NO ESTÁ|SUBTOTAL|TOTAL GENERAL|GRAN TOTAL|RESUMEN|SUMMARY|CONTROL|OBSERVACION|OBSERVACIONES)/;

const SUMMARY_LABEL =
  /^(CRUCE|NO ESTA EN EL OTRO|NO ESTÁ EN EL OTRO|DIFERENCIAS|CRUZA ENTRE|DEBITO - CREDITO|DÉBITO - CRÉDITO)/;

const FILE_BREAK = "__FILE_BREAK__";

const DEFAULT_CODE_HEADER_ALIASES = [
  "EMPLOYEE CODE",
  "EMPLOYEE ID",
  "EMPLOYEE NUMBER",
  "CODE",
  "CODIGO",
  "CÓDIGO",
  "CODIGO EMPLEADO",
  "CÓDIGO EMPLEADO",
  "COD EMPLEADO",
  "COD. EMPLEADO",
  "ID EMPLEADO",
  "ID DEL EMPLEADO",
  "CEDULA",
  "CÉDULA",
  "DOCUMENTO",
  "DOCUMENTO IDENTIDAD",
  "DOCUMENTO DE IDENTIDAD",
  "NO. IDENTIFICACION",
  "NO IDENTIFICACION",
  "NO. IDENTIFICACIÓN",
  "IDENTIFICACION",
  "IDENTIFICACIÓN"
];

const DEFAULT_NAME_HEADER_ALIASES = [
  "NAME",
  "EMPLOYEE NAME",
  "NOMBRE",
  "NOMBRE EMPLEADO",
  "NOMBRE DEL EMPLEADO",
  "EMPLEADO",
  "NOMBRE COMPLETO",
  "NOMBRE DEL EMPLEADO",
  "NOMBRES Y APELLIDOS",
  "FULL NAME"
];

const MONTHS = [
  ["ENERO", 1],
  ["FEBRERO", 2],
  ["MARZO", 3],
  ["ABRIL", 4],
  ["MAYO", 5],
  ["JUNIO", 6],
  ["JULIO", 7],
  ["AGOSTO", 8],
  ["SEPTIEMBRE", 9],
  ["SETIEMBRE", 9],
  ["OCTUBRE", 10],
  ["NOVIEMBRE", 11],
  ["DICIEMBRE", 12]
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
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") {
    return 0;
  }

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  let text = String(value).trim();

  if (!text) return 0;

  // Quita símbolos monetarios.
  text = text
    .replace(/\$/g, "")
    .replace(/\s/g, "");

  // Si viene como número colombiano:
  // 1.234.567,89
  if (
    text.includes(".") &&
    text.includes(",") &&
    text.lastIndexOf(",") > text.lastIndexOf(".")
  ) {
    text = text
      .replace(/\./g, "")
      .replace(",", ".");
  } else if (
    text.includes(",") &&
    !text.includes(".")
  ) {
    // 123456,78
    text = text.replace(",", ".");
  } else {
    // 1,234,567
    text = text.replace(/,/g, "");
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
  return (
    typeof value === "string" &&
    /^[0-9A-F]{6}$/i.test(value)
  );
}

function statusFill(fill) {
  if (!fill) return null;

  const normalized = String(fill)
    .replace("#", "")
    .toUpperCase();

  return STATUS_COLORS[normalized]
    ? normalized
    : null;
}

// ============================================================================
// MESES
// ============================================================================

function parseMonthLabel(text) {
  if (
    text === null ||
    text === undefined
  ) {
    return null;
  }

  const t = norm(text);

  if (!t || t.length > 50) {
    return null;
  }

  for (const [name, num] of MONTHS) {
    if (t.includes(name)) {
      const y = /(20\d{2})/.exec(t);

      return {
        month: num,
        year: y ? Number(y[1]) : null
      };
    }
  }

  return null;
}

// Busca un mes en toda la fila.
// Primero intenta columna A, pero también revisa las demás columnas.
function findMonthLabelInRow(row) {
  if (!Array.isArray(row)) return null;

  // Primero columna A.
  const first = parseMonthLabel(
    cellValue(row[0])
  );

  if (first) {
    return first;
  }

  // Después el resto.
  for (let c = 1; c < row.length; c++) {
    const value = cellValue(row[c]);

    if (
      typeof value !== "string" ||
      !value.trim()
    ) {
      continue;
    }

    const parsed = parseMonthLabel(value);

    if (parsed) {
      return parsed;
    }
  }

  return null;
}

// ============================================================================
// DETECCIÓN DE ENCABEZADOS DE NÓMINA
// ============================================================================

function isCodeHeader(text, aliases) {
  const t = norm(text);

  if (!t) return false;

  if (aliases.includes(t)) {
    return true;
  }

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

  if (aliases.includes(t)) {
    return true;
  }

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

  if (NON_CONCEPT_EXACT.has(t)) {
    return true;
  }

  if (
    t.startsWith("EE STATUS") ||
    t.startsWith("TOTAL") ||
    t.startsWith("SUBTOTAL")
  ) {
    return true;
  }

  if (
    t.startsWith("FEE") ||
    t.includes("EXCHANGE") ||
    t.includes("USD") ||
    t.includes("RATE %")
  ) {
    return true;
  }

  return false;
}

// Detecta la fila de encabezado de un bloque.
function readHeader(
  row,
  codeAliases,
  nameAliases
) {
  let codeCol = -1;
  let nameCol = -1;

  if (!Array.isArray(row)) {
    return null;
  }

  for (
    let c = 0;
    c < row.length;
    c++
  ) {
    const value = cellValue(row[c]);
    const t = norm(value);

    if (!t) continue;

    if (
      codeCol < 0 &&
      isCodeHeader(t, codeAliases)
    ) {
      codeCol = c;
      continue;
    }

    if (
      nameCol < 0 &&
      isNameHeader(t, nameAliases)
    ) {
      nameCol = c;
    }
  }

  if (
    codeCol < 0 ||
    nameCol < 0
  ) {
    return null;
  }

  const leaves = [];

  for (
    let c = nameCol + 1;
    c < row.length;
    c++
  ) {
    const v = cellValue(row[c]);

    if (
      typeof v !== "string" ||
      !clean(v)
    ) {
      continue;
    }

    leaves.push({
      idx: c,
      label: clean(v),
      upper: norm(v)
    });
  }

  const tec = leaves.find(
    (l) =>
      l.upper === "TOTAL EMPLOYEE COST"
  );

  const payments = leaves.find(
    (l) =>
      l.upper === "PAYMENTS"
  );

  const limit = tec
    ? tec.idx
    : Infinity;

  let concepts = leaves.filter(
    (l) => {
      if (l.idx >= limit) {
        return false;
      }

      if (
        looksLikeStructuralHeader(
          l.upper
        )
      ) {
        return false;
      }

      return true;
    }
  );

  // Si existe SALARY junto con:
  // INTEGRATED SALARY / ORDINARY SALARY,
  // se evita duplicar SALARY.
  const hasBreakdown =
    concepts.some(
      (l) =>
        l.upper ===
          "INTEGRATED SALARY" ||
        l.upper ===
          "ORDINARY SALARY"
    );

  if (hasBreakdown) {
    concepts = concepts.filter(
      (l) =>
        l.upper !== "SALARY"
    );
  }

  if (
    concepts.length === 0
  ) {
    return null;
  }

  return {
    codeCol,
    nameCol,
    concepts,
    tecCol: tec
      ? tec.idx
      : null,
    paymentsCol:
      payments
        ? payments.idx
        : null
  };
}

// ============================================================================
// DETECCIÓN DE EMPLEADOS
// ============================================================================

function looksLikeEmployeeName(
  value
) {
  const name = clean(value);

  if (!name) return false;

  const upper = norm(name);

  if (
    upper.startsWith("#")
  ) {
    return false;
  }

  if (
    NOT_AN_EMPLOYEE.test(
      upper
    )
  ) {
    return false;
  }

  if (
    /^(NAME|NOMBRE|EMPLOYEE|EMPLEADO)$/i.test(
      name
    )
  ) {
    return false;
  }

  // Evitar fechas.
  if (
    /^\d{1,4}[\/-]\d{1,2}[\/-]\d{1,4}$/.test(
      name
    )
  ) {
    return false;
  }

  // Evitar números.
  if (
    /^[\d.,\s\-+$()%]+$/.test(
      name
    )
  ) {
    return false;
  }

  // Evitar textos demasiado cortos.
  if (
    name.length < 2
  ) {
    return false;
  }

  return true;
}

function rowHasNumericConcept(
  row,
  concepts
) {
  return concepts.some(
    (c) => {
      const value =
        cellValue(
          row[c.idx]
        );

      if (
        typeof value ===
        "number"
      ) {
        return Number.isFinite(
          value
        );
      }

      if (
        typeof value ===
        "string" &&
        value.trim()
      ) {
        return (
          toNumber(value) !== 0
        );
      }

      return false;
    }
  );
}

// ============================================================================
// NÓMINA -> FORMATO LARGO
// ============================================================================

export function convertNominaRows(
  rows,
  options = {}
) {
  if (
    !Array.isArray(rows) ||
    rows.length === 0
  ) {
    return null;
  }

  const unifyNames =
    options.unifyNamesByCode !== false;

  const defaultYear =
    options.defaultYear ||
    new Date().getFullYear();

  const codeAliases = (
    options.codeHeaderAliases ||
    [
      ...DEFAULT_CODE_HEADER_ALIASES,
      ...(options.extraCodeHeaderAliases ||
        [])
    ]
  ).map(norm);

  const nameAliases = (
    options.nameHeaderAliases ||
    [
      ...DEFAULT_NAME_HEADER_ALIASES,
      ...(options.extraNameHeaderAliases ||
        [])
    ]
  ).map(norm);

  const labelEvents = [];
  const entries = [];

  const blockInfo =
    new Map();

  let header = null;
  let blockId = 0;
  let curLabel = -1;

  // --------------------------------------------------------------------------
  // PASADA 1: encontrar bloques y empleados
  // --------------------------------------------------------------------------

  for (
    let r = 0;
    r < rows.length;
    r++
  ) {
    const row =
      rows[r] || [];

    // Separador entre archivos.
    if (
      norm(
        cellValue(row[0])
      ) === FILE_BREAK
    ) {
      header = null;
      curLabel = -1;
      continue;
    }

    // Resumen final.
    if (
      SUMMARY_LABEL.test(
        norm(
          cellValue(row[0])
        )
      )
    ) {
      header = null;
      continue;
    }

    // Detectar mes.
    const label =
      findMonthLabelInRow(
        row
      );

    if (label) {
      labelEvents.push({
        ...label
      });

      curLabel =
        labelEvents.length - 1;
    }

    // Detectar encabezado.
    const h =
      readHeader(
        row,
        codeAliases,
        nameAliases
      );

    if (h) {
      header = h;

      blockId += 1;

      blockInfo.set(
        blockId,
        {
          employees: 0,
          checkFill: null
        }
      );

      continue;
    }

    if (!header) {
      continue;
    }

    const nameRaw =
      cellValue(
        row[
          header.nameCol
        ]
      );

    const name =
      typeof nameRaw ===
      "string"
        ? clean(nameRaw)
        : clean(nameRaw);

    const code =
      clean(
        cellValue(
          row[
            header.codeCol
          ]
        )
      );

    const hasNumbers =
      rowHasNumericConcept(
        row,
        header.concepts
      );

    const validName =
      looksLikeEmployeeName(
        name
      );

    // Para considerar una fila empleado:
    //
    // 1. Debe tener nombre válido.
    // 2. Debe tener código o algún valor numérico
    //    en las columnas de conceptos.
    //
    // Esto evita tomar filas de títulos, subtotales
    // o información de control.
    const isEmployee =
      validName &&
      (Boolean(code) ||
        hasNumbers);

    if (!isEmployee) {
      const info =
        blockInfo.get(
          blockId
        );

      if (
        info &&
        header.tecCol !== null &&
        !info.checkFill
      ) {
        const totalCell =
          row[
            header.tecCol
          ];

        if (
          typeof cellValue(
            totalCell
          ) === "number" ||
          toNumber(
            cellValue(
              totalCell
            )
          ) !== 0
        ) {
          info.checkFill =
            statusFill(
              cellFill(
                totalCell
              )
            );
        }
      }

      continue;
    }

    const info =
      blockInfo.get(
        blockId
      );

    if (info) {
      info.employees += 1;
    }

    entries.push({
      blockId,
      labelIdx:
        curLabel,
      header,
      rowNum:
        r + 1,
      row,
      code,
      name
    });
  }

  if (
    entries.length === 0
  ) {
    return null;
  }

  // --------------------------------------------------------------------------
  // DEDUCIR AÑOS DE LOS MESES
  // --------------------------------------------------------------------------

  // Primera pasada hacia atrás.
  let refYear = null;
  let refMonth = null;

  for (
    let i =
      labelEvents.length - 1;
    i >= 0;
    i--
  ) {
    const ev =
      labelEvents[i];

    if (ev.year) {
      refYear =
        ev.year;
      refMonth =
        ev.month;
    } else if (
      refYear !== null
    ) {
      ev.year =
        ev.month <=
        refMonth
          ? refYear
          : refYear - 1;

      ev.inferred =
        true;

      refYear =
        ev.year;

      refMonth =
        ev.month;
    }
  }

  // Segunda pasada hacia adelante.
  refYear = null;
  refMonth = null;

  for (
    const ev of labelEvents
  ) {
    if (ev.year) {
      refYear =
        ev.year;
      refMonth =
        ev.month;
    } else if (
      refYear !== null
    ) {
      ev.year =
        ev.month >=
        refMonth
          ? refYear
          : refYear + 1;

      ev.inferred =
        true;

      refYear =
        ev.year;

      refMonth =
        ev.month;
    }
  }

  let noYearAtAll =
    false;

  for (
    const ev of labelEvents
  ) {
    if (!ev.year) {
      ev.year =
        defaultYear;

      ev.inferred =
        true;

      noYearAtAll =
        true;
    }
  }

  // --------------------------------------------------------------------------
  // UNIFICAR NOMBRES POR CÓDIGO
  // --------------------------------------------------------------------------

  const latestName =
    new Map();

  const latestRank =
    new Map();

  for (
    const e of entries
  ) {
    if (
      !e.code ||
      e.labelIdx < 0
    ) {
      continue;
    }

    const ev =
      labelEvents[
        e.labelIdx
      ];

    if (!ev) {
      continue;
    }

    const rank =
      ev.year * 12 +
      ev.month;

    if (
      !latestRank.has(
        e.code
      ) ||
      rank >=
        latestRank.get(
          e.code
        )
    ) {
      latestRank.set(
        e.code,
        rank
      );

      latestName.set(
        e.code,
        e.name
      );
    }
  }

  // --------------------------------------------------------------------------
  // ARMAR FORMATO LARGO
  // --------------------------------------------------------------------------

  const records = [];
  const notes = [];

  const colorCount =
    new Map();

  const unknownFills =
    new Map();

  let uncolored = 0;
  let paymentMismatch = 0;

  const totalMismatch =
    [];

  let noLabelRows = 0;
  let unifiedNames = 0;

  const months =
    new Set();

  for (
    const e of entries
  ) {
    const ev =
      e.labelIdx >= 0
        ? labelEvents[
            e.labelIdx
          ]
        : null;

    if (!ev) {
      noLabelRows += 1;
      continue;
    }

    const monthDate =
      new Date(
        Date.UTC(
          ev.year,
          ev.month - 1,
          1
        )
      );

    const monthId =
      `${ev.year}-${String(
        ev.month
      ).padStart(2, "0")}`;

    months.add(
      monthId
    );

    let shownName =
      e.name;

    if (
      unifyNames &&
      e.code &&
      latestName.get(
        e.code
      ) &&
      latestName.get(
        e.code
      ) !== e.name
    ) {
      shownName =
        latestName.get(
          e.code
        );

      unifiedNames +=
        1;
    }

    const {
      concepts,
      tecCol,
      paymentsCol
    } = e.header;

    let total = 0;
    let payrollSum = 0;

    const rowsOut = [];

    // ------------------------------------------------------------------------
    // CONCEPTOS
    // ------------------------------------------------------------------------

    for (
      const c of concepts
    ) {
      const cell =
        e.row[c.idx];

      const value =
        round2(
          toNumber(
            cellValue(
              cell
            )
          )
        );

      if (
        paymentsCol !== null &&
        c.idx < paymentsCol
      ) {
        payrollSum =
          round2(
            payrollSum +
              value
          );
      }

      // Los ceros no salen al formato largo.
      if (
        value === 0
      ) {
        continue;
      }

      total =
        round2(
          total +
            value
        );

      const rawFill =
        cellFill(
          cell
        );

      const fill =
        statusFill(
          rawFill
        );

      if (
        !fill &&
        rawFill &&
        isHexFill(
          rawFill
        ) &&
        !STRUCTURAL_FILLS.has(
          String(
            rawFill
          ).toUpperCase()
        )
      ) {
        const hex =
          String(
            rawFill
          ).toUpperCase();

        unknownFills.set(
          hex,
          (
            unknownFills.get(
              hex
            ) || 0
          ) + 1
        );
      }

      rowsOut.push({
        label:
          c.label,
        value,
        fill
      });
    }

    // ------------------------------------------------------------------------
    // PAYMENTS
    // ------------------------------------------------------------------------

    if (
      paymentsCol !== null
    ) {
      const paid =
        round2(
          toNumber(
            cellValue(
              e.row[
                paymentsCol
              ]
            )
          )
        );

      if (
        Math.abs(
          paid -
            payrollSum
        ) > 1
      ) {
        paymentMismatch +=
          1;
      }
    }

    total =
      round2(
        total
      );

    // ------------------------------------------------------------------------
    // TOTAL EMPLOYEE COST
    // ------------------------------------------------------------------------

    if (
      tecCol !== null
    ) {
      const sheetTotal =
        round2(
          toNumber(
            cellValue(
              e.row[
                tecCol
              ]
            )
          )
        );

      if (
        sheetTotal !== 0 &&
        Math.abs(
          sheetTotal -
            total
        ) > 1
      ) {
        totalMismatch.push({
          month:
            monthId,
          name:
            shownName,
          diff:
            round2(
              total -
                sheetTotal
            )
        });
      }
    }

    // ------------------------------------------------------------------------
    // RECORDS DE CONCEPTOS
    // ------------------------------------------------------------------------

    for (
      const o of rowsOut
    ) {
      records.push({
        "Mes elaboración":
          monthDate,

        Concepto:
          o.label,

        Empleado:
          shownName,

        "Valor Concepto":
          o.value,

        "Valor Totales":
          0,

        _fill:
          o.fill
      });

      if (o.fill) {
        colorCount.set(
          o.fill,
          (
            colorCount.get(
              o.fill
            ) || 0
          ) + 1
        );
      } else {
        uncolored += 1;
      }
    }

    // ------------------------------------------------------------------------
    // TOTAL EMPLOYEE COST
    // ------------------------------------------------------------------------

    if (
      total !== 0
    ) {
      let totalFill =
        tecCol !== null
          ? statusFill(
              cellFill(
                e.row[
                  tecCol
                ]
              )
            )
          : null;

      const info =
        blockInfo.get(
          e.blockId
        );

      if (
        !totalFill &&
        info &&
        info.employees === 1
      ) {
        totalFill =
          info.checkFill;
      }

      records.push({
        "Mes elaboración":
          monthDate,

        Concepto:
          "TOTAL EMPLOYEE COST",

        Empleado:
          shownName,

        "Valor Concepto":
          0,

        "Valor Totales":
          total,

        _fill:
          totalFill
      });

      if (totalFill) {
        colorCount.set(
          totalFill,
          (
            colorCount.get(
              totalFill
            ) || 0
          ) + 1
        );
      } else {
        uncolored += 1;
      }
    }
  }

  // --------------------------------------------------------------------------
  // AVISOS
  // --------------------------------------------------------------------------

  const sortedMonths =
    [...months].sort();

  notes.push({
    type: "info",
    text:
      `Se leyeron ${entries.length} empleado(s)-mes de ${sortedMonths.length} mes(es)` +
      (
        sortedMonths.length
          ? ` (${sortedMonths[0]} a ${sortedMonths[sortedMonths.length - 1]})`
          : ""
      ) +
      "."
  });

  if (
    colorCount.size > 0 ||
    uncolored > 0
  ) {
    const parts =
      [...colorCount.entries()]
        .sort(
          (a, b) =>
            b[1] - a[1]
        )
        .map(
          ([hex, n]) =>
            `${STATUS_COLORS[hex] || `#${hex}`}: ${n}`
        );

    notes.push({
      type: "info",
      text:
        `Colores tomados de la nómina — ` +
        `${
          parts.length
            ? parts.join(", ")
            : "ninguno"
        }` +
        (
          uncolored > 0
            ? `; sin color: ${uncolored}`
            : ""
        ) +
        "."
    });
  }

  if (
    unifiedNames > 0
  ) {
    notes.push({
      type: "info",
      text:
        `Se unificó el nombre en ${unifiedNames} fila(s) usando el nombre más reciente de cada código de empleado.`
    });
  }

  if (
    labelEvents.some(
      (ev) =>
        ev.inferred
    )
  ) {
    notes.push({
      type:
        noYearAtAll
          ? "warn"
          : "info",

      text:
        noYearAtAll
          ? `Ningún rótulo de mes trae el año; se usó ${defaultYear}. Revisa la columna Mes elaboración.`
          : "Algunos rótulos de mes no traen el año; se dedujo por el orden de los bloques."
    });
  }

  if (
    paymentMismatch > 0
  ) {
    notes.push({
      type: "warn",
      text:
        `${paymentMismatch} fila(s) donde los conceptos de nómina no suman el PAYMENTS de la hoja.`
    });
  }

  if (
    totalMismatch.length > 0
  ) {
    const ex =
      totalMismatch
        .slice(0, 4)
        .map(
          (m) =>
            `${m.month} ${m.name} (${m.diff > 0 ? "+" : ""}${formatMoney(m.diff)})`
        )
        .join("; ");

    notes.push({
      type: "warn",
      text:
        `${totalMismatch.length} fila(s) donde el TOTAL EMPLOYEE COST calculado no coincide con el de la hoja. Ejemplos: ${ex}.`
    });
  }

  if (
    unknownFills.size > 0
  ) {
    const list =
      [...unknownFills.entries()]
        .map(
          ([hex, n]) =>
            `#${hex} (${n})`
        )
        .join(", ");

    notes.push({
      type: "warn",
      text:
        `Colores de la nómina que no están en la leyenda del cruce y se dejaron sin color: ${list}.`
    });
  }

  if (
    noLabelRows > 0
  ) {
    notes.push({
      type: "warn",
      text:
        `${noLabelRows} empleado(s) aparecen antes de cualquier rótulo de mes y se omitieron.`
    });
  }

  return {
    records,
    notes
  };
}

// ============================================================================
// CONSOLIDAR VARIAS NÓMINAS
// ============================================================================

export function convertNominaFiles(
  files,
  options = {}
) {
  const usable =
    (files || []).filter(
      (f) =>
        f &&
        Array.isArray(
          f.rows
        ) &&
        f.rows.length > 0
    );

  if (
    usable.length === 0
  ) {
    return null;
  }

  const allRows = [];

  for (
    const f of usable
  ) {
    allRows.push([
      {
        v: FILE_BREAK,
        f: null
      }
    ]);

    allRows.push(
      ...f.rows
    );
  }

  const result =
    convertNominaRows(
      allRows,
      options
    );

  if (!result) {
    return null;
  }

  result.records =
    [...result.records].sort(
      (a, b) =>
        a["Mes elaboración"] -
        b["Mes elaboración"]
    );

  const names =
    usable
      .map(
        (f, i) =>
          f.name ||
          `archivo ${i + 1}`
      )
      .join(", ");

  result.notes.unshift({
    type: "info",
    text:
      `Se consolidaron ${usable.length} archivo(s) de entrada en un solo resultado (${names}).`
  });

  return result;
}

// ============================================================================
// MOVIMIENTO CC
// ============================================================================

export const CONCEPT_KEYWORDS = [
  {
    keyword:
      "EO THIRD PARTIES",
    concepto: null,
    group: "excluir"
  },

  {
    keyword:
      "INGRESOS RECIBIDOS",
    concepto: null,
    group: "excluir"
  },

  {
    keyword:
      "001050 - SALARIO",
    concepto:
      "SALARY",
    group:
      "empleado"
  },

  {
    keyword:
      "SALARIO",
    concepto:
      "SALARY",
    group:
      "empleado"
  },

  {
    keyword:
      "SUBSIDIO DE TRANSPORTE",
    concepto:
      "Transport allowance",
    group:
      "empleado"
  },

  {
    keyword:
      "AUXILIO EXTRALEGAL",
    concepto:
      "Allowance (No salarial)",
    group:
      "empleado"
  },

  {
    keyword:
      "DIAS HABILES EN VACACIONES",
    concepto:
      "Vacation (días hábiles)",
    group:
      "empleado"
  },

  {
    keyword:
      "DIAS NO HABILES EN VACACIONES",
    concepto:
      "Vacation (días no hábiles)",
    group:
      "empleado"
  },

  {
    keyword:
      "LICENCIA REMUNERADA",
    concepto:
      "Paid leave",
    group:
      "empleado"
  },

  {
    keyword:
      "PRIMA DE SERVICIOS",
    concepto:
      "13TH SALARY",
    group:
      "empleado"
  },

  {
    keyword:
      "INTERESES CESANTIAS",
    concepto:
      "INTEREST ON 14TH SALARY",
    group:
      "empleado"
  },

  {
    keyword:
      "CONSIGNACION CESANTIAS",
    concepto:
      "14TH SALARY",
    group:
      "empleado"
  },

  {
    keyword:
      "FPP",
    concepto:
      "Other (People Pass / dotación)",
    group:
      "empleado",
    wholeWord:
      true
  },

  {
    keyword:
      "APORTES A FONDOS DE",
    concepto:
      "PENSION COST",
    group:
      "aporte"
  },

  {
    keyword:
      "APORTES A ENTIDADES PROMOTORAS DE SALUD",
    concepto:
      "HEALTH COST",
    group:
      "aporte"
  },

  {
    keyword:
      "APORTES A CAJAS DE COMPENSACION",
    concepto:
      "FAMILY FUND COST",
    group:
      "aporte"
  },

  {
    keyword:
      "APORTES A ADMINISTRADORAS DE RIESGOS",
    concepto:
      "LABOR RISK COST",
    group:
      "aporte"
  }
];

const MOVIMIENTO_REQUIRED_HEADERS = [
  "COMPROBANTE",
  "FECHA ELABORACION",
  "DESCRIPCION",
  "DEBITO",
  "CREDITO"
];

const MOVIMIENTO_STOP_LABEL =
  /^(CRUCE OK|NO ESTA EN EL OTRO LADO|DIFERENCIAS|CRUZA ENTRE MESES|CRUZA ENTRE ANO|DEBITO - CREDITO SE ANULAN)/;

// ============================================================================
// FECHAS MOVIMIENTO CC
// ============================================================================

function dateFromExcelSerial(
  value
) {
  if (
    value instanceof Date
  ) {
    return value;
  }

  if (
    typeof value ===
      "number" &&
    Number.isFinite(value)
  ) {
    const ms =
      Date.UTC(
        1899,
        11,
        30
      ) +
      Math.round(value) *
        86400000;

    return new Date(ms);
  }

  if (
    typeof value ===
      "string" &&
    value.trim()
  ) {
    const text =
      value.trim();

    // dd/mm/yyyy
    // dd-mm-yyyy
    const m =
      /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/.exec(
        text
      );

    if (m) {
      const day =
        Number(m[1]);

      const month =
        Number(m[2]);

      const year =
        Number(m[3]);

      const d =
        new Date(
          Date.UTC(
            year,
            month - 1,
            day
          )
        );

      if (
        !Number.isNaN(
          d.getTime()
        )
      ) {
        return d;
      }
    }

    const parsed =
      new Date(
        text
      );

    if (
      !Number.isNaN(
        parsed.getTime()
      )
    ) {
      return parsed;
    }
  }

  return null;
}

function firstOfMonth(
  date
) {
  return new Date(
    Date.UTC(
      date.getUTCFullYear(),
      date.getUTCMonth(),
      1
    )
  );
}

function monthKey(
  date
) {
  return `${date.getUTCFullYear()}-${String(
    date.getUTCMonth() + 1
  ).padStart(2, "0")}`;
}

function dateKey(
  date
) {
  return date
    .toISOString()
    .slice(0, 10);
}

// ============================================================================
// ENCABEZADO MOVIMIENTO CC
// ============================================================================

function readMovimientoHeader(
  row
) {
  const map = {};

  for (
    let c = 0;
    c < row.length;
    c++
  ) {
    const t =
      norm(
        cellValue(
          row[c]
        )
      );

    if (
      t &&
      map[t] === undefined
    ) {
      map[t] = c;
    }
  }

  for (
    const required of
      MOVIMIENTO_REQUIRED_HEADERS
  ) {
    if (
      map[required] ===
      undefined
    ) {
      return null;
    }
  }

  return {
    comprobanteCol:
      map[
        "COMPROBANTE"
      ],

    fechaCol:
      map[
        "FECHA ELABORACION"
      ],

    descripcionCol:
      map[
        "DESCRIPCION"
      ],

    terceroCol:
      map[
        "TERCERO"
      ] !== undefined
        ? map[
            "TERCERO"
          ]
        : null,

    debitoCol:
      map["DEBITO"],

    creditoCol:
      map["CREDITO"]
  };
}

// ============================================================================
// BUSCAR CONCEPTO EN DESCRIPCIÓN
// ============================================================================

function findMovimientoConcept(
  descripcionNorm,
  conceptKeywords
) {
  const words =
    ` ${descripcionNorm.replace(
      /[^A-Z0-9]+/g,
      " "
    )} `;

  for (
    const entry of
      conceptKeywords
  ) {
    const k =
      norm(
        entry.keyword
      );

    if (!k) continue;

    const hit =
      entry.wholeWord
        ? words.includes(
            ` ${k} `
          )
        : descripcionNorm.includes(
            k
          );

    if (hit) {
      return entry;
    }
  }

  return null;
}

// ============================================================================
// EXTRAER EMPLEADO DESDE DESCRIPCIÓN
// ============================================================================
//
// Caso real BUBBLE:
//
//   001050 - Salario OFIR ELIZABETH ESPAÑA LOPEZ
//
//   001300 - Subsidio de Transporte OFIR ELIZABETH ESPAÑA LOPEZ
//
// Tercero puede venir vacío.
// ============================================================================

function extractEmployeeFromDescription(
  descripcion,
  match
) {
  if (
    !descripcion ||
    !match ||
    match.group !==
      "empleado"
  ) {
    return "";
  }

  let text =
    clean(
      descripcion
    );

  const normalized =
    norm(text);

  const keyword =
    norm(
      match.keyword
    );

  // Buscar la palabra/frase del concepto.
  const index =
    normalized.indexOf(
      keyword
    );

  if (
    index >= 0
  ) {
    // Usamos la misma posición sobre el texto original.
    // Como norm elimina acentos pero conserva estructura,
    // la longitud suele coincidir suficientemente para este caso.
    text =
      text.slice(
        index +
          keyword.length
      );
  } else {
    // Si el keyword no se encontró exactamente,
    // intentamos eliminar el código inicial.
    text =
      text.replace(
        /^\s*\d+\s*[-:]\s*/,
        ""
      );
  }

  // Elimina separadores típicos.
  text =
    text
      .replace(
        /^\s*[-:.;]+\s*/,
        ""
      )
      .replace(
        /^\s*\.+\s*/,
        ""
      )
      .replace(
        /\s+/g,
        " "
      )
      .trim();

  // Si después de quitar el concepto no queda nada,
  // no inventamos empleado.
  if (
    !text ||
    text.length < 3
  ) {
    return "";
  }

  // Evitar resultados que claramente siguen siendo conceptos
  // o información contable.
  const upper =
    norm(text);

  if (
    /^(SALARIO|SUBSIDIO|AUXILIO|APORTES|APORTE|PENSION|SALUD|RIESGO|PRIMA|VACACIONES|LICENCIA)/.test(
      upper
    )
  ) {
    return "";
  }

  return clean(text);
}

// ============================================================================
// MOVIMIENTO CC -> FORMATO LARGO
// ============================================================================

export function convertMovimientoRows(
  rows,
  options = {}
) {
  const conceptKeywords =
    options.conceptKeywords ||
    CONCEPT_KEYWORDS;

  const matchConceptWith =
    (descripcionNorm) =>
      findMovimientoConcept(
        descripcionNorm,
        conceptKeywords
      );

  let header = null;
  let started = false;

  const parsedRows = [];

  const employeesByDate =
    new Map();

  const employeesByMonth =
    new Map();

  const unknownFills =
    new Map();

  let excludedRows = 0;
  let skippedNoDate = 0;
  let dataRows = 0;

  let unclassifiedCount = 0;

  const unclassifiedExamples =
    new Map();

  // --------------------------------------------------------------------------
  // PASADA 1
  // --------------------------------------------------------------------------

  for (
    let r = 0;
    r < rows.length;
    r++
  ) {
    const row =
      rows[r] || [];

    if (!started) {
      const h =
        readMovimientoHeader(
          row
        );

      if (h) {
        header = h;
        started = true;
      }

      continue;
    }

    // Si aparece otro encabezado dentro del archivo.
    if (
      readMovimientoHeader(
        row
      )
    ) {
      continue;
    }

    const comprobanteVal =
      clean(
        cellValue(
          row[
            header
              .comprobanteCol
          ]
        )
      );

    const descripcionVal =
      clean(
        cellValue(
          row[
            header
              .descripcionCol
          ]
        )
      );

    if (
      !comprobanteVal &&
      MOVIMIENTO_STOP_LABEL.test(
        norm(
          descripcionVal
        )
      )
    ) {
      break;
    }

    if (
      !comprobanteVal &&
      !descripcionVal
    ) {
      continue;
    }

    const fecha =
      dateFromExcelSerial(
        cellValue(
          row[
            header
              .fechaCol
          ]
        )
      );

    if (!fecha) {
      skippedNoDate +=
        1;
      continue;
    }

    const monthDate =
      firstOfMonth(
        fecha
      );

    const dKey =
      dateKey(
        fecha
      );

    const terceroVal =
      header.terceroCol !==
      null
        ? clean(
            cellValue(
              row[
                header
                  .terceroCol
              ]
            )
          )
        : "";

    const debitoCell =
      row[
        header
          .debitoCol
      ];

    const creditoCell =
      row[
        header
          .creditoCol
      ];

    const debito =
      toNumber(
        cellValue(
          debitoCell
        )
      );

    const credito =
      toNumber(
        cellValue(
          creditoCell
        )
      );

    const value =
      round2(
        debito -
          credito
      );

    if (
      value === 0
    ) {
      continue;
    }

    const descripcionNorm =
      norm(
        descripcionVal
      );

    const match =
      matchConceptWith(
        descripcionNorm
      );

    if (
      match &&
      match.group ===
        "excluir"
    ) {
      excludedRows +=
        1;
      continue;
    }

    const group =
      match
        ? match.group
        : "otro";

    const concepto =
      match
        ? match.concepto
        : descripcionVal;

    if (!match) {
      unclassifiedCount +=
        1;

      unclassifiedExamples.set(
        descripcionVal,
        (
          unclassifiedExamples.get(
            descripcionVal
          ) || 0
        ) + 1
      );
    }

    // ------------------------------------------------------------------------
    // NUEVO:
    // Si Tercero está vacío, intentar sacar empleado de Descripción.
    // ------------------------------------------------------------------------

    let empleadoDesdeDescripcion =
      "";

    if (
      group ===
        "empleado" &&
      !terceroVal
    ) {
      empleadoDesdeDescripcion =
        extractEmployeeFromDescription(
          descripcionVal,
          match
        );
    }

    const resolvedEmployee =
      terceroVal ||
      empleadoDesdeDescripcion;

    // ------------------------------------------------------------------------
    // Registrar empleados candidatos por fecha y mes.
    // ------------------------------------------------------------------------

    if (
      group ===
        "empleado" &&
      resolvedEmployee
    ) {
      if (
        !employeesByDate.has(
          dKey
        )
      ) {
        employeesByDate.set(
          dKey,
          new Set()
        );
      }

      employeesByDate
        .get(dKey)
        .add(
          resolvedEmployee
        );

      const mKey =
        monthKey(
          monthDate
        );

      if (
        !employeesByMonth.has(
          mKey
        )
      ) {
        employeesByMonth.set(
          mKey,
          new Set()
        );
      }

      employeesByMonth
        .get(mKey)
        .add(
          resolvedEmployee
        );
    }

    // ------------------------------------------------------------------------
    // COLOR
    // ------------------------------------------------------------------------

    const rawFill =
      cellFill(
        debito !== 0
          ? debitoCell
          : creditoCell
      ) ||
      cellFill(
        debitoCell
      ) ||
      cellFill(
        creditoCell
      );

    const fill =
      statusFill(
        rawFill
      );

    if (
      !fill &&
      rawFill &&
      isHexFill(
        rawFill
      ) &&
      !STRUCTURAL_FILLS.has(
        String(
          rawFill
        ).toUpperCase()
      )
    ) {
      const hex =
        String(
          rawFill
        ).toUpperCase();

      unknownFills.set(
        hex,
        (
          unknownFills.get(
            hex
          ) || 0
        ) + 1
      );
    }

    dataRows +=
      1;

    parsedRows.push({
      monthDate,
      dKey,
      comprobante:
        comprobanteVal,
      concepto,
      group,
      tercero:
        terceroVal,
      empleadoDirecto:
        resolvedEmployee,
      value,
      fill
    });
  }

  if (
    !header ||
    dataRows === 0
  ) {
    return null;
  }

  // --------------------------------------------------------------------------
  // PASADA 2
  // Resolver aportes patronales.
  // --------------------------------------------------------------------------

  let ambiguousAporte =
    0;

  let unresolvedAporte =
    0;

  let monthFallback =
    0;

  for (
    const pr of parsedRows
  ) {
    // Conceptos normales de empleado.
    if (
      pr.group !==
      "aporte"
    ) {
      pr.empleado =
        pr.empleadoDirecto ||
        pr.tercero ||
        "";

      continue;
    }

    // Si el propio movimiento trae tercero o empleado
    // en descripción, usarlo primero.
    if (
      pr.empleadoDirecto
    ) {
      pr.empleado =
        pr.empleadoDirecto;

      continue;
    }

    const byDate =
      employeesByDate.get(
        pr.dKey
      );

    const byMonth =
      employeesByMonth.get(
        monthKey(
          pr.monthDate
        )
      );

    const pool =
      byDate &&
      byDate.size > 0
        ? byDate
        : byMonth;

    if (
      pool &&
      pool.size === 1
    ) {
      pr.empleado =
        [...pool][0];

      if (
        pool === byMonth &&
        !(
          byDate &&
          byDate.size > 0
        )
      ) {
        monthFallback +=
          1;
      }
    } else {
      pr.empleado =
        "";

      if (
        pool &&
        pool.size > 1
      ) {
        ambiguousAporte +=
          1;
      } else {
        unresolvedAporte +=
          1;
      }
    }
  }

  // --------------------------------------------------------------------------
  // AGRUPAR
  // --------------------------------------------------------------------------

  const grouped =
    new Map();

  const totalsByMonthEmployee =
    new Map();

  const colorCount =
    new Map();

  let uncolored =
    0;

  const months =
    new Set();

  for (
    const pr of parsedRows
  ) {
    months.add(
      monthKey(
        pr.monthDate
      )
    );

    const key =
      `${monthKey(
        pr.monthDate
      )}|${pr.concepto}|${pr.empleado}`;

    if (
      !grouped.has(
        key
      )
    ) {
      grouped.set(
        key,
        {
          monthDate:
            pr.monthDate,

          concepto:
            pr.concepto,

          empleado:
            pr.empleado,

          value: 0,

          fill:
            pr.fill
        }
      );
    }

    const g =
      grouped.get(
        key
      );

    g.value =
      round2(
        g.value +
          pr.value
      );

    if (
      !g.fill &&
      pr.fill
    ) {
      g.fill =
        pr.fill;
    }

    if (
      pr.empleado
    ) {
      const teKey =
        `${monthKey(
          pr.monthDate
        )}|${pr.empleado}`;

      totalsByMonthEmployee.set(
        teKey,
        round2(
          (
            totalsByMonthEmployee.get(
              teKey
            ) || 0
          ) +
            pr.value
        )
      );
    }
  }

  // --------------------------------------------------------------------------
  // RECORDS
  // --------------------------------------------------------------------------

  const records = [];

  for (
    const g of grouped.values()
  ) {
    if (
      g.value === 0
    ) {
      continue;
    }

    records.push({
      "Mes elaboración":
        g.monthDate,

      Concepto:
        g.concepto,

      Empleado:
        g.empleado ||
        SIN_EMPLEADO,

      "Valor Concepto":
        g.value,

      "Valor Totales":
        0,

      _fill:
        g.fill
    });

    if (g.fill) {
      colorCount.set(
        g.fill,
        (
          colorCount.get(
            g.fill
          ) || 0
        ) + 1
      );
    } else {
      uncolored +=
        1;
    }
  }

  // --------------------------------------------------------------------------
  // TOTAL EMPLOYEE COST
  // --------------------------------------------------------------------------

  for (
    const [
      key,
      total
    ] of totalsByMonthEmployee
  ) {
    if (
      total === 0
    ) {
      continue;
    }

    const parts =
      key.split("|");

    const mKey =
      parts[0];

    const empleado =
      parts
        .slice(1)
        .join("|");

    const [
      y,
      m
    ] =
      mKey
        .split("-")
        .map(Number);

    records.push({
      "Mes elaboración":
        new Date(
          Date.UTC(
            y,
            m - 1,
            1
          )
        ),

      Concepto:
        "TOTAL EMPLOYEE COST",

      Empleado:
        empleado ||
        SIN_EMPLEADO,

      "Valor Concepto":
        0,

      "Valor Totales":
        total,

      _fill:
        null
    });
  }

  // --------------------------------------------------------------------------
  // ORDEN
  // --------------------------------------------------------------------------

  const firstSeen =
    new Map();

  parsedRows.forEach(
    (pr, idx) => {
      const key =
        `${monthKey(
          pr.monthDate
        )}|${
          pr.empleado ||
          SIN_EMPLEADO
        }|${pr.concepto}`;

      if (
        !firstSeen.has(
          key
        )
      ) {
        firstSeen.set(
          key,
          idx
        );
      }
    }
  );

  const sortKey =
    (rec) => {
      const d =
        rec[
          "Mes elaboración"
        ];

      const monthNum =
        d.getUTCFullYear() *
          12 +
        d.getUTCMonth();

      const isTotal =
        rec.Concepto ===
        "TOTAL EMPLOYEE COST"
          ? 1
          : 0;

      const orderKey =
        `${monthKey(
          d
        )}|${rec.Empleado}|${rec.Concepto}`;

      const order =
        firstSeen.has(
          orderKey
        )
          ? firstSeen.get(
              orderKey
            )
          : Number.MAX_SAFE_INTEGER;

      return [
        monthNum,
        rec.Empleado,
        isTotal,
        order
      ];
    };

  records.sort(
    (a, b) => {
      const ka =
        sortKey(a);

      const kb =
        sortKey(b);

      for (
        let i = 0;
        i < ka.length;
        i++
      ) {
        if (
          ka[i] < kb[i]
        ) {
          return -1;
        }

        if (
          ka[i] > kb[i]
        ) {
          return 1;
        }
      }

      return 0;
    }
  );

  // --------------------------------------------------------------------------
  // AVISOS
  // --------------------------------------------------------------------------

  const notes = [];

  const sortedMonths =
    [...months].sort();

  notes.push({
    type: "info",
    text:
      `Se leyeron ${dataRows} movimiento(s) en ${sortedMonths.length} mes(es)` +
      (
        sortedMonths.length
          ? ` (${sortedMonths[0]} a ${sortedMonths[sortedMonths.length - 1]})`
          : ""
      ) +
      ` y se agruparon en ${grouped.size} fila(s) de concepto + ${totalsByMonthEmployee.size} de TOTAL EMPLOYEE COST.`
  });

  if (
    excludedRows > 0
  ) {
    notes.push({
      type: "info",
      text:
        `${excludedRows} movimiento(s) de facturación/ingresos se excluyeron por no ser costo de un empleado.`
    });
  }

  if (
    colorCount.size > 0 ||
    uncolored > 0
  ) {
    const parts =
      [...colorCount.entries()]
        .sort(
          (a, b) =>
            b[1] - a[1]
        )
        .map(
          ([hex, n]) =>
            `${STATUS_COLORS[hex] || `#${hex}`}: ${n}`
        );

    notes.push({
      type: "info",
      text:
        `Colores tomados del movimiento — ${
          parts.length
            ? parts.join(", ")
            : "ninguno"
        }${
          uncolored > 0
            ? `; sin color: ${uncolored}`
            : ""
        }.`
    });
  }

  if (
    monthFallback > 0
  ) {
    notes.push({
      type: "info",
      text:
        `${monthFallback} movimiento(s) de aporte patronal no tenían un salario en su fecha exacta y se asignaron al único empleado de ese mes.`
    });
  }

  if (
    unresolvedAporte > 0
  ) {
    notes.push({
      type: "warn",
      text:
        `${unresolvedAporte} fila(s) de aporte patronal no se pudieron asignar a ningún empleado porque no hay ninguna fila de salario en esa fecha ni en ese mes; quedaron marcadas como "${SIN_EMPLEADO}".`
    });
  }

  if (
    ambiguousAporte > 0
  ) {
    notes.push({
      type: "warn",
      text:
        `${ambiguousAporte} fila(s) de aporte patronal quedaron marcadas como "${SIN_EMPLEADO}" porque hay más de un empleado candidato.`
    });
  }

  if (
    unclassifiedCount > 0
  ) {
    const ex =
      [...unclassifiedExamples.entries()]
        .slice(0, 5)
        .map(
          ([desc, n]) =>
            `"${desc}" (${n})`
        )
        .join("; ");

    notes.push({
      type: "warn",
      text:
        `${unclassifiedCount} movimiento(s) no coincidieron con ninguna palabra clave de CONCEPT_KEYWORDS y quedaron con el texto tal cual de Descripción: ${ex}.`
    });
  }

  if (
    skippedNoDate > 0
  ) {
    notes.push({
      type: "warn",
      text:
        `${skippedNoDate} fila(s) se omitieron por no tener una fecha válida en "Fecha elaboración".`
    });
  }

  if (
    unknownFills.size > 0
  ) {
    const list =
      [...unknownFills.entries()]
        .map(
          ([hex, n]) =>
            `#${hex} (${n})`
        )
        .join(", ");

    notes.push({
      type: "warn",
      text:
        `Colores del movimiento que no están en la leyenda del cruce y se dejaron sin color: ${list}.`
    });
  }

  notes.push({
    type: "info",
    text:
      "Movimiento CC: el empleado se toma de Tercero cuando existe; si Tercero está vacío, se intenta extraer desde Descripción para conceptos de empleado. Los aportes patronales se asignan por fecha y, cuando solo existe un candidato, por mes."
  });

  return {
    records,
    notes
  };
}

// ============================================================================
// CONSOLIDAR VARIOS MOVIMIENTO CC
// ============================================================================

function trimMovimientoStopSection(
  rows
) {
  let header = null;
  let headerFound = false;

  for (
    let r = 0;
    r < rows.length;
    r++
  ) {
    const row =
      rows[r] || [];

    if (!headerFound) {
      const h =
        readMovimientoHeader(
          row
        );

      if (h) {
        header =
          h;

        headerFound =
          true;
      }

      continue;
    }

    if (
      readMovimientoHeader(
        row
      )
    ) {
      continue;
    }

    const comprobanteVal =
      clean(
        cellValue(
          row[
            header
              .comprobanteCol
          ]
        )
      );

    const descripcionVal =
      clean(
        cellValue(
          row[
            header
              .descripcionCol
          ]
        )
      );

    if (
      !comprobanteVal &&
      MOVIMIENTO_STOP_LABEL.test(
        norm(
          descripcionVal
        )
      )
    ) {
      return rows.slice(
        0,
        r
      );
    }
  }

  return rows;
}

export function convertMovimientoFiles(
  files,
  options = {}
) {
  const usable =
    (files || []).filter(
      (f) =>
        f &&
        Array.isArray(
          f.rows
        ) &&
        f.rows.length > 0
    );

  if (
    usable.length === 0
  ) {
    return null;
  }

  const allRows = [];

  for (
    const f of usable
  ) {
    allRows.push(
      ...trimMovimientoStopSection(
        f.rows
      )
    );
  }

  const result =
    convertMovimientoRows(
      allRows,
      options
    );

  if (!result) {
    return null;
  }

  result.records =
    [...result.records].sort(
      (a, b) =>
        a["Mes elaboración"] -
        b["Mes elaboración"]
    );

  const names =
    usable
      .map(
        (f, i) =>
          f.name ||
          `archivo ${i + 1}`
      )
      .join(", ");

  result.notes.unshift({
    type: "info",
    text:
      `Se consolidaron ${usable.length} archivo(s) de Movimiento CC en un solo resultado (${names}).`
  });

  return result;
}

// ============================================================================
// CUENTA 28
// ============================================================================

export const CUENTA28_STATUS_COLORS = {
  OK: "92D050",
  DIFERENCIA: "FF0000",
  SOLO_NOMINA: "FFFF00",
  SOLO_SIIGO: "FFFF00",
  CRUZA_ENTRE_MESES: "00FF00"
};

export const CUENTA28_ACCUMULATED_CONCEPTS = [
  "13TH SALARY",
  "PRIMA",
  "PRIMA DE SERVICIOS"
];

// ============================================================================
// NORMALIZAR CONCEPTO PARA CUENTA 28
// ============================================================================

function canonicalConcept(
  value
) {
  const t =
    norm(value);

  if (!t) {
    return "";
  }

  // PRIMA
  if (
    t.includes("PRIMA") ||
    t === "13TH SALARY" ||
    t.includes(
      "THIRTEENTH SALARY"
    )
  ) {
    return "13TH SALARY";
  }

  // SALARIO
  if (
    t.includes("SALARIO") ||
    t === "SALARY" ||
    t.includes(
      "ORDINARY SALARY"
    ) ||
    t.includes(
      "INTEGRATED SALARY"
    )
  ) {
    return "SALARY";
  }

  // TRANSPORTE
  if (
    t.includes(
      "TRANSPORT"
    ) ||
    t.includes(
      "SUBSIDIO DE TRANSPORTE"
    )
  ) {
    return "TRANSPORT ALLOWANCE";
  }

  // PENSIÓN
  if (
    t.includes(
      "PENSION"
    ) ||
    t.includes(
      "FONDO DE PENSION"
    ) ||
    t.includes(
      "APORTES A FONDOS"
    )
  ) {
    return "PENSION COST";
  }

  // SALUD
  if (
    t.includes(
      "HEALTH"
    ) ||
    t.includes(
      "SALUD"
    ) ||
    t.includes(
      "EPS"
    ) ||
    t.includes(
      "ENTIDADES PROMOTORAS"
    )
  ) {
    return "HEALTH COST";
  }

  // CAJA
  if (
    t.includes(
      "FAMILY FUND"
    ) ||
    t.includes(
      "CAJA DE COMPENSACION"
    ) ||
    t.includes(
      "CAJAS DE COMPENSACION"
    )
  ) {
    return "FAMILY FUND COST";
  }

  // ARL
  if (
    t.includes(
      "LABOR RISK"
    ) ||
    t.includes(
      "RIESGOS"
    ) ||
    t.includes(
      "RIESGO LABORAL"
    ) ||
    t.includes(
      "ARL"
    ) ||
    t.includes(
      "ADMINISTRADORAS DE RIESGOS"
    )
  ) {
    return "LABOR RISK COST";
  }

  return t;
}

// ============================================================================
// NORMALIZAR EMPLEADO
// ============================================================================

function canonicalEmployee(
  value
) {
  const tokens =
    norm(value)
      .replace(
        /[^A-Z0-9 ]/g,
        " "
      )
      .split(/\s+/)
      .filter(Boolean);

  return tokens
    .sort()
    .join(" ");
}

// ============================================================================
// CRUCE CUENTA 28
// ============================================================================

export function cruzarCuenta28(
  nominaRecords = [],
  movimientoRecords = [],
  options = {}
) {
  const tolerance =
    Number.isFinite(
      options.tolerance
    )
      ? options.tolerance
      : 1;

  const accumulatedConcepts =
    new Set(
      (
        options.accumulatedConcepts ||
        CUENTA28_ACCUMULATED_CONCEPTS
      ).map(
        canonicalConcept
      )
    );

  // --------------------------------------------------------------------------
  // CONSTRUIR MAPA
  // --------------------------------------------------------------------------

  const buildMap =
    (records) => {
      const map =
        new Map();

      for (
        const row of
          records || []
      ) {
        const conceptoOriginal =
          clean(
            row?.Concepto
          );

        if (
          !conceptoOriginal ||
          conceptoOriginal ===
            "TOTAL EMPLOYEE COST"
        ) {
          continue;
        }

        const empleado =
          canonicalEmployee(
            row?.Empleado
          );

        if (
          !empleado ||
          empleado ===
            canonicalEmployee(
              SIN_EMPLEADO
            )
        ) {
          continue;
        }

        const date =
          row?.[
            "Mes elaboración"
          ];

        if (
          !(date instanceof Date) ||
          Number.isNaN(
            date.getTime()
          )
        ) {
          continue;
        }

        const mes =
          monthKey(
            date
          );

        const concepto =
          canonicalConcept(
            conceptoOriginal
          );

        const valor =
          round2(
            toNumber(
              row?.[
                "Valor Concepto"
              ]
            )
          );

        if (
          !concepto ||
          valor === 0
        ) {
          continue;
        }

        const key =
          `${mes}|${empleado}|${concepto}`;

        if (
          !map.has(
            key
          )
        ) {
          map.set(
            key,
            {
              mes,
              empleado,
              concepto,
              empleadoNombre:
                clean(
                  row?.Empleado
                ),
              conceptoOriginal,
              valor: 0
            }
          );
        }

        const item =
          map.get(
            key
          );

        item.valor =
          round2(
            item.valor +
              valor
          );
      }

      return map;
    };

  const nomina =
    buildMap(
      nominaRecords
    );

  const siigo =
    buildMap(
      movimientoRecords
    );

  // --------------------------------------------------------------------------
  // MESES
  // --------------------------------------------------------------------------

  const months =
    [
      ...new Set(
        [
          ...nomina.values(),
          ...siigo.values()
        ].map(
          (x) => x.mes
        )
      )
    ].sort();

  // --------------------------------------------------------------------------
  // DIMENSIONES
  // --------------------------------------------------------------------------

  const dimensions =
    new Map();

  for (
    const item of [
      ...nomina.values(),
      ...siigo.values()
    ]
  ) {
    const key =
      `${item.empleado}|${item.concepto}`;

    if (
      !dimensions.has(
        key
      )
    ) {
      dimensions.set(
        key,
        {
          empleado:
            item.empleado,

          empleadoNombre:
            item.empleadoNombre,

          concepto:
            item.concepto,

          conceptoOriginal:
            item.conceptoOriginal
        }
      );
    }
  }

  // --------------------------------------------------------------------------
  // ACUMULADOS
  // --------------------------------------------------------------------------

  const cumulativeThrough =
    (
      map,
      empleado,
      concepto,
      month
    ) => {
      let total = 0;

      for (
        const item of
          map.values()
      ) {
        if (
          item.empleado ===
            empleado &&
          item.concepto ===
            concepto &&
          item.mes <=
            month
        ) {
          total =
            round2(
              total +
                item.valor
            );
        }
      }

      return total;
    };

  // --------------------------------------------------------------------------
  // CRUCE
  // --------------------------------------------------------------------------

  const rows = [];

  for (
    const dim of
      dimensions.values()
  ) {
    for (
      const mes of months
    ) {
      const key =
        `${mes}|${dim.empleado}|${dim.concepto}`;

      const n =
        nomina.get(
          key
        );

      const s =
        siigo.get(
          key
        );

      const accumulated =
        accumulatedConcepts.has(
          dim.concepto
        );

      const nominaValor =
        accumulated
          ? cumulativeThrough(
              nomina,
              dim.empleado,
              dim.concepto,
              mes
            )
          : round2(
              n?.valor || 0
            );

      const siigoValor =
        accumulated
          ? cumulativeThrough(
              siigo,
              dim.empleado,
              dim.concepto,
              mes
            )
          : round2(
              s?.valor || 0
            );

      const diferencia =
        round2(
          nominaValor -
            siigoValor
        );

      const existeHastaMes =
        nominaValor !== 0 ||
        siigoValor !== 0;

      if (
        !existeHastaMes
      ) {
        continue;
      }

      let estado;

      // ----------------------------------------------------------------------
      // OK / CRUZA ENTRE MESES
      // ----------------------------------------------------------------------

      if (
        Math.abs(
          diferencia
        ) <= tolerance
      ) {
        if (
          accumulated &&
          n &&
          s &&
          Math.abs(
            (n.valor || 0) -
              (s.valor || 0)
          ) > tolerance
        ) {
          estado =
            "CRUZA_ENTRE_MESES";
        } else if (
          accumulated &&
          (!n || !s) &&
          (
            nominaValor !== 0 ||
            siigoValor !== 0
          )
        ) {
          estado =
            "CRUZA_ENTRE_MESES";
        } else {
          estado =
            "OK";
        }
      }

      // ----------------------------------------------------------------------
      // SOLO NÓMINA
      // ----------------------------------------------------------------------

      else if (
        n &&
        !s &&
        !accumulated
      ) {
        estado =
          "SOLO_NOMINA";
      }

      // ----------------------------------------------------------------------
      // SOLO SIIGO
      // ----------------------------------------------------------------------

      else if (
        !n &&
        s &&
        !accumulated
      ) {
        estado =
          "SOLO_SIIGO";
      }

      // ----------------------------------------------------------------------
      // DIFERENCIA
      // ----------------------------------------------------------------------

      else if (
        n &&
        s
      ) {
        estado =
          "DIFERENCIA";
      }

      // ----------------------------------------------------------------------
      // CASOS RESTANTES
      // ----------------------------------------------------------------------

      else {
        estado =
          n
            ? "SOLO_NOMINA"
            : "SOLO_SIIGO";
      }

      rows.push({
        Mes:
          mes,

        Empleado:
          n?.empleadoNombre ||
          s?.empleadoNombre ||
          dim.empleadoNombre ||
          dim.empleado,

        Concepto:
          n?.conceptoOriginal ||
          s?.conceptoOriginal ||
          dim.conceptoOriginal ||
          dim.concepto,

        "Valor Nómina":
          nominaValor,

        "Valor Siigo":
          siigoValor,

        Diferencia:
          diferencia,

        Estado:
          estado,

        Acumulado:
          accumulated,

        "Nómina mes":
          round2(
            n?.valor || 0
          ),

        "Siigo mes":
          round2(
            s?.valor || 0
          ),

        _fill:
          CUENTA28_STATUS_COLORS[
            estado
          ]
      });
    }
  }

  // --------------------------------------------------------------------------
  // ORDENAR
  // --------------------------------------------------------------------------

  rows.sort(
    (a, b) => {
      if (
        a.Mes !==
        b.Mes
      ) {
        return a.Mes.localeCompare(
          b.Mes
        );
      }

      if (
        a.Empleado !==
        b.Empleado
      ) {
        return a.Empleado.localeCompare(
          b.Empleado
        );
      }

      return a.Concepto.localeCompare(
        b.Concepto
      );
    }
  );

  // --------------------------------------------------------------------------
  // RESUMEN
  // --------------------------------------------------------------------------

  const summary =
    rows.reduce(
      (acc, row) => {
        acc.total +=
          1;

        acc[
          row.Estado
        ] =
          (
            acc[
              row.Estado
            ] || 0
          ) + 1;

        return acc;
      },
      {
        total: 0,
        OK: 0,
        DIFERENCIA: 0,
        SOLO_NOMINA: 0,
        SOLO_SIIGO: 0,
        CRUZA_ENTRE_MESES: 0
      }
    );

  return {
    rows,
    summary,
    months
  };
}