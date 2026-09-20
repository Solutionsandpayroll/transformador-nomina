import React, { useState, useMemo, useCallback } from 'react';
import JSZip from 'jszip';
import {
  Info,
  ChevronUp,
  ChevronDown,
  User,
  Upload,
  Download,
  CheckCircle2,
  AlertTriangle,
  Search,
  ChevronLeft,
  ChevronRight,
  FileSpreadsheet,
  X,
  FileX2
} from 'lucide-react';
import { convertSiigoRows } from './siigoConverter';

// ============================================================================
// LECTOR DE .xlsx / .xlsm CON JSZIP (sin librería xlsx/SheetJS)
// ============================================================================
// Un archivo .xlsx/.xlsm es en realidad un .zip con archivos XML adentro.
// Aquí lo desempacamos con JSZip y leemos a mano los XML que necesitamos:
//   - xl/workbook.xml            -> lista de hojas y sus IDs de relación
//   - xl/_rels/workbook.xml.rels -> a qué archivo físico apunta cada hoja
//   - xl/sharedStrings.xml       -> tabla de textos compartidos (Excel no
//                                   repite el mismo texto en cada celda, usa
//                                   un índice a esta tabla)
//   - xl/worksheets/sheetN.xml   -> las celdas de cada hoja

function colLettersToIndex(letters) {
  let result = 0;
  const upper = letters.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    result = result * 26 + (upper.charCodeAt(i) - 64);
  }
  return result - 1; // 0-based
}

function parseCellRef(ref) {
  const match = /^([A-Za-z]+)(\d+)$/.exec(ref || '');
  if (!match) return null;
  return { col: colLettersToIndex(match[1]), row: parseInt(match[2], 10) };
}

function parseSharedStrings(xmlDoc) {
  const siNodes = xmlDoc.getElementsByTagName('si');
  const strings = [];
  for (let i = 0; i < siNodes.length; i++) {
    const tNodes = siNodes[i].getElementsByTagName('t');
    let text = '';
    for (let j = 0; j < tNodes.length; j++) {
      text += tNodes[j].textContent;
    }
    strings.push(text);
  }
  return strings;
}

function parseWorkbookSheetList(xmlDoc) {
  const sheetNodes = xmlDoc.getElementsByTagName('sheet');
  const sheets = [];
  for (let i = 0; i < sheetNodes.length; i++) {
    const node = sheetNodes[i];
    const name = node.getAttribute('name');
    // El atributo r:id puede venir con o sin el prefijo del namespace
    const rId =
      node.getAttribute('r:id') ||
      node.getAttributeNS(
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'id'
      );
    sheets.push({ name, rId });
  }
  return sheets;
}

function parseWorkbookRels(xmlDoc) {
  const relNodes = xmlDoc.getElementsByTagName('Relationship');
  const map = {};
  for (let i = 0; i < relNodes.length; i++) {
    const node = relNodes[i];
    map[node.getAttribute('Id')] = node.getAttribute('Target');
  }
  return map;
}

function resolveWorksheetPath(target) {
  // Los targets suelen venir como "worksheets/sheet1.xml" (relativos a xl/)
  if (target.startsWith('/xl/')) return target.slice(1);
  if (target.startsWith('xl/')) return target;
  return `xl/${target}`;
}

function parseSheetXmlToRows(xmlDoc, sharedStrings) {
  const rowNodes = xmlDoc.getElementsByTagName('row');
  const rows = [];

  for (let i = 0; i < rowNodes.length; i++) {
    const rowNode = rowNodes[i];
    const rowNumAttr = rowNode.getAttribute('r');
    const rowIndex = rowNumAttr ? parseInt(rowNumAttr, 10) - 1 : i;

    const cellNodes = rowNode.getElementsByTagName('c');
    const rowArray = [];

    for (let j = 0; j < cellNodes.length; j++) {
      const cellNode = cellNodes[j];
      const ref = cellNode.getAttribute('r');
      const parsed = ref ? parseCellRef(ref) : null;
      const colIndex = parsed ? parsed.col : j;
      const type = cellNode.getAttribute('t');

      let value = null;

      if (type === 'inlineStr') {
        const isNode = cellNode.getElementsByTagName('is')[0];
        if (isNode) {
          const tNodes = isNode.getElementsByTagName('t');
          let text = '';
          for (let k = 0; k < tNodes.length; k++) text += tNodes[k].textContent;
          value = text;
        }
      } else {
        const vNode = cellNode.getElementsByTagName('v')[0];
        const rawText = vNode ? vNode.textContent : null;

        if (rawText === null || rawText === '') {
          value = null;
        } else if (type === 's') {
          const idx = parseInt(rawText, 10);
          value = sharedStrings[idx] !== undefined ? sharedStrings[idx] : null;
        } else if (type === 'b') {
          value = rawText === '1';
        } else if (type === 'str' || type === 'e') {
          value = rawText;
        } else {
          // numérico por defecto
          const num = Number(rawText);
          value = Number.isFinite(num) ? num : rawText;
        }
      }

      rowArray[colIndex] = value;
    }

    // Rellenar huecos con null para que los índices de columna sean estables
    for (let c = 0; c < rowArray.length; c++) {
      if (rowArray[c] === undefined) rowArray[c] = null;
    }

    rows[rowIndex] = rowArray;
  }

  // Rellenar filas completamente vacías que Excel omitió (no escribió <row>)
  for (let r = 0; r < rows.length; r++) {
    if (!rows[r]) rows[r] = [];
  }

  return rows;
}

async function readWorkbookSheetsWithJSZip(file) {
  const zip = await JSZip.loadAsync(file);
  const parser = new DOMParser();

  const workbookXmlFile = zip.file('xl/workbook.xml');
  if (!workbookXmlFile) {
    throw new Error('El archivo no parece ser un .xlsx/.xlsm válido (falta xl/workbook.xml).');
  }
  const workbookXmlText = await workbookXmlFile.async('text');
  const workbookXmlDoc = parser.parseFromString(workbookXmlText, 'text/xml');
  const sheetList = parseWorkbookSheetList(workbookXmlDoc);

  const relsFile = zip.file('xl/_rels/workbook.xml.rels');
  let relsMap = {};
  if (relsFile) {
    const relsXmlText = await relsFile.async('text');
    const relsXmlDoc = parser.parseFromString(relsXmlText, 'text/xml');
    relsMap = parseWorkbookRels(relsXmlDoc);
  }

  let sharedStrings = [];
  const sharedStringsFile = zip.file('xl/sharedStrings.xml');
  if (sharedStringsFile) {
    const sharedXmlText = await sharedStringsFile.async('text');
    const sharedXmlDoc = parser.parseFromString(sharedXmlText, 'text/xml');
    sharedStrings = parseSharedStrings(sharedXmlDoc);
  }

  const sheets = [];
  for (const { name, rId } of sheetList) {
    const target = relsMap[rId];
    if (!target) continue;
    const path = resolveWorksheetPath(target);
    const sheetFile = zip.file(path);
    if (!sheetFile) continue;

    const sheetXmlText = await sheetFile.async('text');
    const sheetXmlDoc = parser.parseFromString(sheetXmlText, 'text/xml');
    const rows = parseSheetXmlToRows(sheetXmlDoc, sharedStrings);
    sheets.push({ name, rows });
  }

  return sheets;
}

// ============================================================================
// UTILIDADES DE FECHA — "Mes elaboración" debe quedar como fecha real
// ============================================================================
// El rótulo del mes llega como texto libre en la columna A de cada bloque
// ("Abril 2025", "MAYO 2025", "Marzo 2026", e incluso con errores de
// digitación reales como "NNOVIEMBRE 2025"). El archivo de ejemplo que
// compartió el cliente (Hoja2 de BUBBLE_-_Nómina_1.xlsm) espera que esa
// columna sea una fecha (primer día del mes), no el texto tal cual.

const MONTHS_ES = [
  { key: 'ENERO', num: 0 },
  { key: 'FEBRERO', num: 1 },
  { key: 'MARZO', num: 2 },
  { key: 'ABRIL', num: 3 },
  { key: 'MAYO', num: 4 },
  { key: 'JUNIO', num: 5 },
  { key: 'JULIO', num: 6 },
  { key: 'AGOSTO', num: 7 },
  { key: 'SEPTIEMBRE', num: 8 },
  { key: 'SETIEMBRE', num: 8 },
  { key: 'OCTUBRE', num: 9 },
  { key: 'NOVIEMBRE', num: 10 },
  { key: 'DICIEMBRE', num: 11 }
];

function removeAccents(value) {
  return value.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

// Colapsa letras repetidas consecutivas: "NNOVIEMBRE" -> "NOVIEMBRE"
function collapseRepeatedLetters(value) {
  return value.replace(/([A-Z])\1+/g, '$1');
}

// Convierte un rótulo de mes ("Abril 2025", "NNOVIEMBRE 2025", "MAYO 2025")
// en un objeto Date (UTC, día 1 del mes). Devuelve null si no logra
// reconocer mes y año en el texto.
function parseMonthLabelToDate(label) {
  if (label === null || label === undefined) return null;
  const cleaned = removeAccents(String(label).toUpperCase());
  const yearMatch = cleaned.match(/(\d{4})/);
  if (!yearMatch) return null;
  const year = parseInt(yearMatch[1], 10);

  const lettersOnly = cleaned.replace(/[^A-Z]/g, '');

  let found = MONTHS_ES.find((m) => lettersOnly.includes(m.key));
  if (!found) {
    // Tolerar errores de digitación tipo letras dobladas ("NNOVIEMBRE")
    const collapsed = collapseRepeatedLetters(lettersOnly);
    found = MONTHS_ES.find((m) => collapsed.includes(m.key));
  }
  if (!found) return null;

  return new Date(Date.UTC(year, found.num, 1));
}

// Serial de fecha estilo Excel (días desde 1899-12-30), usando aritmética
// UTC en ambos lados para evitar corrimientos por huso horario.
function excelSerialFromDate(date) {
  const excelEpochUTC = Date.UTC(1899, 11, 30);
  return Math.round((date.getTime() - excelEpochUTC) / 86400000);
}

// Formatea cualquier valor de celda para mostrar en tabla / CSV / búsqueda.
function formatCellValue(value) {
  if (value instanceof Date) {
    const yyyy = value.getUTCFullYear();
    const mm = String(value.getUTCMonth() + 1).padStart(2, '0');
    const dd = String(value.getUTCDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }
  if (typeof value === 'number') return value.toLocaleString('es-CO');
  if (value === null || value === undefined) return '';
  return String(value);
}

// ============================================================================
// MODO 1 — MOTOR DE TRANSFORMACIÓN: nómina "ancha" (bloques por mes) -> "largo"
// ============================================================================
//
// Estructura que este motor espera encontrar en cada hoja de un archivo:
//   - En algún punto de la hoja aparece una fila que contiene la etiqueta
//     "EMPLOYEE CODE" (o similar) en alguna columna, seguida de "NAME" en la
//     columna siguiente. Esa fila es el "encabezado" de un bloque de mes.
//   - El nombre del mes suele estar en la columna A de esa misma fila o de la
//     fila inmediatamente anterior (categorías como PAYROLL, TOTALS, etc.)
//   - Después del encabezado vienen 1 o más filas con datos de empleados,
//     hasta que aparece una fila completamente vacía (separador) o el
//     encabezado del siguiente bloque.
//   - Cada bloque puede tener columnas de "conceptos" distintas (un mes trae
//     una prima, otro no la trae, etc.), así que las columnas se detectan
//     dinámicamente leyendo esa fila de encabezado, no una posición fija.
//
// Reglas de negocio aplicadas (confirmadas con el ejemplo real que compartió
// el equipo, Hoja1 -> Hoja2 de BUBBLE_-_Nómina_1.xlsm):
//   1. Las columnas que son subtotales (PAYMENTS, TOTAL, TOTAL COP, TOTAL USD,
//      FEE, FEE USD, EXCHANGE RATE, o cualquier encabezado que contenga la
//      palabra TOTAL) NO se incluyen como "concepto" en el resultado.
//   2. La columna "TOTAL EMPLOYEE COST" es la única excepción: se convierte en
//      una fila especial cuyo valor va en "Valor Totales" en vez de
//      "Valor Concepto".
//   3. Los conceptos sin valor, vacíos o en cero NO generan fila en el
//      resultado.
//   4. Las filas "fantasma" (sin código de empleado y sin nombre, pero con
//      valores repetidos) se descartan porque no se pueden atribuir a nadie.
//   5. Encabezados que son puramente numéricos (residuos de la plantilla) se
//      ignoran, ya que no son nombres de concepto reales.
//   6. "Mes elaboración" se entrega como fecha (primer día del mes), tolerando
//      variaciones de mayúsculas/minúsculas y errores de digitación reales
//      como "NNOVIEMBRE 2025".
//   7. (NUEVO) Las columnas que no son montos (fechas de ingreso, país, estado,
//      tipo de servicio, % de aportes...) se ignoran.
//   8. (NUEVO) El nombre del empleado se unifica por código de empleado: se usa
//      el del último bloque de la hoja, así todos los meses salen igual aunque
//      el archivo cambie el orden "NOMBRE APELLIDOS" / "APELLIDOS NOMBRE".

const HEADER_MARKER = /EMPLOYEE\s*CODE/i;
const NAME_MARKER = /^NAME$/i;
const SPECIAL_TOTAL_LABEL = 'TOTAL EMPLOYEE COST';

const BLACKLIST_EXACT = new Set([
  'PAYMENTS',
  'FEE',
  'FEE USD',
  'EXCHANGE RATE',
  'EMPLOYEE CODE',
  'NAME'
]);

// Encabezados que traen datos que no son dinero (evita filas como
// "Onboarding Date | 45658.2085").
const NON_MONETARY =
  /(DATE|STATUS|COUNTRY|PAYROLL MONTH|SERVICE TYPE|INVOICE TYPE|RATE\s*%|EE RF WID)/i;

function normalizeHeader(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim().toUpperCase();
}

function isNumericLabel(value) {
  if (value === null || value === undefined) return false;
  return /^-?\d+(\.\d+)?$/.test(String(value).trim());
}

function isBlacklistedConcept(normalized) {
  if (normalized === SPECIAL_TOTAL_LABEL) return false;
  if (BLACKLIST_EXACT.has(normalized)) return true;
  if (normalized.includes('TOTAL')) return true;
  return false;
}

function isRowBlank(row, fromCol, toCol) {
  for (let c = fromCol; c <= toCol; c++) {
    const v = row[c];
    if (v !== null && v !== undefined && String(v).trim() !== '') return false;
  }
  return true;
}

function toNumberOrNull(value) {
  if (value === null || value === undefined || value === '') return null;
  const n = typeof value === 'number' ? value : Number(String(value).replace(/,/g, ''));
  return Number.isFinite(n) ? n : null;
}

// Convierte una hoja (array de arrays) en filas largas.
function parseSheetToLongRows(sheetRows, meta) {
  const records = [];
  const latestNameByCode = new Map(); // código de empleado -> último nombre visto
  const n = sheetRows.length;
  let i = 0;

  while (i < n) {
    const row = sheetRows[i] || [];

    // 1. Buscar la columna que contiene "EMPLOYEE CODE" en esta fila
    let codeCol = -1;
    for (let j = 0; j < row.length; j++) {
      if (row[j] !== null && row[j] !== undefined && HEADER_MARKER.test(String(row[j]))) {
        codeCol = j;
        break;
      }
    }

    if (codeCol === -1) {
      i += 1;
      continue;
    }

    // 2. La columna de nombre normalmente es la siguiente; si no calza,
    //    buscarla en el resto de la fila.
    let nameCol = codeCol + 1;
    if (!(row[nameCol] !== undefined && NAME_MARKER.test(String(row[nameCol] || '').trim()))) {
      for (let j = codeCol + 1; j < row.length; j++) {
        if (NAME_MARKER.test(String(row[j] || '').trim())) {
          nameCol = j;
          break;
        }
      }
    }

    // 3. Etiqueta del mes: columna A de esta fila, o de hasta 2 filas arriba
    let monthLabel = null;
    for (let back = 0; back <= 2 && monthLabel === null; back++) {
      const candidateRow = sheetRows[i - back];
      const candidate = candidateRow ? candidateRow[0] : null;
      if (candidate !== null && candidate !== undefined && String(candidate).trim() !== '') {
        monthLabel = String(candidate).trim();
      }
    }
    if (monthLabel === null) monthLabel = `Bloque fila ${i + 1}`;

    // Traducir el rótulo a fecha real; si no se reconoce, se deja el texto
    // original para no perder el dato y que quede visible que hay que revisarlo.
    const parsedMonthDate = parseMonthLabelToDate(monthLabel);
    const mesElaboracion = parsedMonthDate || monthLabel;

    // 4. Detectar columnas de concepto y la columna especial de total
    const concepts = []; // { col, name }
    let totalCol = -1;
    for (let j = nameCol + 1; j < row.length; j++) {
      const raw = row[j];
      if (raw === null || raw === undefined || String(raw).trim() === '') continue;
      if (isNumericLabel(raw)) continue;
      const normalized = normalizeHeader(raw);
      if (normalized === SPECIAL_TOTAL_LABEL) {
        totalCol = j;
      } else if (isBlacklistedConcept(normalized) || NON_MONETARY.test(normalized)) {
        continue;
      } else {
        concepts.push({ col: j, name: String(raw).trim() });
      }
    }

    const lastRelevantCol = Math.max(
      nameCol,
      totalCol,
      ...concepts.map((c) => c.col),
      codeCol
    );

    // 5. Recorrer las filas de datos del bloque hasta encontrar una fila
    //    vacía (separador) o el final de la hoja.
    let k = i + 1;
    while (k < n) {
      const dataRow = sheetRows[k] || [];
      if (isRowBlank(dataRow, codeCol, lastRelevantCol)) break;

      const code = dataRow[codeCol];
      const name = dataRow[nameCol];
      const hasCode = code !== null && code !== undefined && String(code).trim() !== '';
      const hasName = name !== null && name !== undefined && String(name).trim() !== '';

      if (!hasCode && !hasName) {
        // Fila fantasma / duplicada sin identificar a nadie: se descarta.
        k += 1;
        continue;
      }

      const empleado = hasName ? String(name).trim() : String(code).trim();
      const codigo = hasCode ? String(code).replace(/\D/g, '') : '';
      if (codigo) latestNameByCode.set(codigo, empleado);

      for (const concept of concepts) {
        const num = toNumberOrNull(dataRow[concept.col]);
        if (num === null || num === 0) continue; // sin valor -> se excluye
        records.push({
          'Mes elaboración': mesElaboracion,
          Concepto: concept.name,
          Empleado: empleado,
          'Valor Concepto': num,
          'Valor Totales': 0,
          Empresa: meta.empresa,
          Archivo: meta.archivo,
          _codigo: codigo
        });
      }

      if (totalCol !== -1) {
        const totalVal = toNumberOrNull(dataRow[totalCol]);
        if (totalVal !== null && totalVal !== 0) {
          records.push({
            'Mes elaboración': mesElaboracion,
            Concepto: SPECIAL_TOTAL_LABEL,
            Empleado: empleado,
            'Valor Concepto': 0,
            'Valor Totales': totalVal,
            Empresa: meta.empresa,
            Archivo: meta.archivo,
            _codigo: codigo
          });
        }
      }

      k += 1;
    }

    i = k;
  }

  // Unificar el nombre por código de empleado (regla 8) y limpiar el campo interno.
  records.forEach((r) => {
    if (r._codigo && latestNameByCode.has(r._codigo)) {
      r.Empleado = latestNameByCode.get(r._codigo);
    }
    delete r._codigo;
  });

  return records;
}

function companyNameFromFileName(fileName) {
  return fileName.replace(/\.[^/.]+$/, '');
}

function newFileId(file) {
  return `${file.name}-${file.size}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

// --- Modo 1: nómina ancha ---------------------------------------------------
function processWideWorkbook(file, sheets, empresa) {
  let allRecords = [];
  let sheetsUsed = 0;

  for (const { rows } of sheets) {
    const hasMarker = rows.some((row) =>
      (row || []).some((cell) => cell !== null && HEADER_MARKER.test(String(cell)))
    );
    if (!hasMarker) continue;

    sheetsUsed += 1;
    const records = parseSheetToLongRows(rows, { empresa, archivo: file.name });
    allRecords = allRecords.concat(records);
  }

  return {
    fileId: newFileId(file),
    fileName: file.name,
    empresa,
    rows: allRecords,
    notes: [],
    warning:
      sheetsUsed === 0
        ? 'No se encontró en ninguna hoja el patrón de nómina esperado (una fila con "EMPLOYEE CODE"). Revisa que sea el archivo correcto.'
        : null
  };
}

// --- Modo 2: Movimiento CC de Siigo -----------------------------------------
function processSiigoWorkbook(file, sheets, empresa) {
  let converted = null;
  for (const { rows } of sheets) {
    const result = convertSiigoRows(rows);
    if (result) {
      converted = result;
      break; // la primera hoja con encabezado de Movimiento CC (las demás son auxiliares)
    }
  }

  let warning = null;
  if (!converted) {
    warning =
      'No se encontró el encabezado de un Movimiento CC de Siigo (Comprobante, Fecha elaboración, Descripción, Débito, Crédito). Revisa que sea el archivo correcto.';
  } else if (converted.records.length === 0) {
    warning =
      'Se encontró el Movimiento CC, pero ninguna línea de nómina reconocible (comprobantes CC-*).';
  }

  return {
    fileId: newFileId(file),
    fileName: file.name,
    empresa,
    rows: converted ? converted.records : [],
    notes: converted ? converted.notes : [],
    warning
  };
}

async function processWorkbookFile(file, mode) {
  const sheets = await readWorkbookSheetsWithJSZip(file);
  const empresa = companyNameFromFileName(file.name);
  return mode === 'siigo'
    ? processSiigoWorkbook(file, sheets, empresa)
    : processWideWorkbook(file, sheets, empresa);
}

// ============================================================================
// ESCRITOR DE .xlsx CON JSZIP (para la descarga del resultado)
// ============================================================================
// Igual que para leer, generamos a mano el XML mínimo que necesita un .xlsx
// válido: [Content_Types].xml, _rels/.rels, xl/workbook.xml,
// xl/_rels/workbook.xml.rels, xl/styles.xml y xl/worksheets/sheet1.xml.
// Estilos (atributo s): 1 = fecha yyyy-mm-dd; 2 = número con relleno verde
// (viene de Siigo); 3 = número con relleno amarillo (calculado).

function xmlEscape(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function colIndexToLetters(index) {
  let n = index + 1;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

const VALUE_COLUMNS = new Set(['Valor Concepto', 'Valor Totales']);

function buildSheetXml(dataRows, columns) {
  const headerCells = columns
    .map((colName, idx) => {
      const ref = `${colIndexToLetters(idx)}1`;
      return `<c r="${ref}" t="inlineStr"><is><t>${xmlEscape(colName)}</t></is></c>`;
    })
    .join('');
  let xmlRows = `<row r="1">${headerCells}</row>`;

  dataRows.forEach((row, rIdx) => {
    const rowNum = rIdx + 2;
    const cells = columns
      .map((colName, cIdx) => {
        const ref = `${colIndexToLetters(cIdx)}${rowNum}`;
        const val = row[colName];
        if (val === null || val === undefined || val === '') {
          return `<c r="${ref}"/>`;
        }
        if (val instanceof Date) {
          const serial = excelSerialFromDate(val);
          return `<c r="${ref}" s="1"><v>${serial}</v></c>`;
        }
        if (typeof val === 'number') {
          if (row._fill && VALUE_COLUMNS.has(colName)) {
            const style = row._fill === 'yellow' ? 3 : 2;
            return `<c r="${ref}" s="${style}"><v>${val}</v></c>`;
          }
          return `<c r="${ref}"><v>${val}</v></c>`;
        }
        return `<c r="${ref}" t="inlineStr"><is><t>${xmlEscape(String(val))}</t></is></c>`;
      })
      .join('');
    xmlRows += `<row r="${rowNum}">${cells}</row>`;
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${xmlRows}</sheetData></worksheet>`;
}

async function buildXlsxBlobWithJSZip(dataRows, columns) {
  const zip = new JSZip();

  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>';

  const rootRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';

  const workbookXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Nomina largo" sheetId="1" r:id="rId1"/></sheets></workbook>';

  const workbookRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';

  // numFmtId 164 = fecha personalizada (yyyy-mm-dd); numFmtId 4 = #,##0.00 (integrado).
  // fills: 0 = ninguno, 1 = gray125 (obligatorio), 2 = verde, 3 = amarillo.
  // cellXfs: 0 = general, 1 = fecha, 2 = número verde, 3 = número amarillo.
  const stylesXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF92D050"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="4" fontId="0" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/><xf numFmtId="4" fontId="0" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/></cellXfs></styleSheet>';

  const sheetXml = buildSheetXml(dataRows, columns);

  zip.file('[Content_Types].xml', contentTypes);
  zip.file('_rels/.rels', rootRels);
  zip.file('xl/workbook.xml', workbookXml);
  zip.file('xl/_rels/workbook.xml.rels', workbookRels);
  zip.file('xl/styles.xml', stylesXml);
  zip.file('xl/worksheets/sheet1.xml', sheetXml);

  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
}

// ============================================================================
// COMPONENTE PRINCIPAL
// ============================================================================

// Columnas de salida, en el orden y con los nombres exactos del ejemplo real
// (Hoja2 de BUBBLE_-_Nómina_1.xlsm): Mes elaboración, Concepto, Empleado,
// Valor Concepto, Valor Totales.
const BASE_COLUMNS = [
  'Mes elaboración',
  'Concepto',
  'Empleado',
  'Valor Concepto',
  'Valor Totales'
];
// Modo 1 agrega Empresa y Archivo para distinguir las 20+ empresas que se
// consolidan en un solo archivo. Modo 2 (Siigo) sale con las 5 columnas exactas.
const COLUMNS_BY_MODE = {
  nomina: [...BASE_COLUMNS, 'Empresa', 'Archivo'],
  siigo: BASE_COLUMNS
};

const FILL_CLASS = {
  green: 'bg-green-300',
  yellow: 'bg-yellow-300'
};

export default function App() {
  const [mode, setMode] = useState('nomina'); // 'nomina' | 'siigo'
  const [showInstructions, setShowInstructions] = useState(true);
  const [files, setFiles] = useState([]); // { fileId, fileName, empresa, rows, notes, warning }
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const rowsPerPage = 10;

  const outputColumns = COLUMNS_BY_MODE[mode];
  const consolidatedRows = useMemo(() => files.flatMap((f) => f.rows), [files]);

  const changeMode = (nextMode) => {
    if (nextMode === mode) return;
    // Los dos modos generan tablas distintas: se limpia para no mezclarlas.
    setMode(nextMode);
    setFiles([]);
    setSearchTerm('');
    setCurrentPage(1);
  };

  const handleFileUpload = async (e) => {
    const uploaded = Array.from(e.target.files || []);
    if (uploaded.length === 0) return;

    setLoading(true);
    setSearchTerm('');
    setCurrentPage(1);

    try {
      const results = [];
      for (const file of uploaded) {
        try {
          const result = await processWorkbookFile(file, mode);
          results.push(result);
        } catch (err) {
          console.error('Error procesando', file.name, err);
          results.push({
            fileId: `${file.name}-${Date.now()}`,
            fileName: file.name,
            empresa: companyNameFromFileName(file.name),
            rows: [],
            notes: [],
            warning: 'No se pudo leer este archivo. ¿Es un .xlsx/.xlsm válido?'
          });
        }
      }
      setFiles((prev) => [...prev, ...results]);
    } finally {
      setLoading(false);
      e.target.value = '';
    }
  };

  const removeFile = useCallback((fileId) => {
    setFiles((prev) => prev.filter((f) => f.fileId !== fileId));
    setCurrentPage(1);
  }, []);

  const clearAll = () => {
    setFiles([]);
    setSearchTerm('');
    setCurrentPage(1);
  };

  const filteredData = useMemo(() => {
    if (!searchTerm.trim()) return consolidatedRows;
    const term = searchTerm.toLowerCase();
    return consolidatedRows.filter((row) =>
      outputColumns.some((col) => formatCellValue(row[col]).toLowerCase().includes(term))
    );
  }, [consolidatedRows, searchTerm, outputColumns]);

  const totalPages = Math.ceil(filteredData.length / rowsPerPage) || 1;
  const paginatedData = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return filteredData.slice(start, start + rowsPerPage);
  }, [filteredData, currentPage]);

  const downloadXLSX = async () => {
    if (filteredData.length === 0) return;
    setExporting(true);
    try {
      const blob = await buildXlsxBlobWithJSZip(filteredData, outputColumns);
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = 'nomina_formato_largo.xlsx';
      link.click();
      URL.revokeObjectURL(url);
    } finally {
      setExporting(false);
    }
  };

  const downloadCSV = () => {
    if (filteredData.length === 0) return;
    const headerLine = outputColumns.join(',');
    const lines = filteredData.map((row) =>
      outputColumns.map((col) => `"${formatCellValue(row[col]).replace(/"/g, '""')}"`).join(',')
    );
    const csvContent = [headerLine, ...lines].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'nomina_formato_largo.csv';
    link.click();
    URL.revokeObjectURL(url);
  };

  const downloadJSON = () => {
    if (filteredData.length === 0) return;
    // Las fechas se serializan como "yyyy-mm-dd" en vez del ISO string con
    // hora que produciría JSON.stringify por defecto sobre un objeto Date.
    const serializable = filteredData.map((row) => {
      const obj = {};
      outputColumns.forEach((col) => {
        obj[col] = row[col] instanceof Date ? formatCellValue(row[col]) : row[col];
      });
      return obj;
    });
    const blob = new Blob([JSON.stringify(serializable, null, 2)], {
      type: 'application/json'
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'nomina_formato_largo.json';
    link.click();
    URL.revokeObjectURL(url);
  };

  const totalWarnings = files.filter((f) => f.warning).length;
  const hasCalculatedRows = mode === 'siigo' && consolidatedRows.some((r) => r._fill === 'yellow');

  // En el modo Siigo los ceros se muestran como "-" y los montos se colorean
  // (verde = viene de Siigo, amarillo = calculado), igual que en el ejemplo.
  const renderCell = (row, col) => {
    const value = row[col];
    if (mode === 'siigo' && VALUE_COLUMNS.has(col) && value === 0) return '-';
    return formatCellValue(value);
  };
  const cellClass = (row, col) => {
    const base = 'px-4 py-2.5 whitespace-nowrap';
    if (mode === 'siigo' && row._fill && VALUE_COLUMNS.has(col) && row[col] !== 0) {
      return `${base} text-right ${FILL_CLASS[row._fill]}`;
    }
    if (VALUE_COLUMNS.has(col)) return `${base} text-right`;
    return base;
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 font-sans">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 px-8 py-4 flex items-center justify-between shadow-sm">
        <div className="flex items-center gap-3">
          <img src="/logo.jpeg" alt="Logo" className="h-14 w-auto object-contain cursor-pointer" />
        </div>
        <div className="flex items-center gap-2.5 border border-slate-200 bg-slate-50 px-5 py-2 rounded-full text-sm font-semibold text-slate-700 cursor-pointer hover:bg-slate-100 transition-colors">
          <User className="w-4 h-4 text-blue-600" />
          <span>Bienvenido Usuario</span>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-6 py-10 space-y-8">
        {/* Título */}
        <div className="text-center space-y-3">
          <h1 className="text-3xl font-extrabold text-slate-900 tracking-tight">
            Transformador de Nómina: Ancho → Largo
          </h1>
          <p className="text-sm text-slate-600 max-w-2xl mx-auto leading-relaxed">
            Carga uno o varios archivos (.xlsx / .xlsm) y consolida todo en un único formato largo
            listo para el cruce contra Siigo. Elige abajo qué tipo de archivo vas a cargar.
          </p>
        </div>

        {/* Selector de modo */}
        <div className="flex justify-center">
          <div className="inline-flex rounded-xl border border-slate-200 bg-white p-1 shadow-sm">
            <button
              onClick={() => changeMode('nomina')}
              className={`px-4 py-2 text-xs font-semibold rounded-lg transition-colors cursor-pointer ${
                mode === 'nomina' ? 'bg-blue-900 text-white' : 'text-slate-600 hover:bg-slate-100'
              }`}
            >
              Nómina por bloques de mes
            </button>
            <button
              onClick={() => changeMode('siigo')}
              className={`px-4 py-2 text-xs font-semibold rounded-lg transition-colors cursor-pointer ${
                mode === 'siigo' ? 'bg-blue-900 text-white' : 'text-slate-600 hover:bg-slate-100'
              }`}
            >
              Movimiento CC (Siigo)
            </button>
          </div>
        </div>

        {/* Instrucciones */}
        <div className="bg-blue-50/60 border border-blue-200/80 rounded-2xl overflow-hidden transition-all duration-200 shadow-sm">
          <button
            onClick={() => setShowInstructions(!showInstructions)}
            className="w-full px-6 py-4 flex items-center justify-between text-left hover:bg-blue-100/30 transition-colors cursor-pointer"
          >
            <div className="flex items-center gap-3 text-blue-900 font-bold text-base">
              <Info className="w-5 h-5 text-blue-600 shrink-0" />
              <span>¿Cómo usar esta herramienta?</span>
            </div>
            {showInstructions ? (
              <ChevronUp className="w-5 h-5 text-blue-800" />
            ) : (
              <ChevronDown className="w-5 h-5 text-blue-800" />
            )}
          </button>
          {showInstructions && mode === 'nomina' && (
            <div className="px-6 pb-6 pt-2 border-t border-blue-100 text-xs text-slate-700 space-y-2.5 leading-relaxed">
              <p>
                <strong className="text-slate-900">1. Cargar archivos:</strong> puedes seleccionar
                varios archivos a la vez (una empresa puede tener varios .xlsx por mes, o puedes
                subir varias empresas juntas).
              </p>
              <p>
                <strong className="text-slate-900">2. Detección automática:</strong> la herramienta
                busca en cada hoja los bloques que contienen "EMPLOYEE CODE" / "NAME" y a partir de
                ahí identifica el mes y los conceptos de esa nómina, sin importar cuántas columnas
                traiga cada una.
              </p>
              <p>
                <strong className="text-slate-900">3. Filtros de calidad:</strong> se excluyen
                automáticamente los subtotales (Payments, Total, Fee, etc.), los conceptos en cero
                o vacíos, las columnas que no son montos (fechas, país, estado) y las filas sin
                empleado identificado.
              </p>
              <p>
                <strong className="text-slate-900">4. Mes elaboración y empleado:</strong> el mes se
                convierte a fecha real (primer día del mes), tolerando errores de digitación. El
                nombre del empleado se unifica por código, así todos los meses salen igual.
              </p>
              <p>
                <strong className="text-slate-900">5. Consolidado:</strong> todos los archivos
                cargados se acumulan en una sola tabla larga (Mes elaboración, Concepto, Empleado,
                Valor Concepto, Valor Totales, Empresa, Archivo). Puedes quitar un archivo si lo
                subiste por error.
              </p>
              <p>
                <strong className="text-slate-900">6. Exportar:</strong> descarga el resultado en
                Excel, CSV o JSON.
              </p>
            </div>
          )}
          {showInstructions && mode === 'siigo' && (
            <div className="px-6 pb-6 pt-2 border-t border-blue-100 text-xs text-slate-700 space-y-2.5 leading-relaxed">
              <p>
                <strong className="text-slate-900">1. Cargar el Movimiento CC:</strong> el archivo
                de Siigo con las columnas Comprobante, Fecha elaboración, Descripción, Tercero,
                Débito y Crédito. Se usa la primera hoja que tenga ese encabezado.
              </p>
              <p>
                <strong className="text-slate-900">2. Qué se toma de Siigo (verde):</strong> salario,
                subsidio de transporte, auxilio extralegal, vacaciones, licencia remunerada y los
                aportes de pensión, salud, cajas y ARL. Solo comprobantes CC-*; las facturas (FV) y
                notas crédito (NC) se ignoran.
              </p>
              <p>
                <strong className="text-slate-900">3. Qué se calcula (amarillo):</strong> 13TH
                SALARY, 14TH SALARY e INTEREST ON 14TH SALARY no existen en Siigo. Se calculan con
                la base de salario + transporte + vacaciones + licencia. Revísalos: si tu criterio
                contable es otro, el resultado puede diferir en algunos pesos.
              </p>
              <p>
                <strong className="text-slate-900">4. Qué se omite:</strong> prima de servicios,
                intereses de cesantías y consignación de cesantías (ya las cubren las provisiones).
                Lo que no se reconoce aparece como aviso bajo el archivo, nunca se pierde en
                silencio.
              </p>
              <p>
                <strong className="text-slate-900">5. Nombres de concepto:</strong> el auxilio
                extralegal sale como "Alloawance 4 (Other allowances)". Si un mes debe llevar otro
                nombre, se cambia en <code>ALLOWANCE_NAME_BY_MONTH</code> dentro de
                siigoConverter.js.
              </p>
            </div>
          )}
        </div>

        {/* Zona de carga */}
        <div className="bg-white border-2 border-dashed border-slate-300 rounded-2xl p-10 shadow-sm text-center relative hover:border-blue-500 transition-colors cursor-pointer">
          <input
            type="file"
            accept=".xlsx,.xlsm"
            multiple
            onChange={handleFileUpload}
            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
          />
          <div className="flex flex-col items-center gap-3">
            <div className="p-4 bg-blue-50 text-blue-600 rounded-full">
              <Upload className="w-7 h-7" />
            </div>
            <div>
              <p className="font-semibold text-base text-slate-800">
                {mode === 'siigo'
                  ? 'Arrastra el Movimiento CC de Siigo (.xlsx / .xlsm) o haz clic para buscar'
                  : 'Arrastra tus archivos de nómina (.xlsx / .xlsm) o haz clic para buscar'}
              </p>
              <p className="text-xs text-slate-500 mt-1">
                Puedes seleccionar varios archivos a la vez — se van sumando al consolidado
              </p>
            </div>
          </div>
        </div>

        {loading && (
          <div className="flex items-center justify-center gap-2 text-blue-700 py-4">
            <div className="w-4 h-4 border-2 border-blue-600 border-t-transparent rounded-full animate-spin" />
            <span className="text-xs font-medium">Leyendo y transformando archivos...</span>
          </div>
        )}

        {/* Lista de archivos cargados */}
        {files.length > 0 && !loading && (
          <div className="bg-white border border-slate-200 rounded-2xl shadow-sm p-4 space-y-2">
            <div className="flex items-center justify-between mb-1">
              <h2 className="text-sm font-bold text-slate-800">
                Archivos cargados ({files.length})
              </h2>
              <button
                onClick={clearAll}
                className="text-xs font-medium text-slate-500 hover:text-red-600 transition-colors cursor-pointer"
              >
                Quitar todos
              </button>
            </div>
            {files.map((f) => (
              <div
                key={f.fileId}
                className={`px-3 py-2 rounded-lg border text-xs space-y-1.5 ${
                  f.warning ? 'bg-amber-50 border-amber-200' : 'bg-slate-50 border-slate-200'
                }`}
              >
                <div className="flex items-center justify-between gap-3">
                  <div className="flex items-center gap-2 min-w-0">
                    {f.warning ? (
                      <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0" />
                    ) : (
                      <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
                    )}
                    <div className="min-w-0">
                      <p className="font-semibold text-slate-800 truncate">{f.fileName}</p>
                      <p className="text-slate-500">
                        {f.warning
                          ? f.warning
                          : `${f.rows.length} filas generadas · Empresa: ${f.empresa}`}
                      </p>
                    </div>
                  </div>
                  <button
                    onClick={() => removeFile(f.fileId)}
                    className="p-1 rounded hover:bg-white text-slate-400 hover:text-red-600 transition-colors cursor-pointer shrink-0"
                    title="Quitar este archivo"
                  >
                    <X className="w-4 h-4" />
                  </button>
                </div>
                {f.notes && f.notes.length > 0 && (
                  <ul className="pl-6 space-y-1">
                    {f.notes.map((note, idx) => (
                      <li
                        key={idx}
                        className={note.type === 'warn' ? 'text-amber-800' : 'text-slate-600'}
                      >
                        {note.text}
                      </li>
                    ))}
                  </ul>
                )}
              </div>
            ))}
            {totalWarnings > 0 && (
              <p className="text-xs text-amber-700 pt-1">
                {totalWarnings} archivo(s) no generaron filas — revisa que sean del formato
                esperado.
              </p>
            )}
          </div>
        )}

        {/* Tabla consolidada */}
        {consolidatedRows.length > 0 && !loading && (
          <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
            <div className="p-4 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-4">
              <div className="flex items-center gap-2 text-emerald-600 text-xs font-semibold">
                <CheckCircle2 className="w-4 h-4" />
                <span>{filteredData.length} registros consolidados</span>
              </div>

              <div className="relative flex-1 max-w-xs">
                <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  type="text"
                  placeholder="Buscar empleado, concepto, mes..."
                  value={searchTerm}
                  onChange={(e) => {
                    setSearchTerm(e.target.value);
                    setCurrentPage(1);
                  }}
                  className="w-full pl-9 pr-3 py-1.5 text-xs bg-white border border-slate-200 rounded-lg focus:outline-none focus:border-blue-500"
                />
              </div>

              <div className="flex items-center gap-2">
                <button
                  onClick={downloadXLSX}
                  disabled={exporting}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-blue-900 hover:bg-blue-800 disabled:opacity-50 text-white font-medium text-xs rounded-lg transition-colors cursor-pointer"
                >
                  <FileSpreadsheet className="w-3.5 h-3.5" />
                  {exporting ? 'Generando...' : 'Excel'}
                </button>
                <button
                  onClick={downloadCSV}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-emerald-700 hover:bg-emerald-800 text-white font-medium text-xs rounded-lg transition-colors cursor-pointer"
                >
                  <FileSpreadsheet className="w-3.5 h-3.5" />
                  CSV
                </button>
                <button
                  onClick={downloadJSON}
                  className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-700 hover:bg-slate-800 text-white font-medium text-xs rounded-lg transition-colors cursor-pointer"
                >
                  <Download className="w-3.5 h-3.5" />
                  JSON
                </button>
              </div>
            </div>

            {hasCalculatedRows && (
              <div className="px-4 py-2 bg-white border-b border-slate-200 flex flex-wrap items-center gap-4 text-xs text-slate-600">
                <span className="flex items-center gap-1.5">
                  <span className="inline-block w-3 h-3 rounded-sm bg-green-300" />
                  Viene de Siigo
                </span>
                <span className="flex items-center gap-1.5">
                  <span className="inline-block w-3 h-3 rounded-sm bg-yellow-300" />
                  Calculado (no existe en Siigo)
                </span>
              </div>
            )}

            <div className="overflow-x-auto max-h-96">
              <table className="w-full text-left text-xs text-slate-700">
                <thead className="bg-slate-100 uppercase text-slate-500 sticky top-0 border-b border-slate-200 font-semibold">
                  <tr>
                    {outputColumns.map((col) => (
                      <th key={col} className="px-4 py-3 whitespace-nowrap">
                        {col}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {paginatedData.map((row, idx) => (
                    <tr key={idx} className="hover:bg-slate-50 transition-colors">
                      {outputColumns.map((col) => (
                        <td key={col} className={cellClass(row, col)}>
                          {renderCell(row, col)}
                        </td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="p-3 bg-slate-50 border-t border-slate-200 flex items-center justify-between text-xs text-slate-600">
              <span>
                Página {currentPage} de {totalPages}
              </span>
              <div className="flex items-center gap-1">
                <button
                  disabled={currentPage === 1}
                  onClick={() => setCurrentPage((prev) => Math.max(prev - 1, 1))}
                  className="p-1.5 rounded border border-slate-200 hover:bg-white disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                >
                  <ChevronLeft className="w-4 h-4" />
                </button>
                <button
                  disabled={currentPage === totalPages}
                  onClick={() => setCurrentPage((prev) => Math.min(prev + 1, totalPages))}
                  className="p-1.5 rounded border border-slate-200 hover:bg-white disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                >
                  <ChevronRight className="w-4 h-4" />
                </button>
              </div>
            </div>
          </div>
        )}

        {files.length > 0 && consolidatedRows.length === 0 && !loading && (
          <div className="bg-white border border-slate-200 rounded-2xl shadow-sm p-10 text-center space-y-2">
            <FileX2 className="w-8 h-8 text-slate-300 mx-auto" />
            <p className="text-sm font-semibold text-slate-700">Ningún archivo generó registros</p>
            <p className="text-xs text-slate-500">
              {mode === 'siigo'
                ? 'Revisa los avisos de arriba: probablemente el archivo no es un Movimiento CC de Siigo.'
                : 'Revisa los mensajes de advertencia arriba: probablemente el archivo no trae la etiqueta "EMPLOYEE CODE" que la herramienta usa para detectar el formato.'}
            </p>
          </div>
        )}
      </main>
    </div>
  );
}