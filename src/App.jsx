import React, { useState, useMemo, useCallback, useRef, useEffect } from 'react';
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
import {
  convertNominaRows,
  convertNominaFiles,
  convertMovimientoRows,
  convertMovimientoFiles,
  STATUS_COLORS,
  cruzarCuenta28,
  CUENTA28_STATUS_COLORS
} from './nominaConverter';

// ============================================================================
// LECTOR DE .xlsx / .xlsm CON JSZIP (sin librería xlsx/SheetJS)
// ============================================================================
// Un archivo .xlsx/.xlsm es en realidad un .zip con archivos XML adentro.
// Aquí lo desempacamos con JSZip y leemos a mano los XML que necesitamos:
//   - xl/workbook.xml            -> lista de hojas y sus IDs de relación
//   - xl/_rels/workbook.xml.rels -> a qué archivo físico apunta cada hoja
//   - xl/sharedStrings.xml       -> tabla de textos compartidos
//   - xl/styles.xml              -> rellenos (colores) y a qué estilo apunta cada celda
//   - xl/worksheets/sheetN.xml   -> las celdas de cada hoja
//
// Cada celda se devuelve como { v: valor, f: relleno } donde f es el color de
// fondo ('RRGGBB', 'theme:N') o null. Los colores importan: en la nómina el
// equipo pinta el resultado del cruce con Siigo y ese color se conserva.
//
// Nota: los valores de celda se devuelven crudos (número, texto o booleano),
// tal como vienen en el XML. Una celda con formato de fecha llega como el
// número de serie de Excel, no como un objeto Date — quien la use (por
// ejemplo convertMovimientoRows) debe decodificarla.

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

// Paleta antigua de Excel (colores "indexed") para los pocos que se usan como relleno.
const INDEXED_COLORS = {
  2: 'FF0000',
  3: '00FF00',
  5: 'FFFF00',
  7: '00FFFF',
  10: 'FF0000',
  11: '00FF00',
  13: 'FFFF00',
  15: '00FFFF'
};

// Devuelve un array donde la posición N es el color de relleno del estilo N
// (cellXfs) o null si ese estilo no tiene relleno sólido.
function parseStyleFills(xmlDoc) {
  const fillColors = [];
  const fillsNode = xmlDoc.getElementsByTagName('fills')[0];
  if (fillsNode) {
    const fillNodes = fillsNode.getElementsByTagName('fill');
    for (let i = 0; i < fillNodes.length; i++) {
      const pattern = fillNodes[i].getElementsByTagName('patternFill')[0];
      let color = null;
      if (pattern && pattern.getAttribute('patternType') === 'solid') {
        const fg = pattern.getElementsByTagName('fgColor')[0];
        if (fg) {
          const rgb = fg.getAttribute('rgb');
          const indexed = fg.getAttribute('indexed');
          const theme = fg.getAttribute('theme');
          if (rgb) color = rgb.slice(-6).toUpperCase();
          else if (indexed !== null) color = INDEXED_COLORS[Number(indexed)] || `indexed:${indexed}`;
          else if (theme !== null) color = `theme:${theme}`;
        }
      }
      fillColors.push(color);
    }
  }

  const styleFills = [];
  const xfsNode = xmlDoc.getElementsByTagName('cellXfs')[0];
  if (xfsNode) {
    const xfNodes = xfsNode.getElementsByTagName('xf');
    for (let i = 0; i < xfNodes.length; i++) {
      const fillId = Number(xfNodes[i].getAttribute('fillId') || 0);
      styleFills.push(fillColors[fillId] ?? null);
    }
  }
  return styleFills;
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

function parseSheetXmlToRows(xmlDoc, sharedStrings, styleFills) {
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
      const styleIdx = Number(cellNode.getAttribute('s') || 0);
      const fill = styleFills[styleIdx] ?? null;

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

      rowArray[colIndex] = { v: value, f: fill };
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

  let styleFills = [];
  const stylesFile = zip.file('xl/styles.xml');
  if (stylesFile) {
    const stylesXmlText = await stylesFile.async('text');
    const stylesXmlDoc = parser.parseFromString(stylesXmlText, 'text/xml');
    styleFills = parseStyleFills(stylesXmlDoc);
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
    const rows = parseSheetXmlToRows(sheetXmlDoc, sharedStrings, styleFills);
    sheets.push({ name, rows });
  }

  return sheets;
}

// ============================================================================
// UTILIDADES
// ============================================================================

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

function companyNameFromFileName(fileName) {
  return fileName.replace(/\.[^/.]+$/, '');
}

function newFileId(file) {
  return `${file.name}-${file.size}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

// Primer renglón con varios textos de cada hoja: sirve para decirle al usuario
// qué encabezados vio la app cuando no reconoce el formato de un archivo.
function headerHint(sheets) {
  for (const { name, rows } of sheets) {
    for (let r = 0; r < Math.min(rows.length, 60); r++) {
      const cells = (rows[r] || [])
        .map((c) => (c && typeof c.v === 'string' ? c.v.trim() : ''))
        .filter(Boolean);
      if (cells.length >= 4) {
        return `hoja "${name}", fila ${r + 1}: ${cells.slice(0, 8).join(' | ')}`;
      }
    }
  }
  return null;
}

// ============================================================================
// PROCESAMIENTO: Nómina (formato ancho) o Movimiento CC -> formato largo
// ============================================================================
// La lógica (qué es un bloque, cómo salen los conceptos, los colores) vive en
// nominaConverter.js. Aquí solo se lee el libro y se intenta reconocer el
// formato: primero nómina (bloques EMPLOYEE CODE / NAME); si ninguna hoja
// tiene eso, se intenta como Movimiento CC de Siigo (Comprobante / Fecha
// elaboración / Descripción / Débito / Crédito).

async function readNominaFile(file) {
  const sheets = await readWorkbookSheetsWithJSZip(file);
  return {
    fileId: newFileId(file),
    fileName: file.name,
    empresa: companyNameFromFileName(file.name),
    sheets
  };
}

function convertLoadedFile(loaded, unifyNames) {
  if (loaded.readError) {
    return {
      ...loaded,
      rows: [],
      notes: [],
      warning: loaded.readError,
      sourceType: null,
      nominaRows: null,
      movimientoRows: null
    };
  }

  let converted = null;
  let sourceType = null;
  let nominaRows = null; // filas de la hoja de nómina usada (para consolidar varios archivos)
  let movimientoRows = null; // filas de la hoja de Movimiento CC usada (para consolidar varios archivos)

  for (const { rows } of loaded.sheets) {
    const result = convertNominaRows(rows, { unifyNamesByCode: unifyNames });
    if (result) {
      converted = result;
      sourceType = 'nomina';
      nominaRows = rows;
      break; // la primera hoja con bloques de nómina (las demás son auxiliares o el formato largo)
    }
  }

  if (!converted) {
    for (const { rows } of loaded.sheets) {
      const result = convertMovimientoRows(rows);
      if (result) {
        converted = result;
        sourceType = 'movimiento';
        movimientoRows = rows;
        break; // la primera hoja que sea un Movimiento CC (Comprobante/Fecha/Descripción/Débito/Crédito)
      }
    }
  }

  let warning = null;
  if (!converted) {
    warning =
      'No se reconoció el formato del archivo: ni bloques de nómina (encabezado con EMPLOYEE CODE y NAME) ni un Movimiento CC de Siigo (encabezado con Comprobante, Fecha elaboración, Descripción, Débito y Crédito).';
    const hint = headerHint(loaded.sheets);
    if (hint) {
      warning += ` Primer encabezado que vi: ${hint}. Si es una nómina, agrega esos nombres a los alias de nominaConverter.js.`;
    }
  } else if (converted.records.length === 0) {
    warning = 'Se reconoció el formato, pero no se generó ningún registro con valor.';
  }

  return {
    ...loaded,
    rows: converted ? converted.records : [],
    notes: converted ? converted.notes : [],
    warning,
    sourceType,
    nominaRows,
    movimientoRows
  };
}

// ============================================================================
// ESCRITOR DE .xlsx CON JSZIP (para la descarga del resultado)
// ============================================================================
// Igual que para leer, generamos a mano el XML mínimo que necesita un .xlsx
// válido. Estilos (atributo s): 0 = general; 1 = fecha mmm-yy alineada a la
// izquierda; 2 = número contable sin relleno; 3 en adelante = número contable
// con el relleno de cada color de STATUS_COLORS (en el mismo orden).

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

const VALUE_COLUMNS = new Set([
  'Valor Concepto', 'Valor Totales',
  'Valor Nómina', 'Valor Siigo', 'Diferencia', 'Nómina mes', 'Siigo mes'
]);
const FILL_HEXES = Object.keys(STATUS_COLORS);
const NUMBER_STYLE_PLAIN = 2;
const NUMBER_STYLE_FIRST_FILL = 3;
const STATUS_STYLE_FIRST_FILL = NUMBER_STYLE_FIRST_FILL + FILL_HEXES.length;

// Ancho de cada columna en el Excel exportado (en caracteres).
const EXCEL_COLUMN_WIDTH = {
  'Mes elaboración': 15,
  Concepto: 44,
  Empleado: 34,
  'Valor Concepto': 16,
  'Valor Totales': 16
};

function numberStyleFor(row) {
  const idx = row._fill ? FILL_HEXES.indexOf(row._fill) : -1;
  return idx >= 0 ? NUMBER_STYLE_FIRST_FILL + idx : NUMBER_STYLE_PLAIN;
}

function buildSheetXml(dataRows, columns) {
  const colsXml = columns
    .map((colName, idx) => {
      const width = EXCEL_COLUMN_WIDTH[colName] || 20;
      return `<col min="${idx + 1}" max="${idx + 1}" width="${width}" customWidth="1"/>`;
    })
    .join('');

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
          const style = VALUE_COLUMNS.has(colName) ? numberStyleFor(row) : 0;
          return `<c r="${ref}" s="${style}"><v>${val}</v></c>`;
        }
        const statusStyle = colName === 'Estado' && row._fill
          ? STATUS_STYLE_FIRST_FILL + Math.max(0, FILL_HEXES.indexOf(row._fill))
          : 0;
        return `<c r="${ref}" s="${statusStyle}" t="inlineStr"><is><t>${xmlEscape(String(val))}</t></is></c>`;
      })
      .join('');
    xmlRows += `<row r="${rowNum}">${cells}</row>`;
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>${colsXml}</cols><sheetData>${xmlRows}</sheetData></worksheet>`;
}

function buildStylesXml() {
  const fills =
    '<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    FILL_HEXES.map(
      (hex) =>
        `<fill><patternFill patternType="solid"><fgColor rgb="FF${hex}"/><bgColor indexed="64"/></patternFill></fill>`
    ).join('');
  const numberXf = (fillId) =>
    `<xf numFmtId="164" fontId="0" fillId="${fillId}" borderId="0" xfId="0" applyNumberFormat="1"${
      fillId ? ' applyFill="1"' : ''
    }/>`;
  const textFillXf = (fillId) =>
    `<xf numFmtId="0" fontId="0" fillId="${fillId}" borderId="0" xfId="0" applyFill="1"/>`;
  const cellXfs =
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="17" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="left"/></xf>' +
    numberXf(0) +
    FILL_HEXES.map((_, i) => numberXf(i + 2)).join('') +
    FILL_HEXES.map((_, i) => textFillXf(i + 2)).join('');

  // numFmtId 164 = formato contable (ceros como "-"), igual que la Hoja2 de ejemplo;
  // numFmtId 17 = mmm-yy (integrado).
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="_-* #,##0.00_-;\\-* #,##0.00_-;_-* &quot;-&quot;??_-;_-@_-"/></numFmts>' +
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
    `<fills count="${FILL_HEXES.length + 2}">${fills}</fills>` +
    '<borders count="1"><border/></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    `<cellXfs count="${FILL_HEXES.length * 2 + 3}">${cellXfs}</cellXfs>` +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '</styleSheet>'
  );
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

  zip.file('[Content_Types].xml', contentTypes);
  zip.file('_rels/.rels', rootRels);
  zip.file('xl/workbook.xml', workbookXml);
  zip.file('xl/_rels/workbook.xml.rels', workbookRels);
  zip.file('xl/styles.xml', buildStylesXml());
  zip.file('xl/worksheets/sheet1.xml', buildSheetXml(dataRows, columns));

  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
}

// ============================================================================
// COMPONENTE PRINCIPAL
// ============================================================================

// Columnas de salida, en el orden y con los nombres exactos del ejemplo real
// (Hoja2 de BUBBLE_-_Nómina_1.xlsm).
const OUTPUT_COLUMNS = [
  'Mes elaboración',
  'Concepto',
  'Empleado',
  'Valor Concepto',
  'Valor Totales'
];

const CUENTA28_COLUMNS = [
  'Empresa',
  'Mes',
  'Empleado',
  'Concepto',
  'Valor Nómina',
  'Valor Siigo',
  'Diferencia',
  'Estado',
  'Nómina mes',
  'Siigo mes'
];

// Ancho mínimo de cada columna en la tabla de la pantalla (px).
const COLUMN_MIN_WIDTH = {
  'Mes elaboración': 150,
  Concepto: 300,
  Empleado: 280,
  'Valor Concepto': 170,
  'Valor Totales': 170
};

// Etiqueta que se muestra junto al nombre del archivo, según qué formato se reconoció.
const SOURCE_TYPE_LABEL = {
  nomina: 'Nómina',
  movimiento: 'Movimiento CC'
};

export default function App() {
  const [showInstructions, setShowInstructions] = useState(true);
  const [files, setFiles] = useState([]); // { fileId, fileName, empresa, sheets } | { ..., readError }
  const [unifyNames, setUnifyNames] = useState(true);
  const [modo, setModo] = useState('cuenta28');
  const [estadoFilter, setEstadoFilter] = useState('TODOS');
  const [mesFilter, setMesFilter] = useState('TODOS');
  const [conceptoFilter, setConceptoFilter] = useState('TODOS');
  const [empresaFilter, setEmpresaFilter] = useState('TODOS');
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const rowsPerPage = 10;

  // Para bajar solos hasta el resultado cuando termina de procesar.
  const fileListRef = useRef(null);
  const resultsRef = useRef(null);
  const scrollPending = useRef(false);

  // Los archivos se leen una vez; la conversión se recalcula si cambia la opción de nombres.
  const processed = useMemo(
    () => files.map((f) => convertLoadedFile(f, unifyNames)),
    [files, unifyNames]
  );

  const empresas = useMemo(
    () => [...new Set(processed.map((f) => f.empresa).filter(Boolean))].sort((a, b) => a.localeCompare(b)),
    [processed]
  );

  const sourceRows = useMemo(() => {
    const relevant = empresaFilter === 'TODOS'
      ? processed
      : processed.filter((f) => f.empresa === empresaFilter);
    const nominaRows = relevant
      .filter((f) => f.sourceType === 'nomina')
      .flatMap((f) => f.rows);
    const movimientoRows = relevant
      .filter((f) => f.sourceType === 'movimiento')
      .flatMap((f) => f.rows);
    return { nominaRows, movimientoRows };
  }, [processed, empresaFilter]);

  const cuenta28 = useMemo(() => {
    // El cruce se hace por empresa. No mezclamos personas/conceptos de compañías
    // distintas aunque el usuario haya cargado todos los archivos juntos.
    const relevant = empresaFilter === 'TODOS' ? processed : processed.filter((f) => f.empresa === empresaFilter);
    if (empresaFilter !== 'TODOS') {
      return cruzarCuenta28(sourceRows.nominaRows, sourceRows.movimientoRows);
    }

    const groups = new Map();
    for (const f of relevant) {
      if (!groups.has(f.empresa)) groups.set(f.empresa, { nomina: [], movimiento: [] });
      if (f.sourceType === 'nomina') groups.get(f.empresa).nomina.push(...f.rows);
      if (f.sourceType === 'movimiento') groups.get(f.empresa).movimiento.push(...f.rows);
    }

    const rows = [];
    const summary = { total: 0, OK: 0, DIFERENCIA: 0, SOLO_NOMINA: 0, SOLO_SIIGO: 0, CRUZA_ENTRE_MESES: 0 };
    const months = new Set();
    for (const [empresa, group] of groups) {
      const result = cruzarCuenta28(group.nomina, group.movimiento);
      for (const row of result.rows) rows.push({ ...row, Empresa: empresa });
      for (const m of result.months) months.add(m);
      Object.keys(summary).forEach((key) => { summary[key] += result.summary[key] || 0; });
    }
    return { rows, summary, months: [...months].sort() };
  }, [processed, empresaFilter, sourceRows]);

  const consolidatedRows = useMemo(() => {
    if (modo === 'cuenta28') {
      return cuenta28.rows.map((row) => ({ ...row, Empresa: row.Empresa || empresaFilter }));
    }
    return [...sourceRows.nominaRows, ...sourceRows.movimientoRows];
  }, [modo, cuenta28.rows, sourceRows]);

  // Colores presentes en el resultado, para la leyenda.
  const legendColors = useMemo(() => {
    const used = new Set();
    for (const row of consolidatedRows) if (row._fill) used.add(row._fill);
    return Object.keys(STATUS_COLORS).filter((hex) => used.has(hex));
  }, [consolidatedRows]);

  const cuenta28Months = cuenta28.months;
  const cuenta28Concepts = useMemo(
    () => [...new Set(cuenta28.rows.map((r) => r.Concepto))].sort((a, b) => a.localeCompare(b)),
    [cuenta28.rows]
  );

  useEffect(() => {
    if (loading || !scrollPending.current) return;
    const target = resultsRef.current || fileListRef.current;
    if (target) {
      scrollPending.current = false;
      target.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }, [files, loading]);

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
          results.push(await readNominaFile(file));
        } catch (err) {
          console.error('Error procesando', file.name, err);
          results.push({
            fileId: `${file.name}-${Date.now()}`,
            fileName: file.name,
            empresa: companyNameFromFileName(file.name),
            sheets: [],
            readError: 'No se pudo leer este archivo. ¿Es un .xlsx/.xlsm válido?'
          });
        }
      }
      scrollPending.current = true;
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
    setEmpresaFilter('TODOS');
    setEstadoFilter('TODOS');
    setMesFilter('TODOS');
    setConceptoFilter('TODOS');
  };

  const filteredData = useMemo(() => {
    let data = consolidatedRows;
    if (modo === 'cuenta28') {
      if (estadoFilter !== 'TODOS') data = data.filter((r) => r.Estado === estadoFilter);
      if (mesFilter !== 'TODOS') data = data.filter((r) => r.Mes === mesFilter);
      if (conceptoFilter !== 'TODOS') data = data.filter((r) => r.Concepto === conceptoFilter);
    }
    if (!searchTerm.trim()) return data;
    const term = searchTerm.toLowerCase();
    return data.filter((row) => {
      const values = modo === 'cuenta28'
        ? [row.Mes, row.Empleado, row.Concepto, row['Valor Nómina'], row['Valor Siigo'], row.Diferencia, row.Estado]
        : OUTPUT_COLUMNS.map((col) => row[col]);
      return values.some((v) => formatCellValue(v).toLowerCase().includes(term));
    });
  }, [consolidatedRows, searchTerm, modo, estadoFilter, mesFilter, conceptoFilter]);

  const totalPages = Math.ceil(filteredData.length / rowsPerPage) || 1;
  const paginatedData = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return filteredData.slice(start, start + rowsPerPage);
  }, [filteredData, currentPage]);

  const downloadXLSX = async () => {
    if (filteredData.length === 0) return;
    setExporting(true);
    try {
      const exportColumns = modo === 'cuenta28' ? CUENTA28_COLUMNS : OUTPUT_COLUMNS;
      const blob = await buildXlsxBlobWithJSZip(filteredData, exportColumns);
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = modo === 'cuenta28' ? 'cuenta_28_cruce.xlsx' : 'nomina_formato_largo.xlsx';
      link.click();
      URL.revokeObjectURL(url);
    } finally {
      setExporting(false);
    }
  };

  const downloadCSV = () => {
    if (filteredData.length === 0) return;
    const columns = modo === 'cuenta28' ? CUENTA28_COLUMNS : OUTPUT_COLUMNS;
    const headerLine = columns.join(',');
    const lines = filteredData.map((row) =>
      columns.map((col) =>
        typeof row[col] === 'number' ? String(row[col]) : `"${formatCellValue(row[col]).replace(/"/g, '""')}"`
      ).join(',')
    );
    const csvContent = [headerLine, ...lines].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = modo === 'cuenta28' ? 'cuenta_28_cruce.csv' : 'nomina_formato_largo.csv';
    link.click();
    URL.revokeObjectURL(url);
  };

  const downloadJSON = () => {
    if (filteredData.length === 0) return;
    const columns = modo === 'cuenta28' ? CUENTA28_COLUMNS : OUTPUT_COLUMNS;
    const serializable = filteredData.map((row) => {
      const obj = {};
      columns.forEach((col) => {
        obj[col] = row[col] instanceof Date ? formatCellValue(row[col]) : row[col];
      });
      return obj;
    });
    const blob = new Blob([JSON.stringify(serializable, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = modo === 'cuenta28' ? 'cuenta_28_cruce.json' : 'nomina_formato_largo.json';
    link.click();
    URL.revokeObjectURL(url);
  };

  const totalWarnings = processed.filter((f) => f.warning).length;

  const tableColumns = modo === 'cuenta28' ? CUENTA28_COLUMNS : OUTPUT_COLUMNS;
  const renderCell = (row, col) => {
    const value = row[col];
    if (modo === 'cuenta28' && ['Valor Nómina', 'Valor Siigo', 'Diferencia', 'Nómina mes', 'Siigo mes'].includes(col)) {
      return Number(value || 0).toLocaleString('es-CO');
    }
    if (modo !== 'cuenta28' && VALUE_COLUMNS.has(col) && value === 0) return '-';
    return formatCellValue(value);
  };
  const cellClass = (col) => {
    const base = 'px-4 py-2.5 whitespace-nowrap';
    return (modo === 'cuenta28' && ['Valor Nómina', 'Valor Siigo', 'Diferencia', 'Nómina mes', 'Siigo mes'].includes(col)) ||
      (modo !== 'cuenta28' && VALUE_COLUMNS.has(col)) ? `${base} text-right` : base;
  };
  const cellStyle = (row, col) => {
    const style = { minWidth: modo === 'cuenta28' ? 150 : COLUMN_MIN_WIDTH[col] };
    if (modo === 'cuenta28' && col === 'Estado' && row._fill) style.backgroundColor = `#${row._fill}`;
    if (modo !== 'cuenta28' && VALUE_COLUMNS.has(col) && row._fill) style.backgroundColor = `#${row._fill}`;
    return style;
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
            Automatización y cruce de Cuenta 28
          </h1>
          <p className="text-sm text-slate-600 max-w-2xl mx-auto leading-relaxed">
            Carga la nómina y el Movimiento CC de Siigo. La herramienta convierte ambos formatos,
            cruza mes + empleado + concepto y muestra cruces correctos, diferencias, registros que
            solo aparecen en un lado y conceptos que cruzan entre meses.
          </p>
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
          {showInstructions && (
            <div className="px-6 pb-6 pt-2 border-t border-blue-100 text-xs text-slate-700 space-y-2.5 leading-relaxed">
              <p>
                <strong className="text-slate-900">1. Cargar el archivo:</strong> puede ser la
                nómina que se envía a facturación (un bloque por mes, encabezado con EMPLOYEE CODE
                y NAME) o el Movimiento CC de Siigo (encabezado con Comprobante, Fecha elaboración,
                Descripción, Débito y Crédito). La app prueba primero como nómina y, si no
                encuentra bloques, como Movimiento CC. Puedes seleccionar varios archivos a la vez,
                incluso mezclando los dos tipos. Si son varios archivos de una misma empresa —
                varias nóminas (p. ej. los de RemoFirst) o varios Movimiento CC (p. ej. uno por
                rango de fechas) — se pueden cargar juntos; la herramienta los agrupa por empresa
                antes de hacer el cruce.
              </p>
              <p>
                <strong className="text-slate-900">2. Qué sale de la nómina:</strong> una fila por
                mes, empleado y concepto, con los nombres de concepto tal como aparecen en el
                encabezado de la nómina, y al final de cada empleado su TOTAL EMPLOYEE COST. Los
                conceptos en cero no se listan.
              </p>
              <p>
                <strong className="text-slate-900">3. Qué sale del Movimiento CC:</strong> cada
                Descripción se clasifica a un concepto tipo nómina (SALARY, PENSION COST, HEALTH
                COST…) con la tabla CONCEPT_KEYWORDS, se suma por mes, concepto y empleado (Débito −
                Crédito) y se agrega TOTAL EMPLOYEE COST por mes y empleado. Los aportes patronales,
                y las filas de salario/prestación que vengan sin Tercero, se asignan al empleado del
                mismo día (si hay un único candidato); si no se puede, salen como "(sin asignar)".
                Lo que no se reconoce queda con el texto de Descripción y se avisa.
              </p>
              <p>
                <strong className="text-slate-900">4. Cruce Cuenta 28:</strong> se compara por empresa,
                mes, empleado y concepto. Verde = OK, rojo = diferencia, amarillo = solo aparece en
                un lado y verde = cruza entre meses para conceptos acumulativos como la prima.
              </p>
              <p>
                <strong className="text-slate-900">5. Colores del archivo:</strong> son los que ya trae el
                archivo (el resultado del cruce con Siigo): verde y azul = cruce ok, amarillo = no
                está en el otro lado, rojo = diferencias, verde limón = cruza entre meses, morado =
                débito y crédito se anulan. Si el archivo aún no está pintado, las filas salen sin
                color.
              </p>
              <p>
                <strong className="text-slate-900">6. Los avisos</strong> bajo cada archivo indican
                qué formato se reconoció y si algo no cuadra (un total que no coincide, colores
                fuera de la leyenda, filas sin fecha válida, etc.). Nada se descarta en silencio.
              </p>
              <p>
                <strong className="text-slate-900">7. Exportar:</strong> descarga el resultado
                consolidado en Excel (con los colores y el formato contable), CSV o JSON.
              </p>
            </div>
          )}
        </div>

        {/* Modo y filtros Cuenta 28 */}
        <div className="bg-white border border-slate-200 rounded-2xl p-5 shadow-sm space-y-4">
          <div className="flex flex-wrap items-center gap-3">
            <span className="text-sm font-bold text-slate-800">Modo:</span>
            <button onClick={() => setModo('cuenta28')} className={`px-4 py-2 rounded-lg text-xs font-semibold ${modo === 'cuenta28' ? 'bg-blue-900 text-white' : 'bg-slate-100 text-slate-600'}`}>Cuenta 28</button>
            <button onClick={() => setModo('largo')} className={`px-4 py-2 rounded-lg text-xs font-semibold ${modo === 'largo' ? 'bg-blue-900 text-white' : 'bg-slate-100 text-slate-600'}`}>Formato largo</button>
          </div>
          {modo === 'cuenta28' && (
            <>
              <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
                <select value={empresaFilter} onChange={(e) => { setEmpresaFilter(e.target.value); setCurrentPage(1); }} className="border border-slate-200 rounded-lg px-3 py-2 text-xs">
                  <option value="TODOS">Todas las empresas</option>
                  {empresas.map((e) => <option key={e} value={e}>{e}</option>)}
                </select>
                <select value={estadoFilter} onChange={(e) => { setEstadoFilter(e.target.value); setCurrentPage(1); }} className="border border-slate-200 rounded-lg px-3 py-2 text-xs">
                  <option value="TODOS">Todos los estados</option>
                  <option value="OK">🟢 Cruce OK</option>
                  <option value="DIFERENCIA">🔴 Diferencias</option>
                  <option value="SOLO_NOMINA">🟡 Solo nómina</option>
                  <option value="SOLO_SIIGO">🟡 Solo Siigo</option>
                  <option value="CRUZA_ENTRE_MESES">🟢 Cruza entre meses</option>
                </select>
                <select value={mesFilter} onChange={(e) => { setMesFilter(e.target.value); setCurrentPage(1); }} className="border border-slate-200 rounded-lg px-3 py-2 text-xs">
                  <option value="TODOS">Todos los meses</option>
                  {cuenta28Months.map((m) => <option key={m} value={m}>{m}</option>)}
                </select>
                <select value={conceptoFilter} onChange={(e) => { setConceptoFilter(e.target.value); setCurrentPage(1); }} className="border border-slate-200 rounded-lg px-3 py-2 text-xs">
                  <option value="TODOS">Todos los conceptos</option>
                  {cuenta28Concepts.map((c) => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
                <div className="grid grid-cols-2 md:grid-cols-6 gap-2">
                {[['Total', cuenta28.summary.total], ['OK', cuenta28.summary.OK], ['Diferencias', cuenta28.summary.DIFERENCIA], ['Solo nómina', cuenta28.summary.SOLO_NOMINA], ['Solo Siigo', cuenta28.summary.SOLO_SIIGO], ['Cruza meses', cuenta28.summary.CRUZA_ENTRE_MESES]].map(([label, value]) => (
                  <div key={label} className="bg-slate-50 border border-slate-200 rounded-lg p-3">
                    <div className="text-[11px] text-slate-500">{label}</div>
                    <div className="text-xl font-extrabold text-slate-800">{value}</div>
                  </div>
                ))}
              </div>
            </>
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
                Arrastra la nómina o el Movimiento CC (.xlsx / .xlsm) o haz clic para buscar
              </p>
              <p className="text-xs text-slate-500 mt-1">
                Puedes seleccionar varios archivos a la vez, incluso mezclando los dos tipos — se
                van sumando al consolidado
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
          <div
            ref={fileListRef}
            className="bg-white border border-slate-200 rounded-2xl shadow-sm p-4 space-y-2 scroll-mt-6"
          >
            <div className="flex items-center justify-between mb-1">
              <h2 className="text-sm font-bold text-slate-800">
                Archivos cargados ({files.length})
              </h2>
              <div className="flex items-center gap-4 flex-wrap justify-end">
                <label className="flex items-center gap-1.5 text-xs text-slate-600 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={unifyNames}
                    onChange={(e) => setUnifyNames(e.target.checked)}
                  />
                  Unificar el nombre de cada empleado por código (solo nómina)
                </label>
                <button
                  onClick={clearAll}
                  className="text-xs font-medium text-slate-500 hover:text-red-600 transition-colors cursor-pointer"
                >
                  Quitar todos
                </button>
              </div>
            </div>
            {processed.map((f) => (
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
                      <p className="font-semibold text-slate-800 truncate">
                        {f.fileName}
                        {f.sourceType && (
                          <span className="ml-2 inline-block px-1.5 py-0.5 rounded text-[10px] font-semibold uppercase tracking-wide bg-slate-200 text-slate-600 align-middle">
                            {SOURCE_TYPE_LABEL[f.sourceType]}
                          </span>
                        )}
                      </p>
                      <p className="text-slate-500">
                        {f.warning ? f.warning : `${f.rows.length} filas generadas`}
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
                {totalWarnings} archivo(s) no generaron filas — revisa que sean la nómina o el
                Movimiento CC en el formato esperado.
              </p>
            )}
          </div>
        )}

        {/* Tabla consolidada */}
        {consolidatedRows.length > 0 && !loading && (
          <div ref={resultsRef} className="space-y-6 scroll-mt-6">
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
                    placeholder="Buscar empleado, concepto, mes, estado..."
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

              {legendColors.length > 0 && (
                <div className="px-4 py-2 bg-white border-b border-slate-200 flex flex-wrap items-center gap-4 text-xs text-slate-600">
                  {legendColors.map((hex) => (
                    <span key={hex} className="flex items-center gap-1.5">
                      <span
                        className="inline-block w-3 h-3 rounded-sm"
                        style={{ backgroundColor: `#${hex}` }}
                      />
                      {STATUS_COLORS[hex]}
                    </span>
                  ))}
                </div>
              )}

              <div className="overflow-x-auto max-h-[28rem]">
                <table className="w-full text-left text-xs text-slate-700">
                  <thead className="bg-slate-100 uppercase text-slate-500 sticky top-0 border-b border-slate-200 font-semibold">
                    <tr>
                      {tableColumns.map((col) => (
                        <th
                          key={col}
                          className="px-5 py-3 whitespace-nowrap"
                          style={{ minWidth: COLUMN_MIN_WIDTH[col] }}
                        >
                          {col}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {paginatedData.map((row, idx) => (
                      <tr key={idx} className="hover:bg-slate-50 transition-colors">
                        {tableColumns.map((col) => (
                          <td key={col} className={cellClass(col)} style={cellStyle(row, col)}>
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
          </div>
        )}

        {files.length > 0 && consolidatedRows.length === 0 && !loading && (
          <div className="bg-white border border-slate-200 rounded-2xl shadow-sm p-10 text-center space-y-2">
            <FileX2 className="w-8 h-8 text-slate-300 mx-auto" />
            <p className="text-sm font-semibold text-slate-700">Ningún archivo generó registros</p>
            <p className="text-xs text-slate-500">
              Revisa los avisos de arriba: probablemente el archivo no es la nómina ni el
              Movimiento CC en el formato que la app reconoce.
            </p>
          </div>
        )}
      </main>
    </div>
  );
}