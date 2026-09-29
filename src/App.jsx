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
  FileX2,
  Plus,
  Trash2,
  FileText
} from 'lucide-react';
import {
  convertNominaRows,
  convertNominaFiles,
  convertMovimientoRows,
  convertMovimientoFiles,
  STATUS_COLORS
} from './nominaConverter';
import {
  convertFacturacionRows,
  convertFacturacionFiles
} from './facturacionConverter';

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
// ejemplo convertMovimientoRows o convertFacturacionRows) debe decodificarla.

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
// PROCESAMIENTO: Facturación EOR / Nómina (formato ancho) / Movimiento CC -> formato largo
// ============================================================================
// La lógica de cada formato vive en su conversor:
//   - facturacionConverter.js : hoja INVOICING de facturación EOR
//   - nominaConverter.js      : nómina (bloques por mes) y Movimiento CC de Siigo
// Aquí solo se lee el libro y se intenta reconocer el formato, del más
// específico al más general:
//   1) Facturación EOR: encabezado con EMPLOYEE CODE, NAME y PAYROLL MONTH
//      (+ SERVICE TYPE / CUSTOMER NAME). Va primero porque también tiene
//      EMPLOYEE CODE y NAME y, si no, la nómina se la quedaría.
//   2) Nómina: bloques EMPLOYEE CODE / NAME.
//   3) Movimiento CC de Siigo: Comprobante / Fecha elaboración / Descripción /
//      Débito / Crédito.

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
      movimientoRows: null,
      facturacionRows: null
    };
  }

  let converted = null;
  let sourceType = null;
  let nominaRows = null; // filas de la hoja de nómina usada (para consolidar varios archivos)
  let movimientoRows = null; // filas de la hoja de Movimiento CC usada (para consolidar varios archivos)
  let facturacionRows = null; // filas de la hoja de facturación usada (para consolidar varios archivos)

  // 1) Facturación EOR (el nombre del archivo aporta el año: "... Agosto 2026.xlsx")
  for (const { rows } of loaded.sheets) {
    const result = convertFacturacionRows(rows, { fileName: loaded.fileName });
    if (result) {
      converted = result;
      sourceType = 'facturacion';
      facturacionRows = rows;
      break; // la primera hoja con encabezado de facturación (la otra suele ser una tabla dinámica)
    }
  }

  // 2) Nómina
  if (!converted) {
    for (const { rows } of loaded.sheets) {
      const result = convertNominaRows(rows, { unifyNamesByCode: unifyNames });
      if (result) {
        converted = result;
        sourceType = 'nomina';
        nominaRows = rows;
        break; // la primera hoja con bloques de nómina (las demás son auxiliares o el formato largo)
      }
    }
  }

  // 3) Movimiento CC
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
      'No se reconoció el formato del archivo: ni facturación EOR (encabezado con EMPLOYEE CODE, NAME y Payroll Month), ni bloques de nómina (EMPLOYEE CODE y NAME), ni un Movimiento CC de Siigo (Comprobante, Fecha elaboración, Descripción, Débito y Crédito).';
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
    movimientoRows,
    facturacionRows
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

const VALUE_COLUMNS = new Set(['Valor Concepto', 'Valor Totales']);
const FILL_HEXES = Object.keys(STATUS_COLORS);
const NUMBER_STYLE_PLAIN = 2;
const NUMBER_STYLE_FIRST_FILL = 3;

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
        return `<c r="${ref}" t="inlineStr"><is><t>${xmlEscape(String(val))}</t></is></c>`;
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
  const cellXfs =
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="17" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="left"/></xf>' +
    numberXf(0) +
    FILL_HEXES.map((_, i) => numberXf(i + 2)).join('');

  // numFmtId 164 = formato contable (ceros como "-"), igual que la Hoja2 de ejemplo;
  // numFmtId 17 = mmm-yy (integrado).
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="_-* #,##0.00_-;\\-* #,##0.00_-;_-* &quot;-&quot;??_-;_-@_-"/></numFmts>' +
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
    `<fills count="${FILL_HEXES.length + 2}">${fills}</fills>` +
    '<borders count="1"><border/></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    `<cellXfs count="${FILL_HEXES.length + 3}">${cellXfs}</cellXfs>` +
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
  movimiento: 'Movimiento CC',
  facturacion: 'Facturación'
};

// Color de la etiqueta de cada tipo de archivo.
const SOURCE_TYPE_BADGE = {
  nomina: 'bg-indigo-50 text-indigo-700 ring-indigo-200',
  movimiento: 'bg-sky-50 text-sky-700 ring-sky-200',
  facturacion: 'bg-teal-50 text-teal-700 ring-teal-200'
};

const KNOWN_SOURCE_TYPES = ['nomina', 'movimiento', 'facturacion'];

// Encabezado de cada tarjeta: número de paso + título + acciones a la derecha.
function SectionHeader({ step, title, subtitle, children }) {
  return (
    <div className="flex flex-wrap items-center justify-between gap-3 px-6 py-4 border-b border-slate-100">
      <div className="flex items-center gap-3 min-w-0">
        <span className="flex items-center justify-center w-7 h-7 rounded-full bg-blue-900 text-white text-xs font-bold shrink-0">
          {step}
        </span>
        <div className="min-w-0">
          <h2 className="text-base font-bold text-slate-900 leading-tight">{title}</h2>
          {subtitle && <p className="text-xs text-slate-500 mt-0.5">{subtitle}</p>}
        </div>
      </div>
      {children && <div className="flex items-center gap-2 flex-wrap">{children}</div>}
    </div>
  );
}

export default function App() {
  const [showInstructions, setShowInstructions] = useState(false);
  const [files, setFiles] = useState([]); // { fileId, fileName, empresa, sheets } | { ..., readError }
  const [unifyNames, setUnifyNames] = useState(true);
  const [consolidateNomina, setConsolidateNomina] = useState(false);
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [dragging, setDragging] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const rowsPerPage = 10;

  // Para bajar solos hasta el resultado cuando termina de procesar.
  const fileInputRef = useRef(null);
  const fileListRef = useRef(null);
  const resultsRef = useRef(null);
  const scrollPending = useRef(false);

  // Los archivos se leen una vez; la conversión se recalcula si cambia la opción de nombres.
  const processed = useMemo(
    () => files.map((f) => convertLoadedFile(f, unifyNames)),
    [files, unifyNames]
  );

  // Consolidación de varios archivos de una misma empresa bajo la misma casilla:
  // varias nóminas anchas (p. ej. los 4 a 6 archivos mensuales de RemoFirst),
  // varios Movimiento CC (p. ej. un archivo por rango de fechas) y/o varias
  // facturaciones EOR (p. ej. una por mes). Es opcional porque los códigos de
  // empleado / nombres de empresas distintas podrían repetirse y mezclar datos.
  // Para el Movimiento CC, además, consolidar antes de convertir es lo que
  // permite que la inferencia de empleado por fecha (en aportes patronales y en
  // filas de salario sin Tercero) vea todas las filas de una misma fecha aunque
  // hayan llegado en archivos distintos.
  const nominaConsolidation = useMemo(() => {
    if (!consolidateNomina) return null;
    const inputs = processed
      .filter((f) => f.sourceType === 'nomina' && f.nominaRows)
      .map((f) => ({ rows: f.nominaRows, name: f.fileName }));
    if (inputs.length < 2) return null;
    return convertNominaFiles(inputs, { unifyNamesByCode: unifyNames });
  }, [processed, consolidateNomina, unifyNames]);

  const movimientoConsolidation = useMemo(() => {
    if (!consolidateNomina) return null;
    const inputs = processed
      .filter((f) => f.sourceType === 'movimiento' && f.movimientoRows)
      .map((f) => ({ rows: f.movimientoRows, name: f.fileName }));
    if (inputs.length < 2) return null;
    return convertMovimientoFiles(inputs);
  }, [processed, consolidateNomina]);

  const facturacionConsolidation = useMemo(() => {
    if (!consolidateNomina) return null;
    const inputs = processed
      .filter((f) => f.sourceType === 'facturacion' && f.facturacionRows)
      .map((f) => ({ rows: f.facturacionRows, name: f.fileName }));
    if (inputs.length < 2) return null;
    return convertFacturacionFiles(inputs);
  }, [processed, consolidateNomina]);

  const consolidatedRows = useMemo(() => {
    const nominaRows = nominaConsolidation
      ? nominaConsolidation.records
      : processed.filter((f) => f.sourceType === 'nomina').flatMap((f) => f.rows);
    const movimientoRows = movimientoConsolidation
      ? movimientoConsolidation.records
      : processed.filter((f) => f.sourceType === 'movimiento').flatMap((f) => f.rows);
    const facturacionRows = facturacionConsolidation
      ? facturacionConsolidation.records
      : processed.filter((f) => f.sourceType === 'facturacion').flatMap((f) => f.rows);
    const otherRows = processed
      .filter((f) => !KNOWN_SOURCE_TYPES.includes(f.sourceType))
      .flatMap((f) => f.rows);
    return [...nominaRows, ...movimientoRows, ...facturacionRows, ...otherRows];
  }, [processed, nominaConsolidation, movimientoConsolidation, facturacionConsolidation]);

  // Colores presentes en el resultado, para la leyenda.
  const legendColors = useMemo(() => {
    const used = new Set();
    for (const row of consolidatedRows) if (row._fill) used.add(row._fill);
    return Object.keys(STATUS_COLORS).filter((hex) => used.has(hex));
  }, [consolidatedRows]);

  useEffect(() => {
    if (loading || !scrollPending.current) return;
    const target = resultsRef.current || fileListRef.current;
    if (target) {
      scrollPending.current = false;
      target.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }, [files, loading]);

  // Lee y agrega archivos (viene del selector, del botón "Nuevo archivo" o de arrastrar y soltar).
  const processUploadedFiles = async (uploaded) => {
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
    }
  };

  const handleFileUpload = async (e) => {
    const uploaded = Array.from(e.target.files || []);
    await processUploadedFiles(uploaded);
    e.target.value = '';
  };

  const handleDrop = async (e) => {
    e.preventDefault();
    setDragging(false);
    const dropped = Array.from(e.dataTransfer.files || []).filter((f) =>
      /\.(xlsx|xlsm)$/i.test(f.name)
    );
    await processUploadedFiles(dropped);
  };

  const openFilePicker = () => fileInputRef.current?.click();

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
      OUTPUT_COLUMNS.some((col) => formatCellValue(row[col]).toLowerCase().includes(term))
    );
  }, [consolidatedRows, searchTerm]);

  const totalPages = Math.ceil(filteredData.length / rowsPerPage) || 1;
  const paginatedData = useMemo(() => {
    const start = (currentPage - 1) * rowsPerPage;
    return filteredData.slice(start, start + rowsPerPage);
  }, [filteredData, currentPage]);

  const downloadXLSX = async () => {
    if (filteredData.length === 0) return;
    setExporting(true);
    try {
      const blob = await buildXlsxBlobWithJSZip(filteredData, OUTPUT_COLUMNS);
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
    const headerLine = OUTPUT_COLUMNS.join(',');
    // Los montos van como número plano (sin puntos de miles) para poder sumar y filtrar en Excel.
    const lines = filteredData.map((row) =>
      OUTPUT_COLUMNS.map((col) =>
        VALUE_COLUMNS.has(col)
          ? String(row[col])
          : `"${formatCellValue(row[col]).replace(/"/g, '""')}"`
      ).join(',')
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
      OUTPUT_COLUMNS.forEach((col) => {
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

  const totalWarnings = processed.filter((f) => f.warning).length;
  const hasFiles = files.length > 0;
  const hasResults = consolidatedRows.length > 0;

  // Los ceros se muestran como "-" (igual que el formato contable del Excel) y los
  // montos llevan el color que traía la nómina, en las dos columnas de valor.
  const renderCell = (row, col) => {
    const value = row[col];
    if (VALUE_COLUMNS.has(col) && value === 0) return '-';
    return formatCellValue(value);
  };
  const cellClass = (col) => {
    const base = 'px-5 py-2.5 whitespace-nowrap';
    return VALUE_COLUMNS.has(col) ? `${base} text-right tabular-nums` : base;
  };
  const cellStyle = (row, col) => {
    const style = { minWidth: COLUMN_MIN_WIDTH[col] };
    if (VALUE_COLUMNS.has(col) && row._fill) style.backgroundColor = `#${row._fill}`;
    return style;
  };

  // Panel de avisos de una consolidación (nómina, Movimiento CC o facturación).
  const renderConsolidationPanel = (title, consolidation) =>
    consolidation && (
      <div className="px-4 py-3 rounded-xl border border-blue-200 bg-blue-50/60 text-xs space-y-1">
        <p className="font-semibold text-blue-900">{title}</p>
        <ul className="pl-4 space-y-1 list-disc">
          {consolidation.notes.map((n, i) => (
            <li key={i} className={n.type === 'warn' ? 'text-amber-800' : 'text-slate-600'}>
              {n.text}
            </li>
          ))}
        </ul>
      </div>
    );

  // Progreso de los tres pasos.
  const steps = [
    { label: 'Cargar archivos', done: hasFiles },
    { label: 'Revisar avisos', done: hasResults },
    { label: 'Exportar', done: false }
  ];
  const activeStep = !hasFiles ? 0 : !hasResults ? 1 : 2;

  return (
    <div className="min-h-screen bg-slate-100 text-slate-800 font-sans">
      {/* Selector de archivos compartido por la zona de carga y el botón "Nuevo archivo" */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xlsm"
        multiple
        onChange={handleFileUpload}
        className="hidden"
      />

      {/* Header */}
      <header className="bg-white border-b border-slate-200 px-6 sm:px-8 py-3 flex items-center justify-between">
        <img src="/logo.jpeg" alt="Logo" className="h-12 w-auto object-contain" />
        <div className="flex items-center gap-2.5 border border-slate-200 bg-slate-50 px-4 py-2 rounded-full text-sm font-semibold text-slate-700">
          <User className="w-4 h-4 text-blue-600" />
          <span>Bienvenido Usuario</span>
        </div>
      </header>

      {/* Banda de título */}
      <section className="bg-blue-950 text-white">
        <div className="max-w-6xl mx-auto px-6 pt-10 pb-24 flex flex-wrap items-end justify-between gap-8">
          <div className="max-w-2xl space-y-3">
            <h1 className="text-3xl sm:text-4xl font-extrabold tracking-tight leading-tight">
              Convertidor a formato largo
            </h1>
            <p className="text-sm text-blue-100/80 leading-relaxed">
              Carga la nómina, la facturación EOR o el Movimiento CC de Siigo. La app reconoce
              sola cada formato y te devuelve una tabla lista para filtrar y cruzar la cuenta 28.
            </p>
          </div>

          {/* Pasos */}
          <ol className="flex items-center gap-2 text-xs font-medium">
            {steps.map((s, i) => (
              <li key={s.label} className="flex items-center gap-2">
                <span
                  className={`flex items-center gap-2 px-3 py-1.5 rounded-full ring-1 transition-colors ${
                    i === activeStep
                      ? 'bg-white text-blue-950 ring-white'
                      : s.done
                      ? 'bg-blue-900 text-blue-100 ring-blue-700'
                      : 'bg-transparent text-blue-200/70 ring-blue-800'
                  }`}
                >
                  {s.done && i !== activeStep ? (
                    <CheckCircle2 className="w-3.5 h-3.5" />
                  ) : (
                    <span className="font-bold">{i + 1}</span>
                  )}
                  {s.label}
                </span>
                {i < steps.length - 1 && <span className="w-4 h-px bg-blue-800" />}
              </li>
            ))}
          </ol>
        </div>
      </section>

      <main className="max-w-6xl mx-auto px-6 -mt-16 pb-16 space-y-6">
        {/* Instrucciones */}
        <div className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
          <button
            onClick={() => setShowInstructions(!showInstructions)}
            className="w-full px-6 py-3.5 flex items-center justify-between text-left hover:bg-slate-50 transition-colors cursor-pointer"
          >
            <div className="flex items-center gap-3 text-slate-900 font-semibold text-sm">
              <Info className="w-4 h-4 text-blue-600 shrink-0" />
              <span>¿Cómo usar esta herramienta?</span>
            </div>
            {showInstructions ? (
              <ChevronUp className="w-4 h-4 text-slate-500" />
            ) : (
              <ChevronDown className="w-4 h-4 text-slate-500" />
            )}
          </button>
          {showInstructions && (
            <div className="px-6 pb-6 pt-4 border-t border-slate-100 text-xs text-slate-700 space-y-3 leading-relaxed">
              <p>
                <strong className="text-slate-900">1. Cargar el archivo:</strong> puede ser la
                facturación EOR (hoja INVOICING, con EMPLOYEE CODE, NAME y Payroll Month), la
                nómina que se envía a facturación (un bloque por mes, encabezado con EMPLOYEE CODE
                y NAME) o el Movimiento CC de Siigo (encabezado con Comprobante, Fecha elaboración,
                Descripción, Débito y Crédito). La app prueba primero como facturación, luego como
                nómina y, si no encuentra bloques, como Movimiento CC. Puedes seleccionar varios
                archivos a la vez, incluso mezclando los tipos. Si son varios archivos de una misma
                empresa — varias nóminas (p. ej. los de RemoFirst), varias facturaciones (una por
                mes) o varios Movimiento CC (p. ej. uno por rango de fechas) — marca la casilla de
                consolidar para que se junten antes de convertir.
              </p>
              <p>
                <strong className="text-slate-900">2. Qué sale de la nómina:</strong> una fila por
                mes, empleado y concepto, con los nombres de concepto tal como aparecen en el
                encabezado de la nómina, y al final de cada empleado su TOTAL EMPLOYEE COST. Los
                conceptos en cero no se listan.
              </p>
              <p>
                <strong className="text-slate-900">3. Qué sale de la facturación EOR:</strong> lo
                mismo que de la nómina (una fila por mes, empleado y concepto, más el TOTAL
                EMPLOYEE COST), leyendo todas las columnas de costo hasta TOTAL EMPLOYEE COST
                (salario, seguridad social, prestaciones y otros costos del empleador). No se
                incluyen FEE, BANKING TAX, IVA ni las columnas en USD. El mes sale de la columna
                Payroll Month y el año del nombre del archivo (p. ej. "... Agosto 2026.xlsx"); las
                filas de "Monthly Adjustment" de meses anteriores quedan en su propio mes.
              </p>
              <p>
                <strong className="text-slate-900">4. Qué sale del Movimiento CC:</strong> cada
                Descripción se clasifica a un concepto tipo nómina (SALARY, PENSION COST, HEALTH
                COST…) con la tabla CONCEPT_KEYWORDS, se suma por mes, concepto y empleado (Débito −
                Crédito) y se agrega TOTAL EMPLOYEE COST por mes y empleado. Los aportes patronales,
                y las filas de salario/prestación que vengan sin Tercero, se asignan al empleado del
                mismo día (si hay un único candidato); si no se puede, salen como "(sin asignar)".
                Lo que no se reconoce queda con el texto de Descripción y se avisa.
              </p>
              <p>
                <strong className="text-slate-900">5. Colores:</strong> son los que ya trae el
                archivo (el resultado del cruce con Siigo): verde y azul = cruce ok, amarillo = no
                está en el otro lado, rojo = diferencias, verde limón = cruza entre meses, morado =
                débito y crédito se anulan. Si el archivo aún no está pintado (o es una
                facturación), las filas salen sin color.
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

        {/* Paso 1: carga */}
        <section className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
          <SectionHeader
            step="1"
            title="Cargar archivos"
            subtitle="Formatos .xlsx y .xlsm. Puedes mezclar nóminas, facturación y Movimiento CC."
          />
          <div className="p-6">
            <div
              role="button"
              tabIndex={0}
              onClick={openFilePicker}
              onKeyDown={(e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                  e.preventDefault();
                  openFilePicker();
                }
              }}
              onDragOver={(e) => {
                e.preventDefault();
                setDragging(true);
              }}
              onDragLeave={() => setDragging(false)}
              onDrop={handleDrop}
              className={`border-2 border-dashed rounded-xl text-center cursor-pointer transition-colors focus:outline-none focus-visible:ring-2 focus-visible:ring-blue-500 ${
                hasFiles ? 'p-5' : 'p-10'
              } ${
                dragging
                  ? 'border-blue-500 bg-blue-50'
                  : 'border-slate-300 bg-slate-50 hover:border-blue-400 hover:bg-blue-50/40'
              }`}
            >
              <div
                className={`flex items-center justify-center gap-4 ${
                  hasFiles ? 'flex-row' : 'flex-col'
                }`}
              >
                <div className="p-3 bg-blue-100 text-blue-700 rounded-full">
                  <Upload className={hasFiles ? 'w-5 h-5' : 'w-6 h-6'} />
                </div>
                <div className={hasFiles ? 'text-left' : ''}>
                  <p className="font-semibold text-sm text-slate-800">
                    {hasFiles
                      ? 'Arrastra más archivos aquí o haz clic para agregarlos'
                      : 'Arrastra tus archivos aquí o haz clic para buscarlos'}
                  </p>
                  <p className="text-xs text-slate-500 mt-1">
                    Se suman al consolidado que ya tienes cargado.
                  </p>
                </div>
              </div>
            </div>

            {loading && (
              <div className="flex items-center justify-center gap-2 text-blue-700 pt-5">
                <div className="w-4 h-4 border-2 border-blue-600 border-t-transparent rounded-full animate-spin" />
                <span className="text-xs font-medium">Leyendo y transformando archivos...</span>
              </div>
            )}
          </div>
        </section>

        {/* Paso 2: archivos cargados */}
        {hasFiles && !loading && (
          <section
            ref={fileListRef}
            className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden scroll-mt-6"
          >
            <SectionHeader
              step="2"
              title={`Archivos cargados (${files.length})`}
              subtitle="Revisa que cada archivo se haya reconocido bien."
            >
              <button
                onClick={openFilePicker}
                className="flex items-center gap-1.5 px-3.5 py-2 bg-blue-900 hover:bg-blue-800 text-white font-semibold text-xs rounded-lg transition-colors cursor-pointer"
              >
                <Plus className="w-4 h-4" />
                Nuevo archivo
              </button>
              <button
                onClick={clearAll}
                className="flex items-center gap-1.5 px-3.5 py-2 bg-white border border-red-200 text-red-600 hover:bg-red-50 font-semibold text-xs rounded-lg transition-colors cursor-pointer"
              >
                <Trash2 className="w-4 h-4" />
                Quitar todos
              </button>
            </SectionHeader>

            <div className="p-6 space-y-4">
              {/* Opciones */}
              <div className="flex flex-wrap gap-x-6 gap-y-2 px-4 py-3 rounded-xl bg-slate-50 border border-slate-200">
                <label className="flex items-center gap-2 text-xs text-slate-700 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={unifyNames}
                    onChange={(e) => setUnifyNames(e.target.checked)}
                    className="accent-blue-900"
                  />
                  Unificar el nombre de cada empleado por código (solo nómina)
                </label>
                <label className="flex items-center gap-2 text-xs text-slate-700 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={consolidateNomina}
                    onChange={(e) => setConsolidateNomina(e.target.checked)}
                    className="accent-blue-900"
                  />
                  Varios archivos de una misma empresa: consolidar
                </label>
              </div>

              {/* Lista */}
              <div className="space-y-2.5">
                {processed.map((f) => (
                  <div
                    key={f.fileId}
                    className={`rounded-xl border px-4 py-3 text-xs space-y-2 ${
                      f.warning ? 'bg-amber-50/70 border-amber-200' : 'bg-white border-slate-200'
                    }`}
                  >
                    <div className="flex items-center justify-between gap-3">
                      <div className="flex items-center gap-3 min-w-0">
                        <div
                          className={`p-2 rounded-lg shrink-0 ${
                            f.warning ? 'bg-amber-100 text-amber-700' : 'bg-emerald-50 text-emerald-600'
                          }`}
                        >
                          {f.warning ? (
                            <AlertTriangle className="w-4 h-4" />
                          ) : (
                            <FileText className="w-4 h-4" />
                          )}
                        </div>
                        <div className="min-w-0">
                          <div className="flex items-center gap-2 min-w-0">
                            <p className="font-semibold text-slate-800 truncate text-sm">
                              {f.fileName}
                            </p>
                            {f.sourceType && (
                              <span
                                className={`shrink-0 px-2 py-0.5 rounded-full text-[11px] font-semibold ring-1 ${SOURCE_TYPE_BADGE[f.sourceType]}`}
                              >
                                {SOURCE_TYPE_LABEL[f.sourceType]}
                              </span>
                            )}
                          </div>
                          <p className={f.warning ? 'text-amber-800 mt-0.5' : 'text-slate-500 mt-0.5'}>
                            {f.warning ? f.warning : `${f.rows.length} filas generadas`}
                          </p>
                        </div>
                      </div>
                      <button
                        onClick={() => removeFile(f.fileId)}
                        className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-200 bg-white text-slate-600 hover:text-red-600 hover:border-red-200 hover:bg-red-50 font-medium transition-colors cursor-pointer shrink-0"
                        title="Quitar este archivo"
                      >
                        <X className="w-3.5 h-3.5" />
                        Quitar archivo
                      </button>
                    </div>
                    {f.notes && f.notes.length > 0 && (
                      <ul className="pl-12 space-y-1 list-disc marker:text-slate-300">
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
              </div>

              {renderConsolidationPanel('Consolidación de nóminas', nominaConsolidation)}
              {renderConsolidationPanel('Consolidación de Movimiento CC', movimientoConsolidation)}
              {renderConsolidationPanel('Consolidación de Facturación', facturacionConsolidation)}
              {totalWarnings > 0 && (
                <p className="text-xs text-amber-700">
                  {totalWarnings} archivo(s) con avisos — revisa que sean la nómina, la
                  facturación o el Movimiento CC en el formato esperado.
                </p>
              )}
            </div>
          </section>
        )}

        {/* Paso 3: resultado */}
        {hasResults && !loading && (
          <section
            ref={resultsRef}
            className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden scroll-mt-6"
          >
            <SectionHeader
              step="3"
              title="Resultado consolidado"
              subtitle={`${filteredData.length} registros${
                searchTerm.trim() ? ' que coinciden con la búsqueda' : ''
              }`}
            >
              <button
                onClick={downloadXLSX}
                disabled={exporting}
                className="flex items-center gap-1.5 px-3.5 py-2 bg-blue-900 hover:bg-blue-800 disabled:opacity-50 text-white font-semibold text-xs rounded-lg transition-colors cursor-pointer"
              >
                <FileSpreadsheet className="w-4 h-4" />
                {exporting ? 'Generando...' : 'Descargar Excel'}
              </button>
              <button
                onClick={downloadCSV}
                className="flex items-center gap-1.5 px-3.5 py-2 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 font-semibold text-xs rounded-lg transition-colors cursor-pointer"
              >
                <Download className="w-4 h-4" />
                CSV
              </button>
              <button
                onClick={downloadJSON}
                className="flex items-center gap-1.5 px-3.5 py-2 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 font-semibold text-xs rounded-lg transition-colors cursor-pointer"
              >
                <Download className="w-4 h-4" />
                JSON
              </button>
            </SectionHeader>

            {/* Búsqueda y leyenda */}
            <div className="px-6 py-3 bg-slate-50 border-b border-slate-100 flex flex-wrap items-center justify-between gap-4">
              <div className="relative w-full max-w-sm">
                <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                <input
                  type="text"
                  placeholder="Buscar empleado, concepto, mes..."
                  value={searchTerm}
                  onChange={(e) => {
                    setSearchTerm(e.target.value);
                    setCurrentPage(1);
                  }}
                  className="w-full pl-9 pr-3 py-2 text-xs bg-white border border-slate-200 rounded-lg focus:outline-none focus:border-blue-500 focus:ring-2 focus:ring-blue-100"
                />
              </div>
              {legendColors.length > 0 && (
                <div className="flex flex-wrap items-center gap-x-4 gap-y-1 text-xs text-slate-600">
                  {legendColors.map((hex) => (
                    <span key={hex} className="flex items-center gap-1.5">
                      <span
                        className="inline-block w-3 h-3 rounded-sm ring-1 ring-black/10"
                        style={{ backgroundColor: `#${hex}` }}
                      />
                      {STATUS_COLORS[hex]}
                    </span>
                  ))}
                </div>
              )}
            </div>

            <div className="overflow-x-auto max-h-[28rem]">
              <table className="w-full text-left text-xs text-slate-700">
                <thead className="bg-slate-100 text-slate-600 sticky top-0 border-b border-slate-200 font-semibold">
                  <tr>
                    {OUTPUT_COLUMNS.map((col) => (
                      <th
                        key={col}
                        className={`px-5 py-3 whitespace-nowrap ${
                          VALUE_COLUMNS.has(col) ? 'text-right' : ''
                        }`}
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
                      {OUTPUT_COLUMNS.map((col) => (
                        <td key={col} className={cellClass(col)} style={cellStyle(row, col)}>
                          {renderCell(row, col)}
                        </td>
                      ))}
                    </tr>
                  ))}
                  {paginatedData.length === 0 && (
                    <tr>
                      <td
                        colSpan={OUTPUT_COLUMNS.length}
                        className="px-5 py-10 text-center text-slate-500"
                      >
                        Ningún registro coincide con "{searchTerm}". Prueba con otro término.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            <div className="px-6 py-3 bg-slate-50 border-t border-slate-100 flex items-center justify-between text-xs text-slate-600">
              <span>
                Página {currentPage} de {totalPages}
              </span>
              <div className="flex items-center gap-1">
                <button
                  disabled={currentPage === 1}
                  onClick={() => setCurrentPage((prev) => Math.max(prev - 1, 1))}
                  className="p-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                  aria-label="Página anterior"
                >
                  <ChevronLeft className="w-4 h-4" />
                </button>
                <button
                  disabled={currentPage === totalPages}
                  onClick={() => setCurrentPage((prev) => Math.min(prev + 1, totalPages))}
                  className="p-1.5 rounded-lg border border-slate-200 bg-white hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                  aria-label="Página siguiente"
                >
                  <ChevronRight className="w-4 h-4" />
                </button>
              </div>
            </div>
          </section>
        )}

        {hasFiles && !hasResults && !loading && (
          <div className="bg-white border border-slate-200 rounded-2xl shadow-sm p-10 text-center space-y-3">
            <FileX2 className="w-8 h-8 text-slate-300 mx-auto" />
            <p className="text-sm font-semibold text-slate-700">Ningún archivo generó registros</p>
            <p className="text-xs text-slate-500 max-w-md mx-auto">
              Revisa los avisos de arriba: probablemente el archivo no es la nómina, la facturación
              ni el Movimiento CC en el formato que la app reconoce.
            </p>
            <button
              onClick={openFilePicker}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-blue-900 hover:bg-blue-800 text-white font-semibold text-xs rounded-lg transition-colors cursor-pointer"
            >
              <Plus className="w-4 h-4" />
              Nuevo archivo
            </button>
          </div>
        )}
      </main>
    </div>
  );
}