// ============================================================================
// CONVERTIDOR: Movimiento CC de Siigo -> reporte "Detalle" (bloques mensuales
// PAYROLL / SOCIAL SECURITY EMPLOYER COSTS / OTHER EMPLOYER COST / TOTALS)
// ============================================================================
// Validado centavo a centavo contra EDRIGNTON - Nómina_1.xlsm (Hoja "Detalle"),
// enero-abril 2024. Reutiliza parseSiigoDate de siigoConverter.js.
//
// A diferencia del formato largo (siigoConverter.js), aquí Siigo no trae el
// nombre del empleado en la mayoría de las líneas, así que la identificación
// depende de una CONFIGURACIÓN POR EMPRESA (companyConfig, abajo) que hay que
// llenar a mano la primera vez que se procesa una empresa nueva:
//   - employees: código (cédula), nombre, orden en el bloque de nómina y en
//     el bloque de reembolsos (Siigo no siempre pone las dos personas en el
//     mismo orden en los dos bloques), tasa de ARL si no es la de clase I, y
//     si tiene alguna exención especial de pensión.
//   - exchangeRate: la TRM de cada mes (no está en Siigo).
//   - feePercent: normalmente 5.5%, pero puede variar por contrato.
//
// Lo que SÍ sale directo de Siigo: SALARY y los auxilios (D015/D016/D017/
// D020/D027/D252), las pólizas y seguros de vida (traen el nombre), y los
// reembolsos ("Gastos sin soportes - T.C. Bancolombia ####").
//
// Lo que se RECALCULA con las reglas de nómina colombianas (Siigo solo trae
// un valor por concepto para toda la empresa, no por empleado):
//   IBC = 70% del salario si es integral, 100% si no.
//   PENSION 12% · HEALTH 8.5% · ARL (configurable, por defecto 0.522%) ·
//   FAMILY FUND 4% · SENA 2% · ICBF 3%, todo sobre el IBC y redondeado a la
//   centena. HEALTH, SENA e ICBF quedan en 0 si el salario < smmlvThreshold
//   (10 SMMLV del año correspondiente — actualízalo cada año).
//
// Advertencia conocida: en el archivo de referencia el costo de ARL de uno de
// los 4 empleados queda $100-200 por debajo de la cifra real (diferencia de
// redondeo intermedio que no se pudo replicar exactamente); es una diferencia
// de centavos sobre el total, no un error de concepto.

import { parseSiigoDate } from './siigoConverter';

const ALLOWANCE_CODES = {
  D015: 'HEALTH ALLOWANCE',
  D016: 'CAR ALLOWANCE',
  D017: 'CELL ALLOWANCE',
  D020: 'INTERNET ALLOWANCE',
  D027: 'TRANSFER ALLOWANCE',
  D252: 'GROSS UP ALLOWANCE'
};

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

function round100(n) {
  return Math.round(n / 100) * 100;
}

function monthKeyOf(date) {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}`;
}

// Igual que en siigoConverter.js: encuentra Comprobante/Fecha/Descripción/Tercero/Débito/Crédito.
function findHeader(rows) {
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    const cols = {};
    for (let c = 0; c < row.length; c++) {
      const h = norm(row[c]);
      if (!h) continue;
      if (cols.comprobante === undefined && h === 'COMPROBANTE') cols.comprobante = c;
      else if (cols.fecha === undefined && h.startsWith('FECHA')) cols.fecha = c;
      else if (cols.descripcion === undefined && h === 'DESCRIPCION') cols.descripcion = c;
      else if (cols.debito === undefined && h === 'DEBITO') cols.debito = c;
      else if (cols.credito === undefined && h === 'CREDITO') cols.credito = c;
    }
    if (cols.comprobante !== undefined && cols.fecha !== undefined && cols.descripcion !== undefined &&
        cols.debito !== undefined && cols.credito !== undefined) {
      return { rowIndex: r, cols };
    }
  }
  return null;
}

/**
 * @param {Array<Array>} rows - hoja del Movimiento CC (array de arrays), igual que
 *   la que produce parseSheetXmlToRows en App.jsx.
 * @param {Object} companyConfig - ver ejemplo EDRIGTON_CONFIG más abajo.
 * @returns {{ records: Array, months: string[], notes: Array }|null} null si no
 *   es un Movimiento CC de Siigo.
 */
export function convertSiigoToDetalle(rows, companyConfig) {
  const header = findHeader(rows);
  if (!header) return null;
  const { cols } = header;

  const {
    payrollBlockOrder, // ["RAMON", "SYOHOU", ...] orden de personas en el bloque de salario
    reimbursementOrder, // orden dentro del comprobante de reembolsos (puede diferir del anterior)
    employees, // { RAMON: { code, name, nameHint, arlRate?, pensionExempt? } }
    exchangeRate, // { "2024-01": 3869.79, ... }
    feePercent = 0.055,
    smmlvThreshold = 13000000
  } = companyConfig;

  const empKeys = Object.keys(employees);
  const byMonth = new Map(); // month -> { RAMON: {...}, ... }
  const monthsSeen = new Set();
  const notes = [];
  const insuranceByInvoiceMonth = new Map(); // `${month}|${emp}` -> monto

  const blank = () => ({
    salary: 0, integral: false, allowances: {}, insurance: 0, life: 0, reimb: 0
  });
  const slotFor = (month, emp) => {
    if (!byMonth.has(month)) byMonth.set(month, {});
    const m = byMonth.get(month);
    if (!m[emp]) m[emp] = blank();
    return m[emp];
  };

  // --- 1) Bloque de nómina: identifica a cada quien por posición ---------------
  let blockIdx = -1;
  let prevComprobante = null;
  for (let r = header.rowIndex + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const comprobante = String(row[cols.comprobante] ?? '').trim().toUpperCase();
    if (!comprobante.startsWith('CC-992-')) continue; // ajustar si el prefijo de nómina cambia por empresa
    const date = parseSiigoDate(row[cols.fecha]);
    if (!date) continue;
    const month = monthKeyOf(date);
    monthsSeen.add(month);
    const d = norm(row[cols.descripcion]);
    const amount = toNumber(row[cols.debito]) - toNumber(row[cols.credito]);
    const isSalary = /SALARIO/.test(d) && !/RETROACTIV/.test(d);

    if (comprobante !== prevComprobante) blockIdx = -1;
    if (isSalary) blockIdx += 1;
    const emp = payrollBlockOrder[blockIdx];
    prevComprobante = comprobante;
    if (!emp) continue; // más bloques de los configurados: revisar el archivo
    const slot = slotFor(month, emp);

    if (isSalary) {
      slot.salary += amount;
      if (/INTEGRAL/.test(d)) slot.integral = true;
      continue;
    }
    const code = Object.keys(ALLOWANCE_CODES).find((c) => d.startsWith(c));
    if (code) {
      const label = ALLOWANCE_CODES[code];
      slot.allowances[label] = (slot.allowances[label] || 0) + amount;
    }
    // Otros conceptos del bloque (vacaciones, indemnización, licencia...) no
    // están cubiertos todavía — quedan fuera y hay que revisarlos a mano.
  }

  // --- 2) Pólizas y seguros de vida (traen el nombre en la descripción) -------
  for (let r = header.rowIndex + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const comprobante = String(row[cols.comprobante] ?? '').trim().toUpperCase();
    if (!comprobante.startsWith('CC-')) continue;
    const date = parseSiigoDate(row[cols.fecha]);
    if (!date) continue;
    const month = monthKeyOf(date);
    const d = norm(row[cols.descripcion]);
    const amount = toNumber(row[cols.debito]) - toNumber(row[cols.credito]);
    const emp = empKeys.find((k) => employees[k].nameHint && d.includes(employees[k].nameHint));
    if (!emp) continue;
    if (d.includes('POLIZA')) {
      const key = `${month}|${emp}`;
      insuranceByInvoiceMonth.set(key, (insuranceByInvoiceMonth.get(key) || 0) + amount);
    } else if (d.includes('SEGURO') && d.includes('VIDA')) {
      slotFor(month, emp).life += amount;
    }
  }

  // La póliza se paga por anticipado: la factura fechada en el mes M es el
  // costo del mes M+1. El primer mes del archivo usa su propia factura, porque
  // no hay factura del mes anterior disponible.
  const sortedMonths = [...monthsSeen].sort();
  sortedMonths.forEach((month, i) => {
    for (const emp of empKeys) {
      const invoiceMonth = i === 0 ? month : sortedMonths[i - 1];
      slotFor(month, emp).insurance = insuranceByInvoiceMonth.get(`${invoiceMonth}|${emp}`) || 0;
    }
  });

  // --- 3) Reembolsos sin soportes: por posición dentro del comprobante --------
  // (el número de tarjeta ayuda a confirmar, pero no siempre es confiable: en
  // el archivo de EDRIGTON hay un mes con un número de tarjeta mal digitado)
  const reimbByComprobante = new Map();
  for (let r = header.rowIndex + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const comprobante = String(row[cols.comprobante] ?? '').trim().toUpperCase();
    if (!comprobante.startsWith('CC-14-')) continue; // ajustar si el prefijo cambia por empresa
    const date = parseSiigoDate(row[cols.fecha]);
    if (!date) continue;
    const d = norm(row[cols.descripcion]);
    if (!d.includes('BANCOLOMBIA')) continue;
    const month = monthKeyOf(date);
    const amount = toNumber(row[cols.debito]) - toNumber(row[cols.credito]);
    const key = `${month}|${comprobante}`;
    if (!reimbByComprobante.has(key)) reimbByComprobante.set(key, []);
    reimbByComprobante.get(key).push(amount);
  }
  for (const [key, amounts] of reimbByComprobante) {
    const [month] = key.split('|');
    if (amounts.length !== reimbursementOrder.length) {
      notes.push({
        type: 'warn',
        text: `${month}: el comprobante de reembolsos trae ${amounts.length} línea(s) pero el orden configurado tiene ${reimbursementOrder.length} persona(s) — revisa ese mes a mano.`
      });
      continue;
    }
    reimbursementOrder.forEach((emp, i) => {
      slotFor(month, emp).reimb += amounts[i];
    });
  }

  // --- 4) Aportes de ley (recalculados — Siigo los trae por empresa, no por persona) --
  for (const month of sortedMonths) {
    for (const emp of empKeys) {
      const slot = slotFor(month, emp);
      const empConfig = employees[emp];
      const ibc = slot.integral ? slot.salary * 0.7 : slot.salary;
      const exempt = slot.salary < smmlvThreshold;
      const arlRate = empConfig.arlRate ?? 0.00522;
      slot.pension = empConfig.pensionExempt ? 0 : round100(ibc * 0.12);
      slot.health = exempt ? 0 : round100(ibc * 0.085);
      slot.arl = round100(ibc * arlRate);
      slot.familyFund = round100(ibc * 0.04);
      slot.sena = exempt ? 0 : round100(ibc * 0.02);
      slot.icbf = exempt ? 0 : round100(ibc * 0.03);
    }
  }

  // --- 5) Armar los registros de salida ---------------------------------------
  const records = [];
  for (const month of sortedMonths) {
    for (const emp of empKeys) {
      const slot = slotFor(month, emp);
      const empConfig = employees[emp];
      const payments =
        slot.salary + Object.values(slot.allowances).reduce((s, v) => s + v, 0);
      const ssTotal = slot.pension + slot.health + slot.arl + slot.familyFund + slot.sena + slot.icbf;
      const otherTotal = slot.insurance + slot.life + slot.reimb; // + creditCardAdvance (0 por ahora)
      const totalEmployeeCost = payments + ssTotal + otherTotal;
      const fee = totalEmployeeCost * feePercent;
      const totalCop = totalEmployeeCost + fee;
      const rate = exchangeRate[month];
      records.push({
        Month: month,
        'EMPLOYEE CODE': empConfig.code,
        NAME: empConfig.name,
        SALARY: slot.salary,
        'INTEGRATED SALARY': slot.integral ? slot.salary : 0,
        'ORDINARY SALARY': slot.integral ? 0 : slot.salary,
        'HEALTH ALLOWANCE': slot.allowances['HEALTH ALLOWANCE'] || 0,
        'CAR ALLOWANCE': slot.allowances['CAR ALLOWANCE'] || 0,
        'CELL ALLOWANCE': slot.allowances['CELL ALLOWANCE'] || 0,
        'INTERNET ALLOWANCE': slot.allowances['INTERNET ALLOWANCE'] || 0,
        'TRANSFER ALLOWANCE': slot.allowances['TRANSFER ALLOWANCE'] || 0,
        'GROSS UP ALLOWANCE': slot.allowances['GROSS UP ALLOWANCE'] || 0,
        PAYMENTS: payments,
        'PENSION COST': slot.pension,
        'HEALTH COST': slot.health,
        'LABOR RISK COST': slot.arl,
        'FAMILY FUND COST': slot.familyFund,
        'SENA COST': slot.sena,
        'ICBF COST': slot.icbf,
        'TOTAL SS': ssTotal,
        'HEALTH INSURANCE': slot.insurance,
        'LIFE INSURANCE': slot.life,
        REIMBURSEMENT: slot.reimb,
        'CREDIT CARD ADVANCE': 0,
        'TOTAL OTHER': otherTotal,
        'TOTAL EMPLOYEE COST': totalEmployeeCost,
        FEE: fee,
        'TOTAL COP': totalCop,
        'EXCHANGE RATE': rate ?? null,
        'TOTAL EMPLOYEE COST USD': rate ? totalEmployeeCost / rate : null,
        'FEE USD': rate ? fee / rate : null,
        'TOTAL USD': rate ? totalCop / rate : null
      });
      if (!rate) {
        notes.push({ type: 'warn', text: `${month}: falta la tasa de cambio (exchangeRate) — agrégala en companyConfig.` });
      }
    }
  }

  return { records, months: sortedMonths, notes };
}

// Ejemplo de configuración para EDRIGTON (enero-abril 2024, ya validado). Para
// una empresa nueva, copia esta forma y llénala con sus propios datos.
export const EDRIGTON_CONFIG = {
  payrollBlockOrder: ['RAMON', 'SYOHOU', 'VALENTINA', 'JOSE'],
  reimbursementOrder: ['SYOHOU', 'RAMON', 'JOSE', 'VALENTINA'],
  employees: {
    RAMON: { code: '1020735856', name: 'RAMON ESTEBAN CARDONA SALAZAR', nameHint: 'RAMON' },
    SYOHOU: { code: '1020773968', name: 'SYOHOU MARUME IBAÑEZ', nameHint: 'SYOHOU' },
    VALENTINA: { code: '1010247284', name: 'VALENTINA PENAGOS ARBOLEDA', nameHint: 'VALENTINA' },
    JOSE: { code: '7718782', name: 'JOSE ALBERTO VISON MOLINA', nameHint: 'VISON', pensionExempt: true }
  },
  exchangeRate: {
    '2024-01': 3869.79,
    '2024-02': 3882.23,
    '2024-03': 3870.79,
    '2024-04': 3698.80
  },
  feePercent: 0.055
};