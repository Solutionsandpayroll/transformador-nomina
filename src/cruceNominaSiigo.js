// ============================================================================
// cruceNominaSiigo.js
// Cruza los registros en formato largo de la nómina/facturación con los del
// Movimiento CC de Siigo, por empresa, mes, persona y concepto.
// ============================================================================
// Persona: se identifica por DOCUMENTO (cédula/NIT) cuando existe en ambos lados.
// Si un lado no trae documento, se busca por nombre (palabras en cualquier
// orden) contra los nombres que sí tienen documento; si no, se cruza por nombre.
// ============================================================================

export const CRUCE_COLORS = {
  OK: '92D050',
  DIFERENCIA: 'FF0000',
  SOLO_NOMINA: 'FFFF00',
  SOLO_SIIGO: 'FFFF00',
  CRUZA_ENTRE_MESES: '00FF00'
};

export const CRUCE_LABEL = {
  OK: 'Cuadra',
  DIFERENCIA: 'Diferencia',
  SOLO_NOMINA: 'Solo en nómina',
  SOLO_SIIGO: 'Solo en Siigo',
  CRUZA_ENTRE_MESES: 'Cuadra acumulado'
};

// Conceptos que Siigo registra en un mes distinto al de la nómina: se comparan acumulados.
const ACCUMULATED = new Set(['13TH SALARY', '14TH SALARY', 'INTEREST ON 14TH SALARY']);

const TOTAL = 'TOTAL EMPLOYEE COST';

function norm(v) {
  return String(v ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/\s+/g, ' ')
    .trim();
}

const round2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
const digits = (v) => String(v ?? '').replace(/\D/g, '');

function canonName(v) {
  return norm(v).replace(/[^A-Z0-9 ]/g, ' ').split(/\s+/).filter(Boolean).sort().join(' ');
}

// Nombres de concepto distintos en nómina y Siigo -> uno común.
export function canonConcept(value) {
  const t = norm(value);
  if (!t) return '';
  if (t === TOTAL) return TOTAL;
  if (/PRIMA|13TH/.test(t)) return '13TH SALARY';
  if (/INTEREST|INTERESES/.test(t)) return 'INTEREST ON 14TH SALARY';
  if (/14TH|CESANTIA/.test(t)) return '14TH SALARY';
  // Seguros y examen médico: Siigo los registra juntos como "Seguro(s) de vida".
  if (/INSURANCE|SEGURO|POLIZA|MEDICAL TEST/.test(t)) return 'HEALTH INSURANCE + MEDICAL TEST';
  if (/GASTOS|EXPENSES/.test(t)) return 'EXPENSES';
  if (/ALLOWANCE|AUXILIO|GROSS UP/.test(t) && !/TRANSPORT/.test(t)) return 'ALLOWANCE';
  if (/TRANSPORT/.test(t)) return 'TRANSPORT ALLOWANCE';
  if (/SALARY|SALARIO/.test(t)) return 'SALARY'; // incluye INTEGRATED / ORDINARY SALARY
  if (/PENSION/.test(t)) return 'PENSION COST';
  if (/HEALTH COST|SALUD|\bEPS\b/.test(t)) return 'HEALTH COST';
  if (/FAMILY FUND|CAJA/.test(t)) return 'FAMILY FUND COST';
  if (/LABOR RISK|RIESGO|\bARL\b/.test(t)) return 'LABOR RISK COST';
  if (/SENA/.test(t)) return 'SENA COST';
  if (/ICBF/.test(t)) return 'ICBF COST';
  if (/VACATION|ANUAL LEAVE|VACACION/.test(t)) return 'VACATION';
  if (/PAID LEAVE|LICENCIA/.test(t)) return 'PAID LEAVE';
  return t;
}

function monthKey(d) {
  if (!(d instanceof Date) || Number.isNaN(d.getTime())) return null;
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`;
}

export function cruzarNominaSiigo(nominaRecords = [], siigoRecords = [], options = {}) {
  const tol = Number.isFinite(options.tolerance) ? options.tolerance : 1;

  // nombre canónico -> documento (primero lo que dice la nómina, luego Siigo)
  const docByName = new Map();
  for (const r of [...nominaRecords, ...siigoRecords]) {
    const d = digits(r.Documento);
    const cn = canonName(r.Empleado);
    if (d && cn && !docByName.has(cn)) docByName.set(cn, d);
  }

  // Nombre incompleto ("Ramon Cardona") -> documento, si sus palabras están en
  // un único nombre completo conocido.
  const docBySubset = (name) => {
    const t = canonName(name).split(' ').filter(Boolean);
    if (t.length < 2) return '';
    const hits = [...docByName].filter(([cn]) => {
      const k = cn.split(' ');
      return t.every((x) => k.includes(x));
    });
    return hits.length === 1 ? hits[0][1] : '';
  };

  const personOf = (r) => {
    const cn = canonName(r.Empleado);
    const d = digits(r.Documento) || docByName.get(cn) || docBySubset(r.Empleado);
    return d ? `D:${d}` : `N:${cn}`;
  };

  // grupo (empresa|persona|concepto) -> mes -> { n, s }
  const groups = new Map();
  const names = new Map(); // persona -> { nombre, doc }

  const add = (records, side) => {
    for (const r of records) {
      const mes = monthKey(r['Mes elaboración']);
      const concepto = canonConcept(r.Concepto);
      if (!mes || !concepto) continue;
      const emp = String(r.Empleado ?? '').trim();
      if (!emp || /SIN ASIGNAR|SIN TERCERO/.test(norm(emp))) continue;
      const valor = round2(concepto === TOTAL ? r['Valor Totales'] : r['Valor Concepto']);
      if (valor === 0) continue;

      const person = personOf(r);
      const empresa = norm(r.Empresa);
      const gk = `${empresa}|${person}|${concepto}`;
      if (!groups.has(gk)) groups.set(gk, { empresa: r.Empresa || '', person, concepto, months: new Map() });
      const g = groups.get(gk);
      if (!g.months.has(mes)) g.months.set(mes, { n: 0, s: 0, hasN: false, hasS: false });
      const cell = g.months.get(mes);
      if (side === 'n') { cell.n = round2(cell.n + valor); cell.hasN = true; }
      else { cell.s = round2(cell.s + valor); cell.hasS = true; }

      const prev = names.get(person);
      if (!prev || side === 'n') {
        names.set(person, { nombre: emp, doc: digits(r.Documento) || (person.startsWith('D:') ? person.slice(2) : '') });
      }
    }
  };
  add(nominaRecords, 'n');
  add(siigoRecords, 's');

  const rows = [];
  for (const g of groups.values()) {
    const acc = ACCUMULATED.has(g.concepto);
    let cumN = 0;
    let cumS = 0;
    for (const mes of [...g.months.keys()].sort()) {
      const c = g.months.get(mes);
      cumN = round2(cumN + c.n);
      cumS = round2(cumS + c.s);
      const vN = acc ? cumN : c.n;
      const vS = acc ? cumS : c.s;
      if (vN === 0 && vS === 0) continue;

      const dif = round2(vN - vS);
      let estado;
      if (Math.abs(dif) <= tol) {
        estado = acc && Math.abs(c.n - c.s) > tol ? 'CRUZA_ENTRE_MESES' : 'OK';
      } else if (c.hasN && !c.hasS && !acc) estado = 'SOLO_NOMINA';
      else if (!c.hasN && c.hasS && !acc) estado = 'SOLO_SIIGO';
      else estado = 'DIFERENCIA';

      const info = names.get(g.person) || {};
      rows.push({
        Empresa: g.empresa,
        Mes: mes,
        Documento: info.doc || '',
        Empleado: info.nombre || '',
        Concepto: g.concepto,
        'Valor Nómina': vN,
        'Valor Siigo': vS,
        Diferencia: dif,
        Estado: estado,
        Acumulado: acc,
        _fill: CRUCE_COLORS[estado]
      });
    }
  }

  rows.sort(
    (a, b) =>
      a.Empresa.localeCompare(b.Empresa) ||
      a.Mes.localeCompare(b.Mes) ||
      a.Empleado.localeCompare(b.Empleado) ||
      a.Concepto.localeCompare(b.Concepto)
  );

  const summary = { total: rows.length, OK: 0, DIFERENCIA: 0, SOLO_NOMINA: 0, SOLO_SIIGO: 0, CRUZA_ENTRE_MESES: 0 };
  for (const r of rows) summary[r.Estado] += 1;

  return { rows, summary };
}