/**
 * FINANZAS PERSONALES - Apps Script
 * Base de datos: una hoja de Google Sheets (se crea sola la primera vez).
 *  - Hoja "Movimientos": cada ingreso/gasto registrado.
 *  - Hoja "Programados": ingresos/gastos fijos que se registran solos.
 */

const HOJA_MOV = 'Movimientos';
const HOJA_REC = 'Programados';
const COLS_MOV = ['id', 'fecha', 'tipo', 'clase', 'categoria', 'descripcion', 'monto', 'origenId', 'periodo'];
const COLS_REC = ['id', 'tipo', 'descripcion', 'categoria', 'monto', 'dia', 'inicio', 'fin', 'activo'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Finanzas Personales')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- Base de datos ---------- */

function getSS_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* se recrea */ }
  }
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!ss) ss = SpreadsheetApp.create('Finanzas Personales - Datos');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function getHoja_(nombre, cols) {
  const ss = getSS_();
  let sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // Fechas y periodos como texto para evitar problemas de zona horaria
    cols.forEach((c, i) => {
      if (['fecha', 'inicio', 'fin', 'periodo'].indexOf(c) >= 0) {
        sh.getRange(2, i + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
      }
    });
  }
  return sh;
}

function leer_(nombre, cols) {
  const sh = getHoja_(nombre, cols);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, cols.length).getValues().map(r => {
    const o = {};
    cols.forEach((c, i) => {
      let v = r[i];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      o[c] = v;
    });
    return o;
  });
}

function filaDe_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function hoy_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function nuevoId_() { return Utilities.getUuid().slice(0, 8); }

function validarFecha_(f) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(f || '')) throw new Error('Fecha inválida');
  return f;
}
function validarMonto_(m) {
  m = Number(m);
  if (!isFinite(m) || m <= 0) throw new Error('El monto debe ser mayor que 0');
  return m;
}
function validarTipo_(t) {
  if (t !== 'ingreso' && t !== 'gasto') throw new Error('Tipo inválido');
  return t;
}

/* ---------- API para la interfaz ---------- */

function obtenerTodo() {
  procesarProgramados(); // registra lo que ya toca
  return {
    movimientos: leer_(HOJA_MOV, COLS_MOV),
    programados: leer_(HOJA_REC, COLS_REC),
    hoy: hoy_()
  };
}

/** Registro manual: ingreso/gasto variable (o fijo puntual). */
function guardarMovimiento(d) {
  const sh = getHoja_(HOJA_MOV, COLS_MOV);
  const fila = [
    d.id || nuevoId_(),
    validarFecha_(d.fecha),
    validarTipo_(d.tipo),
    d.clase === 'fijo' ? 'fijo' : 'variable',
    String(d.categoria || 'General').slice(0, 40),
    String(d.descripcion || '').slice(0, 200),
    validarMonto_(d.monto),
    d.origenId || '',
    d.periodo || ''
  ];
  if (d.id) {
    const r = filaDe_(sh, d.id);
    if (r < 0) throw new Error('Movimiento no encontrado');
    sh.getRange(r, 1, 1, fila.length).setValues([fila]);
  } else {
    const r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, fila.length).setNumberFormat('@').setValues([fila]);
    sh.getRange(r, 7).setNumberFormat('#,##0');
  }
  return obtenerTodo();
}

function eliminarMovimiento(id) {
  const sh = getHoja_(HOJA_MOV, COLS_MOV);
  const r = filaDe_(sh, id);
  if (r > 0) sh.deleteRow(r);
  return obtenerTodo();
}

/** Crea o edita un ingreso/gasto fijo programado. */
function guardarProgramado(d) {
  const sh = getHoja_(HOJA_REC, COLS_REC);
  const dia = Math.round(Number(d.dia));
  if (!(dia >= 1 && dia <= 31)) throw new Error('El día debe estar entre 1 y 31');
  const fila = [
    d.id || nuevoId_(),
    validarTipo_(d.tipo),
    String(d.descripcion || '').slice(0, 200),
    String(d.categoria || 'General').slice(0, 40),
    validarMonto_(d.monto),
    dia,
    validarFecha_(d.inicio),
    d.fin ? validarFecha_(d.fin) : '',
    d.activo === false ? 'no' : 'si'
  ];
  if (!fila[2]) throw new Error('Agrega una descripción');
  if (d.id) {
    const r = filaDe_(sh, d.id);
    if (r < 0) throw new Error('Programado no encontrado');
    sh.getRange(r, 1, 1, fila.length).setNumberFormat('@').setValues([fila]);
    sh.getRange(r, 5).setNumberFormat('#,##0');
  } else {
    const r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, fila.length).setNumberFormat('@').setValues([fila]);
    sh.getRange(r, 5).setNumberFormat('#,##0');
  }
  return obtenerTodo();
}

/**
 * Elimina un programado. Lo ya registrado se conserva como historial;
 * solo deja de generarse hacia adelante.
 */
function eliminarProgramado(id) {
  const sh = getHoja_(HOJA_REC, COLS_REC);
  const r = filaDe_(sh, id);
  if (r > 0) sh.deleteRow(r);
  return obtenerTodo();
}

function pausarProgramado(id, activo) {
  const sh = getHoja_(HOJA_REC, COLS_REC);
  const r = filaDe_(sh, id);
  if (r < 0) throw new Error('Programado no encontrado');
  sh.getRange(r, 9).setValue(activo ? 'si' : 'no');
  return obtenerTodo();
}

/* ---------- Registro automático ---------- */

/**
 * Para cada programado activo, crea el movimiento de cada mes cuya fecha ya
 * llegó. Es idempotente: no duplica (clave = programado + mes). Cambiar el
 * monto de un programado solo afecta a los meses que aún no se han registrado.
 */
function procesarProgramados() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const progs = leer_(HOJA_REC, COLS_REC).filter(p => p.activo !== 'no');
    if (!progs.length) return 0;
    const movs = leer_(HOJA_MOV, COLS_MOV);
    const hechos = {};
    movs.forEach(m => { if (m.origenId) hechos[m.origenId + '|' + m.periodo] = true; });

    const hoy = hoy_();
    const nuevos = [];
    progs.forEach(p => {
      let [y, mo] = String(p.inicio).split('-').map(Number);
      let guard = 0;
      while (guard++ < 600) {
        const ultimo = new Date(y, mo, 0).getDate();
        const dd = Math.min(Number(p.dia), ultimo);
        const fecha = y + '-' + ('0' + mo).slice(-2) + '-' + ('0' + dd).slice(-2);
        if (fecha > hoy) break;
        if (p.fin && fecha > p.fin) break;
        const periodo = y + '-' + ('0' + mo).slice(-2);
        if (fecha >= p.inicio && !hechos[p.id + '|' + periodo]) {
          nuevos.push([nuevoId_(), fecha, p.tipo, 'fijo', p.categoria, p.descripcion,
            Number(p.monto), p.id, periodo]);
          hechos[p.id + '|' + periodo] = true;
        }
        mo++; if (mo > 12) { mo = 1; y++; }
      }
    });
    if (nuevos.length) {
      const sh = getHoja_(HOJA_MOV, COLS_MOV);
      const r = sh.getLastRow() + 1;
      sh.getRange(r, 1, nuevos.length, COLS_MOV.length).setNumberFormat('@').setValues(nuevos);
      sh.getRange(r, 7, nuevos.length, 1).setNumberFormat('#,##0');
    }
    return nuevos.length;
  } finally {
    lock.releaseLock();
  }
}

/** Ejecuta UNA vez desde el editor: crea el disparador diario (6 a.m.). */
function instalarDisparadorDiario() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'procesarProgramados')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('procesarProgramados').timeBased().everyDays(1).atHour(6).create();
}
