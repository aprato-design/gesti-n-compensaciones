/**
 * Compensaciones 2026 — Backend Apps Script (GENERAL)
 * v1 — Vista completa de la compañía (todos los departamentos), sin exclusiones.
 */

// ════════════════════════════════════════════════════════════════════
//   CONFIG
// ════════════════════════════════════════════════════════════════════

var SHEET_ID  = '15mTvxuwXGa8B8Cuc86qqDyb8eKhKrDlwRekx9I1dZ_0';
var SHEET_TAB = 'Acciones';
var DATA_TAB  = 'Bamboo Compensaciones - Sueldos';

// Tab de la app performance_cases.py — mismo spreadsheet (SHEET_ID).
var CASOS_TAB = 'Casos Performance';

// Sheet externo del que se toma Account/Proyecto + datos de último LC (cruce por email).
var LC_SHEET_ID  = '1goxKg8LylMSZcVzkPVBPczknKlFbB8WBYge_XpAQxTw';
var LC_SHEET_GID = 551301189;

var EMAILS_EDITORES = [
  'aprato@makingsense.com',
  'talentcare@makingsense.com',
];

var EMAILS_VIEWERS = [
  'jgasbarro@makingsense.com',
  'aguerreiro@makingsense.com',
];

// Personas que no deben figurar en este tablero (comparación normalizada: sin acentos,
// sin mayúsculas, sin comas — el orden Apellido/Nombre no importa).
var EXCLUDED_EMPLOYEES = [
  'Consoli Jorge',
  'Guerreiro Agustina',
  'Peña Pereira Luis',
  'Raimondi Juan',
  'Beherengaray Calvo Gaston',
  'Dos Reis German',
  'Antonelli Ignacio',
  'Debert Alexander',
];

// ════════════════════════════════════════════════════════════════════
//   ENTRY
// ════════════════════════════════════════════════════════════════════

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Compensaciones 2026 — General')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ════════════════════════════════════════════════════════════════════
//   ROLES
// ════════════════════════════════════════════════════════════════════

function getUserRole_() {
  var email = (Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  var i;
  for (i = 0; i < EMAILS_EDITORES.length; i++) {
    if (EMAILS_EDITORES[i].toLowerCase().trim() === email) return { role: 'editor', email: email };
  }
  for (i = 0; i < EMAILS_VIEWERS.length; i++) {
    if (EMAILS_VIEWERS[i].toLowerCase().trim() === email) return { role: 'viewer', email: email };
  }
  return { role: 'denied', email: email };
}

// ════════════════════════════════════════════════════════════════════
//   API PÚBLICA
// ════════════════════════════════════════════════════════════════════

function getInitialData() {
  var u = getUserRole_();
  if (u.role === 'denied') {
    return { role: 'denied', email: u.email, actions: [], data: [] };
  }
  return {
    role:    u.role,
    email:   u.email,
    actions: readAllActions_(),
    data:    getCompData_()
  };
}

/**
 * Guarda o actualiza una acción + nota para un empleado en un mes.
 * El frontend llama: saveAction(month, name, actionValue, noteValue)
 * Preserva LC1/LC2 existentes; si TODOS los campos quedan vacíos, elimina la fila.
 */
function saveAction(month, name, actionValue, noteValue) {
  var u = getUserRole_();
  if (u.role !== 'editor') {
    throw new Error('Sin permiso (rol: ' + u.role + ')');
  }

  var sheet    = getActionsSheet_();
  var now      = new Date();
  var foundRow = _findActionRow_(sheet, month, name);

  var newAction = (actionValue || '').trim();
  var newNote   = (noteValue   || '');
  var existing  = foundRow > 0 ? sheet.getRange(foundRow, 5, 1, 4).getValues()[0] : ['', false, '', false];
  var merged    = [newAction, newNote, existing[0], !!existing[1], existing[2], !!existing[3]];

  if (foundRow > 0) {
    if (_rowIsEmpty_(merged)) {
      sheet.deleteRow(foundRow);
    } else {
      sheet.getRange(foundRow, 3, 1, 2).setValues([[newAction, newNote]]);
      sheet.getRange(foundRow, 9, 1, 2).setValues([[u.email, now]]);
    }
  } else if (!_rowIsEmpty_(merged)) {
    sheet.appendRow([month, name, newAction, newNote, '', false, '', false, u.email, now]);
  }

  return true;
}

/**
 * Guarda el valor + checkbox "comunicado" de la 1ra o 2da LC del año para un
 * empleado en un mes. El frontend llama: saveLC(month, name, slot, value, comunicado)
 * slot: 1 o 2. Preserva Accion/Nota/la otra LC existentes.
 */
function saveLC(month, name, slot, value, comunicado) {
  var u = getUserRole_();
  if (u.role !== 'editor') {
    throw new Error('Sin permiso (rol: ' + u.role + ')');
  }
  if (slot !== 1 && slot !== 2) {
    throw new Error('slot inválido (debe ser 1 o 2)');
  }

  var sheet    = getActionsSheet_();
  var now      = new Date();
  var foundRow = _findActionRow_(sheet, month, name);

  var newVal = (value || '').trim();
  var newCom = !!comunicado;
  var existing = foundRow > 0
    ? sheet.getRange(foundRow, 3, 1, 6).getValues()[0]   // Accion,Nota,LC1,LC1Com,LC2,LC2Com
    : ['', '', '', false, '', false];

  var lc1  = slot === 1 ? newVal : existing[2];
  var lc1c = slot === 1 ? newCom : !!existing[3];
  var lc2  = slot === 2 ? newVal : existing[4];
  var lc2c = slot === 2 ? newCom : !!existing[5];
  var merged = [existing[0], existing[1], lc1, lc1c, lc2, lc2c];

  if (foundRow > 0) {
    if (_rowIsEmpty_(merged)) {
      sheet.deleteRow(foundRow);
    } else {
      sheet.getRange(foundRow, 5, 1, 4).setValues([[lc1, lc1c, lc2, lc2c]]);
      sheet.getRange(foundRow, 9, 1, 2).setValues([[u.email, now]]);
    }
  } else if (!_rowIsEmpty_(merged)) {
    sheet.appendRow([month, name, '', '', lc1, lc1c, lc2, lc2c, u.email, now]);
  }

  return true;
}

function _findActionRow_(sheet, month, name) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var range = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < range.length; i++) {
    if (String(range[i][0]) === String(month) && String(range[i][1]) === String(name)) {
      return i + 2;
    }
  }
  return -1;
}

/** vals: [accion, nota, lc1, lc1Comunicado, lc2, lc2Comunicado] */
function _rowIsEmpty_(vals) {
  return !vals[0] && !String(vals[1] || '').trim() && !vals[2] && !vals[3] && !vals[4] && !vals[5];
}

// ════════════════════════════════════════════════════════════════════
//   DATOS DE COMPENSACIÓN
// ════════════════════════════════════════════════════════════════════

/**
 * Lee el sheet externo de LC (Account/Proyecto, Fecha último LC, Aging LC, Status LC)
 * y arma un mapa por email (lowercase, trim) -> {account, fechaLC, agingLC, statusLC}.
 * Columnas: A=Account/Proyecto, C=Email, K=Fecha último LC, L=Aging LC (meses), M=Status LC.
 * Si un email aparece en más de una fila, gana la última (se asume la más reciente).
 */
function getLCData_() {
  var ss     = SpreadsheetApp.openById(LC_SHEET_ID);
  var sheets = ss.getSheets();
  var sh     = null;
  for (var s = 0; s < sheets.length; s++) {
    if (sheets[s].getSheetId() === LC_SHEET_GID) { sh = sheets[s]; break; }
  }
  if (!sh) return {};

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  var data = sh.getRange(1, 1, lastRow, 13).getValues(); // A..M
  var map  = {};
  for (var i = 1; i < data.length; i++) {
    var row   = data[i];
    var email = String(row[2] || '').toLowerCase().trim(); // col C
    if (!email) continue;
    map[email] = {
      account:  row[0]  || '',              // col A
      fechaLC:  _fmtLCDate_(row[10]),        // col K
      // col L — colores fijos por rango, ver agingLcBadge() en el frontend (no se lee color de la celda)
      agingLC:  row[11] === '' || row[11] == null ? '' : row[11],
      statusLC: row[12] || '',              // col M — colores fijos por valor, ver STATUS_LC_CLASS en el frontend
    };
  }
  return map;
}

function _fmtLCDate_(v) {
  if (v instanceof Date) {
    var dd = v.getDate(), mm = v.getMonth() + 1, yy = v.getFullYear();
    return (dd < 10 ? '0' + dd : dd) + '/' + (mm < 10 ? '0' + mm : mm) + '/' + yy;
  }
  return v || '';
}

/**
 * Normaliza un nombre para comparar de forma robusta contra EXCLUDED_EMPLOYEES: saca
 * acentos, pasa a minúsculas, saca comas/puntuación y ordena las palabras — así no importa
 * si viene como "Apellido, Nombre" o "Nombre Apellido", ni diferencias de acentos/espacios.
 */
function _normalizeName_(s) {
  var n = String(s || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[^a-z\s]/g, ' ')
    .trim();
  if (!n) return '';
  return n.split(/\s+/).sort().join(' ');
}

var EXCLUDED_EMPLOYEES_NORM_ = EXCLUDED_EMPLOYEES.map(_normalizeName_);

function _isExcludedEmployee_(fullName) {
  var norm = _normalizeName_(fullName);
  if (!norm) return false;
  return EXCLUDED_EMPLOYEES_NORM_.indexOf(norm) !== -1;
}

// ════════════════════════════════════════════════════════════════════
//   CASOS PERFORMANCE (app performance_cases.py)
// ════════════════════════════════════════════════════════════════════

/**
 * Lee la tab "Casos Performance" (misma spreadsheet que Compensaciones) y arma un mapa
 * email → lista de casos [{createdAt, closedAt, status}]. Un empleado puede tener varios
 * casos a lo largo del tiempo (uno cerrado, otro abierto después).
 */
function getCasosData_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(CASOS_TAB);
  if (!sh) return {};

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  var lastCol = sh.getLastColumn();
  var data    = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data[0];
  var emailIdx   = headers.indexOf('employee_email');
  var createdIdx = headers.indexOf('created_at');
  var closedIdx  = headers.indexOf('closed_at');
  var statusIdx  = headers.indexOf('status');
  if (emailIdx === -1 || createdIdx === -1) return {};

  var map = {};
  for (var i = 1; i < data.length; i++) {
    var row   = data[i];
    var email = String(row[emailIdx] || '').toLowerCase().trim();
    if (!email) continue;
    var createdAt = _parseCasoDate_(row[createdIdx]);
    if (!createdAt) continue;
    var closedAt = closedIdx > -1 ? _parseCasoDate_(row[closedIdx]) : null;
    var status   = statusIdx > -1 ? String(row[statusIdx] || '') : '';
    if (!map[email]) map[email] = [];
    map[email].push({ createdAt: createdAt, closedAt: closedAt, status: status });
  }
  return map;
}

/** performance_cases.py escribe created_at/closed_at como texto "YYYY-MM-DD HH:MM". */
function _parseCasoDate_(v) {
  if (v instanceof Date) return v;
  var s = String(v || '').trim();
  if (!s) return null;
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]);
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  return null;
}

/**
 * ¿Había algún caso abierto para esta persona al momento "cutoff"? Se basa únicamente en
 * created_at/closed_at (hechos inmutables), no en el status actual — así un mes pasado
 * queda fijo aunque el caso se haya cerrado después, y el mes en curso es dinámico si
 * "cutoff" es "ahora".
 */
function _hasOpenCaseAt_(cases, cutoff) {
  if (!cases) return false;
  for (var i = 0; i < cases.length; i++) {
    var c = cases[i];
    if (c.createdAt.getTime() > cutoff.getTime()) continue;
    if (!c.closedAt || c.closedAt.getTime() > cutoff.getTime()) return true;
  }
  return false;
}

function getCompData_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(DATA_TAB);
  if (!sh) return [];

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var lcMap = getLCData_();

  var numCols = 36;
  var data    = sh.getRange(1, 1, lastRow, numCols).getValues();
  var headers = data[0];
  var result  = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] && !row[1]) continue;

    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var v = row[j];
      if (v === undefined || v === null) {
        obj[headers[j]] = '';
      } else if (v instanceof Date && headers[j] === 'Ultimo Cambio Proactivo') {
        var MONTH_ABBR_EN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
        obj[headers[j]] = MONTH_ABBR_EN[v.getMonth()] + '-' + v.getFullYear();
      } else if (v instanceof Date && headers[j] === 'Hire Date') {
        // Hire Date se carga desde BambooHR como MM/DD/YYYY (ver _fmt_date en
        // actualizar_mes_sheets.py) — mantener ese orden, no DD/MM/YYYY como el resto.
        var hdd = v.getDate(), hmm = v.getMonth() + 1, hyy = v.getFullYear();
        obj[headers[j]] = (hmm < 10 ? '0' + hmm : hmm) + '/' + (hdd < 10 ? '0' + hdd : hdd) + '/' + hyy;
      } else if (v instanceof Date) {
        var d  = v;
        var dd = d.getDate();
        var mm = d.getMonth() + 1;
        var yy = d.getFullYear();
        obj[headers[j]] = (dd < 10 ? '0' + dd : dd) + '/' + (mm < 10 ? '0' + mm : mm) + '/' + yy;
      } else {
        obj[headers[j]] = v;
      }
    }

    var rawMonth = String(obj['Month'] || '').trim();
    var mMatch;
    mMatch = rawMonth.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
    if (mMatch) {
      obj['Month'] = mMatch[3] + '-' + mMatch[2];
    } else {
      mMatch = rawMonth.match(/^(\d{4})-(\d{2})/);
      if (mMatch) {
        obj['Month'] = mMatch[1] + '-' + mMatch[2];
      }
    }

    if (_isExcludedEmployee_(obj['Last name, First name'])) continue;

    var lcEmail = String(obj['Employee #'] || '').toLowerCase().trim();
    var lc = lcMap[lcEmail];
    obj['Account / Proyecto']  = lc ? lc.account  : '';
    obj['Fecha ultimo LC']     = lc ? lc.fechaLC  : '';
    obj['Aging LC (meses)']    = lc ? lc.agingLC  : '';
    obj['Status LC']           = lc ? lc.statusLC : '';

    result.push(obj);
  }

  // Caso Abierto: fijo para meses pasados (a fin de ese mes), dinámico ("ahora") para el
  // último mes presente en la data — ver _hasOpenCaseAt_().
  var casosMap  = getCasosData_();
  var maxMonth  = '';
  for (var mi = 0; mi < result.length; mi++) {
    if (result[mi].Month > maxMonth) maxMonth = result[mi].Month;
  }
  var now = new Date();
  for (var ri = 0; ri < result.length; ri++) {
    var r = result[ri];
    var remail = String(r['Employee #'] || '').toLowerCase().trim();
    var cutoff;
    if (r.Month === maxMonth) {
      cutoff = now;
    } else {
      var ym = String(r.Month).split('-');
      cutoff = new Date(parseInt(ym[0], 10), parseInt(ym[1], 10), 0, 23, 59, 59);
    }
    r['Caso Abierto'] = _hasOpenCaseAt_(casosMap[remail], cutoff);
  }

  return result;
}

// ════════════════════════════════════════════════════════════════════
//   SHEET HELPERS
// ════════════════════════════════════════════════════════════════════

var ACTIONS_HEADERS_ = ['Mes', 'Nombre', 'Accion', 'Nota', 'LC1', 'LC1Comunicado', 'LC2', 'LC2Comunicado', 'EditadoPor', 'EditadoEn'];

function getActionsSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_TAB);
  var n  = ACTIONS_HEADERS_.length;

  if (!sh) {
    sh = ss.insertSheet(SHEET_TAB);
    sh.getRange(1, 1, 1, n).setValues([ACTIONS_HEADERS_]);
    sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#102532').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, n, 140);
    sh.setColumnWidth(4, 320);
    return sh;
  }

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, n).setValues([ACTIONS_HEADERS_]);
    sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#102532').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    return sh;
  }

  // Migración: esquema viejo (6 cols, sin LC1/LC2) -> nuevo (10 cols). Inserta las 4
  // columnas de LC antes de EditadoPor/EditadoEn, que se corren solas a I/J.
  var lastCol = sh.getLastColumn();
  if (lastCol < n) {
    var current = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    if (current[2] === 'Accion' && current[3] === 'Nota') {
      sh.insertColumnsBefore(5, 4);
      sh.getRange(1, 5, 1, 4).setValues([['LC1', 'LC1Comunicado', 'LC2', 'LC2Comunicado']]);
      sh.getRange(1, 5, 1, 4).setFontWeight('bold').setBackground('#102532').setFontColor('#ffffff');
      sh.setColumnWidths(5, 4, 140);
    }
  }
  return sh;
}

function readAllActions_() {
  var sh   = getActionsSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var data   = sh.getRange(2, 1, last - 1, 8).getValues();
  var result = [];
  for (var i = 0; i < data.length; i++) {
    if (!data[i][0] || !data[i][1]) continue;
    var mesRaw = data[i][0];
    var mes;
    if (mesRaw instanceof Date) {
      var yy = mesRaw.getFullYear();
      var mm = mesRaw.getMonth() + 1;
      mes = yy + '-' + (mm < 10 ? '0' + mm : '' + mm);
    } else {
      var s = String(mesRaw).trim();
      var m1 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
      if (m1) {
        mes = m1[3] + '-' + m1[1];
      } else {
        mes = s;
      }
    }
    result.push({
      Mes:    mes,
      Nombre: String(data[i][1]),
      Accion: String(data[i][2] || ''),
      Nota:   String(data[i][3] || ''),
      LC1:           String(data[i][4] || ''),
      LC1Comunicado: !!data[i][5],
      LC2:           String(data[i][6] || ''),
      LC2Comunicado: !!data[i][7],
    });
  }
  return result;
}

// ════════════════════════════════════════════════════════════════════
//   DIAGNÓSTICO
// ════════════════════════════════════════════════════════════════════

function _whoAmI() { Logger.log(JSON.stringify(getUserRole_())); }

function _dumpData() {
  var d  = getCompData_();
  var ms = {};
  for (var i = 0; i < d.length; i++) { ms[d[i].Month] = true; }
  var meses = Object.keys(ms).sort();
  Logger.log('Filas: ' + d.length + ' | Meses: ' + JSON.stringify(meses));
  if (d.length > 0) Logger.log('Primer obj: ' + JSON.stringify(d[0]));
}

function _dumpActions() {
  var a = readAllActions_();
  Logger.log('Total acciones: ' + a.length);
}
