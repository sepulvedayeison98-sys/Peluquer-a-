/**
 * BARBER STYLE — Google Apps Script Backend
 * Despliega como Web App → Acceso: Cualquier persona
 *
 * COLUMNAS de la hoja "Citas":
 * A:id  B:code  C:name  D:email  E:tel  F:svc  G:price
 * H:dur  I:date  J:time  K:status  L:notes  M:createdAt
 */

const SHEET_NAME = 'Citas';
const COL = {
  id:1, code:2, name:3, email:4, tel:5, svc:6,
  price:7, dur:8, date:9, time:10, status:11, notes:12, createdAt:13
};

// ─── Respuestas JSON (Apps Script maneja CORS automáticamente) ────
function ok(data) {
  return ContentService.createTextOutput(JSON.stringify({ok:true,data}))
    .setMimeType(ContentService.MimeType.JSON);
}
function err(msg) {
  return ContentService.createTextOutput(JSON.stringify({ok:false,error:msg}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── GET handler ──────────────────────────────────────────────────
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || 'getAll';
    if (action === 'getAll') return ok(getAllApts());
    if (action === 'ping')   return ok({status:'alive',ts:new Date().toISOString()});
    return err('Accion GET no reconocida: ' + action);
  } catch(ex) { return err(ex.message); }
}

// ─── POST handler ─────────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'save')         return ok(saveApt(body.apt));
    if (action === 'updateStatus') return ok(updateStatus(body.code, body.status));
    if (action === 'reschedule')   return ok(rescheduleApt(body.code, body.date, body.time));
    if (action === 'cancel')       return ok(updateStatus(body.code, 'cancelled'));
    return err('Accion POST no reconocida: ' + action);
  } catch(ex) { return err(ex.message); }
}

// ─── Obtener todas las citas ──────────────────────────────────────
function getAllApts() {
  const sheet = getSheet();
  const last  = sheet.getLastRow();
  if (last < 2) return [];
  const data = sheet.getRange(2, 1, last-1, 13).getValues();
  return data.filter(r => r[COL.code-1]).map(rowToObj);
}

// ─── Guardar nueva cita ───────────────────────────────────────────
function saveApt(apt) {
  const sheet = getSheet();
  const id    = getNextId(sheet);
  sheet.appendRow([id, apt.code||'', apt.name||'', apt.email||'', apt.tel||'',
    apt.svc||'', Number(apt.price)||0, apt.dur||'', apt.date||'', apt.time||'',
    apt.status||'pending', apt.notes||'', apt.createdAt||new Date().toISOString()]);
  return {id, code:apt.code};
}

// ─── Actualizar estado ────────────────────────────────────────────
function updateStatus(code, newStatus) {
  const row = findRowByCode(code);
  if (!row) throw new Error('Cita no encontrada: ' + code);
  getSheet().getRange(row, COL.status).setValue(newStatus);
  return {code, status:newStatus};
}

// ─── Reagendar ────────────────────────────────────────────────────
function rescheduleApt(code, newDate, newTime) {
  const row = findRowByCode(code);
  if (!row) throw new Error('Cita no encontrada: ' + code);
  const sheet = getSheet();
  sheet.getRange(row, COL.date).setValue(newDate);
  sheet.getRange(row, COL.time).setValue(newTime);
  sheet.getRange(row, COL.status).setValue('rescheduled');
  return {code, date:newDate, time:newTime, status:'rescheduled'};
}

// ─── Helpers ──────────────────────────────────────────────────────
function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SHEET_ID');
  if (ssId) {
    try { return SpreadsheetApp.openById(ssId); } catch(e) {}
  }
  // Crear nueva hoja automáticamente si no existe
  const ss = SpreadsheetApp.create('Barber Style — Citas');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function getSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id','code','name','email','tel','svc','price','dur','date','time','status','notes','createdAt']);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,13)
      .setBackground('#1a1410').setFontColor('#c9a96e').setFontWeight('bold');
  }
  return sheet;
}

function getNextId(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return 1;
  const ids = sheet.getRange(2, COL.id, last-1, 1).getValues().map(r => Number(r[0])||0);
  return Math.max(...ids) + 1;
}

function findRowByCode(code) {
  const sheet = getSheet();
  const last  = sheet.getLastRow();
  if (last < 2) return null;
  const codes = sheet.getRange(2, COL.code, last-1, 1).getValues();
  for (let i=0; i<codes.length; i++) {
    if (codes[i][0] === code) return i+2;
  }
  return null;
}

function rowToObj(r) {
  return {id:r[0],code:r[1],name:r[2],email:r[3],tel:r[4],svc:r[5],
    price:Number(r[6])||0,dur:r[7],date:r[8],time:r[9],status:r[10],
    notes:r[11],createdAt:r[12]?new Date(r[12]).toISOString():''};
}
