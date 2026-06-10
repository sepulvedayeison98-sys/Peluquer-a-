/**
 * WHITE LABEL — Google Apps Script Backend
 * Despliega como Web App → Acceso: Cualquier persona
 *
 * ⚙️  CONFIGURACIÓN POR CLIENTE — edita solo estas 3 constantes:
 *    BUSINESS_NAME  → nombre del negocio (usado en emails)
 *    BARBERO_EMAIL  → correo que recibe notificaciones
 *    N8N_WEBHOOK    → URL del webhook de N8N del cliente
 *
 * COLUMNAS de la hoja "Citas":
 * A:id  B:code  C:name  D:email  E:tel  F:svc  G:price
 * H:dur  I:date  J:time  K:status  L:notes  M:createdAt
 */

const BUSINESS_NAME = 'Barber Style';          // ← nombre del negocio
const SHEET_NAME    = 'Citas';
const BARBERO_EMAIL = 'sepulvedayeison98@gmail.com'; // ← correo del negocio
const N8N_WEBHOOK   = 'https://kodrefe.app.n8n.cloud/webhook/barber-style'; // ← webhook N8N
const COL = {
  id:1, code:2, name:3, email:4, tel:5, svc:6,
  price:7, dur:8, date:9, time:10, status:11, notes:12, createdAt:13
};

// ─── Respuestas JSON ──────────────────────────────────────────────
function ok(data) {
  return ContentService.createTextOutput(JSON.stringify({ok:true,data}))
    .setMimeType(ContentService.MimeType.JSON);
}
function err(msg) {
  return ContentService.createTextOutput(JSON.stringify({ok:false,error:msg}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── GET handler ──────────────────────────────────────────────────
// Nota: los navegadores convierten POST→GET en redirects 302.
// Para evitar ese problema, el frontend envía las mutaciones como GET
// con el payload codificado en el parámetro "p".
function doGet(e) {
  try {
    // Mutaciones enviadas desde el navegador (workaround redirect POST→GET)
    if (e.parameter && e.parameter.p) {
      const body = JSON.parse(e.parameter.p);
      return handleAction(body);
    }
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
    return handleAction(body);
  } catch(ex) { return err(ex.message); }
}

// ─── Lógica de acciones compartida por doGet y doPost ─────────────
function handleAction(body) {
  const action = body.action;
  if (action === 'save')           return ok(saveApt(body.apt));
  if (action === 'updateStatus')   return ok(updateStatus(body.code, body.status));
  if (action === 'reschedule')     return ok(rescheduleApt(body.code, body.date, body.time));
  if (action === 'cancel')         return ok(cancelApt(body.code));
  if (action === 'verifyLogin')    return ok(verifyLogin(body.pass));
  if (action === 'changePassword') return ok(changePassword(body.currentPass, body.newPass));
  if (action === 'clearOld')       return ok(clearOldApts(body.days));
  return err('Accion no reconocida: ' + action);
}

// ─── Obtener todas las citas ──────────────────────────────────────
function getAllApts() {
  const sheet = getSheet();
  const last  = sheet.getLastRow();
  if (last < 2) return [];
  const data = sheet.getRange(2, 1, last-1, 13).getValues();
  return data.filter(r => r[COL.code-1]).map(rowToObj);
}

// ─── Disparar webhook N8N ─────────────────────────────────────────
function notifyN8N(payload) {
  if (!N8N_WEBHOOK) return;
  try {
    // deadline:10 evita que el webhook bloquee la respuesta al cliente más de 10 segundos
    UrlFetchApp.fetch(N8N_WEBHOOK, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      deadline: 10
    });
  } catch(e) {
    Logger.log('N8N webhook error: ' + e.message);
  }
}

// ─── Guardar nueva cita + notificar vía N8N ───────────────────────
function saveApt(apt) {
  const sheet = getSheet();
  const id    = getNextId(sheet);
  sheet.appendRow([id, apt.code||'', apt.name||'', apt.email||'', apt.tel||'',
    apt.svc||'', Number(apt.price)||0, apt.dur||'', apt.date||'', apt.time||'',
    apt.status||'pending', apt.notes||'', apt.createdAt||new Date().toISOString()]);

  // Notificar a N8N (envía emails + WhatsApp)
  notifyN8N({ action: 'save', apt: apt });

  return {id, code:apt.code};
}

// ─── Actualizar estado + notificar vía N8N ───────────────────────
function updateStatus(code, newStatus) {
  const row = findRowByCode(code);
  if (!row) throw new Error('Cita no encontrada: ' + code);
  const sheet = getSheet();
  sheet.getRange(row, COL.status).setValue(newStatus);
  const apt = rowToObj(sheet.getRange(row, 1, 1, 13).getValues()[0]);
  notifyN8N({ action: 'updateStatus', apt: apt, status: newStatus });
  return {code, status:newStatus};
}

// ─── Cancelar cita + notificar vía N8N ───────────────────────────
function cancelApt(code) {
  const row = findRowByCode(code);
  if (!row) throw new Error('Cita no encontrada: ' + code);
  const sheet = getSheet();
  sheet.getRange(row, COL.status).setValue('cancelled');
  const apt = rowToObj(sheet.getRange(row, 1, 1, 13).getValues()[0]);
  notifyN8N({ action: 'cancel', apt: apt });
  return {code, status:'cancelled'};
}

// ─── Reagendar + notificar vía N8N ───────────────────────────────
function rescheduleApt(code, newDate, newTime) {
  const row = findRowByCode(code);
  if (!row) throw new Error('Cita no encontrada: ' + code);
  const sheet = getSheet();
  sheet.getRange(row, COL.date).setValue(newDate);
  sheet.getRange(row, COL.time).setValue(newTime);
  sheet.getRange(row, COL.status).setValue('rescheduled');
  const apt = rowToObj(sheet.getRange(row, 1, 1, 13).getValues()[0]);
  notifyN8N({ action: 'reschedule', apt: apt, date: newDate, time: newTime });
  return {code, date:newDate, time:newTime, status:'rescheduled'};
}

// ═══════════════════════════════════════════════════════════════════
// ─── EMAILS ────────────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════

function fmtDateEs(dateStr) {
  if (!dateStr) return '';
  const [y,m,d] = dateStr.split('-');
  const meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  return d + ' de ' + meses[parseInt(m)-1] + ' de ' + y;
}

function emailBase(contenido) {
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background:#f4f0eb;font-family:Georgia,serif;">' +
    '<table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f0eb;padding:40px 20px;">' +
    '<tr><td align="center"><table width="520" cellpadding="0" cellspacing="0" style="background:#0a0806;border-radius:4px;overflow:hidden;">' +
    '<tr><td style="background:#1a1410;padding:28px 40px;text-align:center;border-bottom:1px solid rgba(201,169,110,.2);">' +
    '<span style="font-family:Georgia,serif;font-size:28px;color:#c9a96e;letter-spacing:4px;font-style:italic;">' + BUSINESS_NAME + '</span>' +
    '</td></tr>' +
    '<tr><td style="padding:36px 40px;color:#fff;">' + contenido + '</td></tr>' +
    '<tr><td style="background:#1a1410;padding:20px 40px;text-align:center;border-top:1px solid rgba(201,169,110,.2);">' +
    '<p style="margin:0;font-size:11px;color:rgba(255,255,255,.35);letter-spacing:.15em;text-transform:uppercase;">' + BUSINESS_NAME + '</p>' +
    '</td></tr>' +
    '</table></td></tr></table></body></html>';
}

function sendConfirmationToClient(apt) {
  const html = emailBase(
    '<h2 style="margin:0 0 20px;font-size:20px;color:#c9a96e;font-weight:400;letter-spacing:.05em;">Cita agendada ✓</h2>' +
    '<p style="color:rgba(255,255,255,.7);font-size:13px;line-height:1.7;margin:0 0 24px;">Hola <strong style="color:#fff;">' + apt.name + '</strong>, tu cita ha sido registrada exitosamente.</p>' +
    '<table style="width:100%;border-collapse:collapse;margin-bottom:28px;">' +
    '<tr><td style="padding:10px 0;border-bottom:1px solid rgba(255,255,255,.06);color:rgba(255,255,255,.5);font-size:12px;letter-spacing:.1em;text-transform:uppercase;">Servicio</td>' +
    '<td style="padding:10px 0;border-bottom:1px solid rgba(255,255,255,.06);color:#fff;font-size:13px;text-align:right;">' + apt.svc + '</td></tr>' +
    '<tr><td style="padding:10px 0;border-bottom:1px solid rgba(255,255,255,.06);color:rgba(255,255,255,.5);font-size:12px;letter-spacing:.1em;text-transform:uppercase;">Fecha</td>' +
    '<td style="padding:10px 0;border-bottom:1px solid rgba(255,255,255,.06);color:#fff;font-size:13px;text-align:right;">' + fmtDateEs(apt.date) + '</td></tr>' +
    '<tr><td style="padding:10px 0;border-bottom:1px solid rgba(255,255,255,.06);color:rgba(255,255,255,.5);font-size:12px;letter-spacing:.1em;text-transform:uppercase;">Hora</td>' +
    '<td style="padding:10px 0;border-bottom:1px solid rgba(255,255,255,.06);color:#fff;font-size:13px;text-align:right;">' + apt.time + '</td></tr>' +
    '<tr><td style="padding:10px 0;color:rgba(255,255,255,.5);font-size:12px;letter-spacing:.1em;text-transform:uppercase;">Código</td>' +
    '<td style="padding:10px 0;color:#c9a96e;font-size:13px;text-align:right;font-family:monospace;">' + apt.code + '</td></tr>' +
    '</table>' +
    '<p style="color:rgba(255,255,255,.5);font-size:12px;line-height:1.6;margin:0;">Guarda tu código de cita para consultas o cambios. Esperamos verte pronto.</p>'
  );

  MailApp.sendEmail({
    to: apt.email,
    subject: 'Cita confirmada — ' + BUSINESS_NAME + ' · ' + fmtDateEs(apt.date),
    htmlBody: html,
    name: BUSINESS_NAME
  });
}

function sendNotificationToBarbero(apt) {
  const html = emailBase(
    '<h2 style="margin:0 0 20px;font-size:18px;color:#c9a96e;font-weight:400;">Nueva cita agendada</h2>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr><td style="padding:8px 0;border-bottom:1px solid rgba(255,255,255,.06);color:rgba(255,255,255,.5);font-size:12px;text-transform:uppercase;letter-spacing:.1em;">Cliente</td>' +
    '<td style="padding:8px 0;border-bottom:1px solid rgba(255,255,255,.06);color:#fff;font-size:13px;text-align:right;">' + apt.name + '</td></tr>' +
    '<tr><td style="padding:8px 0;border-bottom:1px solid rgba(255,255,255,.06);color:rgba(255,255,255,.5);font-size:12px;text-transform:uppercase;letter-spacing:.1em;">Servicio</td>' +
    '<td style="padding:8px 0;border-bottom:1px solid rgba(255,255,255,.06);color:#fff;font-size:13px;text-align:right;">' + apt.svc + '</td></tr>' +
    '<tr><td style="padding:8px 0;border-bottom:1px solid rgba(255,255,255,.06);color:rgba(255,255,255,.5);font-size:12px;text-transform:uppercase;letter-spacing:.1em;">Fecha</td>' +
    '<td style="padding:8px 0;border-bottom:1px solid rgba(255,255,255,.06);color:#fff;font-size:13px;text-align:right;">' + fmtDateEs(apt.date) + '</td></tr>' +
    '<tr><td style="padding:8px 0;border-bottom:1px solid rgba(255,255,255,.06);color:rgba(255,255,255,.5);font-size:12px;text-transform:uppercase;letter-spacing:.1em;">Hora</td>' +
    '<td style="padding:8px 0;border-bottom:1px solid rgba(255,255,255,.06);color:#fff;font-size:13px;text-align:right;">' + apt.time + '</td></tr>' +
    '<tr><td style="padding:8px 0;color:rgba(255,255,255,.5);font-size:12px;text-transform:uppercase;letter-spacing:.1em;">Tel</td>' +
    '<td style="padding:8px 0;color:#fff;font-size:13px;text-align:right;">' + (apt.tel||'—') + '</td></tr>' +
    '</table>'
  );

  MailApp.sendEmail({
    to: BARBERO_EMAIL,
    subject: 'Nueva cita: ' + apt.name + ' — ' + fmtDateEs(apt.date) + ' ' + apt.time,
    htmlBody: html,
    name: BUSINESS_NAME
  });
}

function sendStatusUpdateToClient(apt, nuevoEstado) {
  const esConfirmada = nuevoEstado === 'confirmada';
  const emoji = esConfirmada ? '✅' : '❌';
  const color = esConfirmada ? '#4caf50' : '#e53935';

  const html = emailBase(
    '<h2 style="margin:0 0 20px;font-size:20px;color:' + color + ';font-weight:400;">' + emoji + ' Cita ' + nuevoEstado + '</h2>' +
    '<p style="color:rgba(255,255,255,.7);font-size:13px;line-height:1.7;margin:0 0 24px;">Hola <strong style="color:#fff;">' + apt.name + '</strong>, tu cita ha sido <strong>' + nuevoEstado + '</strong>.</p>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr><td style="padding:8px 0;border-bottom:1px solid rgba(255,255,255,.06);color:rgba(255,255,255,.5);font-size:12px;text-transform:uppercase;letter-spacing:.1em;">Servicio</td>' +
    '<td style="padding:8px 0;border-bottom:1px solid rgba(255,255,255,.06);color:#fff;font-size:13px;text-align:right;">' + apt.svc + '</td></tr>' +
    '<tr><td style="padding:8px 0;border-bottom:1px solid rgba(255,255,255,.06);color:rgba(255,255,255,.5);font-size:12px;text-transform:uppercase;letter-spacing:.1em;">Fecha</td>' +
    '<td style="padding:8px 0;border-bottom:1px solid rgba(255,255,255,.06);color:#fff;font-size:13px;text-align:right;">' + fmtDateEs(apt.date) + '</td></tr>' +
    '<tr><td style="padding:8px 0;color:rgba(255,255,255,.5);font-size:12px;text-transform:uppercase;letter-spacing:.1em;">Hora</td>' +
    '<td style="padding:8px 0;color:#fff;font-size:13px;text-align:right;">' + apt.time + '</td></tr>' +
    '</table>' +
    (esConfirmada ? '<p style="margin:24px 0 0;color:rgba(255,255,255,.5);font-size:12px;">Te esperamos puntual. Gracias por elegir Barber Style.</p>'
                  : '<p style="margin:24px 0 0;color:rgba(255,255,255,.5);font-size:12px;">Si deseas reagendar, puedes hacerlo en nuestro sitio web.</p>')
  );

  MailApp.sendEmail({
    to: apt.email,
    subject: emoji + ' Tu cita fue ' + nuevoEstado + ' — ' + BUSINESS_NAME,
    htmlBody: html,
    name: BUSINESS_NAME
  });
}

function sendRescheduleToClient(apt, newDate, newTime) {
  const html = emailBase(
    '<h2 style="margin:0 0 20px;font-size:20px;color:#c9a96e;font-weight:400;">📅 Cita reagendada</h2>' +
    '<p style="color:rgba(255,255,255,.7);font-size:13px;line-height:1.7;margin:0 0 24px;">Hola <strong style="color:#fff;">' + apt.name + '</strong>, tu cita ha sido reagendada a la siguiente fecha:</p>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr><td style="padding:8px 0;border-bottom:1px solid rgba(255,255,255,.06);color:rgba(255,255,255,.5);font-size:12px;text-transform:uppercase;letter-spacing:.1em;">Servicio</td>' +
    '<td style="padding:8px 0;border-bottom:1px solid rgba(255,255,255,.06);color:#fff;font-size:13px;text-align:right;">' + apt.svc + '</td></tr>' +
    '<tr><td style="padding:8px 0;border-bottom:1px solid rgba(255,255,255,.06);color:rgba(255,255,255,.5);font-size:12px;text-transform:uppercase;letter-spacing:.1em;">Nueva fecha</td>' +
    '<td style="padding:8px 0;border-bottom:1px solid rgba(255,255,255,.06);color:#c9a96e;font-size:13px;text-align:right;font-weight:bold;">' + fmtDateEs(newDate) + '</td></tr>' +
    '<tr><td style="padding:8px 0;color:rgba(255,255,255,.5);font-size:12px;text-transform:uppercase;letter-spacing:.1em;">Nueva hora</td>' +
    '<td style="padding:8px 0;color:#c9a96e;font-size:13px;text-align:right;font-weight:bold;">' + newTime + '</td></tr>' +
    '</table>'
  );

  MailApp.sendEmail({
    to: apt.email,
    subject: '📅 Cita reagendada — ' + BUSINESS_NAME + ' · ' + fmtDateEs(newDate),
    htmlBody: html,
    name: BUSINESS_NAME
  });
}

// ═══════════════════════════════════════════════════════════════════
// ─── AUTENTICACIÓN ─────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════

// Contraseña maestra hardcodeada (fallback si no hay clave en PropertiesService)
const DEFAULT_PASS = 'bs2025*';
const PASS_KEY     = 'PANEL_PASSWORD';

function getStoredPass() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty(PASS_KEY) || DEFAULT_PASS;
}

function verifyLogin(pass) {
  if (!pass) return { valid: false };
  return { valid: pass === getStoredPass() };
}

function changePassword(currentPass, newPass) {
  if (!currentPass || !newPass) throw new Error('Faltan parámetros');
  if (newPass.length < 6)       throw new Error('La contraseña debe tener al menos 6 caracteres');
  if (currentPass !== getStoredPass()) throw new Error('Contraseña actual incorrecta');

  // Guardar nueva contraseña en PropertiesService (servidor, no visible en código)
  PropertiesService.getScriptProperties().setProperty(PASS_KEY, newPass);

  // Notificar al propietario por correo
  const html = emailBase(
    '<h2 style="margin:0 0 16px;font-size:18px;color:#c9a96e;font-weight:400;">🔐 Contraseña actualizada</h2>' +
    '<p style="color:rgba(255,255,255,.7);font-size:13px;line-height:1.7;margin:0 0 20px;">La contraseña de acceso al panel de <strong style="color:#fff;">' + BUSINESS_NAME + '</strong> fue cambiada exitosamente.</p>' +
    '<p style="color:rgba(255,255,255,.5);font-size:12px;line-height:1.6;margin:0;">Si no realizaste este cambio, actualiza la contraseña de inmediato desde el panel en Configuración.</p>'
  );
  MailApp.sendEmail({
    to: BARBERO_EMAIL,
    subject: '🔐 Contraseña del panel actualizada — ' + BUSINESS_NAME,
    htmlBody: html,
    name: BUSINESS_NAME
  });

  return { changed: true };
}

// ═══════════════════════════════════════════════════════════════════
// ─── LIMPIEZA DE CITAS ──────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════

function clearOldApts(days) {
  if (!days || isNaN(days) || days < 1) throw new Error('Número de días inválido');
  const sheet   = getSheet();
  const last    = sheet.getLastRow();
  if (last < 2)  return { deleted: 0 };

  const tz      = Session.getScriptTimeZone();
  const cutoff  = new Date();
  cutoff.setDate(cutoff.getDate() - Number(days));
  const cutoffStr = Utilities.formatDate(cutoff, tz, 'yyyy-MM-dd');

  // Recorrer de abajo hacia arriba para no desfasar índices al borrar
  let deleted = 0;
  for (let r = last; r >= 2; r--) {
    const dateVal = sheet.getRange(r, COL.date).getValue();
    const dateStr = dateVal instanceof Date
      ? Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd')
      : String(dateVal || '');
    if (dateStr && dateStr < cutoffStr) {
      sheet.deleteRow(r);
      deleted++;
    }
  }
  return { deleted, cutoff: cutoffStr };
}

// ═══════════════════════════════════════════════════════════════════
// ─── HELPERS ───────────────────────────────────────────────────────
// ═══════════════════════════════════════════════════════════════════

function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SHEET_ID');
  if (ssId) {
    try { return SpreadsheetApp.openById(ssId); } catch(e) {}
  }
  const ss = SpreadsheetApp.create(BUSINESS_NAME + ' — Citas');
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
  // Google Sheets puede devolver fecha/hora como objetos Date — hay que formatearlo
  const tz = Session.getScriptTimeZone();
  const dateVal = r[8];
  const timeVal = r[9];
  const dateStr = dateVal instanceof Date
    ? Utilities.formatDate(dateVal, tz, 'yyyy-MM-dd')
    : String(dateVal || '');
  const timeStr = timeVal instanceof Date
    ? Utilities.formatDate(timeVal, tz, 'HH:mm')
    : String(timeVal || '');
  return {id:r[0],code:r[1],name:r[2],email:r[3],tel:r[4],svc:r[5],
    price:Number(r[6])||0,dur:r[7],date:dateStr,time:timeStr,status:r[10],
    notes:r[11],createdAt:r[12]?new Date(r[12]).toISOString():''};
}

