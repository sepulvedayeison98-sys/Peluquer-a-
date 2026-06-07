/**
 * ══════════════════════════════════════════════════════════════
 *  STATUS BARBER STUDIO — Backend Google Apps Script
 *  Maneja: notificaciones Gmail + WhatsApp Business + Google Sheets
 *
 *  CÓMO DESPLEGAR:
 *  1. Ir a script.google.com → Nuevo proyecto
 *  2. Pegar este código
 *  3. Configurar CONSTANTS abajo con tus datos
 *  4. Desplegar → Nueva implementación → Aplicación web
 *     - Ejecutar como: "Yo"
 *     - Acceso: "Cualquier usuario"
 *  5. Copiar la URL y pegarla en barberia-agenda-v3.html y barberia-admin.html
 *     donde dice: const APPS_SCRIPT_URL = 'TU_DEPLOYMENT_URL'
 * ══════════════════════════════════════════════════════════════
 */

// ── CONFIGURACIÓN ──────────────────────────────────────────────
const CONFIG = {
  // Email del barbero (quien recibe notificaciones de clientes)
  BARBER_EMAIL: 'barbero@statusbarberstudio.com',

  // Nombre que aparece en los correos
  BARBER_NAME: 'Status Barber Studio',

  // ID del Google Sheet donde se guardan las citas
  // (Crear un sheet en drive.google.com → copiar el ID del URL)
  SHEET_ID: 'TU_GOOGLE_SHEET_ID',

  // WhatsApp Business API (CallMeBot - gratuito para pruebas)
  // Registro en: https://www.callmebot.com/blog/free-api-whatsapp-messages/
  WA_API_KEY_BARBER: 'TU_API_KEY_CALLMEBOT',
  WA_PHONE_BARBER: '+573001234567', // Con código de país

  // Para WhatsApp al cliente, se usa el mismo método
  // El cliente debe registrar su número primero en CallMeBot
  // En producción: usar Meta WhatsApp Business API
  WA_USE_CALLMEBOT: true,

  // Meta WhatsApp Business API (producción)
  META_WA_TOKEN: 'TU_META_WA_TOKEN',
  META_PHONE_ID: 'TU_META_PHONE_NUMBER_ID',
};

// Nombre de la hoja en Google Sheets
const SHEET_NAME = 'Citas';

// ──────────────────────────────────────────────────────────────

/**
 * Punto de entrada HTTP POST desde las apps HTML
 * Recibe el payload JSON y dispara las notificaciones
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    Logger.log('Payload recibido: ' + JSON.stringify(payload));

    const result = processNotification(payload);
    saveToCitas(payload);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('Error en doPost: ' + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * GET para verificar que el script está activo
 */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, service: 'Status Barber Studio API', version: '1.0' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Procesa el payload y dispara los 4 canales de notificación:
 * - Email al cliente
 * - Email al barbero
 * - WhatsApp al cliente
 * - WhatsApp al barbero
 */
function processNotification(payload) {
  const { trigger, apt, toClient, toBarber, oldDate, oldTime } = payload;
  const msgs = buildMessages(trigger, apt, oldDate, oldTime);
  const results = {};

  // ── Email al cliente ──
  if (msgs.emailClient && toClient.email) {
    try {
      sendEmail(toClient.email, msgs.emailClient.subject, msgs.emailClient.body);
      results.emailClient = 'sent';
    } catch(e) { results.emailClient = 'error: ' + e.toString(); }
  }

  // ── Email al barbero ──
  if (msgs.emailBarber) {
    try {
      sendEmail(CONFIG.BARBER_EMAIL, msgs.emailBarber.subject, msgs.emailBarber.body);
      results.emailBarber = 'sent';
    } catch(e) { results.emailBarber = 'error: ' + e.toString(); }
  }

  // ── WhatsApp al cliente ──
  if (msgs.waClient && toClient.tel) {
    try {
      if (CONFIG.WA_USE_CALLMEBOT) {
        sendWACallMeBot(toClient.tel, msgs.waClient);
      } else {
        sendWAMeta(toClient.tel, msgs.waClient);
      }
      results.waClient = 'sent';
    } catch(e) { results.waClient = 'error: ' + e.toString(); }
  }

  // ── WhatsApp al barbero ──
  if (msgs.waBarber) {
    try {
      if (CONFIG.WA_USE_CALLMEBOT) {
        sendWACallMeBot(CONFIG.WA_PHONE_BARBER, msgs.waBarber);
      } else {
        sendWAMeta(CONFIG.WA_PHONE_BARBER, msgs.waBarber);
      }
      results.waBarber = 'sent';
    } catch(e) { results.waBarber = 'error: ' + e.toString(); }
  }

  Logger.log('Resultados: ' + JSON.stringify(results));
  return results;
}

// ══════════════════════════════════════════════════════════════
//  GMAIL
// ══════════════════════════════════════════════════════════════
function sendEmail(to, subject, body) {
  const htmlBody = body
    .replace(/\n/g, '<br>')
    .replace(/\*(.+?)\*/g, '<strong>$1</strong>');

  GmailApp.sendEmail(to, subject, body, {
    name: CONFIG.BARBER_NAME,
    htmlBody: `
      <div style="font-family:Georgia,serif;max-width:520px;margin:0 auto;color:#1a1714;">
        <div style="background:#1a1714;padding:24px 32px;text-align:center;">
          <div style="font-size:22px;color:#ffffff;letter-spacing:0.08em;">STATUS</div>
          <div style="font-size:12px;color:#b8965a;letter-spacing:0.3em;text-transform:uppercase;margin-top:4px;">Barber Studio</div>
        </div>
        <div style="padding:32px;background:#faf9f7;border:1px solid #e2d9cc;border-top:3px solid #b8965a;">
          <div style="font-size:15px;line-height:1.8;color:#3d352c;">${htmlBody}</div>
        </div>
        <div style="padding:16px 32px;background:#f2ede6;border:1px solid #e2d9cc;border-top:none;font-size:11px;color:#8c7d6e;text-align:center;letter-spacing:0.1em;">
          Status Barber Studio · Carrera 55 24A-06, Medellín · statusbarberstudio.com
        </div>
      </div>`,
    replyTo: CONFIG.BARBER_EMAIL
  });
}

// ══════════════════════════════════════════════════════════════
//  WHATSAPP — CallMeBot (gratuito, para pruebas)
//  Docs: https://www.callmebot.com/blog/free-api-whatsapp-messages/
// ══════════════════════════════════════════════════════════════
function sendWACallMeBot(phone, message) {
  // Limpiar teléfono (solo dígitos y +)
  const cleanPhone = phone.replace(/[\s\-\(\)]/g, '');
  const encodedMsg = encodeURIComponent(message);
  const url = `https://api.callmebot.com/whatsapp.php?phone=${cleanPhone}&text=${encodedMsg}&apikey=${CONFIG.WA_API_KEY_BARBER}`;

  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  Logger.log('CallMeBot response: ' + response.getContentText());

  if (response.getResponseCode() !== 200) {
    throw new Error('CallMeBot error: ' + response.getContentText());
  }
}

// ══════════════════════════════════════════════════════════════
//  WHATSAPP — Meta Business API (producción)
//  Docs: https://developers.facebook.com/docs/whatsapp/cloud-api
// ══════════════════════════════════════════════════════════════
function sendWAMeta(phone, message) {
  const cleanPhone = phone.replace(/[\s\-\(\)\+]/g, '');
  const url = `https://graph.facebook.com/v18.0/${CONFIG.META_PHONE_ID}/messages`;

  const payload = JSON.stringify({
    messaging_product: 'whatsapp',
    to: cleanPhone,
    type: 'text',
    text: { body: message }
  });

  const options = {
    method: 'POST',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + CONFIG.META_WA_TOKEN },
    payload: payload,
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  Logger.log('Meta WA response: ' + response.getContentText());

  if (response.getResponseCode() !== 200) {
    throw new Error('Meta WA error: ' + response.getContentText());
  }
}

// ══════════════════════════════════════════════════════════════
//  GOOGLE SHEETS — Guardar y actualizar citas
// ══════════════════════════════════════════════════════════════
function saveToCitas(payload) {
  const { trigger, apt } = payload;
  if (!CONFIG.SHEET_ID || CONFIG.SHEET_ID === 'TU_GOOGLE_SHEET_ID') {
    Logger.log('Sheet ID no configurado, saltando guardado.');
    return;
  }

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);

    // Crear hoja si no existe
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Código', 'Cliente', 'Teléfono', 'Email', 'Servicio',
        'Precio', 'Duración', 'Fecha', 'Hora', 'Estado',
        'Notas', 'Trigger', 'Última actualización'
      ]);
      sheet.getRange(1, 1, 1, 13).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const now = new Date().toLocaleString('es-CO', { timeZone: 'America/Bogota' });

    // Buscar si ya existe la cita por código
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(apt.code || apt.id)) {
        rowIndex = i + 1; // 1-based
        break;
      }
    }

    const rowData = [
      apt.code || apt.id,
      apt.name || apt.nombre,
      apt.tel,
      apt.email || '',
      apt.svc,
      apt.price,
      apt.dur,
      apt.date,
      apt.time,
      apt.status,
      apt.notes || apt.notas || '',
      trigger,
      now
    ];

    if (rowIndex > 0) {
      // Actualizar fila existente
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      Logger.log('Cita actualizada en fila ' + rowIndex);
    } else {
      // Nueva fila
      sheet.appendRow(rowData);
      Logger.log('Nueva cita guardada en Sheets.');
    }

    // Colorear fila según estado
    const lastRow = rowIndex > 0 ? rowIndex : sheet.getLastRow();
    const statusColors = {
      confirmed:   '#edf4f0',
      pending:     '#fdf3e3',
      cancelled:   '#fdf0f0',
      rescheduled: '#f0f4fd',
      done:        '#f2ede6'
    };
    const color = statusColors[apt.status] || '#ffffff';
    sheet.getRange(lastRow, 1, 1, 13).setBackground(color);

  } catch(e) {
    Logger.log('Error guardando en Sheets: ' + e.toString());
  }
}

// ══════════════════════════════════════════════════════════════
//  PLANTILLAS DE MENSAJES
// ══════════════════════════════════════════════════════════════
function buildMessages(trigger, apt, oldDate, oldTime) {
  const name    = apt.name || apt.nombre || 'Cliente';
  const fd      = formatDate(apt.date);
  const oldFd   = oldDate ? formatDate(oldDate) : '';

  const templates = {

    // ─ Cliente cancela ─
    cancel_by_client: {
      emailClient: {
        subject: `Cancelación confirmada — ${apt.code}`,
        body:
          `Hola ${name},\n\n` +
          `Tu cita ha sido cancelada exitosamente.\n\n` +
          `*Código:* ${apt.code}\n` +
          `*Servicio:* ${apt.svc}\n` +
          `*Fecha cancelada:* ${fd} a las ${apt.time}\n\n` +
          `Si deseas reagendar, visita nuestro sitio.\n\n` +
          `— Status Barber Studio`
      },
      emailBarber: {
        subject: `⚠ Cancelación de cita — ${apt.code}`,
        body:
          `*Cliente canceló su cita*\n\n` +
          `*Cliente:* ${name}\n` +
          `*Tel:* ${apt.tel}\n` +
          `*Servicio:* ${apt.svc}\n` +
          `*Fecha:* ${fd} ${apt.time}\n` +
          `*Código:* ${apt.code}`
      },
      waClient:
        `✂️ *Status Barber Studio*\n\n` +
        `Hola ${name}, tu cita *${apt.code}* ha sido cancelada.\n\n` +
        `📅 ${fd} · ${apt.time}\n` +
        `💈 ${apt.svc}\n\n` +
        `Puedes reagendar en nuestro sitio web cuando quieras.`,
      waBarber:
        `⚠️ *Cancelación de cita*\n\n` +
        `Cliente: *${name}*\n` +
        `Código: ${apt.code}\n` +
        `Servicio: ${apt.svc}\n` +
        `Fecha: ${fd} ${apt.time}\n\n` +
        `El cliente canceló su cita.`
    },

    // ─ Barbero cancela ─
    cancel_by_barber: {
      emailClient: {
        subject: `Tu cita fue cancelada — Status Barber Studio`,
        body:
          `Hola ${name},\n\n` +
          `Lamentamos informarte que tu cita fue cancelada por el barbero.\n\n` +
          `*Servicio:* ${apt.svc}\n` +
          `*Fecha:* ${fd} ${apt.time}\n\n` +
          `Por favor escríbenos para reagendar. Disculpa los inconvenientes.\n\n` +
          `— Status Barber Studio`
      },
      emailBarber: {
        subject: `Cita cancelada — ${name}`,
        body: `Cancelaste la cita de ${name}\nFecha: ${fd} ${apt.time}\nServicio: ${apt.svc}`
      },
      waClient:
        `⚠️ *Status Barber Studio*\n\n` +
        `Hola ${name}, lamentamos que tu cita fue cancelada por el barbero.\n\n` +
        `📅 ${fd} · ${apt.time}\n` +
        `💈 ${apt.svc}\n\n` +
        `Por favor escríbenos para reagendar. Disculpa el inconveniente 🙏`,
      waBarber:
        `✅ Cita de *${name}* cancelada.\nSe notificó al cliente.`
    },

    // ─ Cliente solicita reagendar ─
    reschedule_by_client: {
      emailClient: {
        subject: `Solicitud de reagendamiento recibida — ${apt.code}`,
        body:
          `Hola ${name},\n\n` +
          `Tu solicitud de reagendamiento fue recibida.\n\n` +
          `*Fecha anterior:* ${oldFd} ${oldTime}\n` +
          `*Nueva fecha solicitada:* ${fd} ${apt.time}\n` +
          `*Servicio:* ${apt.svc}\n\n` +
          `El barbero confirmará tu nuevo horario pronto.\n\n` +
          `— Status Barber Studio`
      },
      emailBarber: {
        subject: `📅 Solicitud de reagendamiento — ${name}`,
        body:
          `*${name}* solicita reagendar su cita.\n\n` +
          `*Tel:* ${apt.tel}\n` +
          `*Servicio:* ${apt.svc}\n` +
          `*Fecha anterior:* ${oldFd} ${oldTime}\n` +
          `*Nueva fecha:* ${fd} ${apt.time}\n\n` +
          `Ingresa al panel para confirmar.`
      },
      waClient:
        `📅 *Status Barber Studio*\n\n` +
        `Hola ${name}, tu solicitud de reagendamiento fue recibida.\n\n` +
        `🔄 De: ${oldFd} ${oldTime}\n` +
        `✅ A: *${fd} ${apt.time}*\n` +
        `💈 ${apt.svc}\n\n` +
        `El barbero confirmará pronto. ✂️`,
      waBarber:
        `📅 *Reagendamiento solicitado*\n\n` +
        `Cliente: *${name}*\n` +
        `De: ${oldFd} ${oldTime}\n` +
        `A: *${fd} ${apt.time}*\n` +
        `Servicio: ${apt.svc}\n\n` +
        `Confirma en el panel de administración.`
    },

    // ─ Barbero reagenda ─
    reschedule_by_barber: {
      emailClient: {
        subject: `Tu cita fue reagendada — Status Barber Studio`,
        body:
          `Hola ${name},\n\n` +
          `El barbero reagendó tu cita.\n\n` +
          `*Fecha anterior:* ${oldFd} ${oldTime}\n` +
          `*Nueva fecha:* ${fd} ${apt.time}\n` +
          `*Servicio:* ${apt.svc}\n\n` +
          `Si no puedes en la nueva fecha, por favor escríbenos.\n\n` +
          `— Status Barber Studio`
      },
      emailBarber: {
        subject: `Cita reagendada — ${name}`,
        body: `Reagendaste la cita de ${name}\nDe: ${oldFd} ${oldTime}\nA: ${fd} ${apt.time}`
      },
      waClient:
        `📅 *Status Barber Studio*\n\n` +
        `Hola ${name}, el barbero reagendó tu cita.\n\n` +
        `🔄 De: ${oldFd} ${oldTime}\n` +
        `✅ A: *${fd} ${apt.time}*\n` +
        `💈 ${apt.svc}\n\n` +
        `Si tienes alguna duda escríbenos. ✂️`,
      waBarber:
        `✅ Reagendamiento de *${name}* confirmado.\nNueva fecha: ${fd} ${apt.time}\nCliente notificado.`
    },

    // ─ Barbero confirma cita ─
    confirm_by_barber: {
      emailClient: {
        subject: `¡Cita confirmada! — Status Barber Studio`,
        body:
          `Hola ${name},\n\n` +
          `Tu cita está confirmada. ¡Te esperamos!\n\n` +
          `*Servicio:* ${apt.svc}\n` +
          `*Fecha:* ${fd} ${apt.time}\n` +
          `*Duración:* ${apt.dur}\n` +
          `*Dirección:* Carrera 55 24A-06, Medellín, Bello\n\n` +
          `Recuerda llegar 5 minutos antes. ✂️\n\n` +
          `— Status Barber Studio`
      },
      emailBarber: {
        subject: `Cita confirmada — ${name}`,
        body: `Confirmaste la cita de ${name}\nFecha: ${fd} ${apt.time}\nServicio: ${apt.svc}`
      },
      waClient:
        `✅ *Status Barber Studio*\n\n` +
        `Hola ${name}, tu cita está *confirmada*. ¡Te esperamos!\n\n` +
        `📅 ${fd} · ${apt.time}\n` +
        `💈 ${apt.svc}\n` +
        `📍 Carrera 55 24A-06, Medellín\n\n` +
        `Recuerda llegar 5 minutos antes. ✂️`,
      waBarber:
        `✅ Cita de *${name}* confirmada.\n${fd} ${apt.time} — ${apt.svc}`
    },

    // ─ Nueva reserva del cliente ─
    new_booking: {
      emailClient: {
        subject: `Reserva recibida — ${apt.code}`,
        body:
          `Hola ${name},\n\n` +
          `Tu reserva fue registrada exitosamente.\n\n` +
          `*Código:* ${apt.code}\n` +
          `*Servicio:* ${apt.svc}\n` +
          `*Fecha:* ${fd} ${apt.time}\n` +
          `*Duración:* ${apt.dur}\n` +
          `*Total:* $${Number(apt.price).toLocaleString('es-CO')}\n\n` +
          `El barbero confirmará tu cita pronto.\n\n` +
          `Guarda tu código para gestionar tu cita.\n\n` +
          `— Status Barber Studio`
      },
      emailBarber: {
        subject: `Nueva reserva — ${name}`,
        body:
          `*Nueva cita reservada*\n\n` +
          `*Cliente:* ${name}\n` +
          `*Tel:* ${apt.tel}\n` +
          `*Email:* ${apt.email||'—'}\n` +
          `*Servicio:* ${apt.svc}\n` +
          `*Fecha:* ${fd} ${apt.time}\n` +
          `*Código:* ${apt.code}`
      },
      waClient:
        `✂️ *Status Barber Studio*\n\n` +
        `Hola ${name}, tu reserva fue recibida.\n\n` +
        `🔖 Código: *${apt.code}*\n` +
        `📅 ${fd} · ${apt.time}\n` +
        `💈 ${apt.svc}\n\n` +
        `El barbero confirmará pronto. ¡Nos vemos! 💈`,
      waBarber:
        `🆕 *Nueva reserva*\n\n` +
        `Cliente: *${name}*\n` +
        `Tel: ${apt.tel}\n` +
        `Servicio: ${apt.svc}\n` +
        `Fecha: ${fd} ${apt.time}\n` +
        `Código: ${apt.code}`
    }
  };

  return templates[trigger] || {};
}

// ══════════════════════════════════════════════════════════════
//  HELPERS
// ══════════════════════════════════════════════════════════════
function formatDate(ds) {
  if (!ds) return '—';
  const months = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  const [y, m, d] = ds.split('-');
  return `${parseInt(d)} de ${months[parseInt(m)-1]} ${y}`;
}

/**
 * Trigger programado: recordatorio 24h antes de cada cita
 * Configurar en Apps Script → Triggers → Cada hora
 */
function sendReminders() {
  if (!CONFIG.SHEET_ID || CONFIG.SHEET_ID === 'TU_GOOGLE_SHEET_ID') return;

  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, 'America/Bogota', 'yyyy-MM-dd');

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const [code, name, tel, email, svc, price, dur, date, time, status] = data[i];
    if (date === tomorrowStr && status === 'confirmed') {
      const apt = { code, name, tel, email, svc, price, dur, date, time, status };
      const fd = formatDate(date);

      const reminderMsg =
        `⏰ *Recordatorio — Status Barber Studio*\n\n` +
        `Hola ${name}, mañana tienes cita.\n\n` +
        `📅 ${fd} · ${time}\n` +
        `💈 ${svc}\n` +
        `📍 Carrera 55 24A-06, Medellín\n\n` +
        `¡Te esperamos puntual! ✂️`;

      try {
        if (email) {
          sendEmail(email, 'Recordatorio de cita mañana — Status Barber Studio',
            `Hola ${name},\n\nTe recordamos que mañana tienes cita.\n\nFecha: ${fd} ${time}\nServicio: ${svc}\nDirección: Carrera 55 24A-06, Medellín\n\n¡Te esperamos! ✂️\n\n— Status Barber Studio`
          );
        }
        if (tel && CONFIG.WA_USE_CALLMEBOT) {
          sendWACallMeBot(tel, reminderMsg);
        }
        Logger.log('Recordatorio enviado a ' + name);
      } catch(e) {
        Logger.log('Error recordatorio ' + name + ': ' + e.toString());
      }
    }
  }
}

/**
 * TEST — ejecutar manualmente para probar envíos
 */
function testNotification() {
  const testPayload = {
    trigger: 'new_booking',
    apt: {
      code: 'SBS-TEST01',
      name: 'Cliente Test',
      tel: CONFIG.WA_PHONE_BARBER,
      email: CONFIG.BARBER_EMAIL,
      svc: 'Corte + Barba',
      price: 42000,
      dur: '60 min',
      date: '2026-06-10',
      time: '10:00',
      status: 'pending',
      notes: 'Prueba del sistema'
    },
    toClient: { email: CONFIG.BARBER_EMAIL, tel: CONFIG.WA_PHONE_BARBER },
    toBarber: { email: CONFIG.BARBER_EMAIL, tel: CONFIG.WA_PHONE_BARBER }
  };

  const result = processNotification(testPayload);
  Logger.log('Test result: ' + JSON.stringify(result));
}
