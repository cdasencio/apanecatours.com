const DESTINATION_EMAIL = "ratatouillextremetours@gmail.com";
const COMMENTS_SHEET_NAME = "Comentarios";

function doPost(e) {
  try {
    const data = getCommentData_(e);

    if (!data.comentario) {
      throw new Error("Comentario vacio.");
    }

    saveComment_(data);
    sendCommentEmail_(data);

    return json_({
      ok: true,
      message: "Comentario recibido correctamente",
      receivedAt: data.fecha
    });
  } catch (error) {
    const message = String(error && error.message ? error.message : error);
    notifyError_(message, e);
    return json_({
      ok: false,
      error: message
    });
  }
}

function doGet() {
  return json_({
    ok: true,
    message: "Script de comentarios Ratatouille activo",
    email: DESTINATION_EMAIL,
    time: new Date().toISOString()
  });
}

function getCommentData_(e) {
  const params = e && e.parameter ? e.parameter : {};
  let payload = {};

  if (params.payload) {
    try {
      payload = JSON.parse(params.payload);
    } catch (error) {
      payload = {};
    }
  }

  const now = new Date();
  const comentario = clean_(params.comentario || params.comment || params.mensaje || payload.comment || payload.message);

  return {
    fecha: Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"),
    comentario: comentario,
    pagina: clean_(params.pagina || params.page || payload.sourceUrl || payload.page || ""),
    idioma: clean_(params.idioma || params.lang || payload.pageLanguage || ""),
    dispositivo: clean_(params.dispositivo || payload.deviceType || ""),
    navegador: clean_(params.navegador || params.userAgent || payload.userAgent || ""),
    origen: clean_(params.origen || "Formulario de comentarios - Ratatouille Xtreme Tours")
  };
}

function saveComment_(data) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) return;

  let sheet = spreadsheet.getSheetByName(COMMENTS_SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(COMMENTS_SHEET_NAME);

  const headers = ["fecha", "comentario", "pagina", "idioma", "dispositivo", "navegador", "origen"];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }

  sheet.appendRow(headers.map(function(header) {
    return data[header] || "";
  }));
}

function sendCommentEmail_(data) {
  const subject = "Nuevo comentario - Ratatouille Xtreme Tours";
  const body = [
    "NUEVO COMENTARIO DESDE LA PAGINA WEB",
    "",
    "Comentario:",
    data.comentario,
    "",
    "Pagina: " + data.pagina,
    "Idioma: " + data.idioma,
    "Dispositivo: " + data.dispositivo,
    "Navegador: " + data.navegador,
    "Fecha: " + data.fecha
  ].join("\n");

  const htmlBody = [
    '<div style="font-family:Arial,sans-serif;color:#111827;background:#f8fafc;padding:24px;">',
    '<div style="max-width:640px;margin:auto;background:#ffffff;border:1px solid #e5e7eb;border-radius:16px;overflow:hidden;">',
    '<div style="background:#071b2e;color:#ffffff;padding:20px 24px;">',
    '<div style="font-size:12px;font-weight:900;color:#ffd000;text-transform:uppercase;letter-spacing:.8px;">Ratatouille Xtreme Tours Apaneca</div>',
    '<h1 style="font-size:24px;line-height:1.15;margin:8px 0 0;">Nuevo comentario</h1>',
    '</div>',
    '<div style="padding:22px 24px;line-height:1.65;">',
    '<p style="margin:0 0 16px;"><strong>Comentario:</strong><br>' + escapeHtml_(data.comentario).replace(/\n/g, "<br>") + '</p>',
    '<p style="margin:0;color:#475569;font-size:13px;"><strong>Pagina:</strong> ' + escapeHtml_(data.pagina) + '</p>',
    '<p style="margin:0;color:#475569;font-size:13px;"><strong>Idioma:</strong> ' + escapeHtml_(data.idioma) + '</p>',
    '<p style="margin:0;color:#475569;font-size:13px;"><strong>Dispositivo:</strong> ' + escapeHtml_(data.dispositivo) + '</p>',
    '<p style="margin:0;color:#475569;font-size:13px;"><strong>Fecha:</strong> ' + escapeHtml_(data.fecha) + '</p>',
    '</div>',
    '</div>',
    '</div>'
  ].join("");

  MailApp.sendEmail({
    to: DESTINATION_EMAIL,
    subject: subject,
    body: body,
    htmlBody: htmlBody
  });
}

function notifyError_(message, e) {
  try {
    MailApp.sendEmail({
      to: DESTINATION_EMAIL,
      subject: "Error en comentarios - Ratatouille Xtreme Tours",
      body: "Error: " + message + "\n\nParametros recibidos:\n" + JSON.stringify(e && e.parameter ? e.parameter : {}, null, 2)
    });
  } catch (error) {
    // No bloquear la respuesta si falla la notificacion de error.
  }
}

function clean_(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function escapeHtml_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
