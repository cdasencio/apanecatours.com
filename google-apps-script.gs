const RESERVATION_EMAIL = "ratatouillextremetours@gmail.com";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || "{}");
    const subject = data.subject || "Orden de reserva - Ratatouille Xtreme Tours";
    const message = data.message || "Orden recibida sin detalle.";

    MailApp.sendEmail({
      to: RESERVATION_EMAIL,
      subject: subject,
      body: message,
      name: "Ratatouille Xtreme Tours"
    });

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    MailApp.sendEmail({
      to: RESERVATION_EMAIL,
      subject: "Error recibiendo orden - Ratatouille Xtreme Tours",
      body: String(error)
    });

    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(error) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
