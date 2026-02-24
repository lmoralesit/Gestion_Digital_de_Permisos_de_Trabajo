function enviarCorreoNotificacion(destino, idPT, titulo, mensaje, link, textoBoton) {
  const htmlBody = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
      <div style="background-color: #003366; padding: 20px; text-align: center; color: white;">
        <h2 style="margin: 0;">Sistema de Permisos de Trabajo (PTD)</h2>
      </div>
      <div style="padding: 30px; background-color: #ffffff; color: #333333;">
        <h3 style="color: #00509e; border-bottom: 2px solid #f0f0f0; padding-bottom: 10px;">${titulo}: ${idPT}</h3>
        <p style="font-size: 16px; line-height: 1.5;">${mensaje}</p>
        <div style="text-align: center; margin-top: 30px; margin-bottom: 20px;">
          <a href="${link}" style="background-color: #28a745; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 16px; display: inline-block;">${textoBoton}</a>
        </div>
        <p style="font-size: 12px; color: #777777; text-align: center;">Este es un mensaje automático. Por favor no responda a este correo.</p>
      </div>
    </div>
  `;
  MailApp.sendEmail({ to: destino, subject: `[PTD Notificación] - ${idPT}`, htmlBody: htmlBody });
}