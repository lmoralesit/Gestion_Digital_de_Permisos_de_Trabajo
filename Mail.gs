function enviarCorreoNotificacion(destino, idPT, titulo, mensaje, link, textoBoton) {
  const htmlBody = `
    <div style="font-family:'Inter','Segoe UI',system-ui,sans-serif;max-width:600px;margin:0 auto;background:#f1f5f9;">
      <!-- Header -->
      <div style="background:linear-gradient(135deg,#0c1929,#1a365d);padding:28px 32px;text-align:center;">
        <h2 style="margin:0;color:#ffffff;font-size:18px;font-weight:700;letter-spacing:.3px;">
          ⛑️ Sistema de Permisos de Trabajo
        </h2>
        <p style="margin:6px 0 0;color:#94a3b8;font-size:12px;font-weight:400;">Plataforma Digital PTD — FO-COR-SSA-015</p>
      </div>
      <!-- Body -->
      <div style="background:#ffffff;padding:36px 32px;border-left:1px solid #e2e8f0;border-right:1px solid #e2e8f0;">
        <div style="background:#f0f9ff;border-left:4px solid #0d9488;border-radius:0 6px 6px 0;padding:14px 18px;margin-bottom:24px;">
          <p style="margin:0;font-size:13px;color:#475569;font-weight:600;">PERMISO Nº</p>
          <p style="margin:4px 0 0;font-size:22px;color:#0c1929;font-weight:800;">${idPT}</p>
        </div>
        <h3 style="color:#0c1929;font-size:16px;font-weight:700;margin:0 0 12px;">${titulo}</h3>
        <p style="font-size:14px;line-height:1.7;color:#475569;margin:0 0 28px;">${mensaje}</p>
        <div style="text-align:center;margin:32px 0;">
          <a href="${link}" style="background:linear-gradient(135deg,#0d9488,#14b8a6);color:#ffffff;padding:14px 36px;text-decoration:none;border-radius:8px;font-weight:700;font-size:15px;display:inline-block;letter-spacing:.3px;">${textoBoton}</a>
        </div>
      </div>
      <!-- Footer -->
      <div style="background:#f8fafc;padding:20px 32px;text-align:center;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 8px 8px;">
        <p style="margin:0;font-size:11px;color:#94a3b8;line-height:1.6;">
          Mensaje automático del Sistema PTD · No responda a este correo<br>
          © ${new Date().getFullYear()} · Gestión SSA
        </p>
      </div>
    </div>
  `;
  MailApp.sendEmail({ to: destino, subject: '[PTD] ' + titulo + ' — ' + idPT, htmlBody: htmlBody });
}