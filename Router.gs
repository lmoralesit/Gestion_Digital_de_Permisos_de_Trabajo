function doGet(e) {
  let vista = e.parameter.vista || 'inicio';
  let html;
  
  try {
    switch(vista) {
      case 'medico':
        html = HtmlService.createTemplateFromFile('VistaMedico');
        html.token = e.parameter.token;
        break;
      case 'aprobacion':
        html = HtmlService.createTemplateFromFile('VistaAprobacion');
        html.token = e.parameter.token;
        break;
      case 'auditoria':
        html = HtmlService.createTemplateFromFile('VistaAuditoria');
        html.token = e.parameter.token;
        break;
      case 'cierre':
        html = HtmlService.createTemplateFromFile('VistaCierre');
        html.token = e.parameter.token;
        break;
      default:
        html = HtmlService.createTemplateFromFile('Index');
    }
    return html.evaluate()
      .setTitle('Sistema Corporativo PTD')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput('<h2>Error</h2><p>' + error.message + '</p>');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}