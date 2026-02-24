function doGet(e) {
  let vista = e.parameter.vista || 'inicio';
  let html;
  
  try {
    if (vista === 'medico') {
      html = HtmlService.createTemplateFromFile('VistaMedico');
      html.token = e.parameter.token;
    } else if (vista === 'aprobacion') {
      html = HtmlService.createTemplateFromFile('VistaAprobacion');
      html.token = e.parameter.token;
    } else {
      html = HtmlService.createTemplateFromFile('Index');
    }
    return html.evaluate().setTitle('Sistema Corporativo PTD').addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } catch (error) {
    return HtmlService.createHtmlOutput(`<h2>Error 404</h2><p>La vista solicitada no existe o el archivo HTML falta en el proyecto.</p>`);
  }
}