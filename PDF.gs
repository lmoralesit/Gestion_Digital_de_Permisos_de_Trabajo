/**
 * Genera el PDF del Permiso de Trabajo usando la plantilla de Google Docs.
 * Mapea TODAS las etiquetas de la plantilla FO-COR-SSA-015 Rev.5
 */
function generarPDF_PT(idPT, filaMaestro) {
  try {
    const ss = conectarDB();
    const hM = ss.getSheetByName('PT_Maestro');
    const datos = hM.getRange(filaMaestro, 1, 1, 42).getValues()[0];

    // Leer trabajadores
    const dataTrab = ss.getSheetByName('PT_Trabajadores').getDataRange().getValues();
    let trabajadores = [];
    for (let i = 1; i < dataTrab.length; i++) {
      if (dataTrab[i][1] === idPT) {
        trabajadores.push({
          nombre: dataTrab[i][2] || '', cedula: dataTrab[i][3] || '', cargo: dataTrab[i][4] || '',
          urlFirma: dataTrab[i][5] || '', // Column 6 - Firma_Trabajador URL
          tension: dataTrab[i][6] || '', fc: dataTrab[i][7] || '',
          aptitud: dataTrab[i][8] || '', obsMedica: dataTrab[i][9] || '',
          validadoPor: dataTrab[i][10] || ''
        });
      }
    }

    // Leer riesgos
    const dataRiesgos = ss.getSheetByName('PT_Riesgos').getDataRange().getValues();
    let riesgos = [];
    for (let i = 1; i < dataRiesgos.length; i++) {
      if (dataRiesgos[i][1] === idPT) {
        riesgos.push({
          peligro: dataRiesgos[i][2] || '', 
          observacion: dataRiesgos[i][3] || '',
          severidad: dataRiesgos[i][4] || '',
          probabilidad: dataRiesgos[i][5] || '', 
          inherente: dataRiesgos[i][6] || '',
          jerarquia: dataRiesgos[i][7] || '', 
          controles: dataRiesgos[i][8] || '',
          residual: dataRiesgos[i][9] || ''
        });
      }
    }

    // Leer observaciones de auditoría
    const dataAud = ss.getSheetByName('PT_Auditoria').getDataRange().getValues();
    let obsAuditoria = '';
    for (let i = 1; i < dataAud.length; i++) {
      if (dataAud[i][1] === idPT && dataAud[i][10]) obsAuditoria += dataAud[i][10] + ' ';
    }

    // Crear copia temporal del template
    const tempFile = DriveApp.getFileById(CONFIG.DOC_TEMPLATE_ID).makeCopy('TEMP_' + idPT);
    const doc = DocumentApp.openById(tempFile.getId());
    const body = doc.getBody();
    const fmt = 'dd/MM/yyyy HH:mm';
    const tz = 'America/Caracas';

    // ===== TÍTULO =====
    reemplazar(body, '00001', idPT);

    // ===== INFORMACIÓN GENERAL =====
    reemplazar(body, '<<Unidad de Negocio>>', datos[7]);
    reemplazar(body, '<<Localidad>>', datos[8]);
    reemplazar(body, '<<Sede>>', datos[9]);
    reemplazar(body, '<<Fecha/Hora Inicio>>', datos[31] ? Utilities.formatDate(new Date(datos[31]), tz, fmt) : 'Pendiente');
    reemplazar(body, '<<Nombre Empresa>>', datos[4]);
    reemplazar(body, '<<Área / Equipo>>', datos[10]);

    // Estructura organizativa (replaceText reemplaza TODAS las ocurrencias - correcto)
    reemplazar(body, '<<Solicitante del Servicio>>', datos[12]);
    reemplazar(body, '<<Proceso Solicitante del Servicio>>', datos[14]);
    reemplazar(body, '<<Dpto. Solicitante del Servicio>>', datos[13]);
    reemplazar(body, '<<Dueño del Área>>', datos[15]);
    reemplazar(body, '<<Proceso Dueño del Área>>', datos[17]);
    reemplazar(body, '<<Dpto. Dueño del Área>>', datos[16]);
    reemplazar(body, '<<Nombre Analista>>', datos[5]);
    reemplazar(body, '<<Cédula Analista>>', datos[6]);

    // Firma del analista SST
    reemplazar(body, '<<Firma de Analista>>', datos[27] ? '✓ Firmado digitalmente' : '—');

    // Personal Ejecutante 1 y 2
    let t1 = trabajadores[0] || {};
    let t2 = trabajadores[1] || {};
    reemplazar(body, '<<Personal Ejecutante 1>>', t1.nombre || '—');
    reemplazar(body, '<<Cédula Personal Ejecutante 1>>', t1.cedula || '—');
    reemplazar(body, '<<Cargo Personal Ejecutante 1>>', t1.cargo || '—');
    reemplazar(body, '<<Firma de Personal Ejecutante 1>>', t1.urlFirma ? 'Ver firma: ' + t1.urlFirma : '—');
    reemplazar(body, '<<Personal Ejecutante 2>>', t2.nombre || '—');
    reemplazar(body, '<<Cédula Personal Ejecutante 2>>', t2.cedula || '—');
    reemplazar(body, '<<Cargo Personal Ejecutante 2>>', t2.cargo || '—');
    reemplazar(body, '<<Firma de Personal Ejecutante 2>>', t2.urlFirma ? 'Ver firma: ' + t2.urlFirma : '—');

    // Riesgos (slots 1 y 2)
    let r1 = riesgos[0] || {};
    let r2 = riesgos[1] || {};
    reemplazar(body, '<<Peligro Identificado 1>>', (r1.peligro + ' - ' + r1.observacion).trim() || '—');
    reemplazar(body, '<<Severidad 1>>', r1.severidad || '—');
    reemplazar(body, '<<Probabilidad 1>>', r1.probabilidad || '—');
    reemplazar(body, '<<Riesgo Inherente 1>>', r1.inherente ? r1.inherente.toString() : '—');
    reemplazar(body, '<<Jerarquía de Control 1>>', r1.jerarquia || '—');
    reemplazar(body, '<<Controles Aplicados 1>>', r1.controles || '—');
    reemplazar(body, '<<Riesgo Residual 1>>', r1.residual ? r1.residual.toString() : '—');
    reemplazar(body, '<<Peligro Identificado 2>>', (r2.peligro + ' - ' + r2.observacion).trim() || '—');
    reemplazar(body, '<<Severidad 2>>', r2.severidad || '—');
    reemplazar(body, '<<Probabilidad 2>>', r2.probabilidad || '—');
    reemplazar(body, '<<Riesgo Inherente 2>>', r2.inherente ? r2.inherente.toString() : '—');
    reemplazar(body, '<<Jerarquía de Control 2>>', r2.jerarquia || '—');
    reemplazar(body, '<<Controles Aplicados 2>>', r2.controles || '—');
    reemplazar(body, '<<Riesgo Residual 2>>', r2.residual ? r2.residual.toString() : '—');

    // Medición de atmósfera
    let gases = (datos[23] || datos[24] || datos[25] || datos[26])
      ? 'O₂: ' + (datos[23]||'N/A') + '% | LEL: ' + (datos[24]||'N/A') + '% | CO: ' + (datos[25]||'N/A') + ' ppm | H₂S: ' + (datos[26]||'N/A') + ' ppm'
      : 'No aplica';
    reemplazar(body, '<<Medición de Atmósfera>>', gases);

    // ===== VALIDACIÓN MÉDICA (slots 1 y 2) =====
    reemplazar(body, '<<Trabajador a Evaluar 1>>', t1.nombre || '—');
    reemplazar(body, '<<Tensión Arterial Trabajador 1>>', t1.tension || '—');
    reemplazar(body, '<<Frecuencia Cardiaca Trabajador 1>>', t1.fc ? t1.fc.toString() : '—');
    reemplazar(body, '<<Aptitud Médica Trabajador 1>>', t1.aptitud || '—');
    reemplazar(body, '<<Observaciones Médicas Trabajador 1>>', t1.obsMedica || '—');
    reemplazar(body, '<<Trabajador a Evaluar 2>>', t2.nombre || '—');
    reemplazar(body, '<<Tensión Arterial Trabajador 2>>', t2.tension || '—');
    reemplazar(body, '<<Frecuencia Cardiaca Trabajador 2>>', t2.fc ? t2.fc.toString() : '—');
    reemplazar(body, '<<Aptitud Médica Trabajador 2>>', t2.aptitud || '—');
    reemplazar(body, '<<Observaciones Médicas Trabajador 2>>', t2.obsMedica || '—');
    reemplazar(body, '<<Validado por>>', t1.validadoPor || '—');

    // ===== APROBACIÓN =====
    // <<Solicitante del Servicio>> y <<Dueño del Área>> ya fueron reemplazados arriba
    reemplazar(body, '<<Firma de Solicitante del Servicio>>', datos[28] ? '✓ Firmado digitalmente' : '—');
    let fechaAprobacion = datos[31] ? Utilities.formatDate(new Date(datos[31]), tz, fmt) : 'Pendiente';
    reemplazar(body, '<<Fecha/Hora>>', fechaAprobacion);
    reemplazar(body, '<<Firma  Dueño del Área>>', datos[29] ? '✓ Firmado digitalmente' : '—');

    // <<Firma SSA>> aparece en Aprobación Y Cierre - manejar con findText secuencial
    reemplazarPrimero(body, '<<Firma SSA>>', datos[30] ? '✓ Firmado digitalmente' : '—');

    // ===== CIERRE =====
    reemplazar(body, '<<Estatus del Permiso de Trabajo>>', datos[33] || 'Pendiente');
    reemplazar(body, '<<Firma de Analista de Seguridad de la empresa ejecutante>>', datos[34] ? '✓ Firmado' : '—');
    reemplazar(body, '<<Firma del Solicitante del Servicio>>', datos[35] ? '✓ Firmado' : '—');
    // Segunda ocurrencia de Firma SSA (cierre)
    reemplazarPrimero(body, '<<Firma SSA>>', datos[36] ? '✓ Firmado' : '—');
    reemplazar(body, '<<Fecha/Hora del Cierre>>', datos[37] ? Utilities.formatDate(new Date(datos[37]), tz, fmt) : 'Pendiente');
    reemplazar(body, '<<Observaciones>>', obsAuditoria.trim() || '—');

    doc.saveAndClose();

    // Generar PDF
    const pdfBlob = tempFile.getAs(MimeType.PDF);
    pdfBlob.setName(idPT + '.pdf');

    // Organizar en Drive: Año > Mes > Contratista
    const fecha = new Date();
    const carpetaRaiz = DriveApp.getFolderById(CONFIG.FOLDER_RAIZ_ID);
    const carpetaAnio = buscarOCrearCarpeta(carpetaRaiz, fecha.getFullYear().toString());
    const carpetaMes = buscarOCrearCarpeta(carpetaAnio, ('0' + (fecha.getMonth() + 1)).slice(-2));
    const carpetaContratista = buscarOCrearCarpeta(carpetaMes, datos[4] || 'SinEmpresa');

    // Si ya existe un PDF anterior, eliminarlo
    let urlAnterior = datos[32];
    if (urlAnterior) {
      try {
        let idAnterior = urlAnterior.match(/[-\w]{25,}/);
        if (idAnterior) DriveApp.getFileById(idAnterior[0]).setTrashed(true);
      } catch(ex) { /* PDF anterior no encontrado, ignorar */ }
    }

    const archivoPDF = carpetaContratista.createFile(pdfBlob);
    hM.getRange(filaMaestro, 33).setValue(archivoPDF.getUrl());
    tempFile.setTrashed(true);

    return archivoPDF.getUrl();
  } catch (e) {
    Logger.log('Error PDF ' + idPT + ': ' + e.message);
    return '';
  }
}

/** Reemplazo de texto simple (escapa regex) */
function reemplazar(body, placeholder, valor) {
  let regex = escapeRegex(placeholder);
  body.replaceText(regex, (valor !== undefined && valor !== null && valor !== '') ? valor.toString() : '—');
}

/** Reemplazo para placeholders con caracteres regex como ¿? */
function reemplazarEspecial(body, placeholder, valor) {
  reemplazar(body, placeholder, valor);
}

/** Reemplaza SOLO la primera ocurrencia (para tags duplicados con valores distintos) */
function reemplazarPrimero(body, placeholder, valor) {
  let regex = escapeRegex(placeholder);
  let found = body.findText(regex);
  if (!found) return;
  let element = found.getElement();
  let start = found.getStartOffset();
  let end = found.getEndOffsetInclusive();
  let texto = (valor !== undefined && valor !== null && valor !== '') ? valor.toString() : '—';
  element.asText().deleteText(start, end);
  element.asText().insertText(start, texto);
}

/** 
 * Reemplaza un placeholder por una imagen basándose en su código Base64.
 * Si no hay Base64, coloca un guion "—".
 */
function reemplazarFirma(body, placeholder, base64Data) {
  let regex = escapeRegex(placeholder);
  let found = body.findText(regex);
  if (!found) return;

  let element = found.getElement();
  let start = found.getStartOffset();
  let end = found.getEndOffsetInclusive();

  if (base64Data && base64Data.indexOf('data:image') === 0) {
    try {
      // Borrar texto placeholder
      element.asText().deleteText(start, end);
      
      // Decodificar imagen
      let base64 = base64Data.split(',')[1];
      let blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/png');
      
      // Insertar imagen
      let parent = element.getParent();
      // Verificamos si es un párrafo u otro elemento válido para hospedar una imagen inline
      if(parent.getType() === DocumentApp.ElementType.PARAGRAPH || parent.getType() === DocumentApp.ElementType.LIST_ITEM) {
         let inlineImage = parent.asParagraph().insertInlineImage(start, blob);
         // Dimensiones proporcionales (ancho máximo 120px pa q quepa en caja)
         let width = inlineImage.getWidth();
         let height = inlineImage.getHeight();
         let newWidth = 120;
         let newHeight = (height * newWidth) / width;
         inlineImage.setWidth(newWidth);
         inlineImage.setHeight(newHeight);
      } else {
         element.asText().insertText(start, '✓ Firmado (Error Render)');
      }
    } catch(e) {
      Logger.log("Error decodificando firma " + placeholder + ": " + e.message);
      element.asText().insertText(start, '—');
    }
  } else {
    element.asText().deleteText(start, end);
    element.asText().insertText(start, '—');
  }
}

/** Escapa caracteres especiales para regex de Java (usado en replaceText) */
function escapeRegex(text) {
  return text.replace(/[.*+?^${}()|[\]\\¿¡]/g, '\\$&');
}

function buscarOCrearCarpeta(padre, nombre) {
  const carpetas = padre.getFoldersByName(nombre);
  return carpetas.hasNext() ? carpetas.next() : padre.createFolder(nombre);
}