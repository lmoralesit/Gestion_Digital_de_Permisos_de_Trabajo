function generarPDF_PT(idPT, filaMaestro) {
  try {
    const hM = conectarDB().getSheetByName('PT_Maestro');
    const datos = hM.getRange(filaMaestro, 1, 1, 40).getValues()[0];
    
    // Crear copia temporal del documento
    const tempFile = DriveApp.getFileById(CONFIG.DOC_TEMPLATE_ID).makeCopy(`TEMP_${idPT}`);
    const doc = DocumentApp.openById(tempFile.getId());
    const body = doc.getBody();

    // Reemplazo básico de etiquetas (Módulo 1-3)
    body.replaceText('<<Unidad de Negocio>>', datos[7]);
    body.replaceText('<<Localidad>>', datos[8]);
    body.replaceText('<<Sede>>', datos[9]);
    body.replaceText('<<Nombre Empresa>>', datos[4]);
    body.replaceText('<<Área / Equipo>>', datos[10]);
    body.replaceText('<<Descripción de la Actividad>>', datos[18]);
    
    doc.saveAndClose();

    // Generar PDF
    const pdfBlob = tempFile.getAs(MimeType.PDF);
    pdfBlob.setName(`${idPT}.pdf`);
    
    // Organizar en Drive: Año > Mes > Contratista
    const fecha = new Date();
    const carpetaRaiz = DriveApp.getFolderById(CONFIG.FOLDER_RAIZ_ID);
    const carpetaAnio = buscarOCrearCarpeta(carpetaRaiz, fecha.getFullYear().toString());
    const carpetaMes = buscarOCrearCarpeta(carpetaAnio, ('0' + (fecha.getMonth() + 1)).slice(-2));
    const carpetaContratista = buscarOCrearCarpeta(carpetaMes, datos[4]); // Nombre Empresa
    
    const archivoPDF = carpetaContratista.createFile(pdfBlob);
    
    // Guardar URL en Sheet y borrar temporal
    hM.getRange(filaMaestro, 33).setValue(archivoPDF.getUrl());
    tempFile.setTrashed(true);

  } catch (e) { Logger.log("Error generando PDF: " + e.message); }
}

function buscarOCrearCarpeta(padre, nombre) {
  const carpetas = padre.getFoldersByName(nombre);
  return carpetas.hasNext() ? carpetas.next() : padre.createFolder(nombre);
}