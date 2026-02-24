function obtenerListasConfig() {
  const data = conectarDB().getSheetByName('Config_Listas').getDataRange().getValues();
  let listas = { empresas: [], unidades: [], localidades: [], sedesMap: [], tiposTrabajo: [], peligros: [], severidad: [], probabilidad: [], jerarquia: [] };
  for(let i=1; i<data.length; i++) {
    if (data[i][0]) listas.empresas.push(data[i][0]);
    if (data[i][1]) listas.unidades.push(data[i][1]);
    if (data[i][2]) listas.localidades.push(data[i][2]);
    if (data[i][3]) listas.sedesMap.push(data[i][3]); 
    if (data[i][4]) listas.tiposTrabajo.push(data[i][4]);
    if (data[i][5]) listas.peligros.push(data[i][5]);
    if (data[i][6]) listas.severidad.push(data[i][6]);
    if (data[i][7]) listas.probabilidad.push(data[i][7]);
    if (data[i][8]) listas.jerarquia.push(data[i][8]);
  }
  return listas;
}

function obtenerColaboradoresBuk() {
  const data = conectarDB().getSheetByName('Colaboradores').getDataRange().getValues();
  let colab = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][1]) colab.push({ email: data[i][0], nombre: data[i][1], proceso: data[i][6] || 'N/A', departamento: data[i][7] || 'N/A' });
  }
  return colab;
}

function guardarNuevoPT(datos) {
  try {
    const ss = conectarDB();
    const hojaMaestro = ss.getSheetByName('PT_Maestro');
    const year = new Date().getFullYear();
    let nextNumber = 1;
    const lastRow = hojaMaestro.getLastRow();
    
    if (lastRow > 1) {
      const lastId = hojaMaestro.getRange(lastRow, 1).getValue().toString();
      if (lastId.includes(year.toString())) { let num = parseInt(lastId.split('-')[2]); if(!isNaN(num)) nextNumber = num + 1; }
    }
    const idPT = `PT-${year}-${nextNumber.toString().padStart(4, '0')}`;
    let requiereMedico = (datos.altoRiesgo === 'SI' && (datos.tipoRiesgo.includes('Altura') || datos.tipoRiesgo.includes('Confinado')));
    let tokenFlujo = Utilities.getUuid(); 
    
    let estatus = requiereMedico ? 'Pendiente Médico' : 'Pendiente Aprobaciones';
    let tokenMedico = requiereMedico ? tokenFlujo : '';
    let tokenAprobacion = requiereMedico ? '' : tokenFlujo;

    hojaMaestro.appendRow([
      idPT, new Date(), estatus, datos.correoLogin, datos.empresa, datos.nombreSST, datos.cedulaSST, datos.unidadNegocio, datos.localidad, datos.sede, datos.areaEquipo, '', datos.solicitanteInterno, datos.deptoSolicitante, datos.procesoSolicitante, datos.duenoArea, datos.deptoDueno, datos.procesoDueno, datos.descripcion, datos.altoRiesgo, datos.tipoRiesgo, '', '', datos.gases.o2, datos.gases.lel, datos.gases.co, datos.gases.h2s, datos.firmaBase64, '', '', '', '', '', '', '', '', '', '', tokenMedico, tokenAprobacion
    ]);

    const hTrab = ss.getSheetByName('PT_Trabajadores');
    if (datos.trabajadores.length > 0) {
      let trF = datos.trabajadores.map((t, i) => [`${idPT}-T${i+1}`, idPT, t.nombre, t.cedula, t.cargo, '', '', '', '', '', '', '']);
      hTrab.getRange(hTrab.getLastRow() + 1, 1, trF.length, trF[0].length).setValues(trF);
    }

    const hRiesgos = ss.getSheetByName('PT_Riesgos');
    if (datos.riesgos.length > 0) {
      let rF = datos.riesgos.map((r, i) => {
        let inherente = parseInt(r.severidad) * parseInt(r.probabilidad);
        return [`${idPT}-R${i+1}`, idPT, r.peligro, r.severidad, r.probabilidad, inherente, r.jerarquia, r.controles, inherente];
      });
      hRiesgos.getRange(hRiesgos.getLastRow() + 1, 1, rF.length, rF[0].length).setValues(rF);
    }

    let urlBase = ScriptApp.getService().getUrl();
    if(requiereMedico) {
      enviarCorreoNotificacion(CONFIG.EMAIL_PRUEBA, idPT, "Evaluación Médica Requerida", "Se ha registrado un trabajo de Alto Riesgo. Requiere validación médica antes de su aprobación.", `${urlBase}?vista=medico&token=${tokenMedico}`, "Ir a Módulo Médico");
    } else {
      enviarCorreoNotificacion(CONFIG.EMAIL_PRUEBA, idPT, "Aprobación de PT Requerida", "Un nuevo Permiso de Trabajo ha sido solicitado y requiere las firmas de los Validadores.", `${urlBase}?vista=aprobacion&token=${tokenAprobacion}`, "Ir a Firmar");
    }

    return { success: true, idPT: idPT, requiereMedico: requiereMedico };
  } catch (e) { return { success: false, message: e.message }; }
}

function obtenerDatosPorTokenMedico(token) {
  const dataM = conectarDB().getSheetByName('PT_Maestro').getDataRange().getValues();
  let idPT = null, datosPT = null;
  for(let i = 1; i < dataM.length; i++) {
    if(dataM[i][38] === token) { idPT = dataM[i][0]; datosPT = { id: idPT, empresa: dataM[i][4], ubicacion: dataM[i][8] }; break; }
  }
  if(!idPT) return { success: false, message: "Token inválido o expirado." };

  const dataT = conectarDB().getSheetByName('PT_Trabajadores').getDataRange().getValues();
  let trab = [];
  for(let i = 1; i < dataT.length; i++) {
    if(dataT[i][1] === idPT) trab.push({ filaIndex: i + 1, nombre: dataT[i][2], cedula: dataT[i][3] });
  }
  return { success: true, pt: datosPT, trabajadores: trab };
}

function guardarEvaluacionMedica(evaluaciones, token, emailMedico) {
  const ss = conectarDB();
  const hTrab = ss.getSheetByName('PT_Trabajadores');
  evaluaciones.forEach(ev => hTrab.getRange(ev.fila, 7, 1, 6).setValues([[ev.tension, ev.fc, ev.aptitud, ev.observacion, emailMedico, new Date()]]));

  const hM = ss.getSheetByName('PT_Maestro');
  const data = hM.getRange("AM:AM").getValues(); 
  for(let i=0; i<data.length; i++){
    if(data[i][0] === token){
      let newToken = Utilities.getUuid();
      hM.getRange(i+1, 3).setValue('Pendiente Aprobaciones'); hM.getRange(i+1, 39).setValue(''); hM.getRange(i+1, 40).setValue(newToken);
      let idPT = hM.getRange(i+1, 1).getValue();
      enviarCorreoNotificacion(CONFIG.EMAIL_PRUEBA, idPT, "Aprobación Final de PT", "El Servicio Médico ha validado a los trabajadores. El permiso requiere sus firmas finales.", `${ScriptApp.getService().getUrl()}?vista=aprobacion&token=${newToken}`, "Ir a Firmar");
      break;
    }
  }
  return true;
}

function obtenerDatosAprobacion(token) {
  const data = conectarDB().getSheetByName('PT_Maestro').getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][39] === token) return { success: true, pt: data[i][0], desc: data[i][18], area: data[i][10] };
  }
  return { success: false, message: "Token inválido" };
}

function guardarAprobacion(token, fSol, fDueno, fSSA) {
  const hM = conectarDB().getSheetByName('PT_Maestro');
  const data = hM.getRange("AN:AN").getValues(); 
  for(let i=0; i<data.length; i++){
    if(data[i][0] === token){
      hM.getRange(i+1, 29).setValue(fSol); hM.getRange(i+1, 30).setValue(fDueno); hM.getRange(i+1, 31).setValue(fSSA);
      hM.getRange(i+1, 32).setValue(new Date()); hM.getRange(i+1, 3).setValue('Permiso Activo'); hM.getRange(i+1, 40).setValue('');
      let idPT = hM.getRange(i+1, 1).getValue();
      
      // Llamamos a la generación de PDF
      generarPDF_PT(idPT, i+1);

      return { success: true, idPT: idPT };
    }
  }
  return { success: false };
}