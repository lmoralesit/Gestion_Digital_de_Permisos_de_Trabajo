// ===================== LECTURA DE DATOS =====================

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
    if (data[i][0] && data[i][1]) {
      colab.push({
        email: data[i][0],
        nombre: data[i][1],
        proceso: data[i][6] || 'N/A',
        departamento: data[i][7] || 'N/A'
      });
    }
  }
  return colab;
}

// ===================== CREAR NUEVO PERMISO (Módulos 1-3) =====================

function guardarNuevoPT(datos) {
  try {
    const ss = conectarDB();
    const hojaMaestro = ss.getSheetByName('PT_Maestro');
    const year = new Date().getFullYear();
    let nextNumber = 1;
    const lastRow = hojaMaestro.getLastRow();
    
    if (lastRow > 1) {
      const lastId = hojaMaestro.getRange(lastRow, 1).getValue().toString();
      if (lastId.includes(year.toString())) {
        let num = parseInt(lastId.split('-')[2]);
        if(!isNaN(num)) nextNumber = num + 1;
      }
    }
    const idPT = `PT-${year}-${nextNumber.toString().padStart(4, '0')}`;
    let requiereMedico = (datos.altoRiesgo === 'SI' && (datos.tipoRiesgo.includes('Altura') || datos.tipoRiesgo.includes('Confinado')));
    let tokenFlujo = Utilities.getUuid(); 
    
    let estatus = requiereMedico ? 'Pendiente Médico' : 'Pendiente Aprobaciones';
    let tokenMedico = requiereMedico ? tokenFlujo : '';
    let tokenAprobacion = requiereMedico ? '' : tokenFlujo;

    // Subir foto del área si existe
    let urlFoto = '';
    if (datos.fotoAreaBase64) {
      urlFoto = subirFotoArea(datos.fotoAreaBase64, idPT);
    }

    // Fila: 42 columnas (incluyendo Token_Auditoria y Token_Cierre)
    hojaMaestro.appendRow([
      idPT, new Date(), estatus, datos.correoLogin,
      datos.empresa, datos.nombreSST, datos.cedulaSST,
      datos.unidadNegocio, datos.localidad, datos.sede,
      datos.areaEquipo, urlFoto,
      datos.solicitanteInterno, datos.deptoSolicitante, datos.procesoSolicitante,
      datos.duenoArea, datos.deptoDueno, datos.procesoDueno,
      datos.descripcion, datos.altoRiesgo, datos.tipoRiesgo,
      datos.tareasPrincipales, datos.equiposHerramientas,
      datos.gases.o2, datos.gases.lel, datos.gases.co, datos.gases.h2s,
      datos.firmaBase64, '', '', '',
      '', '', '', '', '', '', '',
      tokenMedico, tokenAprobacion, '', ''
    ]);

    // Guardar trabajadores
    const hTrab = ss.getSheetByName('PT_Trabajadores');
    if (datos.trabajadores && datos.trabajadores.length > 0) {
      let trF = datos.trabajadores.map((t, i) => [
        `${idPT}-T${i+1}`, idPT, t.nombre, t.cedula, t.cargo,
        '', '', '', '', '', '', ''
      ]);
      hTrab.getRange(hTrab.getLastRow() + 1, 1, trF.length, trF[0].length).setValues(trF);
    }

    // Guardar riesgos
    const hRiesgos = ss.getSheetByName('PT_Riesgos');
    if (datos.riesgos && datos.riesgos.length > 0) {
      let rF = datos.riesgos.map((r, i) => {
        let severidadNum = parseInt(r.severidad) || 0;
        let probabilidadNum = parseInt(r.probabilidad) || 0;
        let inherente = severidadNum * probabilidadNum;
        let jerarquiaNum = parseInt(r.jerarquia) || 0;
        let residual = Math.max(1, inherente - jerarquiaNum);
        return [
          `${idPT}-R${i+1}`, idPT, r.peligro, r.severidad,
          r.probabilidad, inherente, r.jerarquia, r.controles, residual
        ];
      });
      hRiesgos.getRange(hRiesgos.getLastRow() + 1, 1, rF.length, rF[0].length).setValues(rF);
    }

    // Enviar notificación
    let urlBase = ScriptApp.getService().getUrl();
    if(requiereMedico) {
      enviarCorreoNotificacion(CONFIG.EMAIL_PRUEBA, idPT,
        "Evaluación Médica Requerida",
        "Se ha registrado un trabajo de Alto Riesgo. Requiere validación médica antes de su aprobación.",
        `${urlBase}?vista=medico&token=${tokenMedico}`, "Ir a Módulo Médico"
      );
    } else {
      enviarCorreoNotificacion(CONFIG.EMAIL_PRUEBA, idPT,
        "Aprobación de PT Requerida",
        "Un nuevo Permiso de Trabajo ha sido solicitado y requiere las firmas de los Validadores.",
        `${urlBase}?vista=aprobacion&token=${tokenAprobacion}`, "Ir a Firmar"
      );
    }

    return { success: true, idPT: idPT, requiereMedico: requiereMedico };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ===================== SUBIR FOTO AL DRIVE =====================

function subirFotoArea(base64Data, idPT) {
  try {
    let contentType = 'image/png';
    let b64 = base64Data;
    if (base64Data.indexOf(',') > -1) {
      let parts = base64Data.split(',');
      let match = parts[0].match(/:(.*?);/);
      if (match) contentType = match[1];
      b64 = parts[1];
    }
    let blob = Utilities.newBlob(Utilities.base64Decode(b64), contentType, `foto_area_${idPT}.png`);
    let carpetaRaiz = DriveApp.getFolderById(CONFIG.FOLDER_RAIZ_ID);
    let archivo = carpetaRaiz.createFile(blob);
    return archivo.getUrl();
  } catch(e) {
    Logger.log('Error subiendo foto: ' + e.message);
    return '';
  }
}

// ===================== MÓDULO 4: VALIDACIÓN MÉDICA =====================

function obtenerDatosPorTokenMedico(token) {
  const dataM = conectarDB().getSheetByName('PT_Maestro').getDataRange().getValues();
  let idPT = null, datosPT = null;
  for(let i = 1; i < dataM.length; i++) {
    if(dataM[i][38] === token) {
      idPT = dataM[i][0];
      datosPT = { id: idPT, empresa: dataM[i][4], ubicacion: dataM[i][8] };
      break;
    }
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
  evaluaciones.forEach(ev => {
    hTrab.getRange(ev.fila, 7, 1, 6).setValues([
      [ev.tension, ev.fc, ev.aptitud, ev.observacion, emailMedico, new Date()]
    ]);
  });

  const hM = ss.getSheetByName('PT_Maestro');
  const data = hM.getDataRange().getValues();
  for(let i = 1; i < data.length; i++){
    if(data[i][38] === token){
      let newToken = Utilities.getUuid();
      hM.getRange(i+1, 3).setValue('Pendiente Aprobaciones');
      hM.getRange(i+1, 39).setValue(''); // limpiar token médico
      hM.getRange(i+1, 40).setValue(newToken); // token aprobación
      let idPT = data[i][0];
      enviarCorreoNotificacion(CONFIG.EMAIL_PRUEBA, idPT,
        "Aprobación Final de PT",
        "El Servicio Médico ha validado a los trabajadores. El permiso requiere sus firmas finales.",
        `${ScriptApp.getService().getUrl()}?vista=aprobacion&token=${newToken}`, "Ir a Firmar"
      );
      break;
    }
  }
  return true;
}

// ===================== MÓDULO 5: APROBACIÓN =====================

function obtenerDatosAprobacion(token) {
  const ss = conectarDB();
  const data = ss.getSheetByName('PT_Maestro').getDataRange().getValues();
  for(let i = 1; i < data.length; i++) {
    if(data[i][39] === token) {
      // Obtener trabajadores para mostrar resumen
      let trab = [];
      const dataTrab = ss.getSheetByName('PT_Trabajadores').getDataRange().getValues();
      for(let j = 1; j < dataTrab.length; j++) {
        if(dataTrab[j][1] === data[i][0]) {
          trab.push({ nombre: dataTrab[j][2], cedula: dataTrab[j][3], cargo: dataTrab[j][4] });
        }
      }
      return {
        success: true,
        pt: data[i][0],
        empresa: data[i][4],
        localidad: data[i][8],
        sede: data[i][9],
        area: data[i][10],
        desc: data[i][18],
        altoRiesgo: data[i][19],
        tipoRiesgo: data[i][20],
        tareas: data[i][21],
        equipos: data[i][22],
        solicitanteInterno: data[i][12],
        duenoArea: data[i][15],
        trabajadores: trab
      };
    }
  }
  return { success: false, message: "Token inválido" };
}

function guardarAprobacion(token, fSol, fDueno, fSSA) {
  const ss = conectarDB();
  const hM = ss.getSheetByName('PT_Maestro');
  const data = hM.getDataRange().getValues();
  for(let i = 1; i < data.length; i++){
    if(data[i][39] === token){
      let fila = i + 1;
      hM.getRange(fila, 29).setValue(fSol);     // Firma Solicitante
      hM.getRange(fila, 30).setValue(fDueno);    // Firma Dueño
      hM.getRange(fila, 31).setValue(fSSA);      // Firma SSA
      hM.getRange(fila, 32).setValue(new Date()); // Fecha Aprobación
      hM.getRange(fila, 3).setValue('Permiso Activo');
      hM.getRange(fila, 40).setValue('');         // Limpiar token aprobación

      // Generar tokens para auditoría y cierre
      let tokenAud = Utilities.getUuid();
      let tokenCierre = Utilities.getUuid();
      hM.getRange(fila, 41).setValue(tokenAud);
      hM.getRange(fila, 42).setValue(tokenCierre);

      let idPT = data[i][0];
      
      // Generar PDF
      generarPDF_PT(idPT, fila);

      // Notificar al auditor SSA
      let urlBase = ScriptApp.getService().getUrl();
      enviarCorreoNotificacion(CONFIG.EMAIL_PRUEBA, idPT,
        "Permiso de Trabajo Activo - Auditoría Disponible",
        "El Permiso de Trabajo ha sido aprobado y se encuentra ACTIVO. Puede realizar la auditoría de campo cuando lo requiera.",
        `${urlBase}?vista=auditoria&token=${tokenAud}`, "Ir a Auditoría"
      );

      // Notificar sobre el cierre
      enviarCorreoNotificacion(CONFIG.EMAIL_PRUEBA, idPT,
        "Permiso de Trabajo Activo - Cierre Disponible",
        "El Permiso de Trabajo ha sido aprobado y se encuentra ACTIVO. Al finalizar la jornada, ingrese para cerrar el permiso.",
        `${urlBase}?vista=cierre&token=${tokenCierre}`, "Ir a Cierre"
      );

      return { success: true, idPT: idPT };
    }
  }
  return { success: false };
}

// ===================== MÓDULO 6: AUDITORÍA DE CAMPO =====================

function obtenerDatosAuditoria(token) {
  const ss = conectarDB();
  const data = ss.getSheetByName('PT_Maestro').getDataRange().getValues();
  for(let i = 1; i < data.length; i++) {
    if(data[i][40] === token) {
      return {
        success: true,
        pt: data[i][0],
        empresa: data[i][4],
        localidad: data[i][8],
        sede: data[i][9],
        area: data[i][10],
        desc: data[i][18],
        estatus: data[i][2]
      };
    }
  }
  return { success: false, message: "Token inválido o permiso no encontrado." };
}

function guardarAuditoria(datos, token) {
  try {
    const ss = conectarDB();

    // Buscar el PT por token de auditoría
    const hM = ss.getSheetByName('PT_Maestro');
    const dataM = hM.getDataRange().getValues();
    let idPT = null;
    for(let i = 1; i < dataM.length; i++) {
      if(dataM[i][40] === token) {
        idPT = dataM[i][0];
        break;
      }
    }
    if(!idPT) return { success: false, message: "Token inválido." };

    // Calcular % cumplimiento
    let items = [datos.epp, datos.higiene, datos.equipos, datos.procedimiento, datos.emergencia, datos.ambiental];
    let cumple = items.filter(v => v === 'Cumple').length;
    let pctCumplimiento = Math.round((cumple / 6) * 1000) / 10; // regla de tres: X/6 * 100

    // Guardar en hoja PT_Auditoria
    const hA = ss.getSheetByName('PT_Auditoria');
    let idAud = `${idPT}-AUD-${(hA.getLastRow())}`;
    hA.appendRow([
      idAud, idPT, new Date(), datos.auditorEmail,
      datos.epp, datos.higiene, datos.equipos,
      datos.procedimiento, datos.emergencia, datos.ambiental,
      datos.observaciones, pctCumplimiento
    ]);

    return { success: true, idAuditoria: idAud, porcentaje: pctCumplimiento };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// ===================== MÓDULO 7: CIERRE DEL PERMISO =====================

function obtenerDatosCierre(token) {
  const ss = conectarDB();
  const data = ss.getSheetByName('PT_Maestro').getDataRange().getValues();
  for(let i = 1; i < data.length; i++) {
    if(data[i][41] === token) {
      // Buscar auditoría si existe
      let auditorias = [];
      const dataAud = ss.getSheetByName('PT_Auditoria').getDataRange().getValues();
      for(let j = 1; j < dataAud.length; j++) {
        if(dataAud[j][1] === data[i][0]) {
          auditorias.push({
            fecha: dataAud[j][2],
            porcentaje: dataAud[j][11]
          });
        }
      }

      return {
        success: true,
        pt: data[i][0],
        empresa: data[i][4],
        area: data[i][10],
        desc: data[i][18],
        estatus: data[i][2],
        fechaInicio: data[i][31],
        auditorias: auditorias
      };
    }
  }
  return { success: false, message: "Token inválido o permiso no encontrado." };
}

function guardarCierre(datos, token) {
  try {
    const ss = conectarDB();
    const hM = ss.getSheetByName('PT_Maestro');
    const dataM = hM.getDataRange().getValues();
    
    for(let i = 1; i < dataM.length; i++) {
      if(dataM[i][41] === token) {
        let fila = i + 1;
        hM.getRange(fila, 3).setValue('Cerrado');            // Estatus
        hM.getRange(fila, 34).setValue(datos.estatusCierre); // Culminado/Suspendido/Extendido
        hM.getRange(fila, 35).setValue(datos.firmaContratista);
        hM.getRange(fila, 36).setValue(datos.firmaSolicitante);
        hM.getRange(fila, 37).setValue(datos.firmaSSA);
        hM.getRange(fila, 38).setValue(new Date());          // Fecha Cierre
        hM.getRange(fila, 42).setValue('');                   // Limpiar token cierre

        let idPT = dataM[i][0];

        // Regenerar PDF con todos los datos (incluye cierre y auditoría)
        try { generarPDF_PT(idPT, fila); } catch(ex) { Logger.log('Error regenerando PDF: ' + ex.message); }

        return { success: true, idPT: idPT };
      }
    }
    return { success: false, message: "Token inválido." };
  } catch(e) {
    return { success: false, message: e.message };
  }
}