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

// ===================== ROLES Y CORREOS =====================

function getCorreoActual() {
  return Session.getActiveUser().getEmail() || '';
}

function obtenerCorreoRol(rol) {
  const data = conectarDB().getSheetByName('Config_Roles').getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === rol) return data[i][2];
  }
  return CONFIG.EMAIL_PRUEBA; // Fallback
}

function extraerCorreoDeString(texto, rolDefault) {
  if (!texto) return obtenerCorreoRol(rolDefault) || CONFIG.EMAIL_PRUEBA;
  if (texto.includes('(')) {
    return texto.split('(')[1].replace(')', '').trim();
  }
  return obtenerCorreoRol(rolDefault) || CONFIG.EMAIL_PRUEBA;
}

// ===================== BANCO DE FIRMAS =====================

function obtenerFirmaGuardada(correo) {
  if(!correo) return null;
  const data = conectarDB().getSheetByName('Banco_Firmas').getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === correo.toLowerCase()) {
      return data[i][2]; // Firma_Base64
    }
  }
  return null;
}

function guardarFirmaBanco(correo, nombre, firmaBase64) {
  if(!correo || !firmaBase64) return false;
  const hBanco = conectarDB().getSheetByName('Banco_Firmas');
  const data = hBanco.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === correo.toLowerCase()) {
      hBanco.getRange(i + 1, 3).setValue(firmaBase64); // Actualizar firma
      hBanco.getRange(i + 1, 4).setValue(new Date());  // Actualizar fecha
      return true;
    }
  }
  // Si no existe, agregar fila
  hBanco.appendRow([correo.toLowerCase(), nombre || 'N/A', firmaBase64, new Date()]);
  return true;
}

// ===================== CREAR NUEVO PERMISO (Módulos 1-3) =====================

function guardarNuevoPT(datos) {
  try {
    // SECURITY PATCH: Prevent IDOR / Suplantación de identidad
    const correoSesion = Session.getActiveUser().getEmail();
    if (!correoSesion) return { success: false, message: 'Usuario no autenticado en Google.' };
    
    const auth = validarAcceso(correoSesion);
    if (!auth.autorizado) return { success: false, message: auth.motivo };
    
    datos.correoLogin = correoSesion;
    if (auth.tipo === 'Contratista') {
        datos.empresa = auth.empresa; // Forzar empresa si es contratista
    }

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
      enviarCorreoNotificacion(obtenerCorreoRol('Medico') || CONFIG.EMAIL_PRUEBA, idPT,
        "Evaluación Médica Requerida",
        "Se ha registrado un trabajo de Alto Riesgo. Requiere validación médica antes de su aprobación.",
        `${urlBase}?vista=medico&token=${tokenMedico}`, "Ir a Módulo Médico"
      );
    } else {
      let correoSol = extraerCorreoDeString(datos.solicitanteInterno, 'Solicitante');
      enviarCorreoNotificacion(correoSol, idPT,
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


// ===================== HELPER BUSQUEDA O(1) =====================
function buscarFilaPorTokenO1(nombreHoja, token) {
  const hoja = conectarDB().getSheetByName(nombreHoja);
  if (!hoja) return null;
  const match = hoja.createTextFinder(token).matchEntireCell(true).findNext();
  if (match) {
    const fila = match.getRow();
    const headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    const valores = hoja.getRange(fila, 1, 1, hoja.getLastColumn()).getValues()[0];
    return { fila: fila, vals: valores };
  }
  return null;
}

// ===================== MÓDULO 4: VALIDACIÓN MÉDICA =====================

function obtenerDatosMedicos(token) {
  const ss = conectarDB();
  const hM = ss.getSheetByName('PT_Maestro');

  // ⚡ Optimización Bolt: Búsqueda indexada con TextFinder en lugar de loop O(n)
  const range = hM.getRange(1, 39, hM.getLastRow()).createTextFinder(token).matchEntireCell(true).findNext();

  if(!range) return { success: false, message: "Token inválido o expirado." };

  const rowIndex = range.getRow();
  const rowData = hM.getRange(rowIndex, 1, 1, 39).getValues()[0];
  const idPT = rowData[0];

  // Obtener trabajadores vinculados
  const dataT = ss.getSheetByName('PT_Trabajadores').getDataRange().getValues();
  let trab = [];
  for(let i = 1; i < dataT.length; i++) {
    // ⚡ Alineación Bolt: Usar 'fila' en lugar de 'filaIndex' para compatibilidad con frontend
    if(dataT[i][1] === idPT) trab.push({ fila: i + 1, nombre: dataT[i][2], cedula: dataT[i][3] });
  }

  // ⚡ Alineación Bolt: Estructura de respuesta compatible con VistaMedico.html
  return {
    success: true,
    pt: idPT,
    empresa: rowData[4],
    tipoTrabajo: rowData[20],
    ubicacion: rowData[8], // Restaurado para evitar regresiones
    trabajadores: trab
  };
}

function guardarEvaluacionMedica(evaluaciones, token, emailMedico) {
  const ss = conectarDB();

  // ⚡ Optimización Bolt: Batch write para trabajadores (O(1) en lugar de O(n) llamadas)
  if (evaluaciones && evaluaciones.length > 0) {
    const hTrab = ss.getSheetByName('PT_Trabajadores');
    const startRow = evaluaciones[0].fila;
    const numRows = evaluaciones.length;

    // Mapear datos para setValues (2D array)
    const values = evaluaciones.map(ev => [
      ev.tension, ev.fc, ev.aptitud, ev.observaciones || ev.observacion, emailMedico, new Date()
    ]);

    hTrab.getRange(startRow, 7, numRows, 6).setValues(values);
  }

  const hM = ss.getSheetByName('PT_Maestro');
  // ⚡ Optimización Bolt: Búsqueda rápida de la fila del PT
  const range = hM.getRange(1, 39, hM.getLastRow()).createTextFinder(token).matchEntireCell(true).findNext();

  if(range){
    const rowIndex = range.getRow();
    const idPT = hM.getRange(rowIndex, 1).getValue();
    const nombreSol = hM.getRange(rowIndex, 13).getValue();
    const newToken = Utilities.getUuid();

    // ⚡ Optimización Bolt: Batch update para columnas contiguas 39 y 40
    hM.getRange(rowIndex, 3).setValue('Pendiente Aprobaciones');
    hM.getRange(rowIndex, 39, 1, 2).setValues([['', newToken]]);

    const correoSol = extraerCorreoDeString(nombreSol, 'Solicitante');
    enviarCorreoNotificacion(correoSol, idPT,
      "Aprobación Final de PT",
      "El Servicio Médico ha validado a los trabajadores. El permiso requiere sus firmas finales.",
      `${ScriptApp.getService().getUrl()}?vista=aprobacion&token=${newToken}`, "Ir a Firmar"
    );
  }
  return { success: true };
}

// ===================== MÓDULO 5: APROBACIÓN SECUENCIAL =====================

function obtenerDatosAprobacion(token) {
  const ss = conectarDB();
  const match = buscarFilaPorTokenO1('PT_Maestro', token);
  if(match && match.vals[39] === token) {
      let dataI = match.vals;
      // Determinar qué actor le toca en base al estado de las firmas
      let fSol = dataI[28];
      let fDueno = dataI[29];
      let fSSA = dataI[30];
      
      let etapa = '';
      if(!fSol) etapa = 'Solicitante';
      else if(!fDueno) etapa = 'Dueno';
      else if(!fSSA) etapa = 'SSA';
      else return { success: false, message: "Este permiso ya fue aprobado por todos los actores." };

      let trab = [];
      const dataTrab = ss.getSheetByName('PT_Trabajadores').getDataRange().getValues();
      for(let j = 1; j < dataTrab.length; j++) {
        if(dataTrab[j][1] === dataI[0]) {
          trab.push({ nombre: dataTrab[j][2], cedula: dataTrab[j][3], cargo: dataTrab[j][4] });
        }
      }
      return {
        success: true,
        etapa: etapa, // Informar al frontend quién debe firmar
        pt: dataI[0],
        empresa: dataI[4],
        localidad: dataI[8],
        sede: dataI[9],
        area: dataI[10],
        desc: dataI[18],
        altoRiesgo: dataI[19],
        tipoRiesgo: dataI[20],
        tareas: dataI[21],
        equipos: dataI[22],
        solicitanteNom: dataI[12], // Nombre
        solicitanteDoc: dataI[11] || '', // Correo (Nota: columna 12 es nombre, pero correo lo guardamos en otra parte ahora. Ajustaremos según Data.gs)
        duenoNom: dataI[15],
        duenoDoc: dataI[14] || '',
        ssaDoc: dataI[5] || '', // Correo del analista inicial
        trabajadores: trab
      };
  }
  return { success: false, message: "Token inválido o ya utilizado." };
}

function guardarAprobacionSecuencial(tokenActual, firmaBase64) {
  const ss = conectarDB();
  const hM = ss.getSheetByName('PT_Maestro');
  const match = buscarFilaPorTokenO1('PT_Maestro', tokenActual);
  if(match && match.vals[39] === tokenActual){
      let fila = match.fila;
      let fSol = match.vals[28];
      let fDueno = match.vals[29];
      let fSSA = match.vals[30];
      let idPT = match.vals[0];
      let urlBase = ScriptApp.getService().getUrl();
      let matchValsForDueño = match.vals; // Local reference to use in replacements

      // Guardar firma en el hueco correspondiente y avanzar
      if(!fSol) {
        hM.getRange(fila, 29).setValue(firmaBase64); // Columna AC (29) = Firma Solicitante
        // Generar token para Dueño
        let newToken = Utilities.getUuid();
        hM.getRange(fila, 40).setValue(newToken); // Pisar token actual
        let correoSiguiente = obtenerCorreoDueño(matchValsForDueño); // Necesitamos extraer su correo (o config)
        enviarCorreoNotificacion(correoSiguiente, idPT, "Aprobación Requerida (Dueño del Área)", "El Solicitante ha firmado. Falta su firma como Dueño del Área.", `${urlBase}?vista=aprobacion&token=${newToken}`, "Ir a Firmar");
        return { success: true, ptActivo: false, idPT: idPT };
      } 
      else if(!fDueno) {
        hM.getRange(fila, 30).setValue(firmaBase64); // Firma Dueño
        // Generar token para SSA
        let newToken = Utilities.getUuid();
        hM.getRange(fila, 40).setValue(newToken);
        let correoSiguiente = matchValsForDueño[5] || obtenerCorreoRol('SST'); // Correo del analista original o rol general SST
        enviarCorreoNotificacion(correoSiguiente, idPT, "Aprobación Requerida (SSA)", "El Solicitante y Dueño han firmado. Falta la firma de Seguridad y Salud (SSA) para activar el PT.", `${urlBase}?vista=aprobacion&token=${newToken}`, "Ir a Firmar");
        return { success: true, ptActivo: false, idPT: idPT };
      }
      else if(!fSSA) {
        hM.getRange(fila, 31).setValue(firmaBase64); // Firma SSA
        hM.getRange(fila, 32).setValue(new Date()); // Fecha Aprobación Real
        hM.getRange(fila, 3).setValue('Permiso Activo');
        hM.getRange(fila, 40).setValue(''); // Limpiar token de aprobación ya finalizado

        // Generar tokens de vida activa: Auditoría y Cierre
        let tokenAud = Utilities.getUuid();
        let tokenCierre = Utilities.getUuid();
        hM.getRange(fila, 41).setValue(tokenAud);
        hM.getRange(fila, 42).setValue(tokenCierre);
        
        // Generar PDF con las firmas completas
        generarPDF_PT(idPT, fila);

        // Notificar activación y enviar URL de auditoría al SSA y cierre al Contratista
        enviarCorreoNotificacion(matchValsForDueño[5] || obtenerCorreoRol('SST'), idPT, "Permiso de Trabajo ACTIVO - Auditoría", "El PT ha sido aprobado por todos los actores. Puede auditar con este enlace.", `${urlBase}?vista=auditoria&token=${tokenAud}`, "Ir a Auditoría");
        enviarCorreoNotificacion(matchValsForDueño[3] || obtenerCorreoRol('Supervisor'), idPT, "Permiso de Trabajo ACTIVO - Cierre", "Al finalizar la jornada laboral o el trabajo, por favor inicie el cierre del formulario.", `${urlBase}?vista=cierre&token=${tokenCierre}`, "Ir a Cierre");

      }
    }
  return { success: false, message: "Error procesando la firma secuencial." };
}

// Función auxiliar simple para leer correos de dueños si están en Config u otra columna.  
// Asume que si no lo encuentra usa rol por defecto o el campo de login.
function obtenerCorreoDueño(filaDatos) {
  // Aquí idealmente deberíamos haber guardado el email del dueño. Si no, usamos fallback
  return obtenerCorreoRol('Dueño') || CONFIG.EMAIL_PRUEBA; 
}

// ===================== MÓDULO 6: AUDITORÍA DE CAMPO =====================

function obtenerDatosAuditoria(token) {
  const match = buscarFilaPorTokenO1('PT_Maestro', token);
  if(match && match.vals[40] === token) {
      let dataI = match.vals;
      return {
        success: true,
        pt: dataI[0],
        empresa: dataI[4],
        localidad: dataI[8],
        sede: dataI[9],
        area: dataI[10],
        desc: dataI[18],
        estatus: dataI[2]
      };
  }
  return { success: false, message: "Token inválido o permiso no encontrado." };
}

function guardarAuditoria(datos, token) {
  try {
    const ss = conectarDB();

    // Buscar el PT por token de auditoría
    const match = buscarFilaPorTokenO1('PT_Maestro', token);
    let idPT = null;
    if(match && match.vals[40] === token) {
        idPT = match.vals[0];
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

// ===================== MÓDULO 7: CIERRE SECUENCIAL DEL PERMISO =====================

function obtenerDatosCierre(token) {
  const ss = conectarDB();
  const data = ss.getSheetByName('PT_Maestro').getDataRange().getValues();
  for(let i = 1; i < data.length; i++) {
    if(data[i][41] === token) { // Col 42 para Cierre
      
      let fContratista = data[i][34]; // Col 35 - OJO: en DB maestra 34 es EstatusCierre. 35 es FirmaContratista. Modificaremos abajo según matriz.
      let fSolicitante = data[i][35]; // Col 36
      let fSSA = data[i][36];         // Col 37

      let etapa = '';
      if(!fContratista) etapa = 'Contratista';
      else if(!fSolicitante) etapa = 'Solicitante';
      else if(!fSSA) etapa = 'SSA';
      else return { success: false, message: "Este permiso ya fue cerrado por todos los actores." };

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
        etapa: etapa,
        pt: data[i][0],
        empresa: data[i][4],
        area: data[i][10],
        desc: data[i][18],
        estatus: data[i][2],
        fechaInicio: data[i][31],
        auditorias: auditorias,
        // Datos para identificar al firmante
        estatusCierreActual: data[i][33], // Si fue Culminado, Suspendido etc.
        contratistaNom: data[i][4],
        contratistaDoc: data[i][3], // correo contratista original
        solicitanteNom: data[i][12],
        solicitanteDoc: data[i][11] || obtenerCorreoRol('Solicitante') || '',
        ssaDoc: data[i][5] || obtenerCorreoRol('SST') || ''
      };
    }
  }
  return { success: false, message: "Token inválido o permiso no encontrado." };
}

function guardarCierreSecuencial(tokenActual, firmaBase64, estatusCierreSelect) {
  try {
    const ss = conectarDB();
    const hM = ss.getSheetByName('PT_Maestro');
    const match = buscarFilaPorTokenO1('PT_Maestro', tokenActual);
    
    if(match && match.vals[41] === tokenActual) {
        let fila = match.fila;
        let dataM_i = match.vals;
        let fContratista = dataM_i[34];
        let fSolicitante = dataM_i[35];
        let fSSA = dataM_i[36];
        let idPT = dataM_i[0];
        let urlBase = ScriptApp.getService().getUrl();

        if(!fContratista) {
          hM.getRange(fila, 34).setValue(estatusCierreSelect); // Culminado / etc.
          hM.getRange(fila, 35).setValue(firmaBase64); // Firma Contratista
          
          let newToken = Utilities.getUuid();
          hM.getRange(fila, 42).setValue(newToken);
          let correoSig = dataM_i[11] || obtenerCorreoRol('Solicitante');
          enviarCorreoNotificacion(correoSig, idPT, "Cierre Requerido (Solicitante)", "El contratista ha finalizado labores y firmado retiro. Confirme recepción de área.", `${urlBase}?vista=cierre&token=${newToken}`, "Ir a Cierre");
          return { success: true, ptCerrado: false, idPT: idPT };
        }
        else if(!fSolicitante) {
          hM.getRange(fila, 36).setValue(firmaBase64); // Firma Solicitante
          
          let newToken = Utilities.getUuid();
          hM.getRange(fila, 42).setValue(newToken);
          let correoSig = dataM_i[5] || obtenerCorreoRol('SST');
          enviarCorreoNotificacion(correoSig, idPT, "Cierre Requerido (SSA)", "El Solicitante aprobó la recepción del área. Requiere firma de cierre administrativo SSA.", `${urlBase}?vista=cierre&token=${newToken}`, "Ir a Cierre");
          return { success: true, ptCerrado: false, idPT: idPT };
        }
        else if(!fSSA) {
          hM.getRange(fila, 37).setValue(firmaBase64); // Firma SSA
          hM.getRange(fila, 3).setValue('Cerrado (' + dataM_i[33] + ')'); // Estatus General
          hM.getRange(fila, 38).setValue(new Date()); // Fecha Cierre Total
          hM.getRange(fila, 42).setValue(''); // Limpiar token cierre, fin de ciclo.

          // Regenerar PDF final completo
          try { generarPDF_PT(idPT, fila); } catch(ex) { Logger.log('Error regenerando PDF: ' + ex.message); }

          return { success: true, ptCerrado: true, idPT: idPT, estatusElegido: dataM_i[33] };
        }
      }
    return { success: false, message: "Token inválido." };
  } catch(e) {
    return { success: false, message: e.message };
  }
}