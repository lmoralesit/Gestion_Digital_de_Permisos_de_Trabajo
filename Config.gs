const CONFIG = {
  FOLDER_RAIZ_ID: '17dcS0wmYr4Vi918zUxMSZaGwI3yD6rXB', 
  DOC_TEMPLATE_ID: '1FRjYXbYkMBxZu93TUMoRD4A3AW7KYpN7DZ2VAr_f5R4',
  SHEET_DB_ID: '1hbC0FYNYjRy4cUWZyjbG_rlJaXjM_Eqf-x5S3QIawmw',
  BUK_SHEET_ID: '1ZQkiBPUVdOWV8PlSeWDu9Lc6Jr424fhE99yVDHT0ax0',
  EMAIL_PRUEBA: 'lmorales@alfonzorivas.com'
};

function conectarDB() { return SpreadsheetApp.openById(CONFIG.SHEET_DB_ID); }

function configurarBaseDeDatos() {
  const ss = conectarDB();
  const estructuraDB = {
    'PT_Maestro': [
      'ID_PT', 'Fecha_Creacion', 'Estatus_PT', 'Correo_Solicitante',
      'Nombre_Empresa', 'Analista_SST', 'Cedula_Analista',
      'Unidad_Negocio', 'Localidad', 'Sede', 'Area_Equipo', 'URL_Foto_Area',
      'Solicitante_Interno', 'Depto_Solicitante', 'Proceso_Solicitante',
      'Dueno_Area', 'Depto_Dueno', 'Proceso_Dueno',
      'Descripcion_Actividad', 'Es_Alto_Riesgo', 'Tipo_Trabajo_AR',
      'Tareas_Principales', 'Equipos_Herramientas',
      'O2', 'LEL', 'CO', 'H2S',
      'Firma_Analista_SST', 'Firma_Solicitante', 'Firma_Dueno', 'Firma_SSA',
      'Fecha_Aprobacion_Final', 'URL_PDF',
      'Estatus_Cierre', 'Firma_Cierre_Contratista', 'Firma_Cierre_Solicitante',
      'Firma_Cierre_SSA', 'Fecha_Cierre',
      'Token_Medico', 'Token_Aprobacion', 'Token_Auditoria', 'Token_Cierre'
    ],
    'PT_Trabajadores': [
      'ID_Trabajador', 'ID_PT', 'Nombre_Trabajador', 'Cedula', 'Cargo',
      'Firma_Trabajador', 'Tension_Arterial', 'Frecuencia_Cardiaca',
      'Aptitud_Medica', 'Observaciones_Medicas', 'Validado_Por_Medico', 'Fecha_Validacion'
    ],
    'PT_Riesgos': [
      'ID_Riesgo', 'ID_PT', 'Peligro_Identificado', 'Severidad',
      'Probabilidad', 'Riesgo_Inherente', 'Jerarquia_Control',
      'Controles_Aplicados', 'Riesgo_Residual'
    ],
    'PT_Auditoria': [
      'ID_Auditoria', 'ID_PT', 'Fecha_Auditoria', 'Auditor_SSA',
      'Uso_EPP', 'Higiene_Industrial', 'Equipos_Herramientas',
      'Ejecucion_Procedimiento', 'Controles_Emergencia', 'Gestion_Ambiental',
      'Observaciones', 'Porcentaje_Cumplimiento'
    ],
    'Config_Listas': [
      'Empresa_Contratista', 'Unidad_Negocio', 'Localidad', 'Mapeo_Sedes',
      'Tipo_Trabajo_AR', 'Peligros', 'Severidad', 'Probabilidad', 'Jerarquia'
    ]
  };

  const datosListas = [
    ['Contratista Alfa CA', 'ARCO', 'Turmero', 'Turmero|Planta Turmero', 'Trabajo en Altura', 'Peligros Físicos', '1 - Insignificante', '1 - Baja', '4 - Barreras Duras (Preventivas)'],
    ['Servicios Beta SA', 'Indelma', 'La California', 'La California|Planta la California', 'Trabajo en Caliente', 'Peligros Químicos', '2 - Menor', '2 - Media', '3 - Barreras Físicas (Ingeniería)'],
    ['Mantenimiento Gamma', '', 'Chuao', 'Chuao|Oficinas Chuao', 'Espacio Confinado', 'Peligros Biológicos', '3 - Moderado', '3 - Alta', '2 - Barreras Blandas (Administrativas)'],
    ['', '', 'Cagua', 'Cagua|Cagua 1', 'Trabajo de Izamiento', 'Peligros Disergonómicos', '4 - Mayor / Catastrófico', '4 - Muy Alta', '1 - Última Barrera (Mitigación) EPP'],
    ['', '', '', 'Cagua|Cagua 2', 'Trabajo de Excavación', 'Peligros Mecánicos', '', '', ''],
    ['', '', '', 'Cagua|Cagua 3', 'Trabajo con Químicos', 'Peligros Eléctricos', '', '', ''],
    ['', '', '', 'Cagua|Cagua 6', 'Trabajos Eléctricos', 'Peligros Psicosociales', '', '', ''],
    ['', '', '', 'Cagua|Cagua Indelma', '', 'Peligros de Incendio y Explosión', '', '', '']
  ];

  for (let nombreHoja in estructuraDB) {
    let hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) hoja = ss.insertSheet(nombreHoja);
    let encabezados = estructuraDB[nombreHoja];
    hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]).setFontWeight('bold').setBackground('#003366').setFontColor('#ffffff');
    hoja.setFrozenRows(1); hoja.autoResizeColumns(1, encabezados.length);
    if (nombreHoja === 'Config_Listas' && hoja.getLastRow() === 1) hoja.getRange(2, 1, datosListas.length, datosListas[0].length).setValues(datosListas);
  }

  let hojaColab = ss.getSheetByName('Colaboradores');
  if (!hojaColab) hojaColab = ss.insertSheet('Colaboradores');
  hojaColab.getRange('A1').setFormula('=QUERY(IMPORTRANGE("https://docs.google.com/spreadsheets/d/1ZQkiBPUVdOWV8PlSeWDu9Lc6Jr424fhE99yVDHT0ax0/edit?gid=0"; "buk!A:V"); "SELECT Col3, Col2, Col17, Col12, Col13, Col18, Col19, Col20, Col21, Col22")');
  
  let hoja1 = ss.getSheetByName('Hoja 1'); if(hoja1) ss.deleteSheet(hoja1);
}