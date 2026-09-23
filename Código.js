/** Configuración central de hojas, columnas, estados y concurrencia. */
var CONFIG = {
  HOJAS: {
    SEGUIMIENTO: 'CEG-F-013 PLANILLADE SEGUIMIENT',
    ASESORES: 'asesores'
  },
  PRIMERA_FILA_DATOS: 5,
  COLUMNAS: {
    COL_FECHA: 1,
    COL_DOC: 2,
    COL_NOMBRE: 3,
    COL_TELEFONO: 4,
    COL_TELEFONO_ALT: 5,
    COL_MONTURA: 6,
    COL_LENTE: 7,
    COL_ACCESORIOS: 8,
    COL_VALOR_MONTURA: 9,
    COL_VALOR_LENTE: 10,
    COL_VALOR_ACCESORIOS: 11,
    COL_TOTAL: 12,
    COL_RESPONSABLE_REGISTRO: 13,
    COL_RESPONSABLE_PAGO: 14,
    COL_ESTADO_FACT: 15,
    COL_FACTURADOR: 16,
    COL_ORDEN_LAB: 17,
    COL_LAB: 18,
    COL_FECHA_ENVIO: 19,
    COL_FECHA_INGRESO: 20,
    COL_ASESOR_LLAMADO: 21,
    COL_FECHA_LLAMADO: 22,
    COL_FECHA_ENTREGA: 23,
    COL_ESTADO_ENTREGA: 24,
    COL_RESPONSABLE_ENTREGA: 25,
    COL_OBSERVACIONES: 26,
    COL_TOTAL_REGISTRADO: 27,
    COL_ABONO_ACUMULADO: 28,
    COL_SALDO: 29
  },
  ESTADOS_FACTURACION: ['PENDIENTE', 'ABONO', 'HACER'],
  TIEMPO_BLOQUEO_MS: 10000
};

/**
 * Renderiza la aplicación web de Google Apps Script.
 * @return {HtmlOutput} Interfaz HTML evaluada para el navegador.
 */
function doGet() {
  try {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Sistema de Gestión de Órdenes - Óptica')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput(
      '<h2>Error al cargar la aplicación</h2>' +
      '<p>' + escapeHtml_(error.message || error.toString()) + '</p>'
    );
  }
}

/**
 * Escapa texto antes de incorporarlo a una respuesta HTML.
 * @param {*} value Valor que se convertirá a texto.
 * @return {string} Texto seguro para HTML.
 */
function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Incluye un archivo HTML parcial en una plantilla de Apps Script.
 * @param {string} filename Nombre del archivo HTML.
 * @return {string} Contenido HTML del archivo.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Obtiene una hoja por nombre y conserva compatibilidad con hojas antiguas.
 * @param {string} nombreHoja Nombre lógico de la hoja.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} Hoja seleccionada.
 */
function obtenerHoja(nombreHoja) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(nombreHoja) || ss.getSheets()[0];
}

/**
 * Normaliza el modo operativo recibido desde el cliente.
 * @param {*} modo Valor recibido en la solicitud.
 * @return {string} Modo en minúsculas y sin espacios externos.
 */
function normalizarModo_(modo) {
  return String(modo || '').trim().toLowerCase();
}

/**
 * Valida que una fila pertenezca al área de datos.
 * @param {*} fila Índice de fila recibido desde el cliente.
 * @return {number} Índice de fila validado.
 * @throws {Error} Si la fila no es un registro válido.
 */
function validarFilaDatos_(fila) {
  var rowIndex = Number(fila);

  if (!Number.isInteger(rowIndex) || rowIndex < CONFIG.PRIMERA_FILA_DATOS) {
    throw new Error('La fila debe ser un registro válido desde la fila ' + CONFIG.PRIMERA_FILA_DATOS + '.');
  }

  return rowIndex;
}

/**
 * Valida las credenciales contra la hoja de asesores.
 * @param {string} usuario Identificador del asesor.
 * @param {string} pass Contraseña proporcionada por el usuario.
 * @return {{success: boolean, usuarioValido?: string, message?: string}} Resultado de autenticación.
 */
function validarLogin(usuario, pass) {
  var sheet = obtenerHoja(CONFIG.HOJAS.ASESORES); 
  var data = sheet.getDataRange().getValues();
  
  var userClean = usuario ? usuario.toString().trim().toLowerCase() : '';
  var passClean = pass ? pass.toString().trim() : '';

  for (var i = CONFIG.PRIMERA_FILA_DATOS - 1; i < data.length; i++) {
    var userSheet = data[i][0] ? data[i][0].toString().trim().toLowerCase() : '';
    var passSheet = data[i][1] ? data[i][1].toString().trim() : '';
    var nombreAsesor = data[i][2] ? data[i][2].toString().trim() : '';
    
    if (userSheet === userClean && passSheet === passClean) {
      return { success: true, usuarioValido: nombreAsesor };
    }
  }
  return { success: false, message: 'Usuario o contraseña incorrectos.' };
}

/**
 * Lee en bloque el rango de datos de seguimiento.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hoja Hoja de seguimiento.
 * @return {Array<Array<*>>} Filas A:AC desde la primera fila de datos.
 */
function obtenerDatosSeguimiento_(hoja) {
  var ultimaFila = hoja.getLastRow();
  var primeraFila = CONFIG.PRIMERA_FILA_DATOS;
  var cantidadFilas = ultimaFila - primeraFila + 1;

  if (cantidadFilas <= 0) return [];

  return hoja.getRange(
    primeraFila,
    1,
    cantidadFilas,
    CONFIG.COLUMNAS.COL_SALDO
  ).getValues();
}

/**
 * Obtiene órdenes pendientes para una fase del flujo operativo.
 * @param {string} fase Identificador de fase: facturacion, envio, llegada o entrega.
 * @return {Array<Object>} Órdenes que cumplen las condiciones de la fase.
 */
function obtenerPendientesPorFase(fase) {
  var sheet = obtenerHoja(CONFIG.HOJAS.SEGUIMIENTO);
  if (!sheet) return [];

  var datos = obtenerDatosSeguimiento_(sheet);
  var columnas = CONFIG.COLUMNAS;
  var lista = [];

  for (var i = 0; i < datos.length; i++) {
    var fila = datos[i];
    var numFila = CONFIG.PRIMERA_FILA_DATOS + i;
    var doc = fila[columnas.COL_DOC - 1];
    var nombre = fila[columnas.COL_NOMBRE - 1];
    
    var estadoFact = fila[columnas.COL_ESTADO_FACT - 1] ? fila[columnas.COL_ESTADO_FACT - 1].toString().trim() : 'PENDIENTE';
    var fechaEnvio = fila[columnas.COL_FECHA_ENVIO - 1] ? fila[columnas.COL_FECHA_ENVIO - 1].toString().trim() : 'PENDIENTE';
    var fechaIngreso = fila[columnas.COL_FECHA_INGRESO - 1] ? fila[columnas.COL_FECHA_INGRESO - 1].toString().trim() : 'PENDIENTE';
    var fechaEntrega = fila[columnas.COL_FECHA_ENTREGA - 1] ? fila[columnas.COL_FECHA_ENTREGA - 1].toString().trim() : 'PENDIENTE';
    var estadoGeneral = fila[columnas.COL_ESTADO_ENTREGA - 1] ? fila[columnas.COL_ESTADO_ENTREGA - 1].toString().trim() : '';

    if (!doc || estadoGeneral === 'ANULADO' || estadoFact === 'ANULADO') continue;

    var cumpleFase = false;

    if (fase === 'facturacion' && (estadoFact === 'PENDIENTE' || estadoFact === '' || estadoFact === 'ABONO')) {
      cumpleFase = true;
    } else if (fase === 'envio' && estadoFact === 'HACER' && (fechaEnvio === 'PENDIENTE' || fechaEnvio === '')) {
      cumpleFase = true;
    }else if (fase === 'llegada' && estadoFact !== 'PENDIENTE' && (fechaIngreso === 'PENDIENTE' || fechaIngreso === '')) {
      cumpleFase = true;
    } else if (fase === 'entrega' && fechaIngreso !== 'PENDIENTE' && (fechaEntrega === 'PENDIENTE' || fechaEntrega === '')) {
      cumpleFase = true;
    }

    if (cumpleFase) {
      lista.push({
        filaExacta: numFila,
        fecha: fila[columnas.COL_FECHA - 1] ? formatDate(fila[columnas.COL_FECHA - 1]) : '',
        documento: doc,
        nombre: nombre,
        telefono: fila[columnas.COL_TELEFONO - 1] || '',
        montura: fila[columnas.COL_MONTURA - 1] || '',
        lente: fila[columnas.COL_LENTE - 1] || '',
        accesorios: fila[columnas.COL_ACCESORIOS - 1] || '',
        estadoFact: estadoFact,
        colQ: fila[columnas.COL_ORDEN_LAB - 1] ? fila[columnas.COL_ORDEN_LAB - 1].toString().trim() : '',
        colR: fila[columnas.COL_LAB - 1] ? fila[columnas.COL_LAB - 1].toString().trim() : '',
        fechaIngreso: fechaIngreso,
        colS: fila[columnas.COL_FECHA_ENVIO - 1] ? formatDate(fila[columnas.COL_FECHA_ENVIO - 1]) : '',
        colZ: fila[columnas.COL_OBSERVACIONES - 1] || '',
        total: fila[columnas.COL_TOTAL - 1] || 0,
        valorTotal: fila[columnas.COL_TOTAL_REGISTRADO - 1] || fila[columnas.COL_TOTAL - 1] || 0,
        abonoAcumulado: fila[columnas.COL_ABONO_ACUMULADO - 1] || 0
      });
    }
  }
  return lista;
}

/**
 * Anula una orden y registra la razón en el historial de observaciones.
 * @param {number|string} fila Fila de la orden.
 * @param {string} motivo Motivo de la anulación.
 * @param {string} usuario Usuario que ejecuta la acción.
 * @return {{success: boolean, message: string}} Resultado de la operación.
 */
function anularOrden(fila, motivo, usuario) {
  var lock = LockService.getScriptLock();

  if (!lock.tryLock(CONFIG.TIEMPO_BLOQUEO_MS)) {
    return { success: false, message: 'No se pudo obtener el bloqueo de la hoja. Intente nuevamente.' };
  }

  try {
    var sheet = obtenerHoja(CONFIG.HOJAS.SEGUIMIENTO);
    var columnas = CONFIG.COLUMNAS;
    var rowIndex = validarFilaDatos_(fila);

    var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    var marcaAnulado = '[ORDEN ANULADA por ' + usuario + ' el ' + ahora + '. Motivo: ' + motivo + ']';
    
    var obsAntigua = sheet.getRange(rowIndex, columnas.COL_OBSERVACIONES).getValue().toString().trim(); 
    var obsFinal = (obsAntigua && obsAntigua !== 'SIN OBSERVACIONES') ? (obsAntigua + ' // ' + marcaAnulado) : marcaAnulado;

    sheet.getRange(rowIndex, columnas.COL_ESTADO_FACT).setValue('ANULADO');
    sheet.getRange(rowIndex, columnas.COL_ESTADO_ENTREGA).setValue('ANULADO');
    sheet.getRange(rowIndex, columnas.COL_OBSERVACIONES).setValue(obsFinal);

    return { success: true, message: 'Orden anulada correctamente.' };
  } catch (err) {
    return { success: false, message: 'Error al anular: ' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Calcula los indicadores operativos del dashboard en una sola lectura.
 * @return {{hoy: number, fact: number, rec: number, ent: number}} Métricas agregadas.
 */
function obtenerMetricasDashboard() {
  var sheet = obtenerHoja(CONFIG.HOJAS.SEGUIMIENTO);
  if (!sheet) return { hoy: 0, fact: 0, rec: 0, ent: 0 };

  var datos = obtenerDatosSeguimiento_(sheet);
  var columnas = CONFIG.COLUMNAS;
  var hoyStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
  var hoyCount = 0, pendFact = 0, pendRec = 0, pendEnt = 0;

  for (var i = 0; i < datos.length; i++) {
    var fila = datos[i];
    var doc = fila[columnas.COL_DOC - 1];
    var fechaIng = fila[columnas.COL_FECHA - 1] ? formatDate(fila[columnas.COL_FECHA - 1]) : '';
    var estadoFact = fila[columnas.COL_ESTADO_FACT - 1] ? fila[columnas.COL_ESTADO_FACT - 1].toString().trim() : 'PENDIENTE';
    var fechaIngProd = fila[columnas.COL_FECHA_INGRESO - 1] ? fila[columnas.COL_FECHA_INGRESO - 1].toString().trim() : 'PENDIENTE';
    var fechaEntrega = fila[columnas.COL_FECHA_ENTREGA - 1] ? fila[columnas.COL_FECHA_ENTREGA - 1].toString().trim() : 'PENDIENTE';
    var estadoGeneral = fila[columnas.COL_ESTADO_ENTREGA - 1] ? fila[columnas.COL_ESTADO_ENTREGA - 1].toString().trim() : '';

    if (!doc || estadoGeneral === 'ANULADO' || estadoFact === 'ANULADO') continue;

    if (fechaIng === hoyStr) hoyCount++;

    if (estadoFact === 'PENDIENTE' || estadoFact === '' || estadoFact === 'ABONO') {
      pendFact++;
    } else if (estadoFact === 'HACER' && (fechaIngProd === 'PENDIENTE' || fechaIngProd === '')) {
      pendRec++;
    } else if (fechaEntrega === 'PENDIENTE' || fechaEntrega === '') {
      pendEnt++;
    }
  }

  return { hoy: hoyCount, fact: pendFact, rec: pendRec, ent: pendEnt };
}

/**
 * Convierte un valor en un número no negativo utilizable por el dominio financiero.
 * @param {*} valor Valor recibido desde la hoja o el cliente.
 * @return {number} Número normalizado o cero.
 */
function numeroSeguro_(valor) {
  var numero = Number(valor);
  return isFinite(numero) && numero >= 0 ? numero : 0;
}

/**
 * Obtiene el resumen financiero persistido de una orden.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Hoja de seguimiento.
 * @param {number} rowIndex Fila de la orden.
 * @return {{total: number, abonado: number, saldo: number}} Resumen financiero.
 */
function obtenerResumenFinanciero_(sheet, rowIndex) {
  var totalRegistrado = sheet.getRange(rowIndex, 27).getValue();
  var totalBase = sheet.getRange(rowIndex, 12).getValue();
  var abonoRegistrado = sheet.getRange(rowIndex, 28).getValue();
  var totalFuente = totalRegistrado === '' || totalRegistrado === null ? totalBase : totalRegistrado;
  var total = numeroSeguro_(totalFuente);
  var abonado = numeroSeguro_(abonoRegistrado);

  return {
    total: total,
    abonado: abonado,
    saldo: Math.max(0, total - abonado)
  };
}

/**
 * Agrega una observación al historial existente de una orden.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hoja Hoja de seguimiento.
 * @param {number} fila Fila de la orden.
 * @param {string} observacionNueva Texto adicional para el historial.
 * @return {void}
 */
function actualizarObservaciones_(hoja, fila, observacionNueva) {
  if (!observacionNueva) return;

  var columna = CONFIG.COLUMNAS.COL_OBSERVACIONES;
  var anterior = hoja.getRange(fila, columna).getValue().toString().trim();
  var final = (anterior && anterior !== 'SIN OBSERVACIONES')
    ? anterior + ' // ' + observacionNueva
    : observacionNueva;

  hoja.getRange(fila, columna).setValue(final);
}

/**
 * Crea una orden o corrige sus datos iniciales bajo el lock de la operación.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hoja Hoja de seguimiento.
 * @param {Object} data Datos normalizados del formulario.
 * @return {{success: boolean, message: string, filaExacta?: number}} Resultado de registro.
 */
function procesarRegistroInicial_(hoja, data) {
  var columnas = CONFIG.COLUMNAS;
  var filaActiva = data.filaActiva;

  if (filaActiva && filaActiva !== '') {
    var fila = validarFilaDatos_(filaActiva);
    var totalCorregido = numeroSeguro_(data.colJ) + numeroSeguro_(data.colK);

    hoja.getRange(fila, columnas.COL_FECHA, 1, columnas.COL_TOTAL).setValues([[
      data.colA, data.colB, data.colC, data.colD, data.colE,
      data.colF, data.colG, data.colH, 0,
      numeroSeguro_(data.colJ), numeroSeguro_(data.colK), totalCorregido
    ]]);
    hoja.getRange(fila, columnas.COL_RESPONSABLE_PAGO).setValue(data.colN);
    hoja.getRange(fila, columnas.COL_ORDEN_LAB, 1, 3).setValues([[data.colQ, data.colR, data.colS]]);

    var ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    var marca = '[CORRECCIÓN FASE 1 por ' + data.usuarioActivo + ' el ' + ahora + ']';
    var anterior = hoja.getRange(fila, columnas.COL_OBSERVACIONES).getValue().toString().trim();
    var detalle = data.colZ_nueva ? marca + ': ' + data.colZ_nueva : marca;
    var observaciones = (anterior && anterior !== 'SIN OBSERVACIONES')
      ? anterior + ' // ' + detalle
      : detalle;

    hoja.getRange(fila, columnas.COL_OBSERVACIONES).setValue(observaciones);
    return { success: true, message: '¡Datos iniciales corregidos y auditoría registrada!' };
  }

  var cantidad = hoja.getMaxRows() - CONFIG.PRIMERA_FILA_DATOS + 1;
  var valoresDoc = hoja.getRange(CONFIG.PRIMERA_FILA_DATOS, columnas.COL_DOC, cantidad, 1).getValues();
  var nuevaFila = 0;

  for (var i = 0; i < valoresDoc.length; i++) {
    var valor = valoresDoc[i][0];
    if (valor === '' || valor === null || valor.toString().trim() === '') {
      nuevaFila = CONFIG.PRIMERA_FILA_DATOS + i;
      break;
    }
  }

  if (nuevaFila === 0) {
    nuevaFila = hoja.getMaxRows() + 1;
    hoja.insertRowsAfter(hoja.getMaxRows(), 1);
  }

  var total = numeroSeguro_(data.colJ) + numeroSeguro_(data.colK);
  var estadoFact = data.colO ? data.colO.toString().trim().toUpperCase() : 'PENDIENTE';
  if (CONFIG.ESTADOS_FACTURACION.indexOf(estadoFact) === -1) estadoFact = 'PENDIENTE';

  var linea = [
    data.colA, data.colB, data.colC, data.colD, data.colE,
    data.colF, data.colG, data.colH, 0,
    numeroSeguro_(data.colJ), numeroSeguro_(data.colK), total,
    data.usuarioActivo, data.colN, estadoFact,
    data.colP || data.usuarioActivo || 'PENDIENTE',
    data.colQ, data.colR, data.colS,
    'PENDIENTE', 'PENDIENTE', 'PENDIENTE', 'PENDIENTE', 'PENDIENTE', 'PENDIENTE',
    data.colZ_nueva || 'SIN OBSERVACIONES', total, 0, total
  ];

  hoja.getRange(nuevaFila, columnas.COL_FECHA, 1, linea.length).setValues([linea]);
  var documentoGuardado = hoja.getRange(nuevaFila, columnas.COL_DOC).getDisplayValue().toString().trim();
  if (documentoGuardado !== String(data.colB || '').trim()) {
    throw new Error('No se pudo verificar la identificación guardada en la fila ' + nuevaFila + '.');
  }

  SpreadsheetApp.flush();
  return { success: true, message: '¡Orden registrada exitosamente!', filaExacta: nuevaFila };
}

/**
 * Procesa el estado de facturación, abonos, saldo e historial financiero.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hoja Hoja de seguimiento.
 * @param {number} fila Fila de la orden.
 * @param {Object} data Datos de facturación enviados por el cliente.
 * @return {{success: boolean, message: string, financiero?: Object}} Resultado financiero.
 */
function procesarFacturacion_(hoja, fila, data) {
  var columnas = CONFIG.COLUMNAS;
  var estado = data.colO ? data.colO.toString().trim().toUpperCase() : 'PENDIENTE';
  if (CONFIG.ESTADOS_FACTURACION.indexOf(estado) === -1) {
    return { success: false, message: 'Estado de facturación no válido.' };
  }

  var resumen = obtenerResumenFinanciero_(hoja, fila);
  var nuevoAbono = numeroSeguro_(data.nuevoAbono);
  if (estado === 'ABONO' && nuevoAbono <= 0) {
    return { success: false, message: 'El abono debe ser mayor que cero.' };
  }
  if (nuevoAbono > resumen.saldo) {
    return { success: false, message: 'El abono no puede superar el saldo pendiente de $' + resumen.saldo.toLocaleString() };
  }

  var abonado = resumen.abonado + nuevoAbono;
  var saldo = Math.max(0, resumen.total - abonado);
  var estadoFinal = (resumen.total === 0 || saldo === 0) ? 'HACER' : estado;
  hoja.getRange(fila, columnas.COL_ESTADO_FACT, 1, 2).setValues([[estadoFinal, data.usuarioActivo]]);
  if (estadoFinal === 'HACER') hoja.getRange(fila, columnas.COL_FECHA_ENVIO).setValue(data.colS || 'PENDIENTE');
  hoja.getRange(fila, columnas.COL_TOTAL_REGISTRADO, 1, 3).setValues([[resumen.total, abonado, saldo]]);

  var anterior = hoja.getRange(fila, columnas.COL_OBSERVACIONES).getValue().toString().trim();
  var observaciones = data.colZ_nueva ? (anterior ? anterior + ' // ' + data.colZ_nueva : data.colZ_nueva) : anterior;
  if (nuevoAbono > 0 || estadoFinal === 'HACER') {
    var fechaHora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    observaciones += ' | [' + fechaHora + ' - ' + data.usuarioActivo + ']: PAGO/ABONO REGISTRADO: +$' + nuevoAbono.toLocaleString() +
      ' (Total Abonado: $' + abonado.toLocaleString() + ' / Saldo: $' + saldo.toLocaleString() + ' - Estado: ' + estadoFinal + ')';
  }
  hoja.getRange(fila, columnas.COL_OBSERVACIONES).setValue(observaciones);

  return {
    success: true,
    message: 'Abono procesado e historial actualizado correctamente.',
    financiero: { total: resumen.total, abonado: abonado, saldo: saldo, estado: estadoFinal }
  };
}

/**
 * Actualiza los datos de envío al laboratorio.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hoja Hoja de seguimiento.
 * @param {number} fila Fila de la orden.
 * @param {Object} data Datos Q:S del formulario.
 * @return {{success: boolean, message: string}} Resultado de la actualización.
 */
function procesarEnvioLab_(hoja, fila, data) {
  hoja.getRange(fila, CONFIG.COLUMNAS.COL_ORDEN_LAB, 1, 3).setValues([[data.colQ, data.colR, data.colS]]);
  actualizarObservaciones_(hoja, fila, data.colZ_nueva);
  return { success: true, message: '¡Datos de envío a laboratorio actualizados correctamente!' };
}

/**
 * Registra la llegada del producto terminado y la trazabilidad del asesor.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hoja Hoja de seguimiento.
 * @param {number} fila Fila de la orden.
 * @param {Object} data Datos T:V del formulario.
 * @return {{success: boolean, message: string}} Resultado de la actualización.
 */
function procesarLlegadaProducto_(hoja, fila, data) {
  hoja.getRange(fila, CONFIG.COLUMNAS.COL_FECHA_INGRESO, 1, 3).setValues([[data.colT, data.usuarioActivo, data.colV]]);
  actualizarObservaciones_(hoja, fila, data.colZ_nueva);
  return { success: true, message: '¡Recepción de producto registrada!' };
}

/**
 * Registra la entrega de la orden al paciente.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} hoja Hoja de seguimiento.
 * @param {number} fila Fila de la orden.
 * @param {Object} data Datos W:Y del formulario.
 * @return {{success: boolean, message: string}} Resultado de la actualización.
 */
function procesarEntregaPaciente_(hoja, fila, data) {
  hoja.getRange(fila, CONFIG.COLUMNAS.COL_FECHA_ENTREGA, 1, 3).setValues([[data.colW, data.colX, data.usuarioActivo]]);
  actualizarObservaciones_(hoja, fila, data.colZ_nueva);
  return { success: true, message: '¡Entrega al paciente completada!' };
}

/**
 * Punto de entrada del backend para las operaciones del ciclo de vida de una orden.
 * @param {Object} data Payload enviado por el frontend.
 * @return {{success: boolean, message: string, filaExacta?: number, financiero?: Object}} Resultado de la operación.
 */
function procesarFormulario(data) {
  var lock = LockService.getScriptLock();

  if (!lock.tryLock(CONFIG.TIEMPO_BLOQUEO_MS)) {
    return { success: false, message: 'Hay otra operación en curso. Espere unos segundos e intente nuevamente.' };
  }

  try {
    if (!data) return { success: false, message: 'No se recibieron datos del formulario.' };
    var modo = normalizarModo_(data.mode);
    var sheet = obtenerHoja(CONFIG.HOJAS.SEGUIMIENTO);
    
    if (!sheet) return { success: false, message: 'Error: No se encontró la pestaña de seguimiento.' };

    if (modo === 'registro') return procesarRegistroInicial_(sheet, data);

    var fila = validarFilaDatos_(data.filaActiva);
    if (modo === 'facturacion') return procesarFacturacion_(sheet, fila, data);
    if (modo === 'envio') return procesarEnvioLab_(sheet, fila, data);
    if (modo === 'llegada') return procesarLlegadaProducto_(sheet, fila, data);
    if (modo === 'entrega') return procesarEntregaPaciente_(sheet, fila, data);

    return { success: false, message: 'Modo de operación no válido.' };

  } catch (error) {
    return { success: false, message: 'Error en el servidor: ' + error.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Formatea fechas de Sheets para el contrato JSON del frontend.
 * @param {*} dateVal Valor de fecha almacenado en la hoja.
 * @return {*} Fecha en formato yyyy-MM-dd o valor original.
 */
function formatDate(dateVal) {
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return dateVal;
}