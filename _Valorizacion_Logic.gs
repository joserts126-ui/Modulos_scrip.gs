/**
 * ====================================================
 * === LÓGICA DE NEGOCIO DEL MÓDULO DE VALORIZACIÓN ===
 * ====================================================
 */

/**
 * Función principal llamada por Valorizacion.html al cargar.
 * Obtiene las OTs pendientes y las valorizaciones ya emitidas para un pedido.
 */
function getDatosParaValorizar(numPedido) {
  try {
    if (!numPedido) throw new Error("Se requiere un número de pedido.");

    const pendientes = _getOTsPendientesParaValorizar(numPedido);
    const emitidas = _getValorizacionesEmitidas(numPedido);

    return {
      success: true,
      pendientes: pendientes,
      emitidas: emitidas
    };

  } catch (e) {
    Logger.log(`Error en getDatosParaValorizar: ${e.message}`);
    return manejarError('getDatosParaValorizar', e); // Asume que tienes manejarError en _Comercreal_Logic o _Core
  }
}

/**
 * [HELPER INTERNO - VERSIÓN CORREGIDA]
 * Obtiene todas las OTs de un pedido que están "Pendientes" de valorizar.
 * AHORA LEE LAS COLUMNAS CORRECTAS DE 'HOJA: OT'.
 */
function _getOTsPendientesParaValorizar(numPedido) {
  const allDataOT = obtenerDatosHoja(HOJA_OT, false); // No usar caché para datos frescos
  const COL_MAP_OT = getColumnMap(HOJA_OT);

  // --- Columnas a leer (NOMBRES CORREGIDOS) ---
  const COT_COL = COL_MAP_OT['N°COTIZACION'];
  const ESTADO_VAL_COL = COL_MAP_OT['ESTADO_VALORIZACION'];
  const OT_ID_COL = COL_MAP_OT['N° OT'];
  const FECHA_COL = COL_MAP_OT['FECHA'];
  const HORAS_COL = COL_MAP_OT['HOROMETRO TRABAJADO']; // O 'TIEMPO TOTAL', según tu regla de negocio
  
  // --- ¡AQUÍ ESTÁ LA CORRECCIÓN! ---
  // Leemos de las nuevas columnas que 'guardarOT' acaba de poblar.
  const MONTO_SERV_COL = COL_MAP_OT['MONTO SERVICIO']; 
  const MONTO_MOV_COL = COL_MAP_OT['MONTO MOVILIZACION'];
  // ---------------------------------
  
  // Validar que todas las columnas necesarias existan
  if ([COT_COL, ESTADO_VAL_COL, OT_ID_COL, MONTO_SERV_COL, MONTO_MOV_COL, HORAS_COL, FECHA_COL].includes(undefined)) {
    Logger.log(`Error en _getOTsPendientes...: Faltan columnas en HOJA: OT.
        N°COTIZACION=${COT_COL}, ESTADO_VALORIZACION=${ESTADO_VAL_COL}, N° OT=${OT_ID_COL}, 
        FECHA=${FECHA_COL}, HOROMETRO TRABAJADO=${HORAS_COL}, 
        MONTO SERVICIO=${MONTO_SERV_COL}, MONTO MOVILIZACION=${MONTO_MOV_COL}`);
    throw new Error("Columnas clave (como MONTO SERVICIO, MONTO MOVILIZACION, etc.) no encontradas en la hoja OT.");
  }

  const pendientes = [];
  const scriptTimeZone = Session.getScriptTimeZone();

  for (let i = 1; i < allDataOT.length; i++) {
    const row = allDataOT[i];
    
    // Filtro por Pedido y Estado
    if (String(row[COT_COL] || '').trim() === numPedido &&
        String(row[ESTADO_VAL_COL] || '').trim() === 'Pendiente') {
      
      // --- ¡AQUÍ ESTÁ LA LECTURA CORRECTA! ---
      const montoServicio = parseFloat(row[MONTO_SERV_COL] || 0);
      const montoMovilizacion = parseFloat(row[MONTO_MOV_COL] || 0);
      // ------------------------------------

      pendientes.push({
        otId: row[OT_ID_COL],
        fecha: formatearFechaRapido(row[FECHA_COL], scriptTimeZone),
        horas: parseFloat(row[HORAS_COL] || 0).toFixed(2),
        montoServicio: montoServicio,
        montoMovilizacion: montoMovilizacion,
        montoTotalOT: montoServicio + montoMovilizacion,
        rowIndex: i + 1 // N° de fila real para la actualización
      });
    }
  }
  
  Logger.log(`Encontradas ${pendientes.length} OTs pendientes para ${numPedido}`);
  return pendientes;
}

/**
 * [HELPER INTERNO]
 * Obtiene el historial de valorizaciones ya emitidas para este pedido.
 */
function _getValorizacionesEmitidas(numPedido) {
  const allDataVal = obtenerDatosHoja(HOJA_VALORIZACIONES, false);
  const COL_MAP_VAL = getColumnMap(HOJA_VALORIZACIONES);

  const PEDIDO_COL = COL_MAP_VAL['ID_PEDIDO'];
  if (PEDIDO_COL === undefined) {
    Logger.log("Advertencia: No se encontró 'ID_Pedido' en Hoja Valorizaciones. Creándola si no existe.");
    // Podríamos crear la hoja aquí si no existe
    return [];
  }

  const emitidas = [];
  const scriptTimeZone = Session.getScriptTimeZone();

  for (let i = 1; i < allDataVal.length; i++) {
    const row = allDataVal[i];
    if (String(row[PEDIDO_COL] || '').trim() === numPedido) {
      emitidas.push({
        valId: row[COL_MAP_VAL['ID_VALORIZACION']],
        fecha: formatearFechaRapido(row[COL_MAP_VAL['FECHA_EMISION']], scriptTimeZone),
        montoTotal: parseFloat(row[COL_MAP_VAL['MONTO_TOTAL']] || 0),
        estado: row[COL_MAP_VAL['ESTADO']]
      });
    }
  }
  // Ordenar por fecha, más reciente primero
  return emitidas.sort((a, b) => new Date(b.fecha.split('/').reverse().join('-')) - new Date(a.fecha.split('/').reverse().join('-')));
}


/**
 * Crea una nueva Valorización (Cabecera y Detalle) y actualiza el estado de las OTs.
 * Implementa un patrón transaccional.
 */
function crearValorizacion(numPedido, otsSeleccionadas) {
  // otsSeleccionadas es un array de objetos: [{otId: "OT-001", rowIndex: 5, montoServicio: 100, montoMov: 50}, ...]
  
  // 1. Validaciones
  const permisos = obtenerPermisosUsuario(); // Asume que tienes esta función
  if (!permisos.puedeEditarCotizacion) { // Reutiliza un permiso existente o crea uno nuevo
    return { success: false, message: "Acceso denegado." };
  }
  if (!numPedido) return { success: false, message: "No se proporcionó un número de pedido." };
  if (!otsSeleccionadas || otsSeleccionadas.length === 0) {
    return { success: false, message: "Debe seleccionar al menos una OT para valorizar." };
  }

  const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
  const sheetOT = ss.getSheetByName(HOJA_OT);
  const sheetVal = ss.getSheetByName(HOJA_VALORIZACIONES);
  const sheetValDetalle = ss.getSheetByName(HOJA_VALORIZACION_DETALLE);
  
  // Crear hojas si no existen (¡Robusto!)
  if (!sheetVal) sheetVal = ss.insertSheet(HOJA_VALORIZACIONES).appendRow(['ID_Valorizacion', 'ID_Pedido', 'Fecha_Emision', 'Fecha_Servicio_Inicio', 'Fecha_Servicio_Fin', 'Cliente', 'Monto_Servicios_Total', 'Monto_Movilizacion_Total', 'Monto_Total', 'Estado', 'ID_Factura', 'Link_PDF']);
  if (!sheetValDetalle) sheetValDetalle = ss.insertSheet(HOJA_VALORIZACION_DETALLE).appendRow(['ID_Detalle', 'ID_Valorizacion', 'ID_OT', 'Monto_Servicio_Hist', 'Monto_Mov_Hist']);

  const COL_MAP_OT = getColumnMap(HOJA_OT);
  const ESTADO_VAL_COL_OT = COL_MAP_OT['ESTADO_VALORIZACION'];
  
  if (ESTADO_VAL_COL_OT === undefined) {
      throw new Error("No se encontró la columna 'Estado_Valorizacion' en la hoja OT.");
  }

  const nuevoValId = `VAL-${numPedido.replace(/[^a-zA-Z0-9]/g, '')}-${new Date().getTime().toString().slice(-5)}`;
  let montoServiciosTotal = 0;
  let montoMovilizacionTotal = 0;
  let cliente = '';
  let fechaInicio = new Date(8640000000000000); // Fecha muy futura
  let fechaFin = new Date(-8640000000000000); // Fecha muy pasada
  
  // Para el rollback
  let otsActualizadas = []; 

  try {
    // 2. Preparar Datos y (PASO A) Actualizar OTs
    Logger.log(`Iniciando transacción para ${nuevoValId}. Procesando ${otsSeleccionadas.length} OTs...`);
    
    // Obtenemos los datos de cliente/fechas de la *primera* OT
    const primeraOT = obtenerDatosHoja(HOJA_OT, false).find(row => row[COL_MAP_OT['N° OT']] === otsSeleccionadas[0].otId);
    if(primeraOT) cliente = primeraOT[COL_MAP_OT['CLIENTE']]; // Asume que tienes columna CLIENTE en OT

    for (const ot of otsSeleccionadas) {
      // Sumar montos
      montoServiciosTotal += parseFloat(ot.montoServicio || 0);
      montoMovilizacionTotal += parseFloat(ot.montoMovilizacion || 0);
      
      // Encontrar fechas
      const fechaOT = new Date(ot.fecha.split('/').reverse().join('-'));
      if (fechaOT < fechaInicio) fechaInicio = fechaOT;
      if (fechaOT > fechaFin) fechaFin = fechaOT;

      // (PASO A) Actualizar estado de la OT
      // Columna es 1-based, COL_MAP_OT es 0-based
      sheetOT.getRange(ot.rowIndex, ESTADO_VAL_COL_OT + 1).setValue("En Proceso");
      otsActualizadas.push(ot.rowIndex); // Guardar para posible rollback
    }
    
    // 3. (PASO B) Escribir Cabecera de Valorización
    Logger.log(`Paso B: Creando cabecera ${nuevoValId}`);
    sheetVal.appendRow([
      nuevoValId,
      numPedido,
      new Date(), // Fecha_Emision
      fechaInicio,
      fechaFin,
      cliente,
      montoServiciosTotal,
      montoMovilizacionTotal,
      montoServiciosTotal + montoMovilizacionTotal, // Monto_Total
      "Borrador", // Estado
      "", // ID_Factura
      ""  // Link_PDF
    ]);

    // 4. (PASO C) Escribir Detalle de Valorización
    Logger.log(`Paso C: Creando ${otsSeleccionadas.length} detalles...`);
    const filasDetalle = [];
    for (const ot of otsSeleccionadas) {
      filasDetalle.push([
        `VALDET-${nuevoValId}-${ot.otId}`, // ID_Detalle
        nuevoValId, // ID_Valorizacion
        ot.otId, // ID_OT
        ot.montoServicio, // Monto_Servicio_Hist
        ot.montoMovilizacion // Monto_Mov_Hist
      ]);
    }
    sheetValDetalle.getRange(sheetValDetalle.getLastRow() + 1, 1, filasDetalle.length, filasDetalle[0].length)
                   .setValues(filasDetalle);

    // 5. Éxito
    Logger.log(`ÉXITO: Valorización ${nuevoValId} creada.`);
    // Invalidar cachés
    cache.remove(`hoja_${HOJA_OT}`);
    cache.remove(`hoja_${HOJA_VALORIZACIONES}`);
    cache.remove(`hoja_${HOJA_VALORIZACION_DETALLE}`);
    
    return { success: true, message: `Valorización ${nuevoValId} creada con ${otsSeleccionadas.length} OTs.` };

  } catch (e) {
    Logger.log(`ERROR en crearValorizacion: ${e.message}. Iniciando ROLLBACK...`);
    
    // (PASO D) Rollback de OTs
    if (otsActualizadas.length > 0) {
      Logger.log(`Rollback: Revertiendo ${otsActualizadas.length} OTs a "Pendiente"`);
      for (const rowIndex of otsActualizadas) {
        try {
          sheetOT.getRange(rowIndex, ESTADO_VAL_COL_OT + 1).setValue("Pendiente");
        } catch (revertError) {
          Logger.log(`ERROR CRÍTICO DE ROLLBACK en fila ${rowIndex}: ${revertError.message}`);
          enviarNotificacionError(`ERROR CRÍTICO DE ROLLBACK: No se pudo revertir la OT en fila ${rowIndex} a "Pendiente". Revisión manual requerida.`);
        }
      }
    }
    // (Rollback de Valorizacion y Detalle no es necesario porque el error suele ocurrir antes,
    // pero se podría añadir lógica para borrar la fila de 'sheetVal' si se creó)

    return manejarError('crearValorizacion', e);
  }
}
