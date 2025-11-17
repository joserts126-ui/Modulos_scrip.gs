/**
 * ==========================================================
 * === LÓGICA DE VALORIZACIÓN (v2 - 100% Migrado a Supabase) ===
 * ==========================================================
 */

/**
 * Función principal llamada por Valorizacion.html al cargar.
 * Obtiene OTs pendientes y valorizaciones emitidas desde Supabase.
 */
function getDatosParaValorizar(numPedido) {
  try {
    if (!numPedido) throw new Error("Se requiere un número de pedido.");
    
    // Llama a los nuevos helpers de Supabase
    const pendientes = _getOTsPendientesSupabase(numPedido);
    const emitidas = _getValorizacionesEmitidasSupabase(numPedido);

    return {
      success: true,
      pendientes: pendientes,
      emitidas: emitidas
    };
  } catch (e) {
    Logger.log(`Error en getDatosParaValorizar (Supabase): ${e.message}`);
    return manejarError('getDatosParaValorizar', e);
  }
}

/**
 * [HELPER INTERNO - Supabase]
 * Obtiene todas las OTs de un pedido que están "Pendientes" de valorizar.
 */
function _getOTsPendientesSupabase(numPedido) {
  Logger.log(`Buscando OTs pendientes para: ${numPedido}`);
  try {
    // Esta consulta busca en Detalle_Pedidos por el 'Cot' y se une (JOIN)
    // a las Ordenes_Trabajo que estén 'Pendiente'.
    const consulta = `
      select=
        Cot_Linea_Ref,
        Ordenes_Trabajo!inner(
          N_OT,
          Fecha,
          Monto_Servicio,
          Monto_Movilizacion,
          Estado_Valorizacion
        )
      &Cot=eq.${numPedido}
      &Ordenes_Trabajo.Estado_Valorizacion=eq.Pendiente
    `.replace(/\s/g, '');

    const lineasConOTsPendientes = supabaseFetch('Detalle_Pedidos', {
      method: 'get',
      params: consulta
    });

    // Mapear el resultado anidado a la estructura plana que espera el frontend
    const pendientes = lineasConOTsPendientes.map(linea => {
      const ot = linea.Ordenes_Trabajo[0]; // Asumimos una OT pendiente por línea por ahora
      if (!ot) return null; // Seguridad

      const montoServicio = parseFloat(ot.Monto_Servicio || 0);
      const montoMovilizacion = parseFloat(ot.Monto_Movilizacion || 0);

      return {
        otId: ot.N_OT,
        fecha: ot.Fecha, // Viene como 'YYYY-MM-DD'
        montoServicio: montoServicio,
        montoMovilizacion: montoMovilizacion,
        montoTotalOT: montoServicio + montoMovilizacion,
        // ¡Importante! Necesitamos el ID numérico de la línea para la OT
        // pero la RPC de creación de Valorización no lo usa,
        // así que solo pasamos el ID de la OT
      };
    }).filter(ot => ot != null); // Filtrar nulos
    
    Logger.log(`Encontradas ${pendientes.length} OTs pendientes para ${numPedido}`);
    return pendientes;
    
  } catch (error) {
    Logger.log(`Error en _getOTsPendientesSupabase: ${error.message}`);
    throw new Error(`Error al buscar OTs pendientes: ${error.message}`);
  }
}

/**
 * [HELPER INTERNO - Supabase]
 * Obtiene el historial de valorizaciones ya emitidas para este pedido.
 */
function _getValorizacionesEmitidasSupabase(numPedido) {
  Logger.log(`Buscando valorizaciones emitidas para: ${numPedido}`);
  try {
    const consulta = `
      select=
        Val_ID,
        Fecha_Emision,
        Monto_Total,
        Estado
      &Cot=eq.${numPedido}
      &order=Fecha_Emision.desc
    `.replace(/\s/g, '');

    const emitidas = supabaseFetch('Valorizaciones', {
      method: 'get',
      params: consulta
    });

    // Mapear al formato simple que espera el frontend
    return emitidas.map(val => ({
      valId: val.Val_ID,
      fecha: formatearFechaRapido(val.Fecha_Emision, Session.getScriptTimeZone()),
      montoTotal: parseFloat(val.Monto_Total || 0),
      estado: val.Estado
    }));
    
  } catch (error) {
    Logger.log(`Error en _getValorizacionesEmitidasSupabase: ${error.message}`);
    throw new Error(`Error al buscar valorizaciones emitidas: ${error.message}`);
  }
}


/**
 * REEMPLAZO (v2 - Supabase RPC)
 * Llama a la RPC 'crear_nueva_valorizacion' para crear la valorización de forma segura.
 */
function crearValorizacion(numPedido, otsSeleccionadas) {
  const permisos = obtenerPermisosUsuario();
  if (!permisos.puedeEditarCotizacion) { // Reutiliza un permiso
    return { success: false, message: "Acceso denegado." };
  }

  if (!numPedido) return { success: false, message: "No se proporcionó un número de pedido." };
  if (!otsSeleccionadas || otsSeleccionadas.length === 0) {
    return { success: false, message: "Debe seleccionar al menos una OT para valorizar." };
  }

  Logger.log(`Iniciando RPC crear_nueva_valorizacion para ${numPedido} con ${otsSeleccionadas.length} OTs...`);
  
  try {
    // 1. Preparar el payload para la RPC
    // El frontend envía datos extra (rowIndex), los filtramos
    const payloadOTs = otsSeleccionadas.map(ot => ({
      "N_OT": ot.otId,
      "Monto_Servicio": ot.montoServicio,
      "Monto_Movilizacion": ot.montoMovilizacion,
      "Fecha": ot.fecha.split('/').reverse().join('-') // Convertir DD/MM/YYYY a YYYY-MM-DD
    }));

    const payloadRPC = {
      "codigo_pedido": numPedido,
      "ots_json": payloadOTs
    };
    
    // 2. Llamar a la RPC
    const resultado = supabaseFetch('rpc/crear_nueva_valorizacion', {
      method: 'post',
      payload: payloadRPC
    });

    // 3. Devolver el resultado de la RPC (que ya tiene formato {success: true, ...})
    Logger.log(`Resultado de la RPC: ${JSON.stringify(resultado)}`);
    return resultado;

  } catch (e) {
    Logger.log(`ERROR en crearValorizacion (RPC): ${e.message}`);
    return manejarError('crearValorizacion', e);
  }
}
