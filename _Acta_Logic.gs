/**
 * ====================================================
 * === LÓGICA DE NEGOCIO DEL MÓDULO DE ACTAS        ===
 * === (PLANIFICACIÓN - Misión 3 Revisada)          ===
 * ====================================================
 */

/**
 * Función principal llamada por Acta.html al cargar.
 * Obtiene el historial de actas generadas para un pedido.
 */
function getHistorialDeActas(numPedido) {
  try {
    if (!numPedido) throw new Error("Se requiere un número de pedido.");

    _crearHojaActasSiNoExiste(); // Asegura que la hoja exista
    const allDataActas = obtenerDatosHoja(HOJA_ACTAS, false); // No caché
    const COL_MAP_ACTAS = getColumnMap(HOJA_ACTAS);

    const PEDIDO_COL = COL_MAP_ACTAS['ID_PEDIDO'];
    if (PEDIDO_COL === undefined) {
      throw new Error("Columna 'ID_Pedido' no encontrada en Hoja: Actas.");
    }

    const historial = [];
    const scriptTimeZone = Session.getScriptTimeZone();

    for (let i = 1; i < allDataActas.length; i++) {
      const row = allDataActas[i];
      if (String(row[PEDIDO_COL] || '').trim() === numPedido) {
        historial.push({
          actaId: row[COL_MAP_ACTAS['ID_ACTA']],
          fecha: formatearFechaRapido(row[COL_MAP_ACTAS['FECHA_EMISION']], scriptTimeZone),
          emitidaPor: row[COL_MAP_ACTAS['EMITIDA_POR']],
          link: row[COL_MAP_ACTAS['LINK_DOCUMENTO']],
          version: row[COL_MAP_ACTAS['VERSION']]
        });
      }
    }
    // Ordenar por versión o fecha, más reciente primero
    return { 
        success: true, 
        historial: historial.sort((a, b) => (b.version || '').localeCompare(a.version || ''))
    };

  } catch (e) {
    Logger.log(`Error en getHistorialDeActas: ${e.message}`);
    return manejarError('getHistorialDeActas', e);
  }
}

/**
 * REEMPLAZO TOTAL (Misión 3 - Corregida v2)
 * Genera un nuevo Google Doc de Acta de Planificación basado en el Pedido.
 * CORREGIDO: Llama a DOS funciones de obtención de datos:
 * 1. obtenerDetallesCompletosDePedido: para el array 'servicios' que necesita 'getDestinationFolder'.
 * 2. obtenerPedidoParaEdicion: para el array 'Lineas' y 'aclaracionesServicio' que necesita la plantilla.
 */
function generarActaPlanificacion(numPedido) {
  const usuario = obtenerEmailSeguro();
  Logger.log(`Iniciando generación de Acta (v2) para ${numPedido} por ${usuario}`);

  try {
    // 1. Obtener Datos del Pedido (EN DOS PARTES)
    
    // PASO 1.A: Llamar a 'obtenerDetallesCompletosDePedido'
    // Esta función nos da los datos de cabecera y el array 'servicios'
    // que 'getDestinationFolder' necesita.
    const datosPedidoParaCarpeta = obtenerDetallesCompletosDePedido(numPedido); // de _Comercial_Logic.gs
    if (!datosPedidoParaCarpeta || !datosPedidoParaCarpeta.success) {
      throw new Error(`No se pudieron obtener los datos (parte 1) del pedido: ${datosPedidoParaCarpeta.message || 'Error desconocido'}`);
    }

    // PASO 1.B: Llamar a 'obtenerPedidoParaEdicion'
    // Esta función nos da el array 'Lineas' (para la tabla) y 'aclaracionesServicio'.
    const datosPedidoParaPlantilla = obtenerPedidoParaEdicion(numPedido); // de _Comercial_Logic.gs
    if (!datosPedidoParaPlantilla || !datosPedidoParaPlantilla.success) {
      throw new Error(`No se pudieron obtener los datos (parte 2) del pedido: ${datosPedidoParaPlantilla.message || 'Error desconocido'}`);
    }

    // 2. Determinar Versión
    const historial = getHistorialDeActas(numPedido).historial || [];
    const nuevaVersion = `V${historial.length + 1}`;
    const nuevoActaId = `ACTA-PLAN-${numPedido.replace(/[^A-Z0-9]/g, '')}-${nuevaVersion}`;
    const fechaEmision = new Date();

    // 3. Obtener Carpeta de Destino
    // ¡CORRECCIÓN! Usamos 'datosPedidoParaCarpeta' que contiene el array 'servicios'.
    const carpetaCot = getDestinationFolder(
        datosPedidoParaCarpeta.ejecutivo, 
        datosPedidoParaCarpeta.empresa, 
        fechaEmision, 
        datosPedidoParaCarpeta // Este objeto SÍ tiene '.servicios'
    );
    const carpetaPedido = carpetaCot.getParents().next(); // Sube a la carpeta del Pedido (Nivel 4)
    const carpetaActas = findOrCreateFolder(carpetaPedido, "Actas de Planificacion"); // Crea subcarpeta

    // 4. Copiar Plantilla
    const plantillaDoc = DriveApp.getFileById(ID_PLANTILLA_ACTA_PLANIFICACION); // de _Constants_Lists.gs
    const nombreArchivo = `Acta de Planificación - ${numPedido} - ${nuevaVersion}`;
    const nuevoArchivo = plantillaDoc.makeCopy(carpetaActas, nombreArchivo);
    const nuevoDoc = DocumentApp.openById(nuevoArchivo.getId());
    
    // 5. Rellenar Placeholders Simples
    // Usamos 'datosPedidoParaPlantilla' que tiene los nombres de propiedad correctos (ej. 'Cliente' con mayúscula)
    const body = nuevoDoc.getBody();
    body.replaceText("{{PEDIDO}}", numPedido);
    body.replaceText("{{FECHA_EMISION}}", formatearFechaRapido(fechaEmision, Session.getScriptTimeZone()));
    body.replaceText("{{EJECUTIVO}}", datosPedidoParaPlantilla.Ejecutivo || '');
    body.replaceText("{{CLIENTE}}", datosPedidoParaPlantilla.Cliente || '');
    body.replaceText("{{RUC}}", datosPedidoParaPlantilla.RUC || '');
    body.replaceText("{{CONTACTO}}", datosPedidoParaPlantilla.Contacto || '');
    body.replaceText("{{LUGAR}}", datosPedidoParaPlantilla.Direccion || ''); // Usa 'Direccion' para 'Lugar'
    body.replaceText("{{ACLARACIONES}}", datosPedidoParaPlantilla.aclaracionesServicio || 'Ninguna.');

    // 6. Rellenar Tabla de Servicios
    const placeholderTabla = body.findText("{{TABLA_SERVICIOS}}");
    if (placeholderTabla) {
      const elemento = placeholderTabla.getElement();
      const parrafo = elemento.getParent();
      const indiceParrafo = body.getChildIndex(parrafo);

      // Crear la tabla
      const tabla = body.insertTable(indiceParrafo, [['Item', 'Código', 'Descripción', 'Cant.', 'Unidad']]);
      tabla.getRow(0).editAsText().setBold(true).setBackgroundColor("#EEEEEE");
      
      // ¡CORRECCIÓN! Leemos de 'datosPedidoParaPlantilla.Lineas'
      const lineasServicio = datosPedidoParaPlantilla.Lineas || [];
      lineasServicio.forEach((linea, index) => {
        const fila = tabla.appendRow([
          (index + 1).toString(),
          linea.cod || '',
          linea.descripcion || '',
          (linea.cantidad || 0).toString(),
          linea.und_medida || ''
        ]);
        fila.editAsText().setBold(false);
      });
      
      // Eliminar el párrafo del placeholder
      parrafo.removeFromParent();

    } else {
      Logger.log(`Advertencia: No se encontró el placeholder {{TABLA_SERVICIOS}} en la plantilla.`);
    }

    // 7. Guardar y Cerrar
    nuevoDoc.saveAndClose();
    Logger.log(`Documento ${nuevoArchivo.getName()} creado exitosamente.`);
    
    // 8. Registrar en Hoja: Actas
    const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
    const sheetActas = _crearHojaActasSiNoExiste(ss); // Usamos el helper para asegurar
    sheetActas.appendRow([
      nuevoActaId,
      numPedido,
      fechaEmision,
      usuario,
      nuevoArchivo.getUrl(),
      nuevaVersion
    ]);

    // Invalidar caché
    cache.remove(`hoja_${HOJA_ACTAS}`);

    // 9. Devolver éxito
    return { success: true, link: nuevoArchivo.getUrl(), id: nuevoActaId };

  } catch (e) {
    Logger.log(`ERROR en generarActaPlanificacion (v2): ${e.message}\nStack: ${e.stack}`);
    // Usamos la función de manejo de errores centralizada
    return manejarError('generarActaPlanificacion', e); 
  }
}

// --- FUNCIONES DE AYUDA PARA ROBUSTEZ ---

function _crearHojaActasSiNoExiste(ss) {
  const sheetName = HOJA_ACTAS;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow([
      'ID_Acta', 'ID_Pedido', 'Fecha_Emision', 'Emitida_Por', 'Link_Documento', 'Version'
    ]);
    Logger.log(`Hoja ${sheetName} creada exitosamente.`);
    // Asegurarse de que el mapa de columnas se refresque
    cache.remove(`header_map_V5_${sheetName}`);
  }
  return sheet;
}
