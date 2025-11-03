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
 * Genera un nuevo Google Doc de Acta de Planificación basado en el Pedido.
 */
function generarActaPlanificacion(numPedido) {
  const usuario = obtenerEmailSeguro();
  Logger.log(`Iniciando generación de Acta para ${numPedido} por ${usuario}`);

  try {
    // 1. Obtener Datos del Pedido
    // REUTILIZAMOS la función de _Comercial_Logic.gs
    // Esta función ya devuelve un objeto { success: true, ...datos }
    const datosPedido = obtenerPedidoParaEdicion(numPedido);
    if (!datosPedido || !datosPedido.success) {
      throw new Error(`No se pudieron obtener los datos del pedido: ${datosPedido.message || 'Error desconocido'}`);
    }

    // 2. Determinar Versión
    const historial = getHistorialDeActas(numPedido).historial || [];
    const nuevaVersion = `V${historial.length + 1}`;
    const nuevoActaId = `ACTA-PLAN-${numPedido.replace(/[^A-Z0-9]/g, '')}-${nuevaVersion}`;
    const fechaEmision = new Date();

    // 3. Obtener Carpeta de Destino
    // Reutilizamos la lógica de _Comercial_Logic.gs para encontrar la carpeta del pedido
    // Nota: getDestinationFolder devuelve la subcarpeta "cot", así que subimos un nivel.
    const carpetaCot = getDestinationFolder(datosPedido.Ejecutivo, datosPedido.Empresa, fechaEmision, datosPedido);
    const carpetaPedido = carpetaCot.getParents().next(); // Sube a la carpeta del Pedido (Nivel 4)
    const carpetaActas = findOrCreateFolder(carpetaPedido, "Actas de Planificacion"); // Crea subcarpeta

    // 4. Copiar Plantilla
    const plantillaDoc = DriveApp.getFileById(ID_PLANTILLA_ACTA_PLANIFICACION);
    const nombreArchivo = `Acta de Planificación - ${numPedido} - ${nuevaVersion}`;
    const nuevoArchivo = plantillaDoc.makeCopy(carpetaActas, nombreArchivo);
    const nuevoDoc = DocumentApp.openById(nuevoArchivo.getId());
    
    // 5. Rellenar Placeholders Simples
    const body = nuevoDoc.getBody();
    body.replaceText("{{PEDIDO}}", numPedido);
    body.replaceText("{{FECHA_EMISION}}", formatearFechaRapido(fechaEmision, Session.getScriptTimeZone()));
    body.replaceText("{{EJECUTIVO}}", datosPedido.Ejecutivo || '');
    body.replaceText("{{CLIENTE}}", datosPedido.Cliente || '');
    body.replaceText("{{RUC}}", datosPedido.RUC || '');
    body.replaceText("{{CONTACTO}}", datosPedido.Contacto || '');
    body.replaceText("{{LUGAR}}", datosPedido.Direccion || ''); // Usa 'Direccion' para 'Lugar'
    body.replaceText("{{ACLARACIONES}}", datosPedido.aclaracionesServicio || 'Ninguna.');

    // 6. Rellenar Tabla de Servicios
    const placeholderTabla = body.findText("{{TABLA_SERVICIOS}}");
    if (placeholderTabla) {
      const elemento = placeholderTabla.getElement();
      const parrafo = elemento.getParent();
      const indiceParrafo = body.getChildIndex(parrafo);

      // Crear la tabla
      const tabla = body.insertTable(indiceParrafo, [['Item', 'Código', 'Descripción', 'Cant.', 'Unidad']]);
      // Aplicar estilo de cabecera
      tabla.getRow(0).editAsText().setBold(true).setBackgroundColor("#EEEEEE");
      
      const lineasServicio = datosPedido.Lineas || [];
      lineasServicio.forEach((linea, index) => {
        const fila = tabla.appendRow([
          index + 1,
          linea.cod,
          linea.descripcion,
          linea.cantidad,
          linea.und_medida
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
    const sheetActas = ss.getSheetByName(HOJA_ACTAS);
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
    Logger.log(`ERROR en generarActaPlanificacion: ${e.message}\nStack: ${e.stack}`);
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
