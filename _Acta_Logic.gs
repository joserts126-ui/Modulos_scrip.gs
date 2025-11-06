/**
 * ====================================================
 * === LÓGICA DE NEGOCIO DEL MÓDULO DE ACTAS        ===
 * === (PLANIFICACIÓN - Misión 3 Revisada v2)       ===
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
      if (String(row[PEDIDO_COL] || '').trim().toUpperCase() === numPedido.toUpperCase()) {
        historial.push({
          actaId: row[COL_MAP_ACTAS['ID_ACTA']],
          fecha: formatearFechaRapido(row[COL_MAP_ACTAS['FECHA_EMISION']], scriptTimeZone),
          emitidaPor: row[COL_MAP_ACTAS['EMITIDA_POR']],
          link: row[COL_MAP_ACTAS['LINK_DOCUMENTO']],
          version: row[COL_MAP_ACTAS['VERSION']]
        });
      }
    }
    // Ordenar por versión (V1, V2, etc.)
    return { 
        success: true, 
        historial: historial.sort((a, b) => (a.version || '').localeCompare(b.version || ''))
    };

  } catch (e) {
    Logger.log(`Error en getHistorialDeActas: ${e.message}`);
    return manejarError('getHistorialDeActas', e);
  }
}

/**
 * Genera una nueva Hoja de Cálculo de Acta de Planificación basado en el Pedido.
 */
function generarActaPlanificacion(numPedido) {
  const usuario = obtenerEmailSeguro();
  Logger.log(`Iniciando generación de Acta (Sheets) para ${numPedido} por ${usuario}`);

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

    // 3. Seleccionar Plantilla de HOJA DE CÁLCULO correcta
    const empresa = (datosPedidoParaPlantilla.Empresa || '').toUpperCase(); // Usamos 'Empresa' de datosPedidoParaPlantilla
    let plantillaId;
    switch (empresa) {
        case 'ALPAMAYO':
            plantillaId = ID_PLANTILLA_ACTA_ALP;
            break;
        case 'GYM':
            plantillaId = ID_PLANTILLA_ACTA_GYM;
            break;
        case 'SAN JOSE':
            plantillaId = ID_PLANTILLA_ACTA_GSJ;
            break;
        default:
             Logger.log(`Empresa "${empresa}" no reconocida. Usando plantilla ALPAMAYO por defecto.`);
             plantillaId = ID_PLANTILLA_ACTA_ALP;
    }
    
    if (!plantillaId || plantillaId.includes("TU_ID_DE_PLANTILLA")) {
        throw new Error("Error Crítico: El ID de la plantilla de Acta no está configurado en _Constants_Lists.gs.");
    }

    // 4. Obtener Carpeta de Destino
    // Usamos 'datosPedidoParaCarpeta' que contiene el array 'servicios'.
    const carpetaCot = getDestinationFolder(
        datosPedidoParaCarpeta.ejecutivo, 
        datosPedidoParaCarpeta.empresa, 
        fechaEmision, 
        datosPedidoParaCarpeta // Este objeto SÍ tiene '.servicios'
    );
    const carpetaPedido = carpetaCot.getParents().next(); // Sube a la carpeta del Pedido (Nivel 4)
    const carpetaActas = findOrCreateFolder(carpetaPedido, "Actas de Planificacion"); // Crea subcarpeta

    // 5. Copiar Plantilla (Spreadsheet.copy)
    const plantillaSS = SpreadsheetApp.openById(plantillaId);
    const nombreArchivo = `Acta de Planificación - ${numPedido} - ${nuevaVersion}`;
    const nuevaHojaId = plantillaSS.copy(nombreArchivo).getId(); // Copia el *archivo* completo
    
    // Mover el archivo a la carpeta correcta
    const nuevoArchivo = DriveApp.getFileById(nuevaHojaId);
    carpetaActas.addFile(nuevoArchivo);
    DriveApp.getRootFolder().removeFile(nuevoArchivo); // Limpiar la raíz

    // 6. Rellenar la Hoja de Cálculo copiada
    const nuevoSS = SpreadsheetApp.openById(nuevaHojaId);
    // Asumimos que la plantilla tiene una sola hoja, o la primera es la que se usa.
    const sheet = nuevoSS.getSheets()[0]; 

    // --- Lógica de Reemplazo en Celdas (Placeholders) ---
    // Esto es mucho más rápido y robusto que replaceText en GDocs
    // Busca y reemplaza en *toda* la hoja.
    const textFinder = sheet.createTextFinder("{{PEDIDO}}");
    textFinder.replaceAllWith(numPedido);
    
    sheet.createTextFinder("{{FECHA_EMISION}}").replaceAllWith(formatearFechaRapido(fechaEmision, Session.getScriptTimeZone()));
    sheet.createTextFinder("{{EJECUTIVO}}").replaceAllWith(datosPedidoParaPlantilla.Ejecutivo || '');
    sheet.createTextFinder("{{CLIENTE}}").replaceAllWith(datosPedidoParaPlantilla.Cliente || '');
    sheet.createTextFinder("{{RUC}}").replaceAllWith(datosPedidoParaPlantilla.RUC || '');
    sheet.createTextFinder("{{CONTACTO}}").replaceAllWith(datosPedidoParaPlantilla.Contacto || '');
    sheet.createTextFinder("{{LUGAR}}").replaceAllWith(datosPedidoParaPlantilla.Direccion || '');
    sheet.createTextFinder("{{ACLARACIONES}}").replaceAllWith(datosPedidoParaPlantilla.aclaracionesServicio || 'Ninguna.');

    // --- Lógica para Rellenar la Tabla de Servicios ---
    // (Asumimos que la plantilla tiene placeholders {{ITEM_1}}, {{COD_1}}, {{DESC_1}}, etc.)
    // O, mejor, un rango con nombre o una celda de inicio.
    // Usaremos una celda de inicio (ej. A10)
    
    // **TU TAREA:** Define esta celda en tu plantilla. 
    // Sugiero poner el texto "{{INICIO_TABLA}}" en la celda A10.
    const celdaInicioTabla = sheet.createTextFinder("{{INICIO_TABLA}}").findNext();
    
    if (celdaInicioTabla) {
        const filaInicio = celdaInicioTabla.getRow();
        const colInicio = celdaInicioTabla.getColumn();
        
        // Leemos de 'datosPedidoParaPlantilla.Lineas'
        const lineasServicio = datosPedidoParaPlantilla.Lineas || [];
        
        if (lineasServicio.length > 0) {
            // Crear el array de datos para la tabla (Item, Código, Descripción, Cant., Unidad)
            const datosTabla = lineasServicio.map((linea, index) => {
                return [
                    index + 1,
                    linea.cod,
                    linea.descripcion,
                    linea.cantidad,
                    linea.und_medida
                ];
            });
            
            // Escribir todos los datos de la tabla de una sola vez
            sheet.getRange(filaInicio, colInicio, datosTabla.length, datosTabla[0].length)
                 .setValues(datosTabla);
                 
            // Limpiar el placeholder
            celdaInicioTabla.setValue(""); 
        } else {
             celdaInicioTabla.setValue("No hay líneas de servicio registradas.");
        }
    } else {
        Logger.log("Advertencia: No se encontró la celda con el placeholder {{INICIO_TABLA}}.");
    }

    SpreadsheetApp.flush(); // Guardar cambios

    // 7. Registrar en Hoja: Actas
    const sheetActasRegistro = ss.getSheetByName(HOJA_ACTAS);
    sheetActasRegistro.appendRow([
      nuevoActaId,
      numPedido,
      fechaEmision,
      usuario,
      nuevoArchivo.getUrl(), // URL de la *Hoja de Cálculo*
      nuevaVersion
    ]);
    
    cache.remove(`hoja_${HOJA_ACTAS}`); // Invalidar caché

    // 8. Devolver éxito
    return { success: true, link: nuevoArchivo.getUrl(), id: nuevoActaId };

  } catch (e) {
    Logger.log(`ERROR en generarActaPlanificacion (Sheets): ${e.message}\nStack: ${e.stack}`);
    return manejarError('generarActaPlanificacion', e);
  }
}


// --- FUNCIONES DE AYUDA PARA ROBUSTEZ ---

/**
 * Crea la hoja de registro de Actas si no existe.
 */
function _crearHojaActasSiNoExiste(ss) {
  const sheetName = HOJA_ACTAS;
  let sheet = ss ? ss.getSheetByName(sheetName) : SpreadsheetApp.openById(HOJA_ID_PRINCIPAL).getSheetByName(sheetName);
  
  if (!sheet) {
    if (!ss) ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
    sheet = ss.insertSheet(sheetName);
    // Columnas según Misión 3 Revisada
    sheet.appendRow([
      'ID_Acta', 'ID_Pedido', 'Fecha_Emision', 'Emitida_Por', 'Link_Documento', 'Version'
    ]);
    Logger.log(`Hoja ${sheetName} creada exitosamente.`);
    // Asegurarse de que el mapa de columnas se refresque
    cache.remove(`header_map_V5_${sheetName}`);
  }
  return sheet;
}
