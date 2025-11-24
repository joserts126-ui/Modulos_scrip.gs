/**
 * ====================================================
 * === LÓGICA DE NEGOCIO DEL MÓDULO DE ACTAS        ===
 * === (100% MIGRADO A SUPABASE)                    ===
 * ====================================================
 */

/**
 * Función principal llamada por Acta.html al cargar.
 * Obtiene el historial de Actas para un pedido específico desde Supabase.
 * Nota: Usa la tabla Valorizaciones para simular el historial de Actas.
 */
function getHistorialDeActas(numPedido) {
    try {
        if (!numPedido) throw new Error("Se requiere un número de pedido.");

        // Asumimos que Actas se gestiona a través de la tabla Valorizaciones
        const params = `Cot=eq.${numPedido}&select=Val_ID, Fecha_Emision, Estado, Monto_Total&order=Fecha_Emision.desc`; 
        
        const resultado = supabaseFetch('Valorizaciones', {
            method: 'get',
            params: params
        });

        // Mapear al formato que el frontend Acta.html espera (simulando versión/link)
        const historial = (resultado || []).map((acta, index) => ({
          actaId: acta.Val_ID,
          fecha: formatearFechaRapido(acta.Fecha_Emision, Session.getScriptTimeZone()),
          emitidaPor: acta.Estado || 'Emitida', 
          link: '#', 
          version: `V${resultado.length - index}` 
        }));

        return { 
            success: true, 
            historial: historial
        };
        
    } catch (e) {
        Logger.log(`Error en getHistorialDeActas (Supabase): ${e.message}`);
        return { success: false, message: e.message || 'Error al cargar historial.' };
    }
}

/**
 * Genera una nueva Hoja de Cálculo de Acta de Planificación basado en el Pedido,
 * utilizando datos 100% de Supabase.
 */
function generarActaPlanificacion(numPedido) {
    const usuario = obtenerEmailSeguro();
    Logger.log(`Iniciando generación de Acta (Sheets) para ${numPedido} por ${usuario}`);

    try {
        // 1. Obtener Datos del Pedido (Una única llamada a la función migrada de Supabase)
        const datosPedido = obtenerPedidoParaEdicion(numPedido); 
        
        if (!datosPedido || !datosPedido.success) {
          throw new Error(`No se pudieron obtener los datos del pedido: ${datosPedido.message || 'Error desconocido'}`);
        }

        const datosPedidoParaPlantilla = datosPedido;
        const datosPedidoParaCarpeta = datosPedido; 

        // 2. Determinar Versión
        const historial = getHistorialDeActas(numPedido).historial || [];
        const nuevaVersion = `V${historial.length + 1}`;
        const nuevoActaId = `ACTA-PLAN-${numPedido.replace(/[^A-Z0-9]/g, '')}-${nuevaVersion}`;
        const fechaEmision = new Date();

        // 3. Seleccionar Plantilla de HOJA DE CÁLCULO correcta
        const empresa = (datosPedidoParaPlantilla.Empresa || '').toUpperCase(); 
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
        const carpetaCot = getDestinationFolder(
            datosPedidoParaCarpeta.Ejecutivo, 
            datosPedidoParaCarpeta.Empresa, 
            fechaEmision, 
            datosPedidoParaCarpeta
        );
        const carpetaPedido = carpetaCot.getParents().next(); 
        const carpetaActas = findOrCreateFolder(carpetaPedido, "Actas de Planificacion"); 

        // 5. Copiar Plantilla (Spreadsheet.copy)
        const plantillaSS = SpreadsheetApp.openById(plantillaId);
        const nombreArchivo = `Acta de Planificación - ${numPedido} - ${nuevaVersion}`;
        const nuevaHojaId = plantillaSS.copy(nombreArchivo).getId(); 
        
        // Mover el archivo a la carpeta correcta
        const nuevoArchivo = DriveApp.getFileById(nuevaHojaId);
        carpetaActas.addFile(nuevoArchivo);
        DriveApp.getRootFolder().removeFile(nuevoArchivo); 

        // 6. Rellenar la Hoja de Cálculo copiada
        const nuevoSS = SpreadsheetApp.openById(nuevaHojaId);
        const sheet = nuevoSS.getSheets()[0]; 

        // --- Lógica de Reemplazo en Celdas (Placeholders) ---
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
        const celdaInicioTabla = sheet.createTextFinder("{{INICIO_TABLA}}").findNext();
        
        if (celdaInicioTabla) {
            const filaInicio = celdaInicioTabla.getRow();
            const colInicio = celdaInicioTabla.getColumn();
            
            const lineasServicio = datosPedidoParaPlantilla.Lineas || [];
            
            if (lineasServicio.length > 0) {
                const datosTabla = lineasServicio.map((linea, index) => {
                    return [
                        index + 1,
                        linea.cod,
                        linea.descripcion,
                        linea.cantidad,
                        linea.und_medida
                    ];
                });
                
                sheet.getRange(filaInicio, colInicio, datosTabla.length, datosTabla[0].length)
                     .setValues(datosTabla);
                     
                celdaInicioTabla.setValue(""); 
            } else {
                 celdaInicioTabla.setValue("No hay líneas de servicio registradas.");
            }
        } else {
            Logger.log("Advertencia: No se encontró la celda con el placeholder {{INICIO_TABLA}}.");
        }

        SpreadsheetApp.flush(); 

        // 7. Registrar en Supabase: Actas
        const registroActa = {
            "ID_Acta": nuevoActaId,
            "Cot": numPedido, 
            "Fecha_Emision": fechaEmision.toISOString(),
            "Emitida_Por": usuario,
            "Link_Documento": nuevoArchivo.getUrl(), 
            "Version": nuevaVersion
        };
        
        try {
            supabaseFetch('Actas_Planificacion', {
                method: 'post',
                payload: registroActa
            });
            Logger.log(`Registro de Acta ${nuevoActaId} guardado en Supabase.`);

        } catch (e) {
            // El fallback a Sheets ha sido eliminado por completo.
            Logger.log(`ADVERTENCIA/ERROR: Falló el registro del Acta en Supabase. El acta se generó en Drive, pero el registro de metadatos falló. Error: ${e.message}`);
        }

        // 8. Devolver éxito
        return { success: true, link: nuevoArchivo.getUrl(), id: nuevoActaId };

    } catch (e) {
        Logger.log(`ERROR en generarActaPlanificacion (Sheets): ${e.message}\nStack: ${e.stack}`);
        return manejarError('generarActaPlanificacion', e);
    }
}


// --- FUNCIONES DE ACCESO AL CRUD DE ACTAS ---

/**
 * Verifica si ya existe un Acta registrada para un Pedido y Línea.
 * @returns {string|null} El ActaID si existe, null si no.
 */
function verificarActaExistente(cotNumero, cotLineaIndex) {
    try {
        const params = `COT_Numero=eq.${cotNumero}&COT_LineaIndex=eq.${cotLineaIndex}&select=ActaID`;

        const resultado = supabaseFetch('Actas_Planificacion', {
            method: 'get',
            params: params
        });

        if (resultado && resultado.length > 0) {
            return resultado[0].ActaID; 
        }
        return null;

    } catch (e) {
        Logger.log(`Error en verificarActaExistente (Supabase): ${e.message}`);
        return null;
    }
}
