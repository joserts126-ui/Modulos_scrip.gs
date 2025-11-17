// ==================================================================================================
// === ESTE ARCHIVO CONTIENE LA LÓGICA DE NEGOCIO DEL MÓDULO COMERCIAL (EX MODULOS_SCRIP.GS) ===
// === NOTA: ESTE ARCHIVO DEPENDE DE QUE EXISTAN _Core.gs y _Constants_Lists.gs ===
// ==================================================================================================
const MODO_DEBUG_PDF = true;
// ====================================================
// === FUNCIONES DE PROCESAMIENTO AUXILIAR CRÍTICO ===
// ====================================================

/**
 * Normaliza el nombre del turno. Necesario para obtenerPedidoParaEdicion.
 */
function normalizarTurno(turno) {
    if (!turno || typeof turno !== 'string') return 'Diurno';
    const turnoUpper = turno.toUpperCase().trim();
    if (turnoUpper.includes('DIURNO')) return 'Diurno';
    if (turnoUpper.includes('NOCTURNO')) return 'Nocturno';
    if (turnoUpper.includes('DOBLE')) return 'Doble Turno';
    return 'Diurno';
}

/**
 * Formatea la fecha rápidamente (necesario para Resumen).
 */
function formatearFechaRapido(rawFecha, timeZone) {
    if (!rawFecha) return '';
    if (rawFecha instanceof Date) {
        return Utilities.formatDate(rawFecha, timeZone, 'dd/MM/yyyy');
    }
    try {
        const dateObj = new Date(rawFecha);
        if (!isNaN(dateObj)) {
             return Utilities.formatDate(dateObj, timeZone, 'dd/MM/yyyy');
        }
    } catch(e) {
        return String(rawFecha).substring(0, 10);
    }
    return String(rawFecha).substring(0, 10);
}

/**
 * Formatea un valor de fecha para un input HTML type="date" (yyyy-MM-dd).
 * Devuelve string vacío si la fecha es inválida o nula.
 */
function formatearParaInputDate(rawFecha) {
    if (!rawFecha) return '';
    let dateObj;
    if (rawFecha instanceof Date) {
        dateObj = rawFecha;
    } else {
        try {
            // Intentar parsear la fecha (string o número)
            dateObj = new Date(rawFecha);
        } catch (e) {
            return ''; // Inválido
        }
    }
    
    // Validar que el objeto de fecha sea válido
    if (dateObj && !isNaN(dateObj)) {
        try {
            // Usar Utilities.formatDate para obtener yyyy-MM-dd en la zona horaria del script
            // Esto evita errores de "un día antes/después" por UTC
            return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } catch (e) {
             // Fallback si Utilities.formatDate falla
             const y = dateObj.getFullYear();
             const m = String(dateObj.getMonth() + 1).padStart(2, '0');
             const d = String(dateObj.getDate()).padStart(2, '0');
             return `${y}-${m}-${d}`;
        }
    }
    return ''; // Fecha inválida
}

/**
 * Función unificada para obtener datos filtrados (usa crudHoja en _Core.gs)
 */
function obtenerDatosFiltrados(sheetName, filtro = {}) {
    return crudHoja('FILTER', sheetName, null, filtro);
}

/**
 * Función optimizada para búsqueda (usa obtenerDatosHoja en _Core.gs)
 */
function buscarRegistro(sheetName, criterio, columnaBusqueda = 0) {
    const allData = obtenerDatosHoja(sheetName);
    if (allData.length <= 1) return null;
    
    const criterioStr = String(criterio).trim();
    for (let i = 1; i < allData.length; i++) {
        const valor = String(allData[i][columnaBusqueda] || '').trim();
        if (valor === criterioStr) {
            return {
                datos: allData[i],
                indiceFila: i + 1,
                encabezados: allData[0]
            };
        }
    }
    return null;
}

// ====================================================
// === FUNCIONES DE PROCESAMIENTO AUXILIAR CRÍTICO ===
// === (Aseguran que la edición no se caiga) ===
// ====================================================

/**
 * Procesa el monto rápidamente (necesario para Resumen).
 */
function procesarMontoRapido(rawMonto) {
    if (typeof rawMonto === 'number') return rawMonto;
    if (!rawMonto) return 0;
    const strMonto = String(rawMonto);
    let numero = '';
    for (let i = 0; i < strMonto.length; i++) {
        const char = strMonto[i];
        if ((char >= '0' && char <= '9') || char === '.' || char === '-') {
            numero += char;
        }
    }
    return parseFloat(numero) || 0;
}

// ====================================================
// === FUNCIONES DE BÚSQUEDA Y FILTRO DE DATOS ===
// ====================================================

/**
 * Función unificada para obtener datos filtrados (usa crudHoja en _Core.gs)
 */
function obtenerDatosFiltrados(sheetName, filtro = {}) {
    // Esta función llama a crudHoja('FILTER') definido en _Core.gs
    return crudHoja('FILTER', sheetName, null, filtro);
}

/**
 * Función optimizada para búsqueda (usa obtenerDatosHoja en _Core.gs)
 */
function buscarRegistro(sheetName, criterio, columnaBusqueda = 0) {
    // obtenerDatosHoja se asume definido en _Core.gs
    const allData = obtenerDatosHoja(sheetName);
    if (allData.length <= 1) return null;
    
    const criterioStr = String(criterio).trim();
    for (let i = 1; i < allData.length; i++) {
        const valor = String(allData[i][columnaBusqueda] || '').trim();
        if (valor === criterioStr) {
            return {
                datos: allData[i],
                indiceFila: i + 1,
                encabezados: allData[0]
            };
        }
    }
    return null;
}

// ====================================================
// === CÁLCULO DE NEGOCIO (MOVIDO DE FRONTEND) ===
// ====================================================

/**
 * Calcula los días cotizados basándose en las unidades y horas mínimas.
 */
function calcularDiasCotizadosGS(cantidad, undMedida, hMinNum, hMinUnd) {
    let diasCotizados = 0;
    const undMedidaUpper = undMedida.toUpperCase().trim();
    const hMinUndUpper = hMinUnd.toUpperCase().trim();
    
    const necesitaTiempo = undMedidaUpper === 'HORAS' || undMedidaUpper === 'DÍAS';

    if (necesitaTiempo) {
        if (undMedidaUpper === 'DÍAS') {
            diasCotizados = cantidad;
        } else if (undMedidaUpper === 'HORAS') {
            let factorDias = 0;
            switch (hMinUndUpper) {
                case 'MENSUAL': factorDias = VALOR_DIAS_MES; break;
                case 'SEMANAL': factorDias = 7; break;
                case 'DIARIAS': factorDias = 1; break;
            }
            if (hMinNum > 0 && factorDias > 0) {
                diasCotizados = (cantidad / hMinNum) * factorDias;
            } else {
                diasCotizados = cantidad / VALOR_HORAS_DIA; 
            }
        }
    }
    
    return Math.ceil(diasCotizados);
}

// ====================================================
// === FUNCIONES DE OPERACIONES DE TRABAJO (OT) ===
// ====================================================

/**
 * Lógica para la obtención de datos iniciales del formulario OT (RegistrarOT.html).
 */
function getDatosInicialesOT() {
    try {
        return {
            success: true,
            clientes: obtenerDatosHoja(HOJA_CLIENTES),
            servicios: obtenerDatosHoja(HOJA_SERVICIOS)
            // 'pedidos' se elimina de aquí, se cargarán dinámicamente
        };
    } catch (e) {
        return manejarError('getDatosInicialesOT', e);
    }
}


// ====================================================
// === UTILIDADES DE DATOS SIMPLIFICADAS (USANDO _Core.gs) ===
// ====================================================

function getListaHoraSegun() {
    return getListaValoresUnicos(HOJA_SERVICIOS, 5);
}

function getListaUnidadesDeMedida() {
    return getListaValoresUnicos(HOJA_SERVICIOS, 6);
}

function buscarServicioPorCodigo(codServicio) {
    return buscarRegistro(HOJA_SERVICIOS, codServicio, 0); 
}

function getListaServicios() {
    return obtenerDatosHoja(HOJA_SERVICIOS);
}

// ====================================================
// === FUNCIONES INICIALIZACIÓN DE MÓDULO COMERCIAL ===
// ====================================================

function getDatosInicialesComercial() {
    try {
        // NOTA: Quitamos Clientes y Servicios de aquí
        return {
            success: true,
            formasPago: LISTA_FORMAS_PAGO, 
            listaEmpresas: LISTA_EMPRESAS,
            listaTurnos: LISTA_TURNOS,
            listaHorasMinimas: LISTA_HORAS_MINIMAS_UND,
            listaEjecutivos: LISTA_EJECUTIVOS,
            listaEstadosCot: LISTA_ESTADOS_COT
        };
    } catch (e) {
        return manejarError('getDatosInicialesComercial', e);
    }
}

function getDatosPesadosComercial() {
  try {
    Logger.log("Ejecutando getDatosPesadosComercial (Versión Supabase)...");

    // 1. Obtener Clientes (de la tabla "Clientes")
    const clientes = supabaseFetch('Clientes', {
      method: 'get',
      params: 'select=RUC_DNI,Nombre_RazonSocial' 
    });
    Logger.log(`Clientes cargados: ${clientes.length}`);

    // 2. Obtener Servicios (de la tabla "Servicios")
    const servicios = supabaseFetch('Servicios', {
      method: 'get',
      params: 'select=ID_servicios,Nombre_Servicio' 
    });
    Logger.log(`Servicios cargados: ${servicios.length}`);

    // 3. Devolver los datos
    // ¡CORRECCIÓN AQUÍ!
    // Ahora devolvemos las constantes estáticas que acabamos de definir
    return {
      success: true,
      clientes: clientes,     // Array de Objetos JSON
      servicios: servicios,    // Array de Objetos JSON
      listaUndMedida: LISTA_UND_MEDIDA,  // <-- Constante de _Constants_Lists.gs
      listaHorasSegun: LISTA_HORAS_SEGUN // <-- Constante de _Constants_Lists.gs
    };

  } catch (e) {
    Logger.log(`ERROR en getDatosPesadosComercial (Supabase): ${e.message}`);
    return manejarError('getDatosPesadosComercial', e);
  }
}

/**
 * REFACTORIZADO (v3 - Supabase)
 * Obtiene los contactos para un RUC específico desde la tabla "Contactos"
 * y "traduce" los nombres de las columnas para el frontend.
 */
function getContactosParaComercial(ruc) {
  try {
    if (!ruc) return []; // Devuelve un array vacío si no hay RUC

    Logger.log(`Ejecutando getContactosParaComercial (Supabase) para RUC: ${ruc}`);

    // 1. Definir la consulta a Supabase
    const tabla = 'Contactos';
    // Pide las columnas de 'Contactos_rows.csv'
    const columnas = 'Nombre_Contacto,Cargo,Correo,Celular'; 
    // Filtra por el RUC (usando el nombre de tu columna en Supabase)
    const filtro = `RUC_DNI=eq.${ruc}`;

    // 2. Usar el traductor
    const contactosDeSupabase = supabaseFetch(tabla, {
      method: 'get',
      params: `${filtro}&select=${columnas}`
    });

    Logger.log(`Se encontraron ${contactosDeSupabase.length} contactos.`);

    // 3. ¡TRADUCCIÓN!
    // El frontend (Comercial.html) espera: {nombre, cargo, email, telefono, display}
    // Supabase devuelve: {Nombre_Contacto, Cargo, Correo, Celular}
    
    const contactosTraducidos = contactosDeSupabase.map(contacto => {
      const nombre = contacto.Nombre_Contacto || '';
      const cargo = contacto.Cargo || '';
      
      return {
        nombre: nombre,
        cargo: cargo,
        email: contacto.Correo || '',
        telefono: contacto.Celular || '',
        display: `${nombre}${cargo ? ' - ' + cargo : ''}`.trim() // El frontend espera esto
      };
    });
    
    // 4. Devolver el array de objetos "traducido"
    return contactosTraducidos;

  } catch (e) {
    Logger.log(`ERROR en getContactosParaComercial (Supabase): ${e.message}`);
    // Devuelve un array vacío en caso de error para que el frontend no se rompa
    return []; 
  }
}

/**
 * FUNCIÓN DE AYUDA (PRIORIDAD 2)
 * Crea o actualiza la fila de resumen en la hoja 'ResumenCot'.
 * Esta función es llamada por 'guardarCotizacion'.
 * @param {string} codigoPedido - El ID del pedido (ej. "COT.ALP...").
 * @param {Object} datosSanitizados - Los datos de cabecera del formulario.
 * @param {number} filaInicioDataCot - El N° de fila donde se insertó la *primera* línea en DataCot.
 */
function _actualizarResumenCot(codigoPedido, datosSanitizados, filaInicioDataCot) {
    try {
        const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
        const sheetResumen = ss.getSheetByName(HOJA_RESUMEN_COT);
        
        if (!sheetResumen) {
            Logger.log(`Error: No se encontró la hoja de resumen '${HOJA_RESUMEN_COT}'. No se puede actualizar el resumen.`);
            return; // Salir silenciosamente para no detener el guardado principal
        }

        const COL_MAP_RESUMEN = getColumnMap(HOJA_RESUMEN_COT);
        const ID_PEDIDO_COL = COL_MAP_RESUMEN['ID_PEDIDO'];

        if (ID_PEDIDO_COL === undefined) {
             Logger.log(`Error: No se encontró la columna 'ID_Pedido' en '${HOJA_RESUMEN_COT}'.`);
             return;
        }

        const allDataResumen = sheetResumen.getDataRange().getValues();
        let rowIndexToUpdate = -1;

        // Buscar la fila existente
        for (let i = 1; i < allDataResumen.length; i++) {
            if (String(allDataResumen[i][ID_PEDIDO_COL] || '').trim() === codigoPedido) {
                rowIndexToUpdate = i + 1; // N° de fila (1-based)
                break;
            }
        }

        // Preparar la fila con los datos más recientes
        const montoTotal = parseFloat(datosSanitizados.Total) || 0;
        const numColsResumen = sheetResumen.getLastColumn() > 0 ? sheetResumen.getLastColumn() : 10;
        const filaResumen = new Array(numColsResumen).fill('');

        filaResumen[COL_MAP_RESUMEN['ID_PEDIDO']] = codigoPedido;
        filaResumen[COL_MAP_RESUMEN['FECHA_COT']] = datosSanitizados.numPedido ? (allDataResumen[rowIndexToUpdate -1][COL_MAP_RESUMEN['FECHA_COT']] || new Date()) : new Date(); // Conservar fecha original si se actualiza
        filaResumen[COL_MAP_RESUMEN['CLIENTE']] = datosSanitizados.Cliente || '';
        filaResumen[COL_MAP_RESUMEN['ID_CLIENTE']] = datosSanitizados.RUC || '';
        filaResumen[COL_MAP_RESUMEN['EJECUTIVO']] = datosSanitizados.Ejecutivo || '';
        filaResumen[COL_MAP_RESUMEN['EMPRESA']] = datosSanitizados.Empresa || '';
        filaResumen[COL_MAP_RESUMEN['MONTO_TOTAL']] = montoTotal;
        filaResumen[COL_MAP_RESUMEN['MONEDA']] = datosSanitizados.Moneda || '';
        filaResumen[COL_MAP_RESUMEN['ESTADO_PEDIDO']] = datosSanitizados.Estado || '';
        
        // Guardar el RowIndex de DataCot solo en la creación
        if (rowIndexToUpdate === -1) {
             filaResumen[COL_MAP_RESUMEN['DATACOT_ROWINDEX']] = filaInicioDataCot;
        } else {
             // Conservar el RowIndex original si ya existía
             filaResumen[COL_MAP_RESUMEN['DATACOT_ROWINDEX']] = allDataResumen[rowIndexToUpdate -1][COL_MAP_RESUMEN['DATACOT_ROWINDEX']];
        }


        // Escribir en la hoja de resumen
        if (rowIndexToUpdate > 1) {
            // Actualizar fila existente
            sheetResumen.getRange(rowIndexToUpdate, 1, 1, filaResumen.length).setValues([filaResumen]);
        } else {
            // Crear nueva fila
            sheetResumen.appendRow(filaResumen);
        }

        // Invalidar caché de resumen
        cache.remove(`hoja_${HOJA_RESUMEN_COT}`);
        Logger.log(`ResumenCot actualizado para ${codigoPedido}.`);

    } catch (e) {
        Logger.log(`ERROR al actualizar ResumenCot: ${e.message}`);
        enviarNotificacionError(`Fallo al actualizar ResumenCot para ${codigoPedido}: ${e.message}`);
    }
}

/**
 * Devuelve la representación de fecha segura (ISO string) o un valor por defecto.
 */
function getSafeDateString(value, defaultValue) {
    if (value instanceof Date) {
        return value.toISOString();
    }
    if (value) {
        try {
            const date = new Date(value);
            if (!isNaN(date)) {
                return date.toISOString();
            }
        } catch (e) {
            // Ignorar error de parseo
        }
    }
    return new Date(defaultValue || new Date()).toISOString();
}
/**
 * Guarda o actualiza una Dirección usando crudHoja.
 */
function guardarOActualizarDireccion(data) {
    try {
        const datosSanitizados = sanitizarDatos(data);
        const rowIndex = parseInt(datosSanitizados.rowIndex);
        
        // Generar un ID solo si es un nuevo registro
        const idDireccion = datosSanitizados.ID_DIRECCION || (rowIndex > 1 ? datosSanitizados.ID_DIRECCION : `DIR-${new Date().getTime()}`);
        
        // Valores alineados con las columnas de la hoja Direcciones
        const nuevosValores = [
            idDireccion,                            // Columna 1: ID
            datosSanitizados.RUC,                   // Columna 2: RUC
            datosSanitizados.TIPO,                  // Columna 3: TIPO
            datosSanitizados.DIRECCION,             // Columna 4: DIRECCION
            datosSanitizados.CIUDAD                 // Columna 5: CIUDAD
        ];
        
        const operacion = rowIndex > 1 ? 'UPDATE' : 'CREATE';
        
        const resultado = crudHoja(operacion, HOJA_DIRECCIONES, { rowIndex: rowIndex, valores: nuevosValores });

        if (resultado.success) {
            return { success: true, message: "Dirección guardada exitosamente" };
        } else {
            throw new Error(resultado.message);
        }
    } catch (error) {
        throw new Error("Error al guardar la dirección: " + error.message);
    }
}

// ====================================================
// === FUNCIONES DE SERVICIOS OPTIMIZADAS (CRUD) ===
// ====================================================

function getDatosInicialesServicios() {
    try {
        const allServiciosData = obtenerDatosHoja(HOJA_SERVICIOS);
        if (allServiciosData.length === 0) {
            return {
                servicios: [['ID Servicio', 'Descripción del Servicio', 'Precio Unitario', 'Costo Unitario', 'Hora según', 'Und. Medida']],
                horasSegun: [],
                undMedida: [],
            };
        }
        
        return {
            servicios: allServiciosData,
            horasSegun: getListaValoresUnicosOptimizada(allServiciosData, 4),
            undMedida: getListaValoresUnicosOptimizada(allServiciosData, 5)
        };
    } catch (error) {
        return manejarError('getDatosInicialesServicios', error);
    }
}

function guardarOActualizarServicio(dataObject) {
  // --- NUEVA VERIFICACIÓN DE PERMISO ---
    const permisos = obtenerPermisosUsuario();
    if (!permisos.puedeEditarServicios) {
        // Devolver un error manejable
        return { success: false, message: "Acceso denegado. No tiene permiso para editar servicios." };
    }
    // --- FIN DE VERIFICACIÓN ---
    try {
        const codigoServicio = (dataObject['ID Servicio'] || '').toString().trim();
        const descripcion = (dataObject['Descripción del Servicio'] || '').toString().trim();
        
        if (!codigoServicio) throw new Error("El código del servicio es requerido");
        if (!descripcion) throw new Error("La descripción del servicio es requerida");
        
        if (parseInt(dataObject.rowIndex) <= 1) { 
            const servicioExistente = buscarServicioPorCodigo(codigoServicio);
            if (servicioExistente) throw new Error(`El código ${codigoServicio} ya existe. Use un código único.`);
        }
        
        const valores = [
            codigoServicio,
            descripcion,
            parseFloat(dataObject['Precio Unitario'] || 0),
            parseFloat(dataObject['Costo Unitario'] || 0),
            dataObject['Hora según'] || '',
            dataObject['Und. Medida'] || '',
            dataObject['Notas'] || ''
        ];
        
        const resultado = crudHoja(
            parseInt(dataObject.rowIndex) > 1 ? 'UPDATE' : 'CREATE',
            HOJA_SERVICIOS, 
            { rowIndex: dataObject.rowIndex, valores: valores }
        );
        return resultado;
        
    } catch (error) {
        throw error;
    }
}

function obtenerServicioPorRowIndex(rowIndex) {
    try {
        const rowData = crudHoja('READ_ROW', HOJA_SERVICIOS, { rowIndex: rowIndex });
        if (!rowData) throw new Error("No se encontró el servicio.");

        const allData = obtenerDatosHoja(HOJA_SERVICIOS);
        const encabezados = allData[0];
        
        return { row: rowData, rowIndex: rowIndex, encabezados: encabezados };
    } catch (error) {
        throw new Error("No se pudo cargar el servicio: " + error.message);
    }
}


// ====================================================
// === FUNCIONES DE MÓDULO CONTACTOS (DETALLE) ===
// ====================================================

function getContactosYDirecciones(ruc) {
    try {
        const allContactos = obtenerDatosHoja(HOJA_CONTACTOS);
        const allDirecciones = obtenerDatosHoja(HOJA_DIRECCIONES);

        // Función auxiliar para añadir el rowIndex
        const obtenerDetallesConRowIndex = (allData, colRUC) => {
            if (allData.length <= 1) return [allData[0]];
            const headers = allData[0];
            const filteredRows = allData.slice(1).filter(row => String(row[colRUC]) == String(ruc));
            
            const rowsWithIndex = filteredRows.map(row => {
                // Buscamos el índice original en la hoja completa para el CRUD
                const rowIndex = allData.findIndex(dataRow => JSON.stringify(dataRow) === JSON.stringify(row));
                return row.concat(rowIndex + 1); 
            });
            // Devuelve encabezados con ROW_INDEX y filas de datos
            return [headers.concat("ROW_INDEX")].concat(rowsWithIndex);
        };
        
        // Usamos los índices RUC de las constantes
        const contactos = obtenerDetallesConRowIndex(allContactos, CONTACTO_COLS.RUC);
        const direcciones = obtenerDetallesConRowIndex(allDirecciones, DIRECCION_COLS.RUC);

        return { contactos: contactos, direcciones: direcciones };
    } catch (error) {
        return manejarError('getContactosYDirecciones', error);
    }
}

function getFilaPorRowIndex(ruc, rowIndex, tipo) {
    try {
        let sheetName;
        if (tipo === 'contacto') sheetName = HOJA_CONTACTOS;
        else if (tipo === 'direccion') sheetName = HOJA_DIRECCIONES;
        else if (tipo === 'servicio') sheetName = HOJA_SERVICIOS;
        else throw new Error("Tipo de búsqueda inválido.");
        
        const rowData = crudHoja('READ_ROW', sheetName, { rowIndex: rowIndex });

        if (!rowData) throw new Error("Fila no encontrada o fuera de rango.");

        return { row: rowData, rowIndex: rowIndex };
    } catch (error) {
        return manejarError('getFilaPorRowIndex', error);
    }
}

/**
 * REFACTORIZADO (v3 - Supabase)
 * Obtiene la lista de cotizaciones desde la tabla "Pedidos"
 * y la cruza con "Clientes" para obtener el nombre.
 */
function getListaCotizacionesResumen() {
  try {
    Logger.log("Ejecutando getListaCotizacionesResumen (Versión Supabase)...");

    // 1. Definimos la consulta a Supabase.
    // Esta consulta especial le pide a Supabase que:
    // "Selecciona todas estas columnas de 'Pedidos', y de la tabla 'Clientes' 
    // (que está conectada por el RUC), tráeme solo 'Nombre_RazonSocial'"
    //
    // ¡IMPORTANTE! Esto solo funciona si creaste una Foreign Key en Supabase
    // desde 'Pedidos.RUC' hacia 'Clientes.RUC_DNI'.
    
    const consulta = 'select=Cot,Fecha_Creacion,Ejecutivo,Empresa,Total_Cot,Moneda,Estado_Cot,RUC,Clientes(Nombre_RazonSocial)';

    // 2. Usamos el traductor
    const pedidos = supabaseFetch('Pedidos', {
      method: 'get',
      params: consulta
    });

    Logger.log(`Se encontraron ${pedidos.length} pedidos.`);
    
    // 3. Devolvemos el JSON. El frontend (HTML) se encargará de procesarlo.
    return pedidos; // Esto es un Array de Objetos JSON

  } catch (e) {
    Logger.log(`❌ ERROR FATAL en getListaCotizacionesResumen (Supabase): ${e.message} \n ${e.stack}`);
    return manejarError('getListaCotizacionesResumen', e);
  }
}

// ====================================================
// === LÓGICA DE CÓDIGO DE PEDIDO (SIN CAMBIOS ESTRUCTURALES) ===
// ====================================================

/**
 * Genera código de pedido único
 */
function generarCodigoPedido(empresa) {
    try {
        const props = PropertiesService.getScriptProperties();
        const today = new Date();
        const anio = today.getFullYear();
        const mes = String(today.getMonth() + 1).padStart(2, '0');
        
        const usuarioActual = obtenerUsuarioActual();
        const prefijoEmpresa = obtenerPrefijoEmpresa(empresa);
        const { rangoMin, rangoMax } = obtenerRangoUsuario(usuarioActual);
        const keyContador = `contador_${usuarioActual}_${prefijoEmpresa}_${anio}_${mes}`;
        const numSecuencia = obtenerSiguienteNumero(props, keyContador, rangoMin, rangoMax);
        const codigoFinal = `COT.${prefijoEmpresa}.${anio}.${mes}.${numSecuencia}`;
        
        Logger.log(`Código generado: ${codigoFinal} para usuario: ${usuarioActual}`);
        return codigoFinal;
    } catch (error) {
        Logger.log(`Error en generarCodigoPedido: ${error.message}`);
        return `COT.${empresa?.substring(0, 3) || 'GEN'}.${new Date().getTime()}`;
    }
}

/**
 * Función SIMPLIFICADA para obtener usuario actual
 */
function obtenerUsuarioActual() {
    try {
        const email = obtenerEmailSeguro();
        if (!email) {
            return 'default';
        }
        
        // Mapeo simple de usuarios
        if (email.toLowerCase().includes('carmen')) {
            return 'carmen';
        } else if (email.toLowerCase().includes('anthony') || email.toLowerCase().includes('antony')) {
            return 'anthony';
        } else if (email.toLowerCase().includes('renato')) {
            return 'renato';
        } else if (email.toLowerCase().includes('joserts126')) {
            return 'anthony';
        }
        
        return email.split('@')[0].toLowerCase();
    } catch (error) {
        Logger.log(`Error al obtener usuario: ${error.message}`);
        return 'default';
    }
}

/**
 * Obtiene prefijo de empresa
 */
function obtenerPrefijoEmpresa(empresa) {
    if (!empresa) return 'GEN';
    const empresaLower = empresa.toLowerCase().trim();
    if (empresaLower.includes('alpamayo') || empresaLower === 'alp') return 'ALP';
    if (empresaLower.includes('san jose') || empresaLower === 'sj') return 'SJ';
    if (empresaLower.includes('gym') || empresaLower === 'gym') return 'GYM';
    return empresa.substring(0, 3).toUpperCase();
}

/**
 * Obtiene rango de números por usuario
 */
function obtenerRangoUsuario(usuario) {
    const usuarioLower = usuario.toLowerCase();
    if (usuarioLower.includes('carmen')) return { rangoMin: 2001, rangoMax: 3000 };
    if (usuarioLower.includes('anthony') || usuarioLower.includes('antony')) return { rangoMin: 3001, rangoMax: 4000 };
    return { rangoMin: 4001, rangoMax: 5000 };
}

/**
 * Obtiene siguiente número de secuencia
 */
function obtenerSiguienteNumero(props, keyContador, rangoMin, rangoMax) {
    let numSecuencia = parseInt(props.getProperty(keyContador)) || (rangoMin - 1);
    numSecuencia++;
    if (numSecuencia > rangoMax) {
        numSecuencia = rangoMin;
        Logger.log(`Contador reiniciado al rango mínimo: ${rangoMin}`);
    }
    props.setProperty(keyContador, numSecuencia.toString());
    return numSecuencia;
}

/**
 * REFACTORIZADO (v3 - Supabase)
 * Verifica si un código de cotización ya existe en la tabla "Pedidos".
 */
function verificarCodigoUnico(codigoGenerado) {
  try {
    Logger.log(`Verificando unicidad de código: ${codigoGenerado}`);

    // 1. Llama a la tabla "Pedidos" y filtra por "Cot"
    const resultado = supabaseFetch('Pedidos', {
      method: 'get',
      // Pedimos solo la columna "Cot" y filtramos
      params: `select=Cot&Cot=eq.${codigoGenerado}`
    });
    
    // 2. Si el array devuelto está vacío (length 0), el código NO existe,
    // por lo tanto, ES único.
    if (resultado.length === 0) {
      Logger.log("Código es único.");
      return true; // Es único
    } else {
      Logger.log("Código duplicado encontrado.");
      return false; // No es único
    }

  } catch (error) {
    Logger.log(`Error en verificarCodigoUnico (Supabase): ${error.message}`);
    // Si hay un error de base de datos, es más seguro
    // asumir que el código NO es único para evitar colisiones.
    return false; 
  }
}

// ====================================================
// === VALIDACIÓN Y SANITIZACIÓN ===
// ====================================================

/**
 * Valida datos de cotización
 */
function validarDatosCotizacion(datos) {
    const errores = [];
    // Validar campos requeridos
    const camposRequeridos = ['Empresa', 'RUC', 'Cliente', 'Moneda', 'Forma_Pago'];
    camposRequeridos.forEach(campo => {
        if (!datos[campo] || datos[campo].toString().trim() === '') {
            errores.push(`El campo ${campo} es obligatorio`);
        }
    });
    // Validar RUC/DNI
    if (datos.RUC && !validarRUC(datos.RUC)) {
        errores.push('El RUC/DNI no tiene un formato válido');
    }
    
    // Validar líneas de servicio
    if (!datos.Lineas || datos.Lineas.length === 0) {
        errores.push('Debe agregar al menos un servicio a la cotización');
    } else {
        datos.Lineas.forEach((linea, index) => {
            if (!linea.cod || linea.cod.trim() === '') {
                errores.push(`Línea ${index + 1}: El código de servicio es obligatorio`);
            }
            if (!linea.cantidad || parseFloat(linea.cantidad) <= 0) {
              
              errores.push(`Línea ${index + 1}: La cantidad debe ser mayor a 0`);
            }
        });
    }
    
    return errores;
}

/**
 * Valida formato de RUC/DNI
 */
function validarRUC(ruc) {
    const cleanRuc = ruc.toString().trim();
    if (!cleanRuc) return false;
    // Validar DNI (8 dígitos)
    if (cleanRuc.length === 8) return /^\d+$/.test(cleanRuc);
    // Validar RUC (11 dígitos)
    if (cleanRuc.length === 11) return /^\d+$/.test(cleanRuc);
    
    return false;
}

/**
 * Sanitiza datos para prevenir XSS
 */
function sanitizarDatos(datos) {
    const sanitizados = {};
    Object.keys(datos).forEach(key => {
        if (typeof datos[key] === 'string') {
            // Eliminar scripts y caracteres peligrosos
            sanitizados[key] = datos[key]
                .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
                .replace(/javascript:/gi, '')
                .replace(/on\w+=/gi, '');
   
         } else {
            sanitizados[key] = datos[key];
        }
    });
    return sanitizados;
}

// ====================================================
// === MANEJO DE ERRORES CENTRALIZADO ===
// ====================================================

/**
 * Maneja errores de forma centralizada
 */
function manejarError(contexto, error) {
    const timestamp = new Date().toISOString();
    const mensajeError = `[${timestamp}] ERROR en ${contexto}: ${error.message}`;
    
    Logger.log(mensajeError);
    
    // Enviar email de notificación para errores críticos
    if (error.message.includes('CRÍTICO') || error.message.includes('FATAL')) {
        enviarNotificacionError(mensajeError);
    }
    
    return {
        success: false,
        message: 'Ocurrió un error inesperado. Por favor, intente nuevamente.',
        errorId: timestamp
    };
}

/**
 * Envía notificación de error por email
 */
function enviarNotificacionError(mensaje) {
    try {
        MailApp.sendEmail({
            to: Session.getEffectiveUser().getEmail(),
            subject: '🚨 Error en Sistema de Cotizaciones',
            body: mensaje
        });
    } catch (e) {
        Logger.log('Error al enviar notificación: ' + e.message);
    }
}

// ====================================================
// === FUNCIONES DE OPERACIONES DE TRABAJO (OT) ===
// ====================================================

/**
 * REEMPLAZO (v2 - Supabase): Obtiene la lista COMPLETA de Órdenes de Trabajo.
 * Realiza JOINs anidados para obtener los nombres de Cliente, Servicio y Pedido.
 */
function getListaOT() {
  try {
    Logger.log("Ejecutando getListaOT (v2 - Supabase)...");

    // Esta consulta anidada trae los datos relacionados que necesitamos para la tabla
    const consulta = `
      select=
        N_OT,
        Fecha,
        Horometro_Trabajado_Horas,
        Detalle_Pedidos!inner (
          Cot,
          Servicios!inner (Nombre_Servicio),
          Pedidos!inner (
            Clientes!inner (Nombre_RazonSocial)
          )
        )
      &order=Fecha.desc
    `.replace(/\s/g, ''); // Limpiar espacios

    const ordenesDeTrabajo = supabaseFetch('Ordenes_Trabajo', {
      method: 'get',
      params: consulta
    });

    Logger.log(`Se encontraron ${ordenesDeTrabajo.length} OTs en Supabase.`);
    return ordenesDeTrabajo; // Devuelve el JSON de objetos

  } catch (e) {
    return manejarError('getListaOT', e);
  }
}

/**
 * Obtiene una Orden de Trabajo por su número de OT (para edición/visualización).
 * Implementa mapeo para extracción robusta.
 */
function obtenerOTPorNumero(numeroOT) {
    try {
        const OT_SHEET = HOJA_OT; 
        const COL_MAP = getColumnMap(OT_SHEET);
        const allData = obtenerDatosHoja(OT_SHEET);
        const OT_COL = COL_MAP['N° OT'] || 0; 
        
        const filaOT = allData.slice(1).find(row => String(row[OT_COL] || '').trim() === numeroOT);
        
        if (!filaOT) throw new Error(`OT ${numeroOT} no encontrada.`);
        
        const getValue = (colName) => getFilaValue(filaOT, COL_MAP, colName);

        const datos = {
            numeroOT: String(getValue('N° OT') || ''),
            fecha: getSafeDateString(getValue('Fecha')),
            cliente: String(getValue('Cliente') || ''), 
            servicio: String(getValue('Servicio/Máquina') || ''),
            pedido: String(getValue('Pedido') || ''),
            horaInicio: String(getValue('Hora Inicio') || ''),
            horaFin: String(getValue('Hora Fin') || ''),
            tiempoRefrigerio: parseFloat(getValue('Tiempo Refrigerio (min)') || 0) || 0,
            horometroInicio: parseFloat(getValue('Horómetro Inicio') || 0) || 0,
            horometroFin: parseFloat(getValue('Horómetro Fin') || 0) || 0,
            esCamionGrua: String(getValue('Es Camión Grúa') || 'NO'),
            horometroInicioCamion: parseFloat(getValue('Horómetro Inicio Camión') || 0) || 0,
            horometroFinCamion: parseFloat(getValue('Horómetro Fin Camión') || 0) || 0,
            horometroInicioGrua: parseFloat(getValue('Horómetro Inicio Grúa') || 0) || 0,
            horometroFinGrua: parseFloat(getValue('Horómetro Fin Grúa') || 0) || 0,
            tiempoTotal: parseFloat(getValue('Horas Trab.') || 0) || 0,
            horometroTrabajado: parseFloat(getValue('Horómetro Trab.') || 0) || 0,
            usuario: String(getValue('Usuario Registro') || ''),
            fechaRegistro: getSafeDateString(getValue('Fecha Registro'))
        };

        return { success: true, data: datos }; 
    } catch (e) {
        return manejarError('obtenerOTPorNumero', e);
    }
}

/**
 * Filtra pedidos por cliente y servicio (Lógica de OT para selects)
 */
function filtrarPedidosPorClienteYServicio(rucCliente, idServicio) {
    try {
        const COT_SHEET = HOJA_COTIZACIONES;
        const allData = obtenerDatosHoja(COT_SHEET, true, 5); // Usar caché
        
        if (allData.length <= 1) return [];

        const COL_MAP = getColumnMap(COT_SHEET);
        const RUC_COL = COL_MAP['ID CLIENTE'];
        const COD_COL = COL_MAP['COD'];
        const COT_COL = COL_MAP['COT'];

        if (RUC_COL === undefined || COD_COL === undefined || COT_COL === undefined) {
             throw new Error("Columnas 'ID CLIENTE', 'COD' o 'COT' no encontradas en DataCot.");
        }

        const rucBuscado = String(rucCliente).trim();
        const servicioBuscado = String(idServicio).trim();
        const pedidosEncontrados = new Set();
        
        // Iterar desde el final para obtener los más recientes primero
        for (let i = allData.length - 1; i > 0; i--) {
            const row = allData[i];
            const rucFila = String(row[RUC_COL] || '').trim();
            const codFila = String(row[COD_COL] || '').trim();

            if (rucFila === rucBuscado && codFila === servicioBuscado) {
                const numPedido = String(row[COT_COL] || '').trim();
                if (numPedido) {
                    pedidosEncontrados.add(numPedido);
                }
            }
        }
        
        // Devolvemos un array de arrays para que coincida con el frontend
        return Array.from(pedidosEncontrados).map(pedido => [pedido]);

    } catch (e) {
        return manejarError('filtrarPedidosPorClienteYServicio', e);
    }
}

/**
 * [NUEVO HELPER]
 * Busca una línea en DataCot por su LineaID y devuelve detalles clave.
 * DEBE ser añadido a _Comercial_Logic.gs
 * @param {string} lineaID El ID único de la línea (ej. "COT...-L1").
 * @returns {object} Un objeto {success, precio, cliente, ruc, horasSegun}
 */
function _getDetallesDeLineaCot(lineaID) {
    Logger.log(`Buscando detalles de DataCot para lineaID: ${lineaID}`);
    if (!lineaID) {
        return { success: false, message: "lineaID nulo o vacío." };
    }
    
    // Usamos obtenerDatosHoja sin caché para asegurar el precio más reciente al *crear* la OT
    const allDataCot = obtenerDatosHoja(HOJA_COTIZACIONES, false); 
    const COL_MAP_COT = getColumnMap(HOJA_COTIZACIONES);

    const LINEAID_COL = COL_MAP_COT['NUM'];
    const PRECIO_COL = COL_MAP_COT['PRECIO'];
    const CLIENTE_COL = COL_MAP_COT['CLIENTE'];
    const RUC_COL = COL_MAP_COT['ID CLIENTE'];
    const HORAS_SEGUN_COL = COL_MAP_COT['HORAS SEGÚN'];

    // Validar que las columnas esenciales existan en DataCot
    if ([LINEAID_COL, PRECIO_COL, CLIENTE_COL, RUC_COL, HORAS_SEGUN_COL].includes(undefined)) {
        Logger.log(`Error de configuración en _getDetallesDeLineaCot: Faltan columnas en DataCot.
            NUM=${LINEAID_COL}, PRECIO=${PRECIO_COL}, CLIENTE=${CLIENTE_COL}, ID CLIENTE=${RUC_COL}, HORAS SEGÚN=${HORAS_SEGUN_COL}`);
        return { success: false, message: "Error de configuración: Faltan columnas (NUM, PRECIO, CLIENTE, ID CLIENTE, HORAS SEGÚN) en DataCot." };
    }

    for (let i = 1; i < allDataCot.length; i++) {
        const row = allDataCot[i];
        if (String(row[LINEAID_COL] || '').trim() === lineaID) {
            const detalles = {
                success: true,
                precio: parseFloat(row[PRECIO_COL] || 0),
                cliente: String(row[CLIENTE_COL] || ''),
                ruc: String(row[RUC_COL] || ''),
                horasSegun: String(row[HORAS_SEGUN_COL] || '').trim().toUpperCase()
            };
            Logger.log(`Detalles encontrados: ${JSON.stringify(detalles)}`);
            return detalles;
        }
    }
    
    Logger.log(`Error: No se encontró la línea con ID ${lineaID} en DataCot.`);
    return { success: false, message: `No se encontró la línea con ID ${lineaID} en DataCot.` };
}


/**
 * REEMPLAZO TOTAL de guardarOT (v7.1 - Corregido con ID Híbrido)
 * Escribe en la tabla "Ordenes_Trabajo" de Supabase
 * y actualiza "Detalle_Pedidos" usando el ID numérico.
 */
function guardarOT(data) {
    const permisos = obtenerPermisosUsuario();
    if (!permisos.puedeEditarOT) {
        return { success: false, message: "Acceso denegado. No tiene permiso para editar OTs." };
    }

    Logger.log("INICIO guardarOT (v7.1 - Supabase). Datos recibidos: " + JSON.stringify(data));
    
    try {
        const datos = sanitizarDatos(data);
        // Este es el ID de TEXTO (ej. "COT...-L1")
        const lineaIDRef = datos.lineaID; 
        const esModoUpdate = (datos.modo === 'editar' || datos.modo === 'editarGlobal');

        if (!lineaIDRef) {
            return manejarError('guardarOT', new Error("Error Crítico: No se proporcionó un lineaID (referencia). No se puede enlazar al pedido."));
        }
        
        // --- PASO 1: Obtener detalles del Pedido (Precio y ID REAL) ---
        // Usamos la RPC que busca por texto
        Logger.log(`Buscando detalles de línea para: ${lineaIDRef}`);
        const detallesLinea = supabaseFetch('rpc/get_detalles_linea_para_ot', {
            method: 'post',
            payload: { "linea_ref_param": lineaIDRef }
        })[0]; 
        
        if (!detallesLinea || !detallesLinea.success) {
            throw new Error(detallesLinea.message || `No se encontraron detalles para la línea ${lineaIDRef}.`);
        }
        
        // ¡ESTE ES EL ID REAL (el número)!
        const lineaIdRealNumerico = detallesLinea.linea_id_real;
        Logger.log(`ID real (numérico) encontrado: ${lineaIdRealNumerico}`);

        // --- PASO 2: Calcular Montos (Sin cambios) ---
        const horasSegun = (detallesLinea.horas_segun || '').toUpperCase();
        let horasParaActualizar = 0;
        
        if (horasSegun.includes("HORÓMETRO") || horasSegun.includes("HOROMETRO")) {
            horasParaActualizar = parseFloat(datos.horometroTrabajado || 0);
        } else {
            horasParaActualizar = parseFloat(datos.tiempoTotal || 0);
        }

        const precioUnitario = parseFloat(detallesLinea.precio || 0);
        const montoServicioCalculado = horasParaActualizar * precioUnitario;
        const montoMovilizacionForm = parseFloat(datos.montoMovilizacion || 0);
        
        // --- PASO 3: Preparar el payload para la tabla "Ordenes_Trabajo" ---
        const payloadOT = {
            "N_OT": datos.numeroOT,
            "Cot_Linea": lineaIdRealNumerico, // <-- Usamos el ID NUMÉRICO
            "Fecha": datos.fecha,
            "Hora_Inicio": datos.horaInicio || null,
            "Hora_Fin": datos.horaFin || null,
            "Tiempo_Refrigerio": parseFloat(datos.tiempoRefrigerio || 0),
            "Horometro_Inicio": parseFloat(datos.horometroInicio || 0),
            "Horometro_Fin": parseFloat(datos.horometroFin || 0),
            "Tiempo_Total_Horas": parseFloat(datos.tiempoTotal || 0),
            "Horometro_Trabajado_Horas": parseFloat(datos.horometroTrabajado || 0),
            "Monto_Servicio": montoServicioCalculado,
            "Monto_Movilizacion": montoMovilizacionForm,
            "Estado_Valorizacion": 'Pendiente'
        };

        // --- PASO 4: Guardar en Supabase (Crear o Actualizar) ---
        if (esModoUpdate) {
            // --- MODO ACTUALIZAR (PATCH) ---
            const otIDExistente = datos.otIDExistente;
            Logger.log(`Modo UPDATE para OT: ${otIDExistente}`);
            
            delete payloadOT.Estado_Valorizacion; 
            
            supabaseFetch('Ordenes_Trabajo', {
                method: 'patch',
                payload: payloadOT,
                params: `N_OT=eq.${otIDExistente}`
            });
            
            Logger.log(`OT ${otIDExistente} actualizada. La actualización de Detalle_Pedidos en modo edición no está implementada.`);
            return { success: true, message: `OT ${otIDExistente} actualizada.` };

        } else {
            // --- MODO CREAR (POST) ---
            Logger.log(`Modo CREATE para OT: ${datos.numeroOT}`);
            
            payloadOT.Usuario_Registro = obtenerEmailSeguro();
            payloadOT.Fecha_Registro = new Date().toISOString();

            supabaseFetch('Ordenes_Trabajo', {
                method: 'post',
                payload: payloadOT
            });

            // --- PASO 5: Actualizar (ACUMULAR) en "Detalle_Pedidos" ---
            // Usamos la RPC que actualiza por el ID numérico
            Logger.log(`Llamando a RPC 'actualizar_despacho_detalle' para ${lineaIdRealNumerico}...`);
            const payloadAcumular = {
                "linea_id_real_param": lineaIdRealNumerico, // <-- Usamos el ID NUMÉRICO
                "horas_sumar": horasParaActualizar,
                "monto_sumar": montoServicioCalculado,
                "movilizacion_sumar": montoMovilizacionForm
            };
            
            supabaseFetch('rpc/actualizar_despacho_detalle', {
                method: 'post',
                payload: payloadAcumular
            });

            Logger.log(`OT ${datos.numeroOT} creada y Detalle_Pedidos actualizado.`);
            return { success: true, message: `OT ${datos.numeroOT} registrada y Pedido actualizado.` };
        }
        
    } catch (e) {
        Logger.log(`ERROR FATAL en guardarOT (v7.1): ${e.message} \n Stack: ${e.stack}`);
        if (e.message.includes('duplicate key value violates unique constraint "Ordenes_Trabajo_pkey"')) {
            return manejarError('guardarOT', new Error(`El número de OT '${datos.numeroOT}' ya existe. Por favor, ingrese un número único.`));
        }
        return manejarError('guardarOT', e);
    }
}
// ====================================================
// === FUNCIONES DE PDF (SIN CAMBIOS ESTRUCTURALES) ===
// ====================================================

/**
 * Genera un PDF de la cotización usando las plantillas de Google Sheets.
 * Selecciona la plantilla según la ubicación del servicio con un fallback seguro.
 * @param {Object} cotizacionData - Objeto con los datos de la cotización.
 * @returns {Blob} El archivo PDF generado como un Blob.
 */
function generarPDFCotizacion(cotizacionData) {
    
    // 1. Determinar la plantilla a usar según el lugar
    const lugar = cotizacionData.lugar ? cotizacionData.lugar.toUpperCase() : '';
    let nombrePlantilla;

    if (lugar.includes('ALPAMAYO')) {
        nombrePlantilla = HOJA_PLANTILLA_ALPAMAYO;
    } else if (lugar.includes('GYM')) {
        nombrePlantilla = HOJA_PLANTILLA_GYM;
    } else if (lugar.includes('SAN JOSE') || lugar.includes('GSJ')) {
        nombrePlantilla = HOJA_PLANTILLA_SANJOSE;
    } else {
        // ✅ FIX CRÍTICO: Usar ALPAMAYO como plantilla GENÉRICA si la ubicación no se reconoce.
        Logger.log(`⚠️ Ubicación '${cotizacionData.lugar}' no reconocida. Usando plantilla GENÉRICA.`);
        nombrePlantilla = HOJA_PLANTILLA_ALPAMAYO; 
    }

    const ssTemplate = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
    const hojaMaestra = ssTemplate.getSheetByName(nombrePlantilla);
    
    if (!hojaMaestra) {
        throw new Error("❌ No se encontró la hoja de plantilla: " + nombrePlantilla);
    }

    // 2. Crear una copia temporal de la hoja maestra para rellenar
    const nombreTemporal = "TEMP_" + cotizacionData.ruc + "_" + new Date().getTime();
    const hojaTemporal = hojaMaestra.copyTo(ssTemplate);
    hojaTemporal.setName(nombreTemporal);
    hojaTemporal.showSheet();

    const SERVICIOS = cotizacionData.servicios || [];
    const START_ROW = 18;
    let hojaEliminar = hojaTemporal; 

    try {
        // 3. Rellenar los datos de cabecera
        hojaTemporal.getRange('C5').setValue(cotizacionData.cliente || '');
        hojaTemporal.getRange('C6').setValue(cotizacionData.ruc || '');
        hojaTemporal.getRange('C7').setValue(cotizacionData.contacto || '');
        hojaTemporal.getRange('C8').setValue(cotizacionData.correoCelular || '');
        hojaTemporal.getRange('C12').setValue(cotizacionData.lugar || '');
        hojaTemporal.getRange('C13').setValue(cotizacionData.turno || '');
        hojaTemporal.getRange('C14').setValue(cotizacionData.duracion || '');
        
        let fechaCotizacion = cotizacionData.fecha ? new Date(cotizacionData.fecha) : new Date();
        hojaTemporal.getRange('G3').setValue(fechaCotizacion).setNumberFormat('dd/MM/yyyy'); 

        // 4. Agregar las líneas de servicio
        const numServicios = SERVICIOS.length;
        if (numServicios > 1) {
            hojaTemporal.insertRowsAfter(START_ROW, numServicios - 1);
        }

        const dataCompleta = SERVICIOS.map((s, index) => [
            index + 1,            // Columna A
            s.descripcion || '',  // Columna B
            
            '',                   // Columna C
            '',                   // Columna D
            '',                   // Columna E
 
            s.valor || 0          // Columna F
        ]);

        if (dataCompleta.length > 0) {
            hojaTemporal.getRange(START_ROW, 1, dataCompleta.length, 6).setValues(dataCompleta);
        }
        
        let currentRow = START_ROW + numServicios; 

        // 5. Agregar la línea de total
        hojaTemporal.getRange('A' + currentRow).setValue('TOTAL');
        hojaTemporal.getRange('F' + currentRow).setValue(cotizacionData.total || 0);

        // 6. Generar el PDF
        
        const URL = ssTemplate.getUrl();
        const exportUrl = URL.replace('/edit', '/export?exportFormat=pdf' +
            '&gid=' + hojaTemporal.getSheetId() + 
            '&format=pdf' +
            '&size=A4' + 
            '&portrait=true' + 
            '&fitw=true' + 
            '&gridlines=false' +
            '&sheetnames=false');

        const response = UrlFetchApp.fetch(exportUrl, {
            headers: {
                'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(), 
            },
            muteHttpExceptions: true
        });
        const pdfBlob = response.getBlob().setName((cotizacionData.cliente || 'CLIENTE') + "_COTIZACION.pdf");
        
        return pdfBlob;
    } catch (e) {
        Logger.log('Error al generar PDF: ' + e.toString());
        throw new Error("Error en el PDF: " + e.message);
    } finally {
        // 7. Limpieza: Eliminar la hoja temporal
        if (hojaEliminar) {
            ssTemplate.deleteSheet(hojaEliminar);
        }
    }
}


/**
 * Función de prueba para diagnosticar si el sistema de lectura de datos funciona.
 * Ejecutar SÓLO desde el editor de Apps Script (Run -> testCargaDatos).
 */
function testCargaDatos() {
  try {
    // Intenta usar la función central de lectura
    const datosCot = obtenerDatosHoja(HOJA_COTIZACIONES);
    
    // Intenta usar el mapeo de columnas
    const map = getColumnMap(HOJA_COTIZACIONES);
    
    Logger.log("✅ ÉXITO en la carga de datos.");
    Logger.log("Filas cargadas de DataCot: " + datosCot.length);
    Logger.log("Columnas en DataCot: " + datosCot[0].length);
    Logger.log("Índice de columna CONTACTO: " + map['CONTACTO']);

    // Si llega hasta aquí, significa que las bases funcionan.
    return "Éxito: La lectura y el mapeo de columnas funcionan.";
  } catch (e) {
    Logger.log("❌ ERROR CRÍTICO DE CARGA: " + e.message);
    // Si falla, el error aparecerá en los logs con el archivo y la línea.
    return "Fallo: Verifique el log (Ctrl+Enter) en el editor. El error es: " + e.message;
  }
}

/**
 * Obtiene el valor de una columna de forma segura (resiliente a índices fuera de límites).
 * @param {Array} fila El array de datos de la fila actual.
 * @param {Object} COL_MAP El mapa de encabezados.
 * @param {string} colName El nombre de la columna a buscar.
 * @returns {*} El valor de la celda o null si no se encuentra.
 */
function getFilaValue(fila, COL_MAP, colName) {
    const index = COL_MAP[colName];
    // Retorna el valor si el índice existe y no está fuera de los límites de la fila
    return (index !== undefined && index < fila.length) ? fila[index] : null;
}

// ====================================================
// === FUNCIONES DE SEGURIDAD Y UTILIDAD ===
// ====================================================

/**
 * Obtiene email de forma segura
 */
function obtenerEmailSeguro() {
    try {
        return Session.getEffectiveUser().getEmail();
    } catch (error) {
        return 'Usuario Desconocido';
    }
}

// ====================================================
// === MANEJO DE ERRORES CENTRALIZADO (CRÍTICO) ===
// ====================================================

/**
 * Envía notificación de error por email
 */
function enviarNotificacionError(mensaje) {
    try {
        MailApp.sendEmail({
            to: Session.getEffectiveUser().getEmail(),
            subject: '🚨 Error en Sistema de Cotizaciones',
            body: mensaje
        });
    } catch (e) {
        Logger.log('Error al enviar notificación: ' + e.message);
    }
}

/**
 * Maneja errores de forma centralizada
 */
function manejarError(contexto, error) {
    const timestamp = new Date().toISOString();
    Logger.log(`❌ ERROR en ${contexto}: ${error.message}`);
    
    return {
        success: false,
        message: 'Ocurrió un error inesperado. Por favor, intente nuevamente.',
        error: error.message,
        contexto: contexto
    };
}

/**
 * REFACTORIZADO (v4 - Supabase)
 * Genera el PDF. Ahora acepta 'numPedido' (Cot) en lugar de 'rowIndex'.
 */
function generarYDevolverPDF(numPedido) { // <-- Acepta numPedido
    let newSS; 
    try {
        // 1. Obtener N° de Pedido (Ya lo tenemos)
        if (!numPedido) throw new Error("No se pudo encontrar el N° de Pedido (COT).");
        
        // 2. Obtener TODOS los detalles (¡Usando nuestra función refactorizada!)
        // obtenerDetallesCompletosDePedido debe ser refactorizada también.
        // Por ahora, llamaremos a la función que SÍ refactorizamos:
        const cotizacionData = obtenerPedidoParaEdicion(numPedido);
        if (!cotizacionData.success) {
            throw new Error("No se pudieron cargar los datos del pedido para el PDF.");
        }
        cotizacionData.numPedido = numPedido;
        
        // 3. Obtener información de origen y destino (Sin cambios)
        const fechaObj = new Date(cotizacionData.fecha || new Date());
        const { id: sourceTemplateFileId, tab: sourceTemplateTabName } = getTemplateInfo(cotizacionData.Empresa); 
        const destinationFolder = getDestinationFolder(cotizacionData.Ejecutivo, cotizacionData.Empresa, fechaObj, cotizacionData);
        // (Esta lógica de carpetas sigue funcionando si las constantes son correctas)

        // 4. Crear un NUEVO Google Sheet en blanco (Sin cambios)
        const nombreArchivo = `Cotizacion_${numPedido}`;
        newSS = SpreadsheetApp.create(nombreArchivo); 
        const newFileId = newSS.getId();

        // 5. Copiar SÓLO la pestaña de la plantilla al nuevo archivo (Sin cambios)
        const sourceSS = SpreadsheetApp.openById(sourceTemplateFileId);
        const sourceSheet = sourceSS.getSheetByName(sourceTemplateTabName);
        if (!sourceSheet) {
            DriveApp.getFileById(newFileId).setTrashed(true);
            throw new Error(`La plantilla de origen '${sourceTemplateTabName}' no se encontró.`);
        }
        const hojaTemporal = sourceSheet.copyTo(newSS);
        
        // 6. Limpiar y Mover el archivo nuevo Sheet (Sin cambios)
        hojaTemporal.setName(numPedido);
        const defaultSheet = newSS.getSheetByName('Sheet1'); 
        if (defaultSheet) newSS.deleteSheet(defaultSheet);
        
        const newFile = DriveApp.getFileById(newFileId);
        destinationFolder.addFile(newFile); 
        DriveApp.getRootFolder().removeFile(newFile);
        
        // 7. Archivar Versiones Anteriores (Sin cambios)
        // ... (tu lógica de 'findOrCreateFolder' y 'archivosEnCot' está bien)

        // 8. Rellenar la plantilla
        // ¡NECESITAMOS REFACTORIZAR rellenarPlantilla!
        rellenarPlantilla_Refactorizada(hojaTemporal, cotizacionData); // Usamos una nueva versión
        SpreadsheetApp.flush();

        // 9. Crear el PDF (Sin cambios)
        const newSS_Url = newSS.getUrl();
        const exportUrl = newSS_Url.replace('/edit', '/export?exportFormat=pdf&gid=' + hojaTemporal.getSheetId() + '&format=pdf&size=A4&portrait=true&fitw=true&gridlines=false&sheetnames=false');

        const response = UrlFetchApp.fetch(exportUrl, {
            headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
            muteHttpExceptions: true
        });
        const pdfBlob = response.getBlob().setName(nombreArchivo + ".pdf"); 
        const pdfFile = destinationFolder.createFile(pdfBlob);

        // 10. Devolver ambas URLs (Sin cambios)
        return { 
            success: true, 
            pdfUrl: pdfFile.getUrl(), 
            sheetUrl: newFile.getUrl() 
        };
    } catch (e) {
        Logger.log("Error CRÍTICO en generarYDevolverPDF: " + e.toString());
        if (newSS) DriveApp.getFileById(newSS.getId()).setTrashed(true);
        return { success: false, error: e.message };
    }
}

/**
 * REFACTORIZADO: Rellena la plantilla usando los datos de Supabase
 */
function rellenarPlantilla_Refactorizada(hojaTemporal, cotizacionData) {
    const SERVICIOS = cotizacionData.Lineas || []; // ¡Ahora se llama 'Lineas'!
    const START_ROW = 18; // (Asumiendo que esto sigue igual)

    // 3. Rellenar Cabecera
    hojaTemporal.getRange('C3').setValue(cotizacionData.numPedido || ''); 
    hojaTemporal.getRange('C5').setValue(cotizacionData.Cliente || ''); // (Asegúrate que 'obtenerPedidoParaEdicion' devuelva esto)
    hojaTemporal.getRange('C6').setValue(cotizacionData.RUC || '');      
    hojaTemporal.getRange('C7').setValue(cotizacionData.Contacto || ''); // (Asegúrate que 'obtenerPedidoParaEdicion' devuelva esto)
    hojaTemporal.getRange('C8').setValue(cotizacionData.contactoInfo || ''); // (Esto parece faltar en tu nueva lógica)
    
    const fechaObj = new Date(cotizacionData.fecha || new Date());
    hojaTemporal.getRange('G3').setValue(fechaObj).setNumberFormat('dd/MM/yyyy');
    hojaTemporal.getRange('C12').setValue(cotizacionData.Direccion || '');   
    hojaTemporal.getRange('C13').setValue(cotizacionData.Turno || '');   

    // 4. Lógica para agregar líneas de servicio
    let itemNum = 1;
    let currentRow = START_ROW;

    // 4.1 Contar filas necesarias
    let filasNecesarias = 0;
    SERVICIOS.forEach(s => {
        filasNecesarias++; 
        if (s.movilizacion && s.movilizacion > 0) filasNecesarias++; 
    });
    
    // 4.2 Insertar las filas
    if (filasNecesarias > 1) {
        hojaTemporal.insertRowsAfter(START_ROW, filasNecesarias - 1);
    } else if (filasNecesarias === 0) {
        hojaTemporal.getRange(`B${currentRow}`).setValue("No hay líneas de servicio.");
        return; // Salir si no hay servicios
    }

    // 4.3 Llenar las filas
    SERVICIOS.forEach(s => {
        const movilizacion = s.movilizacion || 0;
        const valorServicio = (s.cantidad * s.precio) || 0; // Calculamos el valor del servicio
        
        let descripcionDetallada = "Alquiler de " + s.descripcion + "\n";
        
        if (s.precio > 0) {
            const simboloMoneda = (cotizacionData.Moneda && cotizacionData.Moneda.toUpperCase() === 'SOLES') ? 'S/.' : 'USD$.';
            descripcionDetallada += `Costo por hora de servicio: ${simboloMoneda} ${s.precio.toFixed(2)} por hora.\n`;
        }
        
        if (s.horas_minimas_num > 0 && s.und_horas_minimas) {
             descripcionDetallada += `Horas mínimas de servicio: ${s.horas_minimas_num} horas mínimas (${s.und_horas_minimas}).\n`;
        }
        
        // (Tu lógica de Almuerzo, Duración, Turno, etc. sigue aquí)
        
        // A. LÍNEA DE SERVICIO PRINCIPAL
        hojaTemporal.getRange(`B${currentRow}:E${currentRow}`).merge().setValue(descripcionDetallada.trim())
            .setHorizontalAlignment("left").setVerticalAlignment("top").setWrap(true);
            
        hojaTemporal.getRange(`F${currentRow}:G${currentRow}`).merge();
        hojaTemporal.getRange(`A${currentRow}`).setValue(itemNum);
        hojaTemporal.getRange(`F${currentRow}`).setValue(valorServicio); 
        currentRow++;
        
        // B. LÍNEA DE MOVILIZACIÓN (SI APLICA)
        if (movilizacion > 0) {
            // ... (Tu lógica de movilización sigue aquí)
            hojaTemporal.getRange(`B${currentRow}:E${currentRow}`).merge().setValue("Movilización y Desmovilización")
                 .setHorizontalAlignment("left").setVerticalAlignment("top").setWrap(true);
            hojaTemporal.getRange(`F${currentRow}:G${currentRow}`).merge();
            hojaTemporal.getRange(`A${currentRow}`).setValue(`${itemNum}.1`);
            hojaTemporal.getRange(`F${currentRow}`).setValue(movilizacion);
            currentRow++; 
        }
        itemNum++; 
    });
    
    // 5. Agregar la línea de total (Sin cambios)
    const filaTotal = currentRow;
    const ultimaFilaItems = filaTotal - 1; 
    hojaTemporal.getRange(`A${filaTotal}:E${filaTotal}`).merge();
    hojaTemporal.getRange(`A${filaTotal}`).setValue("SUBTOTAL");
    hojaTemporal.getRange(`F${filaTotal}:G${filaTotal}`).merge();
    hojaTemporal.getRange(`F${filaTotal}`).setFormula(`=SUM(F${START_ROW}:G${ultimaFilaItems})`);
}

/**
 * Busca info de contacto (email/tel) basado en el RUC y el Nombre del Contacto.
 */
function buscarInfoContacto(ruc, nombreContacto) {
    if (!ruc || !nombreContacto) return '';
    try {
        const allData = obtenerDatosHoja(HOJA_CONTACTOS); // Usa cache
        if (allData.length <= 1) return '';
        
        const COL_MAP_CON = getColumnMap(HOJA_CONTACTOS);
        const RUC_COL = COL_MAP_CON['RUC']; //
        const NOMBRE_COL = COL_MAP_CON['NOMBRE']; //
        const EMAIL_COL = COL_MAP_CON['EMAIL']; //
        const TEL_COL = COL_MAP_CON['TELEFONO']; //

        if (RUC_COL === undefined || NOMBRE_COL === undefined) return ''; // Hoja mal configurada

        const rucBuscado = String(ruc).trim();
        const nombreBuscado = String(nombreContacto).trim();

        for (let i = 1; i < allData.length; i++) {
            const row = allData[i];
            if (String(row[RUC_COL] || '').trim() === rucBuscado && 
                String(row[NOMBRE_COL] || '').trim() === nombreBuscado) {
                
                const email = String(row[EMAIL_COL] || '').trim();
                const tel = String(row[TEL_COL] || '').trim();
                
                if (email && tel) return `${email} / ${tel}`;
                if (email) return email;
                if (tel) return tel;
                return ''; // Encontrado pero sin email/tel
            }
        }
        return ''; // No encontrado
    } catch (e) {
        Logger.log(`Error en buscarInfoContacto: ${e.message}`);
        return '';
    }
}

/**
 * REEMPLAZO FINAL v4 de obtenerDetallesCompletosDePedido
 * Corrige la obtención de la descripción (maneja 'DESCRIPCION' y 'DESCRIPCIÓN').
 */
function obtenerDetallesCompletosDePedido(numPedido) {
    // Logger.log(`Iniciando obtenerDetallesCompletosDePedido para: ${numPedido}`); 
    
    const COL_MAP_COT = getColumnMap(HOJA_COTIZACIONES);
    const allDataCot = obtenerDatosHoja(HOJA_COTIZACIONES); 
    const allDataServ = obtenerDatosHoja(HOJA_SERVICIOS);   
    
    const pedidoBuscado = String(numPedido).trim().toUpperCase();
    const CODIGO_PEDIDO_COL = COL_MAP_COT['COT'] || 0; 

    const filasCoincidentes = allDataCot.slice(1)
        .filter(fila => String(fila[CODIGO_PEDIDO_COL] || '').trim().toUpperCase() === pedidoBuscado);
        
    if (filasCoincidentes.length === 0) {
        throw new Error(`No se encontraron líneas de servicio para el pedido: ${numPedido}`);
    }

    // Preparar mapa de servicios (para abreviatura)
    const mapaServicios = new Map();
    const COL_MAP_SERV = getColumnMap(HOJA_SERVICIOS);
    const COD_SERV_COL = COL_MAP_SERV['COD']; 
    const ABREV_SERV_COL = COL_MAP_SERV['ABREVIATURA']; 
    if (COD_SERV_COL === undefined) throw new Error("No se encontró el encabezado 'COD' en la hoja Servicios.");
    const indiceAbreviatura = (ABREV_SERV_COL !== undefined) ? ABREV_SERV_COL : 7; 

    allDataServ.slice(1).forEach(filaServ => {
        const codServ = String(filaServ[COD_SERV_COL] || '').trim();
        if (codServ) mapaServicios.set(codServ, filaServ); 
    });

    // Mapear líneas de cotización Y buscar abreviatura
    const lineas = filasCoincidentes.map((filaCot) => {
        const getValueCot = (colName) => getFilaValue(filaCot, COL_MAP_COT, colName);
        
        const codServicioCot = String(getValueCot('COD') || '').trim(); 
        let abreviaturaEncontrada = '';

        if (codServicioCot && mapaServicios.has(codServicioCot)) {
            const filaServicioCompleta = mapaServicios.get(codServicioCot);
            if (indiceAbreviatura < filaServicioCompleta.length) {
                abreviaturaEncontrada = String(filaServicioCompleta[indiceAbreviatura] || '').trim();
            }
        }
        
        // --- INICIO DE CORRECCIÓN DESCRIPCIÓN ---
        // Intenta obtener con 'DESCRIPCION', si falla, intenta con 'DESCRIPCIÓN' (acento)
        let descripcion = String(getValueCot('DESCRIPCION') || getValueCot('DESCRIPCIÓN') || '').trim(); 
        // --- FIN DE CORRECCIÓN DESCRIPCIÓN ---

        const moneda = String(getValueCot('MONEDA') || 'USD'); 
        const precioUnitario = parseFloat(getValueCot('PRECIO') || 0); 
        const horasMinimasNum = parseFloat(getValueCot('HORAS MINIMAS') || 0); 
        const undHorasMinimas = String(getValueCot('UND. HORAS. MINIMAS') || ''); 
        const totalDias = parseFloat(getValueCot('TOTAL DIAS') || 0); 

        return {
            cod: codServicioCot, 
            descripcion: descripcion, // <-- Usa la descripción corregida
            abreviatura: abreviaturaEncontrada, 
            valor: parseFloat(getValueCot('M. PEDIDO') || 0) || 0, 
            movilizacion: parseFloat(getValueCot('MOV. Y DES. MOV.') || 0) || 0, 
            moneda: moneda,
            precioUnitario: precioUnitario,
            horasMinimasNum: horasMinimasNum,
            undHorasMinimas: undHorasMinimas,
            totalDias: totalDias
        };
    });
    
    // Obtener datos generales (sin cambios)
    const primeraFila = filasCoincidentes[0];
    const getGenValue = (colName) => getFilaValue(primeraFila, COL_MAP_COT, colName);
    const rucCliente = String(getGenValue('ID CLIENTE') || '');
    const nombreContacto = String(getGenValue('CONTACTO') || '');
    const infoContacto = buscarInfoContacto(rucCliente, nombreContacto);

    // Devolver objeto completo
    const resultadoFinal = {
        success: true,
        ruc: rucCliente, 
        cliente: String(getGenValue('CLIENTE') || ''),
        fecha: getSafeDateString(getGenValue('FECHA COT')),
        contacto: nombreContacto, 
        contactoInfo: infoContacto, 
        lugar: String(getGenValue('UBICACIÓN') || ''),
        turno: String(getGenValue('TURNO') || ''), 
        empresa: String(getGenValue('EMPRESA') || ''),
        ejecutivo: String(getGenValue('EJECUTIVO') || ''),
        total: parseFloat(getGenValue('Total servicio') || 0) || 0,
        servicios: lineas 
    };
    return resultadoFinal;
}
/**
 * NUEVA FUNCIÓN DE AYUDA
 * Obtiene el ID del archivo plantilla y el nombre de la PESTAÑA de plantilla
 * basado en la empresa.
 */
function getTemplateInfo(empresa) {
    const empresaUpper = empresa.toUpperCase();
    
    switch (empresaUpper) {
        case 'ALPAMAYO':
            return { id: ID_PLANTILLA_FILE_ALP, tab: HOJA_PLANTILLA_ALPAMAYO }; // 'COT_ALP'
        case 'GYM':
            return { id: ID_PLANTILLA_FILE_GYM, tab: HOJA_PLANTILLA_GYM }; // 'COT_GYM'
        case 'SAN JOSE':
            return { id: ID_PLANTILLA_FILE_SJ, tab: HOJA_PLANTILLA_SANJOSE }; // 'COT_GSJ'
        default:
            Logger.log(`Empresa no reconocida '${empresa}'. Usando Alpamayo como fallback.`);
            return { id: ID_PLANTILLA_FILE_ALP, tab: HOJA_PLANTILLA_ALPAMAYO };
    }
}

/**
 * NUEVA FUNCIÓN DE AYUDA
 * Busca una carpeta por nombre dentro de una carpeta padre.
 * Si no la encuentra, la crea.
 * @param {Folder} parentFolder La carpeta de Drive donde buscar.
 * @param {string} childName El nombre de la subcarpeta a buscar/crear.
 * @returns {Folder} La carpeta encontrada o recién creada.
 */
function findOrCreateFolder(parentFolder, childName) {
  const carpetas = parentFolder.getFoldersByName(childName);
  
  if (carpetas.hasNext()) {
    return carpetas.next(); // La carpeta ya existe, la devuelve
  } else {
    return parentFolder.createFolder(childName); // La carpeta no existe, la crea
  }
}

/**
 * REEMPLAZO de getDestinationFolder
 * Ahora navega los 4 NIVELES de carpetas.
 */
function getDestinationFolder(ejecutivo, empresa, fechaObj, cotizacionData) {
    let parentFolderId;

    const empresaUpper = empresa.toUpperCase();
    
    // Nivel 1: Carpeta del Ejecutivo
    if (ejecutivo && ejecutivo.toUpperCase() === 'CARMEN') {
        parentFolderId = FOLDER_ID_CARMEN;
    } else {
        const empresaUpper = empresa.toUpperCase();
        if (empresaUpper === 'GYM') parentFolderId = FOLDER_ID_GYM;
        else if (empresaUpper === 'SAN JOSE') parentFolderId = FOLDER_ID_SJ;
        else parentFolderId = FOLDER_ID_ALP;
    }
    
    let currentFolder = DriveApp.getFolderById(parentFolderId);

    // Nivel 2: Carpeta de la Empresa (Ej. "GRUAS SAN JOSE PERU SAC")
    const nombreCarpetaEmpresa = MAPA_NOMBRES_EMPRESAS[empresa.toUpperCase()] || empresa;
    currentFolder = findOrCreateFolder(currentFolder, nombreCarpetaEmpresa);

    // Nivel 3: Carpeta Mes/Año (Ej. "COT.SJ.2025.09 COT_SETIEMBRE")
    const prefijos = {"ALPAMAYO": "COT.ALP", "SAN JOSE": "COT.SJ", "GYM": "COT.GYM"};
    const meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
    
    const prefijoEmpresa = prefijos[empresa.toUpperCase()] || 'COT.GEN';
    const anio = fechaObj.getFullYear();
    const mesNum = String(fechaObj.getMonth() + 1).padStart(2, '0');
    const mesNombre = meses[fechaObj.getMonth()];
    const separador = (empresaUpper === 'ALPAMAYO') ? '_' : ' ';
    
    const nombreSubfolderMes = `${prefijoEmpresa}.${anio}.${mesNum} COT_${mesNombre}`;
    currentFolder = findOrCreateFolder(currentFolder, nombreSubfolderMes);

    // Nivel 4: Carpeta de Cotización Específica (Ej. "COT.SJ.2025.09.2006 STRACON GT100TN")
    let servicioNombre = "VARIOS";
    if (cotizacionData.servicios.length === 1) {
        const servicioUnico = cotizacionData.servicios[0];
        servicioNombre = servicioUnico.abreviatura || servicioUnico.cod || servicioUnico.descripcion.substring(0, 10);
    }
    
    const nombreCarpetaCotizacion = `${cotizacionData.numPedido} ${cotizacionData.cliente} ${servicioNombre}`;
    currentFolder = findOrCreateFolder(currentFolder, nombreCarpetaCotizacion); // <-- Crea o encuentra la carpeta de cotización (Nivel 4)

    // --- AÑADIDO: Nivel 5 ---
    // Crear la subcarpeta "cot" DENTRO de la carpeta de cotización
    currentFolder = findOrCreateFolder(currentFolder, "cot"); 
    // --- FIN DE AÑADIDO ---

    return currentFolder; // Devuelve la carpeta "cot" de Nivel 5
}

/**
 * REEMPLAZO de rellenarPlantilla
 * Construye la descripción detallada en la celda B.
 */
function rellenarPlantilla(hojaTemporal, cotizacionData) {
    const SERVICIOS = cotizacionData.servicios || []; 
    const START_ROW = 18;

    // 3. Rellenar Cabecera (Sin cambios)
    hojaTemporal.getRange('C3').setValue(cotizacionData.numPedido || ''); 
    hojaTemporal.getRange('C5').setValue(cotizacionData.cliente || '');   
    hojaTemporal.getRange('C6').setValue(cotizacionData.ruc || '');      
    hojaTemporal.getRange('C7').setValue(cotizacionData.contacto || ''); 
    hojaTemporal.getRange('C8').setValue(cotizacionData.contactoInfo || ''); 
    const fechaObj = new Date(cotizacionData.fecha || new Date());
    hojaTemporal.getRange('G3').setValue(fechaObj).setNumberFormat('dd/MM/yyyy'); 
    hojaTemporal.getRange('C12').setValue(cotizacionData.lugar || '');   
    hojaTemporal.getRange('C13').setValue(cotizacionData.turno || '');   

    // 4. Lógica para agregar líneas de servicio
    let itemNum = 1;
    let currentRow = START_ROW;

    // 4.1 Contar filas necesarias (Sin cambios)
    let filasNecesarias = 0;
    SERVICIOS.forEach(s => {
        filasNecesarias++; 
        if (s.movilizacion && s.movilizacion > 0) filasNecesarias++; 
    });

    // 4.2 Insertar las filas (Sin cambios)
    if (filasNecesarias > 1) {
        hojaTemporal.insertRowsAfter(START_ROW, filasNecesarias - 1);
    }

    // 4.3 Llenar las filas (CON CAMBIOS EN LA DESCRIPCIÓN)
    SERVICIOS.forEach(s => {
        const movilizacion = s.movilizacion || 0;
        const valorServicio = s.valor || 0; 
        
        // --- INICIO DE CONSTRUCCIÓN DE DESCRIPCIÓN DETALLADA ---
        let descripcionDetallada = "Alquiler de " + s.descripcion + "\n"; // Nombre base + salto de línea
        
        // Añadir costo por hora si hay precio unitario
        if (s.precioUnitario > 0) {
            const simboloMoneda = (s.moneda && s.moneda.toUpperCase() === 'SOLES') ? 'S/.' : 'USD$.';
            descripcionDetallada += `Costo por hora de servicio: ${simboloMoneda} ${s.precioUnitario.toFixed(2)} por hora.\n`;
        }
        
        // Añadir horas mínimas si existen
        if (s.horasMinimasNum > 0 && s.undHorasMinimas) {
             descripcionDetallada += `Horas mínimas de servicio: ${s.horasMinimasNum} horas mínimas (${s.undHorasMinimas}).\n`;
        }
        
        // Añadir Almuerzo (Texto fijo, puedes cambiarlo o hacerlo dinámico si añades el dato)
        descripcionDetallada += "Almuerzo: 1 hora diaria.\n"; 
        
        // Añadir Duración (usando Total Días)
        if (s.totalDias > 0) {
             descripcionDetallada += `Duración: ${Math.ceil(s.totalDias)} día(s) de servicio.\n`;
        }
        
        // Añadir Turno (usando el turno general de la cotización)
        if (cotizacionData.turno) {
            descripcionDetallada += `Turno: ${cotizacionData.turno}.\n`;
        }
        
        // Añadir detalle de Movilización (Texto fijo si hay monto > 0)
        if (movilizacion > 0) {
             // Puedes ajustar este texto si necesitas algo más específico
             descripcionDetallada += "Mov. y desmo. Del equipo: Considerado aparte.\n"; 
        }
        // --- FIN DE CONSTRUCCIÓN DE DESCRIPCIÓN ---

        // A. LÍNEA DE SERVICIO PRINCIPAL
        hojaTemporal.getRange(`B${currentRow}:E${currentRow}`).merge().setValue(descripcionDetallada.trim()) // <-- Poner descripción detallada
            .setHorizontalAlignment("left").setVerticalAlignment("top").setWrap(true); // <-- Forzar alineación y ajuste
            
        hojaTemporal.getRange(`F${currentRow}:G${currentRow}`).merge();
        hojaTemporal.getRange(`A${currentRow}`).setValue(itemNum);
        // hojaTemporal.getRange(`B${currentRow}`).setValue(s.descripcion); // <-- Ya no se usa
        hojaTemporal.getRange(`F${currentRow}`).setValue(valorServicio); 
        currentRow++; 

        // B. LÍNEA DE MOVILIZACIÓN (SI APLICA)
        if (movilizacion > 0) {
            const itemMovNum = `${itemNum}.1`; 
            hojaTemporal.getRange(`B${currentRow}:E${currentRow}`).merge().setValue("Movilización y Desmovilización")
                .setHorizontalAlignment("left").setVerticalAlignment("top").setWrap(true); // <-- Formato
                
            hojaTemporal.getRange(`F${currentRow}:G${currentRow}`).merge();
            hojaTemporal.getRange(`A${currentRow}`).setValue(itemMovNum);
            // hojaTemporal.getRange(`B${currentRow}`).setValue("Movilización y Desmovilización"); // <-- Ya no se usa
            hojaTemporal.getRange(`F${currentRow}`).setValue(movilizacion);
            currentRow++; 
        }
        itemNum++; 
    });

    // 5. Agregar la línea de total (Sin cambios)
    const filaTotal = currentRow;
    const ultimaFilaItems = filaTotal - 1; 
    hojaTemporal.getRange(`A${filaTotal}:E${filaTotal}`).merge();
    hojaTemporal.getRange(`A${filaTotal}`).setValue("SUBTOTAL");
    hojaTemporal.getRange(`F${filaTotal}:G${filaTotal}`).merge();
    hojaTemporal.getRange(`F${filaTotal}`).setFormula(`=SUM(F${START_ROW}:G${ultimaFilaItems})`);
}

/**
 * ====================================================
 * === FUNCIONES PARA GESTIÓN DE ACTAS DE SERVICIO ===
 * ====================================================
 */

/**
 * Verifica si ya existe un acta para una línea específica de una COT.
 * Crea la hoja "Actas" si no existe.
 * @param {string} cotNumero El número de la COT (ej. "COT.GYM.2025.10.2003").
 * @param {number} cotLineaIndex El índice de la línea dentro de la cotización (1, 2, 3...).
 * @returns {string|null} El ActaID si existe, null si no.
 */
function verificarActaExistente(cotNumero, cotLineaIndex) {
  try {
    const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
    let sheet = ss.getSheetByName(HOJA_ACTAS);
    
    // Crear hoja y encabezados si no existen
    if (!sheet) {
      sheet = ss.insertSheet(HOJA_ACTAS);
      // Encabezados esenciales + algunos del PDF
      sheet.appendRow([
          'ActaID', 'COT_Numero', 'COT_LineaIndex', 'FechaCreacion', 'UsuarioCreador', 
          'ClienteNombre', 'ClienteRUC', 'ClienteContactoNombre', 'ClienteContactoCorreo', 
          'FechaServicio', 'ServicioAlcance' // Añade más encabezados según necesites
          // 'PuntoPartida', 'PuntoDescarga', 'Direccion', 'Plazo', 'HorasTotales', ... etc.
      ]);
      Logger.log(`Hoja "${HOJA_ACTAS}" creada con encabezados.`);
    }

    const data = obtenerDatosHoja(HOJA_ACTAS, false); // Leer sin caché para asegurar datos frescos
    if (data.length <= 1) return null; // Hoja vacía o solo encabezados

    const COL_MAP = getColumnMap(HOJA_ACTAS); // Obtener mapa de columnas actualizado
    const COT_NUM_COL = COL_MAP['COT_NUMERO'];
    const LINEA_IDX_COL = COL_MAP['COT_LINEAINDEX'];
    const ACTA_ID_COL = COL_MAP['ACTAID'];

    // Validar que las columnas necesarias existan después de (posiblemente) crear la hoja
    if (COT_NUM_COL === undefined || LINEA_IDX_COL === undefined || ACTA_ID_COL === undefined) {
      Logger.log(`Error Crítico: Faltan columnas esenciales ('COT_NUMERO', 'COT_LINEAINDEX', 'ACTAID') en la hoja "${HOJA_ACTAS}". Verifica los encabezados.`);
      // Podrías lanzar un error aquí si prefieres que falle ruidosamente
      // throw new Error(`Faltan columnas esenciales en la hoja "${HOJA_ACTAS}".`);
      return null; 
    }

    // Buscar el acta
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // Comparar como strings y números respectivamente
      if (String(row[COT_NUM_COL] || '').trim().toUpperCase() === String(cotNumero).trim().toUpperCase() && 
          Number(row[LINEA_IDX_COL]) === Number(cotLineaIndex)) {
        Logger.log(`Acta encontrada para ${cotNumero}, Línea ${cotLineaIndex}: ID ${row[ACTA_ID_COL]}`);
        return String(row[ACTA_ID_COL] || ''); // Devuelve el ActaID encontrado
      }
    }
    
    Logger.log(`No se encontró acta para ${cotNumero}, Línea ${cotLineaIndex}.`);
    return null; // No encontrado

  } catch (e) {
    Logger.log(`Error en verificarActaExistente: ${e.message}`);
    return null; // Asumir que no existe en caso de error grave
  }
}

// --- Marcadores para las funciones que faltan ---

/**
 * Crea una nueva fila en la hoja "Actas".
 * Genera un ActaID único.
 * Recibe un objeto 'actaData' con todos los campos del formulario Acta.html.
 * Devuelve { success: true, actaId: nuevoId } o { success: false, error: mensaje }.
 */
function crearNuevaActa(actaData) {
  try {
    const datosSanitizados = sanitizarDatos(actaData);
    const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
    const sheet = ss.getSheetByName(HOJA_ACTAS);
    
    // Validar datos básicos
    if (!datosSanitizados.COT_Numero || !datosSanitizados.COT_LineaIndex) {
      throw new Error("Faltan datos de la cotización (Número o Línea).");
    }

    // Generar ID único
    const nuevoId = `ACTA-${new Date().getTime()}-${Math.random().toString(36).substring(2, 8)}`;
    const fechaCreacion = new Date();
    const usuarioCreador = obtenerEmailSeguro(); // Asume que esta función existe

    // Mapear datos del formulario a la fila de la hoja
    // **IMPORTANTE:** El orden DEBE coincidir con los encabezados de tu hoja "Actas"
    const COL_MAP = getColumnMap(HOJA_ACTAS);
    const nuevaFila = new Array(sheet.getLastColumn()).fill(''); // Crear array del tamaño de las columnas

    // Función auxiliar para asignar valor si la columna existe
    const setVal = (colName, value) => {
        const index = COL_MAP[colName.toUpperCase()];
        if (index !== undefined) nuevaFila[index] = value;
        else Logger.log(`Advertencia (crearNuevaActa): No se encontró el encabezado '${colName}' en ${HOJA_ACTAS}.`);
    };

    // Asignar valores
    setVal('ActaID', nuevoId);
    setVal('COT_Numero', datosSanitizados.COT_Numero);
    setVal('COT_LineaIndex', datosSanitizados.COT_LineaIndex);
    setVal('FechaCreacion', fechaCreacion);
    setVal('UsuarioCreador', usuarioCreador);

    // Asignar campos del formulario (asegúrate que los nombres coincidan)
    setVal('ClienteNombre', datosSanitizados.clienteNombre || '');
    setVal('ClienteRUC', datosSanitizados.clienteRUC || '');
    setVal('ClienteContactoNombre', datosSanitizados.clienteContactoNombre || '');
    setVal('ClienteContactoCorreo', datosSanitizados.clienteContactoCorreo || '');
    setVal('FechaServicio', datosSanitizados.fechaServicio ? new Date(datosSanitizados.fechaServicio) : '');
    setVal('ServicioAlcance', datosSanitizados.servicioAlcance || '');
    setVal('PuntoPartida', datosSanitizados.puntoPartida || '');
    // setVal('MapsPartida', datosSanitizados.mapsPartida || ''); // Si tienes este campo
    setVal('PuntoDescarga', datosSanitizados.puntoDescarga || '');
    // setVal('MapsDescarga', datosSanitizados.mapsDescarga || ''); // Si tienes este campo
    setVal('Direccion', datosSanitizados.direccion || '');
    // setVal('MapsDireccion', datosSanitizados.mapsDireccion || ''); // Si tienes este campo
    setVal('Plazo', datosSanitizados.plazo || '');
    setVal('HorasTotales', datosSanitizados.horasTotales || '');
    setVal('HorasTraslado', datosSanitizados.horasTraslado || '');
    setVal('HorasTrabajo', datosSanitizados.horasTrabajo || '');
    setVal('InicioServicio', datosSanitizados.inicioServicio || ''); // Formato HH:MM
    setVal('Turno', datosSanitizados.turno || '');
    setVal('EquipoSolicitado', datosSanitizados.equipoSolicitado || '');
    setVal('EncargadoOperativo', datosSanitizados.encargadoOperativo || '');
    setVal('CelularContacto', datosSanitizados.celularContacto || '');
    setVal('PersonalRequerido', datosSanitizados.personalRequerido || '');
    setVal('Observaciones', datosSanitizados.observaciones || '');

    // Añadir la fila a la hoja
    sheet.appendRow(nuevaFila);

    return { success: true, actaId: nuevoId, message: "Acta creada exitosamente." };

  } catch (e) {
    Logger.log(`Error en crearNuevaActa: ${e.message}\nStack: ${e.stack}`);
    return manejarError('crearNuevaActa', e);
  }
}

/**
 * Obtiene todos los datos de una fila de la hoja "Actas" usando su ActaID.
 * Devuelve un objeto con los datos { success: true, data: {...} } o { success: false, error: mensaje }.
 */
function obtenerDatosActa(actaId) {
  try {
    if (!actaId) throw new Error("Se requiere un ID de Acta.");

    const data = obtenerDatosHoja(HOJA_ACTAS, false); // Leer sin caché para asegurar datos frescos
    if (data.length <= 1) throw new Error(`La hoja "${HOJA_ACTAS}" está vacía o no se pudo leer.`);

    const COL_MAP = getColumnMap(HOJA_ACTAS);
    const ACTA_ID_COL = COL_MAP['ACTAID'];
    if (ACTA_ID_COL === undefined) throw new Error(`No se encontró la columna 'ActaID' en la hoja "${HOJA_ACTAS}".`);

    const encabezados = data[0];
    let filaEncontrada = null;
    let rowIndex = -1;

    // Buscar la fila por ActaID
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][ACTA_ID_COL] || '').trim() === String(actaId).trim()) {
        filaEncontrada = data[i];
        rowIndex = i + 1; // Guardar el índice real de la fila para posible actualización
        break;
      }
    }

    if (!filaEncontrada) {
      throw new Error(`No se encontró ningún acta con el ID: ${actaId}`);
    }

    // Convertir la fila en un objeto usando los encabezados
    const actaData = {};
    encabezados.forEach((header, index) => {
      if (header) {
        // Usar el nombre del encabezado original como clave (más fácil en JS)
        // Puedes normalizarlo si prefieres (ej. header.toUpperCase().replace...)
        actaData[header] = filaEncontrada[index]; 
      }
    });

    // Añadir el rowIndex al objeto devuelto, útil para la actualización
    actaData.rowIndex = rowIndex; 

    // Formatear fechas si es necesario para el input type="date"
    if (actaData.FechaServicio && actaData.FechaServicio instanceof Date) {
        actaData.FechaServicio = actaData.FechaServicio.toISOString().split('T')[0];
    } else if (actaData.FechaServicio) {
        // Intentar convertir si no es Date
        try {
           actaData.FechaServicio = new Date(actaData.FechaServicio).toISOString().split('T')[0];
        } catch(e){
            actaData.FechaServicio = ''; // Dejar vacío si no se puede convertir
        }
    }


    return { success: true, data: actaData };

  } catch (e) {
    Logger.log(`Error en obtenerDatosActa: ${e.message}\nStack: ${e.stack}`);
    return manejarError('obtenerDatosActa', e);
  }
}

/**
 * Actualiza una fila existente en la hoja "Actas" usando el ActaID (o rowIndex si se pasa).
 * Recibe un objeto 'actaData' con todos los campos del formulario Acta.html.
 * Debe incluir 'ActaID' o 'rowIndex'.
 * Devuelve { success: true } o { success: false, error: mensaje }.
 */
function actualizarActa(actaData) {
  try {
    const datosSanitizados = sanitizarDatos(actaData);
    let rowIndex = datosSanitizados.rowIndex; // Priorizar rowIndex si viene del frontend
    const actaId = datosSanitizados.ActaID;

    if (!rowIndex && !actaId) {
      throw new Error("Se requiere 'rowIndex' o 'ActaID' para actualizar el acta.");
    }

    const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
    const sheet = ss.getSheetByName(HOJA_ACTAS);
    const COL_MAP = getColumnMap(HOJA_ACTAS);

    // Si no tenemos rowIndex, buscarlo usando ActaID
    if (!rowIndex) {
      const data = obtenerDatosHoja(HOJA_ACTAS, false);
      const ACTA_ID_COL = COL_MAP['ACTAID'];
      if (ACTA_ID_COL === undefined) throw new Error("No se encontró la columna 'ActaID'.");
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][ACTA_ID_COL] || '').trim() === String(actaId).trim()) {
          rowIndex = i + 1;
          break;
        }
      }
      if (!rowIndex) throw new Error(`No se encontró el acta con ID ${actaId} para actualizar.`);
    }

    // Obtener la fila actual para no sobrescribir columnas no incluidas en actaData
    const filaActualValores = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Función auxiliar para actualizar valor si la columna existe
    const setVal = (colName, value) => {
        const index = COL_MAP[colName.toUpperCase()];
        if (index !== undefined) filaActualValores[index] = value;
         else Logger.log(`Advertencia (actualizarActa): No se encontró el encabezado '${colName}' en ${HOJA_ACTAS}.`);
    };

    // Asignar valores actualizados (similar a crearNuevaActa)
    // No actualizamos ActaID, COT_Numero, COT_LineaIndex, FechaCreacion, UsuarioCreador
    setVal('ClienteNombre', datosSanitizados.clienteNombre || '');
    setVal('ClienteRUC', datosSanitizados.clienteRUC || '');
    setVal('ClienteContactoNombre', datosSanitizados.clienteContactoNombre || '');
    setVal('ClienteContactoCorreo', datosSanitizados.clienteContactoCorreo || '');
    setVal('FechaServicio', datosSanitizados.fechaServicio ? new Date(datosSanitizados.fechaServicio) : '');
    setVal('ServicioAlcance', datosSanitizados.servicioAlcance || '');
    setVal('PuntoPartida', datosSanitizados.puntoPartida || '');
    setVal('PuntoDescarga', datosSanitizados.puntoDescarga || '');
    setVal('Direccion', datosSanitizados.direccion || '');
    setVal('Plazo', datosSanitizados.plazo || '');
    setVal('HorasTotales', datosSanitizados.horasTotales || '');
    setVal('HorasTraslado', datosSanitizados.horasTraslado || '');
    setVal('HorasTrabajo', datosSanitizados.horasTrabajo || '');
    setVal('InicioServicio', datosSanitizados.inicioServicio || '');
    setVal('Turno', datosSanitizados.turno || '');
    setVal('EquipoSolicitado', datosSanitizados.equipoSolicitado || '');
    setVal('EncargadoOperativo', datosSanitizados.encargadoOperativo || '');
    setVal('CelularContacto', datosSanitizados.celularContacto || '');
    setVal('PersonalRequerido', datosSanitizados.personalRequerido || '');
    setVal('Observaciones', datosSanitizados.observaciones || '');
    // Añadir aquí cualquier otro campo que se pueda editar

    // Escribir la fila actualizada en la hoja
    sheet.getRange(rowIndex, 1, 1, filaActualValores.length).setValues([filaActualValores]);

    return { success: true, message: "Acta actualizada exitosamente." };

  } catch (e) {
    Logger.log(`Error en actualizarActa: ${e.message}\nStack: ${e.stack}`);
    return manejarError('actualizarActa', e);
  }
}


/**
 * REEMPLAZO v2 de obtenerDatosBaseParaActa
 * Obtiene más datos de DataCot y busca la columna "PERSONAL" en Servicios.
 */
function obtenerDatosBaseParaActa(cotNumero, cotLineaIndex) {
  try {
    if (!cotNumero || !cotLineaIndex) {
      throw new Error("Se requiere número de COT y número de línea.");
    }

    const COL_MAP_COT = getColumnMap(HOJA_COTIZACIONES);
    const allDataCot = obtenerDatosHoja(HOJA_COTIZACIONES, false); // Datos frescos
    const CODIGO_PEDIDO_COL = COL_MAP_COT['COT'] || 0; 
    let filaEncontradaCot = null;
    let contadorLineas = 0;

    // Buscar la fila correspondiente a la línea específica en DataCot
    for (let i = 1; i < allDataCot.length; i++) {
        const row = allDataCot[i];
        if (String(row[CODIGO_PEDIDO_COL] || '').trim().toUpperCase() === String(cotNumero).trim().toUpperCase()) {
            contadorLineas++;
            if (contadorLineas === Number(cotLineaIndex)) {
                filaEncontradaCot = row;
                break;
            }
        }
    }

    if (!filaEncontradaCot) {
      throw new Error(`No se encontró la línea ${cotLineaIndex} para la cotización ${cotNumero}.`);
    }

    // Función auxiliar para obtener valores de la fila de DataCot
    const getValueCot = (colName) => getFilaValue(filaEncontradaCot, COL_MAP_COT, colName);

    // --- Extraer datos DIRECTOS de DataCot ---
    const codServicioCot = String(getValueCot('COD') || '').trim();
    const clienteNombre = String(getValueCot('CLIENTE') || '');
    const clienteRUC = String(getValueCot('ID CLIENTE') || '');
    const clienteContactoNombre = String(getValueCot('CONTACTO') || '');
    const direccion = String(getValueCot('UBICACIÓN') || '');
    const turno = String(getValueCot('TURNO') || '');
    const servicioAlcance = String(getValueCot('DESCRIPCION') || getValueCot('DESCRIPCIÓN') || '');

    // --- Buscar datos ADICIONALES (Correo y Personal Requerido) ---
    let clienteContactoCorreo = '';
    let personalRequerido = '';

    // Buscar correo del contacto (lógica existente)
    if (clienteRUC && clienteContactoNombre) {
        const infoContacto = buscarInfoContacto(clienteRUC, clienteContactoNombre);
        if (infoContacto && infoContacto.includes('@')) {
           clienteContactoCorreo = infoContacto.split('/')[0].trim(); 
        }
    }

    // Buscar Personal Requerido en Hoja Servicios
    if (codServicioCot) {
        const allDataServ = obtenerDatosHoja(HOJA_SERVICIOS, true); // Usar caché aquí es seguro
        const COL_MAP_SERV = getColumnMap(HOJA_SERVICIOS);
        const COD_SERV_COL = COL_MAP_SERV['COD']; 
        const PERSONAL_SERV_COL = COL_MAP_SERV['PERSONAL']; // Busca encabezado 'PERSONAL'

        if (COD_SERV_COL !== undefined && PERSONAL_SERV_COL !== undefined) {
             for (let i = 1; i < allDataServ.length; i++) {
                 const filaServ = allDataServ[i];
                 if (String(filaServ[COD_SERV_COL] || '').trim() === codServicioCot) {
                     personalRequerido = String(filaServ[PERSONAL_SERV_COL] || '').trim();
                     break; // Encontrado
                 }
             }
             if (!personalRequerido) {
                 Logger.log(`Advertencia (obtenerDatosBase): No se encontró valor en columna 'PERSONAL' para el servicio ${codServicioCot}.`);
             }
        } else {
             Logger.log("Advertencia (obtenerDatosBase): No se encontró encabezado 'COD' o 'PERSONAL' en Hoja Servicios.");
        }
    }

    // --- Construir el objeto de datos base ---
    const datosBase = {
      clienteNombre: clienteNombre,
      clienteRUC: clienteRUC,
      clienteContactoNombre: clienteContactoNombre, // Para "Personal a Cargo"
      clienteContactoCorreo: clienteContactoCorreo,
      fechaServicio: '', // Dejar vacío, se debe ingresar manualmente
      servicioAlcance: servicioAlcance, // Para "Alcance"
      direccion: direccion,
      turno: turno,
      equipoSolicitado: servicioAlcance, // Usar descripción como equipo inicial
      personalRequerido: personalRequerido // Dato de Hoja Servicios
      // Puedes añadir más campos si los tienes en DataCot (Plazo?, Horas?)
    };

    return { success: true, data: datosBase };

  } catch (e) {
    Logger.log(`Error en obtenerDatosBaseParaActa: ${e.message}\nStack: ${e.stack}`);
    return manejarError('obtenerDatosBaseParaActa', e);
  }
}

/**
 * NUEVA FUNCIÓN: Obtiene TODOS los datos necesarios para pre-llenar
 * el formulario de creación de Acta, incluyendo cod y desc del servicio.
 * Reemplaza la necesidad de pasar cod/desc en la URL.
 * @param {string} cotNumero Número de COT.
 * @param {number} cotLineaIndex Índice de la línea (1-based).
 * @returns {object} { success: true, data: {...} } o { success: false, error: ... }
 */
function obtenerDatosCompletosParaActaCreacion(cotNumero, cotLineaIndex) {
  Logger.log(`Buscando datos base completos para COT ${cotNumero}, Línea ${cotLineaIndex}`);
  try {
    if (!cotNumero || !cotLineaIndex) {
      throw new Error("Se requiere número de COT y número de línea.");
    }

    const COL_MAP_COT = getColumnMap(HOJA_COTIZACIONES);
    const allDataCot = obtenerDatosHoja(HOJA_COTIZACIONES, false); 
    const CODIGO_PEDIDO_COL = COL_MAP_COT['COT'] || 0; 
    let filaEncontradaCot = null;
    let contadorLineas = 0;

    // Buscar la fila específica en DataCot
    for (let i = 1; i < allDataCot.length; i++) {
        const row = allDataCot[i];
        if (String(row[CODIGO_PEDIDO_COL] || '').trim().toUpperCase() === String(cotNumero).trim().toUpperCase()) {
            contadorLineas++;
            if (contadorLineas === Number(cotLineaIndex)) {
                filaEncontradaCot = row;
                break;
            }
        }
    }

    if (!filaEncontradaCot) {
      throw new Error(`No se encontró la línea ${cotLineaIndex} para la cotización ${cotNumero}.`);
    }

    // Extraer datos de la fila de DataCot
    const getValueCot = (colName) => getFilaValue(filaEncontradaCot, COL_MAP_COT, colName);
    const codServicioCot = String(getValueCot('COD') || '').trim();
    const clienteNombre = String(getValueCot('CLIENTE') || '');
    const clienteRUC = String(getValueCot('ID CLIENTE') || '');
    const clienteContactoNombre = String(getValueCot('CONTACTO') || '');
    const direccion = String(getValueCot('UBICACIÓN') || '');
    const turno = String(getValueCot('TURNO') || '');
    const servicioAlcance = String(getValueCot('DESCRIPCION') || getValueCot('DESCRIPCIÓN') || ''); // Descripción base

    // Buscar Correo y Personal Requerido (sin cambios)
    let clienteContactoCorreo = '';
    let personalRequerido = '';
    // ... (lógica existente para buscar correo y personal usando buscarInfoContacto y hoja Servicios)...
     if (clienteRUC && clienteContactoNombre) {
        const infoContacto = buscarInfoContacto(clienteRUC, clienteContactoNombre);
        if (infoContacto && infoContacto.includes('@')) {
           clienteContactoCorreo = infoContacto.split('/')[0].trim(); 
        }
    }
     if (codServicioCot) {
        // ... (lógica existente para buscar personal en HOJA_SERVICIOS) ...
         const allDataServ = obtenerDatosHoja(HOJA_SERVICIOS, true);
         const COL_MAP_SERV = getColumnMap(HOJA_SERVICIOS);
         const COD_SERV_COL = COL_MAP_SERV['COD']; 
         const PERSONAL_SERV_COL = COL_MAP_SERV['PERSONAL'];
         if (COD_SERV_COL !== undefined && PERSONAL_SERV_COL !== undefined) {
             for (let i = 1; i < allDataServ.length; i++) {
                 const filaServ = allDataServ[i];
                 if (String(filaServ[COD_SERV_COL] || '').trim() === codServicioCot) {
                     personalRequerido = String(filaServ[PERSONAL_SERV_COL] || '').trim();
                     break; 
                 }
             }
         }
    }


    // Construir objeto final, incluyendo codServicioCot y servicioAlcance
    const datosCompletos = {
      clienteNombre: clienteNombre,
      clienteRUC: clienteRUC,
      clienteContactoNombre: clienteContactoNombre,
      clienteContactoCorreo: clienteContactoCorreo,
      fechaServicio: '', 
      servicioCod: codServicioCot, // Añadido
      servicioAlcance: servicioAlcance, 
      direccion: direccion,
      turno: turno,
      equipoSolicitado: servicioAlcance, 
      personalRequerido: personalRequerido
    };
    Logger.log("Datos base completos encontrados:", datosCompletos);
    return { success: true, data: datosCompletos };

  } catch (e) {
    Logger.log(`Error en obtenerDatosCompletosParaActaCreacion: ${e.message}\nStack: ${e.stack}`);
    // Devolver el error para que el frontend lo muestre
    return { success: false, error: e.message }; 
  }
}

/**
 * Guarda o actualiza los datos complementarios (notas) en la hoja HOJA_COMPLEMENTOS_COT.
 * Esta función borra y re-escribe la nota para la COT dada.
 */
function guardarDatosComplementarios(codigoPedido, datos) {
    try {
        const sheetName = HOJA_COMPLEMENTOS_COT; // Constante de _Constants_Lists.gs
        const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
        let sheet = ss.getSheetByName(sheetName);

        // Crear la hoja si no existe
        if (!sheet) {
            Logger.log(`Hoja ${sheetName} no encontrada, creando...`);
            sheet = ss.insertSheet(sheetName);
            sheet.appendRow(['COT', 'Plantilla', 'AclaracionesDelServicio']);
        }

        const COL_MAP = getColumnMap(sheetName);
        const allData = sheet.getDataRange().getValues();
        const COT_COL_INDEX = COL_MAP['COT'];

        if (COT_COL_INDEX === undefined) {
             Logger.log(`Error: No se encontró la columna 'COT' en ${sheetName}. No se guardarán las notas.`);
             return; // Salir si la hoja no está bien configurada
        }

        let rowIndex = -1;
        // Buscar si ya existe
        for (let i = 1; i < allData.length; i++) {
            if (String(allData[i][COT_COL_INDEX] || '').trim() === codigoPedido) {
                rowIndex = i + 1; // Fila encontrada (índice 1-based)
                break;
            }
        }

        const nuevosValores = [
            codigoPedido,
            datos.plantillaNotas || '',
            datos.aclaracionesServicio || ''
        ];

        if (rowIndex > 1) {
            // Actualizar usando crudHoja
            crudHoja('UPDATE', sheetName, { rowIndex: rowIndex, valores: nuevosValores });
        } else {
            // Crear usando crudHoja
            crudHoja('CREATE', sheetName, nuevosValores);
        }
        Logger.log(`Datos complementarios guardados para ${codigoPedido}.`);

    } catch (e) {
        Logger.log(`ERROR al guardar datos complementarios para ${codigoPedido}: ${e.message}`);
        // No lanzamos error para no detener el guardado principal
    }
}

/**
 * REEMPLAZO 1 (Backend)
 * Actualiza (acumula) los valores de despacho en la hoja DataCot.
 * ¡CORREGIDO! Ahora busca usando 'lineaID' (columna 'NUM')
 */
function actualizarDataCotDesdeOT(lineaID, horasDespachadas, montoDespachado, montoMovilizacion) {
    Logger.log(`INICIO actualizarDataCotDesdeOT: LineaID=${lineaID}, Horas=${horasDespachadas}, Monto=${montoDespachado}, Mov=${montoMovilizacion}`);
    
    if (!lineaID) {
        Logger.log("ERROR: Faltó lineaID para actualizar DataCot.");
        return false;
    }

    // Modificamos esta lógica: si las horas son 0, PERO los montos no lo son, debe continuar.
    if (horasDespachadas === 0 && montoDespachado === 0 && montoMovilizacion === 0) {
        Logger.log("Advertencia: No hay valores (horas o montos) para actualizar en DataCot. Saliendo.");
        return true; 
    }

    try {
        const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
        const sheet = ss.getSheetByName(HOJA_COTIZACIONES);
        if (!sheet) {
            Logger.log(`ERROR FATAL: No se encontró la hoja ${HOJA_COTIZACIONES}`);
            return false;
        }

        Logger.log("Paso 1: Obteniendo mapa de columnas para " + HOJA_COTIZACIONES);
        const COL_MAP = getColumnMap(HOJA_COTIZACIONES);
        Logger.log("Mapa de DataCot: " + JSON.stringify(COL_MAP));

        // Validación de columnas clave
        const LINEAID_COL = COL_MAP['NUM'];
        const UND_DESPACHO_COL = COL_MAP['UND. DESPACHO'];
        const M_DESPACHO_COL = COL_MAP['M. DESPACHO'];
        const MYDM_VALORIZADA_COL = COL_MAP['MYDM VALORIZADA'];

        if ([LINEAID_COL, UND_DESPACHO_COL, M_DESPACHO_COL, MYDM_VALORIZADA_COL].includes(undefined)) {
            Logger.log("ERROR FATAL: Faltan columnas clave en DataCot. Revisa encabezados: 'NUM', 'UND. DESPACHO', 'M. DESPACHO', 'MYDM VALORIZADA'");
            Logger.log(`Valores actuales: NUM=${LINEAID_COL}, UND. DESPACHO=${UND_DESPACHO_COL}, M. DESPACHO=${M_DESPACHO_COL}, MYDM VALORIZADA=${MYDM_VALORIZADA_COL}`);
            return false;
        }

        Logger.log("Paso 2: Leyendo datos de DataCot para buscar la fila...");
        const allData = sheet.getDataRange().getValues();
        let rowIndexToUpdate = -1;

        for (let i = 1; i < allData.length; i++) {
            const idEnFila = String(allData[i][LINEAID_COL] || '').trim();
            if (idEnFila === lineaID) {
                rowIndexToUpdate = i + 1;
                Logger.log(`Paso 3: Fila encontrada. Coincide en la Fila ${rowIndexToUpdate}`);
                break;
            }
        }

        if (rowIndexToUpdate === -1) {
            Logger.log(`ERROR FATAL: No se encontró la fila en DataCot para LineaID=${lineaID}.`);
            return false;
        }

        // --- Lógica de Actualización ---
        Logger.log("Paso 4: Leyendo valores actuales de la Fila " + rowIndexToUpdate);
        
        // (Col + 1 porque getRange es 1-based, y los índices del mapa son 0-based)
        const undDespachoActual = parseFloat(sheet.getRange(rowIndexToUpdate, UND_DESPACHO_COL + 1).getValue() || 0);
        const mDespachoActual = parseFloat(sheet.getRange(rowIndexToUpdate, M_DESPACHO_COL + 1).getValue() || 0);
        const mydmValorizadaActual = parseFloat(sheet.getRange(rowIndexToUpdate, MYDM_VALORIZADA_COL + 1).getValue() || 0);

        Logger.log(`Valores Actuales: Horas=${undDespachoActual}, Monto=${mDespachoActual}, Mov=${mydmValorizadaActual}`);

        const nuevasUndDespacho = undDespachoActual + horasDespachadas;
        const nuevoMDespacho = mDespachoActual + montoDespachado;
        const nuevaMydmValorizada = mydmValorizadaActual + montoMovilizacion;

        Logger.log(`Paso 5: Escribiendo nuevos valores... Horas=${nuevasUndDespacho}, Monto=${nuevoMDespacho}, Mov=${nuevaMydmValorizada}`);
        
        sheet.getRange(rowIndexToUpdate, UND_DESPACHO_COL + 1).setValue(nuevasUndDespacho);
        sheet.getRange(rowIndexToUpdate, M_DESPACHO_COL + 1).setValue(nuevoMDespacho);
        sheet.getRange(rowIndexToUpdate, MYDM_VALORIZADA_COL + 1).setValue(nuevaMydmValorizada);
        
        Logger.log("ÉXITO: DataCot actualizada exitosamente para Fila " + rowIndexToUpdate);
        return true;
        
    } catch (e) {
        Logger.log(`ERROR CRÍTICO en actualizarDataCotDesdeOT: ${e.message}\nStack: ${e.stack}`);
        enviarNotificacionError(`Error en actualizarDataCotDesdeOT: ${e.message}`);
        return false;
    }
}

/**
 * REEMPLAZO (v2 - Supabase): Obtiene el resumen de líneas de un pedido.
 * Lee directamente de 'Detalle_Pedidos' y sus tablas relacionadas.
 * Es mucho más rápido porque los totales (UND_DESPACHO) ya están calculados.
 */
function getResumenOTPorPedido(numPedido) {
  Logger.log(`INICIO getResumenOTPorPedido (v2 - Supabase) para: ${numPedido}`);
  try {
    if (!numPedido) throw new Error("Número de pedido no proporcionado.");

    // 1. Construir la consulta anidada
    const consulta = `
      select=
        Cot,
        Cot_Linea_Ref,
        Cantidad,
        UND,
        UND_DESPACHO,
        Servicios!inner(ID_servicios, Nombre_Servicio),
        Pedidos!inner(RUC_DNI, Clientes!inner(Nombre_RazonSocial))
      &Cot=eq.${numPedido}
    `.replace(/\s/g, '');

    // 2. Obtener los detalles
    const lineasDetalle = supabaseFetch('Detalle_Pedidos', {
      method: 'get',
      params: consulta
    });

    if (!lineasDetalle || lineasDetalle.length === 0) {
      throw new Error("No se encontraron líneas para el pedido: " + numPedido);
    }

    // 3. Mapear los datos al formato que espera el frontend
    // Obtenemos el nombre del cliente desde la *primera* línea
    let clienteNombre = lineasDetalle[0].Pedidos.Clientes.Nombre_RazonSocial;

    const lineasMapeadas = lineasDetalle.map((linea, index) => {
      // Asumimos que 'horasPedidas' es 'Cantidad' si la unidad es HORAS o DÍAS
      let horasPedidas = (linea.UND === 'HORAS' || linea.UND === 'DÍAS') ? linea.Cantidad : 0;

      return {
        lineaID: linea.Cot_Linea_Ref, // El ID de texto (COT...-L1)
        cod: linea.Servicios.ID_servicios,
        descripcion: linea.Servicios.Nombre_Servicio,
        horasPedidas: horasPedidas,
        horasDespachadas: parseFloat(linea.UND_DESPACHO || 0)
      };
    });

    Logger.log(`ÉXITO: Se encontraron ${lineasMapeadas.length} líneas.`);
    return { 
      success: true, 
      pedido: numPedido,
      cliente: clienteNombre,
      lineas: lineasMapeadas
    };

  } catch (e) {
    Logger.log(`ERROR FATAL en getResumenOTPorPedido (v2): ${e.message} \n Stack: ${e.stack}`);
    return manejarError('getResumenOTPorPedido', e);
  }
}

/**
 * Obtiene los datos de una LÍNEA específica de DataCot para pre-llenar una OT.
 */
function getDatosDeLineaParaOT(lineaID) {
  try {
    if (!lineaID) throw new Error("ID de Línea no proporcionado.");
    
    const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
    const sheet = ss.getSheetByName(HOJA_COTIZACIONES);
    const mapCot = getColumnMap(HOJA_COTIZACIONES);
    const allData = sheet.getDataRange().getValues();

    const LINEAID_COL = mapCot['NUM'];
    if (LINEAID_COL === undefined) throw new Error("No se encontró la columna 'NUM' en DataCot.");

    let filaEncontrada = null;
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][LINEAID_COL] || '').trim() === lineaID) {
        filaEncontrada = allData[i];
        break;
      }
    }

    if (!filaEncontrada) throw new Error("No se encontró la línea con ID: " + lineaID);

    // Extraer los datos que el formulario de OT necesita
    const datosLinea = {
      lineaID: lineaID,
      numPedido: String(filaEncontrada[mapCot['COT']] || ''),
      clienteRUC: String(filaEncontrada[mapCot['ID CLIENTE']] || ''),
      clienteNombre: String(filaEncontrada[mapCot['CLIENTE']] || ''),
      codServicio: String(filaEncontrada[mapCot['COD']] || ''),
      descServicio: String(filaEncontrada[mapCot['DESCRIPCION']] || '')
    };

    return { success: true, data: datosLinea };

  } catch (e) {
    Logger.log(`Error en getDatosDeLineaParaOT: ${e.message}`);
    return manejarError('getDatosDeLineaParaOT', e);
  }
}

/**
 * Elimina permanentemente una Orden de Trabajo por su N° OT.
 * @param {string} numeroOT El N° OT a eliminar (ej. "OT-001").
 * @returns {object} { success: true, message: "..." } o { success: false, message: "..." }
 */
function eliminarOTPorID(numeroOT) {
    // 1. Verificar Permiso
    const permisos = obtenerPermisosUsuario();
    if (!permisos.puedeEditarServicios) { // Usa el mismo permiso que para guardar/editar OT
        return { success: false, message: "Acceso denegado. No tiene permiso para eliminar OTs." };
    }

    if (!numeroOT) {
        return { success: false, message: "No se proporcionó un N° de OT." };
    }

    try {
        const OT_SHEET = HOJA_OT;
        const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
        const sheet = ss.getSheetByName(OT_SHEET);
        if (!sheet) throw new Error(`Hoja ${OT_SHEET} no encontrada.`);

        const COL_MAP = getColumnMap(OT_SHEET);
        const allData = sheet.getDataRange().getValues();
        const OT_COL_INDEX = COL_MAP['N° OT']; // Columna "N° OT"

        if (OT_COL_INDEX === undefined) {
            throw new Error(`No se encontró la columna 'N° OT' en la hoja ${OT_SHEET}.`);
        }

        let rowIndexToDelete = -1;
        
        // 2. Buscar la fila
        for (let i = 1; i < allData.length; i++) { // Empezar en 1 (datos)
            const otEnFila = String(allData[i][OT_COL_INDEX] || '').trim();
            if (otEnFila === numeroOT) {
                rowIndexToDelete = i + 1; // i es 0-based, +1 para el índice de fila real
                break;
            }
        }

        // 3. Eliminar la fila
        if (rowIndexToDelete > 1) { // Mayor a 1 (fila de encabezado)
            sheet.deleteRow(rowIndexToDelete);
            SpreadsheetApp.flush(); // Forzar guardado
            
            // Invalidar caché
            cache.remove(`hoja_${OT_SHEET}`);
            
            Logger.log(`OT ${numeroOT} eliminada exitosamente (Fila ${rowIndexToDelete}).`);
            return { success: true, message: `OT ${numeroOT} eliminada exitosamente.` };
        } else {
            throw new Error(`No se encontró ninguna OT con el número ${numeroOT}.`);
        }

    } catch (e) {
        Logger.log(`Error en eliminarOTPorID: ${e.message}`);
        // Usar manejarError para devolver un objeto de error estándar
        return manejarError('eliminarOTPorID', e);
    }
}

/**
 * Obtiene los datos de una línea específica Y TAMBIÉN la lista de todos 
 * los servicios asociados con el pedido de esa línea.
 * @param {string} lineaID El ID de la línea (ej. "COT.ALP...-L1").
 * @returns {object} { success: true, dataLinea: {...}, serviciosDelPedido: [...] }
 */
function getDatosDePedidoParaOT(lineaID) {
  try {
    if (!lineaID) throw new Error("ID de Línea no proporcionado.");
    
    // 1. Obtener datos de la línea específica (reutilizando tu función existente)
    const resultadoLinea = getDatosDeLineaParaOT(lineaID);
    if (!resultadoLinea.success) {
      throw new Error("No se pudieron encontrar los datos de la línea original: " + lineaID);
    }
    
    const datosLinea = resultadoLinea.data;
    const numPedido = datosLinea.numPedido;
    
    if (!numPedido) {
      throw new Error("La línea " + lineaID + " no tiene un número de pedido asociado.");
    }

    // 2. Buscar TODOS los servicios para ese numPedido
    const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
    const sheet = ss.getSheetByName(HOJA_COTIZACIONES);
    const mapCot = getColumnMap(HOJA_COTIZACIONES);
    const allData = sheet.getDataRange().getValues();

    const COT_COL_COT = mapCot['COT'];
    const COD_COL_COT = mapCot['COD'];
    const DESC_COL_COT = mapCot['DESCRIPCION'] || mapCot['DESCRIPCIÓN']; // Soporta ambos
    
    if (COT_COL_COT === undefined || COD_COL_COT === undefined || DESC_COL_COT === undefined) {
      throw new Error("Faltan columnas 'COT', 'COD' o 'DESCRIPCION' en DataCot.");
    }

    const serviciosDelPedido = [];
    const codigosVistos = new Set(); // Para evitar duplicados si hay varias líneas del mismo servicio

    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      if (String(row[COT_COL_COT] || '').trim() === numPedido) {
        
        const codServicio = String(row[COD_COL_COT] || '');
        
        if (codServicio && !codigosVistos.has(codServicio)) {
          serviciosDelPedido.push({
            cod: codServicio,
            desc: String(row[DESC_COL_COT] || '')
          });
          codigosVistos.add(codServicio);
        }
      }
    }

    Logger.log(`Encontrados ${serviciosDelPedido.length} servicios únicos para el pedido ${numPedido}.`);

    // 3. Devolver ambos resultados
    return { 
      success: true, 
      dataLinea: datosLinea, // Datos de la línea específica
      serviciosDelPedido: serviciosDelPedido // Lista de TODOS los servicios del pedido
    };

  } catch (e) {
    Logger.log(`Error en getDatosDePedidoParaOT: ${e.message}`);
    return manejarError('getDatosDePedidoParaOT', e);
  }
}

/**
 * Busca en DataCot una lineaID específica y devuelve el valor de su
 * columna "HORAS SEGÚN".
 * @param {string} lineaID El ID único de la línea (ej. "COT...-L1").
 * @returns {string} El valor de "HORAS SEGÚN" (ej. "HORÓMETRO" o "INICIO/FIN"), o una cadena vacía.
 */
function obtenerHorasSegunPorLineaID(lineaID) {
  if (!lineaID) return '';

  try {
    const sheetName = HOJA_COTIZACIONES;
    const mapCot = getColumnMap(sheetName);
    const allData = obtenerDatosHoja(sheetName, false); // Leer sin caché para asegurar datos frescos

    const LINEAID_COL = mapCot['NUM'];
    const HORAS_SEGUN_COL = mapCot['HORAS SEGÚN'];

    if (LINEAID_COL === undefined || HORAS_SEGUN_COL === undefined) {
      Logger.log(`ADVERTENCIA (obtenerHorasSegun): No se encontraron las columnas 'NUM' o 'HORAS SEGÚN' en ${sheetName}.`);
      return '';
    }

    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][LINEAID_COL] || '').trim() === lineaID) {
        const horasSegun = String(allData[i][HORAS_SEGUN_COL] || '').trim().toUpperCase();
        Logger.log(`Horas Según encontrado para ${lineaID}: ${horasSegun}`);
        return horasSegun;
      }
    }
    
    Logger.log(`ADVERTENCIA (obtenerHorasSegun): No se encontró la línea ${lineaID} en ${sheetName}.`);
    return ''; // No se encontró la línea

  } catch (e) {
    Logger.log(`ERROR en obtenerHorasSegunPorLineaID: ${e.message}`);
    return ''; // Devolver vacío en caso de error
  }
}
/**
 * GUARDA LA COTIZACIÓN EN SUPABASE (v1.1 - CORREGIDA CON Cot_Linea_Ref)
 * PASO 1 (Crear): Inserta la cabecera y luego las líneas.
 * PASO 1 (Actualizar): Llama a la RPC 'actualizar_cotizacion_y_detalles'
 */
function guardarCotizacion(datos) {
  const permisos = obtenerPermisosUsuario();
  if (!permisos.puedeEditarCotizacion) {
    return { success: false, message: "Acceso denegado." };
  }

  const datosSanitizados = sanitizarDatos(datos);
  const esModoUpdate = !!datosSanitizados.numPedido;
  
  try {
    let codigoPedido;
    
    // --- Preparar las líneas PRIMERO ---
    const lineas = datosSanitizados.Lineas || [];
    if (lineas.length === 0) throw new Error("Debe agregar al menos un servicio.");
    
    //=====================================================
    // MODO ACTUALIZACIÓN (PATCH) - ¡RPC!
    //=====================================================
    if (esModoUpdate) {
      codigoPedido = datosSanitizados.numPedido;
      Logger.log(`Iniciando MODO UPDATE (RPC) para: ${codigoPedido}`);

      // ---- PASO 1: Preparar la cabecera (JSON) ----
      const payloadCabecera = {
        Estado_Cot: datosSanitizados.Estado,
        Total_Cot: parseFloat(datosSanitizados.Total),
        Moneda: datosSanitizados.Moneda,
        Ejecutivo: datosSanitizados.Ejecutivo,
        Fecha_Inicio: datosSanitizados.fechaEjecucion, 
        Forma_De_Pago: datosSanitizados.Forma_Pago,
        Empresa: datosSanitizados.Empresa,
        RUC_DNI: datosSanitizados.RUC, // <-- CORREGIDO A RUC_DNI
        Direccion: datosSanitizados.Direccion, // <-- Columna nueva
        Turno: datosSanitizados.Turno         // <-- Columna nueva
      };

      // ---- PASO 2: Preparar las líneas (JSON) ----
      const payloadDetalles = [];
      for (let i = 0; i < lineas.length; i++) {
        const linea = lineas[i];
        const lineaIDRef = `${codigoPedido}-L${i + 1}`; // ID de Texto
        
        const detalle = {
          "Cot_Linea_Ref": lineaIDRef, // <-- NUEVA COLUMNA DE TEXTO
          "Cot": codigoPedido,
          "ID_Servicio": linea.cod,
          "Cantidad": linea.cantidad,
          "Precio": linea.precio,
          "Monto_Movilizacion": linea.movilizacion,
          "Horas_Segun": linea.hora_segun,
          "UND": linea.und_medida,
          "UND_HORAS_MINIMAS": linea.und_horas_minimas,
          "Horas_minimas": linea.horas_minimas_num
        };
        payloadDetalles.push(detalle);
      }

      // ---- PASO 3: Llamar a la función RPC ----
      const payloadRPC = {
        "codigo_pedido": codigoPedido,
        "datos_cabecera": payloadCabecera,
        "nuevas_lineas": payloadDetalles
      };
      
      Logger.log("Llamando a RPC 'actualizar_cotizacion_y_detalles'...");
      supabaseFetch('rpc/actualizar_cotizacion_y_detalles', {
        method: 'post',
        payload: payloadRPC
      });
      Logger.log(`Cotización ${codigoPedido} actualizada vía RPC.`);

    } 
    //=====================================================
    // MODO CREACIÓN (POST)
    //=====================================================
    else {
      Logger.log("Iniciando MODO CREATE...");
      codigoPedido = generarCodigoPedido(datosSanitizados.Empresa);
      Logger.log(`Nuevo código generado: ${codigoPedido}`);
      
      // ---- PASO 1 (CREATE): Insertar la cabecera ----
      const payloadPedido = {
        Cot: codigoPedido,
        Fecha_Creacion: new Date().toISOString(),
        Estado_Cot: datosSanitizados.Estado,
        Total_Cot: parseFloat(datosSanitizados.Total),
        Moneda: datosSanitizados.Moneda,
        Ejecutivo: datosSanitizados.Ejecutivo,
        Fecha_Inicio: datosSanitizados.fechaEjecucion,
        Forma_De_Pago: datosSanitizados.Forma_De_Pago,
        Empresa: datosSanitizados.Empresa,
        RUC_DNI: datosSanitizados.RUC, // <-- CORREGIDO A RUC_DNI
        Direccion: datosSanitizados.Direccion, // <-- Columna nueva
        Turno: datosSanitizados.Turno         // <-- Columna nueva
      };
      
      supabaseFetch('Pedidos', {
        method: 'post',
        payload: payloadPedido
      });
      Logger.log(`Cabecera ${codigoPedido} creada.`);

      // ---- PASO 2 (CREATE): Preparar e Insertar líneas ----
      const payloadDetalles = [];
      for (let i = 0; i < lineas.length; i++) {
        const linea = lineas[i];
        const lineaIDRef = `${codigoPedido}-L${i + 1}`; // ID de Texto
        
        const detalle = {
          "Cot_Linea_Ref": lineaIDRef, // <-- NUEVA COLUMNA DE TEXTO
          "Cot": codigoPedido,
          "ID_Servicio": linea.cod,
          "Cantidad": linea.cantidad,
          "Precio": linea.precio,
          "Monto_Movilizacion": linea.movilizacion,
          "Horas_Segun": linea.hora_segun,
          "UND": linea.und_medida,
          "UND_HORAS_MINIMAS": linea.und_horas_minimas,
          "Horas_minimas": linea.horas_minimas_num
        };
        payloadDetalles.push(detalle);
      }
      
      supabaseFetch('Detalle_Pedidos', {
        method: 'post',
        payload: payloadDetalles
      });
      Logger.log(`${payloadDetalles.length} líneas de detalle guardadas para ${codigoPedido}.`);
    }
    
    return { 
      success: true, 
      message: esModoUpdate ? "Cotización actualizada" : "Cotización registrada",
      codigoPedido: codigoPedido 
    };
    
  } catch (error) {
    Logger.log(`ERROR en guardarCotizacion (v1.1): ${error.message}\nStack: ${error.stack}`);
    return manejarError('guardarCotizacion', error);
  }
}

/**
 * REFACTORIZADO (v2.1 - CORREGIDO CON RUC_DNI): Obtiene los datos para editar un pedido
 * en UNA SOLA LLAMADA a Supabase, usando joins anidados basados en el esquema real.
 */
function obtenerPedidoParaEdicion(numPedido) {
  Logger.log(`Iniciando obtenerPedidoParaEdicion (v2.1) para: ${numPedido}`);
  try {
    // 1. Construir la consulta anidada.
    // La magia ocurre gracias a las Foreign Keys que creaste.
    // Pedimos todo de 'Pedidos', y Supabase sabe cómo "unir"
    // Clientes, Contactos, y Detalle_Pedidos (con sus Servicios).
    const consulta = `
      select=
        *,
        Clientes!inner(RUC_DNI, Nombre_RazonSocial, Direccion_Fiscal),
        Contactos(ID_Contacto, Nombre_Contacto),
        Detalle_Pedidos (
          *,
          Servicios!inner(ID_servicios, Nombre_Servicio)
        )
      &Cot=eq.${numPedido}
    `.replace(/\s/g, ''); // Limpiar espacios y saltos de línea para la URL

    Logger.log(`Ejecutando consulta anidada: ${consulta}`);

    // 2. Ejecutar la llamada ÚNICA a Supabase
    const resultado = supabaseFetch('Pedidos', {
      method: 'get',
      params: consulta
    });
    
    if (!resultado || resultado.length === 0) {
      throw new Error(`Pedido ${numPedido} no encontrado.`);
    }
    
    const pedido = resultado[0]; // Solo esperamos un resultado

    // 3. Mapear los datos de Supabase al formato que espera tu HTML
    const datosGenerales = {
      fechaEjecucion: pedido.Fecha_Inicio ? pedido.Fecha_Inicio.split('T')[0] : '',
      Estado: pedido.Estado_Cot,
      Empresa: pedido.Empresa,
      RUC: pedido.RUC_DNI, // <-- Usamos el nombre de columna correcto
      Cliente: pedido.Clientes.Nombre_RazonSocial, // Usamos !inner, así que 'Clientes' siempre existirá
      Moneda: pedido.Moneda,
      Forma_De_Pago: pedido.Forma_De_Pago,
      // Usamos la 'Direccion' del pedido (que añadiste) o la 'Direccion_Fiscal' del cliente como fallback
      Direccion: pedido.Direccion || (pedido.Clientes ? pedido.Clientes.Direccion_Fiscal : ''),
      Turno: pedido.Turno || '', // Leerá la nueva columna 'Turno' que añadiste
      Ejecutivo: pedido.Ejecutivo,
      Contacto: pedido.Contactos ? pedido.Contactos.Nombre_Contacto : '' // Contactos puede ser nulo
    };

    // 4. Mapear líneas de detalle
    const lineas = pedido.Detalle_Pedidos.map(linea => {
      const subtotal = (linea.Cantidad * linea.Precio) + linea.Monto_Movilizacion;
      return {
        cod: linea.ID_Servicio,
        descripcion: linea.Servicios.Nombre_Servicio, // Usamos !inner, 'Servicios' siempre existirá
        cantidad: linea.Cantidad,
        und_medida: linea.UND,
        und_horas_minimas: linea.UND_HORAS_MINIMAS,
        dias_cotizados: 0, // Calculado en frontend, no se almacena
        horas_minimas_num: linea.Horas_minimas,
        hora_segun: linea.Horas_Segun,
        movilizacion: linea.Monto_Movilizacion,
        precio: linea.Precio,
        subtotal: subtotal
      };
    });
    
    // 5. Devolver el objeto combinado
    const resultadoFinal = { 
      success: true, 
      ...datosGenerales, 
      Lineas: lineas, 
      Total: pedido.Total_Cot, 
      numPedido: numPedido,
      // (Opcional) Leer las notas si las añadiste a la tabla Pedidos
      plantillaNotas: pedido.plantillaNotas || '',
      aclaracionesServicio: pedido.aclaracionesServicio || ''
    };
    
    Logger.log(`Éxito: Pedido ${numPedido} cargado en una sola llamada.`);
    return resultadoFinal;
        
  } catch (error) {
    Logger.log(`❌ ERROR CRÍTICO al extraer pedido ${numPedido} (v2.1): ${error.message}\nStack: ${error.stack}`);
    return manejarError('obtenerPedidoParaEdicion', error); // Llama a tu manejador de errores
  }
}

/**
 * REFACTORIZADO: Obtiene la lista de clientes de Supabase
 */
function getListaClientes() {
  try {
    const clientes = supabaseFetch('Clientes', { 
      method: 'get',
      params: 'select=RUC_DNI,Nombre_RazonSocial' 
    });
    
    // El frontend (Contactos.html) espera un Array de Arrays, 
    // pero nuestro nuevo frontend (Comercial.html) preferirá JSON.
    // Por ahora, lo mantenemos como JSON (array de objetos).
    return clientes;

  } catch (e) {
    return manejarError('getListaClientes', e); 
  }
}

/**
 * REFACTORIZADO: Guarda un nuevo cliente en Supabase
 */
function guardarNuevoCliente(data) {
  try {
    const datosSanitizados = sanitizarDatos(data); //

    const nuevoCliente = {
      RUC_DNI: datosSanitizados.RUC,
      Nombre_RazonSocial: datosSanitizados.NOMBRE
    };

    const resultado = supabaseFetch('Clientes', {
      method: 'post',
      payload: nuevoCliente,
      params: 'select=RUC_DNI,Nombre_RazonSocial'
    });

    return { 
      success: true, 
      message: "Cliente registrado exitosamente.",
      nuevoCliente: resultado[0] 
    };
  } catch (error) {
    return manejarError("guardarNuevoCliente", error); //
  }
}

/**
 * REFACTORIZADO: Guarda o actualiza un Contacto en Supabase
 */
function guardarOActualizarContacto(data) {
  try {
    const datosSanitizados = sanitizarDatos(data); //
    
    // 'rowIndex' ahora es 'ID_Contacto'
    const idContacto = datosSanitizados.rowIndex; 

    // Mapeo a 'Contactos_rows.csv'
    const payload = {
      RUC_DNI: datosSanitizados.RUC,
      Nombre_Contacto: datosSanitizados.NOMBRE,
      Correo: datosSanitizados.EMAIL,
      Celular: datosSanitizados.TELEFONO,
      Cargo: datosSanitizados.CARGO
    };

    let metodo;
    let params;

    if (idContacto && parseInt(idContacto) > 0) {
      metodo = 'patch';
      params = `ID_Contacto=eq.${idContacto}`;
    } else {
      metodo = 'post';
      params = '';
    }

    supabaseFetch('Contactos', { 
      method: metodo,
      payload: payload,
      params: params
    });

    return { success: true, message: "Contacto guardado exitosamente" }; //

  } catch (error) {
    return manejarError("guardarOActualizarContacto", error); //
  }
}
