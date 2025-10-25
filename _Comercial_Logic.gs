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
        // Se asume la existencia de hoja 'Pedidos'
        return {
            clientes: obtenerDatosHoja(HOJA_CLIENTES),
            servicios: obtenerDatosHoja(HOJA_SERVICIOS),
            pedidos: obtenerDatosHoja(HOJA_PEDIDOS_OT) // Usa la hoja de pedidos para OT
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

/**
 * Obtiene la lista completa de clientes usando la función modular.
 */
function getListaClientes() {
    try {
        const data = obtenerDatosHoja(HOJA_CLIENTES);
        if (data.length <= 1) {
             return [['RUC', 'NOMBRE/RAZÓN SOCIAL']]; 
        }
        return data;
    } catch (e) {
        return manejarError('getListaClientes', e);
    }
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
            listaEjecutivos: LISTA_EJECUTIVOS
        };
    } catch (e) {
        return manejarError('getDatosInicialesComercial', e);
    }
}

function getDatosPesadosComercial() {
    try {
        const serviciosData = obtenerDatosHoja(HOJA_SERVICIOS);
        return {
            success: true,
            clientes: obtenerDatosHoja(HOJA_CLIENTES),
            servicios: serviciosData,
            listaUndMedida: getListaValoresUnicosOptimizada(serviciosData, 5), 
            listaHorasSegun: getListaValoresUnicosOptimizada(serviciosData, 4)
        };
    } catch (e) {
        return manejarError('getDatosPesadosComercial', e);
    }
}

/**
 * Obtiene una lista simplificada de contactos para llenar el select en Comercial.html
 */
function getContactosParaComercial(ruc) {
    try {
        const allData = obtenerDatosHoja(HOJA_CONTACTOS);
        if (allData.length <= 1) return [];

        const COL_MAP = getColumnMap(HOJA_CONTACTOS);
        const RUC_COL = COL_MAP['RUC'];
        const NOMBRE_COL = COL_MAP['NOMBRE'];
        const CARGO_COL = COL_MAP['CARGO'];
        const EMAIL_COL = COL_MAP['EMAIL'];
        const TELEFONO_COL = COL_MAP['TELEFONO'];

        if (RUC_COL === undefined || NOMBRE_COL === undefined) return [];

        const contactos = [];
        const rucBuscado = String(ruc).trim();

        for (let i = 1; i < allData.length; i++) {
            const row = allData[i];
            if (String(row[RUC_COL] || '').trim() === rucBuscado) {
                const nombre = String(row[NOMBRE_COL] || '').trim();
                const cargo = String(row[CARGO_COL] || '').trim();
                
                contactos.push({
                    nombre: nombre,
                    cargo: cargo,
                    email: String(row[EMAIL_COL] || '').trim(),
                    telefono: String(row[TELEFONO_COL] || '').trim(),
                    display: `${nombre}${cargo ? ' - ' + cargo : ''}`.trim()
                });
            }
        }
        return contactos;
    } catch (error) {
        return [];
    }
}


// ====================================================
// === GUARDADO Y EDICIÓN DE COTIZACIÓN (CON MAPEO) ===
// ====================================================

function guardarCotizacion(datos) {
    // Implementación usando getColumnMap (de _Core.gs) y lógica de eliminación/creación
    try {
        const datosSanitizados = sanitizarDatos(datos);
        const erroresValidacion = validarDatosCotizacion(datosSanitizados);
        
        if (erroresValidacion.length > 0) {
            return { success: false, message: "Errores de validación: " + erroresValidacion.join(', ') };
        }

        const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
        let hojaCot = ss.getSheetByName(HOJA_COTIZACIONES);
        
        if (!hojaCot) {
            hojaCot = ss.insertSheet(HOJA_COTIZACIONES);
            const encabezados = [
                'COT', 'NUM', 'FECHA COT', 'EMPRESA', 'EJECUTIVO', 'ID CLIENTE', 'CLIENTE', 'ESTADO COT', 
                'COD', 'DESCRIPCION', 'UND', 'UND. PEDIDO', 'UND. DESPACHO', 'UND. PENDIENTE', 'PRECIO', 
                'MONEDA', 'M. PEDIDO', 'M. DESPACHO', 'M. PENDIENTE', 'MOV. Y DES. MOV.', 'MYDM VALORIZADA', 
                'MYDM x VALORIZAR', 'Total servicio', 'Total Valorizado', 'Total por valorizar', 'TURNO', 
                'FECHA INICIO', 'FECHA FIN', 'PLACA', 'UND. HORAS. MINIMAS', 'HORAS SEGÚN', 'HORAS MINIMAS', 
                'TOTAL HORAS MINIMAS', 'TOTAL DIAS', 'ESTADO DE SERVICIO', 'ACTA', 'VALORIZADO', 'FACTURA', 
                'F. PAGO', 'Link COT', 'Link Act', 'Link Val', 'Link Factura', 'Observaciones', 'UBICACIÓN', 
                'CONTACTO', 'MES', 'AÑO'
            ];
            hojaCot.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
        }

        let codigoPedido;
        const COL_MAP = getColumnMap(HOJA_COTIZACIONES);

        if (datosSanitizados.numPedido) {
            const allData = hojaCot.getDataRange().getValues();
            const CODIGO_COL_INDEX = COL_MAP['COT'] || 0; 
            const filasAEliminar = [];
            for (let i = 1; i < allData.length; i++) {
                if (String(allData[i][CODIGO_COL_INDEX] || '').trim() === datosSanitizados.numPedido) {
                    filasAEliminar.push(i + 1);
                }
            }
            if (filasAEliminar.length > 0) {
                filasAEliminar.sort((a, b) => b - a).forEach(rowIndex => hojaCot.deleteRow(rowIndex));
            }
            codigoPedido = datosSanitizados.numPedido;
        } else {
            let intentos = 0;
            do {
                codigoPedido = generarCodigoPedido(datosSanitizados.Empresa);
                intentos++;
            } while (!verificarCodigoUnico(codigoPedido) && intentos < 5);
            if (intentos >= 5) throw new Error("No se pudo generar un código único.");
        }

        const fechaRegistro = new Date(datosSanitizados.fechaRegistro || new Date());
        const lineas = datosSanitizados.Lineas || [];
        if (lineas.length === 0) throw new Error("Debe agregar al menos un servicio a la cotización");

        const TOTAL_COLS = 48; 
        let filaActual = hojaCot.getLastRow() + 1;
        
        for (let i = 0; i < lineas.length; i++) {
            const linea = lineas[i];
            const filaCompleta = new Array(TOTAL_COLS).fill('');
            
            // ASIGNACIÓN DE DATOS GENERALES USANDO MAPEO (CRÍTICO)
            filaCompleta[COL_MAP['COT']] = codigoPedido;
            filaCompleta[COL_MAP['FECHA COT']] = fechaRegistro;
            filaCompleta[COL_MAP['EMPRESA']] = datosSanitizados.Empresa || '';
            filaCompleta[COL_MAP['EJECUTIVO']] = datosSanitizados.Ejecutivo || '';
            filaCompleta[COL_MAP['ID CLIENTE']] = datosSanitizados.RUC || '';
            filaCompleta[COL_MAP['CLIENTE']] = datosSanitizados.Cliente || '';
            filaCompleta[COL_MAP['ESTADO COT']] = datosSanitizados.Estado || 'COTIZACION';
            filaCompleta[COL_MAP['MONEDA']] = datosSanitizados.Moneda || 'SOLES';
            filaCompleta[COL_MAP['TURNO']] = datosSanitizados.Turno || 'Diurno';
            filaCompleta[COL_MAP['F. PAGO']] = datosSanitizados.Forma_Pago || '';
            filaCompleta[COL_MAP['UBICACIÓN']] = datosSanitizados.Direccion || ''; 
            filaCompleta[COL_MAP['CONTACTO']] = datosSanitizados.Contacto || '';
            
            if (i === 0) filaCompleta[COL_MAP['Total servicio']] = parseFloat(datosSanitizados.Total) || 0;

            // ASIGNACIÓN DE DATOS DE LÍNEA USANDO MAPEO
            filaCompleta[COL_MAP['COD']] = linea.cod || '';
            filaCompleta[COL_MAP['DESCRIPCION']] = linea.descripcion || '';
            filaCompleta[COL_MAP['UND']] = linea.und_medida || 'HORAS';
            filaCompleta[COL_MAP['UND. PEDIDO']] = parseFloat(linea.cantidad) || 0;
            filaCompleta[COL_MAP['PRECIO']] = parseFloat(linea.precio) || 0;
            filaCompleta[COL_MAP['M. PEDIDO']] = parseFloat(linea.subtotal) || 0;
            filaCompleta[COL_MAP['MOV. Y DES. MOV.']] = parseFloat(linea.movilizacion) || 0;
            filaCompleta[COL_MAP['UND. HORAS. MINIMAS']] = linea.und_horas_minimas || '';
            filaCompleta[COL_MAP['HORAS SEGÚN']] = linea.hora_segun || '';
            filaCompleta[COL_MAP['HORAS MINIMAS']] = parseFloat(linea.horas_minimas_num) || 0;
            filaCompleta[COL_MAP['TOTAL DIAS']] = parseFloat(linea.dias_cotizados) || 0;

            hojaCot.getRange(filaActual, 1, 1, filaCompleta.length).setValues([filaCompleta]);
            filaActual++;
        }

        return { success: true, message: datosSanitizados.numPedido ? 'Cotización actualizada con éxito. Código: ' + codigoPedido : 'Cotización registrada con éxito. Código: ' + codigoPedido, codigoPedido: codigoPedido };
    } catch (error) {
        return manejarError('guardarCotizacion', error);
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

function obtenerPedidoParaEdicion(numPedido) {
    try {
        const COL_MAP = getColumnMap(HOJA_COTIZACIONES); 
        const allData = obtenerDatosHoja(HOJA_COTIZACIONES, true, 15); 
        const pedidoBuscado = String(numPedido).trim().toUpperCase();
        const CODIGO_PEDIDO_COL = COL_MAP['COT'] || 0; 
        
        // 1. Encontrar todas las filas coincidentes
        const filasCoincidentes = allData.slice(1)
            .filter(fila => String(fila[CODIGO_PEDIDO_COL] || '').trim().toUpperCase() === pedidoBuscado);
        
        if (filasCoincidentes.length === 0) throw new Error(`Pedido ${numPedido} no encontrado. Verifique el código.`);
        
        const primeraFila = filasCoincidentes[0];
        
        // Función auxiliar para obtener valores de la primera fila de forma segura
        const getGenValue = (colName) => getFilaValue(primeraFila, COL_MAP, colName);

        // 2. EXTRACCIÓN DE DATOS GENERALES RESILIENTE
        const datosGenerales = {
            fechaRegistro: getSafeDateString(getGenValue('FECHA COT')),
            
            Empresa: String(getGenValue('EMPRESA') || ''), 
            RUC: String(getGenValue('ID CLIENTE') || ''),
            Cliente: String(getGenValue('CLIENTE') || ''), 
            Estado: String(getGenValue('ESTADO COT') || 'COTIZACION'),
            Moneda: String(getGenValue('MONEDA') || 'SOLES'), 
            Forma_Pago: String(getGenValue('F. PAGO') || ''),
            Direccion: String(getGenValue('UBICACIÓN') || ''), 
            Turno: String(getGenValue('TURNO') || 'Diurno'), 
            Ejecutivo: String(getGenValue('EJECUTIVO') || ''), 
            Contacto: String(getGenValue('CONTACTO') || '') 
        };
        
        // 3. EXTRACCIÓN DE LÍNEAS DE SERVICIO EXTREMADAMENTE RESILIENTE
        const lineas = filasCoincidentes.map((fila) => {
            const getValue = (colName) => getFilaValue(fila, COL_MAP, colName); // Alias para la fila actual

            return {
                cod: String(getValue('COD') || ''), 
                descripcion: String(getValue('DESCRIPCION') || ''),
                cantidad: parseFloat(getValue('UND. PEDIDO') || 0) || 0, 
                und_medida: String(getValue('UND') || 'HORAS'),
                und_horas_minimas: String(getValue('UND. HORAS. MINIMAS') || ''), 
                dias_cotizados: parseFloat(getValue('TOTAL DIAS') || 0) || 0,
                horas_minimas_num: parseFloat(getValue('HORAS MINIMAS') || 0) || 0, 
                hora_segun: String(getValue('HORAS SEGÚN') || ''),
                movilizacion: parseFloat(getValue('MOV. Y DES. MOV.') || 0) || 0, 
                precio: parseFloat(getValue('PRECIO') || 0) || 0,
                subtotal: parseFloat(getValue('M. PEDIDO') || 0) || 0
            };
        });
        
        // 4. Construcción del resultado final
        const resultado = { 
            success: true, 
            ...datosGenerales, 
            Lineas: lineas, 
            Total: parseFloat(getGenValue('Total servicio') || 0) || 0, 
            totalLineas: lineas.length, 
            numPedido: numPedido, 
            usuario: obtenerEmailSeguro(), 
            autorizado: true 
        };

        return resultado;
        
    } catch (error) {
        // Registrar el error específico en los logs del servidor
        Logger.log(`❌ ERROR CRÍTICO al extraer pedido ${numPedido}: ${error.message}`);
        
        // Devolver un mensaje de error manejado por el frontend
        return manejarError('obtenerPedidoParaEdicion', error);
    }
}

// ====================================================
// === PASO 4: MIGRACIÓN DE CRUD A crudHoja (Contactos/Direcciones/Clientes) ===
// ====================================================

/**
 * Guarda o actualiza un Contacto usando crudHoja.
 */
function guardarOActualizarContacto(data) {
    try {
        const datosSanitizados = sanitizarDatos(data);
        const rowIndex = parseInt(datosSanitizados.rowIndex);
        
        // Generar un ID solo si es un nuevo registro
        const idContacto = datosSanitizados.ID_CONTACTO || (rowIndex > 1 ? datosSanitizados.ID_CONTACTO : `CON-${new Date().getTime()}`);
        
        // Valores alineados con las columnas de la hoja Contactos
        const nuevosValores = [
            idContacto,                             // Columna 1: ID
            datosSanitizados.RUC,                   // Columna 2: RUC
            datosSanitizados.NOMBRE,                // Columna 3: NOMBRE
            datosSanitizados.EMAIL,                 // Columna 4: EMAIL
            datosSanitizados.TELEFONO,              // Columna 5: TELEFONO
            datosSanitizados.CARGO                  // Columna 6: CARGO
        ];
        
        const operacion = rowIndex > 1 ? 'UPDATE' : 'CREATE';
        
        const resultado = crudHoja(operacion, HOJA_CONTACTOS, { rowIndex: rowIndex, valores: nuevosValores });

        if (resultado.success) {
            return { success: true, message: "Contacto guardado exitosamente" };
        } else {
            throw new Error(resultado.message);
        }
    } catch (error) {
        throw new Error("Error al guardar el contacto: " + error.message);
    }
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

/**
 * Guarda un nuevo Cliente usando crudHoja.
 */
function guardarNuevoCliente(data) {
    try {
        const datosSanitizados = sanitizarDatos(data);
        const nuevosValores = [
            datosSanitizados.RUC || '', 
            datosSanitizados.NOMBRE || ''
        ];
        
        if (!nuevosValores[0] || !nuevosValores[1]) {
            throw new Error("RUC y Nombre son campos obligatorios.");
        }
        
        crudHoja('CREATE', HOJA_CLIENTES, nuevosValores);
        
        return { 
            success: true, 
            message: "Cliente registrado exitosamente.",
            nuevoCliente: nuevosValores 
        };
    } catch (error) {
        throw new Error("Error al guardar el cliente: " + error.message);
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
 * REEMPLAZO v2 de getListaCotizacionesResumen
 * Incluye Vendedor y Compañía en los datos devueltos.
 */
function getListaCotizacionesResumen() {
    try {
        const allData = obtenerDatosHoja(HOJA_COTIZACIONES, true, 1); // Usar caché
        const encabezadoBase = ["N° Pedido", "Fecha", "Cliente", "Vendedor", "Compañía", "Monto Total", "Moneda", "Estado", "ID Cliente", "Row Index"];

        if (allData.length <= 1) {
            return [encabezadoBase]; // Devuelve solo encabezados si no hay datos
        }

        const COL_MAP = getColumnMap(HOJA_COTIZACIONES);
        // Asegúrate que los nombres de encabezado coincidan con tu HOJA_COTIZACIONES
        const INDICES = {
            NUM_PEDIDO: COL_MAP['COT'],
            FECHA_CREACION: COL_MAP['FECHA COT'],
            CLIENTE: COL_MAP['CLIENTE'],
            VENDEDOR: COL_MAP['EJECUTIVO'], // Mapea a la columna EJECUTIVO
            COMPANIA: COL_MAP['EMPRESA'],   // Mapea a la columna EMPRESA
            MONTO_TOTAL: COL_MAP['TOTAL SERVICIO'], // Usa TOTAL SERVICIO
            MONEDA: COL_MAP['MONEDA'],
            ESTADO: COL_MAP['ESTADO COT'],
            ID_CLIENTE: COL_MAP['ID CLIENTE'],
        };

        // Validar que se encontraron los índices necesarios
        for (const key in INDICES) {
            if (INDICES[key] === undefined) {
                 Logger.log(`Advertencia: No se encontró el encabezado para '${key}' en HOJA_COTIZACIONES. Usando índice por defecto o causará error.`);
                 // Podrías asignar índices por defecto aquí si es necesario, ej: INDICES.VENDEDOR = 4;
            }
        }


        const resumenMap = new Map();
        const scriptTimeZone = Session.getScriptTimeZone();

        for (let i = 1; i < allData.length; i++) {
            const row = allData[i];
            const numPedido = String(row[INDICES.NUM_PEDIDO] || '').trim();
            if (!numPedido) continue;

            let montoLinea = 0;
            try { montoLinea = procesarMontoRapido(row[INDICES.MONTO_TOTAL]); } catch(e){ montoLinea = 0; }

            if (!resumenMap.has(numPedido)) {
                const fecha = formatearFechaRapido(row[INDICES.FECHA_CREACION], scriptTimeZone);
                const rowIndex = i + 1;
                resumenMap.set(numPedido, {
                    numPedido: numPedido,
                    fecha: fecha,
                    cliente: row[INDICES.CLIENTE] || 'N/A',
                    vendedor: row[INDICES.VENDEDOR] || 'N/A', // Capturar vendedor
                    compania: row[INDICES.COMPANIA] || 'N/A', // Capturar compañía
                    moneda: String(row[INDICES.MONEDA] || 'SOL').toUpperCase().trim(),
                    estado: row[INDICES.ESTADO] || 'Pendiente',
                    idCliente: row[INDICES.ID_CLIENTE] || '',
                    montoTotal: montoLinea,
                    rowIndex: rowIndex
                });
            } else {
                 const existente = resumenMap.get(numPedido);
                 existente.montoTotal += montoLinea;
                 // Opcional: actualizar vendedor/compañía si estaban vacíos
                 if (existente.vendedor === 'N/A' && row[INDICES.VENDEDOR]) existente.vendedor = row[INDICES.VENDEDOR];
                 if (existente.compania === 'N/A' && row[INDICES.COMPANIA]) existente.compania = row[INDICES.COMPANIA];
            }
        }

        const resumenData = [encabezadoBase]; // Usar el encabezado definido al inicio
        resumenMap.forEach(item => {
            resumenData.push([
                item.numPedido, item.fecha, item.cliente,
                item.vendedor, item.compania, // Añadir los nuevos datos
                item.montoTotal.toFixed(2), // Mantener el número aquí, el formato se hace en frontend
                item.moneda,
                item.estado, item.idCliente, item.rowIndex
            ]);
        });
        return resumenData;
    } catch (e) {
        Logger.log(`❌ ERROR FATAL en getListaCotizacionesResumen: ${e.message} \n ${e.stack}`);
        // Devolver encabezado y mensaje de error
        return [ ["N° Pedido", "Fecha", "Cliente", "Vendedor", "Compañía", "Monto Total", "Moneda", "Estado", "ID Cliente", "Row Index"],
                 ["Error", "No se pudieron cargar los datos", e.message, "", "", "", "", "", "", ""] ];
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
 * Verifica si un código es único
 */
function verificarCodigoUnico(codigoGenerado) {
    try {
        const resultado = buscarRegistro(HOJA_COTIZACIONES, codigoGenerado, 0);
        return resultado === null;
    } catch (error) {
        Logger.log(`Error en verificarCodigoUnico: ${error.message}`);
        return true;
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
 * Obtiene la lista completa de Órdenes de Trabajo para la tabla de resumen (OT.html).
 */
function getListaOT() {
    try {
        // Usamos obtenerDatosHoja para lectura con cache
        return obtenerDatosHoja(HOJA_OT); 
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
        const PEDIDOS_SHEET = HOJA_PEDIDOS_OT; 
        const filtro = { 'RUC': rucCliente, 'ID Servicio': idServicio };
        const pedidos = obtenerDatosFiltrados(PEDIDOS_SHEET, filtro);
        return pedidos.slice(1);
    } catch (e) {
        return manejarError('filtrarPedidosPorClienteYServicio', e);
    }
}

/**
 * Guarda o actualiza una Orden de Trabajo.
 * Implementa mapeo para escritura robusta.
 */
function guardarOT(data) {
    try {
        const OT_SHEET = HOJA_OT;
        const COL_MAP = getColumnMap(OT_SHEET);
        const datos = sanitizarDatos(data);
        
        const modo = datos.modo;
        const numOT = datos.numeroOT;
        
        // --- LÓGICA DE VALIDACIÓN Y BUSQUEDA (Omitida por ser extensa) ---

        let rowIndex = 0;
        if (modo === 'editar') {
            const resultado = buscarRegistro(OT_SHEET, datos.otIDExistente, COL_MAP['N° OT'] || 0);
            if (!resultado) throw new Error(`OT a editar ${datos.otIDExistente} no encontrada.`);
            rowIndex = resultado.indiceFila;
        }

        // 2. Crear un array de valores (Asumo 30 columnas como ejemplo, ajústalo si es necesario)
        const filaCompleta = new Array(30).fill('');

        // 3. Mapear datos al array de forma robusta
        const mapAndSet = (colName, value) => {
            const index = COL_MAP[colName];
            if (index !== undefined) filaCompleta[index] = value;
        };
        
        // --- ESCRITURA DE DATOS PRINCIPALES ---
        mapAndSet('N° OT', datos.numeroOT);
        mapAndSet('Fecha', datos.fecha);
        mapAndSet('Cliente', datos.cliente);
        mapAndSet('Servicio/Máquina', datos.servicio);
        mapAndSet('Pedido', datos.pedido);
        mapAndSet('Hora Inicio', datos.horaInicio);
        mapAndSet('Hora Fin', datos.horaFin);
        mapAndSet('Tiempo Refrigerio (min)', parseFloat(datos.tiempoRefrigerio || 0));
        mapAndSet('Horómetro Inicio', parseFloat(datos.horometroInicio || 0));
        mapAndSet('Horómetro Fin', parseFloat(datos.horometroFin || 0));
        mapAndSet('Es Camión Grúa', datos.esCamionGrua ? 'SI' : 'NO');
        // --- Fin Escritura ---

        // 4. Ejecutar CRUD
        if (modo === 'editar') {
            crudHoja('UPDATE', OT_SHEET, { rowIndex: rowIndex, valores: filaCompleta });
            return { success: true, message: `OT ${numOT} actualizada.` };
        } else {
            crudHoja('CREATE', OT_SHEET, filaCompleta);
            return { success: true, message: `OT ${numOT} registrada.` };
        }

    } catch (e) {
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
 * Procesa el monto rápidamente (asume que existe y está correcta)
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

/**
 * Formatea la fecha rápidamente (asume que existe y está correcta)
 */
function formatearFechaRapido(rawFecha, timeZone) {
    if (!rawFecha) return '';
    if (rawFecha instanceof Date) {
        return Utilities.formatDate(rawFecha, timeZone, 'dd/MM/yyyy');
    }
    // Para valores que no son objetos Date (ej. strings o números)
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

/**
 * Normaliza el nombre del turno.
 */
function normalizarTurno(turno) {
    if (!turno || typeof turno !== 'string') return 'Diurno';
    const turnoUpper = turno.toUpperCase().trim();
    if (turnoUpper.includes('DIURNO')) return 'Diurno';
    if (turnoUpper.includes('NOCTURNO')) return 'Nocturno';
    if (turnoUpper.includes('DOBLE')) return 'Doble Turno';
    return 'Diurno';
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
 * REEMPLAZO FINAL v3 de 'generarYDevolverPDF'
 * Añade la lógica para mover archivos existentes a una subcarpeta "Versiones Anteriores".
 */
function generarYDevolverPDF(rowIndex) {
    let newSS; 
    try {
        // 1. Obtener N° de Pedido (Sin cambios)
        const pedidoRow = crudHoja('READ_ROW', HOJA_COTIZACIONES, { rowIndex: rowIndex });
        if (!pedidoRow) throw new Error("No se encontró el pedido con el índice: " + rowIndex);
        
        const COL_MAP = getColumnMap(HOJA_COTIZACIONES);
        const numPedido = pedidoRow[COL_MAP['COT'] || 0]; 
        if (!numPedido) throw new Error("No se pudo encontrar el N° de Pedido (COT).");
        
        // 2. Obtener TODOS los detalles (Sin cambios)
        const cotizacionData = obtenerDetallesCompletosDePedido(numPedido);
        cotizacionData.numPedido = numPedido;
        
        // 3. Obtener información de origen y destino (Sin cambios)
        const fechaObj = new Date(cotizacionData.fecha || new Date());
        const { id: sourceTemplateFileId, tab: sourceTemplateTabName } = getTemplateInfo(cotizacionData.empresa); 
        const destinationFolder = getDestinationFolder(cotizacionData.ejecutivo, cotizacionData.empresa, fechaObj, cotizacionData); // Esta es la carpeta "cot" (Nivel 5)

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
        // Mueve el NUEVO archivo Sheet a la carpeta "cot"
        destinationFolder.addFile(newFile); 
        DriveApp.getRootFolder().removeFile(newFile);

        // --- INICIO DE NUEVA LÓGICA DE ARCHIVADO ---
        // 7. Archivar Versiones Anteriores
        const nombreCarpetaArchivo = "Versiones Anteriores";
        const carpetaArchivo = findOrCreateFolder(destinationFolder, nombreCarpetaArchivo); // Crea "Versiones Anteriores" dentro de "cot"
        
        const archivosEnCot = destinationFolder.getFiles();
        while (archivosEnCot.hasNext()) {
            const archivo = archivosEnCot.next();
            // Mover SÓLO si NO es el archivo que acabamos de crear Y NO es un PDF reciente con el mismo nombre base
            const esNuevoSheet = (archivo.getId() === newFileId);
            const esPDFPotencial = archivo.getName().startsWith(nombreArchivo) && archivo.getMimeType() === MimeType.PDF;

            if (!esNuevoSheet && !esPDFPotencial) { // Mueve todos los archivos excepto el nuevo Sheet y PDFs con nombre similar
                 Logger.log(`Archivando archivo anterior: ${archivo.getName()}`);
                 archivo.moveTo(carpetaArchivo); // Mover a "Versiones Anteriores"
            } else if (esPDFPotencial) {
                 // Si es un PDF con nombre similar, podría ser el de la ejecución anterior, lo archivamos también
                 // Podríamos añadir una comprobación de fecha si fuera necesario, pero por ahora lo movemos.
                 Logger.log(`Archivando PDF anterior: ${archivo.getName()}`);
                 archivo.moveTo(carpetaArchivo);
            }
        }
        // --- FIN DE NUEVA LÓGICA DE ARCHIVADO ---

        // 8. Rellenar la plantilla (Sin cambios)
        rellenarPlantilla(hojaTemporal, cotizacionData);
        SpreadsheetApp.flush();

        // 9. Crear el PDF (Sigue igual, se crea en la carpeta "cot")
        const newSS_Url = newSS.getUrl();
        const exportUrl = newSS_Url.replace('/edit', '/export?exportFormat=pdf&gid=' + hojaTemporal.getSheetId() + '&format=pdf&size=A4&portrait=true&fitw=true&gridlines=false&sheetnames=false');

        const response = UrlFetchApp.fetch(exportUrl, {
            headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
            muteHttpExceptions: true
        });
        
        const pdfBlob = response.getBlob().setName(nombreArchivo + ".pdf"); 
        const pdfFile = destinationFolder.createFile(pdfBlob); // Se guarda en "cot"

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
