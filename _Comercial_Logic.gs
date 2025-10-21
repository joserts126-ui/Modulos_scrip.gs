// ==================================================================================================
// === ESTE ARCHIVO CONTIENE LA LÓGICA DE NEGOCIO DEL MÓDULO COMERCIAL (EX MODULOS_SCRIP.GS) ===
// === NOTA: ESTE ARCHIVO DEPENDE DE QUE EXISTAN _Core.gs y _Constants_Lists.gs ===
// ==================================================================================================

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
        const serviciosData = obtenerDatosHoja(HOJA_SERVICIOS);
        return {
            success: true,
            clientes: obtenerDatosHoja(HOJA_CLIENTES),
            servicios: serviciosData,
            formasPago: LISTA_FORMAS_PAGO, 
            listaEmpresas: LISTA_EMPRESAS,
            listaTurnos: LISTA_TURNOS,
            listaUndMedida: getListaValoresUnicosOptimizada(serviciosData, 5), 
            listaHorasMinimas: LISTA_HORAS_MINIMAS_UND,
            listaHorasSegun: getListaValoresUnicosOptimizada(serviciosData, 4),
            listaEjecutivos: LISTA_EJECUTIVOS
        };
    } catch (e) {
        return manejarError('getDatosInicialesComercial', e);
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
            // Verifica que la fecha sea válida (no NaN)
            if (!isNaN(date)) {
                return date.toISOString();
            }
        } catch (e) {
            // Ignorar error de parseo
        }
    }
    // Devuelve la fecha actual por defecto
    return new Date(defaultValue || new Date()).toISOString();
}

function obtenerPedidoParaEdicion(numPedido) {
    try {
        const COL_MAP = getColumnMap(HOJA_COTIZACIONES); 
        // Usamos cache de 15 minutos para la data principal
        const allData = obtenerDatosHoja(HOJA_COTIZACIONES, true, 15); 
        const pedidoBuscado = String(numPedido).trim().toUpperCase();
        const CODIGO_PEDIDO_COL = COL_MAP['COT'] || 0; 
        
        // 1. Encontrar todas las filas coincidentes
        const filasCoincidentes = allData.slice(1)
            .filter(fila => String(fila[CODIGO_PEDIDO_COL] || '').trim().toUpperCase() === pedidoBuscado);
        
        if (filasCoincidentes.length === 0) throw new Error(`Pedido ${numPedido} no encontrado. Verifique el código.`);
        
        const primeraFila = filasCoincidentes[0];
        
        // Función auxiliar para obtener valores de la primera fila
        const getGenValue = (colName) => getFilaValue(primeraFila, COL_MAP, colName);

        // 2. EXTRACCIÓN DE DATOS GENERALES RESILIENTE
        const datosGenerales = {
            // FIX: Uso de getSafeDateString para evitar crashes por fechas nulas
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
                // Doble || 0 para garantizar que sea un número, incluso si es null o ""
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
        Logger.log(`❌ ERROR FATAL al extraer pedido ${numPedido}: ${error.message}`);
        Logger.log(error.stack);
        
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

// ====================================================
// === FUNCIONES DE RESUMEN COMERCIAL OPTIMIZADAS ===
// ====================================================

function getListaCotizacionesResumen() {
    try {
        const allData = obtenerDatosHoja(HOJA_COTIZACIONES, true, 1); // Reducir cache para depuración
        if (allData.length <= 1) {
            return [["N° Pedido", "Fecha", "Cliente", "Monto Total", "Moneda", "Estado", "ID Cliente", "Row Index"]];
        }
        
        const COL_MAP = getColumnMap(HOJA_COTIZACIONES);
        
        // 🔑 NOTA: Siempre usar el || para el caso en que un encabezado se borre o cambie de nombre.
        const INDICES = {
            NUM_PEDIDO: COL_MAP['COT'] || 0, 
            FECHA_CREACION: COL_MAP['FECHA COT'] || 2, 
            MONEDA: COL_MAP['MONEDA'] || 15, 
            CLIENTE: COL_MAP['CLIENTE'] || 6, 
            ESTADO: COL_MAP['ESTADO COT'] || 7, 
            ID_CLIENTE: COL_MAP['ID CLIENTE'] || 5, 
            MONTO_TOTAL: COL_MAP['Total servicio'] || 22
        };
        
        const resumenMap = new Map();
        const scriptTimeZone = Session.getScriptTimeZone();

        for (let i = 1; i < allData.length; i++) {
            const row = allData[i];
            
            // --- PROTECCIÓN CRÍTICA ---
            const numPedido = String(row[INDICES.NUM_PEDIDO] || '').trim();
            // Si el número de pedido es nulo o vacío, saltar la fila para evitar fallos.
            if (!numPedido) continue; 
            
            // Usamos un try/catch interno para procesar el monto y evitar que una sola fila dañada tumbe todo.
            let montoLinea = 0;
            try {
                montoLinea = procesarMontoRapido(row[INDICES.MONTO_TOTAL]);
            } catch(e) {
                Logger.log(`⚠️ Fila ${i+1}: Error al procesar monto (${e.message}). Se usará 0.`);
                montoLinea = 0;
            }
            
            // Procesamiento principal
            if (!resumenMap.has(numPedido)) {
                
                const fecha = formatearFechaRapido(row[INDICES.FECHA_CREACION], scriptTimeZone);
                const rowIndex = i + 1; // Índice real de la hoja
                
                resumenMap.set(numPedido, {
                    numPedido: numPedido, 
                    fecha: fecha,
                    moneda: String(row[INDICES.MONEDA] || 'SOL').toUpperCase().trim(),
                    cliente: row[INDICES.CLIENTE] || 'Cliente Desconocido',
                    estado: row[INDICES.ESTADO] || 'Pendiente',
                    idCliente: row[INDICES.ID_CLIENTE] || '',
                    montoTotal: montoLinea,
                    rowIndex: rowIndex 
                });
            } else {
                const existente = resumenMap.get(numPedido);
                existente.montoTotal += montoLinea;
            }
        }
        
        const resumenData = [["N° Pedido", "Fecha", "Cliente", "Monto Total", "Moneda", "Estado", "ID Cliente", "Row Index"]];
        resumenMap.forEach(item => {
            resumenData.push([
                item.numPedido, item.fecha, item.cliente, item.montoTotal.toFixed(2), item.moneda,
                item.estado, item.idCliente, item.rowIndex 
            ]);
        });
        
        return resumenData;
    } catch (e) {
        // Si el error llega aquí, es grave (ej. no se puede abrir la hoja).
        Logger.log(`❌ ERROR FATAL en getListaCotizacionesResumen: ${e.message}`);
        // Devolvemos una estructura que el frontend puede manejar sin colapsar.
        return [["N° Pedido", "Fecha", "Cliente", "Monto Total", "Moneda", "Estado", "ID Cliente", "Row Index"]];
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
        
        // Función auxiliar para obtener valores de forma segura.
        const getValue = (colName) => getFilaValue(filaOT, COL_MAP, colName);

        // Mapeo seguro de datos (Ajusta los nombres de las columnas a tu hoja 'OT')
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
        // ✅ FIX: Usar ALPAMAYO como plantilla GENÉRICA si la ubicación no se reconoce.
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

// ====================================================
// === FUNCIONES DE PDF (ROBUSTO CON MAPEO) ===
// ====================================================

/**
 * @param {number} rowIndex El número de fila del pedido en la HOJA_COTIZACIONES.
 * @returns {Object} Un objeto con { success: true, pdfUrl: URL } o { success: false, error: mensaje }.
 */
function generarYDevolverPDF(rowIndex) {
    try {
        // 1. Obtener la fila completa de la cotización usando READ_ROW
        const pedidoRow = crudHoja('READ_ROW', HOJA_COTIZACIONES, { rowIndex: rowIndex });

        if (!pedidoRow) {
            return { success: false, error: "No se encontró el pedido con el índice: " + rowIndex };
        }
        
        // 2. Obtener el MAPEO DE COLUMNAS
        const COL_MAP = getColumnMap(HOJA_COTIZACIONES);
        
        // 3. Mapear los datos al objeto que necesita generarPDFCotizacion
        const dataParaPDF = {
            ruc: pedidoRow[COL_MAP['ID CLIENTE']] || '',
            cliente: pedidoRow[COL_MAP['CLIENTE']] || '', 
            fecha: pedidoRow[COL_MAP['FECHA COT']] || new Date(),
            contacto: pedidoRow[COL_MAP['CONTACTO']] || '',
            correoCelular: '', // Dejamos vacío o lo mapeamos si tienes la columna
            lugar: pedidoRow[COL_MAP['UBICACIÓN']] || '',
            turno: pedidoRow[COL_MAP['TURNO']] || '',
            duracion: '', // No mapeado
            // Servicios y Total requieren una lógica adicional:
            // Dado que un pedido tiene muchas líneas, debemos buscar todas las líneas.
            // Para simplicidad por ahora (y evitar otra función pesada), usamos los datos de la primera línea:
            servicios: [
                { descripcion: "Pedido: " + pedidoRow[COL_MAP['COT']], valor: pedidoRow[COL_MAP['Total servicio']] || 0 }
            ], 
            total: pedidoRow[COL_MAP['Total servicio']] || 0 
        };
        
        // 4. Generar el PDF (retorna un Blob)
        const pdfBlob = generarPDFCotizacion(dataParaPDF);
        const archivoEnDrive = DriveApp.createFile(pdfBlob);
        
        // 5. Devolver la URL
        return { 
            success: true, 
            pdfUrl: archivoEnDrive.getUrl()
        };
    } catch (e) {
        Logger.log("Error en generarYDevolverPDF: " + e.toString());
        return { success: false, error: e.message };
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
    // Retorna el valor si el índice existe y está dentro de los límites de la fila
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
 * Genera una cotización creando una copia de la plantilla de Google Sheets,
 * llenando los datos de cabecera y añadiendo dinámicamente las líneas de servicio.
 * Esta versión incorpora búsqueda de datos de RUC y Contacto de pestañas auxiliares.
 */
function generarCotizacionSheet() {
  
  // ====================================================================
  // 1. CONFIGURACIÓN
  // ====================================================================

  // IDs y Nombres de Archivos/Pestañas
  const ID_PLANTILLA_SHEET = '19N-nRzPwQbcuRHogyUOkkIcpGmrLGyMT';
  const ID_CARPETA_DESTINO = '1Z2M-jSQBI9RNQzMToh5OaXkDvbfftcVo';
  const NOMBRE_PESTAÑA_DATA = 'DataCot';
  const NOMBRE_PESTAÑA_CLIENTES = 'Clientes'; // Nueva pestaña
  const NOMBRE_PESTAÑA_CONTACTOS = 'Contactos'; // Nueva pestaña
  const NOMBRE_PESTAÑA_PLANTILLA = 'COT_ALP';
  
  // Ubicaciones de Datos
  const FILA_INICIO_SERVICIOS_PLANTILLA = 18;
  const INDICE_FILA_CABECERA = 1; // Fila 2 de DataCot
  const INDICE_FILA_SERVICIOS = 2; // Fila 3 de DataCot

  // ====================================================================
  // 2. FUNCIONES DE UTILIDAD PARA MAPEO Y BÚSQUEDA
  // ====================================================================
  
  /**
   * Crea un objeto (mapa) de datos a partir de una fila, usando los encabezados
   * como claves. Normaliza las claves a mayúsculas y quita espacios.
   */
  function createDataMap(headers, dataRow) {
    const dataMap = {};
    headers.forEach((header, index) => {
      if (header) {
        const key = header.trim().toUpperCase().replace(/[^A-Z0-9]+/g, '_');
        dataMap[key] = dataRow[index];
      }
    });
    return dataMap;
  }
  
  /**
   * Busca un valor en una pestaña por el nombre del encabezado (similar a un VLOOKUP).
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - La hoja donde buscar.
   * @param {string} searchHeader - El encabezado de la columna de búsqueda.
   * @param {any} searchValue - El valor a buscar.
   * @param {string} returnHeader - El encabezado de la columna cuyo valor se desea.
   * @returns {any|null} El valor encontrado o null.
   */
  function searchDataByHeader(sheet, searchHeader, searchValue, returnHeader) {
    if (!sheet || !searchValue) return null;
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const normalizedHeaders = headers.map(h => h.trim().toUpperCase().replace(/[^A-Z0-9]+/g, '_'));
    
    // Encuentra los índices de las columnas de búsqueda y retorno
    const searchIndex = normalizedHeaders.indexOf(searchHeader.trim().toUpperCase().replace(/[^A-Z0-9]+/g, '_'));
    const returnIndex = normalizedHeaders.indexOf(returnHeader.trim().toUpperCase().replace(/[^A-Z0-9]+/g, '_'));
    
    if (searchIndex === -1 || returnIndex === -1) {
      Logger.log(`Error: No se encontró el encabezado de búsqueda ('${searchHeader}') o de retorno ('${returnHeader}') en la hoja ${sheet.getName()}.`);
      return null;
    }
    
    // Recorre las filas (empezando por la segunda, después de los encabezados)
    for (let i = 1; i < data.length; i++) {
      if (data[i][searchIndex] && data[i][searchIndex].toString().trim() === searchValue.toString().trim()) {
        return data[i][returnIndex];
      }
    }
    return null; // No encontrado
  }


  // ====================================================================
  // 3. OBTENER DATA DEL GOOGLE SHEET DE ORIGEN
  // ====================================================================

  const ssData = SpreadsheetApp.getActiveSpreadsheet();
  const hojaData = ssData.getSheetByName(NOMBRE_PESTAÑA_DATA);
  const hojaClientes = ssData.getSheetByName(NOMBRE_PESTAÑA_CLIENTES);
  const hojaContactos = ssData.getSheetByName(NOMBRE_PESTAÑA_CONTACTOS);

  if (!hojaData) {
    SpreadsheetApp.getUi().alert(`❌ Error: No se encontró la pestaña de datos "${NOMBRE_PESTAÑA_DATA}".`);
    return;
  }

  const valores = hojaData.getDataRange().getValues();
  const encabezados = valores[0];

  // 3.1. Mapeo de Cabecera (Fila 2)
  const dataRowCabecera = valores[INDICE_FILA_CABECERA];
  const mapCabecera = createDataMap(encabezados, dataRowCabecera);
  
  // 3.2. Extracción y Mapeo de Servicio (Desde Fila 3)
  const dataServicios = valores.slice(INDICE_FILA_SERVICIOS)
    .filter(row => {
      const servicioMap = createDataMap(encabezados, row);
      return servicioMap['DESCRIPCION']; 
    }); 
    
  // 3.3. Obtener RUC y Contacto desde pestañas auxiliares
  const nombreCliente = mapCabecera['CLIENTE'] || 'Cliente Desconocido';
  
  // Buscar RUC usando el NOMBRE del CLIENTE
  const rucCliente = searchDataByHeader(hojaClientes, 'NOMBRE', nombreCliente, 'RUC');
  
  // Buscar datos de contacto usando el RUC encontrado
  const nombreContacto = searchDataByHeader(hojaContactos, 'RUC', rucCliente, 'NOMBRE');
  const emailContacto = searchDataByHeader(hojaContactos, 'RUC', rucCliente, 'EMAIL');
  const cargoContacto = searchDataByHeader(hojaContactos, 'RUC', rucCliente, 'CARGO');


  // ====================================================================
  // 4. CREAR COPIA DE LA PLANTILLA Y CONFIGURACIÓN
  // ====================================================================

  const numeroCotizacion = mapCabecera['NUM_COT'] || '000';
  const nombreArchivoFinal = `Cotización N° ${numeroCotizacion} - ${nombreCliente}`;

  const plantillaFile = DriveApp.getFileById(ID_PLANTILLA_SHEET);
  const archivoCopia = plantillaFile.makeCopy(nombreArchivoFinal, DriveApp.getFolderById(ID_CARPETA_DESTINO));
  const ssCotizacion = SpreadsheetApp.openById(archivoCopia.getId());
  const hojaCotizacion = ssCotizacion.getSheetByName(NOMBRE_PESTAÑA_PLANTILLA);
  
  if (!hojaCotizacion) {
    SpreadsheetApp.getUi().alert(`❌ Error: No se encontró la pestaña de plantilla "${NOMBRE_PESTAÑA_PLANTILLA}" en la copia.`);
    return;
  }
  
  // ====================================================================
  // 5. LLENAR DATOS DE CABECERA
  // ====================================================================

  hojaCotizacion.getRange('C3').setValue(numeroCotizacion);
  hojaCotizacion.getRange('C5').setValue(nombreCliente);
  hojaCotizacion.getRange('C6').setValue(rucCliente || mapCabecera['ID_CLIENTE'] || ''); // RUC encontrado o ID de DataCot
  hojaCotizacion.getRange('C7').setValue(nombreContacto || mapCabecera['CONTACTO'] || ''); // Nombre de contacto encontrado o Contacto de DataCot
  hojaCotizacion.getRange('C8').setValue(mapCabecera['TOTAL_DIAS'] || ''); 
  
  // 🚨 Opcional: Si quieres agregar más datos en la cabecera, úsalos aquí (Ej: Email y Cargo)
  // hojaCotizacion.getRange('C9').setValue(emailContacto || '');
  // hojaCotizacion.getRange('D9').setValue(cargoContacto || '');

  // ====================================================================
  // 6. LLENAR DATOS DE SERVICIOS (Llenado Dinámico)
  // ====================================================================
  
  let filaActual = FILA_INICIO_SERVICIOS_PLANTILLA;
  let itemNum = 1;
  const subtotalesCeldas = []; 

  // Insertar filas necesarias
  if (dataServicios.length > 1) {
    hojaCotizacion.insertRowsAfter(filaActual, dataServicios.length - 1);
  }
  
  dataServicios.forEach((dataRowServicio, index) => {
    const servicioMap = createDataMap(encabezados, dataRowServicio);

    const descripcion = servicioMap['DESCRIPCION'] || '';
    // Usamos TOTAL_SERVICIO de la fila de servicio como el valor de la línea
    const totalLinea = servicioMap['TOTAL_SERVICIO'] || 0; 
    
    // Línea de Servicio Principal
    hojaCotizacion.getRange('A' + filaActual).setValue(itemNum);
    hojaCotizacion.getRange('B' + filaActual).setValue(descripcion);
    hojaCotizacion.getRange('F' + filaActual).setValue(totalLinea);
    subtotalesCeldas.push('F' + filaActual);
    
    // Verificar Movilización
    const esUltimaLinea = (index === dataServicios.length - 1);
    // Usamos MOV_Y_DES_MOV_ de la fila de servicio actual.
    const montoMovilizacion = servicioMap['MOV_Y_DES_MOV_'] || 0; 

    if (esUltimaLinea && montoMovilizacion > 0) {
      
      filaActual++; 
      hojaCotizacion.insertRowAfter(filaActual - 1); 
      
      const itemMovilizacion = itemNum + '.1';
      
      hojaCotizacion.getRange('A' + filaActual).setValue(itemMovilizacion);
      hojaCotizacion.getRange('B' + filaActual).setValue('Movilización y Desmovilización');
      hojaCotizacion.getRange('F' + filaActual).setValue(montoMovilizacion);
      subtotalesCeldas.push('F' + filaActual);
    }

    filaActual++; 
    itemNum++;
  });
  
  // ====================================================================
  // 7. CÁLCULO DEL SUBTOTAL FINAL
  // ====================================================================
  
  // El subtotal final va en la fila inmediatamente posterior a la última línea (F19, F20, etc.)
  const filaSubtotalFinal = filaActual;
  const rangoSubtotales = subtotalesCeldas.join('+');
  
  // Colocar la fórmula en la celda F de la fila correspondiente
  hojaCotizacion.getRange('F' + filaSubtotalFinal).setFormula('=' + rangoSubtotales);
  
  // ====================================================================
  // 8. FINALIZAR
  // ====================================================================

  SpreadsheetApp.getUi().alert(`✅ ¡Cotización generada con éxito! El archivo "${nombreArchivoFinal}" se encuentra en la carpeta de destino.`);
}
