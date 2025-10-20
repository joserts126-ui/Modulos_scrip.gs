// ====================================================
// === CONFIGURACIÓN GLOBAL Y CONSTANTES DE HOJA ===
// ====================================================

const HOJA_ID_PRINCIPAL = "15qfA3idaLkyhvFwAeEZQo6L9BudBBfgnV8DFrs1qV6Y"; 
const HOJA_CLIENTES = "Clientes";
const HOJA_CONTACTOS = "Contactos";
const HOJA_DIRECCIONES = "Direcciones";
const HOJA_SERVICIOS = "Servicios";
const HOJA_COTIZACIONES = "DataCot";

// Nuevas constantes para las pestañas de plantillas
const HOJA_PLANTILLA_ALPAMAYO = 'COT_ALP';
const HOJA_PLANTILLA_GYM = 'COT_GYM';
const HOJA_PLANTILLA_SANJOSE = 'COT_GSJ';

// Eliminadas las constantes de plantilla PDF y carpeta

const CLIENTE_COLS = { RUC: 0, NOMBRE: 1 };
const CONTACTO_COLS = { RUC: 1, NOMBRE: 2, EMAIL: 3, TELEFONO: 4, CARGO: 5 };
const DIRECCION_COLS = { RUC: 1, TIPO: 2, DIRECCION: 3, CIUDAD: 4 };

// ====================================================
// === LISTAS ESTÁTICAS COMO CONSTANTES GLOBALES ===
// ====================================================

const LISTA_TURNOS = ["Diurno", "Nocturno", "Doble Turno"];
const LISTA_EMPRESAS = ["ALPAMAYO", "SAN JOSE", "GYM"];
const LISTA_ESTADOS_COT = ["Cotización", "Pedido", "Cancelado", "Finalizado"];
const LISTA_FORMAS_PAGO = [
    "CONTADO", "50% de adelanto", "CREDITO. 30 DIAS", "CREDITO. 15 DIAS", 
    "CREDITO. 45 DIAS", "CREDITO. 60 DIAS", "CREDITO. 90 DIAS"
];
const LISTA_HORAS_MINIMAS_UND = ["Mensual", "Diarias"]; 

// Cache para optimización
const cache = CacheService.getScriptCache();

// ====================================================
// === CORE: MANEJO DE VISTAS (doGet) Y UTILIDADES ===
// ====================================================

/**
 * Maneja las solicitudes GET y sirve las páginas correspondientes
 */
function doGet(e) {
    const page = e.parameter.page || 'Modulos'; 
    const validPages = ['Modulos', 'Comercial', 'Servicios', 'Contactos', 'ResumenComercial', 'OT', 'RegistrarOT']; 
    
    if (validPages.includes(page)) {
        return HtmlService.createTemplateFromFile(page).evaluate();
    }
    return HtmlService.createTemplateFromFile('Modulos').evaluate();
}

/**
 * Obtiene la URL del script desplegado
 */
function getScriptUrl() {
    return ScriptApp.getService().getUrl();
}

/**
 * Incluye archivos HTML parciales
 */
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ====================================================
// === FUNCIONES UNIFICADAS Y SIMPLIFICADAS ===
// ====================================================

/**
 * Función unificada para operaciones CRUD en hojas
 */
function crudHoja(operacion, sheetName, datos = null, filtro = null) {
    try {
        const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
        const sheet = ss.getSheetByName(sheetName);
        
        if (!sheet) throw new Error(`Hoja '${sheetName}' no encontrada`);
        
        // Invalidar cache
        cache.remove(`hoja_${sheetName}`);
        
        switch(operacion) {
            case 'READ':
                if (sheet.getLastRow() < 1) return [];
                return sheet.getDataRange().getValues();

            case 'READ_ROW': // <--- ¡NUEVA OPERACIÓN AÑADIDA!
                const rowIndexToRead = parseInt(datos.rowIndex);
                if (rowIndexToRead > 1 && rowIndexToRead <= sheet.getLastRow()) {
                    const lastCol = sheet.getLastColumn();
                    // Lee toda la fila (desde la columna 1 hasta la última)
                    return sheet.getRange(rowIndexToRead, 1, 1, lastCol).getValues()[0];
                }
                throw new Error("Índice de fila inválido o fuera de rango para lectura");

            case 'CREATE':
                const valoresCrear = Array.isArray(datos) ? datos : datos.valores;
                sheet.appendRow(valoresCrear);
                return { success: true, message: "Registro creado exitosamente", tipo: 'creacion' };
                
            case 'UPDATE':
                const rowIndex = parseInt(datos.rowIndex);
                if (rowIndex > 1) {
                    sheet.getRange(rowIndex, 1, 1, datos.valores.length)
                         .setValues([datos.valores]);
                    return { success: true, message: "Registro actualizado exitosamente", tipo: 'actualizacion' };
                }
                throw new Error("Índice de fila inválido para actualización");
                
            case 'FILTER':
                const allData = sheet.getDataRange().getValues();
                if (!filtro || Object.keys(filtro).length === 0) return allData;
                
                const headers = allData[0];
                const filtered = allData.slice(1).filter(row => 
                    Object.entries(filtro).every(([key, value]) => {
                        const colIndex = headers.indexOf(key);
                        return colIndex !== -1 && String(row[colIndex]).trim() === String(value).trim();
                    })
                );
                return [headers].concat(filtered);
                
            default:
                throw new Error(`Operación '${operacion}' no soportada`);
        }
    } catch (error) {
        Logger.log(`ERROR en crudHoja(${operacion}, ${sheetName}): ${error.message}`);
        throw error;
    }
}

/**
 * Función unificada para obtener datos con cache OPTIMIZADA
 */
function obtenerDatosHoja(sheetName, useCache = true, cacheMinutes = 10) {
    const cacheKey = `hoja_${sheetName}`;
    const startTime = new Date().getTime();
    
    try {
        if (useCache) {
            const cached = cache.get(cacheKey);
            if (cached) {
                const endTime = new Date().getTime();
                Logger.log(`⚡ CACHE HIT para ${sheetName}: ${endTime - startTime}ms`);
                return JSON.parse(cached);
            }
        }
        
        const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
        const sheet = ss.getSheetByName(sheetName);
        
        if (!sheet || sheet.getLastRow() === 0) {
            return [];
        }
        
        // Limitar cantidad de filas si es muy grande
        const lastRow = Math.min(sheet.getLastRow(), 1000); // Máximo 1000 filas
        const data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
        
        if (useCache && data.length > 0) {
            cache.put(cacheKey, JSON.stringify(data), cacheMinutes * 60);
        }
        
        const endTime = new Date().getTime();
        Logger.log(`📥 DATOS OBTENIDOS ${sheetName}: ${endTime - startTime}ms - ${data.length} filas`);
        
        return data;
        
    } catch (error) {
        Logger.log(`❌ ERROR en obtenerDatosHoja: ${error.message}`);
        return [];
    }
}

/**
 * Función unificada para guardar/actualizar registros (REEMPLAZA guardarEnHoja)
 */
function guardarEnHoja(sheetName, datos, idField = 'rowIndex') {
    const datosSanitizados = sanitizarDatos(datos);
    const operacion = parseInt(datosSanitizados[idField]) > 1 ? 'UPDATE' : 'CREATE';
    return crudHoja(operacion, sheetName, datosSanitizados);
}

/**
 * Obtiene valores únicos de una columna optimizada con cache
 */
function getListaValoresUnicosOptimizada(allData, columnIndex) {
    if (!allData || allData.length <= 1) return []; 
    
    const listaUnica = new Set();
    for (let i = 1; i < allData.length; i++) {
        const valor = allData[i][columnIndex];
        if (valor !== undefined && valor !== null) {
            const valorLimpio = String(valor).trim();
            if (valorLimpio) listaUnica.add(valorLimpio);
        }
    }
    return Array.from(listaUnica);
}

/**
 * Obtiene lista de valores únicos de una hoja específica
 */
function getListaValoresUnicos(sheetName, columnIndex) {
    try {
        const allData = obtenerDatosHoja(sheetName);
        return getListaValoresUnicosOptimizada(allData, columnIndex - 1); 
    } catch (e) {
        Logger.log(`ERROR en getListaValoresUnicos para ${sheetName}: ${e.message}`);
        return []; 
    }
}

/**
 * Función optimizada para obtener datos filtrados (SIMPLIFICAR)
 */
function obtenerDatosFiltrados(sheetName, filtro = {}) {
    return crudHoja('FILTER', sheetName, null, filtro);
}

/**
 * Función optimizada para búsqueda (MANTENER la existente pero simplificar)
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
// === FUNCIONES DE LISTAS DINÁMICAS OPTIMIZADAS ===
// ====================================================

function getListaHoraSegun() {
    return getListaValoresUnicos(HOJA_SERVICIOS, 5);
}

function getListaUnidadesDeMedida() {
    return getListaValoresUnicos(HOJA_SERVICIOS, 6);
}

/**
 * Busca servicio por código con cache
 */
function buscarServicioPorCodigo(codServicio) {
    return buscarRegistro(HOJA_SERVICIOS, codServicio, 0);
}

// ====================================================
// === FUNCIONES UNIFICADAS PARA MÓDULOS ===
// ====================================================

/**
 * Función unificada para obtener datos iniciales
 */
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
            listaEjecutivos: ['ANTHONY', 'CARMEN', 'RENATO']
        };
        
    } catch (e) {
        Logger.log(`Error en getDatosInicialesComercial: ${e.message}`);
        return manejarError('getDatosInicialesComercial', e);
    }
}

// ====================================================
// === FUNCIONES DE SERVICIOS OPTIMIZADAS (CRUD) ===
// ====================================================

/**
 * Obtiene lista de servicios con cache
 */
function getListaServicios() {
    return obtenerDatosHoja(HOJA_SERVICIOS);
}

/**
 * Función unificada para servicios
 */
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
        Logger.log(`Error en getDatosInicialesServicios: ${error.message}`);
        return {
            servicios: [['Error', 'No se pudieron cargar los servicios']],
            horasSegun: [],
            undMedida: []
        };
    }
}

/**
 * Guarda o actualiza un servicio - VERSIÓN CON VALIDACIÓN MEJORADA
 */
function guardarOActualizarServicio(dataObject) {
    try {
        Logger.log("💾 GUARDANDO SERVICIO - Datos recibidos:");
        Logger.log(JSON.stringify(dataObject, null, 2));
        
        // ✅ VALIDACIÓN MEJORADA
        const codigoServicio = (dataObject['ID Servicio'] || '').toString().trim();
        const descripcion = (dataObject['Descripción del Servicio'] || '').toString().trim();
        
        if (!codigoServicio) {
            throw new Error("El código del servicio es requerido");
        }
        
        if (!descripcion) {
            throw new Error("La descripción del servicio es requerida");
        }
        
        // Validar que el código no exista (solo para nuevos servicios)
        if (parseInt(dataObject.rowIndex) <= 1) { // Es un nuevo servicio
            const servicioExistente = buscarServicioPorCodigo(codigoServicio);
            if (servicioExistente) {
                throw new Error(`El código ${codigoServicio} ya existe. Use un código único.`);
            }
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
        
        Logger.log("📝 Valores a guardar:");
        Logger.log(JSON.stringify(valores, null, 2));
        
        const resultado = guardarEnHoja(HOJA_SERVICIOS, {
            ...dataObject,
            valores: valores
        });
        
        Logger.log("✅ Servicio guardado exitosamente");
        return resultado;
        
    } catch (error) {
        Logger.log(`❌ ERROR en guardarOActualizarServicio: ${error.message}`);
        throw error;
    }
}

// ====================================================
// === FUNCIONES DE REGISTRO DE COTIZACIÓN ===
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
// === FUNCIÓN PRINCIPAL DE GUARDADO ===
// ====================================================

function guardarCotizacion(datos) {
    try {
        Logger.log("💾 INICIANDO GUARDADO EN GS...");
        
        // Sanitizar datos primero
        const datosSanitizados = sanitizarDatos(datos);
        Logger.log("📍 DATOS RECIBIDOS Y SANITIZADOS EN GS:");
        Logger.log("📍 Dirección a guardar: " + datosSanitizados.Direccion);
        Logger.log("📞 Contacto a guardar: " + datosSanitizados.Contacto);
        Logger.log("👨‍💼 Ejecutivo a guardar: " + datosSanitizados.Ejecutivo);

        // Validar datos
        const erroresValidacion = validarDatosCotizacion(datosSanitizados);
        if (erroresValidacion.length > 0) {
            return {
                success: false,
                message: "Errores de validación: " + erroresValidacion.join(', ')
            };
        }

        const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
        let hojaCot = ss.getSheetByName(HOJA_COTIZACIONES);

        if (!hojaCot) {
            Logger.log("📄 Creando nueva hoja de cotizaciones...");
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
            Logger.log("✅ Hoja creada con encabezados");
        }

        let codigoPedido;
        
        // Lógica de código de pedido
        if (datosSanitizados.numPedido) {
            // MODO EDICIÓN: Eliminar filas existentes antes de guardar
            Logger.log("🔄 MODO EDICIÓN - Eliminando filas existentes para: " + datosSanitizados.numPedido);
            
            const allData = hojaCot.getDataRange().getValues();
            const filasAEliminar = [];
            
            // Buscar todas las filas del pedido a editar
            for (let i = 1; i < allData.length; i++) {
                const fila = allData[i];
                const codigoPedidoExistente = String(fila[0] || '').trim();
                
                if (codigoPedidoExistente === datosSanitizados.numPedido) {
                    filasAEliminar.push(i + 1);
                }
            }
            
            // Eliminar filas existentes (de abajo hacia arriba)
            if (filasAEliminar.length > 0) {
                filasAEliminar.sort((a, b) => b - a).forEach(rowIndex => {
                    hojaCot.deleteRow(rowIndex);
                });
                Logger.log("✅ Filas eliminadas: " + filasAEliminar.length);
            }
            
            codigoPedido = datosSanitizados.numPedido;
            Logger.log("🔄 Usando código existente: " + codigoPedido);
        } else {
            // Generar nuevo código para modo nuevo
            let intentos = 0;
            const maxIntentos = 5;
            
            do {
                codigoPedido = generarCodigoPedido(datosSanitizados.Empresa);
                intentos++;
                if (intentos > maxIntentos) {
                    throw new Error("No se pudo generar un código único después de " + maxIntentos + " intentos");
                }
            } while (!verificarCodigoUnico(codigoPedido));
            
            Logger.log("🆕 Nuevo código generado: " + codigoPedido);
        }

        const fechaRegistro = new Date(datosSanitizados.fechaRegistro || new Date());
        Logger.log("📅 Fecha de registro: " + fechaRegistro);

        const lineas = datosSanitizados.Lineas || [];
        if (lineas.length === 0) {
            throw new Error("Debe agregar al menos un servicio a la cotización");
        }

        const filaInicial = hojaCot.getLastRow() + 1;
        let filaActual = filaInicial;

        Logger.log("📝 Escribiendo " + lineas.length + " líneas desde fila: " + filaInicial);

        for (let i = 0; i < lineas.length; i++) {
            const linea = lineas[i];
            const filaCompleta = new Array(48).fill('');
            
            Logger.log(`   📋 Procesando línea ${i + 1}: ${linea.cod} - ${linea.descripcion}`);
            
            // ✅✅✅ CORRECCIÓN: ASIGNACIÓN CORRECTA DE COLUMNAS
            // DATOS GENERALES
            filaCompleta[0] = codigoPedido;                       // A - COT
            filaCompleta[2] = fechaRegistro;                      // C - FECHA COT
            filaCompleta[3] = datosSanitizados.Empresa || '';     // D - EMPRESA
            filaCompleta[4] = datosSanitizados.Ejecutivo || '';   // E - EJECUTIVO
            filaCompleta[5] = datosSanitizados.RUC || '';         // F - ID CLIENTE
            filaCompleta[6] = datosSanitizados.Cliente || '';     // G - CLIENTE
            filaCompleta[7] = datosSanitizados.Estado || 'COTIZACION'; // H - ESTADO COT
            filaCompleta[15] = datosSanitizados.Moneda || 'SOLES'; // P - MONEDA
            filaCompleta[25] = datosSanitizados.Turno || 'Diurno'; // Z - TURNO
            filaCompleta[38] = datosSanitizados.Forma_Pago || '';  // AM - F. PAGO
            filaCompleta[44] = datosSanitizados.Direccion || '';   // ✅ AS - UBICACIÓN (CORREGIDO)
            filaCompleta[45] = datosSanitizados.Contacto || '';    // ✅ AT - CONTACTO (CORREGIDO)
            
            // Total servicio solo en primera línea
            if (i === 0) {
                filaCompleta[22] = parseFloat(datosSanitizados.Total) || 0; // W - Total servicio
            }

            // DATOS DE LÍNEA
            filaCompleta[8] = linea.cod || '';                    // I - COD
            filaCompleta[9] = linea.descripcion || '';            // J - DESCRIPCION
            filaCompleta[10] = linea.und_medida || 'HORAS';       // K - UND
            filaCompleta[11] = parseFloat(linea.cantidad) || 0;   // L - UND. PEDIDO
            filaCompleta[14] = parseFloat(linea.precio) || 0;     // O - PRECIO
            filaCompleta[16] = parseFloat(linea.subtotal) || 0;   // Q - M. PEDIDO
            filaCompleta[19] = parseFloat(linea.movilizacion) || 0; // T - MOV. Y DES. MOV.
            filaCompleta[29] = linea.und_horas_minimas || '';     // AD - UND. HORAS. MINIMAS
            filaCompleta[30] = linea.hora_segun || '';            // AE - HORAS SEGÚN
            filaCompleta[31] = parseFloat(linea.horas_minimas_num) || 0; // AF - HORAS MINIMAS
            filaCompleta[33] = parseFloat(linea.dias_cotizados) || 0; // AH - TOTAL DIAS

            // Log de verificación de columnas críticas
            if (i === 0) {
                Logger.log("✅ VERIFICACIÓN COLUMNAS CRÍTICAS:");
                Logger.log("   Col 43 (AR) - Dirección: " + filaCompleta[43]);
                Logger.log("   Col 44 (AS) - Contacto: " + filaCompleta[44]);
                Logger.log("   Col 4 (E) - Ejecutivo: " + filaCompleta[4]);
            }

            // Escribir en la hoja
            hojaCot.getRange(filaActual, 1, 1, filaCompleta.length).setValues([filaCompleta]);
            Logger.log(`   📝 Fila ${filaActual} escrita exitosamente`);
            
            filaActual++;
        }

        Logger.log("✅✅✅ GUARDADO EXITOSO - " + lineas.length + " líneas guardadas");
        Logger.log("🎯 Código: " + codigoPedido);
        Logger.log("📍 Dirección guardada en columna AS (44): " + datosSanitizados.Direccion);
        Logger.log("📞 Contacto guardado en columna AT (45): " + datosSanitizados.Contacto);

        return { 
            success: true, 
            message: datosSanitizados.numPedido ? 
                'Cotización actualizada con éxito. Código: ' + codigoPedido :
                'Cotización registrada con éxito. Código: ' + codigoPedido,
            codigoPedido: codigoPedido
        };

    } catch (error) {
        Logger.log(`❌❌❌ ERROR en guardarCotizacion: ${error.message}`);
        return manejarError('guardarCotizacion', error);
    }
}

/**
 * Función alias para mantener compatibilidad
 */
function guardarRegistro(datos) {
    Logger.log("📤 Función guardarRegistro llamada, redirigiendo a guardarCotizacion...");
    return guardarCotizacion(datos);
}

// ====================================================
// === FUNCIONES DE MÓDULO CONTACTOS OPTIMIZADAS ===
// ====================================================

function getContactosYDirecciones(ruc) {
    try {
        const contactos = obtenerDatosFiltrados(HOJA_CONTACTOS, { 'RUC': ruc });
        const direcciones = obtenerDatosFiltrados(HOJA_DIRECCIONES, { 'RUC': ruc });
        
        return {
            contactos: contactos,
            direcciones: direcciones
        };
    } catch (error) {
        Logger.log(`ERROR en getContactosYDirecciones: ${error.message}`);
        return { contactos: [[]], direcciones: [[]] };
    }
}

function getFilaPorRowIndex(ruc, rowIndex, tipo) {
    try {
        const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
        let sheetName;
        if (tipo === 'contacto') sheetName = HOJA_CONTACTOS;
        else if (tipo === 'direccion') sheetName = HOJA_DIRECCIONES;
        else if (tipo === 'servicio') sheetName = HOJA_SERVICIOS;
        else throw new Error("Tipo de búsqueda inválido.");
        
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet || rowIndex < 2 || rowIndex > sheet.getLastRow()) return null;
        const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
        return { row: rowData, rowIndex: rowIndex };
    } catch (error) {
        Logger.log(`Error en getFilaPorRowIndex: ${error.message}`);
        return null;
    }
}

function guardarOActualizarContacto(data) {
    try {
        const datosSanitizados = sanitizarDatos(data);
        
        const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL); 
        const sheet = ss.getSheetByName(HOJA_CONTACTOS);
        if (!sheet) throw new Error(`La hoja '${HOJA_CONTACTOS}' no fue encontrada.`);
        const idContacto = datosSanitizados.ID_CONTACTO || `CON-${new Date().getTime()}`;
        const nuevosValores = [idContacto, datosSanitizados.RUC, datosSanitizados.NOMBRE, datosSanitizados.EMAIL, datosSanitizados.TELEFONO, datosSanitizados.CARGO];
        const rowIndex = parseInt(datosSanitizados.rowIndex);
        if (rowIndex > 1) { 
            sheet.getRange(rowIndex, 1, 1, nuevosValores.length).setValues([nuevosValores]);
        } else {
            sheet.appendRow(nuevosValores);
        }
        return { success: true, message: "Contacto guardado exitosamente" };
    } catch (error) {
        Logger.log(`Error en guardarOActualizarContacto: ${error.message}`);
        throw new Error("Error al guardar el contacto: " + error.message);
    }
}

function guardarOActualizarDireccion(data) {
    try {
        const datosSanitizados = sanitizarDatos(data);
        
        const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL); 
        const sheet = ss.getSheetByName(HOJA_DIRECCIONES);
        if (!sheet) throw new Error(`La hoja '${HOJA_DIRECCIONES}' no fue encontrada.`);
        const idDireccion = datosSanitizados.ID_DIRECCION || `DIR-${new Date().getTime()}`;
        const nuevosValores = [idDireccion, datosSanitizados.RUC, datosSanitizados.TIPO, datosSanitizados.DIRECCION, datosSanitizados.CIUDAD];
        const rowIndex = parseInt(datosSanitizados.rowIndex);
        if (rowIndex > 1) { 
            sheet.getRange(rowIndex, 1, 1, nuevosValores.length).setValues([nuevosValores]);
        } else {
            sheet.appendRow(nuevosValores);
        }
        return { success: true, message: "Dirección guardada exitosamente" };
    } catch (error) {
        Logger.log(`Error en guardarOActualizarDireccion: ${error.message}`);
        throw new Error("Error al guardar la dirección: " + error.message);
    }
}

// ====================================================
// === FUNCIONES DE CLIENTES OPTIMIZADAS ===
// ====================================================

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
        
        const resultado = guardarEnHoja(HOJA_CLIENTES, {
            ...datosSanitizados,
            valores: nuevosValores,
            rowIndex: 0 // Forzar creación nueva
        });
        
        return { 
            success: true, 
            message: "Cliente registrado exitosamente.",
            nuevoCliente: nuevosValores
        };
    } catch (error) {
        Logger.log(`ERROR en guardarNuevoCliente: ${error.message}`);
        throw new Error("Error al guardar el cliente: " + error.message);
    }
}

// ====================================================
// === FUNCIONES DE RESUMEN COMERCIAL OPTIMIZADAS ===
// ====================================================

function getListaCotizacionesResumen() {
    try {
        Logger.log("⚡ CARGANDO RESUMEN OPTIMIZADO...");
        const startTime = new Date().getTime();
        
        const allData = obtenerDatosHoja(HOJA_COTIZACIONES, true, 10); // Cache de 10 minutos
        const endGetData = new Date().getTime();
        Logger.log(`📊 Datos obtenidos en: ${endGetData - startTime}ms`);
        
        if (allData.length <= 1) {
            // Se añade la columna "Row Index" al encabezado
            return [["N° Pedido", "Fecha", "Cliente", "Monto Total", "Moneda", "Estado", "ID Cliente", "Row Index"]];
        }
        
        const INDICES = {
            NUM_PEDIDO: 0, FECHA_CREACION: 2, MONEDA: 15, CLIENTE: 6, 
            ESTADO: 7, ID_CLIENTE: 5, MONTO_TOTAL: 22
        };
        
        const resumenMap = new Map();
        const scriptTimeZone = Session.getScriptTimeZone();

        // Procesamiento optimizado
        for (let i = 1; i < allData.length; i++) {
            const row = allData[i];
            const numPedido = String(row[INDICES.NUM_PEDIDO] || '').trim();
            if (!numPedido) continue;
            
            const montoLinea = procesarMontoRapido(row[INDICES.MONTO_TOTAL]);
            
            if (!resumenMap.has(numPedido)) {
                const fecha = formatearFechaRapido(row[INDICES.FECHA_CREACION], scriptTimeZone);
                
                // 🔑 Se guarda el índice de fila de la hoja de cálculo (i + 1)
                const rowIndex = i + 1; 

                resumenMap.set(numPedido, {
                    numPedido: numPedido,
                    fecha: fecha,
                    moneda: String(row[INDICES.MONEDA] || 'SOL').toUpperCase().trim(),
                    cliente: row[INDICES.CLIENTE] || 'Cliente Desconocido',
                    estado: row[INDICES.ESTADO] || 'Pendiente',
                    idCliente: row[INDICES.ID_CLIENTE] || '',
                    montoTotal: montoLinea,
                    rowIndex: rowIndex // <-- ¡NUEVO CAMPO AGREGADO!
                });
            } else {
                const existente = resumenMap.get(numPedido);
                existente.montoTotal += montoLinea;
            }
        }
        
        // Crear array final de forma eficiente
        // Se añade la columna "Row Index" al encabezado
        const resumenData = [["N° Pedido", "Fecha", "Cliente", "Monto Total", "Moneda", "Estado", "ID Cliente", "Row Index"]];
        
        resumenMap.forEach(item => {
            resumenData.push([
                item.numPedido,
                item.fecha,
                item.cliente,
                item.montoTotal.toFixed(2),
                item.moneda,
                item.estado,
                item.idCliente,
                item.rowIndex // <-- ¡NUEVO VALOR AGREGADO!
            ]);
        });
        
        const endTime = new Date().getTime();
        Logger.log(`✅ RESUMEN GENERADO en ${endTime - startTime}ms - ${resumenData.length - 1} pedidos únicos`);
        
        return resumenData;

    } catch (e) {
        Logger.log(`❌ ERROR en resumen: ${e.message}`);
        return [["N° Pedido", "Fecha", "Cliente", "Monto Total", "Moneda", "Estado", "ID Cliente", "Row Index"]];
    }
}

// Funciones auxiliares optimizadas
function procesarMontoRapido(rawMonto) {
    if (typeof rawMonto === 'number') return rawMonto;
    if (!rawMonto) return 0;
    
    // Conversión rápida sin regex pesado
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

function formatearFechaRapido(rawFecha, timeZone) {
    if (!rawFecha) return '';
    if (rawFecha instanceof Date) {
        return Utilities.formatDate(rawFecha, timeZone, 'dd/MM/yyyy');
    }
    return String(rawFecha).substring(0, 10);
}

function getCotizacionPorCodigo(codigoCotizacion) {
    try {
        const resultado = buscarRegistro(HOJA_COTIZACIONES, codigoCotizacion, 0);
        if (!resultado) return null;
        
        const headers = resultado.encabezados;
        const row = resultado.datos;
        
        const columnMapping = {
            cot: headers.indexOf('COT'),
            num: headers.indexOf('NUM'),
            fecha: headers.indexOf('FECHA COT'),
            empresa: headers.indexOf('EMPRESA'),
            ejecutivo: headers.indexOf('EJECUTIVO'),
            ruc: headers.indexOf('ID CLIENTE'),
            cliente: headers.indexOf('CLIENTE'),
            estado: headers.indexOf('ESTADO COT'),
            cod: headers.indexOf('COD'),
            descripcion: headers.indexOf('DESCRIPCION'),
            und: headers.indexOf('UND'),
            cantidad: headers.indexOf('UND. PEDIDO'),
            precio: headers.indexOf('PRECIO'),
            moneda: headers.indexOf('MONEDA'),
            subtotal: headers.indexOf('M. PEDIDO'),
            movilizacion: headers.indexOf('MOV. Y DES. MOV.'),
            turno: headers.indexOf('TURNO'),
            und_horas_minimas: headers.indexOf('UND. HORAS. MINIMAS'),
            hora_segun: headers.indexOf('HORAS SEGÚN'),
            horas_minimas: headers.indexOf('HORAS MINIMAS'),
            dias_cotizados: headers.indexOf('TOTAL DIAS'),
            forma_pago: headers.indexOf('F. PAGO'),
            ubicacion: headers.indexOf('UBICACIÓN'),
            contacto: headers.indexOf('CONTACTO'),
            total: headers.indexOf('Total servicio')
        };
        
        return {
            COT: row[columnMapping.cot] || '',
            NUM: row[columnMapping.num] || '',
            FECHA: row[columnMapping.fecha] || '',
            EMPRESA: row[columnMapping.empresa] || '',
            EJECUTIVO: row[columnMapping.ejecutivo] || '',
            RUC: row[columnMapping.ruc] || '',
            CLIENTE: row[columnMapping.cliente] || '',
            ESTADO: row[columnMapping.estado] || '',
            COD: row[columnMapping.cod] || '',
            DESCRIPCION: row[columnMapping.descripcion] || '',
            UND: row[columnMapping.und] || '',
            CANTIDAD: row[columnMapping.cantidad] || 0,
            PRECIO: row[columnMapping.precio] || 0,
            MONEDA: row[columnMapping.moneda] || '',
            SUBTOTAL: row[columnMapping.subtotal] || 0,
            MOVILIZACION: row[columnMapping.movilizacion] || 0,
            TURNO: row[columnMapping.turno] || '',
            UND_HORAS_MINIMAS: row[columnMapping.und_horas_minimas] || '',
            HORA_SEGUN: row[columnMapping.hora_segun] || '',
            HORAS_MINIMAS: row[columnMapping.horas_minimas] || 0,
            DIAS_COTIZADOS: row[columnMapping.dias_cotizados] || 0,
            FORMA_PAGO: row[columnMapping.forma_pago] || '',
            UBICACION: row[columnMapping.ubicacion] || '',
            CONTACTO: row[columnMapping.contacto] || '',
            TOTAL: row[columnMapping.total] || 0
        };
    } catch (error) {
        Logger.log(`ERROR en getCotizacionPorCodigo: ${error.message}`);
        return null;
    }
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
        Logger.log("No se pudo obtener el email: " + error.message);
        return null;
    }
}

/**
 * Limpia cache global
 */
function limpiarCache() {
    try {
        cache.removeAll();
        Logger.log("✅ Cache limpiado exitosamente");
        return { success: true, message: "Cache limpiado" };
    } catch (error) {
        Logger.log("❌ Error al limpiar cache: " + error.message);
        return { success: false, message: "Error al limpiar cache" };
    }
}

function diagnosticarResumen() {
    try {
        Logger.log("🩺 INICIANDO DIAGNÓSTICO DE RESUMEN...");
        
        const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
        const hojaCot = ss.getSheetByName(HOJA_COTIZACIONES);
        
        if (!hojaCot) {
            Logger.log("❌ NO SE ENCUENTRA LA HOJA: " + HOJA_COTIZACIONES);
            return { error: "Hoja no encontrada: " + HOJA_COTIZACIONES };
        }
        
        const lastRow = hojaCot.getLastRow();
        const lastCol = hojaCot.getLastColumn();
        
        Logger.log("📊 DIMENSIONES DE LA HOJA:");
        Logger.log(" - Última fila: " + lastRow);
        Logger.log(" - Última columna: " + lastCol);
        
        if (lastRow <= 1) {
            Logger.log("⚠️ HOJA VACÍA (solo encabezados)");
            return { vacia: true };
        }
        
        // Verificar encabezados
        const headers = hojaCot.getRange(1, 1, 1, lastCol).getValues()[0];
        Logger.log("📋 ENCABEZADOS ENCONTRADOS:");
        headers.forEach((header, index) => {
            Logger.log(`   Col ${index + 1}: "${header}"`);
        });
        
        // Verificar primeros 5 pedidos
        const data = hojaCot.getRange(2, 1, Math.min(5, lastRow-1), lastCol).getValues();
        Logger.log("🔍 PRIMEROS 5 REGISTROS:");
        data.forEach((row, index) => {
            Logger.log(`   Fila ${index + 2}: ${row[0]} | ${row[2]} | ${row[6]} | ${row[22]}`);
        });
        
        // Probar la función de resumen
        const resultadoResumen = getListaCotizacionesResumen();
        
        return {
            success: true,
            dimensiones: { filas: lastRow, columnas: lastCol },
            totalPedidos: resultadoResumen.length - 1,
            muestraPrimeros: data.length
        };
        
    } catch (error) {
        Logger.log("❌ ERROR EN DIAGNÓSTICO: " + error.message);
        return { error: error.message };
    }
}

/**
 * Cargar resumen por partes (paginación) - MÁS RÁPIDO
 */
function getResumenPaginado(pagina = 1, filasPorPagina = 50) {
    try {
        Logger.log(`⚡ CARGANDO PÁGINA ${pagina} (${filasPorPagina} filas)`);
        const startTime = new Date().getTime();
        
        const allData = obtenerDatosHoja(HOJA_COTIZACIONES, true, 10);
        
        if (allData.length <= 1) {
            return {
                datos: [["N° Pedido", "Fecha", "Cliente", "Monto Total", "Moneda", "Estado", "ID Cliente"]],
                pagina: 1,
                totalPaginas: 1,
                totalFilas: 0
            };
        }
        
        const INDICES = { NUM_PEDIDO: 0, FECHA_CREACION: 2, MONEDA: 15, CLIENTE: 6, ESTADO: 7, ID_CLIENTE: 5, MONTO_TOTAL: 22 };
        const resumenMap = new Map();
        const scriptTimeZone = Session.getScriptTimeZone();

        // Procesar solo las filas necesarias
        const startIndex = 1; // Saltar encabezados
        const endIndex = Math.min(allData.length, startIndex + (pagina * filasPorPagina));
        
        for (let i = startIndex; i < endIndex; i++) {
            const row = allData[i];
            const numPedido = String(row[INDICES.NUM_PEDIDO] || '').trim();
            if (!numPedido) continue;
            
            const montoLinea = procesarMontoRapido(row[INDICES.MONTO_TOTAL]);
            
            if (!resumenMap.has(numPedido)) {
                resumenMap.set(numPedido, {
                    numPedido: numPedido,
                    fecha: formatearFechaRapido(row[INDICES.FECHA_CREACION], scriptTimeZone),
                    moneda: String(row[INDICES.MONEDA] || 'SOL').toUpperCase().trim(),
                    cliente: row[INDICES.CLIENTE] || 'Cliente Desconocido',
                    estado: row[INDICES.ESTADO] || 'Pendiente',
                    idCliente: row[INDICES.ID_CLIENTE] || '',
                    montoTotal: montoLinea
                });
            } else {
                resumenMap.get(numPedido).montoTotal += montoLinea;
            }
        }
        
        // Convertir a array
        const resumenData = [["N° Pedido", "Fecha", "Cliente", "Monto Total", "Moneda", "Estado", "ID Cliente"]];
        resumenMap.forEach(item => {
            resumenData.push([
                item.numPedido,
                item.fecha,
                item.cliente,
                item.montoTotal.toFixed(2),
                item.moneda,
                item.estado,
                item.idCliente
            ]);
        });
        
        const endTime = new Date().getTime();
        Logger.log(`✅ PÁGINA ${pagina} en ${endTime - startTime}ms - ${resumenData.length} filas`);
        
        return {
            datos: resumenData,
            pagina: pagina,
            totalPaginas: Math.ceil((allData.length - 1) / filasPorPagina),
            totalFilas: allData.length - 1
        };
        
    } catch (error) {
        Logger.log(`❌ ERROR en resumen paginado: ${error.message}`);
        return {
            datos: [["Error", "Error al cargar", "", "", "", "", ""]],
            pagina: 1,
            totalPaginas: 1,
            totalFilas: 0
        };
    }
}

/**
 * Limpiar cache específico para forzar recarga
 */
function limpiarCacheResumen() {
    try {
        cache.remove(`hoja_${HOJA_COTIZACIONES}`);
        Logger.log("✅ Cache de resumen limpiado");
        return { success: true, message: "Cache limpiado, la próxima carga será fresca" };
    } catch (error) {
        Logger.log("❌ Error limpiando cache: " + error.message);
        return { success: false, message: error.message };
    }
}

function obtenerPedidoParaEdicion(numPedido) {
    const startTime = new Date().getTime();
    Logger.log(`⚡ SOLICITANDO EDICIÓN OPTIMIZADA: ${numPedido}`);
    
    try {
        // Cache específico para pedidos individuales
        const cacheKey = `pedido_${numPedido}`;
        const cached = cache.get(cacheKey);
        
        if (cached) {
            const endTime = new Date().getTime();
            Logger.log(`✅ CACHE HIT - Pedido ${numPedido} en ${endTime - startTime}ms`);
            return cached; // Ya viene como string JSON
        }
        
        const allData = obtenerDatosHoja(HOJA_COTIZACIONES, true, 15); // Cache más largo
        const endGetData = new Date().getTime();
        Logger.log(`📊 Datos obtenidos en: ${endGetData - startTime}ms`);
        
        const pedidoBuscado = String(numPedido).trim().toUpperCase();
        const filasCoincidentes = [];
        
        // Búsqueda optimizada - solo columnas necesarias
        for (let i = 1; i < allData.length; i++) {
            const fila = allData[i];
            const codigoPedido = String(fila[0] || '').trim().toUpperCase();
            
            if (codigoPedido === pedidoBuscado) {
                filasCoincidentes.push({
                    data: fila,
                    rowIndex: i + 1
                });
                
                // Si ya encontramos todas las líneas, salir temprano
                if (filasCoincidentes.length > 50) { // Límite razonable
                    Logger.log("⚠️ Límite de líneas alcanzado (50)");
                    break;
                }
            }
        }
        
        Logger.log(`📋 Filas encontradas: ${filasCoincidentes.length}`);
        
        if (filasCoincidentes.length === 0) {
            throw new Error("Pedido no encontrado: " + pedidoBuscado);
        }
        
        // Procesar datos generales (solo primera fila)
        const primeraFila = filasCoincidentes[0].data;
        
        // ✅ CORRECCIÓN: Obtener datos de columnas correctas
        const ejecutivo = primeraFila[4] || '';
        const direccion = primeraFila[44] || ''; // Columna AS
        const contacto = primeraFila[45] || '';  // Columna AT
        
        const datosGenerales = {
            fechaRegistro: primeraFila[2] ? new Date(primeraFila[2]).toISOString() : new Date().toISOString(),
            Empresa: primeraFila[3] || '',
            RUC: String(primeraFila[5] || ''),
            Cliente: primeraFila[6] || '',
            Estado: primeraFila[7] || 'COTIZACION',
            Moneda: primeraFila[15] || 'SOLES',
            Forma_Pago: primeraFila[38] || '',
            Direccion: direccion,
            Turno: normalizarTurno(primeraFila[25] || 'Diurno'),
            Ejecutivo: ejecutivo,
            Contacto: contacto
        };
        
        // Procesar líneas de servicios (optimizado)
        const lineas = [];
        filasCoincidentes.forEach((filaObj, index) => {
            const fila = filaObj.data;
            
            // Solo procesar si tiene datos relevantes
            if (fila[8] || fila[9] || parseFloat(fila[11] || 0) > 0) {
                lineas.push({
                    cod: String(fila[8] || ''),
                    descripcion: String(fila[9] || ''),
                    cantidad: parseFloat(fila[11] || 1),
                    und_medida: String(fila[10] || 'HORAS'),
                    und_horas_minimas: String(fila[29] || ''),
                    dias_cotizados: parseFloat(fila[33] || 0),
                    horas_minimas_num: parseFloat(fila[31] || 0),
                    hora_segun: String(fila[30] || ''),
                    movilizacion: parseFloat(fila[19] || 0),
                    precio: parseFloat(fila[14] || 0),
                    subtotal: parseFloat(fila[16] || 0)
                });
            }
        });
        
        const resultado = {
            success: true,
            ...datosGenerales,
            Lineas: lineas,
            Total: parseFloat(primeraFila[22] || 0),
            totalLineas: lineas.length,
            numPedido: numPedido,
            usuario: obtenerEmailSeguro(),
            autorizado: true,
            tiempoCarga: new Date().getTime() - startTime
        };
        
        const resultadoJSON = JSON.stringify(resultado);
        
        // Guardar en cache por 10 minutos
        cache.put(cacheKey, resultadoJSON, 600);
        
        Logger.log(`✅✅✅ EDICIÓN EXITOSA en ${new Date().getTime() - startTime}ms - ${lineas.length} líneas`);
        
        return resultadoJSON;
        
    } catch (error) {
        Logger.log(`❌❌❌ ERROR en edición: ${error.message} - Tiempo: ${new Date().getTime() - startTime}ms`);
        
        return JSON.stringify({
            success: false,
            message: "Error al cargar el pedido: " + error.message,
            error: error.message,
            tiempoCarga: new Date().getTime() - startTime
        });
    }
}

// Función auxiliar optimizada
function normalizarTurno(turno) {
    if (!turno || typeof turno !== 'string') return 'Diurno';
    
    const turnoUpper = turno.toUpperCase().trim();
    if (turnoUpper.includes('DIURNO')) return 'Diurno';
    if (turnoUpper.includes('NOCTURNO')) return 'Nocturno';
    if (turnoUpper.includes('DOBLE')) return 'Doble Turno';
    return 'Diurno';
}
/**
 * Función específica para obtener servicio por rowIndex
 */
function obtenerServicioPorRowIndex(rowIndex) {
    try {
        Logger.log(`🔍 Buscando servicio en fila: ${rowIndex}`);
        
        const allData = obtenerDatosHoja(HOJA_SERVICIOS);
        
        if (rowIndex < 2 || rowIndex > allData.length) {
            throw new Error("Índice de fila inválido para servicios");
        }
        
        const filaDatos = allData[rowIndex - 1]; // -1 porque allData[0] son encabezados
        const encabezados = allData[0];
        
        Logger.log(`✅ Servicio encontrado: ${filaDatos[0]} - ${filaDatos[1]}`);
        
        return {
            row: filaDatos,
            rowIndex: rowIndex,
            encabezados: encabezados
        };
        
    } catch (error) {
        Logger.log(`❌ ERROR en obtenerServicioPorRowIndex: ${error.message}`);
        throw new Error("No se pudo cargar el servicio: " + error.message);
    }
}

function getListaClientes() {
    try {
        const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
        const sheet = ss.getSheetByName(HOJA_CLIENTES);
        if (!sheet || sheet.getLastRow() < 1) {
            return [['RUC', 'NOMBRE/RAZÓN SOCIAL'], ['Error:', 'Verifique su configuración']];
        }
        return sheet.getDataRange().getValues();
    } catch (e) {
        Logger.log("Error en getListaClientes: " + e.message);
        return [['RUC', 'NOMBRE/RAZÓN SOCIAL'], ['Error Grave:', e.message]];
    }
}

// Función para obtener contactos y direcciones de un cliente
function getContactosYDirecciones(ruc) {
    try {
        const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
        const datos = { contactos: [[]], direcciones: [[]] };
        
        const obtenerDatosFiltrados = (sheetName, colMap) => {
            const sheet = ss.getSheetByName(sheetName);
            if (!sheet || sheet.getLastRow() < 1) return [[]];
            const allData = sheet.getDataRange().getValues();
            const headers = allData[0];
            if (allData.length === 1) return [headers];
            const filteredRows = allData.slice(1).filter(row => String(row[colMap.RUC]) == String(ruc));
            const rowsWithIndex = filteredRows.map(row => {
                const rowIndex = allData.findIndex(dataRow => JSON.stringify(dataRow) === JSON.stringify(row));
                return row.concat(rowIndex + 1); 
            });
            return [headers].concat(rowsWithIndex);
        };

        datos.contactos = obtenerDatosFiltrados(HOJA_CONTACTOS, CONTACTO_COLS);
        datos.direcciones = obtenerDatosFiltrados(HOJA_DIRECCIONES, DIRECCION_COLS);
        return datos;
    } catch (error) {
        Logger.log(`Error en getContactosYDirecciones: ${error.message}`);
        return { contactos: [[]], direcciones: [[]] };
    }
}
/**
 * Genera un PDF de la cotización usando las plantillas de Google Sheets.
 * Selecciona la plantilla según la ubicación del servicio.
 * * @param {Object} cotizacionData - Objeto con los datos de la cotización. 
 * Debe contener: ruc, cliente, fecha, contacto, 
 * correoCelular, lugar, turno, duracion, total, 
 * y un array de objetos 'servicios' (descripcion, valor).
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
        throw new Error("❌ Ubicación no reconocida para la plantilla de cotización: " + cotizacionData.lugar);
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
    let hojaEliminar = hojaTemporal; // Bandera para la limpieza final

    try {
        // 3. Rellenar los datos de cabecera
        hojaTemporal.getRange('C5').setValue(cotizacionData.cliente || '');
        hojaTemporal.getRange('C6').setValue(cotizacionData.ruc || '');
        hojaTemporal.getRange('C7').setValue(cotizacionData.contacto || '');
        hojaTemporal.getRange('C8').setValue(cotizacionData.correoCelular || '');
        hojaTemporal.getRange('C12').setValue(cotizacionData.lugar || '');
        hojaTemporal.getRange('C13').setValue(cotizacionData.turno || '');
        hojaTemporal.getRange('C14').setValue(cotizacionData.duracion || '');

        // Formatear y colocar la fecha en G3
        let fechaCotizacion = cotizacionData.fecha ? new Date(cotizacionData.fecha) : new Date();
        hojaTemporal.getRange('G3').setValue(fechaCotizacion).setNumberFormat('dd/MM/yyyy'); // Ajustar formato de fecha

        // 4. Agregar las líneas de servicio
        const numServicios = SERVICIOS.length;
        if (numServicios > 1) {
            // Si hay N servicios, se insertan N-1 filas nuevas debajo de la fila de inicio (18)
            hojaTemporal.insertRowsAfter(START_ROW, numServicios - 1);
        }

        const dataServicios = SERVICIOS.map((s, index) => [
            index + 1,            // A: Número de Item (A18, A19...)
            s.descripcion || '',  // B: Descripción del Servicio (B18, B19...)
            s.valor || 0          // F: Valor del Servicio (F18, F19...)
        ]);
        
        // Rango de la columna A a la F
        if (dataServicios.length > 0) {
            // Escribir los datos en el rango A18:F(18 + numServicios - 1)
            // Se asume que las columnas C, D y E se mantienen vacías en el loop para escribir solo A, B y F
            // Se usa setValues() con un array 2D de 6 columnas y se asume que las plantillas manejan las columnas vacías.
            // Para ser más eficientes, solo se escriben las columnas necesarias (A, B, F). 
            
            // Creamos un array 2D completo de A a F para el setValues
            const dataCompleta = SERVICIOS.map((s, index) => [
                index + 1,            // Columna A
                s.descripcion || '',  // Columna B
                '',                   // Columna C
                '',                   // Columna D
                '',                   // Columna E
                s.valor || 0          // Columna F
            ]);
            
            hojaTemporal.getRange(START_ROW, 1, dataCompleta.length, 6).setValues(dataCompleta);
        }
        
        let currentRow = START_ROW + numServicios; // Fila inmediatamente debajo del último servicio

        // 5. Agregar la línea de total
        hojaTemporal.getRange('A' + currentRow).setValue('TOTAL');
        hojaTemporal.getRange('F' + currentRow).setValue(cotizacionData.total || 0);

        // 6. Generar el PDF
        
        // Obtener URL de exportación con parámetros específicos para la hoja temporal
        const URL = ssTemplate.getUrl();
        const exportUrl = URL.replace('/edit', '/export?exportFormat=pdf' +
            '&gid=' + hojaTemporal.getSheetId() + // ID de la hoja TEMPORAL (clave para exportar solo esta)
            '&format=pdf' +
            '&size=A4' + // Formato A4
            '&portrait=true' + // Orientación vertical
            '&fitw=true' + // Ajustar a ancho
            '&gridlines=false' + // Ocultar cuadrículas
            '&sheetnames=false'); // Ocultar nombres de hoja

        const response = UrlFetchApp.fetch(exportUrl, {
            headers: {
                'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(), // Necesario para acceso a Drive
            },
            muteHttpExceptions: true
        });

        const pdfBlob = response.getBlob().setName((cotizacionData.cliente || 'CLIENTE') + "_COTIZACION.pdf");
        
        return pdfBlob; 

    } catch (e) {
        Logger.log('Error al generar PDF: ' + e.toString());
        // Propaga el error para que el cliente lo sepa
        throw new Error("Error en el PDF: " + e.message); 
    } finally {
        // 7. Limpieza: Eliminar la hoja temporal
        if (hojaEliminar) {
            ssTemplate.deleteSheet(hojaEliminar);
        }
    }
}

function pruebaGenerarPDF() {
    const dataEjemplo = {
        ruc: '20512345678',
        cliente: 'Empresa Demo S.A.C.',
        fecha: new Date(),
        contacto: 'Juan Pérez',
        correoCelular: 'juan.perez@demo.com / 987654321',
        lugar: 'San Jose', // Esto determinará que se use la plantilla COT_GSJ
        turno: 'Día',
        duracion: '4 horas',
        servicios: [
            { descripcion: 'Instalación de tubería 4 pulgadas', valor: 1500.00 },
            { descripcion: 'Soporte técnico y mantenimiento', valor: 500.00 },
            { descripcion: 'Servicio de consultoría y diseño', valor: 250.00 }
        ],
        total: 2250.00
    };
    
    try {
        const pdfFile = generarPDFCotizacion(dataEjemplo);
        
        // Opcional: Guardar el archivo en la raíz de tu Drive para verificación
        const archivoEnDrive = DriveApp.createFile(pdfFile);
        Logger.log('✅ PDF generado y guardado: ' + archivoEnDrive.getUrl());
        
        // Si quieres adjuntarlo a un correo:
        // MailApp.sendEmail('tu_correo@example.com', 'Cotización Generada', 'Adjunto la cotización.', {
        //     attachments: [pdfFile]
        // });

    } catch (error) {
        Logger.log(error);
        SpreadsheetApp.getUi().alert('Error al generar PDF: ' + error.message);
    }
}
/**
 * @param {number} rowIndex El número de fila del pedido en la HOJA_COTIZACIONES.
 * @returns {Object} Un objeto con { success: true, pdfUrl: URL } o { success: false, error: mensaje }.
 */
function generarYDevolverPDF(rowIndex) {
    try {
        // 1. Obtener la data del pedido (usando la función que ya debes tener)
        // Se asume que tienes una función para obtener una fila completa por rowIndex.
        const pedidoRow = crudHoja('READ_ROW', HOJA_COTIZACIONES, { rowIndex: rowIndex });

if (!pedidoRow) {
    return { success: false, error: "No se encontró el pedido con el índice: " + rowIndex };
}
        
        // 2. Formatear los datos al objeto que necesita generarPDFCotizacion
        // ADVERTENCIA: Debes mapear tu fila (array) de datos a un objeto. 
        // Este es un ejemplo basado en los datos que usa tu función:
        const dataParaPDF = {
            ruc: pedidoRow[2], // Asume RUC en columna 3 (índice 2)
            cliente: pedidoRow[3], // Asume Cliente en columna 4 (índice 3)
            fecha: pedidoRow[4], // Asume Fecha en columna 5 (índice 4)
            contacto: pedidoRow[5], // Asume Contacto en columna 6 (índice 5)
            correoCelular: pedidoRow[6], // Asume Correo/Celular en columna 7 (índice 6)
            lugar: pedidoRow[7], // Asume Lugar en columna 8 (índice 7)
            turno: pedidoRow[8], // Asume Turno en columna 9 (índice 8)
            duracion: pedidoRow[9], // Asume Duración en columna 10 (índice 9)
            // Se asume que necesitarás buscar los servicios asociados a este pedido
            // Por simplicidad, se usará data fija, pero en una implementación real, 
            // necesitarías una función para obtener el detalle de la cotización.
            servicios: [
                { descripcion: "Servicio 1", valor: 100 }, 
                { descripcion: "Servicio 2", valor: 50 } 
            ],
            total: 150 // Asume Total en alguna columna (debes ajustarlo)
        };
        
        // 3. Generar el PDF (retorna un Blob)
        const pdfBlob = generarPDFCotizacion(dataParaPDF);

        // 4. Guardar el PDF Blob temporalmente en Drive
        const archivoEnDrive = DriveApp.createFile(pdfBlob);
        
        // 5. Devolver la URL temporal
        return { 
            success: true, 
            pdfUrl: archivoEnDrive.getUrl()
        };

    } catch (e) {
        Logger.log("Error en generarYDevolverPDF: " + e.toString());
        return { success: false, error: e.message };
    }
}