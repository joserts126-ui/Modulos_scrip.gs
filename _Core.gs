// ====================================================
// === CORE: MANEJO DE VISTAS (doGet) Y UTILIDADES ===
// ====================================================

// Cache para optimización
const cache = CacheService.getScriptCache();

/**
 * Maneja las solicitudes GET y sirve las páginas correspondientes
 */
function doGet(e) {
    const page = e.parameter.page || 'Modulos';
    
    // Añadir 'Acta' a la lista
    const validPages = ['Modulos', 'Comercial', 'Servicios', 'Contactos', 'ResumenComercial', 'OT', 'RegistrarOT', 'Acta', 'Valorizacion']
    
    if (validPages.includes(page)) {
        const tmpl = HtmlService.createTemplateFromFile(page);
        
        // --- INYECCIÓN DE PARÁMETROS (LA CORRECCIÓN) ---
        // Pasa todos los parámetros de la URL (e.g., 'pedido', 'editar', 'linea')
        // a una variable 'parametros' dentro del HTML.
        tmpl.parametros = e.parameter || {}; 
        // --- FIN DE LA CORRECCIÓN ---
        
        tmpl.permisos = obtenerPermisosUsuario();
        
        return tmpl.evaluate()
            .setTitle(page)
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    
    // Página por defecto si 'page' no es válida
    const defaultTmpl = HtmlService.createTemplateFromFile('Modulos');
    defaultTmpl.parametros = {}; // Pasar objeto vacío
    defaultTmpl.permisos = obtenerPermisosUsuario();
    return defaultTmpl.evaluate()
        .setTitle('Módulos Principales')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
// === FUNCIONES NUCLEO DE DATOS (CRUD y CACHE) ===
// ====================================================

/**
 * Obtiene el mapa de columnas (encabezado -> índice) para una hoja.
 * @param {string} sheetName El nombre de la hoja.
 * @returns {Object} Un mapa de {HEADER_NAME_UPPERCASE: index}.
 */
function getColumnMap(sheetName) {
    const cacheKey = `header_map_V5_${sheetName}`;
    let cachedMap = cache.get(cacheKey);

    if (cachedMap) {
        return JSON.parse(cachedMap);
    }

    try {
        const ss = SpreadsheetApp.openById(HOJA_ID_PRINCIPAL);
        const sheet = ss.getSheetByName(sheetName);

        if (!sheet || sheet.getLastRow() < 1) {
            Logger.log(`Advertencia: Hoja ${sheetName} vacía o no existe.`);
            return {};
        }

        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const map = {};
        
        headers.forEach((header, index) => {
            if (header) {
                map[header.trim().toUpperCase()] = index;
            }
        });

        cache.put(cacheKey, JSON.stringify(map), 3600); 
        return map;

    } catch (e) {
        Logger.log(`Error creando mapa de columnas para ${sheetName}: ${e.message}`);
        return {};
    }
}

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

            case 'READ_ROW': 
                const rowIndexToRead = parseInt(datos.rowIndex);
                if (rowIndexToRead > 1 && rowIndexToRead <= sheet.getLastRow()) {
                    const lastCol = sheet.getLastColumn();
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
        
        // Limitar cantidad de filas si es muy grande (máximo 1000 filas)
        const lastRow = Math.min(sheet.getLastRow(), 1000);
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

function forzarAutorizacion() {
  // Esta función solo existe para forzar la re-autorización
  try {
    const ss = SpreadsheetApp.openById("15qfA3idaLkyhvFwAeEZQo6L9BudBBfgnV8DFrs1qV6Y");
    const folder = DriveApp.getFolderById("1h4LZiA9Iwx54jHqOyHqFvDJykGipM6YO");
    Logger.log("Hoja: " + ss.getName() + ", Carpeta: " + folder.getName());
  } catch (e) {
    Logger.log("Error de autorización: " + e.message);
  }
}

/**
 * Obtiene el objeto de permisos para el usuario actual.
 * Usa caché para alto rendimiento.
 */
function obtenerPermisosUsuario() {
    const email = obtenerEmailSeguro(); // Asumo que esta función ya existe en tu lógica
    const cacheKey = 'permisos_' + email;
    const cache = CacheService.getScriptCache();
    
    // 1. Intentar obtener de la caché
    const cachedPermisos = cache.get(cacheKey);
    if (cachedPermisos) {
        return JSON.parse(cachedPermisos);
    }

    // 2. Si no está en caché, leer la hoja
    const HOJA_PERMISOS = "Permisos"; // Asegúrate que coincida
    const permisosData = obtenerDatosHoja(HOJA_PERMISOS, false); // false = no usar caché para leer la hoja
    const COL_MAP = getColumnMap(HOJA_PERMISOS);

    let permisosUsuario = {
        email: email,
        rol: 'Invitado', // Rol por defecto
        puedeEditarServicios: false,
        puedeEditarOT: false,
        puedeVerReportes: false,
        puedeEditarCotizacion: false
    };

    if (permisosData.length > 1) {
        const filaUsuario = permisosData.slice(1).find(fila => 
            String(fila[COL_MAP['EMAIL']] || '').trim().toLowerCase() === email.toLowerCase()
        );

        if (filaUsuario) {
            permisosUsuario.rol = filaUsuario[COL_MAP['ROL']] || 'Invitado';
            permisosUsuario.puedeEditarServicios = filaUsuario[COL_MAP['PUEDEEDITARSERVICIOS']] === true;
            permisosUsuario.puedeEditarOT = filaUsuario[COL_MAP['PUEDEEDITAROT']] === true;
            permisosUsuario.puedeVerReportes = filaUsuario[COL_MAP['PUEDEVERREPORTES']] === true;
            permisosUsuario.puedeEditarCotizacion = filaUsuario[COL_MAP['PUEDEEDITARCOTIZACION']] === true;
        }
    }
    
    // 3. Guardar en caché por 1 hora (3600 segundos)
    cache.put(cacheKey, JSON.stringify(permisosUsuario), 3600);
    
    return permisosUsuario;
}
