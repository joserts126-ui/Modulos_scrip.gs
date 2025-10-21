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
// === FUNCIONES NUCLEO DE DATOS (CRUD y CACHE) ===
// ====================================================

/**
 * Obtiene el mapa de columnas (encabezado -> índice) para una hoja.
 * @param {string} sheetName El nombre de la hoja.
 * @returns {Object} Un mapa de {HEADER_NAME_UPPERCASE: index}.
 */
function getColumnMap(sheetName) {
    const cacheKey = `header_map_${sheetName}`;
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
