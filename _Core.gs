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
    
    const validPages = ['Modulos', 'Comercial', 'Servicios', 'Contactos', 'ResumenComercial', 'OT', 'RegistrarOT', 'Acta'];
    
    let tmpl;
    if (validPages.includes(page)) {
        tmpl = HtmlService.createTemplateFromFile(page);
    } else {
        tmpl = HtmlService.createTemplateFromFile('Modulos'); // Página por defecto
    }

    // --- CORRECCIÓN DEFINITIVA ---
    // Pasamos el objeto de parámetros tal cual.
    tmpl.parametros = e.parameter || {}; 
    // --- FIN DE LA CORRECCIÓN ---
    
    tmpl.permisos = obtenerPermisosUsuario();
    
    return tmpl.evaluate()
        .setTitle(page)
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
 * Obtiene el objeto de permisos para el usuario activo (v4).
 * Lee columnas específicas de vendedor (ANTHONY, RENATO, etc.) desde "Permisos".
 * CORREGIDO: Reemplazado el operador '...' por Object.assign() para compatibilidad ES5.
 */
function obtenerPermisosUsuario() {
  const email = obtenerEmailSeguro();
  const cache = CacheService.getScriptCache();
  const cacheKey = "permisos_v3_" + email; // Mantenemos la clave v3, no hay problema

  // 1. Intentar obtener desde la caché
  const permisosCacheados = cache.get(cacheKey);
  if (permisosCacheados) {
    return JSON.parse(permisosCacheados);
  }

  // 2. Definir permisos por defecto
  const permisosPorDefecto = {
    puedeEditarCotizacion: false,
    puedeEditarServicios: false,
    puedeEditarOT: false,
    puedeVerReportes: false,
    puedeVerTodasLasCotizaciones: false,
    nombreVendedorExacto: null,
    visibilidadVendedores: {} 
  };
  
  try {
    const data = obtenerDatosHoja(HOJA_PERMISOS); 
    if (data.length <= 1) return permisosPorDefecto;
    
    const COL_MAP = getColumnMap(HOJA_PERMISOS);
    const COL_EMAIL = COL_MAP['EMAIL'];
    const COL_NOMBRE_VENDEDOR = COL_MAP['NOMBREVENDEDOREXACTO'];

    if (COL_EMAIL === undefined) {
      Logger.log("Error de Permisos: La hoja 'Permisos' no tiene la columna 'EMAIL'.");
      return permisosPorDefecto;
    }

    const encabezados = data[0];
    const encabezadosUpper = encabezados.map(h => h ? h.toUpperCase().replace(/\s+/g, '') : '');
    let filaUsuario = null;

    // 3. Buscar la fila del usuario
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][COL_EMAIL] || '').trim().toLowerCase() === email.toLowerCase()) {
        filaUsuario = data[i];
        break;
      }
    }

    if (!filaUsuario) {
      Logger.log(`Permisos v3: Usuario '${email}' no encontrado. Usando permisos por defecto.`);
      return permisosPorDefecto;
    }

    // 4. Construir el objeto de permisos
    
    // --- INICIO DE LA CORRECCIÓN ---
    // const permisosReales = { ...permisosPorDefecto, visibilidadVendedores: {} }; // ESTA LÍNEA DABA EL ERROR
    
    // Esta es la versión compatible (ES5) que hace lo mismo:
    const permisosReales = Object.assign({}, permisosPorDefecto, {
      visibilidadVendedores: {}
    });
    // --- FIN DE LA CORRECCIÓN ---

    const indiceInicioVendedores = encabezadosUpper.indexOf('PUEDEVERVENTASDE');

    encabezados.forEach((headerOriginal, index) => {
      if (!headerOriginal) return; 
      const headerKey = encabezadosUpper[index]; 
      const valor = filaUsuario[index];

      if (headerKey.startsWith("PUEDE")) {
         const camelCaseKey = headerOriginal.charAt(0).toLowerCase() + headerOriginal.slice(1).replace(/([A-Z])/g, '$1').replace(/\s+/g, '');
         permisosReales[camelCaseKey] = (String(valor).toUpperCase() === 'VERDADERO' || String(valor).toUpperCase() === 'TRUE');
  
      }
      else if (headerKey === 'NOMBREVENDEDOREXACTO') {
        permisosReales.nombreVendedorExacto = String(valor || '').trim();
      }
      else if (indiceInicioVendedores !== -1 && index > indiceInicioVendedores) {
         const nombreVendedorColumna = headerOriginal.trim().toUpperCase();
         if (nombreVendedorColumna) { 
             permisosReales.visibilidadVendedores[nombreVendedorColumna] = (String(valor).toUpperCase() === 'VERDADERO' || String(valor).toUpperCase() === 'TRUE');
         }
      }
    });
    
    // 5. Guardar en caché y devolver
    cache.put(cacheKey, JSON.stringify(permisosReales), 300);
    Logger.log("Permisos v3 (ES5) cacheados para " + email + ": " + JSON.stringify(permisosReales));
    return permisosReales;

  } catch (e) {
    Logger.log(`Error crítico en obtenerPermisosUsuario v3: ${e.message}\n${e.stack}`);
    return permisosPorDefecto;
  }
}
