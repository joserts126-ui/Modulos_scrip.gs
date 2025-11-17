// Pega esto en un nuevo archivo _Auditoria.gs

/**
 * ====================================================
 * === SCRIPTS DE AUDITORÍA DE FASE 0 ("EL RADAR") ===
 * ====================================================
 * * Instrucciones:
 * 1. Pega este código en un nuevo archivo _Auditoria.gs en tu proyecto.
 * 2. Selecciona la función 'EJECUTAR_AUDITORIA_COMPLETA' en el editor.
 * 3. Haz clic en 'Ejecutar'.
 * 4. Revisa los registros (Logger.log o 'Ver' > 'Registros').
 * 5. Copia y pega los resultados para nuestro siguiente paso.
 * * NOTA: Estos scripts asumen que tus constantes (HOJA_ID_PRINCIPAL, HOJA_COTIZACIONES, 
 * HOJA_CLIENTES) están definidas en _Constants_Lists.gs o similar.
 */

/**
 * FUNCIÓN PRINCIPAL DE AUDITORÍA
 * Ejecuta todas las auditorías en secuencia.
 */
function EJECUTAR_AUDITORIA_COMPLETA() {
  Logger.log('🚀 INICIANDO AUDITORÍA COMPLETA FASE 0 - "EL RADAR"');
  Logger.log('===================================================');
  
  const resultados = {};
  
  try {
    // 1. Auditoría de Rendimiento
    Logger.log('\n📊 1. EJECUTANDO AUDITORÍA DE RENDIMIENTO...');
    resultados.rendimiento = auditarRendimiento();
    
    // 2. Auditoría de Datos Huérfanos
    Logger.log('\n🔍 2. EJECUTANDO AUDITORÍA DE DATOS HUÉRFANOS...');
    resultados.huerfanos = auditarDatosHuerfanos();
    
    // 3. Auditoría de Corrupción Texto Libre
    Logger.log('\n📝 3. EJECUTANDO AUDITORÍA DE CORRUPCIÓN TEXTO LIBRE...');
    resultados.corrupcion = auditarCorrupcionTextoLibre();
    
    Logger.log('\n🎯 AUDITORÍA COMPLETA FASE 0 FINALIZADA');
    Logger.log('====================================');
    Logger.log('📋 RESUMEN EJECUTIVO PARA GERENCIA:');
    Logger.log('• Rendimiento: Tiempos de funciones críticas registrados.');
    Logger.log('• Datos Huérfanos: ' + (resultados.huerfanos ? resultados.huerfanos.totalHuerfanos : 'ERROR') + ' registros problemáticos encontrados.');
    Logger.log('• Corrupción Texto: ' + (resultados.corrupcion ? resultados.corrupcion.totalVariantes : 'ERROR') + ' variantes de texto diferentes para el mismo servicio.');
    
  } catch (error) {
    Logger.log(`💥 ERROR CRÍTICO EN AUDITORÍA COMPLETA: ${error.message}\n${error.stack}`);
  }
}

/**
 * 1. AUDITORÍA DE RENDIMIENTO
 * Mide el tiempo de ejecución de las funciones más lentas.
 */
function auditarRendimiento() {
  Logger.log("--- Iniciando Auditoría de Rendimiento ---");
  try {
    // Prueba 1: Carga del resumen (una de las vistas más pesadas)
    console.time('Rendimiento: getListaCotizacionesResumen');
    getListaCotizacionesResumen(); // [Función real de _Comercial_Logic.gs]
    console.timeEnd('Rendimiento: getListaCotizacionesResumen');

    // Prueba 2: Carga de un pedido para edición (lectura intensiva de DataCot)
    const numPedidoPrueba = 'COT.GYM.2025.07.4003'; // Usa un pedido real de tu DataCot.csv
    console.time(`Rendimiento: obtenerPedidoParaEdicion (${numPedidoPrueba})`);
    obtenerPedidoParaEdicion(numPedidoPrueba); // [Función real de _Comercial_Logic.gs]
    console.timeEnd(`Rendimiento: obtenerPedidoParaEdicion (${numPedidoPrueba})`);

    // Prueba 3: Carga de datos pesados (función que tú mismo creaste)
    console.time('Rendimiento: getDatosPesadosComercial');
    getDatosPesadosComercial(); // [Función real de _Comercial_Logic.gs]
    console.timeEnd('Rendimiento: getDatosPesadosComercial');

    Logger.log("--- Auditoría de Rendimiento Finalizada ---");
    return "Completada";
    
  } catch (e) {
    Logger.log(`❌ Error en auditoría de rendimiento: ${e.message}`);
    return "Error";
  }
}

/**
 * 2. AUDITORÍA DE DATOS HUÉRFANOS (Riesgo #2)
 * Busca cotizaciones (DataCot) cuyo 'ID CLIENTE' no existe en la hoja 'Clientes'.
 */
function auditarDatosHuerfanos() {
  Logger.log("--- Iniciando Auditoría de Datos Huérfanos ---");
  try {
    // 1. Obtener mapa de columnas (usando tu función de _Core.gs)
    const mapCot = getColumnMap(HOJA_COTIZACIONES);
    const mapClientes = getColumnMap(HOJA_CLIENTES);

    // 2. Validar encabezados (basado en tus CSVs)
    const ID_CLIENTE_COL_COT = mapCot['ID CLIENTE']; // De DataCot.csv
    const COT_COL_COT = mapCot['COT']; // De DataCot.csv
    const RUC_COL_CLIENTES = mapClientes['RUC']; // De Clientes.csv

    if (ID_CLIENTE_COL_COT === undefined) throw new Error(`No se encontró la columna 'ID CLIENTE' en ${HOJA_COTIZACIONES}`);
    if (RUC_COL_CLIENTES === undefined) throw new Error(`No se encontró la columna 'RUC' en ${HOJA_CLIENTES}`);
    
    // 3. Leer datos (usando tu función de _Core.gs, sin caché)
    const allDataCot = obtenerDatosHoja(HOJA_COTIZACIONES, false);
    const allDataClientes = obtenerDatosHoja(HOJA_CLIENTES, false);

    // 4. Crear un Set de RUCs válidos para búsqueda rápida
    console.time('Construcción Set Clientes');
    const rucsValidos = new Set();
    for (let i = 1; i < allDataClientes.length; i++) {
      const ruc = allDataClientes[i][RUC_COL_CLIENTES];
      if (ruc) rucsValidos.add(String(ruc).trim());
    }
    console.timeEnd('Construcción Set Clientes');
    Logger.log(`Total de ${rucsValidos.size} RUCs válidos encontrados en ${HOJA_CLIENTES}.`);

    // 5. Buscar Huérfanos
    console.time('Búsqueda de Huérfanos');
    const huerfanos = [];
    for (let i = 1; i < allDataCot.length; i++) {
      const row = allDataCot[i];
      const idCliente = String(row[ID_CLIENTE_COL_COT] || '').trim();
      
      if (idCliente && !rucsValidos.has(idCliente)) {
        huerfanos.push({
          fila: i + 1,
          cotizacion: row[COT_COL_COT],
          idClienteHuerfano: idCliente
        });
      }
    }
    console.timeEnd('Búsqueda de Huérfanos');

    // 6. Reportar Resultados
    Logger.log("--- Reporte de Datos Huérfanos ---");
    Logger.log(`Total de cotizaciones analizadas: ${allDataCot.length - 1}`);
    Logger.log(`¡RIESGO ENCONTRADO! Total de cotizaciones huérfanas: ${huerfanos.length}`);
    
    if (huerfanos.length > 0) {
      Logger.log("--- MUESTRA DE DATOS HUÉRFANOS (primeros 20) ---");
      huerfanos.slice(0, 20).forEach(h => {
        Logger.log(`Fila ${h.fila} (COT: ${h.cotizacion}) tiene el ID CLIENTE '${h.idClienteHuerfano}', que no existe en la hoja '${HOJA_CLIENTES}'.`);
      });
    }
    Logger.log("--- Auditoría de Datos Huérfanos Finalizada ---");
    
    return {
      totalHuerfanos: huerfanos.length,
      muestra: huerfanos.slice(0, 20)
    };

  } catch (e) {
    Logger.log(`❌ Error en auditoría de datos huérfanos: ${e.message}\n${e.stack}`);
    return null;
  }
}

/**
 * 3. AUDITORÍA DE CORRUPCIÓN POR TEXTO LIBRE (Riesgo #3)
 * Busca variaciones en la columna 'DESCRIPCION' de DataCot.
 */
function auditarCorrupcionTextoLibre() {
  Logger.log("--- Iniciando Auditoría de Corrupción de Texto ---");
  try {
    // 1. Obtener mapa y datos
    const mapCot = getColumnMap(HOJA_COTIZACIONES);
    const allDataCot = obtenerDatosHoja(HOJA_COTIZACIONES, false);

    // 2. Validar encabezado (tu CSV usa 'DESCRIPCION')
    const DESC_COL_COT = mapCot['DESCRIPCION'];
    if (DESC_COL_COT === undefined) throw new Error(`No se encontró la columna 'DESCRIPCION' en ${HOJA_COTIZACIONES}`);

    // 3. Contar frecuencias
    console.time('Conteo de Frecuencias');
    const frecuencias = {};
    for (let i = 1; i < allDataCot.length; i++) {
      const descripcion = allDataCot[i][DESC_COL_COT];
      if (descripcion && typeof descripcion === 'string') {
        
        // Normalizar: minúsculas, sin acentos, sin espacios extra
        const normalizada = descripcion
          .toLowerCase()
          .trim()
          .replace(/\s+/g, ' ')
          .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
        
        if (normalizada) {
          if (!frecuencias[normalizada]) {
            frecuencias[normalizada] = { count: 0, originales: new Set() };
          }
          frecuencias[normalizada].count++;
          frecuencias[normalizada].originales.add(descripcion); // Guardar la versión original
        }
      }
    }
    console.timeEnd('Conteo de Frecuencias');

    // 4. Ordenar y filtrar
    const ordenado = Object.entries(frecuencias)
      .map(([normalizada, data]) => ({
        normalizada,
        count: data.count,
        originales: Array.from(data.originales)
      }))
      .sort((a, b) => b.count - a.count);

    // 5. Reportar Resultados
    Logger.log("--- Reporte de Corrupción de Texto ---");
    Logger.log(`Total de descripciones únicas (normalizadas): ${ordenado.length}`);
    
    const variantesCriticas = ordenado.filter(item => item.originales.length > 1);
    Logger.log(`¡RIESGO ENCONTRADO! Total de servicios con múltiples variantes de texto: ${variantesCriticas.length}`);
    
    if (variantesCriticas.length > 0) {
      Logger.log("--- MUESTRA DE VARIANTES CRÍTICAS (primeros 10) ---");
      variantesCriticas.slice(0, 10).forEach(v => {
        Logger.log(`Texto Normalizado: "${v.normalizada}" (Aparece ${v.count} veces)`);
        v.originales.forEach(original => {
          Logger.log(`  -> Variante Original: "${original}"`);
        });
      });
    }
    Logger.log("--- Auditoría de Corrupción de Texto Finalizada ---");
    
    return {
      totalVariantes: ordenado.length,
      variantesCriticas: variantesCriticas.length,
      muestra: variantesCriticas.slice(0, 10)
    };

  } catch (e) {
    Logger.log(`❌ Error en auditoría de corrupción de texto: ${e.message}\n${e.stack}`);
    return null;
  }
}
