/**
 * ====================================================
 * === CLIENTE DE SUPABASE (Traductor de API) ===
 * ====================================================
 * Este archivo centraliza la conexión con Supabase.
 */

// 1. OBTÉN ESTOS DATOS DE TU CONFIGURACIÓN DE SUPABASE
// (Guárdalos en Archivo > Propiedades > Propiedades del secuencias de comandos)
const SCRIPT_PROPS = PropertiesService.getScriptProperties();
const SUPABASE_URL = SCRIPT_PROPS.getProperty('SUPABASE_URL'); // Tu URL: https://lexwph...
const SUPABASE_KEY = SCRIPT_PROPS.getProperty('SUPABASE_KEY'); // Tu SERVICE_ROLE key

/**
 * Función central para TODAS las interacciones con Supabase.
 *
 * @param {string} tabla - El nombre de tu tabla en Supabase (ej. "Clientes").
 * @param {object} opciones - Configuración de la llamada.
 * @param {string} [opciones.method] - 'get', 'post', 'patch', 'delete'.
 * @param {string} [opciones.params] - Filtros de URL (ej. "select=*" o "RUC_DNI=eq.123").
 * @param {object|array} [opciones.payload] - El objeto (o array de objetos) JSON para enviar.
 * @returns {object|array} El resultado JSON de la API.
 */
function supabaseFetch(tabla, opciones = {}) {
  const { method = 'get', params = 'select=*', payload = null } = opciones;
  
  // Construir la URL del endpoint
  let url = `${SUPABASE_URL}/rest/v1/${tabla}`;
  if (params) {
    url += `?${params}`;
  }

  // Configurar las opciones de UrlFetchApp
  const fetchOptions = {
    'method': method,
    'headers': {
      'apikey': SUPABASE_KEY,
      'Authorization': `Bearer ${SUPABASE_KEY}`,
      'Prefer': 'return=representation' // Pide que Supabase devuelva el objeto(s)
    },
    'muteHttpExceptions': true // ¡Importante! Para capturar errores
  };

  if (payload) {
    fetchOptions.contentType = 'application/json';
    fetchOptions.payload = JSON.stringify(payload);
  }

  // Ejecutar la llamada
  const response = UrlFetchApp.fetch(url, fetchOptions);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();

  // Manejo de Errores
  if (responseCode >= 200 && responseCode < 300) {
    if (responseBody) {
      return JSON.parse(responseBody);
    }
    return true; // Éxito sin contenido (ej. en un DELETE)
  } else {
    Logger.log(`ERROR ${responseCode} en supabaseFetch (${tabla}): ${responseBody}`);
    // Lanzar un error claro que el frontend pueda entender
    let errorMsg = responseBody;
    try {
      errorMsg = JSON.parse(responseBody).message;
    } catch(e) {
      // No era JSON, usa el texto plano
    }
    throw new Error(`Error de Supabase: ${errorMsg}`);
  }
}
