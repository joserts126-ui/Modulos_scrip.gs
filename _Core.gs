// ====================================================
// === CORE: MANEJO DE VISTAS (doGet) Y UTILIDADES ===
// ====================================================

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
 * REFACTORIZADO (v3 - Supabase)
 * Obtiene el objeto de permisos para el usuario actual desde la tabla "Permisos".
 * Ya no usa caché para simplificar.
 */
function obtenerPermisosUsuario() {
  const email = obtenerEmailSeguro(); //
  Logger.log(`Obteniendo permisos para: ${email}`);
  
  // 1. Definir los permisos por defecto (Rol Invitado)
  let permisosUsuario = {
    email: email,
    rol: 'Invitado', //
    puedeEditarServicios: false, //
    puedeEditarOT: false, //
    puedeVerReportes: false, //
    puedeEditarCotizacion: false //
  };

  try {
    // 2. Buscar al usuario en la tabla "Permisos" de Supabase
    // (Usamos el "traductor" de _Supabase_Client.gs)
    const resultado = supabaseFetch('Permisos', {
      method: 'get',
      // Pide todas las columnas, filtrando por el email
      params: `select=*&Email=eq.${encodeURIComponent(email)}`
    });

    // 3. Si se encuentra al usuario, mapear los permisos
    if (resultado && resultado.length > 0) {
      const filaUsuario = resultado[0];
      Logger.log("Usuario encontrado en tabla Permisos.");
      
      // Mapeamos los nombres de tu tabla Supabase
      permisosUsuario.rol = filaUsuario.Rol || 'Invitado';
      permisosUsuario.puedeEditarServicios = filaUsuario.PuedeEditarServicios === true;
      permisosUsuario.puedeEditarOT = filaUsuario.PuedeEditarOT === true;
      permisosUsuario.puedeVerReportes = filaUsuario.PuedeVerReportes === true;
      permisosUsuario.puedeEditarCotizacion = filaUsuario.PuedeEditarCotizacion === true;
    
    } else {
      Logger.log("Usuario no encontrado en Permisos. Usando rol 'Invitado'.");
    }

  } catch (e) {
    Logger.log(`ERROR al buscar permisos en Supabase: ${e.message}. Usando rol 'Invitado'.`);
    // Si hay un error de BD, se queda con los permisos de 'Invitado' por seguridad
  }
  
  // 4. Devolver el objeto de permisos
  Logger.log(`Permisos finales: ${JSON.stringify(permisosUsuario)}`);
  return permisosUsuario;
}
function manejarError(contexto, error) {
    Logger.log(`❌ ERROR en ${contexto}: ${error.message}`);
    return {
        success: false,
        message: `Ocurrió un error en ${contexto}.`,
        error: error.message
    };
}
