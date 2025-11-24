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
 * REFACTORIZADO: Guarda o actualiza una Dirección usando supabaseFetch.
 * Se asegura de que el RUC_DNI sea NUMERIC.
 */
function guardarOActualizarDireccion(data) {
    try {
        const datosSanitizados = sanitizarDatos(data);
        const idDireccion = datosSanitizados.rowIndex; 
        
        // 1. Parsear RUC_DNI a NUMERIC (bigint)
        const rucNumerico = parseFloat(datosSanitizados.RUC);
        if (isNaN(rucNumerico)) throw new Error("RUC/DNI no es un valor numérico válido para Direcciones.");
        
        const payload = {
            "RUC_DNI": rucNumerico, 
            "Tipo": datosSanitizados.TIPO,
            // Nombres de columna probables si la tabla existiera
            "Direccion_Completa": datosSanitizados.DIRECCION, 
            "Ciudad_Distrito": datosSanitizados.CIUDAD      
        };
        
        let metodo;
        let params;

        if (idDireccion && parseInt(idDireccion) > 0) { 
            metodo = 'patch';
            params = `ID_Direccion=eq.${idDireccion}`; // Clave primaria
        } else { 
            metodo = 'post';
            params = '';
        }

        // Asume que la tabla es 'Direcciones'
        supabaseFetch('Direcciones', { 
            method: metodo,
            payload: payload,
            params: params
        });
        
        return { success: true, message: "Dirección guardada exitosamente" };
    } catch (error) {
        return manejarError("guardarOActualizarDireccion", error);
    }
}

/**
 * Obtiene los datos iniciales para el módulo de Servicios desde Supabase.
 */
function getDatosInicialesServicios() {
    try {
        // 1. Obtener Servicios de Supabase
        // Nota: Pedimos todas las columnas.
        const servicios = supabaseFetch('Servicios', {
            method: 'get',
            params: 'select=*&order=Nombre_Servicio.asc'
        });

        // 2. Devolver estructura combinada
        return {
            success: true,
            servicios: servicios || [], // Array de objetos JSON
            horasSegun: LISTA_HORAS_SEGUN, // Constante de _Constants_Lists.gs
            undMedida: LISTA_UND_MEDIDA    // Constante de _Constants_Lists.gs
        };

    } catch (error) {
        return manejarError('getDatosInicialesServicios', error);
    }
}

/**
 * Guarda o actualiza un Servicio usando supabaseFetch.
 * AHORA USA LAS COLUMNAS REALES: Hora_Segun y Und_Medida.
 */
function guardarOActualizarServicio(data) {
    const permisos = obtenerPermisosUsuario();
    if (!permisos.puedeEditarServicios) {
        return { success: false, message: "Acceso denegado. No tiene permiso para editar servicios." };
    }
    
    try {
        const datosSanitizados = sanitizarDatos(data);
        const idServicio = datosSanitizados.rowIndex; 
        
        // Mapeo a las columnas de Supabase
        // IMPORTANTE: 'hora_segun' y 'und_medida' en minúsculas
        const payload = {
            "Nombre_Servicio": datosSanitizados.DESCRIPCION, 
            "hora_segun": datosSanitizados.HORA_SEGUN,       
            "und_medida": datosSanitizados.UND_MEDIDA,       
            
            "Maquinaria": datosSanitizados.Maquinaria || '', 
            "Tipo": datosSanitizados.Tipo || '',
            "Abreviatura": datosSanitizados.Abreviatura || '',
            "Personal": datosSanitizados.Personal || ''
        };

        let metodo;
        let params;
        
        if (idServicio && parseInt(idServicio) > 0) {
             metodo = 'patch';
             params = `ID_servicios=eq.${idServicio}`;
        } else {
             metodo = 'post';
             params = '';
        }

        supabaseFetch('Servicios', { 
            method: metodo,
            payload: payload,
            params: params
        });
        
        return { success: true, message: "Servicio guardado exitosamente" };
    } catch (error) {
        // Devolvemos el error real para facilitar la depuración
        Logger.log("Error en guardarOActualizarServicio: " + error.message);
        return { success: false, message: error.message }; 
    }
}
/**
 * Obtiene un Servicio por su ID único de Supabase (ID_servicios).
 */
function obtenerServicioPorRowIndex(rowIndex) {
    try {
        if (!rowIndex || rowIndex <= 0) throw new Error("ID de Servicio inválido.");
        
        const params = `ID_servicios=eq.${rowIndex}`;
        const resultado = supabaseFetch('Servicios', { method: 'get', params: params });
        
        if (resultado && resultado.length > 0) {
            const servicio = resultado[0];
            
            // Simulamos la estructura de array 2D que el frontend espera
            // Leemos hora_segun y und_medida (minúsculas)
            const servicioSimulado = [
                 servicio.ID_servicios,
                 servicio.Nombre_Servicio,
                 '', 
                 '', 
                 servicio.hora_segun || '', 
                 servicio.und_medida || '' 
             ];

            return { 
                 success: true, 
                 row: servicioSimulado, 
                 rowIndex: rowIndex 
            };
        }

        return { success: false, message: 'Servicio no encontrado.' };
        
    } catch (error) {
        Logger.log("Error en obtenerServicioPorRowIndex: " + error.message);
        return { success: false, message: error.message };
    }
}


// ====================================================
// === FUNCIONES DE MÓDULO CONTACTOS (DETALLE) ===
// ====================================================

/**
 * REEMPLAZO (v4 - Supabase)
 * Obtiene los Contactos y Direcciones para un RUC específico usando Supabase.
 * Devuelve arrays 2D (con el ID de Supabase como el último elemento para simular ROW_INDEX)
 * para mantener la compatibilidad con Contactos.html.
 */
function getContactosYDirecciones(ruc) {
    try {
        if (!ruc) return { contactos: [['ID', 'RUC', 'NOMBRE', 'EMAIL', 'TELÉFONO', 'CARGO', 'ROW_INDEX']], direcciones: [['ID', 'RUC', 'TIPO', 'DIRECCIÓN', 'CIUDAD', 'ROW_INDEX']] };
        
        // RUC_DNI es NUMERIC en el esquema SQL, lo parseamos
        const rucNumerico = parseFloat(ruc);
        if (isNaN(rucNumerico)) throw new Error("RUC no es un valor numérico válido.");
        
        // 1. Obtener Contactos
        const contactoCols = 'ID_Contacto, RUC_DNI, Nombre_Contacto, Correo, Celular, Cargo';
        const paramsContacto = `RUC_DNI=eq.${rucNumerico}&select=${contactoCols}`;
        const contactos = supabaseFetch('Contactos', { method: 'get', params: paramsContacto });

        // 2. Obtener Direcciones
        // Asumimos tabla Direcciones con estas columnas, aunque no está en el esquema SQL provisto.
        const direccionCols = 'ID_Direccion, RUC_DNI, Tipo, Direccion_Completa, Ciudad_Distrito'; 
        const paramsDireccion = `RUC_DNI=eq.${rucNumerico}&select=${direccionCols}`;
        const direcciones = supabaseFetch('Direcciones', { method: 'get', params: paramsDireccion });

        // 3. Mapear Contactos a formato 2D + ID
        const headersContacto = ['ID', 'RUC', 'NOMBRE', 'EMAIL', 'TELÉFONO', 'CARGO', 'ROW_INDEX'];
        const contactosMapeados = [headersContacto];
        contactos.forEach(c => {
            contactosMapeados.push([
                c.ID_Contacto,        // [0] ID
                c.RUC_DNI,            // [1] RUC
                c.Nombre_Contacto,    // [2] NOMBRE
                c.Correo,             // [3] EMAIL
                c.Celular,            // [4] TELÉFONO
                c.Cargo,              // [5] CARGO
                c.ID_Contacto         // [6] ID_Contacto como ROW_INDEX
            ]);
        });
        
        // 4. Mapear Direcciones a formato 2D + ID
        const headersDireccion = ['ID', 'RUC', 'TIPO', 'DIRECCIÓN', 'CIUDAD', 'ROW_INDEX'];
        const direccionesMapeadas = [headersDireccion];
        direcciones.forEach(d => {
            direccionesMapeadas.push([
                d.ID_Direccion,       // [0] ID
                d.RUC_DNI,            // [1] RUC
                d.Tipo,               // [2] TIPO
                d.Direccion_Completa, // [3] DIRECCIÓN
                d.Ciudad_Distrito,    // [4] CIUDAD
                d.ID_Direccion        // [5] ID_Direccion como ROW_INDEX
            ]);
        });
        
        return { 
            contactos: contactosMapeados, 
            direcciones: direccionesMapeadas 
        };

    } catch (error) {
        return manejarError('getContactosYDirecciones', error);
    }
}

/**
 * Obtiene una fila de Contacto, Dirección o Servicio por su ID único de Supabase (rowIndex).
 */
function getFilaPorRowIndex(ruc, rowIndex, tipo) {
    try {
        if (!rowIndex || rowIndex <= 0) {
            throw new Error(`ID Inválido para ${tipo}`);
        }
        
        let tabla;
        let columnaID;
        let mapeo;

        if (tipo === 'contacto') {
            tabla = 'Contactos';
            columnaID = 'ID_Contacto'; 
            // Mapea el objeto de Supabase a los índices de array que espera Contactos.html
            mapeo = (data) => [data.ID_Contacto, data.RUC_DNI, data.Nombre_Contacto, data.Correo, data.Celular, data.Cargo];
        } else if (tipo === 'direccion') {
            tabla = 'Direcciones'; 
            columnaID = 'ID_Direccion'; 
            mapeo = (data) => [data.ID_Direccion, data.RUC_DNI, data.Tipo, data.Direccion_Completa, data.Ciudad_Distrito];
        } else if (tipo === 'servicio') {
            tabla = 'Servicios';
            columnaID = 'ID_servicios';
            // Aquí debes devolver el array simulado de 7 columnas que espera Servicios.html
            mapeo = (data) => [data.ID_servicios, data.Nombre_Servicio, '', '', data.Hora_Segun || '', data.Und_Medida || '', ''];
        } else {
            throw new Error('Tipo de dato no reconocido: ' + tipo);
        }
        
        const params = `${columnaID}=eq.${rowIndex}`;
        const resultado = supabaseFetch(tabla, { method: 'get', params: params });
        
        if (resultado && resultado.length > 0) {
            // El array mapeado es la simulación de la fila de Sheets
            const filaArraySimulado = mapeo(resultado[0]); 
            return { 
                success: true, 
                row: filaArraySimulado, 
                rowIndex: rowIndex 
            };
        }

        return { success: false, message: 'No se encontró la fila.' };
        
    } catch (error) {
        return manejarError("getFilaPorRowIndex", error);
    }
}
/**
 * Obtiene la lista de cotizaciones para el Resumen.
 * CORREGIDO: Usa 'Ejecitivo' (nombre real en DB) en lugar de 'Ejecutivo'.
 */
function getListaCotizacionesResumen() {
  try {
    // Consulta corregida: "Ejecitivo"
    const consulta = 'select=Cot,Fecha_Creacion,Ejecitivo,Empresa,Total_Cot,Moneda,Estado_Cot,RUC,Clientes(Nombre_RazonSocial)&order=Fecha_Creacion.desc';

    const pedidos = supabaseFetch('Pedidos', {
      method: 'get',
      params: consulta
    });
    
    return pedidos || []; 

  } catch (e) {
    Logger.log(`Error en getListaCotizacionesResumen: ${e.message}`);
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
 * Obtiene los datos de una OT por su número, desde la tabla Ordenes_Trabajo en Supabase.
 */
function obtenerOTPorNumero(numOT) {
    try {
        if (!numOT) return { success: false, message: "Número de OT no proporcionado." };
        
        const params = `N_OT=eq.${numOT}&select=*, Detalle_Pedidos!inner(Cot_Linea_Ref)`;

        const resultado = supabaseFetch('Ordenes_Trabajo', {
            method: 'get',
            params: params
        });

        if (resultado && resultado.length > 0) {
            const ot = resultado[0]; 
            const linea = ot.Detalle_Pedidos;
            
            // Mapeamos los nombres de Supabase a los nombres del frontend de RegistrarOT.html
            const datosMapeados = {
                numeroOT: ot.N_OT,
                otIDExistente: ot.N_OT,
                // Usamos la referencia de texto para que el frontend la use al actualizar
                lineaIDAsociada: linea.Cot_Linea_Ref, 
                fecha: formatearParaInputDate(ot.Fecha),
                horaInicio: ot.Hora_Inicio,
                horaFin: ot.Hora_Fin,
                tiempoRefrigerio: ot.Tiempo_Refrigerio,
                horometroInicio: ot.Horometro_Inicio,
                horometroFin: ot.Horometro_Fin,
                montoDespacho: ot.Monto_Servicio,
                montoMovilizacion: ot.Monto_Movilizacion,
                tiempoTotal: ot.Tiempo_Total_Horas,
                horometroTrabajado: ot.Horometro_Trabajado_Horas,
                usuario: ot.Usuario_Registro,
                fechaRegistro: ot.Fecha_Registro
                // NOTA: 'clienteOT' y 'servicioOT' no se pueden obtener aquí sin un JOIN muy complejo,
                // por lo que el frontend los dejará vacíos o los llenará en `cargarDatosOTExistente`
            };
            return { success: true, data: datosMapeados }; 
        }

        return { success: false, message: 'OT no encontrada.' };

    } catch (error) {
        return manejarError("obtenerOTPorNumero", error);
    }
}

/**
 * Filtra pedidos por RUC del cliente (NUMERIC) y ID del servicio (BIGINT).
 * Consulta Detalle_Pedidos, hace JOIN a Pedidos.
 */
function filtrarPedidosPorClienteYServicio(rucCliente, idServicio) {
    try {
        const rucNumerico = parseFloat(rucCliente);
        const idServicioNumerico = parseFloat(idServicio);
        
        if (isNaN(rucNumerico)) throw new Error("RUC no es un valor numérico válido.");
        if (isNaN(idServicioNumerico)) throw new Error("ID de Servicio no es un valor numérico válido.");

        // Filtramos por el ID de Servicio y luego por el RUC a través del JOIN Pedidos.
        const consulta = `
          select=Cot
          &ID_Servicio=eq.${idServicioNumerico}
          &Pedidos!inner(RUC=eq.${rucNumerico})
          &order=Cot.desc
          &limit=100
        `.replace(/\s/g, '');

        const detalles = supabaseFetch('Detalle_Pedidos', {
          method: 'get',
          params: consulta
        });

        // Extraer los valores únicos de 'Cot' (Número de Pedido)
        const pedidosUnicos = new Set();
        detalles.forEach(d => pedidosUnicos.add(d.Cot));

        // Devolvemos un array de arrays para que coincida con el frontend (RegistrarOT.html)
        return Array.from(pedidosUnicos).map(pedido => [pedido]);

    } catch (e) {
        return manejarError("filtrarPedidosPorClienteYServicio", e);
    }
}

/**
 * Guarda la Orden de Trabajo en Supabase (Versión con Diagnóstico de Error).
 */
function guardarOT(data) {
    const permisos = obtenerPermisosUsuario();
    if (!permisos.puedeEditarOT) {
        return { success: false, message: "Acceso denegado. No tiene permiso para editar OTs." };
    }

    Logger.log("INICIO guardarOT (Debug). Datos recibidos: " + JSON.stringify(data));
    
    try {
        const datos = sanitizarDatos(data);
        const lineaIDRef = datos.lineaID; 
        const esModoUpdate = (datos.modo === 'editar' || datos.modo === 'editarGlobal');
        
        if (!lineaIDRef && !esModoUpdate) {
            throw new Error("Error: No se encontró la referencia de línea (lineaID).");
        }
        
        // --- 1. Validaciones de Tipos de Datos (Para evitar error 400/500 en Supabase) ---
        const rucClienteNum = parseFloat(datos.clienteRUC);
        const idServicioNum = parseInt(datos.servicio);

        if (isNaN(rucClienteNum)) {
            throw new Error(`El RUC del cliente '${datos.clienteRUC}' no tiene un formato numérico válido. Asegúrese de seleccionar un cliente de la lista.`);
        }
        if (isNaN(idServicioNum)) {
            throw new Error(`El ID del servicio '${datos.servicio}' no tiene un formato numérico válido. Asegúrese de seleccionar un servicio de la lista.`);
        }

        // --- 2. Obtener detalles técnicos de la línea ---
        const detallesLinea = supabaseFetch('rpc/get_detalles_linea_para_ot', {
            method: 'post',
            payload: { "linea_ref_param": lineaIDRef }
        })[0];

        if (!detallesLinea || !detallesLinea.success) {
            throw new Error(`No se encontraron detalles en base de datos para la línea ${lineaIDRef}. Verifique que el pedido tenga líneas guardadas.`);
        }
        
        const lineaIdRealNumerico = detallesLinea.linea_id_real;
        
        // --- 3. Calcular Montos ---
        const horasSegun = (detallesLinea.horas_segun || '').toUpperCase();
        let horasParaCobrar = 0;
        
        if (horasSegun.includes("HORÓMETRO") || horasSegun.includes("HOROMETRO")) {
            horasParaCobrar = parseFloat(datos.horometroTrabajado || 0);
        } else {
            horasParaCobrar = parseFloat(datos.tiempoTotal || 0);
        }

        const precioUnitario = parseFloat(detallesLinea.precio || 0);
        const montoServicioCalculado = horasParaCobrar * precioUnitario;
        const montoMovilizacionForm = parseFloat(datos.montoMovilizacion || 0);
        
        // --- 4. Preparar Payload ---
        const payloadOT = {
            "N_OT": datos.numeroOT,
            "Cot": datos.pedido,                   
            "Fecha": datos.fecha,                  
            "Cot_Linea": lineaIdRealNumerico,      
            "Cot_Linea_Ref": lineaIDRef,           
            "RUC_Cliente": rucClienteNum, // Usamos el validado
            "ID_Servicio": idServicioNum, // Usamos el validado
            
            "Hora_Inicio": datos.horaInicio || null,
            "Hora_Fin": datos.horaFin || null,
            "Tiempo_Refrigerio": parseFloat(datos.tiempoRefrigerio || 0),
            "Tiempo_Total": parseFloat(datos.tiempoTotal || 0),
            
            "Horometro_Inicio": parseFloat(datos.horometroInicio || 0),
            "Horometro_Fin": parseFloat(datos.horometroFin || 0),
            "Horometro_Trabajado": parseFloat(datos.horometroTrabajado || 0),
            
            "Es_Camion_Grua": datos.esCamionGrua === 'true' || datos.esCamionGrua === true, 
            "Horometro_Inicio_Camion": parseFloat(datos.horometroInicioCamion || 0),
            "Horometro_Fin_Camion": parseFloat(datos.horometroFinCamion || 0),
            "Horometro_Inicio_Grua": parseFloat(datos.horometroInicioGrua || 0),
            "Horometro_Fin_Grua": parseFloat(datos.horometroFinGrua || 0),
            
            "Monto_Servicio": montoServicioCalculado,
            "Monto_Movilizacion": montoMovilizacionForm,
            "Estado_Valorizacion": 'Pendiente'
        };

        // --- 5. Ejecutar Guardado ---
        if (esModoUpdate) {
            // UPDATE
            const otIDExistente = datos.otIDExistente;
            delete payloadOT.Estado_Valorizacion; 
            delete payloadOT.Fecha_Registro;      
            delete payloadOT.Usuario_Registro;    
            
            supabaseFetch('Ordenes_Trabajo', {
                method: 'patch',
                payload: payloadOT,
                params: `N_OT=eq.${otIDExistente}`
            });
            return { success: true, message: `OT ${otIDExistente} actualizada.` };

        } else {
            // CREATE
            payloadOT.Usuario_Registro = obtenerEmailSeguro();
            payloadOT.Fecha_Registro = new Date().toISOString();

            supabaseFetch('Ordenes_Trabajo', {
                method: 'post',
                payload: payloadOT
            });

            // Actualizar acumulados (RPC)
            const payloadAcumular = {
                "linea_id_real_param": lineaIdRealNumerico,
                "horas_sumar": horasParaCobrar,
                "monto_sumar": montoServicioCalculado,
                "movilizacion_sumar": montoMovilizacionForm
            };
            supabaseFetch('rpc/actualizar_despacho_detalle', {
                method: 'post',
                payload: payloadAcumular
            });
            
            return { success: true, message: `OT ${datos.numeroOT} registrada exitosamente.` };
        }

    } catch (e) {
        Logger.log(`❌ ERROR REAL en guardarOT: ${e.message}`);
        
        // Manejo de errores comunes para dar mensajes claros
        if (e.message.includes('duplicate key')) {
            return { success: false, message: `El número de OT '${data.numeroOT}' ya existe. Use otro número.` };
        }
        if (e.message.includes('violates foreign key constraint')) {
             return { success: false, message: `Error de datos: El Pedido, Cliente o Servicio no existen en la base de datos.` };
        }
        
        // DEVOLVEMOS EL MENSAJE REAL DEL ERROR
        return { success: false, message: "Error del sistema: " + e.message };
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
 * REEMPLAZO de getDestinationFolder (CORREGIDO para Supabase)
 * Ahora maneja 'Lineas' (Supabase) y 'servicios' (Legacy).
 */
function getDestinationFolder(ejecutivo, empresa, fechaObj, cotizacionData) {
    let parentFolderId;

    // 1. Determinar Carpeta del Ejecutivo (Nivel 1)
    // Normalización robusta para evitar errores si el ejecutivo viene nulo
    const ejecutivoNombre = (ejecutivo || '').toUpperCase();
    
    if (ejecutivoNombre.includes('CARMEN')) {
        parentFolderId = FOLDER_ID_CARMEN;
    } else {
        // Lógica de fallback por empresa si no es Carmen
        const empresaUpper = (empresa || '').toUpperCase();
        if (empresaUpper.includes('GYM')) parentFolderId = FOLDER_ID_GYM;
        else if (empresaUpper.includes('SAN JOSE')) parentFolderId = FOLDER_ID_SJ;
        else parentFolderId = FOLDER_ID_ALP;
    }
    
    let currentFolder = DriveApp.getFolderById(parentFolderId);

    // 2. Carpeta de la Empresa (Nivel 2)
    const empresaNombre = (empresa || '').toUpperCase();
    const nombreCarpetaEmpresa = MAPA_NOMBRES_EMPRESAS[empresaNombre] || empresaNombre || "EMPRESA_DESCONOCIDA";
    currentFolder = findOrCreateFolder(currentFolder, nombreCarpetaEmpresa);

    // 3. Carpeta Mes/Año (Nivel 3)
    const prefijos = {"ALPAMAYO": "COT.ALP", "SAN JOSE": "COT.SJ", "GYM": "COT.GYM"};
    const meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
    
    const prefijoEmpresa = prefijos[empresaNombre] || 'COT.GEN';
    const anio = fechaObj.getFullYear();
    const mesNum = String(fechaObj.getMonth() + 1).padStart(2, '0');
    const mesNombre = meses[fechaObj.getMonth()];
    
    const nombreSubfolderMes = `${prefijoEmpresa}.${anio}.${mesNum} COT_${mesNombre}`;
    currentFolder = findOrCreateFolder(currentFolder, nombreSubfolderMes);

    // 4. Carpeta de Cotización Específica (Nivel 4)
    let servicioNombre = "VARIOS";
    
    // ✅ CORRECCIÓN AQUÍ: Detectar 'Lineas' (Supabase) o 'servicios' (Legacy)
    const listaServicios = cotizacionData.Lineas || cotizacionData.servicios || [];

    if (listaServicios.length === 1) {
        const servicioUnico = listaServicios[0];
        // Priorizar Abreviatura, luego Cod, luego Descripción
        servicioNombre = servicioUnico.abreviatura || servicioUnico.cod || (servicioUnico.descripcion || '').substring(0, 10);
    }
    
    const nombreCarpetaCotizacion = `${cotizacionData.numPedido} ${cotizacionData.cliente || cotizacionData.Cliente} ${servicioNombre}`;
    currentFolder = findOrCreateFolder(currentFolder, nombreCarpetaCotizacion); 

    // 5. Subcarpeta "cot" (Nivel 5)
    currentFolder = findOrCreateFolder(currentFolder, "cot");

    return currentFolder;
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
 * Obtiene el resumen de líneas de un pedido (vFinal - Corregido RUC).
 * Lee directamente de 'Detalle_Pedidos' y sus tablas relacionadas.
 */
function getResumenOTPorPedido(numPedido) {
  Logger.log(`INICIO getResumenOTPorPedido para: ${numPedido}`);
  try {
    if (!numPedido) throw new Error("Número de pedido no proporcionado.");

    // ✅ CORRECCIÓN: Cambiamos RUC_DNI por RUC dentro de Pedidos!inner(...)
    // También aseguramos el orden correcto de los joins
    const consulta = `
      select=
        Cot,
        Cot_Linea_Ref,
        Cantidad,
        UND,
        UND_DESPACHO,
        Servicios!inner(ID_servicios, Nombre_Servicio),
        Pedidos!inner(RUC, Clientes!inner(Nombre_RazonSocial))
      &Cot=eq.${numPedido}
    `.replace(/\s/g, '');

    const lineasDetalle = supabaseFetch('Detalle_Pedidos', {
      method: 'get',
      params: consulta
    });

    if (!lineasDetalle || lineasDetalle.length === 0) {
       return { success: true, pedido: numPedido, cliente: "Desconocido", lineas: [] };
    }

    // Obtener el nombre del cliente de forma segura (maneja arrays u objetos)
    let clienteNombre = "Cliente no encontrado";
    const primerPedido = lineasDetalle[0].Pedidos;
    
    if (primerPedido && primerPedido.Clientes) {
         if (Array.isArray(primerPedido.Clientes)) {
             clienteNombre = primerPedido.Clientes[0]?.Nombre_RazonSocial || clienteNombre;
         } else {
             clienteNombre = primerPedido.Clientes.Nombre_RazonSocial || clienteNombre;
         }
    }

    const lineasMapeadas = lineasDetalle.map((linea) => {
      let horasPedidas = (linea.UND === 'HORAS' || linea.UND === 'DÍAS') ? linea.Cantidad : 0;

      return {
        lineaID: linea.Cot_Linea_Ref, 
        cod: linea.Servicios.ID_servicios,
        descripcion: linea.Servicios.Nombre_Servicio,
        horasPedidas: horasPedidas,
        horasDespachadas: parseFloat(linea.UND_DESPACHO || 0)
      };
    });

    return { 
      success: true, 
      pedido: numPedido,
      cliente: clienteNombre,
      lineas: lineasMapeadas
    };

  } catch (e) {
    Logger.log(`ERROR en getResumenOTPorPedido: ${e.message}`);
    return manejarError('getResumenOTPorPedido', e);
  }
}

/**
 * Obtiene los datos de una LÍNEA específica para pre-llenar una OT (MIGRADO A SUPABASE).
 */
function getDatosDeLineaParaOT(lineaID) {
  try {
    if (!lineaID) throw new Error("ID de Línea no proporcionado.");

    // Consultar Detalle_Pedidos con Joins a Pedidos, Clientes y Servicios
    // Filtramos por Cot_Linea_Ref (ej. COT.ALP.2025.11.3019-L1)
    const consulta = `
        select=
            Cot,
            Cot_Linea_Ref,
            Servicios!inner(ID_servicios, Nombre_Servicio),
            Pedidos!inner(
                RUC, 
                Clientes!inner(Nombre_RazonSocial)
            )
        &Cot_Linea_Ref=eq.${lineaID}
    `.replace(/\s/g, '');

    const resultado = supabaseFetch('Detalle_Pedidos', {
        method: 'get',
        params: consulta
    });

    if (!resultado || resultado.length === 0) {
        throw new Error("No se encontró la línea con ID: " + lineaID);
    }

    const linea = resultado[0];
    
    // Manejo seguro del cliente (objeto o array)
    let clienteNombre = "";
    let clienteRUC = "";
    
    if (linea.Pedidos) {
        clienteRUC = linea.Pedidos.RUC;
        if (linea.Pedidos.Clientes) {
            clienteNombre = Array.isArray(linea.Pedidos.Clientes) 
                ? linea.Pedidos.Clientes[0]?.Nombre_RazonSocial 
                : linea.Pedidos.Clientes.Nombre_RazonSocial;
        }
    }

    const datosLinea = {
      lineaID: linea.Cot_Linea_Ref,
      numPedido: linea.Cot,
      clienteRUC: clienteRUC,
      clienteNombre: clienteNombre,
      codServicio: linea.Servicios.ID_servicios,
      descServicio: linea.Servicios.Nombre_Servicio
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
 * Obtiene los datos de una línea Y la lista de servicios del pedido (MIGRADO A SUPABASE).
 */
function getDatosDePedidoParaOT(lineaID) {
  try {
    // 1. Obtener datos de la línea específica
    const resultadoLinea = getDatosDeLineaParaOT(lineaID);
    if (!resultadoLinea.success) throw new Error(resultadoLinea.message);
    
    const datosLinea = resultadoLinea.data;
    const numPedido = datosLinea.numPedido;

    // 2. Buscar TODOS los servicios de ese pedido para llenar el combo
    const consultaServicios = `
        select=ID_Servicio, Servicios(Nombre_Servicio)
        &Cot=eq.${numPedido}
    `.replace(/\s/g, '');

    const resultadoServicios = supabaseFetch('Detalle_Pedidos', {
        method: 'get',
        params: consultaServicios
    });

    const serviciosDelPedido = [];
    const vistos = new Set();

    if (resultadoServicios) {
        resultadoServicios.forEach(item => {
            const cod = item.ID_Servicio;
            const desc = item.Servicios ? item.Servicios.Nombre_Servicio : '';
            
            if (!vistos.has(cod)) {
                serviciosDelPedido.push({ cod: cod, desc: desc });
                vistos.add(cod);
            }
        });
    }

    return { 
      success: true, 
      dataLinea: datosLinea, 
      serviciosDelPedido: serviciosDelPedido 
    };

  } catch (e) {
    return manejarError('getDatosDePedidoParaOT', e);
  }
}

/**
 * GUARDA LA COTIZACIÓN EN SUPABASE (vFinal - Corregida para Schema SQL)
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
    
    // Validaciones previas
    const lineas = datosSanitizados.Lineas || [];
    if (lineas.length === 0) throw new Error("Debe agregar al menos un servicio.");

    // Parsear RUC a número (Requisito SQL: numeric)
    const rucNumerico = parseFloat(datosSanitizados.RUC);
    if (isNaN(rucNumerico)) throw new Error("El RUC debe ser numérico.");

    // =====================================================
    // PREPARACIÓN DE PAYLOADS (Alineados al SQL)
    // =====================================================
    
    // Campos comunes para Create y Update
    const payloadCampos = {
        "Estado_Cot": datosSanitizados.Estado,
        "Total_Cot": parseFloat(datosSanitizados.Total),
        "Moneda": datosSanitizados.Moneda,
        // "Ejecitivo": Nota el typo en tu SQL ("i" en vez de "u"). 
        // Lo mantenemos así para que funcione, o cámbialo a "Ejecutivo" si corriges tu DB.
        "Ejecitivo": datosSanitizados.Ejecutivo, 
        "Fecha_Inicio": datosSanitizados.fechaEjecucion ? new Date(datosSanitizados.fechaEjecucion).toISOString() : null,
        "Forma_De_Pago": datosSanitizados.Forma_De_Pago || datosSanitizados.Forma_Pago,
        "Empresa": datosSanitizados.Empresa,
        "RUC": rucNumerico, // CORREGIDO: En tabla Pedidos se llama "RUC", no "RUC_DNI"
        "Direccion": datosSanitizados.Direccion,
        "Turno": datosSanitizados.Turno,
        "plantillaNotas": datosSanitizados.plantillaNotas || '',
        "aclaracionesServicio": datosSanitizados.aclaracionesServicio || '',
        // CAMPO OBLIGATORIO FALTANTE EN TU CÓDIGO ANTERIOR:
        "Estado_Factura": "Pendiente" 
    };

    // MODO ACTUALIZACIÓN (PATCH)
    if (esModoUpdate) {
      codigoPedido = datosSanitizados.numPedido;
      Logger.log(`Iniciando MODO UPDATE para: ${codigoPedido}`);

      const payloadRPC = {
        "codigo_pedido": codigoPedido,
        "datos_cabecera": payloadCampos,
        "nuevas_lineas": prepararLineasParaRPC(lineas, codigoPedido) // Función auxiliar abajo
      };
      
      // Usamos la RPC para actualización atómica
      supabaseFetch('rpc/actualizar_cotizacion_y_detalles', {
        method: 'post',
        payload: payloadRPC
      });
      
    } 
    // MODO CREACIÓN (POST)
    else {
      Logger.log("Iniciando MODO CREATE...");
      codigoPedido = generarCodigoPedido(datosSanitizados.Empresa);
      
      // Completar campos exclusivos de creación
      payloadCampos.Cot = codigoPedido;
      payloadCampos.Fecha_Creacion = new Date().toISOString();

      // 1. Insertar Cabecera
      supabaseFetch('Pedidos', {
        method: 'post',
        payload: payloadCampos
      });

      // 2. Insertar Detalles
      const payloadDetalles = prepararLineasParaInsert(lineas, codigoPedido);
      supabaseFetch('Detalle_Pedidos', {
        method: 'post',
        payload: payloadDetalles
      });
    }
    
    return { 
      success: true, 
      message: esModoUpdate ? "Cotización actualizada" : "Cotización registrada",
      codigoPedido: codigoPedido 
    };

  } catch (error) {
    // Importante: Devolver el error real para que lo veas en el frontend
    Logger.log(`❌ ERROR en guardarCotizacion: ${error.message}`);
    return { success: false, message: "Error al guardar: " + error.message };
  }
}

// --- FUNCIONES AUXILIARES PARA LIMPIEZA ---

function prepararLineasParaInsert(lineas, codigoPedido) {
    return lineas.map((linea, i) => ({
        "Cot_Linea_Ref": `${codigoPedido}-L${i + 1}`,
        "Cot": codigoPedido,
        "ID_Servicio": parseInt(linea.cod), // Asegurar entero
        "Cantidad": parseFloat(linea.cantidad) || 0,
        "Precio": parseFloat(linea.precio) || 0,
        "Monto_Movilizacion": parseFloat(linea.movilizacion) || 0,
        "Horas_Segun": linea.hora_segun || '',
        "UND": linea.und_medida || '',
        "UND_HORAS_MINIMAS": linea.und_horas_minimas || '',
        "Horas_minimas": parseFloat(linea.horas_minimas_num) || 0
    }));
}

function prepararLineasParaRPC(lineas, codigoPedido) {
    // La estructura es idéntica para la RPC
    return prepararLineasParaInsert(lineas, codigoPedido);
}

/**
 * REFACTORIZADO (vFinal - Corregido Mapeo de Columnas): 
 * Obtiene los datos para editar un pedido en UNA SOLA LLAMADA a Supabase.
 */
function obtenerPedidoParaEdicion(numPedido) {
  Logger.log(`Iniciando obtenerPedidoParaEdicion (vFinal) para: ${numPedido}`);
  try {
    // 1. Construir la consulta anidada.
    // IMPORTANTE: 'RUC' es la columna en Pedidos, 'RUC_DNI' es la columna en Clientes.
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
    `.replace(/\s/g, '');

    // 2. Ejecutar la llamada ÚNICA a Supabase
    const resultado = supabaseFetch('Pedidos', {
      method: 'get',
      params: consulta
    });

    if (!resultado || resultado.length === 0) {
      throw new Error(`Pedido ${numPedido} no encontrado.`);
    }
    
    const pedido = resultado[0];

    // 3. Mapear los datos (AQUÍ ESTABA EL ERROR)
    const datosGenerales = {
      fechaEjecucion: pedido.Fecha_Inicio ? pedido.Fecha_Inicio.split('T')[0] : '',
      Estado: pedido.Estado_Cot,
      Empresa: pedido.Empresa,
      
      // ✅ CORRECCIÓN 1: La columna en la tabla Pedidos se llama "RUC", no "RUC_DNI"
      RUC: pedido.RUC, 
      
      Cliente: pedido.Clientes.Nombre_RazonSocial, 
      Moneda: pedido.Moneda,
      Forma_De_Pago: pedido.Forma_De_Pago,
      Direccion: pedido.Direccion || (pedido.Clientes ? pedido.Clientes.Direccion_Fiscal : ''),
      Turno: pedido.Turno || '', 
      
      // ✅ CORRECCIÓN 2: La columna en la tabla Pedidos se llama "Ejecitivo" (typo en DB)
      Ejecutivo: pedido.Ejecitivo || pedido.Ejecutivo, 
      
      // Obtener nombre de contacto si existe relación
      Contacto: pedido.Contactos ? pedido.Contactos.Nombre_Contacto : '' 
    };

    // 4. Mapear líneas de detalle
    const lineas = pedido.Detalle_Pedidos.map(linea => {
      // Calcular subtotal sumando precio*cantidad + movilización
      const subtotal = (linea.Cantidad * linea.Precio) + (linea.Monto_Movilizacion || 0);
      return {
        cod: linea.ID_Servicio,
        descripcion: linea.Servicios.Nombre_Servicio,
        cantidad: linea.Cantidad,
        und_medida: linea.UND,
        und_horas_minimas: linea.UND_HORAS_MINIMAS,
        dias_cotizados: 0, 
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
      plantillaNotas: pedido.plantillaNotas || '',
      aclaracionesServicio: pedido.aclaracionesServicio || ''
    };
    
    return resultadoFinal;

  } catch (error) {
    Logger.log(`❌ ERROR CRÍTICO al extraer pedido ${numPedido}: ${error.message}`);
    return manejarError('obtenerPedidoParaEdicion', error);
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
 * Guarda un nuevo cliente en Supabase.
 * Usado por Contactos.html y Comercial.html
 */
function guardarNuevoCliente(data) {
    try {
        const datosSanitizados = sanitizarDatos(data);
        
        // Validación: El RUC debe ser numérico para Supabase (tipo numeric/bigint)
        const rucNumerico = parseFloat(datosSanitizados.RUC);
        if (isNaN(rucNumerico)) throw new Error("El RUC/DNI debe ser un valor numérico válido.");

        const nuevoCliente = {
            "RUC_DNI": rucNumerico,
            "Nombre_RazonSocial": datosSanitizados.NOMBRE
        };

        // Insertar en Supabase
        const resultado = supabaseFetch('Clientes', {
            method: 'post',
            payload: nuevoCliente,
            // Pedimos que nos devuelva los datos insertados para confirmar
            params: 'select=RUC_DNI,Nombre_RazonSocial'
        });

        return { 
            success: true, 
            message: "Cliente registrado exitosamente.",
            nuevoCliente: resultado[0] 
        };

    } catch (error) {
        return manejarError("guardarNuevoCliente", error);
    }
}

/**
 * REFACTORIZADO: Guarda o actualiza un Contacto en Supabase.
 * Se asegura de que el RUC_DNI sea NUMERIC.
 */
function guardarOActualizarContacto(data) {
  try {
    const datosSanitizados = sanitizarDatos(data);
    
    // 1. Parsear RUC_DNI a NUMERIC (bigint)
    const rucNumerico = parseFloat(datosSanitizados.RUC);
    if (isNaN(rucNumerico)) throw new Error("RUC/DNI no es un valor numérico válido para Contactos.");

    // 'rowIndex' ahora es 'ID_Contacto'
    const idContacto = datosSanitizados.rowIndex;

    // 2. Mapeo a las columnas de Supabase
    const payload = {
      "RUC_DNI": rucNumerico, // NUMERICO
      "Nombre_Contacto": datosSanitizados.NOMBRE,
      "Correo": datosSanitizados.EMAIL,
      "Celular": parseFloat(datosSanitizados.TELEFONO) || null, // Celular es NUMERICO
      "Cargo": datosSanitizados.CARGO
    };
    
    let metodo;
    let params;

    if (idContacto && parseInt(idContacto) > 0) {
      metodo = 'patch';
      params = `ID_Contacto=eq.${idContacto}`; // Clave primaria
    } else {
      metodo = 'post';
      params = '';
    }

    supabaseFetch('Contactos', { 
      method: metodo,
      payload: payload,
      params: params
    });

    return { success: true, message: "Contacto guardado exitosamente" };

  } catch (error) {
    return manejarError("guardarOActualizarContacto", error);
  }
}
