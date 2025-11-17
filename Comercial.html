<!DOCTYPE html>
<html>
<head>
    <base target="_top">
    <title>Registro de Cotización</title>
    <?!= HtmlService.createHtmlOutputFromFile("Estilos").getContent(); ?>
</head>

<body>
    <div class="barraTopComercial">
        <div class="topBarra-izquierda">
            <? var url = getScriptUrl(); ?>
            <a href='<?= url ?>?page=Modulos' class="retornarModulos iconoTopBarra">Módulos</a>

            <div style="display: flex; align-items: center; gap: 20px;">
                <h3 id="tituloPrincipalContenido" class="titulo-contenido" style="margin-bottom: 0; border-bottom: none;">Nueva Cotización</h3>
                <div class="form-group" style="margin-bottom: 0; background-color: #f0f0f0; border-radius: 8px; padding: 5px 10px;">
                    <h2 style="color: black; margin: 0 !important; padding: 0 !important; font-weight: bold !important; border: none !important; font-size: 1em !important;">
                        ESTADO:
                    </h2>
                    <select id="estadoCotizacion" name="Estado" style="min-width: 150px; background-color: white; border: 2px solid #5cb85c; font-weight: bold; height: 32px !important; padding: 5px 8px !important;">
                        </select>
                </div>
            </div>
        </div>
        
        <div class="topBarra-derecha">
            <button onclick="navegarAResumen()" class="boton-regresar"> ← Resumen </button>
            <button type="button" onclick="llamarGenerarPDF()" class="btn-guardar-top btn-pdf" id="btnGenerarPDF" style="display: none;"> 📄 PDF </button> 
            <button type="button" onclick="gestionarActaGlobal()" class="btn-guardar-top btn-info" id="btnGestionarActa" style="display: none; background-color: #17a2b8;"> 📋 Acta </button> 
            <button type="button" onclick="gestionarOTsParaPedido()" class="btn-guardar-top btn-warning" id="btnGestionarOT" style="display: none; background-color: #f0ad4e;"> ⚙️ OTs </button>
            <button type="button" onclick="gestionarValorizaciones()" class="btn-guardar-top btn-success" id="btnGestionarValorizacion">    💲 Valorizaciones</button>
            </div>
        </div>
        <div class="contenedor-principal contenedor-compacto">
        
        <form id="formCotizacion" class="formulario-cotizacion">
            
            <div class="card datos-generales">
                <h3>Datos de la Cotización</h3>
                <div class="formulario-campos dos-columnas">
                    
                    <div>
                        <div class="form-group">
                            <h2>MONEDA</h2>
                            <select id="moneda" name="Moneda" required>
                                <option value="Soles">Soles (S/.)</option>
                                <option value="Dolares">Dólares ($)</option>
                            </select>
                        </div>
                        <div class="form-group">
                            <h2>RUC / DNI</h2>
                            <input type="text" id="rucCliente" name="RUC" placeholder="Escriba RUC o DNI" required
                                list="datalist-rucs-clientes"
                                onchange="manejarSeleccionCliente(this, 'clienteNombre')">
                        </div>
                        <div class="form-group">
                            <h2>DIRECCIÓN</h2>
                            <input type="text" id="direccion" name="Direccion" placeholder="Dirección de la obra" required>
                        </div>
                        <div class="form-group">
                            <h2>TURNO</h2>
                            <select id="turno" name="Turno" required></select>
                        </div>
                        <div class="form-group">
                            <h2>EJECUTIVO/ASESOR</h2>
                            <select id="ejecutivo" name="Ejecutivo" required>
                            <option value="" disabled selected>-- Seleccione --</option>
                            <option value="ANTHONY">ANTHONY</option>
                            <option value="CARMEN">CARMEN</option>
                            </select>
                        </div>
                    </div>
                    
                    <div>
                        <div class="form-group">
                            <h2>FORMA DE PAGO</h2>
                            <select id="formaPago" name="Forma_Pago" required></select>
                        </div>
                        <div class="form-group">
                            <h2>CLIENTE / RAZÓN SOCIAL</h2>
                            <input type="text" id="clienteNombre" name="Cliente" placeholder="Escriba el nombre o Razón Social" required
                                list="datalist-nombres-clientes"
                                onchange="manejarSeleccionCliente(this, 'rucCliente')">
                        </div>
                        <div class="form-group">
                            <h2>EMPRESA (Operadora)</h2>
                            <select id="empresa" name="Empresa" required></select>
                        </div>
                        <div class="form-group">
                             <h2>FECHA EJECUCIÓN SERVICIO</h2>
                             <input type="date" id="fechaEjecucion" name="fechaEjecucion" value="" required>
                        </div>
                        <div class="form-group">
                           <h2>CONTACTO</h2>
                           <select id="contacto" name="Contacto" required>
                           <option value="" disabled selected>-- Seleccione Contacto --</option>
                           </select>
                        </div>
                    </div>

                </div>
            </div>

            <div class="card detalle-servicios">
                <h3>Detalle de Servicios</h3>
                <table id="tablaDetalle" class="tabla-listado">
                    <thead>
                        <tr>
                            <th class="col-num">#</th>
                            <th class="col-servicio-cod">Servicio/Código</th>
                            <th class="col-nombre-servicio">Nombre del Servicio</th>
                            <th class="col-cant">Cant.</th>
                            <th class="col-und-medida">UND. MEDIDA</th>
                            <th class="col-precio-unit">Precio Unit.</th>
                            <th class="col-h-minimas">H. Mínimas</th>
                            <th class="col-und-h-min">UND. H. MIN.</th>
                            <th class="col-h-segun">H. SEGÚN</th>
                            <th class="col-dias-cotizados">Días Cotizados</th>
                            <th class="col-movilizacion">Movilización</th>
                            <th class="col-subtotal">Subtotal</th>
                            <th class="col-accion">Acción</th>
                            </tr>
                    </thead>
                    <tbody id="cuerpoDetalle"></tbody>
                </table>
                <button type="button" onclick="agregarLineaDetalle()" class="btn-guardar-top btn-agregar-servicio">➕ Agregar Servicio</button>
            </div>

            <div class="card resumen-total">
                 <h3 class="total-label">TOTAL: <span id="totalCotizacion">0.00</span></h3>
                 <input type="hidden" id="total" name="Total" value="0.00">
            </div>

            <div class="card notas-aclaratorias" style="margin-top: 20px; background-color: #fdfdfd;">
                <h3 style="border-bottom: 2px solid #007bff; padding-bottom: 5px;">Notas Aclaratorias del Servicio</h3>
                
                <div class="form-group" style="display: block; margin-bottom: 15px;">
                    <label for="plantillaNotas" style="font-weight: bold; font-size: 0.9em; color: #555; display: block; margin-bottom: 5px;">
                        Nombre de Plantilla (Opcional):
                    </label>
                    <input type="text" id="plantillaNotas" name="plantillaNotas" placeholder="Ej: Condiciones Estándar Grúa 50T" style="width: 100%;">
                </div>

                <div class="form-group" style="display: block;">
                    <label for="aclaracionesServicio" style="font-weight: bold; font-size: 0.9em; color: #555; display: block; margin-bottom: 5px;">
                        Aclaraciones (Texto Libre):
                    </label>
                    <textarea id="aclaracionesServicio" name="aclaracionesServicio" rows="8" style="width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px;"></textarea>
                </div>
            </div>
            <div class="contenedor-boton-guardar">
                 <button type="button" onclick="guardarRegistroCotizacion()" class="btn-guardar-top btn-guardar-principal">GUARDAR COTIZACIÓN</button>
            </div>
        </form>
    </div>

    <datalist id="datalist-codigos-servicio"></datalist>
    <datalist id="datalist-descripciones-servicio"></datalist>
    <datalist id="datalist-rucs-clientes"></datalist>
    <datalist id="datalist-nombres-clientes"></datalist>

    <div id="modalNuevoCliente" class="modal">
        <div class="modal-content">
            <span class="close-button" onclick="cerrarModalNuevoCliente()">&times;</span>
            <h3>Registrar Nuevo Cliente</h3>
            <form id="formNuevoCliente">
                <div class="formulario-campos">
                    <div>
                        <label for="newRUC">RUC / DNI *</label>
                        <input type="text" id="newRUC" name="RUC" required>
                    </div>
                    <div>
                        <label for="newNombre">Razón Social / Nombre *</label>
                        <input type="text" id="newNombre" name="NOMBRE" required>
                    </div>
                </div>
                <button type="button" onclick="registrarNuevoCliente()">Guardar y Seleccionar Cliente</button>
            </form>
        </div>
    </div>

    <div id="loadingOverlay" class="loading-overlay">
        <div class="loading-spinner">
            <p>Guardando cotización...</p>
            <div class="spinner"></div>
        </div>
    </div>

    <div id="errorModal" class="error-modal">
        <div class="error-content">
            <span class="close-button" onclick="cerrarErrorModal()">&times;</span>
            <h3>Errores de Validación</h3>
            <div id="listaErrores"></div>
            <button type="button" onclick="cerrarErrorModal()">Cerrar</button>
        </div>
    </div>

<script>
  function gestionarValorizaciones() {
    if (!estadoApp.esEdicion || !estadoApp.numPedidoEdicion) return;
    const urlBase = '<?!= getScriptUrl() ?>';
    const urlDestino = `${urlBase}?page=Valorizacion&pedido=${encodeURIComponent(estadoApp.numPedidoEdicion)}`;
    window.open(urlDestino, '_blank');
    }

  function navegarAResumen() {
        const urlBase = '<?!= getScriptUrl() ?>';
        window.top.location.href = urlBase + '?page=ResumenComercial';
    }
  
  function llamarGenerarPDF() {
         if (!estadoApp.esEdicion || !estadoApp.numPedidoEdicion) {
            alert('Solo se puede generar PDF para una cotización existente (en modo edición).');
            return;
        }
        generarPDFDesdeComercial(estadoApp.numPedidoEdicion);
     }
  
  function generarPDFDesdeComercial(numPedido) {
        mostrarLoading(true, 'Generando PDF...');
        google.script.run
            .withSuccessHandler(function(resultado) { 
                mostrarLoading(false); 
                if(resultado.success){ 
                    window.open(resultado.pdfUrl, '_blank'); 
                    if(resultado.sheetUrl) { window.open(resultado.sheetUrl, '_blank'); }
                    alert('Archivos generados exitosamente.'); 
                } else { 
                    alert('Error al generar PDF: '+ (resultado.error || 'Error desconocido')); 
                } 
            })
            .withFailureHandler(function(error) { 
                mostrarLoading(false); 
                alert('Error Conexión (PDF): '+error.message); 
            })
            // Asegúrate que esta función exista en tu .gs
            // Si no, usa la que te di: generarPDFPorNumeroPedido
            .generarYDevolverPDF(numPedido); 
    }
/**
 * REEMPLAZO (Misión 3 - Revisada)
 * Abre la página de Actas de Planificación, pasando el N° de Pedido actual.
 */
function gestionarActaGlobal() {
    // 1. Verifica que estemos en modo edición
    if (!estadoApp.esEdicion || !estadoApp.numPedidoEdicion) {
        alert('Debe estar en modo de edición de un pedido existente para gestionar sus Actas.');
        return;
    }

    // 2. Obtiene la URL base y el N° de Pedido
    const urlBase = '<?!= getScriptUrl() ?>';
    const numPedido = estadoApp.numPedidoEdicion;

    // 3. Construye la URL de destino
    const urlDestino = `${urlBase}?page=Acta&pedido=${encodeURIComponent(numPedido)}`;

    // 4. Abre en una nueva pestaña
    console.log(`Navegando a la gestión de Actas para el pedido: ${numPedido}`);
    window.open(urlDestino, '_blank');
}

/**
 * NUEVA FUNCIÓN
 * Abre la página de resumen de OTs, pasando el número de pedido
 * actual para que esa página se filtre sola.
 */
function gestionarOTsParaPedido() {
    // 1. Verifica que estemos en modo edición y tengamos un número de pedido
    if (!estadoApp.esEdicion || !estadoApp.numPedidoEdicion) {
        alert('Debe estar en modo de edición de un pedido existente para gestionar sus OTs.');
        return;
    }

    // 2. Obtiene la URL base
    const urlBase = '<?!= getScriptUrl() ?>';
    const numPedido = estadoApp.numPedidoEdicion;

    // 3. Construye la URL de destino
    // Asumimos que tu página de resumen de OTs se llama 'OT'
    // Le pasamos el número de pedido como un parámetro URL
    const urlDestino = `${urlBase}?page=OT&pedido=${encodeURIComponent(numPedido)}`;

    // 4. Abre en una nueva pestaña
    console.log(`Navegando a la gestión de OTs para el pedido: ${numPedido}`);
    window.open(urlDestino, '_blank');
}

// Función para generar PDF
function generarPDF() {
    const numPedido = estadoApp.esEdicion ? estadoApp.numPedidoEdicion : 
                     document.querySelector('input[name="numPedido"]')?.value;
    
    if (!numPedido) {
        alert('No se puede generar PDF: número de pedido no disponible');
        return;
    }
    
    mostrarLoading(true, 'Generando PDF...');
    
    google.script.run
        .withSuccessHandler(function(resultado) {
            mostrarLoading(false);
            if (resultado.success) {
                // Abrir el PDF en una nueva pestaña
                window.open(resultado.pdfUrl, '_blank');
                alert('PDF generado exitosamente');
            } else {
                alert('Error al generar PDF: ' + resultado.message);
            }
        })
        .withFailureHandler(function(error) {
            mostrarLoading(false);
            alert('Error al generar PDF: ' + error.message);
        })
        .generarCotizacionPDF(numPedido);
}


        // =============================================================================
        // VARIABLES GLOBALES Y CONFIGURACIÓN
        // =============================================================================
const estadoApp = {
            lineaCounter: 0,
            datosCargados: false,
            guardando: false,
            esEdicion: false,
            numPedidoEdicion: null,
            configuracion: {
                VALOR_DIAS_MES: 30,
                VALOR_HORAS_DIA: 24,
                decimales: 2
            }
        };

        let listaClientes = [];
        let listaServicios = [];
        let listaFormasPago = [];
        let listaEmpresas = [];
        let listaTurnos = [];
        let listaUndMedida = [];
        let listaHorasMinimas = [];
        let listaHorasSegun = [];
        let listaContactos = [];
        let listaEjecutivos = ['ANTHONY', 'CARMEN'];

        // =============================================================================
        // FUNCIONES DE UTILIDAD Y MANEJO DE ERRORES
        // =============================================================================
        function mostrarLoading(mostrar, mensaje = 'Cargando...') {
            const overlay = document.getElementById('loadingOverlay');
            if (overlay) {
                overlay.style.display = mostrar ? 'flex' : 'none';
                const mensajeElement = overlay.querySelector('p');
                if (mensajeElement && mensaje) {
                    mensajeElement.textContent = mensaje;
                }
            }
            estadoApp.guardando = mostrar;
        }

        function mostrarError(error, contexto = '') {
            console.error(`Error en ${contexto}:`, error);
            const mensaje = error.message || error;
            const mensajeAmigable = traducirError(mensaje);
            alert(`Error en ${contexto}: ${mensajeAmigable}`);
        }

        function traducirError(mensaje) {
            const erroresComunes = {
                'Exception:': 'Error del sistema:',
                'Timed out': 'La operación tardó demasiado tiempo',
                'Access denied': 'Acceso denegado',
                'Service invoked too many times': 'Límite de servicio excedido'
            };
            for (const [key, value] of Object.entries(erroresComunes)) {
                if (mensaje.includes(key)) return value;
            }
            return mensaje;
        }

        function mostrarErroresValidacion(errores) {
            const modal = document.getElementById('errorModal');
            const lista = document.getElementById('listaErrores');
            if (modal && lista) {
                lista.innerHTML = errores.map(error => `<p class="error-item">• ${error}</p>`).join('');
                modal.style.display = 'flex';
            } else {
                alert('Errores de validación:\n' + errores.join('\n• '));
            }
        }

        function cerrarErrorModal() {
            const modal = document.getElementById('errorModal');
            if (modal) modal.style.display = 'none';
        }

function mostrarExito(respuesta) {
            // 1. Mostrar el mensaje de éxito
            const mensaje = respuesta.message || "Cotización guardada exitosamente";
            alert(mensaje);
            
            // 2. NO REDIRIGIR AUTOMÁTICAMENTE.
            // En su lugar, actualizamos el estado de la app a "modo edición".
            
            // Si la respuesta del servidor incluye el nuevo codigoPedido 
            // y aún no estábamos en modo edición...
            if (respuesta.codigoPedido && !estadoApp.esEdicion) {
                console.log(`Guardado exitoso. Cambiando a modo edición para: ${respuesta.codigoPedido}`);
                estadoApp.esEdicion = true;
                estadoApp.numPedidoEdicion = respuesta.codigoPedido;

                // Actualizar la UI para reflejar el modo edición
                const tituloContenido = document.getElementById('tituloPrincipalContenido');
                if (tituloContenido) {
                     tituloContenido.textContent = `Pedido: ${respuesta.codigoPedido}`;
                     tituloContenido.style.color = '#e74c3c'; // Color rojo para indicar "edición"
                }
                
                // Mostrar los botones de acción que estaban ocultos
                document.getElementById('btnGenerarPDF').style.display = 'inline-block';
                document.getElementById('btnGestionarActa').style.display = 'inline-block';
                document.getElementById('btnGestionarOT').style.display = 'inline-block';
                document.getElementById('btnGestionarValorizacion').style.display = 'inline-block';

                // Deshabilitar el campo RUC
                document.getElementById('rucCliente').readOnly = true;
                document.getElementById('rucCliente').style.backgroundColor = '#f5f5f5';

            } else {
                 // Si ya estábamos en modo edición, solo mostramos el alert de "Actualizado"
                 console.log("Actualización completada.");
            }

            // El usuario ahora puede usar los botones "Resumen", "OTs", "PDF", etc.
            // NO USAR: window.top.location.href = ...;
        }

        function puedeRealizarAccion(accion) {
            if (estadoApp.guardando) {
                mostrarError('El sistema está guardando datos. Espere por favor.', accion);
                return false;
            }
            if (!estadoApp.datosCargados && accion !== 'cargarDatos') {
                mostrarError('Los datos iniciales no se han cargado completamente', accion);
                return false;
            }
            return true;
        }

        function inicializarFecha() {
            const fechaInput = document.getElementById('fechaRegistro');
            if (fechaInput) {
                const ahora = new Date();
                const offset = ahora.getTimezoneOffset();
                const fechaLocal = new Date(ahora.getTime() - (offset * 60 * 1000));
                fechaInput.value = fechaLocal.toISOString().split('T')[0];
            }
        }

        // =============================================================================
        // FUNCIÓN DE DEBUG MEJORADA
        // =============================================================================
        function verificarDatosCargados() {
            console.log("🔍 VERIFICACIÓN DE DATOS CARGADOS:");
    
            const direccionElem = document.getElementById('direccion');
            const turnoElem = document.getElementById('turno');
            const empresaElem = document.getElementById('empresa');
            const clienteElem = document.getElementById('clienteNombre');
            const fechaElem = document.getElementById('fechaEjecucion');
            const monedaElem = document.getElementById('moneda');
            const formaPagoElem = document.getElementById('formaPago');
            const ejecutivoElem = document.getElementById('ejecutivo');
            const contactoElem = document.getElementById('contacto');
    
            console.log("📍 Dirección:", direccionElem ? direccionElem.value : "Elemento no encontrado");
            console.log("🕐 Turno:", turnoElem ? turnoElem.value : "Elemento no encontrado");
            console.log("🏢 Empresa:", empresaElem ? empresaElem.value : "Elemento no encontrado");
            console.log("👤 Cliente:", clienteElem ? clienteElem.value : "Elemento no encontrado");
            console.log("📅 Fecha:", fechaElem ? fechaElem.value : "Elemento no encontrado");
            console.log("💰 Moneda:", monedaElem ? monedaElem.value : "Elemento no encontrado");
            console.log("💳 Forma Pago:", formaPagoElem ? formaPagoElem.value : "Elemento no encontrado");
            console.log("👨‍💼 Ejecutivo:", ejecutivoElem ? ejecutivoElem.value : "Elemento no encontrado");
            console.log("📞 Contacto:", contactoElem ? contactoElem.value : "Elemento no encontrado");
    
            console.log("📊 Estado App:", estadoApp);
    
            const lineas = document.querySelectorAll('#cuerpoDetalle tr');
            console.log("📋 Líneas en tabla:", lineas.length);
        }

        // =============================================================================
        // FUNCIONES DE VALIDACIÓN
        // =============================================================================
        function validarFormulario() {
            const errores = [];
            
            // ✅ Validar campos generales
            const empresa = document.getElementById('empresa').value;
            const ruc = document.getElementById('rucCliente').value.trim();
            const cliente = document.getElementById('clienteNombre').value.trim();
            const moneda = document.getElementById('moneda').value;
            const formaPago = document.getElementById('formaPago').value;
            const direccion = document.getElementById('direccion').value.trim();
            const ejecutivo = document.getElementById('ejecutivo').value;
            const contacto = document.getElementById('contacto').value;
            
            if (!empresa) errores.push('La empresa es obligatoria');
            if (!ruc) errores.push('El RUC/DNI es obligatorio');
            if (!cliente) errores.push('El cliente es obligatorio');
            if (!moneda) errores.push('La moneda es obligatoria');
            if (!formaPago) errores.push('La forma de pago es obligatoria');
            if (!direccion) errores.push('La dirección es obligatoria');
            if (!ejecutivo) errores.push('El ejecutivo es obligatorio');
            if (!contacto) errores.push('El contacto es obligatorio');
            
            // ✅ Validar líneas de detalle
            const lineas = obtenerLineasDetalleCorregidas();
            if (lineas.length === 0) {
                errores.push('Debe agregar al menos un servicio válido');
            } else {
                lineas.forEach((linea, index) => {
                    if (!linea.cod) errores.push(`Línea ${index + 1}: El código de servicio es obligatorio`);
                    if (linea.cantidad <= 0) errores.push(`Línea ${index + 1}: La cantidad debe ser mayor a 0`);
                    if (linea.precio < 0) errores.push(`Línea ${index + 1}: El precio no puede ser negativo`);
                });
            }
            
            console.log("✅ VALIDACIÓN COMPLETADA - Errores:", errores.length);
            return errores;
        }

        function validarRUC(ruc) {
            const cleanRuc = ruc.trim();
            if (!cleanRuc) return false;
            if (cleanRuc.length === 8) return /^\d+$/.test(cleanRuc); // DNI
            if (cleanRuc.length === 11) return /^\d+$/.test(cleanRuc); // RUC
            return true;
        }

// =============================================================================
// FUNCIONES DE INICIALIZACIÓN Y CARGA DE DATOS - CORREGIDAS
// =============================================================================
document.addEventListener("DOMContentLoaded", function() {
    console.log("🔄 Comercial.html CARGADO - Iniciando proceso...");
    // inicializarFecha(); // Ya no es necesario si no hay campo fechaRegistro
    
    let editId = null;
    const urlParams = new URLSearchParams(window.location.search);
    editId = urlParams.get('editar');
    if (!editId) {
        editId = sessionStorage.getItem('pedidoAEditar');
        if (editId) {
            console.log("📦 EditId obtenido de sessionStorage:", editId);
            sessionStorage.removeItem('pedidoAEditar');
        }
    }
    console.log("🎯 ID final para edición:", editId);
    
    const tituloContenido = document.getElementById('tituloPrincipalContenido'); 

    // --- Lógica de Edición (Botones Visibles) ---
    if (editId) {
        estadoApp.esEdicion = true;
        estadoApp.numPedidoEdicion = editId;
        if (tituloContenido) {
             tituloContenido.textContent = `Pedido: ${editId}`;
             tituloContenido.style.color = '#e74c3c';
        }
        
        // Mostrar botones de Edición
        const btnPDF = document.getElementById('btnGenerarPDF');
        if (btnPDF) btnPDF.style.display = 'inline-block';
        
        const btnActa = document.getElementById('btnGestionarActa');
        if (btnActa) btnActa.style.display = 'inline-block';
        
        const btnOT = document.getElementById('btnGestionarOT');
        if (btnOT) btnOT.style.display = 'inline-block';
        
    // --- Lógica de Creación (Botones Ocultos) ---
    } else {
         if (tituloContenido) {
             tituloContenido.textContent = 'Nueva Cotización';
         }
         
         // Ocultar botones de Edición
         const btnPDF = document.getElementById('btnGenerarPDF');
         if (btnPDF) btnPDF.style.display = 'none';
         
         const btnActa = document.getElementById('btnGestionarActa');
         if (btnActa) btnActa.style.display = 'none';

         const btnOT = document.getElementById('btnGestionarOT');
         if (btnOT) btnOT.style.display = 'none';
    }

    // --- Resto de tu función (Llamadas al servidor) ---
    mostrarLoading(true, "Cargando datos iniciales...");
    // 1. PRIMERA LLAMADA (Ligera)
    google.script.run
        .withSuccessHandler(function(dataInicial) {
            if (dataInicial && dataInicial.success) {
                console.log("✅ Datos LIGEROS cargados");
                listaFormasPago = dataInicial.formasPago || [];
                listaEmpresas = dataInicial.listaEmpresas || [];
                listaTurnos = dataInicial.listaTurnos || [];
                listaHorasMinimas = dataInicial.listaHorasMinimas || [];
                listaEjecutivos = dataInicial.listaEjecutivos || ['ANTHONY', 'CARMEN'];
                
                // --- CORRECCIÓN AQUÍ ---
                const listaEstados = dataInicial.listaEstadosCot || []; // Obtener la lista de estados
                llenarSelect('estadoCotizacion', listaEstados); // Llenar el select de estado ANTES de la llamada pesada
                // --- FIN CORRECCIÓN ---

                llenarSelect('formaPago', listaFormasPago);
                llenarSelect('empresa', listaEmpresas);
                llenarSelect('turno', listaTurnos);
                llenarSelect('ejecutivo', listaEjecutivos);

                // 2. SEGUNDA LLAMADA (Pesada)
                mostrarLoading(true, "Cargando Clientes y Servicios...");
                google.script.run
                    .withSuccessHandler(function(dataPesada) {
                        mostrarLoading(false);
                        if (dataPesada && dataPesada.success) {
                            console.log("✅ Datos PESADOS cargados");
                            listaClientes = dataPesada.clientes || [];
                            listaServicios = dataPesada.servicios || [];
                            listaUndMedida = dataPesada.listaUndMedida || [];
                            listaHorasSegun = dataPesada.listaHorasSegun || [];
                            llenarDatalistsClientes();
                            llenarDatalistsServicios();
                            estadoApp.datosCargados = true;

                            // 3. Lógica de Edición/Nuevo
                            if (estadoApp.esEdicion) {
                                console.log("🚀 CARGANDO DATOS DE EDICIÓN:", estadoApp.numPedidoEdicion);
                                cargarPedidoParaEdicion(estadoApp.numPedidoEdicion);
                            } else {
                                console.log("📝 MODO NUEVA COTIZACIÓN");
                                // Inicializar fecha de ejecución si es nuevo pedido
                                const fechaEjecInput = document.getElementById('fechaEjecucion');
                                if(fechaEjecInput && !fechaEjecInput.value){
                                     const ahora = new Date();
                                     const offset = ahora.getTimezoneOffset();
                                     const fechaLocal = new Date(ahora.getTime() - (offset * 60 * 1000));
                                     fechaEjecInput.value = fechaLocal.toISOString().split('T')[0];
                                }
                                // Poner estado por defecto si es nuevo
                                const estadoSelect = document.getElementById('estadoCotizacion');
                                if (estadoSelect && estadoSelect.options.length > 1) {
                                     estadoSelect.value = estadoSelect.options[1].value; // Selecciona la primera opción real
                                }
                                agregarLineaDetalle();
                            }
                        } else { mostrarError(dataPesada?.error || 'Error al cargar Clientes/Servicios', 'inicialización pesada'); }
                    })
                    .withFailureHandler(function(error) { mostrarLoading(false); mostrarError(error, 'cargar datos pesados'); })
                    .getDatosPesadosComercial();
            } else { mostrarLoading(false); mostrarError(dataInicial?.error || 'Error al cargar datos iniciales', 'inicialización ligera'); }
        })
        .withFailureHandler(function(error) { mostrarLoading(false); console.error("❌ Error cargando datos iniciales:", error); mostrarError(error, 'cargar datos iniciales'); })
        .getDatosInicialesComercial();
});


function procesarDatosInicialesComercial(data) {
    console.log("🔄 Procesando datos iniciales");
    
    if (!data || !data.success) {
        mostrarError(data?.error || 'Datos inválidos recibidos', 'procesar datos iniciales');
        return;
    }
    
    // Cargar todas las listas necesarias
    listaClientes = data.clientes || [];
    listaServicios = data.servicios || [];
    listaFormasPago = data.formasPago || [];
    listaEmpresas = data.listaEmpresas || [];
    listaTurnos = data.listaTurnos || [];
    listaUndMedida = data.listaUndMedida || [];
    listaHorasMinimas = data.listaHorasMinimas || [];
    listaHorasSegun = data.listaHorasSegun || [];
    listaEjecutivos = data.listaEjecutivos || ['ANTHONY', 'CARMEN'];

    // Llenar selects
    llenarSelect('formaPago', listaFormasPago);
    llenarSelect('empresa', listaEmpresas);
    llenarSelect('turno', listaTurnos);
    llenarSelect('ejecutivo', listaEjecutivos);

    // Llenar datalists
    llenarDatalistsClientes();
    llenarDatalistsServicios();

    estadoApp.datosCargados = true;
    console.log("✅ Datos iniciales procesados correctamente");
    
    // NO manejar edición aquí - ya se maneja después en el success handler
}
        // =============================================================================
        // FUNCIONES DE GESTIÓN DE CONTACTOS
        // =============================================================================

        function llenarDatalistContactos() {
            const contactoSelect = document.getElementById('contacto');
            
            if (!contactoSelect) {
                console.warn("❌ Select de contacto no encontrado en el DOM");
                return;
            }

            // Limpiar opciones existentes
            contactoSelect.innerHTML = '<option value="" disabled selected>-- Seleccione Contacto --</option>';

            console.log("👥 Contactos disponibles para llenar:", listaContactos);

            if (listaContactos.length === 0) {
                const optEmpty = document.createElement('option');
                optEmpty.value = "";
                optEmpty.textContent = "No hay contactos registrados";
                optEmpty.disabled = true;
                contactoSelect.appendChild(optEmpty);
                console.log("ℹ️ No hay contactos para mostrar");
                return;
            }

            // Llenar con contactos disponibles
            listaContactos.forEach(contacto => {
                const displayText = contacto.display || `${contacto.nombre}${contacto.cargo ? ' - ' + contacto.cargo : ''}`;
                
                // Para select
                const optSelect = document.createElement('option');
                optSelect.value = displayText;
                optSelect.textContent = displayText;
                optSelect.setAttribute('data-email', contacto.email || '');
                optSelect.setAttribute('data-telefono', contacto.telefono || '');
                contactoSelect.appendChild(optSelect);
            });
            
            console.log("✅ Datalist de contactos actualizado con " + listaContactos.length + " contactos");
        }

        function cargarContactosPorRUC(ruc) {
            if (!ruc) return;
            
            console.log("📞 Cargando contactos para RUC:", ruc);
            
            google.script.run
                .withSuccessHandler(function(contactosData) { // <--- ESTO YA LO TIENES
                    console.log("✅ Contactos recibidos:", contactosData);
                    
                    if (contactosData && contactosData.length > 0) {
                        listaContactos = contactosData;
                        llenarDatalistContactos();
                    } else {
                        console.log("ℹ️ No se encontraron contactos para este RUC");
                        listaContactos = [];
                        llenarDatalistContactos();
                    }
                })
                .withFailureHandler(function(error) { // <--- AÑADE ESTO
                    console.error("❌ Error cargando contactos:", error);
                    alert("Error al cargar contactos: " + error.message); // Muestra el error
                    listaContactos = [];
                    llenarDatalistContactos();
                })
                .getContactosParaComercial(ruc);
        }

        function seleccionarContactoEspecifico(contactoBuscado) {
            const contactoSelect = document.getElementById('contacto');
            if (!contactoSelect) {
                console.warn("❌ Select de contacto no encontrado");
                return;
            }
            
            console.log(`🔍 Buscando contacto específico: "${contactoBuscado}"`);
            
            // Buscar coincidencia exacta
            let opcionContacto = Array.from(contactoSelect.options).find(opt => 
                opt.value === contactoBuscado || opt.text === contactoBuscado
            );
            
            // Si no encuentra exacto, buscar coincidencia parcial
            if (!opcionContacto) {
                opcionContacto = Array.from(contactoSelect.options).find(opt => 
                    opt.value.includes(contactoBuscado) || 
                    opt.text.includes(contactoBuscado) ||
                    contactoBuscado.includes(opt.value) ||
                    contactoBuscado.includes(opt.text)
                );
            }
            
            if (opcionContacto) {
                contactoSelect.value = opcionContacto.value;
                console.log(`✅ Contacto seleccionado: "${opcionContacto.value}"`);
            } else {
                console.warn(`⚠️ No se pudo encontrar contacto: "${contactoBuscado}"`);
            }
        }

        function llenarSelect(id, opciones) {
            const select = document.getElementById(id);
            if (!select) {
                console.warn(`⚠️ Select con id '${id}' no encontrado`);
                return;
            }
            
            select.innerHTML = '<option value="" disabled selected>-- Seleccione --</option>';
            if (opciones && opciones.length > 0) {
                opciones.forEach(opcion => {
                    const opt = document.createElement('option');
                    opt.value = opcion;
                    opt.textContent = opcion;
                    select.appendChild(opt);
                });
            } else {
                console.warn(`⚠️ No hay opciones para llenar el select '${id}'`);
            }
        }

        // =============================================================================
        // FUNCIONES DE GESTIÓN DE CLIENTES
        // =============================================================================
        function llenarDatalistsClientes() {
    const rucsDatalist = document.getElementById('datalist-rucs-clientes');
    const nombresDatalist = document.getElementById('datalist-nombres-clientes');
    if (!rucsDatalist || !nombresDatalist) return;

    rucsDatalist.innerHTML = '';
    nombresDatalist.innerHTML = '';

    // ANTES: listaClientes.slice(1).forEach(cliente => { ... cliente[0], cliente[1] })
    // AHORA:
    listaClientes.forEach(cliente => {
        // Usamos los nombres de columna de tu CSV/Supabase
        const ruc = cliente.RUC_DNI; 
        const nombre = cliente.Nombre_RazonSocial;

        if (ruc && nombre) {
            const optRuc = document.createElement('option');
            optRuc.value = ruc;
            optRuc.label = nombre;
            rucsDatalist.appendChild(optRuc);

            const optNom = document.createElement('option');
            optNom.value = nombre;
            optNom.label = ruc;
            nombresDatalist.appendChild(optNom);
        }
    });

    // (La lógica de "-- CLIENTE NUEVO --" sigue igual)
    const optNuevoRuc = document.createElement('option');
    optNuevoRuc.value = "-- CLIENTE NUEVO --";
    optNuevoRuc.label = "Registrar nuevo RUC/DNI";
    rucsDatalist.appendChild(optNuevoRuc);

    const optNuevoNom = document.createElement('option');
    optNuevoNom.value = "-- CLIENTE NUEVO --";
    optNuevoNom.label = "Registrar nuevo Nombre/Razón Social";
    nombresDatalist.appendChild(optNuevoNom);
}

        function manejarSeleccionCliente(inputElement, targetId) {
    if (!puedeRealizarAccion('seleccionar cliente')) return;

    const selectedValue = inputElement.value.trim();
    const targetInput = document.getElementById(targetId);

    if (selectedValue === "-- CLIENTE NUEVO --") {
        abrirModalNuevoCliente();
        inputElement.value = '';
        return;
    }

    if (!selectedValue) {
        targetInput.value = "";
        return;
    }

    const esRucInput = (inputElement.id === 'rucCliente');

    // ANTES: Buscaba en un array de arrays
    // AHORA: Busca en un array de objetos
    let clienteEncontrado;
    if (esRucInput) {
        clienteEncontrado = listaClientes.find(c => 
            String(c.RUC_DNI).trim().toUpperCase() === selectedValue.toUpperCase()
        );
    } else {
        clienteEncontrado = listaClientes.find(c => 
            String(c.Nombre_RazonSocial).trim().toUpperCase() === selectedValue.toUpperCase()
        );
    }

    if (clienteEncontrado) {
        if (esRucInput) {
            targetInput.value = clienteEncontrado.Nombre_RazonSocial;
        } else {
            targetInput.value = clienteEncontrado.RUC_DNI;
        }

        const rucFinal = clienteEncontrado.RUC_DNI;
        if (rucFinal) {
            cargarContactosPorRUC(rucFinal); // Esta función ya está bien
        }

    } else {
        // (La lógica de "Cliente no registrado" sigue igual)
        const rucActual = document.getElementById('rucCliente').value.trim();
        const nombreActual = document.getElementById('clienteNombre').value.trim();
        if (rucActual && nombreActual && rucActual !== "-- CLIENTE NUEVO --" && nombreActual !== "-- CLIENTE NUEVO --") {
            const confirmar = confirm(`El dato '${selectedValue}' no está registrado. ¿Desea registrar al nuevo cliente con RUC: ${rucActual} y Nombre: ${nombreActual}?`);
            if (confirmar) abrirModalNuevoCliente(rucActual, nombreActual);
            else inputElement.value = targetInput.value = '';
        }
    }
}

        function abrirModalNuevoCliente(ruc = '', nombre = '') {
            const modal = document.getElementById('modalNuevoCliente');
            if (modal) {
                modal.style.display = 'block';
                document.getElementById('newRUC').value = ruc;
                document.getElementById('newNombre').value = nombre;
            }
        }

        function cerrarModalNuevoCliente() {
            const modal = document.getElementById('modalNuevoCliente');
            if (modal) {
                modal.style.display = 'none';
                document.getElementById('formNuevoCliente').reset();
            }
        }

        function registrarNuevoCliente() {
            const ruc = document.getElementById('newRUC').value.trim();
            const nombre = document.getElementById('newNombre').value.trim();

            if (!ruc || !nombre) {
                alert("RUC y Nombre son obligatorios para registrar un nuevo cliente.");
                return;
            }

            if (!validarRUC(ruc)) {
                alert("El RUC/DNI no tiene un formato válido.");
                return;
            }

            mostrarLoading(true, "Registrando nuevo cliente...");
            google.script.run
                .withSuccessHandler(function(response) {
                    mostrarLoading(false);
                    if (response.success) {
                        listaClientes.push(response.nuevoCliente); 
                        llenarDatalistsClientes();
                        document.getElementById('rucCliente').value = ruc;
                        document.getElementById('clienteNombre').value = nombre;
                        cerrarModalNuevoCliente();
                        alert("Cliente registrado exitosamente");
                    } else {
                        mostrarError(response.message, 'registrar cliente');
                    }
                })
                .withFailureHandler(function(error) {
                    mostrarLoading(false);
                    mostrarError(error, 'registrar cliente');
                })
                .guardarNuevoCliente({ RUC: ruc, NOMBRE: nombre });
        }

        // =============================================================================
        // FUNCIONES DE GESTIÓN DE DETALLES Y SERVICIOS
        // =============================================================================
        function llenarDatalistsServicios() {
    const codigosDatalist = document.getElementById('datalist-codigos-servicio');
    if (!codigosDatalist) return;
    codigosDatalist.innerHTML = '';

    // ANTES: listaServicios.slice(1).forEach(servicio => { ... servicio[0], servicio[1] })
    // AHORA:
    listaServicios.forEach(servicio => {
        const codigo = servicio.ID_servicios; // De tu CSV
        const descripcion = servicio.Nombre_Servicio; // De tu CSV
        if (codigo && descripcion) {
            const optCod = document.createElement('option');
            optCod.value = `${codigo} - ${descripcion}`;
            codigosDatalist.appendChild(optCod);
        }
    });
}

        function generarOpcionesSelect(opciones, valorSeleccionado = '') {
            let html = '<option value="" disabled>Seleccione</option>';
            if (opciones && opciones.length > 0) {
                opciones.forEach(opcion => {
                    const selected = opcion === valorSeleccionado ? 'selected' : '';
                    html += `<option value="${opcion}" ${selected}>${opcion}</option>`;
                });
            }
            return html;
        }

        function optimizarSelectoresFila(row, rowId) {
            return {
                cantidad: row.querySelector(`input[name="cantidad_${rowId}"]`),
                undMedida: row.querySelector(`select[name="und_medida_${rowId}"]`),
                hMinNum: row.querySelector(`input[name="horas_minimas_num_${rowId}"]`),
                precio: row.querySelector(`input[name="precio_unitario_${rowId}"]`),
                movilizacion: row.querySelector(`input[name="movilizacion_${rowId}"]`),
                undHMinSelect: row.querySelector(`select[name="und_horas_minimas_${rowId}"]`),
                diasDisplay: row.querySelector('.dias-cotizados-display'),
                diasHidden: row.querySelector(`input[name="dias_cotizados_val_${rowId}"]`),
                subtotalDisplay: row.querySelector('.subtotal-display'),
                subtotalHidden: row.querySelector(`input[name="subtotal_linea_${rowId}"]`)
            };
        }

        function agregarLineaDetalle(lineaData = null) {
            if (!puedeRealizarAccion('agregar línea') && !lineaData) return;
            
            estadoApp.lineaCounter++;
            const rowId = estadoApp.lineaCounter;
            const tbody = document.getElementById('cuerpoDetalle');
            const tr = document.createElement('tr');
            tr.id = `linea-${rowId}`;

            const def = {
                COD: '', DESC: '', UND: 'HORAS', CANTIDAD: 1, 
                UND_H_MIN: '', DIAS_COT: 0, H_MIN_NUM: 0, H_SEGUN: '', MOV: 0.00, PRECIO: 0.00, SUB: 0.00
            };
            const data = lineaData || def;
            
            const codigoMostrar = data.COD && data.DESC ? `${data.COD} - ${data.DESC}` : '';
            const undMedidaOptions = generarOpcionesSelect(listaUndMedida, data.UND);
            const horasMinOptions = generarOpcionesSelect(listaHorasMinimas, data.UND_H_MIN);
            const horasSegunOptions = generarOpcionesSelect(listaHorasSegun, data.H_SEGUN);
            const necesitaTiempo = data.UND === 'HORAS' || data.UND === 'DÍAS';
            const disabledAttr = !necesitaTiempo ? 'disabled' : '';

            tr.innerHTML = `
                <td>${rowId}</td>
                <td> 
                    <input type="text" class="form-control detalle-input"
                        placeholder="Código - Nombre"
                        list="datalist-codigos-servicio"
                        value="${codigoMostrar}"
                        name="Servicio_Busqueda_${rowId}"
                        onchange="cargarDatosServicio(this)"
                        onblur="this.value = this.value.toUpperCase().trim()"
                        required>
                    <input type="hidden" name="Servicio_cod_${rowId}" value="${data.COD}">
                </td>
                <td><input type="text" name="descripcion_${rowId}" value="${data.DESC}" class="descripcion-servicio detalle-input" readonly></td>
                <td><input type="number" name="cantidad_${rowId}" value="${data.CANTIDAD}" min="0" oninput="calcularTodo(this)" class="detalle-input text-right" required></td>
                <td> 
                    <select name="und_medida_${rowId}" onchange="manejarUnidadMedida(this); calcularTodo(this);" class="detalle-select" required>
                        ${undMedidaOptions}
                    </select>
                </td>
                <td><input type="number" name="precio_unitario_${rowId}" value="${parseFloat(data.PRECIO).toFixed(2)}" min="0" oninput="calcularTodo(this)" step="0.01" class="detalle-input text-right" required></td>
                <td><input type="number" name="horas_minimas_num_${rowId}" value="${data.H_MIN_NUM}" min="0" oninput="calcularTodo(this)" class="detalle-input text-right" ${disabledAttr}></td>
                <td> 
                    <select name="und_horas_minimas_${rowId}" onchange="calcularTodo(this)" class="detalle-select" ${disabledAttr} required>
                        ${horasMinOptions}
                    </select>
                </td>
                <td> 
                    <select name="hora_segun_${rowId}" required onchange="calcularTodo(this)" class="detalle-select">
                        ${horasSegunOptions}
                    </select>
                </td>
                <td><span class="dias-cotizados-display display-only-cell text-right">${Math.ceil(data.DIAS_COT).toFixed(0)}</span></td>
                <td><input type="number" name="movilizacion_${rowId}" value="${parseFloat(data.MOV).toFixed(2)}" min="0" oninput="calcularTodo(this)" step="0.01" class="detalle-input text-right"></td> 
                <td><span class="subtotal-display display-only-cell text-right">${parseFloat(data.SUB).toFixed(2)}</span></td>
                <td><button type="button" onclick="eliminarLineaDetalle('linea-${rowId}')">🗑️</button></td>
                
                <input type="hidden" name="subtotal_linea_${rowId}" value="${parseFloat(data.SUB).toFixed(2)}">
                <input type="hidden" name="dias_cotizados_val_${rowId}" value="${Math.ceil(data.DIAS_COT).toFixed(0)}">
            `;

            tbody.appendChild(tr);
            const undSelect = tr.querySelector(`select[name="und_medida_${rowId}"]`);
            manejarUnidadMedida(undSelect);
        }

        function cargarDatosServicio(inputElement) {
    if (!puedeRealizarAccion('cargar servicio')) return;

    const row = inputElement.closest('tr');
    const valorBusqueda = inputElement.value.trim();
    const servicioCod = valorBusqueda.split(' - ')[0].trim();
    const rowId = row.id.split('-')[1];

    // ANTES: Buscaba en array de arrays: listaServicios.slice(1).find(s => ...)
    // AHORA:
    const servicio = listaServicios.find(s => 
        String(s.ID_servicios).trim().toUpperCase() === servicioCod.toUpperCase()
    );

    const descInput = row.querySelector(`input[name="descripcion_${rowId}"]`);
    const codHidden = row.querySelector(`input[name="Servicio_cod_${rowId}"]`);

    if (servicio) {
        // 1. Rellena el ID y la Descripción
        codHidden.value = servicio.ID_servicios;
        descInput.value = servicio.Nombre_Servicio;

        // 2. Limpia los campos que el usuario debe llenar (porque ya no están en 'Servicios')
        const selectores = optimizarSelectoresFila(row, rowId);
        selectores.precio.value = '0.00';
        selectores.hMinNum.value = '0';
        selectores.movilizacion.value = '0.00';

        // 3. Poner foco en el siguiente campo (Cantidad o Precio)
        selectores.cantidad.focus();

    } else {
        // Limpia todo si el servicio no es válido
        alert("Servicio no encontrado. Por favor, ingrese un código válido de la lista.");
        codHidden.value = '';
        descInput.value = '';
        const selectores = optimizarSelectoresFila(row, rowId);
        selectores.precio.value = '0.00';
        selectores.hMinNum.value = '0';
        selectores.movilizacion.value = '0.00';
    }
}

        function manejarUnidadMedida(selectElement) {
            const row = selectElement.closest('tr');
            const rowId = row.id.split('-')[1];
            const undMedida = selectElement.value.toUpperCase().trim();
            const selectores = optimizarSelectoresFila(row, rowId);
            const necesitaTiempo = undMedida === 'HORAS' || undMedida === 'DÍAS';

            if (undMedida === 'SERVICIO') {
                selectores.cantidad.value = 1;
                selectores.cantidad.setAttribute('readonly', true);
            } else {
                selectores.cantidad.removeAttribute('readonly');
            }

            if (necesitaTiempo) {
                selectores.undHMinSelect.removeAttribute('disabled');
                selectores.hMinNum.removeAttribute('disabled');
            } else {
                selectores.undHMinSelect.setAttribute('disabled', true);
                selectores.hMinNum.setAttribute('disabled', true);
                selectores.diasDisplay.textContent = '0';
                selectores.diasHidden.value = '0';
                selectores.hMinNum.value = '0';
                if(selectores.undHMinSelect.options.length > 0) {
                    selectores.undHMinSelect.value = selectores.undHMinSelect.options[0].value;
                } 
            }
            calcularTodo(selectElement); 
        }

        function calcularTodo(inputOrSelectElement) {
            const row = inputOrSelectElement.closest('tr');
            const rowId = row.id.split('-')[1];
            const selectores = optimizarSelectoresFila(row, rowId);
            
            const cantidad = parseFloat(selectores.cantidad.value) || 0;
            const undMedida = selectores.undMedida.value.toUpperCase().trim();
            const hMinNum = parseFloat(selectores.hMinNum.value) || 0;
            const precio = parseFloat(selectores.precio.value) || 0;
            const movilizacion = parseFloat(selectores.movilizacion.value) || 0;

            let diasCotizados = 0;
            const necesitaTiempo = undMedida === 'HORAS' || undMedida === 'DÍAS';
            const hMinUnd = (necesitaTiempo && selectores.undHMinSelect && !selectores.undHMinSelect.disabled) ? 
                selectores.undHMinSelect.value.toUpperCase().trim() : ""; 

            if (necesitaTiempo) {
                if (undMedida === 'DÍAS') {
                    diasCotizados = cantidad;
                } else if (undMedida === 'HORAS') {
                    let factorDias = 0;
                    switch (hMinUnd) {
                        case 'MENSUAL': factorDias = estadoApp.configuracion.VALOR_DIAS_MES; break;
                        case 'SEMANAL': factorDias = 7; break;
                        case 'DIARIAS': factorDias = 1; break;
                    }
                    if (hMinNum > 0 && factorDias > 0) {
                        diasCotizados = (cantidad / hMinNum) * factorDias;
                    } else {
                        diasCotizados = cantidad / estadoApp.configuracion.VALOR_HORAS_DIA; 
                    }
                }
            } 

            const subtotal = (cantidad * precio) + movilizacion;
            selectores.diasDisplay.textContent = Math.ceil(diasCotizados).toFixed(0);
            selectores.diasHidden.value = Math.ceil(diasCotizados).toFixed(0);
            selectores.subtotalDisplay.textContent = subtotal.toFixed(2);
            selectores.subtotalHidden.value = subtotal.toFixed(2);
            calcularTotalGeneral();
        }

        function calcularTotalGeneral() {
            let total = 0;
            document.querySelectorAll('input[name^="subtotal_linea_"]').forEach(input => {
                total += parseFloat(input.value) || 0;
            });
            document.getElementById('totalCotizacion').textContent = total.toFixed(2);
            document.getElementById('total').value = total.toFixed(2);
        }

        function eliminarLineaDetalle(id) {
            if (!puedeRealizarAccion('eliminar línea')) return;
            const row = document.getElementById(id);
            if (row) {
                row.remove();
                calcularTotalGeneral();
            }
        }

        // =============================================================================
        // FUNCIÓN PRINCIPAL DE GUARDADO
        // =============================================================================
        function guardarRegistroCotizacion() {
            if (!puedeRealizarAccion('guardar cotización')) return;
            
            console.log("💾 INICIANDO GUARDADO...");
            
            // Validar formulario
            const errores = validarFormulario();
            if (errores.length > 0) {
                console.error("❌ ERRORES DE VALIDACIÓN:", errores);
                mostrarErroresValidacion(errores);
                return;
            }
            
            // Verificar que hay líneas válidas
            const lineas = obtenerLineasDetalleCorregidas();
            if (lineas.length === 0) {
                mostrarError("Debe agregar al menos un servicio válido", 'guardar');
                return;
            }
            
            mostrarLoading(true, "Guardando cotización...");
            
            try {
                const datos = prepararDatosEnvio();
                
                console.log("🚀 ENVIANDO DATOS AL SERVIDOR...");
                console.log("📊 Resumen datos:", {
                    empresa: datos.Empresa,
                    cliente: datos.Cliente,
                    moneda: datos.Moneda,
                    direccion: datos.Direccion,
                    contacto: datos.Contacto,
                    ejecutivo: datos.Ejecutivo,
                    totalLineas: datos.Lineas.length,
                    total: datos.Total
                });
                
                google.script.run
                    .withSuccessHandler(handleExitoGuardado)
                    .withFailureHandler(handleErrorGuardado)
                    .guardarCotizacion(datos);
                    
            } catch (error) {
                console.error("💥 ERROR AL PREPARAR DATOS:", error);
                handleErrorGuardado(error);
            }
        }

        /**
 * FUNCIÓN COMPLETA CORREGIDA para que copies y pegues:
 */
function prepararDatosEnvio() {
    console.log("📦 PREPARANDO DATOS PARA ENVÍO...");
    const datos = {};
    
    function obtenerValorElemento(id) {
        const elemento = document.getElementById(id);
        if (!elemento) {
            throw new Error(`Error Interno: No se encontró el elemento con id "${id}" en el HTML.`);
        }
        console.log(`✅ ${id}: "${elemento.value}"`);
        return elemento.value;
    }

    try {
        datos.Empresa = obtenerValorElemento('empresa');
        datos.RUC = obtenerValorElemento('rucCliente');
        datos.Cliente = obtenerValorElemento('clienteNombre');
        datos.Moneda = obtenerValorElemento('moneda');
        
        // ⭐ CORRECCIÓN CRÍTICA AQUÍ ⭐
        datos.Forma_De_Pago = obtenerValorElemento('formaPago');
        
        datos.Direccion = obtenerValorElemento('direccion');
        datos.Turno = obtenerValorElemento('turno');
        datos.Ejecutivo = obtenerValorElemento('ejecutivo');
        datos.Contacto = obtenerValorElemento('contacto');
        datos.Estado = obtenerValorElemento('estadoCotizacion');
        datos.fechaEjecucion = obtenerValorElemento('fechaEjecucion');
        datos.plantillaNotas = obtenerValorElemento('plantillaNotas');
        datos.aclaracionesServicio = obtenerValorElemento('aclaracionesServicio');
        
        if (estadoApp.esEdicion) {
            datos.numPedido = estadoApp.numPedidoEdicion;
        }

        datos.Lineas = obtenerLineasDetalleCorregidas();
        datos.Total = obtenerValorElemento('total') || '0.00';

        console.log("📊 RESUMEN DE DATOS A ENVIAR:");
        console.log(`   • Empresa: ${datos.Empresa}`);
        console.log(`   • Cliente: ${datos.Cliente} (${datos.RUC})`);
        console.log(`   • Forma de Pago: ${datos.Forma_De_Pago}`); // ⭐ Ahora mostrará el nombre correcto
        console.log(`   • Total Líneas: ${datos.Lineas.length}`);
        console.log(`   • Monto Total: ${datos.Total}`);
        
        return datos;

    } catch (error) {
        console.error("💥 ERROR DURANTE LA PREPARACIÓN DE DATOS:", error.message);
        throw error;
    }
}

/**
 * OBTENER LÍNEAS DE DETALLE (v2 - MEJORADO CON VALIDACIÓN)
 */
function obtenerLineasDetalleCorregidas() {
    const lineas = [];
    const filasDetalle = document.querySelectorAll('#cuerpoDetalle tr');
    
    console.log(`🔍 PROCESANDO ${filasDetalle.length} LÍNEAS...`);
    
    filasDetalle.forEach((row, index) => {
        const rowId = row.id.split('-')[1];
        
        // Obtener elementos
        const codHidden = row.querySelector(`input[name="Servicio_cod_${rowId}"]`);
        const descInput = row.querySelector(`input[name="descripcion_${rowId}"]`);
        const cantidadInput = row.querySelector(`input[name="cantidad_${rowId}"]`);
        const undMedidaSelect = row.querySelector(`select[name="und_medida_${rowId}"]`);
        const precioInput = row.querySelector(`input[name="precio_unitario_${rowId}"]`);
        const movilizacionInput = row.querySelector(`input[name="movilizacion_${rowId}"]`);
        
        // Validar elementos existentes
        if (!codHidden || !cantidadInput || !precioInput) {
            console.warn(`⚠️ Línea ${index + 1} - Elementos no encontrados`);
            return;
        }
        
        const cod = codHidden.value.trim();
        const cantidad = parseFloat(cantidadInput.value);
        const precio = parseFloat(precioInput.value);
        
        // ✅ VALIDACIÓN MEJORADA
        if (!cod) {
            console.warn(`⚠️ Línea ${index + 1} omitida - Código vacío`);
            return;
        }
        
        if (isNaN(cantidad) || cantidad <= 0) {
            console.warn(`⚠️ Línea ${index + 1} omitida - Cantidad inválida: ${cantidadInput.value}`);
            return;
        }
        
        if (isNaN(precio)) {
            console.warn(`⚠️ Línea ${index + 1} omitida - Precio inválido: ${precioInput.value}`);
            return;
        }
        
        const linea = {
            cod: cod,
            descripcion: descInput ? descInput.value.trim() : '',
            cantidad: cantidad,
            und_medida: undMedidaSelect ? undMedidaSelect.value : 'HORAS',
            precio: precio,
            movilizacion: movilizacionInput ? parseFloat(movilizacionInput.value) || 0 : 0,
            und_horas_minimas: row.querySelector(`select[name="und_horas_minimas_${rowId}"]`)?.value || '',
            horas_minimas_num: row.querySelector(`input[name="horas_minimas_num_${rowId}"]`) ? 
                parseFloat(row.querySelector(`input[name="horas_minimas_num_${rowId}"]`).value) || 0 : 0,
            hora_segun: row.querySelector(`select[name="hora_segun_${rowId}"]`)?.value || '',
            dias_cotizados: row.querySelector(`input[name="dias_cotizados_val_${rowId}"]`) ? 
                parseFloat(row.querySelector(`input[name="dias_cotizados_val_${rowId}"]`).value) || 0 : 0,
            subtotal: (cantidad * precio) + (movilizacionInput ? parseFloat(movilizacionInput.value) || 0 : 0)
        };
        
        console.log(`✅ Línea ${index + 1} procesada:`, {
            cod: linea.cod,
            cant: linea.cantidad,
            precio: linea.precio,
            subtotal: linea.subtotal
        });
        
        lineas.push(linea);
    });
    
    console.log(`📋 TOTAL LÍNEAS VÁLIDAS: ${lineas.length}`);
    return lineas;
}
/**
 * OBTENER LÍNEAS DE DETALLE (v2 - MEJORADO CON VALIDACIÓN)
 */
function obtenerLineasDetalleCorregidas() {
    const lineas = [];
    const filasDetalle = document.querySelectorAll('#cuerpoDetalle tr');
    
    console.log(`🔍 PROCESANDO ${filasDetalle.length} LÍNEAS...`);
    
    filasDetalle.forEach((row, index) => {
        const rowId = row.id.split('-')[1];
        
        // Obtener elementos
        const codHidden = row.querySelector(`input[name="Servicio_cod_${rowId}"]`);
        const descInput = row.querySelector(`input[name="descripcion_${rowId}"]`);
        const cantidadInput = row.querySelector(`input[name="cantidad_${rowId}"]`);
        const undMedidaSelect = row.querySelector(`select[name="und_medida_${rowId}"]`);
        const precioInput = row.querySelector(`input[name="precio_unitario_${rowId}"]`);
        const movilizacionInput = row.querySelector(`input[name="movilizacion_${rowId}"]`);
        
        // Validar elementos existentes
        if (!codHidden || !cantidadInput || !precioInput) {
            console.warn(`⚠️ Línea ${index + 1} - Elementos no encontrados`);
            return;
        }
        
        const cod = codHidden.value.trim();
        const cantidad = parseFloat(cantidadInput.value);
        const precio = parseFloat(precioInput.value);
        
        // ✅ VALIDACIÓN MEJORADA
        if (!cod) {
            console.warn(`⚠️ Línea ${index + 1} omitida - Código vacío`);
            return;
        }
        
        if (isNaN(cantidad) || cantidad <= 0) {
            console.warn(`⚠️ Línea ${index + 1} omitida - Cantidad inválida: ${cantidadInput.value}`);
            return;
        }
        
        if (isNaN(precio)) {
            console.warn(`⚠️ Línea ${index + 1} omitida - Precio inválido: ${precioInput.value}`);
            return;
        }
        
        const linea = {
            cod: cod,
            descripcion: descInput ? descInput.value.trim() : '',
            cantidad: cantidad,
            und_medida: undMedidaSelect ? undMedidaSelect.value : 'HORAS',
            precio: precio,
            movilizacion: movilizacionInput ? parseFloat(movilizacionInput.value) || 0 : 0,
            und_horas_minimas: row.querySelector(`select[name="und_horas_minimas_${rowId}"]`)?.value || '',
            horas_minimas_num: row.querySelector(`input[name="horas_minimas_num_${rowId}"]`) ? 
                parseFloat(row.querySelector(`input[name="horas_minimas_num_${rowId}"]`).value) || 0 : 0,
            hora_segun: row.querySelector(`select[name="hora_segun_${rowId}"]`)?.value || '',
            dias_cotizados: row.querySelector(`input[name="dias_cotizados_val_${rowId}"]`) ? 
                parseFloat(row.querySelector(`input[name="dias_cotizados_val_${rowId}"]`).value) || 0 : 0,
            subtotal: (cantidad * precio) + (movilizacionInput ? parseFloat(movilizacionInput.value) || 0 : 0)
        };
        
        console.log(`✅ Línea ${index + 1} procesada:`, {
            cod: linea.cod,
            cant: linea.cantidad,
            precio: linea.precio,
            subtotal: linea.subtotal
        });
        
        lineas.push(linea);
    });
    
    console.log(`📋 TOTAL LÍNEAS VÁLIDAS: ${lineas.length}`);
    return lineas;
}

        function obtenerLineasDetalleCorregidas() {
            const lineas = [];
            const filasDetalle = document.querySelectorAll('#cuerpoDetalle tr');
            
            console.log("🔍 PROCESANDO " + filasDetalle.length + " LÍNEAS...");
            
            filasDetalle.forEach((row, index) => {
                const rowId = row.id.split('-')[1];
                
                // Obtener elementos
                const codHidden = row.querySelector(`input[name="Servicio_cod_${rowId}"]`);
                const descInput = row.querySelector(`input[name="descripcion_${rowId}"]`);
                const cantidadInput = row.querySelector(`input[name="cantidad_${rowId}"]`);
                const undMedidaSelect = row.querySelector(`select[name="und_medida_${rowId}"]`);
                const precioInput = row.querySelector(`input[name="precio_unitario_${rowId}"]`);
                const movilizacionInput = row.querySelector(`input[name="movilizacion_${rowId}"]`);
                
                // Validar elementos existentes
                if (!codHidden || !cantidadInput || !precioInput) {
                    console.warn(`⚠️ Línea ${index + 1} - Elementos no encontrados`);
                    return;
                }
                
                const cod = codHidden.value.trim();
                const cantidad = parseFloat(cantidadInput.value);
                const precio = parseFloat(precioInput.value);
                
                if (!cod || isNaN(cantidad) || cantidad <= 0) {
                    console.warn(`⚠️ Línea ${index + 1} omitida - Datos inválidos`);
                    return;
                }
                
                const linea = {
                    cod: cod,
                    descripcion: descInput ? descInput.value.trim() : '',
                    cantidad: cantidad,
                    und_medida: undMedidaSelect ? undMedidaSelect.value : 'HORAS',
                    precio: isNaN(precio) ? 0 : precio,
                    movilizacion: movilizacionInput ? parseFloat(movilizacionInput.value) || 0 : 0,
                    und_horas_minimas: row.querySelector(`select[name="und_horas_minimas_${rowId}"]`)?.value || '',
                    horas_minimas_num: row.querySelector(`input[name="horas_minimas_num_${rowId}"]`) ? parseFloat(row.querySelector(`input[name="horas_minimas_num_${rowId}"]`).value) || 0 : 0,
                    hora_segun: row.querySelector(`select[name="hora_segun_${rowId}"]`)?.value || '',
                    dias_cotizados: row.querySelector(`input[name="dias_cotizados_val_${rowId}"]`) ? parseFloat(row.querySelector(`input[name="dias_cotizados_val_${rowId}"]`).value) || 0 : 0,
                    subtotal: (cantidad * (isNaN(precio) ? 0 : precio)) + (movilizacionInput ? parseFloat(movilizacionInput.value) || 0 : 0)
                };
                
                console.log(`✅ Línea ${index + 1} procesada:`, linea);
                lineas.push(linea);
            });
            
            console.log("📋 TOTAL LÍNEAS VÁLIDAS: " + lineas.length);
            return lineas;
        }

        function handleExitoGuardado(respuesta) {
            mostrarLoading(false);
            mostrarExito(respuesta);
        }

        function handleErrorGuardado(error) {
            mostrarLoading(false);
            mostrarError(error, 'guardar cotización');
        }

        // =============================================================================
        // FUNCIONES DE EDICIÓN - CORREGIDAS
        // =============================================================================
        function cargarPedidoParaEdicion(numPedido) {
            console.log("🔄 INICIANDO CARGA DE EDICIÓN para:", numPedido);
            
            mostrarLoading(true, `Cargando pedido ${numPedido}...`);
            
            google.script.run
                .withSuccessHandler(function(datos) {
                    console.log("✅ RESPUESTA DEL SERVIDOR RECIBIDA");
                    
                    try {
                        // El backend ya debería devolver un objeto JS si es exitoso.
                        if (datos && datos.success !== false) {
                            console.log("🚀 Datos VÁLIDOS - Procediendo a precargar formulario");
                            precargarFormulario(datos);
                        } else {
                            const mensajeError = datos?.message || "No se pudieron cargar los datos del pedido";
                            console.error("❌ ERROR EN DATOS:", mensajeError);
                            mostrarError(mensajeError, 'cargar pedido');
                        }
                    } catch (parseError) {
                        console.error("❌ ERROR PROCESANDO RESPUESTA:", parseError);
                        mostrarError("Error procesando los datos del servidor", 'procesar datos');
                    }
                    mostrarLoading(false);
                })
                .withFailureHandler(function(error) {
                    console.error("❌ ERROR DE CONEXIÓN:", error);
                    mostrarLoading(false);
                    mostrarError(error, 'conexión al cargar pedido');
                })
                .obtenerPedidoParaEdicion(numPedido);
        }

function precargarFormulario(datos) {
            console.log("🎨 INICIANDO PRECARGA DE FORMULARIO (v2 - Anidado)");
            
            // Tus logs de depuración
            console.log("📍 Datos recibidos - Dirección:", datos.Direccion);
            console.log("📞 Datos recibidos - Contacto:", datos.Contacto);
            console.log("👨‍💼 Datos recibidos - Ejecutivo:", datos.Ejecutivo);

            // Pequeño delay para asegurar que el DOM esté listo
            setTimeout(() => {
                try {
                    // 1. Rellenar campos generales
                    console.log("📝 Llenando campos generales...");
                    
                    const campos = {
                        'fechaEjecucion': datos.fechaEjecucion || '', 
                        'empresa': datos.Empresa || '',
                        'rucCliente': datos.RUC || '',
                        'clienteNombre': datos.Cliente, // <-- Leído desde la consulta anidada
                        'formaPago': datos.Forma_De_Pago || '',
                        'direccion': datos.Direccion || '', // <-- Leído de Pedidos o Clientes.Direccion_Fiscal
                        'ejecutivo': datos.Ejecutivo || '',
                    };
                    
                    Object.keys(campos).forEach(campoId => {
                        const elemento = document.getElementById(campoId);
                        if (elemento) {
                            elemento.value = campos[campoId];
                            console.log(`✅ Campo ${campoId}: "${campos[campoId]}"`);
                        } else {
                            console.warn(`⚠️ Campo no encontrado: ${campoId}`);
                        }
                    });

                    // --- LÓGICA DE SELECTS (Estado, Turno, Moneda, Ejecutivo) ---

                    // Lógica de ESTADO
                    const estadoSelect = document.getElementById('estadoCotizacion');
                    if (estadoSelect) {
                         const estadoASeleccionar = datos.Estado || 'COTIZACION';
                         const opcionEncontrada = [...estadoSelect.options].find(option => 
                             option.value.toLowerCase() === estadoASeleccionar.toLowerCase()
                         );
                         if (opcionEncontrada) {
                              estadoSelect.value = opcionEncontrada.value;
                              console.log(`✅ Estado seleccionado: "${estadoSelect.value}"`);
                         } else if (estadoSelect.options.length > 1) {
                              estadoSelect.value = estadoSelect.options[1].value;
                              console.warn(`⚠️ Estado "${estadoASeleccionar}" no encontrado, usando por defecto.`);
                         }
                    }

                    // Lógica de TURNO
                    const turnoSelect = document.getElementById('turno');
                    if (turnoSelect && datos.Turno) {
                        const mapeoTurnos = { 'DIURNO': 'Diurno', 'NOCTURNO': 'Nocturno', 'DOBLE TURNO': 'Doble Turno', 'DOBLETURNO': 'Doble Turno', 'DOBLE': 'Doble Turno', 'DIURNA': 'Diurno', 'NOCTURNA': 'Nocturno' };
                        const turnoRecibido = String(datos.Turno).toUpperCase().trim();
                        const turnoNormalizado = mapeoTurnos[turnoRecibido] || (turnoRecibido.charAt(0) + turnoRecibido.slice(1).toLowerCase());
                        
                        let opcionTurno = [...turnoSelect.options].find(opt => opt.value.toLowerCase() === turnoNormalizado.toLowerCase());
                        if (opcionTurno) {
                           turnoSelect.value = opcionTurno.value;
                           console.log(`✅ Turno seleccionado: "${opcionTurno.value}"`);
                        } else { 
                           console.warn(`⚠️ No se pudo encontrar turno: "${datos.Turno}"`); 
                        }
                    }

                    // Lógica de MONEDA
                    let monedaValue = datos.Moneda || 'Soles';
                    if (monedaValue === 'SOL') monedaValue = 'Soles';
                    if (monedaValue === 'USD') monedaValue = 'Dolares';
                    const monedaSelect = document.getElementById('moneda');
                    if (monedaSelect) {
                        monedaSelect.value = monedaValue;
                        console.log(`✅ Moneda: "${monedaValue}"`);
                    }
                    
                    // Lógica de EJECUTIVO
                    const ejecutivoSelect = document.getElementById('ejecutivo');
                    if (ejecutivoSelect && datos.Ejecutivo) {
                        const ejecutivoNormalizado = String(datos.Ejecutivo).toUpperCase().trim();
                        const opcionEjecutivo = [...ejecutivoSelect.options].find(opt => opt.value.toUpperCase() === ejecutivoNormalizado);
                        if (opcionEjecutivo) { 
                           ejecutivoSelect.value = opcionEjecutivo.value;
                           console.log(`✅ Ejecutivo seleccionado: "${opcionEjecutivo.value}"`); 
                        } else { 
                           console.warn(`⚠️ No se pudo encontrar ejecutivo: "${datos.Ejecutivo}"`); 
                        }
                    }

                    // --- LÓGICA DE CONTACTOS ---
                    if (datos.RUC) {
                        console.log("📞 Cargando lista de contactos para RUC:", datos.RUC);
                        google.script.run
                            .withSuccessHandler(function(contactosData) {
                                console.log("✅ Lista de contactos recibida:", contactosData);
                                listaContactos = contactosData || [];
                                llenarDatalistContactos(); // Llena el select con TODOS los contactos
                                
                                // Ahora selecciona el contacto específico que vino de la carga
                                if (datos.Contacto) {
                                    seleccionarContactoEspecifico(datos.Contacto);
                                }
                            })
                            .withFailureHandler(function(error) {
                                console.error("❌ Error cargando lista de contactos en modo edición:", error);
                                listaContactos = [];
                                llenarDatalistContactos();
                            })
                            .getContactosParaComercial(datos.RUC);
                    } else {
                         listaContactos = [];
                         llenarDatalistContactos();
                    }

                    // Llenar notas (si las añadiste a la tabla Pedidos)
                    document.getElementById('plantillaNotas').value = datos.plantillaNotas || '';
                    document.getElementById('aclaracionesServicio').value = datos.aclaracionesServicio || '';
                    
                    // Deshabilitar RUC (Sin cambios)
                    document.getElementById('rucCliente').readOnly = true;
                    document.getElementById('rucCliente').style.backgroundColor = '#f5f5f5';

                    console.log("✅✅✅ TODOS los campos generales llenados.");

                    // 2. Limpiar tabla (Sin cambios)
                    const tbody = document.getElementById('cuerpoDetalle');
                    if (tbody) {
                        tbody.innerHTML = '';
                        estadoApp.lineaCounter = 0;
                        console.log("✅ Tabla limpiada");
                    }

                    // 3. Cargar líneas de servicios
                    const servicios = datos.Lineas || [];
                    console.log("📦 Cargando", servicios.length, "líneas de servicios");
                    
                    if (servicios.length > 0) {
                        servicios.forEach((servicio, index) => {
                            const lineaData = {
                                COD: servicio.cod || '',
                                DESC: servicio.descripcion || '', // <-- ¡Esto ahora viene resuelto!
                                UND: servicio.und_medida || 'HORAS', 
                                CANTIDAD: servicio.cantidad || 1, 
                                UND_H_MIN: servicio.und_horas_minimas || '', 
                                DIAS_COT: servicio.dias_cotizados || 0, 
                                H_MIN_NUM: servicio.horas_minimas_num || 0, 
                                H_SEGUN: servicio.hora_segun || '', 
                                MOV: servicio.movilizacion || 0, 
                                PRECIO: servicio.precio || 0, 
                                SUB: servicio.subtotal || 0
                            };
                            agregarLineaDetalle(lineaData);
                        });
                    } else {
                        console.warn("⚠️ No hay líneas de servicio para cargar");
                        agregarLineaDetalle(); // Agregar una línea vacía si no hay ninguna
                    }
                    
                    // 4. Calcular total y finalizar
                    setTimeout(() => {
                        calcularTotalGeneral();
                        console.log("🎉🎉🎉 PRECARGA (v2) COMPLETADA EXITOSAMENTE 🎉🎉🎉");
                    }, 500); // Delay extra para renderizado
                    
                } catch (error) {
                    console.error("💥 ERROR en precargarFormulario (v2):", error);
                    console.error("Stack trace:", error.stack);
                    mostrarError(error, 'precargar formulario v2');
                }
            }, 300);
        }

        function formatDateForInput(dateString) {
            if (!dateString) return '';
            try {
                const date = new Date(dateString);
                return date.toISOString().split('T')[0];
            } catch (e) {
                console.error("Error formateando fecha:", e);
                return '';
            }
        }

        // =============================================================================
        // FUNCIONES DE DIAGNÓSTICO
        // =============================================================================
        function debugCompleto() {
            console.log("=== 🐛 DEBUG COMPLETO ===");
            console.log("📊 Estado App:", estadoApp);
            console.log("👥 Clientes cargados:", listaClientes.length);
            console.log("🛠️ Servicios cargados:", listaServicios.length);
            console.log("✏️ Modo edición:", estadoApp.esEdicion);
            console.log("🔢 Pedido edición:", estadoApp.numPedidoEdicion);
            console.log("=== 🐛 FIN DEBUG ===");
        }

        // Ejecutar debug después de la carga
        setTimeout(debugCompleto, 2000);
        
        // Ejecutar verificación después de la carga completa
        setTimeout(verificarDatosCargados, 3000);
        
        /**
 * Función para abrir modal de cotización
 */
function abrirModalCotizacion(ruc, nombreCliente) {
  // Aquí puedes obtener los contactos y direcciones para seleccionar
  const contacto = prompt("Ingrese el nombre del contacto:", "");
  const contactoInfo = prompt("Ingrese email o teléfono del contacto:", "");
  const lugar = prompt("Lugar (ALPAMAYO/GYM/SAN JOSE):", "ALPAMAYO");
  const turno = prompt("Turno:", "MAÑANA");
  const duracion = prompt("Duración:", "1 MES");
  
  if (contacto && contactoInfo && lugar) {
    const datosBase = {
      ruc: ruc,
      nombreCliente: nombreCliente,
      contacto: contacto,
      contactoInfo: contactoInfo,
      lugar: lugar.toUpperCase(),
      turno: turno,
      duracion: duracion,
      items: [] // Aquí podrías agregar un formulario para los items
    };
    
    // Guardar datos temporalmente y abrir editor de items
    sessionStorage.setItem('datosCotizacionTemp', JSON.stringify(datosBase));
    abrirEditorItems();
  }
}

/**
 * Función para generar PDF final
 */
function generarPDFFinal(items) {
  const datosTemp = sessionStorage.getItem('datosCotizacionTemp');
  if (!datosTemp) {
    alert('No hay datos de cotización guardados');
    return;
  }
  
  const datosCotizacion = JSON.parse(datosTemp);
  datosCotizacion.items = items;
  
  // Mostrar loading
  alert('Generando PDF...');
  
  google.script.run
    .withSuccessHandler(function(resultado) {
      if (resultado.success) {
        // Abrir el PDF en nueva pestaña
        window.open(resultado.pdfUrl, '_blank');
        alert('PDF generado exitosamente!');
      } else {
        alert('Error al generar PDF: ' + resultado.error);
      }
    })
    .withFailureHandler(function(error) {
      alert('Error: ' + error.message);
    })
    .generarPDFCotizacion(datosCotizacion);
}

/**
 * Ejemplo de cómo agregar desde tu tabla de clientes
 */
// Modifica el botón en tu tabla para incluir la opción de cotización:
function seleccionarCliente(ruc, nombre) {
  RUC_ACTUAL = ruc;
  NOMBRE_ACTUAL = nombre;
  
  // Podrías agregar un botón adicional o modificar el existente
  const accion = confirm(`Cliente: ${nombre}\nRUC: ${ruc}\n\n¿Qué deseas hacer?\nOK - Ver Detalles\nCancelar - Generar Cotización`);
  
  if (accion) {
    // Código existente para ver detalles
    document.getElementById("tituloDetalleCliente").textContent = `Gestión de Cliente: ${nombre} (RUC: ${ruc})`;
    document.getElementById("clienteRUCDisplay").textContent = ruc;
    document.getElementById("clienteNombreInput").value = nombre;
    
    mostrarModal('modalCliente');
    cargarDetallesCliente();
  } else {
    // Nueva funcionalidad: generar cotización
    abrirModalCotizacion(ruc, nombre);
  }
}
function debugModoEdicion() {
    console.log("=== 🔧 DEBUG MODO EDICIÓN ===");
    console.log("URL completa:", window.location.href);
    console.log("Parámetros URL:", new URLSearchParams(window.location.search).toString());
    console.log("Estado App:", estadoApp);
    console.log("EditId desde URL:", new URLSearchParams(window.location.search).get('editar'));
    console.log("=== FIN DEBUG ===");
}

/**
 * Función auxiliar para escapar caracteres especiales en strings 
 * que se usarán dentro de llamadas a funciones JavaScript en HTML.
 */
function escapeJsString(str) {
    if (!str) return '';
    return str.replace(/\\/g, '\\\\')
              .replace(/'/g, "\\'")
              .replace(/"/g, '\\"')
              .replace(/\n/g, '\\n')
              .replace(/\r/g, '\\r');
}

/**
     * REEMPLAZO v6 de gestionarActa
     * Usa sessionStorage para pasar datos a Acta.html en lugar de parámetros URL largos.
     */
    function gestionarActa(cotNumeroParam, lineaIndex, servicioCod, servicioDesc) {
        console.log(`Gestionando Acta - Parámetros recibidos: COT='${cotNumeroParam}', Línea=${lineaIndex}`);

        let cotNumeroFinal = null;

        // 1. Determinar el número de COT correcto (Sin cambios)
        if (estadoApp.esEdicion && estadoApp.numPedidoEdicion) {
            cotNumeroFinal = estadoApp.numPedidoEdicion;
        } else if (cotNumeroParam && cotNumeroParam !== 'NUEVO_PEDIDO') {
             cotNumeroFinal = cotNumeroParam;
        }

        // 2. Validar si tenemos un número de COT válido (Sin cambios)
        if (!cotNumeroFinal) { /* ... alert ... */ return; }
        if (isNaN(lineaIndex) || lineaIndex < 1) { /* ... alert ... */ return; }

        const btn = document.getElementById(`btnActa-${lineaIndex}`);
        if(btn) { btn.textContent = "Verificando..."; btn.disabled = true; }

        // 3. Llamar al backend para verificar si el acta existe
        google.script.run
            .withSuccessHandler(function(actaId) {
                if(btn) { btn.textContent = "Acta..."; btn.disabled = false; }
                
                const urlBase = '<?!= getScriptUrl() ?>';
                let urlDestino = '';
                
                // --- INICIO: Guardar en sessionStorage y construir URL simple ---
                let datosParaActa = {};

                if (actaId) { // Editar
                    console.log("Acta encontrada:", actaId, "- Preparando para edición.");
                    datosParaActa = {
                        modo: 'editar',
                        actaId: actaId
                        // No necesitamos pasar más datos, Acta.html los cargará por ID
                    };
                    // URL solo necesita indicar la página y el modo (opcionalmente el ID para claridad)
                     urlDestino = `${urlBase}?page=Acta&modo=editar&ref=${actaId}`; 
                     
                } else { // Crear
                    console.log("Acta no encontrada - Preparando para creación.");
                    datosParaActa = {
                        modo: 'crear',
                        cot: cotNumeroFinal, 
                        linea: lineaIndex,
                        // Pasamos cod y desc aquí para que Acta.html los tenga
                        cod: servicioCod || '', 
                        desc: servicioDesc || '' 
                    };
                     // URL solo necesita indicar la página y el modo
                     urlDestino = `${urlBase}?page=Acta&modo=crear`;
                }
                
                // Guardar los datos en sessionStorage ANTES de abrir la ventana
                try {
                    sessionStorage.setItem('datosParaActa', JSON.stringify(datosParaActa));
                    console.log("Datos guardados en sessionStorage:", datosParaActa);
                    
                    // Abrir la ventana con la URL simple
                    console.log("Abriendo URL:", urlDestino);
                    window.open(urlDestino, '_blank'); 

                } catch (e) {
                     console.error("Error al guardar en sessionStorage:", e);
                     mostrarError("No se pudo preparar la información para la página del Acta.", 'sessionStorage');
                     if(btn) { btn.textContent = "Acta..."; btn.disabled = false; } // Restaurar botón si falla aquí
                }
                // --- FIN: Guardar en sessionStorage ---

            })
            .withFailureHandler(function(error) {
                 if(btn) { btn.textContent = "Acta..."; btn.disabled = false; }
                 mostrarError(error, 'verificar acta');
            })
            .verificarActaExistente(cotNumeroFinal, lineaIndex); 
    }
// Ejecutar después de la carga
setTimeout(debugModoEdicion, 1000);
    </script>
</body>
</html>
