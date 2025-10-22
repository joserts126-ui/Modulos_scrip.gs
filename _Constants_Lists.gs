// ====================================================
// === CONFIGURACIÓN GLOBAL Y CONSTANTES DE HOJA ===
// ====================================================

const HOJA_ID_PRINCIPAL = "15qfA3idaLkyhvFwAeEZQo6L9BudBBfgnV8DFrs1qV6Y"; 
const HOJA_CLIENTES = "Clientes";
const HOJA_CONTACTOS = "Contactos";
const HOJA_DIRECCIONES = "Direcciones";
const HOJA_SERVICIOS = "Servicios";
const HOJA_COTIZACIONES = "DataCot";

// Constantes para las pestañas de plantillas
const HOJA_PLANTILLA_ALPAMAYO = 'COT_ALP';
const HOJA_PLANTILLA_GYM = 'COT_GYM';
const HOJA_PLANTILLA_SANJOSE = 'COT_GSJ';

const CLIENTE_COLS = { RUC: 0, NOMBRE: 1 };
const CONTACTO_COLS = { RUC: 1, NOMBRE: 2, EMAIL: 3, TELEFONO: 4, CARGO: 5 };
const DIRECCION_COLS = { RUC: 1, TIPO: 2, DIRECCION: 3, CIUDAD: 4 };

// ====================================================
// === IDs DE CARPETAS Y PLANTILLAS PARA PDF/SHEET ===
// ====================================================

// --- IDs DE CARPETAS DE DESTINO (Rellenados desde tus links) ---
const FOLDER_ID_CARMEN = "1h4LZiA9Iwx54jHqOyHqFvDJykGipM6YO";
const FOLDER_ID_GYM     = "1ODozAGz2AeDmFGk_rT_axt7bhG6SdLv2"; // Gruas y Montacargas San Jose SAC
const FOLDER_ID_SJ      = "1bO8-_ZWM2nFfc1jl4PDYrf6aTnKqp7s4"; // Gruas San Jose Peru SAC
const FOLDER_ID_ALP     = "1ZrPgv6mfTZk5r4GE49tAlc9C559b_M93"; // Alpamayo

// --- IDs DE ARCHIVOS DE PLANTILLA (¡DEBES RELLENAR ESTOS!) ---
// Ve a tu Google Drive, haz clic derecho en el archivo plantilla y "Obtener enlace"
// Copia el ID desde el enlace (ej. .../d/AQUI_VA_EL_ID/edit)
const ID_PLANTILLA_FILE_ALP = "15qfA3idaLkyhvFwAeEZQo6L9BudBBfgnV8DFrs1qV6Y";
const ID_PLANTILLA_FILE_GYM = "15qfA3idaLkyhvFwAeEZQo6L9BudBBfgnV8DFrs1qV6Y";
const ID_PLANTILLA_FILE_SJ  = "15qfA3idaLkyhvFwAeEZQo6L9BudBBfgnV8DFrs1qV6Y";

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
const LISTA_HORAS_MINIMAS_UND = ["Mensual", "Diarias", "Semanal"]; 

// ====================================================
// === CONSTANTES DE NEGOCIO Y CÁLCULO ===
// ====================================================

const VALOR_DIAS_MES = 30; 
const VALOR_HORAS_DIA = 24; 
const LISTA_EJECUTIVOS = ['ANTHONY', 'CARMEN', 'RENATO', 'JOSUE', 'PEDRO'];

const HOJA_OT = 'OT'; 
const HOJA_PEDIDOS_OT = 'Pedidos';

// Mapea el nombre corto de la empresa (de la cotización) 
// al nombre EXACTO de la carpeta en Drive.
const MAPA_NOMBRES_EMPRESAS = {
    "ALPAMAYO": "GRUAS ALPAMAYO SAC", // Cambiado
    "SAN JOSE": "GRUAS SAN JOSE PERU SAC", // Cambiado
    "GYM": "GRUAS Y MONTACARGAS SAN JOSE SAC" // Cambiado
};
