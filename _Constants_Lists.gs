// ====================================================
// === CONFIGURACIÓN GLOBAL Y CONSTANTES DE DRIVE ===
// ====================================================

const HOJA_ID_PRINCIPAL = "15qfA3idaLkyhvFwAeEZQo6L9BudBBfgnV8DFrs1qV6Y"; 
// Constantes para las pestañas de plantillas
const HOJA_PLANTILLA_ALPAMAYO = 'COT_ALP';
const HOJA_PLANTILLA_GYM = 'COT_GYM';
const HOJA_PLANTILLA_SANJOSE = 'COT_GSJ';

// --- IDs DE CARPETAS DE DESTINO ---
const FOLDER_ID_CARMEN = "1h4LZiA9Iwx54jHqOyHqFvDJykGipM6YO";
const FOLDER_ID_GYM     = "1ODozAGz2AeDmFGk_rT_axt7bhG6SdLv2"; 
const FOLDER_ID_SJ      = "1bO8-_ZWM2nFfc1jl4PDYrf6aTnKqp7s4";
const FOLDER_ID_ALP     = "1ZrPgv6mfTZk5r4GE49tAlc9C559b_M93";

// --- IDs DE ARCHIVOS DE PLANTILLA ---
const ID_PLANTILLA_FILE_ALP = "15qfA3idaLkyhvFwAeEZQo6L9BudBBfgnV8DFrs1qV6Y";
const ID_PLANTILLA_FILE_GYM = "15qfA3idaLkyhvFwAeEZQo6L9BudBBfgnV8DFrs1qV6Y";
const ID_PLANTILLA_FILE_SJ  = "15qfA3idaLkyhvFwAeEZQo6L9BudBBfgnV8DFrs1qV6Y";

// ====================================================
// === CONSTANTES DEL MÓDULO DE ACTAS Y VALORIZACIÓN ===
// ====================================================

// --- IDs DE PLANTILLAS DE HOJAS DE CÁLCULO (SHEETS) ---
const ID_PLANTILLA_ACTA_ALP = "ACTA_ALP";
const ID_PLANTILLA_ACTA_GYM = "ACTA_GYM";
const ID_PLANTILLA_ACTA_GSJ = "ACTA_GSJ";

const HOJA_ACTAS = "Actas"; // Aún se usa para el registro en sheets
// Las constantes de Valorización (HOJA_VALORIZACIONES, etc.) se eliminan ya que la lógica migró a Supabase

// ====================================================
// === LISTAS ESTÁTICAS COMO CONSTANTES GLOBALES ===
// ====================================================

const LISTA_UND_MEDIDA = ["HORAS", "DÍAS", "SERVICIO", "GLOBAL", "UND"];
const LISTA_HORAS_SEGUN = ["HORÓMETRO", "INICIO/FIN", "PARTE DIARIO"];
const LISTA_TURNOS = ["Diurno", "Nocturno", "Doble Turno"];
const LISTA_EMPRESAS = ["ALPAMAYO", "SAN JOSE", "GYM"];
const LISTA_ESTADOS_COT = ["Cotización", "Pedido", "Cancelado", "Finalizado"];
const LISTA_FORMAS_PAGO = [
    "CONTADO", "50% de adelanto", "CREDITO. 30 DIAS", "CREDITO. 15 DIAS", 
    "CREDITO. 45 DIAS", "CREDITO. 60 DIAS", "CREDITO. 90 DIAS"
];
const LISTA_HORAS_MINIMAS_UND = ["Mensual", "Diarias", "Semanal"]; 

// ====================================================
// === CONSTANTES DE NEGOCIO Y OT/MAPEO ===
// ====================================================

const VALOR_DIAS_MES = 30;
const VALOR_HORAS_DIA = 24; 
const LISTA_EJECUTIVOS = ['ANTHONY', 'CARMEN', 'RENATO', 'JOSUE', 'PEDRO'];

const MAPA_NOMBRES_EMPRESAS = {
    "ALPAMAYO": "GRUAS ALPAMAYO SAC", 
    "SAN JOSE": "GRUAS SAN JOSE PERU SAC", 
    "GYM": "GRUAS Y MONTACARGAS SAN JOSE SAC" 
};
