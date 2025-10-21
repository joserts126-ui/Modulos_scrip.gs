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
