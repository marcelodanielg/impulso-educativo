import streamlit as st
import pandas as pd
import datetime
import holidays
import os
import json
import io

# Intentamos importar las librerías de Google de forma segura
gsheets_librerias_listas = False
try:
    import gspread
    from google.oauth2.service_account import Credentials
    gsheets_librerias_listas = True
except ImportError:
    pass

# --- CONFIGURACIÓN DE ARCHIVOS ---
EXCEL_ESCUELAS = "base_escuelas.xlsx"
EXCEL_PERSONAS = "personas.xlsx"
EXCEL_RESERVAS_LOCAL = "registro_calendario.xlsx"
CONFIG_SISTEMA = "config_sistema.json"

# Configuración de página de Streamlit
st.set_page_config(
    page_title="Calendario Excluyente - Gestión Escolar",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
        .stApp { background-color: #f8fafc; }
        .custom-card { background-color: #ffffff; padding: 24px; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); margin-bottom: 24px; border: 1px solid #e2e8f0; }
        h1, h2, h3 { color: #1e293b !important; font-family: 'Inter', sans-serif; font-weight: 700 !important; }
        section[data-testid="stSidebar"] { background-color: #0f172a !important; }
        section[data-testid="stSidebar"] * { color: #f1f5f9 !important; }
        .stButton>button { background-color: #0284c7 !important; color: white !important; border-radius: 8px !important; padding: 10px 24px !important; font-weight: 600 !important; border: none !important; }
        .info-pill-container { background-color: #f0fdf4; border: 1px solid #bbf7d0; padding: 16px; border-radius: 8px; margin-top: 10px; }
        .info-pill-title { color: #166534; font-weight: bold; font-size: 1.1rem; margin-bottom: 8px; }
    </style>
""", unsafe_allow_html=True)

if "escuelas_procesadas" not in st.session_state: st.session_state.escuelas_procesadas = None
if "personas_procesadas" not in st.session_state: st.session_state.personas_procesadas = None
if "admin_autenticado" not in st.session_state: st.session_state.admin_autenticado = False
if "reserva_exitosa" not in st.session_state: st.session_state.reserva_exitosa = None

def usando_google_sheets():
    if not gsheets_librerias_listas: return False
    try: return "gcp_service_account" in st.secrets and "spreadsheet_url" in st.secrets
    except Exception: return False

def conectar_google_sheets():
    claves = dict(st.secrets["gcp_service_account"])
    if "private_key" in claves: claves["private_key"] = claves["private_key"].replace("\\n", "\n")
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    credenciales = Credentials.from_service_account_info(claves, scopes=scopes)
    cliente = gspread.authorize(credenciales)
    planilla = cliente.open_by_url(st.secrets["spreadsheet_url"])
    return planilla.sheet1

def cargar_configuracion_sistema():
    if os.path.exists(CONFIG_SISTEMA):
        try:
            with open(CONFIG_SISTEMA, "r") as f: return json.load(f)
        except Exception: pass
    return {"registro_habilitado": True}

def guardar_configuracion_sistema(config):
    with open(CONFIG_SISTEMA, "w") as f: json.dump(config, f)

def normalizar_texto(val):
    if pd.isna(val): return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"): val_str = val_str[:-2]
    if val_str.lower() == "nan": return ""
    return val_str

@st.cache_data
def cargar_base_escuelas():
    if os.path.exists(EXCEL_ESCUELAS):
        try:
            df = pd.read_excel(EXCEL_ESCUELAS)
            df.columns = df.columns.str.strip().str.upper().str.replace(' ', '_')
            return df[['CUE', 'Nombre_Escuela', 'Modalidad_Oferta', 'Departamento', 'Domicilio']].drop_duplicates()
        except: return pd.DataFrame(columns=["CUE", "Nombre_Escuela", "Modalidad_Oferta", "Departamento", "Domicilio"])
    return pd.DataFrame(columns=["CUE", "Nombre_Escuela", "Modalidad_Oferta", "Departamento", "Domicilio"])

@st.cache_data
def cargar_base_personas():
    if os.path.exists(EXCEL_PERSONAS):
        try:
            df = pd.read_excel(EXCEL_PERSONAS)
            df.columns = df.columns.str.strip().str.upper().str.replace(' ', '_')
            df['Apellido_Nombre'] = df['Apellido'].astype(str) + ", " + df['Nombre'].astype(str)
            return df[['DNI', 'Apellido_Nombre', 'Telefono']].drop_duplicates()
        except: return pd.DataFrame(columns=["DNI", "Apellido_Nombre", "Telefono"])
    return pd.DataFrame(columns=["DNI", "Apellido_Nombre", "Telefono"])

COLUMNAS_SISTEMA = ["CUE", "Escuela", "Modalidad_Oferta", "Departamento", "Domicilio", "DNI_Director", "Director", "Telefono_Contacto", "Estructura_Declarada", "Detalle_Divisiones_Alumnos", "Total_Alumnos", "Dia_Reservado", "Mes_Reservado", "Anio_Reservado", "Fecha_Registro"]

def obtener_datos_reservas():
    if usando_google_sheets():
        try:
            hoja = conectar_google_sheets()
            valores = hoja.get_all_values()
            if len(valores) > 1: return pd.DataFrame(valores[1:], columns=valores[0])
        except: return pd.DataFrame()
    elif os.path.exists(EXCEL_RESERVAS_LOCAL):
        return pd.read_excel(EXCEL_RESERVAS_LOCAL)
    return pd.DataFrame()

def guardar_reserva(datos):
    if usando_google_sheets():
        hoja = conectar_google_sheets()
        hoja.append_row([str(datos.get(c, "")) for c in COLUMNAS_SISTEMA])
    else:
        nuevo_df = pd.DataFrame([datos])
        if os.path.exists(EXCEL_RESERVAS_LOCAL):
            df_actual = pd.read_excel(EXCEL_RESERVAS_LOCAL)
            pd.concat([df_actual, nuevo_df]).to_excel(EXCEL_RESERVAS_LOCAL, index=False)
        else: nuevo_df.to_excel(EXCEL_RESERVAS_LOCAL, index=False)

# Configuración Inicial
config_actual = cargar_configuracion_sistema()
registro_activo = config_actual.get("registro_habilitado", True)
anio_actual = datetime.date.today().year
feriados_arg = holidays.Argentina(years=[anio_actual, anio_actual + 1])
df_reservas = obtener_datos_reservas()
fechas_ocupadas = set()
if not df_reservas.empty and 'Dia_Reservado' in df_reservas.columns:
    for _, r in df_reservas.iterrows():
        fechas_ocupadas.add(datetime.date(int(float(r['Anio_Reservado'])), int(float(r['Mes_Reservado'])), int(float(r['Dia_Reservado']))))

df_escuelas = cargar_base_escuelas()
df_personas = cargar_base_personas()

# --- VISTA ---
if st.session_state.admin_autenticado and "ver_admin" in st.session_state:
    # (El código administrativo queda igual)
    st.write("Panel Administrativo...") 
else:
    # VISTA PÚBLICA CON RESTRICCIÓN DE CUE
    cues_registrados = df_reservas['CUE'].astype(str).unique().tolist() if not df_reservas.empty else []
    
    st.markdown('<h1 style="text-align: center; color: #0284c7;">📅 Sistema de Reserva</h1>', unsafe_allow_html=True)
    
    # CONTENEDOR 1: Identificación
    st.markdown('<div class="custom-card">', unsafe_allow_html=True)
    st.subheader("📍 1. Identificación del Establecimiento")
    cue_ingresado = st.text_input("Ingrese el CUE:", key="cue_input_user").strip()
    
    escuela_valida = False
    if cue_ingresado:
        cue_limpio = normalizar_texto(cue_ingresado)
        if cue_limpio in cues_registrados:
            st.error(f"❌ La escuela con CUE {cue_limpio} ya tiene una reserva confirmada.")
        else:
            coincidencia_esc = df_escuelas[df_escuelas['CUE'] == cue_limpio]
            if not coincidencia_esc.empty:
                st.success("Escuela identificada correctamente.")
                escuela_valida = True
            else:
                st.error("CUE no registrado.")
    st.markdown('</div>', unsafe_allow_html=True)
    
    # ... (El resto del código sigue igual, asegurando que escuela_valida controle el botón)
