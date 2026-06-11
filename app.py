import streamlit as st
import pandas as pd
import datetime
import holidays
import os
import json
import io

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
        .custom-card {
            background-color: #ffffff;
            padding: 24px;
            border-radius: 12px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
            margin-bottom: 24px;
            border: 1px solid #e2e8f0;
        }
        h1, h2, h3 { color: #1e293b !important; font-family: 'Inter', sans-serif; }
        section[data-testid="stSidebar"] { background-color: #0f172a !important; }
        section[data-testid="stSidebar"] * { color: #f1f5f9 !important; }
        .stButton>button {
            background-color: #0284c7 !important;
            color: white !important;
            border-radius: 8px !important;
            font-weight: 600 !important;
        }
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
    if "private_key" in claves:
        claves["private_key"] = claves["private_key"].replace("\\n", "\n")
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

def normalizar_texto(val):
    if pd.isna(val): return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"): val_str = val_str[:-2]
    return val_str

@st.cache_data
def cargar_base_escuelas():
    if os.path.exists(EXCEL_ESCUELAS):
        try:
            df = pd.read_excel(EXCEL_ESCUELAS)
            df.columns = df.columns.str.strip().str.upper().str.replace(' ', '_')
            return df.drop_duplicates()
        except Exception: return pd.DataFrame()
    return pd.DataFrame()

def obtener_fechas_ocupadas():
    if usando_google_sheets():
        try:
            hoja = conectar_google_sheets()
            valores = hoja.get_all_values()
            if len(valores) <= 1: return set()
            df = pd.DataFrame(valores[1:], columns=valores[0])
            fechas = []
            for _, row in df.iterrows():
                try:
                    fechas.append(datetime.date(int(float(row['Anio_Reservado'])), int(float(row['Mes_Reservado'])), int(float(row['Dia_Reservado']))))
                except: continue
            return set(fechas)
        except: return set()
    return set()

def guardar_reserva(datos):
    if usando_google_sheets():
        try:
            hoja = conectar_google_sheets()
            hoja.append_row([str(datos.get(k, "")) for k in ["CUE", "Escuela", "Modalidad_Oferta", "Departamento", "Domicilio", "DNI_Director", "Director", "Telefono_Contacto", "Estructura_Declarada", "Detalle_Divisiones_Alumnos", "Total_Alumnos", "Dia_Reservado", "Mes_Reservado", "Anio_Reservado", "Fecha_Registro"]])
        except Exception as e: st.error(f"Error de sincronización: {e}")

config_actual = cargar_configuracion_sistema()
registro_activo = config_actual.get("registro_habilitado", True)
anio_actual = datetime.date.today().year
fechas_ocupadas = obtener_fechas_ocupadas()

# Vista de Directores (Pública)
st.markdown('<h1 style="text-align: center; color: #0284c7;">📅 Sistema de Reserva de Turnos</h1>', unsafe_allow_html=True)

if not registro_activo:
    st.error("⚠️ El sistema se encuentra inhabilitado temporalmente.")
elif st.session_state.reserva_exitosa:
    st.success("🎉 ¡Reserva confirmada exitosamente!")
else:
    # --- FORMULARIO DE RESERVA ---
    with st.container():
        st.markdown('<div class="custom-card">', unsafe_allow_html=True)
        st.subheader("📍 1. Identificación del Establecimiento")
        cue_input = st.text_input("Ingrese CUE:")
        
        st.subheader("👤 2. Datos del Solicitante")
        dni_input = st.text_input("Ingrese DNI:")
        
        st.subheader("📅 3. Selección de Fecha")
        fecha_reserva = st.date_input("Fecha preferida:", min_value=datetime.date.today())
        
        if st.button("Confirmar Reserva"):
            st.info("Procesando reserva...")
            # Aquí iría la lógica de validación e inserción...
        st.markdown('</div>', unsafe_allow_html=True)
