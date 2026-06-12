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

st.set_page_config(
    page_title="Calendario Excluyente - Gestión Escolar",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
        .stApp { background-color: #f8fafc; }
        .custom-card { background-color: #ffffff; padding: 24px; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); margin-bottom: 24px; border: 1px solid #e2e8f0; }
        h1, h2, h3 { color: #1e293b !important; }
        .stButton>button { background-color: #0284c7 !important; color: white !important; }
    </style>
""", unsafe_allow_html=True)

# Inicialización de estado
if "admin_autenticado" not in st.session_state: st.session_state.admin_autenticado = False
if "reserva_exitosa" not in st.session_state: st.session_state.reserva_exitosa = None

# --- FUNCIONES DE BASE DE DATOS ---
def usando_google_sheets():
    return gsheets_librerias_listas and "gcp_service_account" in st.secrets and "spreadsheet_url" in st.secrets

def conectar_google_sheets():
    claves = dict(st.secrets["gcp_service_account"])
    if "private_key" in claves: claves["private_key"] = claves["private_key"].replace("\\n", "\n")
    credenciales = Credentials.from_service_account_info(claves, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
    return gspread.authorize(credenciales).open_by_url(st.secrets["spreadsheet_url"]).sheet1

@st.cache_data
def obtener_fechas_ocupadas():
    """Retorna una lista de objetos date ocupados."""
    fechas = []
    if usando_google_sheets():
        try:
            valores = conectar_google_sheets().get_all_values()
            if len(valores) > 1:
                df = pd.DataFrame(valores[1:], columns=valores[0])
                for _, row in df.iterrows():
                    fechas.append(datetime.date(int(float(row['Anio_Reservado'])), int(float(row['Mes_Reservado'])), int(float(row['Dia_Reservado']))))
        except: pass
    elif os.path.exists(EXCEL_RESERVAS_LOCAL):
        try:
            df = pd.read_excel(EXCEL_RESERVAS_LOCAL)
            for _, row in df.iterrows():
                fechas.append(datetime.date(int(row['Anio_Reservado']), int(row['Mes_Reservado']), int(row['Dia_Reservado'])))
        except: pass
    return fechas

# --- LÓGICA DEL CALENDARIO ---
anio_actual = datetime.date.today().year
feriados_arg = holidays.Argentina(years=[anio_actual, anio_actual + 1])
fechas_ocupadas = obtener_fechas_ocupadas()

# --- VISTA PRINCIPAL ---
# (Se mantiene el flujo original, solo ajustamos el componente de fecha)
# ... [Cargar bases de datos, lógica de autenticación igual al original] ...

# Dentro de la VISTA DE DIRECTORES:
# Se reemplaza la sección de "CONTENEDOR 4" con esta lógica:

# --- Lógica de bloqueo de fechas ---
def es_fecha_disponible(fecha):
    if fecha.weekday() in [5, 6]: return False
    if fecha in feriados_arg: return False
    if fecha in fechas_ocupadas: return False
    return True

st.subheader("📅 4. Selección de Turno Excluyente")
st.write("Seleccione una fecha disponible (lunes a viernes, excluyendo feriados y días ya reservados).")

# Generamos un rango de fechas permitidas para que el usuario elija
fecha_inicio_rango = datetime.date(anio_actual, 8, 1)
fecha_fin_rango = datetime.date(anio_actual, 11, 30)

# El usuario selecciona una fecha
fecha_seleccionada = st.date_input(
    "Seleccione el día:",
    value=None,
    min_value=fecha_inicio_rango,
    max_value=fecha_fin_rango,
    help="Las fechas marcadas en rojo o gris fuera de rango no están disponibles."
)

if fecha_seleccionada:
    if not es_fecha_disponible(fecha_seleccionada):
        st.error(f"🔴 La fecha {fecha_seleccionada.strftime('%d/%m/%Y')} no está disponible (Feriado, Finde o ya reservada).")
        es_valida = False
    else:
        st.success(f"🟢 Fecha {fecha_seleccionada.strftime('%d/%m/%Y')} disponible.")
        es_valida = True
else:
    es_valida = False

# ... [Continuar con el botón de "Confirmar y Registrar" igual que en tu código original] ...

