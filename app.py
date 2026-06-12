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

st.set_page_config(page_title="Calendario Excluyente - Gestión Escolar", page_icon="📅", layout="wide", initial_sidebar_state="collapsed")

# --- ESTILOS ---
st.markdown("""
    <style>
        .stApp { background-color: #f8fafc; }
        .custom-card { background-color: #ffffff; padding: 24px; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); margin-bottom: 24px; border: 1px solid #e2e8f0; }
        h1, h2, h3 { color: #1e293b !important; font-family: 'Inter', sans-serif; }
        section[data-testid="stSidebar"] { background-color: #0f172a !important; }
        section[data-testid="stSidebar"] * { color: #f1f5f9 !important; }
        .stButton>button { background-color: #0284c7 !important; color: white !important; }
        .info-pill-container { background-color: #f0fdf4; border: 1px solid #bbf7d0; padding: 16px; border-radius: 8px; margin-top: 10px; }
    </style>
""", unsafe_allow_html=True)

# --- INICIALIZACIÓN DE ESTADO ---
if "escuelas_procesadas" not in st.session_state: st.session_state.escuelas_procesadas = None
if "personas_procesadas" not in st.session_state: st.session_state.personas_procesadas = None
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

def normalizar_texto(val):
    if pd.isna(val): return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"): val_str = val_str[:-2]
    return "" if val_str.lower() == "nan" else val_str

@st.cache_data
def obtener_fechas_ocupadas():
    fechas = set()
    if usando_google_sheets():
        try:
            hoja = conectar_google_sheets()
            valores = hoja.get_all_values()
            if len(valores) > 1:
                df = pd.DataFrame(valores[1:], columns=valores[0])
                for _, row in df.iterrows():
                    try:
                        fechas.add(datetime.date(int(float(row['Anio_Reservado'])), int(float(row['Mes_Reservado'])), int(float(row['Dia_Reservado']))))
                    except: continue
        except: pass
    elif os.path.exists(EXCEL_RESERVAS_LOCAL):
        try:
            df = pd.read_excel(EXCEL_RESERVAS_LOCAL)
            for _, row in df.iterrows():
                fechas.add(datetime.date(int(row['Anio_Reservado']), int(row['Mes_Reservado']), int(row['Dia_Reservado'])))
        except: pass
    return fechas

# --- CARGA DE DATOS ---
@st.cache_data
def cargar_base_escuelas():
    if os.path.exists(EXCEL_ESCUELAS):
        df = pd.read_excel(EXCEL_ESCUELAS)
        df.columns = df.columns.str.strip().str.upper().str.replace(' ', '_')
        return df.rename(columns={'CUE': 'CUE'})
    return pd.DataFrame()

@st.cache_data
def cargar_base_personas():
    if os.path.exists(EXCEL_PERSONAS):
        df = pd.read_excel(EXCEL_PERSONAS)
        df.columns = df.columns.str.strip().str.upper().str.replace(' ', '_')
        return df
    return pd.DataFrame()

# --- LÓGICA PRINCIPAL ---
fechas_ocupadas = obtener_fechas_ocupadas()
anio_actual = datetime.date.today().year
feriados_arg = holidays.Argentina(years=[anio_actual, anio_actual + 1])

# [EL RESTO DE TU LÓGICA DE SIDEBAR, ADMINISTRADOR Y VISTA DE USUARIO VA AQUÍ]
# (He mantenido el flujo tal cual lo tenías, solo aplica el cambio en el Contenedor 4 que sigue)

# --- CONTENEDOR 4 (COPIA ESTO EN TU ARCHIVO) ---
st.markdown('<div class="custom-card">', unsafe_allow_html=True)
st.subheader("📅 4. Selección de Turno Excluyente")
fecha_minima = datetime.date(anio_actual, 8, 1)
fecha_maxima = datetime.date(anio_actual, 11, 30)

fecha_sel = st.date_input("Seleccione el día:", min_value=fecha_minima, max_value=fecha_maxima)

es_valida = True
motivo = ""

if fecha_sel.weekday() >= 5:
    es_valida = False
    motivo = "Fin de semana no disponible."
elif fecha_sel in feriados_arg:
    es_valida = False
    motivo = f"Feriado: {feriados_arg.get(fecha_sel)}"
elif fecha_sel in fechas_ocupadas:
    es_valida = False
    motivo = "Esta fecha ya ha sido reservada."

if es_valida:
    st.info(f"🟢 La fecha {fecha_sel.strftime('%d/%m/%Y')} está disponible.")
else:
    st.error(f"🔴 {motivo}")

# BOTÓN DE REGISTRO
if st.button("Confirmar y Registrar Agenda", disabled=not es_valida):
    # Aquí iría tu lógica de guardar_reserva(...)
    st.success("Reserva realizada.")
    st.rerun()

st.markdown('</div>', unsafe_allow_html=True)
