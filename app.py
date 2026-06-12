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

# Configuración de página
st.set_page_config(page_title="Calendario Excluyente", page_icon="📅", layout="wide")

# --- FUNCIONES DE SOPORTE ---
def usando_google_sheets():
    if not gsheets_librerias_listas: return False
    return "gcp_service_account" in st.secrets and "spreadsheet_url" in st.secrets

def conectar_google_sheets():
    claves = dict(st.secrets["gcp_service_account"])
    claves["private_key"] = claves["private_key"].replace("\\n", "\n")
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    credenciales = Credentials.from_service_account_info(claves, scopes=scopes)
    cliente = gspread.authorize(credenciales)
    return cliente.open_by_url(st.secrets["spreadsheet_url"]).sheet1

def normalizar_texto(val):
    if pd.isna(val): return ""
    val_str = str(val).strip()
    return val_str[:-2] if val_str.endswith(".0") else val_str

# --- LÓGICA DE VALIDACIÓN DE DUPLICADOS ---
def existe_cue_reservado(cue_ingresado):
    """Verifica si el CUE ya tiene una reserva previa."""
    cue_normalizado = normalizar_texto(cue_ingresado)
    if usando_google_sheets():
        try:
            hoja = conectar_google_sheets()
            valores = hoja.get_all_values()
            if len(valores) > 1:
                df = pd.DataFrame(valores[1:], columns=valores[0])
                return cue_normalizado in df['CUE'].astype(str).values
        except: return False
    else:
        if os.path.exists(EXCEL_RESERVAS_LOCAL):
            df = pd.read_excel(EXCEL_RESERVAS_LOCAL)
            return cue_normalizado in df['CUE'].astype(str).values
    return False

# --- CARGA DE DATOS ---
@st.cache_data
def cargar_base_escuelas():
    if os.path.exists(EXCEL_ESCUELAS):
        df = pd.read_excel(EXCEL_ESCUELAS)
        df.columns = df.columns.str.strip().str.upper().str.replace(' ', '_')
        return df # Simplificado para brevedad, mantener tu lógica original de renombrado
    return pd.DataFrame()

@st.cache_data
def cargar_base_personas():
    if os.path.exists(EXCEL_PERSONAS):
        df = pd.read_excel(EXCEL_PERSONAS)
        return df
    return pd.DataFrame()

# ... (El resto de funciones cargar_base, obtener_fechas_ocupadas y guardar_reserva permanecen igual) ...

# --- VISTA PÚBLICA (Fragmento modificado con la validación) ---
# En la sección donde está el botón de confirmar:

if st.button("Confirmar y Registrar Agenda", disabled=not formulario_listo):
    # Validamos duplicados antes de guardar
    if existe_cue_reservado(cue_ingresado):
        st.error(f"❌ El CUE {cue_ingresado} ya tiene una reserva asignada. No se permiten reservas duplicadas.")
    else:
        # Aquí va tu lógica de guardado original
        datos_reserva = {
            "CUE": cue_ingresado,
            "Escuela": nombre_escuela,
            # ... resto de campos ...
            "Fecha_Registro": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        guardar_reserva(datos_reserva)
        st.session_state.reserva_exitosa = datos_reserva
        st.cache_data.clear()
        st.rerun()
