import streamlit as st
import pandas as pd
import datetime
import holidays
import os
import json
import io

# --- IMPORTACIONES Y CONFIGURACIÓN ---
gsheets_librerias_listas = False
try:
    import gspread
    from google.oauth2.service_account import Credentials
    gsheets_librerias_listas = True
except ImportError:
    pass

EXCEL_ESCUELAS = "base_escuelas.xlsx"
EXCEL_RESERVAS_LOCAL = "registro_calendario.xlsx"

st.set_page_config(page_title="Sistema de Reserva", layout="wide")

# --- FUNCIONES DE SOPORTE ---
def normalizar_texto(val):
    return str(val).strip().replace(".0", "") if pd.notna(val) else ""

@st.cache_data
def cargar_base_escuelas():
    if os.path.exists(EXCEL_ESCUELAS):
        df = pd.read_excel(EXCEL_ESCUELAS)
        df.columns = df.columns.str.strip().str.upper().str.replace(' ', '_')
        df['CUE'] = df['CUE'].apply(normalizar_texto)
        return df
    return pd.DataFrame(columns=['CUE', 'Nombre_Escuela'])

def obtener_reservas_df():
    if os.path.exists(EXCEL_RESERVAS_LOCAL):
        return pd.read_excel(EXCEL_RESERVAS_LOCAL)
    return pd.DataFrame(columns=['CUE'])

# --- INICIO DE LA APP ---
df_escuelas = cargar_base_escuelas()

st.title("📅 Sistema de Reserva de Turnos")

# VISTA DE DIRECTORES
st.markdown("### 📍 1. Identificación del Establecimiento Educativo")

# 1. Obtenemos datos para validar
df_reservas = obtener_reservas_df()
cues_registrados = df_reservas['CUE'].astype(str).tolist() if 'CUE' in df_reservas.columns else []

# 2. Input Único (La clave para evitar el error)
cue_ingresado = st.text_input("Ingrese el CUE de la institución:", key="cue_input_user", placeholder="Ej: 7000123")

# 3. Lógica de validación
if cue_ingresado:
    cue_limpio = normalizar_texto(cue_ingresado)
    
    if cue_limpio in cues_registrados:
        st.error(f"❌ La escuela con CUE {cue_limpio} YA tiene una reserva confirmada.")
    else:
        coincidencia_esc = df_escuelas[df_escuelas['CUE'] == cue_limpio]
        if not coincidencia_esc.empty:
            st.success(f"✅ Escuela identificada: {coincidencia_esc.iloc[0]['Nombre_Escuela']}")
            
            # --- AQUÍ CONTINÚA TU FORMULARIO (CONTENEDORES 2, 3 y 4) ---
            # Ejemplo:
            # dni = st.text_input("DNI:", key="dni_unico")
            # ... resto de tu lógica ...
            
        else:
            st.error("❌ El CUE ingresado no figura en el padrón.")
