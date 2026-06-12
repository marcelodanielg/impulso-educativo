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

st.set_page_config(page_title="Calendario Excluyente", page_icon="📅", layout="wide")

# --- FUNCIONES DE SOPORTE ---
def usando_google_sheets():
    return gsheets_librerias_listas and "gcp_service_account" in st.secrets and "spreadsheet_url" in st.secrets

def conectar_google_sheets():
    claves = dict(st.secrets["gcp_service_account"])
    if "private_key" in claves: claves["private_key"] = claves["private_key"].replace("\\n", "\n")
    credenciales = Credentials.from_service_account_info(claves, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
    return gspread.authorize(credenciales).open_by_url(st.secrets["spreadsheet_url"]).sheet1

@st.cache_data(ttl=600)
def obtener_fechas_ocupadas():
    """Retorna un conjunto de fechas ocupadas para validar disponibilidad."""
    fechas = set()
    if usando_google_sheets():
        try:
            valores = conectar_google_sheets().get_all_values()
            if len(valores) > 1:
                df = pd.DataFrame(valores[1:], columns=valores[0])
                for _, row in df.iterrows():
                    fechas.add(datetime.date(int(float(row['Anio_Reservado'])), int(float(row['Mes_Reservado'])), int(float(row['Dia_Reservado']))))
        except: pass
    elif os.path.exists(EXCEL_RESERVAS_LOCAL):
        try:
            df = pd.read_excel(EXCEL_RESERVAS_LOCAL)
            for _, row in df.iterrows():
                fechas.add(datetime.date(int(row['Anio_Reservado']), int(row['Mes_Reservado']), int(row['Dia_Reservado'])))
        except: pass
    return fechas

# --- LÓGICA PRINCIPAL ---
fechas_ocupadas = obtener_fechas_ocupadas()
feriados_arg = holidays.Argentina(years=[datetime.date.today().year])

# ... (El resto de tu código original se mantiene igual hasta el CONTENEDOR 4) ...

# --- MODIFICACIÓN DEL CONTENEDOR 4 ---
st.markdown('<div class="custom-card">', unsafe_allow_html=True)
st.subheader("📅 4. Selección de Turno Excluyente")

fecha_minima = datetime.date(datetime.date.today().year, 8, 1)
fecha_maxima = datetime.date(datetime.date.today().year, 11, 30)

fecha_seleccionada = st.date_input(
    "Seleccione el día que reservará:",
    min_value=fecha_minima,
    max_value=fecha_maxima
)

# Lógica de validación
es_valida = True
mensaje_error = ""

if fecha_seleccionada.weekday() >= 5:
    es_valida = False
    mensaje_error = "La fecha seleccionada es fin de semana."
elif fecha_seleccionada in feriados_arg:
    es_valida = False
    mensaje_error = f"La fecha es feriado: {feriados_arg.get(fecha_seleccionada)}"
elif fecha_seleccionada in fechas_ocupadas:
    es_valida = False
    mensaje_error = "Esta fecha ya ha sido reservada por otra institución."

if es_valida:
    st.success(f"🟢 La fecha {fecha_seleccionada.strftime('%d/%m/%Y')} está disponible.")
else:
    st.error(f"🔴 No disponible: {mensaje_error}")

# El botón de guardar ahora usa 'es_valida'
if st.button("Confirmar y Registrar Agenda", disabled=not es_valida):
    # ... (tu código de guardado original) ...
    st.success("Reserva realizada con éxito.")
    st.rerun()

st.markdown('</div>', unsafe_allow_html=True)

