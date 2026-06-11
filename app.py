import streamlit as st
import pandas as pd
import datetime
import os
import json

#gsheets_librerias_listas = False
try:
    import gspread
    from google.oauth2.service_account import Credentials
    gsheets_librerias_listas = True
except ImportError:
    pass

EXCEL_ESCUELAS = "base_escuelas.xlsx"
CONFIG_SISTEMA = "config_sistema.json"

st.set_page_config(page_title="Calendario Escolar", page_icon="📅", layout="centered")

#st.markdown("""
    <style>
        .stApp { 
            background: linear-gradient(135deg, #f0f4f8 0%, #d9e2ec 100%);
        }
        .main-card {
            background: rgba(255, 255, 255, 0.85);
            backdrop-filter: blur(10px);
            padding: 30px;
            border-radius: 20px;
            box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.15);
            border: 1px solid rgba(255, 255, 255, 0.18);
        }
        .stButton>button {
            width: 100%;
            border-radius: 10px;
            background-color: #0284c7;
            color: white;
            font-weight: bold;
        }
    </style>
""", unsafe_allow_html=True)

#if "admin_autenticado" not in st.session_state: st.session_state.admin_autenticado = False

def usar_gsheets():
    return gsheets_librerias_listas and "gcp_service_account" in st.secrets

def conectar_gsheets():
    claves = dict(st.secrets["gcp_service_account"])
    claves["private_key"] = claves["private_key"].replace("\\n", "\n")
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    cred = Credentials.from_service_account_info(claves, scopes=scopes)
    return gspread.authorize(cred).open_by_url(st.secrets["spreadsheet_url"]).sheet1

#def obtener_reservas():
    if not usar_gsheets(): return pd.DataFrame()
    try:
        hoja = conectar_gsheets()
        data = hoja.get_all_values()
        if len(data) < 2: return pd.DataFrame(columns=["CUE"])
        return pd.DataFrame(data[1:], columns=data[0])
    except Exception as e:
        st.error(f"Error cargando registros: {e}")
        return pd.DataFrame(columns=["CUE"])

def guardar_reserva(datos):
    if usar_gsheets():
        hoja = conectar_gsheets()
        hoja.append_row([datos.get(k, "") for k in ["CUE", "Escuela", "Email", "DNI_Director", "Dia_Reservado", "Mes_Reservado", "Anio_Reservado"]])

#st.markdown('<div class="main-card">', unsafe_allow_html=True)
st.title("📅 Reserva de Turnos")
st.write("Complete el formulario para agendar su establecimiento.")

with st.form("form_reserva"):
    cue = st.text_input("CUE del establecimiento")
    nombre_escuela = st.text_input("Nombre de la escuela")
    email = st.text_input("Email de contacto")
    dni = st.text_input("DNI del director/a")
    fecha = st.date_input("Fecha preferida", min_value=datetime.date.today())
    
    submit = st.form_submit_button("Confirmar Reserva")

    if submit:
        #        df_existentes = obtener_reservas()
        if cue in df_existentes["CUE"].values:
            st.error("⚠️ Este CUE ya tiene una reserva registrada.")
        elif not cue or not email:
            st.warning("Por favor, complete los campos obligatorios.")
        else:
            datos = {
                "CUE": cue, "Escuela": nombre_escuela, "Email": email,
                "DNI_Director": dni, "Dia_Reservado": fecha.day,
                "Mes_Reservado": fecha.month, "Anio_Reservado": fecha.year
            }
            guardar_reserva(datos)
            st.success("✅ ¡Reserva realizada correctamente!")

st.markdown('</div>', unsafe_allow_html=True)
