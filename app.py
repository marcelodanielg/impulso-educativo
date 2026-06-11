import streamlit as st
import pandas as pd
import datetime
import gspread
from google.oauth2.service_account import Credentials

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Calendario Escolar", page_icon="📅", layout="centered")

# --- DISEÑO CSS (CORREGIDO) ---
st.markdown("""
    <style>
        .stApp { 
            background: linear-gradient(135deg, #e0eafc 0%, #cfdef3 100%);
        }
        .main-card {
            background: rgba(255, 255, 255, 0.9);
            padding: 30px;
            border-radius: 20px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.1);
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

# --- CONEXIÓN GOOGLE SHEETS ---
def conectar_gsheets():
    claves = dict(st.secrets["gcp_service_account"])
    claves["private_key"] = claves["private_key"].replace("\\n", "\n")
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    cred = Credentials.from_service_account_info(claves, scopes=scopes)
    return gspread.authorize(cred).open_by_url(st.secrets["spreadsheet_url"]).sheet1

def obtener_reservas():
    try:
        hoja = conectar_gsheets()
        data = hoja.get_all_values()
        if len(data) < 2: return pd.DataFrame(columns=["CUE"])
        return pd.DataFrame(data[1:], columns=data[0])
    except Exception as e:
        st.error(f"Error de conexión: {e}")
        return pd.DataFrame(columns=["CUE"])

# --- INTERFAZ ---
st.markdown('<div class="main-card">', unsafe_allow_html=True)
st.title("📅 Reserva de Turnos")
st.write("Complete el formulario para agendar su establecimiento.")

df_existentes = obtener_reservas()

with st.form("form_reserva"):
    cue = st.text_input("CUE del establecimiento")
    nombre_escuela = st.text_input("Nombre de la escuela")
    email = st.text_input("Email de contacto")
    dni = st.text_input("DNI del director/a")
    fecha = st.date_input("Fecha preferida", min_value=datetime.date.today())
    
    submit = st.form_submit_button("Confirmar Reserva")

    if submit:
        if cue in df_existentes["CUE"].values:
            st.error("⚠️ Este CUE ya tiene una reserva registrada.")
        elif not cue or not email or not nombre_escuela:
            st.warning("Por favor, complete todos los campos obligatorios.")
        else:
            try:
                hoja = conectar_gsheets()
                hoja.append_row([cue, nombre_escuela, email, dni, fecha.day, fecha.month, fecha.year])
                st.success("✅ ¡Reserva realizada correctamente!")
            except Exception as e:
                st.error(f"Error al guardar: {e}")

st.markdown('</div>', unsafe_allow_html=True)
