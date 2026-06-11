import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

#st.set_page_config(page_title="Calendario Escolar", page_icon="📅", layout="centered")

## El bloque de CSS va al ras del margen para evitar IndentationError
st.markdown("""
<style>
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    }
    .main-card {
        background: white;
        padding: 30px;
        border-radius: 20px;
        box-shadow: 0 10px 25px rgba(0,0,0,0.1);
        margin-top: 20px;
    }
    h1 { color: #0284c7; text-align: center; }
    .stButton>button {
        width: 100%;
        background-color: #0284c7;
        color: white;
        font-weight: bold;
        border-radius: 10px;
    }
</style>
""", unsafe_allow_html=True)

#def connect_gsheet():
    try:
        if "gcp_service_account" not in st.secrets:
            return None
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        return client.open_by_url(st.secrets["spreadsheet_url"]).sheet1
    except Exception as e:
        st.error(f"Error de conexión: {e}")
        return None

#st.markdown('<div class="main-card">', unsafe_allow_html=True)
st.title("📅 Reserva de Turnos")

hoja = connect_gsheet()

if hoja:
    # Obtener datos para validar duplicados
    data = hoja.get_all_values()
    if len(data) > 0:
        df = pd.DataFrame(data[1:], columns=data[0]) if len(data) > 1 else pd.DataFrame(columns=["CUE", "Email"])
    else:
        # Si está vacía, creamos los encabezados
        hoja.append_row(["CUE", "Escuela", "Email", "DNI_Director", "Fecha"])
        df = pd.DataFrame(columns=["CUE", "Email"])

    with st.form("form_reserva"):
        cue = st.text_input("CUE del establecimiento")
        nombre = st.text_input("Nombre de la escuela")
        email = st.text_input("Email de contacto")
        dni = st.text_input("DNI del director/a")
        fecha = st.date_input("Fecha preferida")
        
        submitted = st.form_submit_button("Confirmar Reserva")

        if submitted:
            #            if not cue or not nombre or not email:
                st.warning("⚠️ Todos los campos son obligatorios.")
            elif cue in df["CUE"].values:
                st.error("⚠️ Este CUE ya tiene una reserva registrada.")
            else:
                try:
                    hoja.append_row([cue, nombre, email, dni, str(fecha)])
                    st.success("✅ ¡Reserva realizada con éxito!")
                except Exception as e:
                    st.error(f"Error al guardar: {e}")

st.markdown('</div>', unsafe_allow_html=True)
