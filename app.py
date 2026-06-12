import streamlit as st
import pandas as pd
import datetime
import holidays
import os
import json
import io

# --- CONFIGURACIÓN E IMPORTACIONES ---
try:
    import gspread
    from google.oauth2.service_account import Credentials
    gsheets_librerias_listas = True
except ImportError:
    gsheets_librerias_listas = False

st.set_page_config(page_title="Gestión Escolar Excluyente", layout="wide")

# Inicialización de Estados
if "admin_autenticado" not in st.session_state: st.session_state.admin_autenticado = False
if "reserva_exitosa" not in st.session_state: st.session_state.reserva_exitosa = None

# --- FUNCIONES DE SOPORTE ---
def normalizar_texto(val):
    return str(val).strip().replace(".0", "").lower()

@st.cache_data
def cargar_datos():
    esc = pd.read_excel("base_escuelas.xlsx") if os.path.exists("base_escuelas.xlsx") else pd.DataFrame(columns=["CUE", "Nombre_Escuela"])
    per = pd.read_excel("personas.xlsx") if os.path.exists("personas.xlsx") else pd.DataFrame(columns=["DNI", "Apellido_Nombre"])
    res = pd.read_excel("registro_calendario.xlsx") if os.path.exists("registro_calendario.xlsx") else pd.DataFrame(columns=["CUE", "Dia_Reservado", "Mes_Reservado", "Anio_Reservado"])
    return esc, per, res

df_escuelas, df_personas, df_reservas = cargar_datos()
feriados_arg = holidays.Argentina(years=datetime.date.today().year)

# --- VISTA PÚBLICA ---
st.title("📅 Sistema de Reserva")

# 1. CUE con Validación de Duplicados
cue_input = st.text_input("Ingrese CUE de la institución (Obligatorio):")
cue_limpio = normalizar_texto(cue_input)
escuela_match = df_escuelas[df_escuelas['CUE'].astype(str).str.lower() == cue_limpio] if cue_input else pd.DataFrame()

escuela_valida = False
if cue_input:
    if not escuela_match.empty:
        if cue_limpio in df_reservas['CUE'].astype(str).str.lower().values:
            st.error("❌ ERROR: Este CUE ya tiene una reserva activa.")
        else:
            st.success(f"Escuela: {escuela_match.iloc[0]['Nombre_Escuela']}")
            escuela_valida = True
    else:
        st.error("❌ CUE no encontrado.")

# 2. Autoridad (Editable manual si no existe)
dni_input = st.text_input("Ingrese su DNI (Obligatorio):")
persona_match = df_personas[df_personas['DNI'].astype(str) == normalizar_texto(dni_input)] if dni_input else pd.DataFrame()

if not persona_match.empty:
    nombre_dir = st.text_input("Nombre del directivo:", value=persona_match.iloc[0]['Apellido_Nombre'])
else:
    nombre_dir = st.text_input("DNI no registrado. Ingrese Nombre y Apellido manualmente:")

# 3. Estructura Resaltada (Obligatoria)
st.markdown("---")
st.error("### ⚠️ ATENCIÓN: SELECCIONE SU PLAN DE ESTUDIOS")
estructura = st.radio("Debe elegir una opción para continuar:", ["5° y 6° Año", "6° y 7° Año"], index=None)
st.markdown("---")

# 4. Calendario con Feedback Visual
st.subheader("📅 Selección de Turno")
fecha = st.date_input("Elija su fecha:", value=None)

es_habil = False
if fecha:
    if fecha.weekday() >= 5: st.error("❌ Fines de semana no permitidos.")
    elif fecha in feriados_arg: st.error(f"❌ Feriado: {feriados_arg.get(fecha)}")
    elif any((df_reservas['Dia_Reservado']==fecha.day) & (df_reservas['Mes_Reservado']==fecha.month) & (df_reservas['Anio_Reservado']==fecha.year)):
        st.error("❌ Esta fecha ya ha sido reservada.")
    else:
        st.success("✅ Fecha disponible para reservar.")
        es_habil = True

# --- BOTÓN FINAL ---
formulario_listo = escuela_valida and nombre_dir and estructura and es_habil

if st.button("Confirmar Reserva", disabled=not formulario_listo):
    st.balloons()
    st.success("¡Reserva procesada con éxito!")
    # Aquí iría tu función original guardar_reserva(...)
