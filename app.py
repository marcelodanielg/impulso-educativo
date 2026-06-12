import streamlit as st
import pandas as pd
import datetime
import holidays
import os
import json
import io

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Calendario Escolar", layout="wide")

# Inicialización de estado
if "reserva_exitosa" not in st.session_state: st.session_state.reserva_exitosa = None

# Funciones auxiliares
def normalizar_texto(val):
    if pd.isna(val): return ""
    return str(val).strip().replace(".0", "")

def cargar_datos():
    # Lógica simplificada para cargar DataFrames (ajustar según tus rutas locales)
    df_esc = pd.read_excel("base_escuelas.xlsx") if os.path.exists("base_escuelas.xlsx") else pd.DataFrame(columns=["CUE", "Nombre_Escuela", "Modalidad_Oferta", "Departamento", "Domicilio"])
    df_per = pd.read_excel("personas.xlsx") if os.path.exists("personas.xlsx") else pd.DataFrame(columns=["DNI", "Apellido_Nombre", "Telefono"])
    return df_esc, df_per

df_escuelas, df_personas = cargar_datos()

# --- INTERFAZ PÚBLICA ---
st.markdown("<h1 style='text-align: center; color: #0284c7;'>📅 Sistema de Reserva de Turnos</h1>", unsafe_allow_html=True)

# 1. CUE con validación de duplicados
cue_input = st.text_input("Ingrese el CUE de la institución:", value="", placeholder="Ej: 7000123")
coincidencia_esc = df_escuelas[df_escuelas['CUE'].astype(str) == normalizar_texto(cue_input)]
escuela_valida = not coincidencia_esc.empty

if cue_input:
    if escuela_valida:
        st.success(f"Escuela: {coincidencia_esc.iloc[0]['Nombre_Escuela']}")
    else:
        st.error("CUE no encontrado.")

# 2. Autoridad (Editable si no existe)
dni_input = st.text_input("Ingrese DNI del directivo:", value="", placeholder="Ej: 22333444")
persona_data = df_personas[df_personas['DNI'].astype(str) == normalizar_texto(dni_input)]

if not persona_data.empty:
    nombre_dir = st.text_input("Nombre de la autoridad:", value=persona_data.iloc[0]['Apellido_Nombre'])
else:
    nombre_dir = st.text_input("DNI no encontrado. Ingrese Apellido y Nombre manualmente:")

# 3. Estructura resaltada (Obligatoria)
st.markdown("---")
st.error("### ⚠️ ATENCIÓN: SELECCIONE SU ESTRUCTURA DE PLAN DE ESTUDIOS")
estructura = st.radio("Debe elegir una opción:", ["5° y 6° Año", "6° y 7° Año"], index=None)
st.markdown("---")

# 4. Selección de Turno (Validación Visual)
st.subheader("📅 Selección de Turno")
fecha = st.date_input("Elija una fecha (Agosto - Noviembre):", value=None)

# Lógica de validación
feriados = holidays.Argentina(years=datetime.date.today().year)
fechas_ocupadas = [] # Aquí deberías cargar las ya reservadas desde tu CSV/GSheet

if fecha:
    if fecha.weekday() >= 5:
        st.error("❌ Fines de semana no disponibles.")
    elif fecha in feriados:
        st.error(f"❌ Feriado: {feriados.get(fecha)}")
    elif fecha in fechas_ocupadas:
        st.error("❌ Esta fecha ya ha sido reservada.")
    else:
        st.success(f"✅ Fecha {fecha} disponible.")

# --- BOTÓN FINAL DE CARGA ---
# Validación de obligatoriedad
campos_ok = (escuela_valida and nombre_dir and estructura and fecha and not (fecha.weekday() >= 5 or fecha in feriados or fecha in fechas_ocupadas))

if st.button("Confirmar Reserva", disabled=not campos_ok):
    st.balloons()
    st.success("¡Reserva realizada con éxito!")
    # Aquí iría la lógica de guardado (append a GSheet o Excel)

