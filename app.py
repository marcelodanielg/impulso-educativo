import datetime
import io
import json
import os
import holidays
import pandas as pd
import streamlit as st
from fpdf import FPDF

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
EXCEL_RESERVAS_LOCAL = "registro_reservas_2026.xlsx"
CONFIG_SISTEMA = "config_sistema.json"

# Configuración de página de Streamlit
st.set_page_config(
    page_title="Capacitación en RCP",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Estilos CSS con paleta de colores suave, estilizada e intuitiva
st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        
        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif;
        }
        
        .stApp { background-color: #f8fafc; }
        
        /* Contenedores por Paso con tonos suaves */
        .step-card {
            background-color: #ffffff;
            padding: 24px;
            border-radius: 16px;
            box-shadow: 0 4px 20px -2px rgba(0, 0, 0, 0.03);
            margin-bottom: 24px;
            border: 1px solid #e2e8f0;
            transition: all 0.3s ease;
        }
        
        .step-card-blue {
            border-top: 5px solid #3b82f6;
            background: linear-gradient(180deg, #f0f7ff 0%, #ffffff 100px);
        }
        
        .step-card-purple {
            border-top: 5px solid #8b5cf6;
            background: linear-gradient(180deg, #f5f3ff 0%, #ffffff 100px);
        }
        
        .step-card-amber {
            border-top: 5px solid #f59e0b;
            background: linear-gradient(180deg, #fffbeb 0%, #ffffff 100px);
        }
        
        .step-card-emerald {
            border-top: 5px solid #10b981;
            background: linear-gradient(180deg, #ecfdf5 0%, #ffffff 100px);
        }

        .step-title {
            font-size: 1.25rem;
            font-weight: 700;
            margin-bottom: 12px;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .title-blue { color: #1d4ed8; }
        .title-purple { color: #6d28d9; }
        .title-amber { color: #b45309; }
        .title-emerald { color: #047857; }

        /* Estilo para los botones */
        .stButton>button, .stDownloadButton>button {
            background-color: #2563eb !important;
            color: white !important;
            border-radius: 10px !important;
            padding: 12px 24px !important;
            font-weight: 600 !important;
            border: none !important;
            transition: all 0.2s ease !important;
            box-shadow: 0 4px 12px rgba(37, 99, 235, 0.2) !important;
            width: 100% !important;
        }
        
        .stButton>button:hover, .stDownloadButton>button:hover {
            background-color: #1d4ed8 !important;
            transform: translateY(-2px) !important;
            box-shadow: 0 6px 16px rgba(37, 99, 235, 0.3) !important;
        }

        /* Pills informativos */
        .info-pill {
            background-color: #ffffff;
            border: 1px solid #cbd5e1;
            padding: 14px 18px;
            border-radius: 12px;
            margin-top: 10px;
            margin-bottom: 10px;
        }
        
        .info-pill-success {
            background-color: #f0fdf4;
            border-color: #bbf7d0;
            color: #166534;
        }

        .atencion-box {
            background-color: #fff7ed;
            border: 1px solid #fed7aa;
            padding: 12px 16px;
            border-radius: 10px;
            color: #c2410c;
            font-size: 0.95rem;
            margin-bottom: 15px;
            font-weight: 500;
        }

        section[data-testid="stSidebar"] { background-color: #0f172a !important; }
        section[data-testid="stSidebar"] * { color: #f1f5f9 !important; }
    </style>
""",
    unsafe_allow_html=True,
)

if "escuelas_procesadas" not in st.session_state:
    st.session_state.escuelas_procesadas = None
if "personas_procesadas" not in st.session_state:
    st.session_state.personas_procesadas = None
if "admin_autenticado" not in st.session_state:
    st.session_state.admin_autenticado = False
if "reserva_exitosa" not in st.session_state:
    st.session_state.reserva_exitosa = None


# --- FUNCIÓN GENERADORA DE COMPROBANTE EN PDF ---
def generar_pdf_comprobante(reserva):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_margins(15, 15, 15)

    # Encabezado
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, "IMPULSO EDUCATIVO", ln=True, align="C")
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(
        0,
        6,
        "Capacitación en RCP - Comprobante de Reserva de turno",
        ln=True,
        align="C",
    )
    pdf.ln(4)

    # Línea separadora
    pdf.set_line_width(0.5)
    pdf.line(15, pdf.get_y(), 195, pdf.get_y())
    pdf.ln(8)

    # Título principal
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 8, "¡Reserva Confirmada Exitosamente!", ln=True, align="C")
    pdf.ln(6)

    # Detalle de datos
    pdf.set_font("Helvetica", "", 11)
    motivo_txt = reserva.get("Motivo_Cambio", "")
    motivo_linea = f"\nMotivo Reprogramación: {motivo_txt}" if motivo_txt else ""

    pdf.multi_cell(
        0,
        7,
        f"Establecimiento: {reserva.get('Escuela', '')}\n"
        f"CUE: {reserva.get('CUE', '')}\n"
        f"Director Solicitante: {reserva.get('Director', '')}\n"
        f"Teléfono Contacto: {reserva.get('Telefono_Contacto', '')}\n"
        f"Total Alumnos Registrados: {reserva.get('Total_Alumnos', '')} alumnos.\n"
        f"Detalle de Cursos: {reserva.get('Detalle_Divisiones_Alumnos', '')}"
        f"{motivo_linea}",
    )
    pdf.ln(5)

    # Recuadro Fecha Reservada
    dia = int(float(reserva.get("Dia_Reservado", 0)))
    mes = int(float(reserva.get("Mes_Reservado", 0)))
    anio = int(float(reserva.get("Anio_Reservado", 0)))

    pdf.set_fill_color(240, 253, 244)
    pdf.set_draw_color(187, 247, 208)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(
        0,
        12,
        f"Día Reservado: {dia:02d} / {mes:02d} / {anio}",
        border=1,
        ln=True,
        align="C",
        fill=True,
    )

    pdf.ln(10)
    pdf.set_font("Helvetica", "I", 9)
    pdf.cell(
        0,
        5,
        f"Registro realizado el {reserva.get('Fecha_Registro', '')}",
        ln=True,
        align="C",
    )
    pdf.cell(
        0,
        5,
        "Este comprobante avala la asignación del turno institucional.",
        ln=True,
        align="C",
    )

    return bytes(pdf.output())


# --- DETECCIÓN DE BASE DE DATOS ACTIVA ---
def usando_google_sheets():
    if not gsheets_librerias_listas:
        return False
    try:
        return (
            "gcp_service_account" in st.secrets
            and "spreadsheet_url" in st.secrets
        )
    except Exception:
        return False


def conectar_google_sheets():
    try:
        claves = dict(st.secrets["gcp_service_account"])
        if "private_key" in claves:
            claves["private_key"] = claves["private_key"].replace("\\n", "\n")

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        credenciales = Credentials.from_service_account_info(
            claves, scopes=scopes
        )
        cliente = gspread.authorize(credenciales)
        planilla = cliente.open_by_url(st.secrets["spreadsheet_url"])
        return planilla.sheet1
    except Exception as e:
        st.error(f"Error al conectar con Google Sheets: {e}")
        return None


# --- PERSISTENCIA DE CONFIGURACIONES ---
def cargar_configuracion_sistema():
    if os.path.exists(CONFIG_SISTEMA):
        try:
            with open(CONFIG_SISTEMA, "r") as f:
                return json.load(f)
        except Exception:
            pass
    return {"registro_habilitado": True}


def guardar_configuracion_sistema(config):
    try:
        with open(CONFIG_SISTEMA, "w") as f:
            json.dump(config, f)
    except Exception as e:
        st.error(f"Error al guardar configuración: {e}")


def normalizar_texto(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    if val_str.lower() == "nan":
        return ""
    return val_str


@st.cache_data
def cargar_base_escuelas():
    if os.path.exists(EXCEL_ESCUELAS):
        try:
            df = pd.read_excel(EXCEL_ESCUELAS)
            df.columns = df.columns.str.strip().str.upper().str.replace(" ", "_")

            col_cue = [c for c in df.columns if "CUE" in c]
            col_nombre = [c for c in df.columns if "NOM" in c or "ESC" in c]
            col_mod = [c for c in df.columns if "MOD" in c or "OFER" in c]
            col_depto = [c for c in df.columns if "DEP" in c]
            col_dom = [c for c in df.columns if "DOM" in c or "DIR" in c]

            mapping = {}
            if col_cue:
                mapping[col_cue[0]] = "CUE"
            if col_nombre:
                mapping[col_nombre[0]] = "Nombre_Escuela"
            if col_mod:
                mapping[col_mod[0]] = "Modalidad_Oferta"
            if col_depto:
                mapping[col_depto[0]] = "Departamento"
            if col_dom:
                mapping[col_dom[0]] = "Domicilio"

            df = df.rename(columns=mapping)
            df["CUE"] = df["CUE"].apply(normalizar_texto)

            for col in ["Modalidad_Oferta", "Departamento", "Domicilio"]:
                if col not in df.columns:
                    df[col] = "No especificado"
                else:
                    df[col] = (
                        df[col].fillna("No especificado").astype(str).str.strip()
                    )

            return df[
                [
                    "CUE",
                    "Nombre_Escuela",
                    "Modalidad_Oferta",
                    "Departamento",
                    "Domicilio",
                ]
            ].drop_duplicates()
        except Exception as e:
            st.error(f"Error al leer la base de escuelas: {e}")
            return pd.DataFrame(
                columns=[
                    "CUE",
                    "Nombre_Escuela",
                    "Modalidad_Oferta",
                    "Departamento",
                    "Domicilio",
                ]
            )
    return pd.DataFrame(
        columns=[
            "CUE",
            "Nombre_Escuela",
            "Modalidad_Oferta",
            "Departamento",
            "Domicilio",
        ]
    )


@st.cache_data
def cargar_base_personas():
    if os.path.exists(EXCEL_PERSONAS):
        try:
            df = pd.read_excel(EXCEL_PERSONAS)
            df.columns = df.columns.str.strip().str.upper().str.replace(" ", "_")

            col_dni = [c for c in df.columns if "DNI" in c or "DOC" in c]
            col_apellido = [c for c in df.columns if "APE" in c]
            col_nombre = [
                c for c in df.columns if "NOM" in c and "ESC" not in c
            ]
            col_tel = [c for c in df.columns if "TEL" in c or "CEL" in c]

            mapping = {}
            if col_dni:
                mapping[col_dni[0]] = "DNI"
            if col_apellido:
                mapping[col_apellido[0]] = "Apellido"
            if col_nombre:
                mapping[col_nombre[0]] = "Nombre"
            if col_tel:
                mapping[col_tel[0]] = "Telefono"

            df = df.rename(columns=mapping)

            if (
                "DNI" in df.columns
                and "Apellido" in df.columns
                and "Nombre" in df.columns
            ):
                df["DNI"] = df["DNI"].apply(normalizar_texto)
                df["Apellido"] = (
                    df["Apellido"].fillna("").astype(str).str.strip()
                )
                df["Nombre"] = df["Nombre"].fillna("").astype(str).str.strip()
                df["Apellido_Nombre"] = df["Apellido"] + ", " + df["Nombre"]
                df["Apellido_Nombre"] = df["Apellido_Nombre"].str.strip(", ")

                if "Telefono" not in df.columns:
                    df["Telefono"] = ""
                else:
                    df["Telefono"] = df["Telefono"].apply(normalizar_texto)

                return df[
                    ["DNI", "Apellido_Nombre", "Telefono"]
                ].drop_duplicates()
            else:
                st.error(
                    "El Excel de personas debe tener columnas identificables"
                    " para DNI, APELLIDO y NOMBRE."
                )
                return pd.DataFrame(
                    columns=["DNI", "Apellido_Nombre", "Telefono"]
                )
        except Exception as e:
            st.error(f"Error al leer la base de personas: {e}")
            return pd.DataFrame(columns=["DNI", "Apellido_Nombre", "Telefono"])
    return pd.DataFrame(columns=["DNI", "Apellido_Nombre", "Telefono"])


COLUMNAS_SISTEMA = [
    "CUE",
    "Escuela",
    "Modalidad_Oferta",
    "Departamento",
    "Domicilio",
    "DNI_Director",
    "Director",
    "Telefono_Contacto",
    "Estructura_Declarada",
    "Detalle_Divisiones_Alumnos",
    "Total_Alumnos",
    "Dia_Reservado",
    "Mes_Reservado",
    "Anio_Reservado",
    "Fecha_Registro",
    "Motivo_Cambio",
]


def cargar_reservas_existentes():
    if usando_google_sheets():
        try:
            hoja = conectar_google_sheets()
            if hoja:
                valores = hoja.get_all_values()
                if valores and len(valores) > 1:
                    df = pd.DataFrame(valores[1:], columns=valores[0])
                    if "Motivo_Cambio" not in df.columns:
                        df["Motivo_Cambio"] = ""
                    return df
        except Exception:
            pass
    else:
        if os.path.exists(EXCEL_RESERVAS_LOCAL):
            try:
                df = pd.read_excel(EXCEL_RESERVAS_LOCAL)
                if "Motivo_Cambio" not in df.columns:
                    df["Motivo_Cambio"] = ""
                return df
            except Exception:
                pass
    return pd.DataFrame(columns=COLUMNAS_SISTEMA)


def obtener_fechas_ocupadas(df_reservas):
    fechas = []
    if not df_reservas.empty and "Dia_Reservado" in df_reservas.columns:
        for _, row in df_reservas.iterrows():
            try:
                d = int(float(row["Dia_Reservado"]))
                m = int(float(row["Mes_Reservado"]))
                a = int(float(row["Anio_Reservado"]))
                fechas.append(datetime.date(a, m, d))
            except Exception:
                continue
    return set(fechas)


def guardar_reserva(datos):
    if usando_google_sheets():
        try:
            hoja = conectar_google_sheets()
            if hoja:
                valores = hoja.get_all_values()
                if not valores or len(valores) == 0:
                    hoja.append_row(COLUMNAS_SISTEMA)

                datos_lista = []
                for col in COLUMNAS_SISTEMA:
                    val = datos.get(col, "")
                    datos_lista.append(str(val))
                hoja.append_row(datos_lista)
        except Exception as e:
            st.error(f"Error crítico al registrar en Google Sheets: {e}")
    else:
        nuevo_df = pd.DataFrame([datos])
        if os.path.exists(EXCEL_RESERVAS_LOCAL):
            try:
                df_actual = pd.read_excel(EXCEL_RESERVAS_LOCAL)
                df_final = pd.concat([df_actual, nuevo_df], ignore_index=True)
            except Exception:
                df_final = nuevo_df
        else:
            df_final = nuevo_df
        df_final.to_excel(EXCEL_RESERVAS_LOCAL, index=False)


def actualizar_reserva_existente(cue, nueva_fecha, motivo_cambio):
    fecha_mod_str = (
        datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " (Modificado)"
    )
    if usando_google_sheets():
        try:
            hoja = conectar_google_sheets()
            if hoja:
                valores = hoja.get_all_values()
                if valores and len(valores) > 1:
                    encabezados = valores[0]

                    if "Motivo_Cambio" not in encabezados:
                        hoja.update_cell(1, len(encabezados) + 1, "Motivo_Cambio")
                        encabezados.append("Motivo_Cambio")

                    idx_cue = encabezados.index("CUE")
                    idx_dia = encabezados.index("Dia_Reservado")
                    idx_mes = encabezados.index("Mes_Reservado")
                    idx_anio = encabezados.index("Anio_Reservado")
                    idx_fecha_reg = encabezados.index("Fecha_Registro")
                    idx_motivo = encabezados.index("Motivo_Cambio")

                    for num_fila, fila in enumerate(valores[1:], start=2):
                        if (
                            normalizar_texto(
                                fila[idx_cue] if len(fila) > idx_cue else ""
                            )
                            == cue
                        ):
                            hoja.update_cell(num_fila, idx_dia + 1, nueva_fecha.day)
                            hoja.update_cell(num_fila, idx_mes + 1, nueva_fecha.month)
                            hoja.update_cell(
                                num_fila, idx_anio + 1, nueva_fecha.year
                            )
                            hoja.update_cell(
                                num_fila, idx_fecha_reg + 1, fecha_mod_str
                            )
                            hoja.update_cell(
                                num_fila, idx_motivo + 1, motivo_cambio
                            )
                            break
        except Exception as e:
            st.error(f"Error al actualizar la reserva en Google Sheets: {e}")
    else:
        if os.path.exists(EXCEL_RESERVAS_LOCAL):
            try:
                df = pd.read_excel(EXCEL_RESERVAS_LOCAL)
                df["CUE"] = df["CUE"].apply(normalizar_texto)
                mask = df["CUE"] == cue
                if mask.any():
                    df.loc[mask, "Dia_Reservado"] = nueva_fecha.day
                    df.loc[mask, "Mes_Reservado"] = nueva_fecha.month
                    df.loc[mask, "Anio_Reservado"] = nueva_fecha.year
                    df.loc[mask, "Fecha_Registro"] = fecha_mod_str
                    df.loc[mask, "Motivo_Cambio"] = motivo_cambio
                    df.to_excel(EXCEL_RESERVAS_LOCAL, index=False)
            except Exception as e:
                st.error(f"Error al actualizar la reserva local: {e}")


# Cargar datos
config_actual = cargar_configuracion_sistema()
registro_activo = config_actual.get("registro_habilitado", True)

anio_actual = datetime.date.today().year
feriados_arg = holidays.Argentina(years=[anio_actual, anio_actual + 1])

df_reservas_historico = cargar_reservas_existentes()
fechas_ocupadas = obtener_fechas_ocupadas(df_reservas_historico)

df_escuelas = cargar_base_escuelas()
df_personas = cargar_base_personas()


def generar_fechas_disponibles(inicio, fin, feriados, ocupadas):
    libres = []
    dia_actual = inicio
    while dia_actual <= fin:
        if dia_actual.weekday() < 5 and dia_actual.weekday() != 2:
            if dia_actual not in feriados and dia_actual not in ocupadas:
                libres.append(dia_actual)
        dia_actual += datetime.timedelta(days=1)
    return libres


def formatear_fecha_espanol(fecha):
    dias = [
        "Lunes",
        "Martes",
        "Miércoles",
        "Jueves",
        "Viernes",
        "Sábado",
        "Domingo",
    ]
    meses = [
        "",
        "Enero",
        "Febrero",
        "Marzo",
        "Abril",
        "Mayo",
        "Junio",
        "Julio",
        "Agosto",
        "Septiembre",
        "Octubre",
        "Noviembre",
        "Diciembre",
    ]
    return f"{dias[fecha.weekday()]} {fecha.day} de {meses[fecha.month]}"


# Sidebar de administración
with st.sidebar:
    st.write("### ⚙️ Soporte")
    if usando_google_sheets():
        st.success("☁️ Google Drive Conectado")
    else:
        st.info("💻 Almacenamiento Local Activo")

    with st.expander("Acceso de Sistema", expanded=False):
        pass_admin = st.text_input("Clave de Seguridad:", type="password")
        if pass_admin == "ariel":
            st.session_state.admin_autenticado = True
            st.success("Acceso Habilitado")
        else:
            st.session_state.admin_autenticado = False

    vista_admin = False
    if st.session_state.admin_autenticado:
        st.divider()
        st.write("🛠️ **Opciones de Admin**")
        vista_admin = st.checkbox("Ver Panel de Administración")

# ================= VISTA DE ADMINISTRADOR =================
if st.session_state.admin_autenticado and vista_admin:
    st.title("🔒 Panel de Administración")
    st.write("Gestione los archivos cargados por cada institucion y acceda al reporte de asignación de turnos.")

    st.markdown('<div class="step-card step-card-blue">', unsafe_allow_html=True)
    st.subheader("🌐 Disponibilidad del Formulario en Internet")
    nuevo_estado = st.toggle("Habilitar inscripcion", value=registro_activo)
    if nuevo_estado != registro_activo:
        config_actual["registro_habilitado"] = nuevo_estado
        guardar_configuracion_sistema(config_actual)
        st.success(f"Formulario de registro {'HABILITADO' if nuevo_estado else 'DESHABILITADO'} en internet con éxito.")
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        st.markdown('<div class="step-card step-card-purple">', unsafe_allow_html=True)
        st.subheader("📤 Cargar / Actualizar Base de Escuelas")
        archivo_subido_esc = st.file_uploader("Seleccione base_escuelas.xlsx", type=["xlsx"], key="uploader_escuelas")
        if archivo_subido_esc is not None:
            id_archivo_esc = f"{archivo_subido_esc.name}_{archivo_subido_esc.size}"
            if st.session_state.escuelas_procesadas != id_archivo_esc:
                try:
                    test_df = pd.read_excel(archivo_subido_esc)
                    with open(EXCEL_ESCUELAS, "wb") as f:
                        f.write(archivo_subido_esc.getbuffer())
                    st.session_state.escuelas_procesadas = id_archivo_esc
                    st.success("¡Base de escuelas guardada con éxito!")
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")
        st.markdown("</div>", unsafe_allow_html=True)

    with col_u2:
        st.markdown('<div class="step-card step-card-purple">', unsafe_allow_html=True)
        st.subheader("📤 Cargar / Actualizar Base de Personas")
        archivo_subido_per = st.file_uploader("Seleccione personas.xlsx", type=["xlsx"], key="uploader_personas")
        if archivo_subido_per is not None:
            id_archivo_per = f"{archivo_subido_per.name}_{archivo_subido_per.size}"
            if st.session_state.personas_procesadas != id_archivo_per:
                try:
                    test_df = pd.read_excel(archivo_subido_per)
                    with open(EXCEL_PERSONAS, "wb") as f:
                        f.write(archivo_subido_per.getbuffer())
                    st.session_state.personas_procesadas = id_archivo_per
                    st.success("¡Base de directores guardada con éxito!")
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="step-card step-card-emerald">', unsafe_allow_html=True)
    st.subheader("📥 Registro Histórico y Descargas")
    if not df_reservas_historico.empty:
        if usando_google_sheets():
            st.info("🟢 Los datos mostrados corresponden a la planilla de **Google Sheets** en tiempo real.")
        else:
            st.warning("⚠️ Los datos mostrados se encuentran guardados de forma **Local**.")

        st.dataframe(df_reservas_historico, use_container_width=True)
        buffer_excel = io.BytesIO()
        df_reservas_historico.to_excel(buffer_excel, index=False)
        st.download_button(
            label="📥 Descargar Excel de Reservas Sincronizado",
            data=buffer_excel.getvalue(),
            file_name=f"registro_reservas_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("No se registran reservas agendadas todavía.")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="step-card" style="border-top: 5px solid #ef4444; background-color: #fef2f2;">', unsafe_allow_html=True)
    st.subheader("⚠️ Zona de Peligro: Reiniciar Calendario")
    confirmar_reinicio = st.checkbox("Confirmo que deseo vaciar todo el registro de reservas", key="check_reinicio")
    if st.button("🗑️ Eliminar todas las reservas del Excel", disabled=not confirmar_reinicio):
        if usando_google_sheets():
            try:
                hoja = conectar_google_sheets()
                if hoja:
                    hoja.clear()
                    hoja.append_row(COLUMNAS_SISTEMA)
                    st.success("¡La planilla de Google Sheets ha sido vaciada con éxito!")
                    st.cache_data.clear()
                    st.rerun()
            except Exception as e:
                st.error(f"Error al vaciar la planilla de Google Sheets: {e}")
        else:
            if os.path.exists(EXCEL_RESERVAS_LOCAL):
                try:
                    os.remove(EXCEL_RESERVAS_LOCAL)
                    st.success("¡El archivo de reservas local ha sido eliminado con éxito!")
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al eliminar el archivo local: {e}")
    st.markdown("</div>", unsafe_allow_html=True)

# ================= VISTA DE DIRECTORES (PÚBLICA) =================
else:
    st.markdown(
        '<h1 style="text-align: center; color: #1e293b !important; margin-bottom: 5px;">📅 Sistema de Reserva de Turnos</h1>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p style="text-align: center; color: #64748b; font-size: 1.1rem; margin-bottom: 30px;">Bienvenido al sistema de Reserva de Turnos para la jornada de RCP.</p>',
        unsafe_allow_html=True,
    )

    if not registro_activo:
        st.error("⚠️ **Sistema Desactivado:** El período de agendamiento se encuentra inhabilitado en este momento.")

    elif st.session_state.reserva_exitosa is not None:
        r = st.session_state.reserva_exitosa

        st.markdown(
            f"""
            <div style="background-color: #ffffff; border: 2px solid #22c55e; border-radius: 16px; padding: 32px; max-width: 650px; margin: 0 auto 24px auto; box-shadow: 0 10px 25px -5px rgba(34, 197, 94, 0.1);">
                <div style="text-align: center; margin-bottom: 20px;">
                    <span style="font-size: 3.5rem;">🎉</span>
                    <h2 style="color: #15803d !important; margin-top: 10px;">¡Reserva Confirmada Exitosamente!</h2>
                </div>
                <hr style="border: 0; border-top: 1px solid #e2e8f0; margin-bottom: 20px;">
                <div style="font-size: 1rem; color: #1e293b; line-height: 1.7;">
                    <strong>Establecimiento:</strong> {r['Escuela']}<br>
                    <strong>CUE:</strong> {r['CUE']}<br>
                    <strong>Director Solicitante:</strong> {r['Director']}<br>
                    <strong>Teléfono Contacto:</strong> {r['Telefono_Contacto']}<br>
                    <strong>Total Alumnos Registrados:</strong> {r['Total_Alumnos']} alumnos.<br>
                    <hr style="border: 0; border-top: 1px dashed #cbd5e1; margin: 15px 0;">
                    <div style="background-color: #f0fdf4; border: 1px solid #bbf7d0; padding: 14px; border-radius: 10px; text-align: center; font-size: 1.1rem; color: #166534; font-weight: 700;">
                        📅 Día Reservado: {int(float(r['Dia_Reservado'])):02d} / {int(float(r['Mes_Reservado'])):02d} / {int(float(r['Anio_Reservado']))}
                    </div>
                </div>
            </div>
        """,
            unsafe_allow_html=True,
        )

        pdf_comprobante_bytes = generar_pdf_comprobante(r)
        col_c1, col_c2 = st.columns(2)

        with col_c1:
            st.download_button(
                label="📄 Descargar Comprobante (PDF)",
                data=pdf_comprobante_bytes,
                file_name=f"Comprobante_Reserva_{r['CUE']}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

        with col_c2:
            if st.button("🏁 Finalizar y Cerrar Sesión", use_container_width=True):
                st.session_state.reserva_exitosa = None
                st.rerun()

    else:
        if df_escuelas.empty:
            st.warning("⚠️ No hay base de escuelas cargada en el sistema.")
        elif df_personas.empty:
            st.warning("⚠️ El padrón de autoridades no se encuentra cargado.")
        else:
            # --- PASO 1: Identificación del Establecimiento (AZUL) ---
            st.markdown('<div class="step-card step-card-blue">', unsafe_allow_html=True)
            st.markdown('<div class="step-title title-blue">🏫 1. Identificación del Establecimiento Educativo</div>', unsafe_allow_html=True)
            
            cue_ingresado = st.text_input(
                "Ingrese el CUE de la institución para comenzar:",
                key="cue_input_user",
                placeholder="Ej: 7000123",
            ).strip()

            nombre_escuela = ""
            modalidad = ""
            departamento = ""
            domicilio = ""
            escuela_valida = False
            cue_ya_reservado = False
            reserva_previa = None

            if cue_ingresado:
                cue_limpio = normalizar_texto(cue_ingresado)

                if "CUE" in df_reservas_historico.columns:
                    df_res_temp = df_reservas_historico.copy()
                    df_res_temp["CUE_norm"] = df_res_temp["CUE"].apply(normalizar_texto)
                    coincidencias_res = df_res_temp[df_res_temp["CUE_norm"] == cue_limpio]
                    if not coincidencias_res.empty:
                        cue_ya_reservado = True
                        reserva_previa = coincidencias_res.iloc[-1].to_dict()

                if cue_ya_reservado:
                    dia_res = int(float(reserva_previa["Dia_Reservado"]))
                    mes_res = int(float(reserva_previa["Mes_Reservado"]))
                    anio_res = int(float(reserva_previa["Anio_Reservado"]))

                    st.info(
                        f"📌 **Turno Actual Agendado:** La institución **{reserva_previa['Escuela']}** (CUE: {cue_limpio}) ya tiene un turno reservado para el día **{dia_res:02d}/{mes_res:02d}/{anio_res}**."
                    )

                    pdf_existente = generar_pdf_comprobante(reserva_previa)
                    st.download_button(
                        label="📥 Re-descargar Comprobante de Turno (PDF)",
                        data=pdf_existente,
                        file_name=f"Comprobante_Reserva_{cue_limpio}.pdf",
                        mime="application/pdf",
                        key="btn_redescargar_pdf",
                    )

                    st.divider()

                    st.markdown("### 🔄 Cambiar Turno Asignado por Otro Disponible")
                    st.write("Para reprogramar el turno, valide el DNI del titular registrado:")

                    dni_registrado = normalizar_texto(reserva_previa.get("DNI_Director", ""))

                    dni_validante = st.text_input(
                        "Ingrese el DNI del titular registrado:",
                        placeholder="Ej: 22333444",
                        key="input_dni_validador",
                    ).strip()

                    dni_valido = False
                    if dni_validante:
                        if normalizar_texto(dni_validante) == dni_registrado:
                            st.success("✅ DNI verificado correctamente.")
                            dni_valido = True
                        else:
                            st.error("❌ El DNI ingresado no coincide con el DNI del titular registrado.")

                    motivo_cambio = st.text_area(
                        "Ingrese un motivo breve del cambio de turno:",
                        placeholder="Ej: Superposición con jornada institucional / Evaluación...",
                        key="input_motivo_cambio",
                    ).strip()

                    fecha_inicio = datetime.date(anio_actual, 8, 1)
                    fecha_limite = datetime.date(anio_actual, 11, 30)

                    lista_fechas_libres = generar_fechas_disponibles(
                        fecha_inicio, fecha_limite, feriados_arg, fechas_ocupadas
                    )

                    if len(lista_fechas_libres) > 0:
                        opciones_combo_cambio = {formatear_fecha_espanol(f): f for f in lista_fechas_libres}
                        nueva_fecha_sel = st.selectbox(
                            "Seleccione la NUEVA fecha disponible:",
                            options=list(opciones_combo_cambio.keys()),
                            index=0,
                            key="combo_cambio_turno",
                        )
                        fecha_nueva_obj = opciones_combo_cambio[nueva_fecha_sel]

                        puede_reprogramar = dni_valido and bool(motivo_cambio)

                        if st.button("🔄 Confirmar y Reprogramar Turno", disabled=not puede_reprogramar):
                            actualizar_reserva_existente(cue_limpio, fecha_nueva_obj, motivo_cambio)

                            reserva_actualizada = reserva_previa.copy()
                            reserva_actualizada["Dia_Reservado"] = int(fecha_nueva_obj.day)
                            reserva_actualizada["Mes_Reservado"] = int(fecha_nueva_obj.month)
                            reserva_actualizada["Anio_Reservado"] = int(fecha_nueva_obj.year)
                            reserva_actualizada["Motivo_Cambio"] = motivo_cambio
                            reserva_actualizada["Fecha_Registro"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " (Modificado)"

                            st.session_state.reserva_exitosa = reserva_actualizada
                            st.cache_data.clear()
                            st.rerun()
                    else:
                        st.error("🔴 No hay nuevas fechas disponibles para cambiar su turno.")

                else:
                    coincidencia_esc = df_escuelas[df_escuelas["CUE"] == cue_limpio]

                    if not coincidencia_esc.empty:
                        nombre_escuela = coincidencia_esc.iloc[0]["Nombre_Escuela"]
                        modalidad = coincidencia_esc.iloc[0]["Modalidad_Oferta"]
                        departamento = coincidencia_esc.iloc[0]["Departamento"]
                        domicilio = coincidencia_esc.iloc[0]["Domicilio"]
                        escuela_valida = True

                        st.markdown(
                            f"""
                            <div class="info-pill info-pill-success">
                                <strong>🏫 Escuela Verificada:</strong> {nombre_escuela}<br>
                                <strong>Modalidad:</strong> {modalidad} &nbsp;|&nbsp; <strong>Departamento:</strong> {departamento}
                            </div>
                        """,
                            unsafe_allow_html=True,
                        )
                    else:
                        st.error("❌ El CUE ingresado no figura registrado en el sistema escolar.")
            st.markdown("</div>", unsafe_allow_html=True)

            # --- PASO 2: Datos del Solicitante (VIOLETA) --- Solo si el CUE es válido y no tiene reserva
            if escuela_valida and not cue_ya_reservado:
                st.markdown('<div class="step-card step-card-purple">', unsafe_allow_html=True)
                st.markdown('<div class="step-title title-purple">👤 2. Datos del Solicitante (Autoridad)</div>', unsafe_allow_html=True)
                
                dni_ingresado = st.text_input(
                    "Ingrese su DNI (sin puntos):",
                    key="dni_input_user",
                    placeholder="Ej: 22333444",
                ).strip()

                nombre_director = ""
                telefono_predicho = ""
                persona_valida = False

                if dni_ingresado:
                    dni_limpio = normalizar_texto(dni_ingresado)
                    coincidencia_per = df_personas[df_personas["DNI"] == dni_limpio]

                    if not coincidencia_per.empty:
                        nombre_director = coincidencia_per.iloc[0]["Apellido_Nombre"]
                        telefono_predicho = coincidencia_per.iloc[0]["Telefono"]
                        persona_valida = True

                        st.markdown(
                            f"""
                            <div class="info-pill info-pill-success">
                                <strong>👤 Autoridad Verificada:</strong> {nombre_director}
                            </div>
                        """,
                            unsafe_allow_html=True,
                        )
                        telefono_final = st.text_input(
                            "Verifique o edite su Teléfono de Contacto:",
                            value=telefono_predicho,
                            placeholder="Ej: 2645551234",
                        )
                    else:
                        st.warning("⚠️ El DNI no figura en el padrón precargado. Complete sus datos manualmente:")

                        col_m1, col_m2 = st.columns(2)
                        with col_m1:
                            apellido_manual = st.text_input("Ingrese Apellido/s:", placeholder="Ej: PÉREZ").strip().upper()
                        with col_m2:
                            nombre_manual = st.text_input("Ingrese Nombre/s:", placeholder="Ej: Juan Carlos").strip()

                        telefono_manual = st.text_input("Ingrese Teléfono de Contacto:", placeholder="Ej: 2645551234").strip()

                        if apellido_manual and nombre_manual:
                            nombre_director = f"{apellido_manual}, {nombre_manual}"
                            persona_valida = True

                        telefono_final = telefono_manual
                else:
                    telefono_final = ""
                st.markdown("</div>", unsafe_allow_html=True)

                # --- PASO 3: Relevamiento de Cursos (ÁMBAR) --- Solo si la persona está validada
                if persona_valida:
                    st.markdown('<div class="step-card step-card-amber">', unsafe_allow_html=True)
                    st.markdown('<div class="step-title title-amber">📊 3. Relevamiento de Cursos y Alumnos</div>', unsafe_allow_html=True)

                    st.markdown(
                        '<div class="atencion-box">⚠️ Debe declarar la matrícula obligatoria de alumnos para cada división registrada.</div>',
                        unsafe_allow_html=True,
                    )

                    estructura_opciones = [
                        "Seleccione una opción...",
                        "5° y 6° Año (Secundaria Orientada / Ciclo Superior Común)",
                        "6° y 7° Año (Escuelas Técnicas o de Modalidades Profesionales)",
                    ]

                    estructura_seleccionada = st.selectbox(
                        "Estructura del plan de estudios de la institución:",
                        options=estructura_opciones,
                        index=0,
                        key="estructura_plan_estudios",
                    )

                    datos_cursos = {}
                    total_alumnos_declarados = 0
                    estructura_valida_plan = False
                    hay_campos_alumnos_vacios_o_cero = False
                    ano_bajo, ano_alto = "", ""

                    if estructura_seleccionada != "Seleccione una opción...":
                        if "5° y 6°" in estructura_seleccionada:
                            ano_bajo, ano_alto = "5° Año", "6° Año"
                        else:
                            ano_bajo, ano_alto = "6° Año", "7° Año"

                        # --- CARGA SECUENCIAL DEL PRIMER AÑO ---
                        st.markdown(f"### 📘 Carga para **{ano_bajo}**")
                        cant_div_bajo = st.number_input(
                            f"Cantidad total de divisiones en {ano_bajo}:",
                            min_value=0,
                            max_value=15,
                            value=0,
                            step=1,
                            key="div_bajo",
                        )

                        divs_bajo = []
                        if cant_div_bajo > 0:
                            for i in range(cant_div_bajo):
                                col_i1, col_i2 = st.columns([1, 2])
                                with col_i1:
                                    seccion = st.text_input(
                                        f"División {i+1}:",
                                        value=chr(65 + i) if i < 26 else str(i + 1),
                                        key=f"sec_{ano_bajo}_{i}",
                                    ).strip()
                                with col_i2:
                                    alumnos = st.number_input(
                                        f"Cantidad de alumnos en Div. {seccion}:",
                                        min_value=1,
                                        max_value=100,
                                        value=None,  # Campo sin inicializar en cero
                                        step=1,
                                        placeholder="Ingrese N° de alumnos (Obligatorio)",
                                        key=f"alu_{ano_bajo}_{i}",
                                    )

                                if alumnos is None or alumnos == 0:
                                    hay_campos_alumnos_vacios_o_cero = True
                                else:
                                    total_alumnos_declarados += alumnos

                                divs_bajo.append({"division": seccion, "alumnos": alumnos})
                        datos_cursos[ano_bajo] = divs_bajo

                        st.divider()

                        # --- CARGA SECUENCIAL DEL SEGUNDO AÑO ---
                        st.markdown(f"### 📗 Carga para **{ano_alto}**")
                        cant_div_alto = st.number_input(
                            f"Cantidad total de divisiones en {ano_alto}:",
                            min_value=0,
                            max_value=15,
                            value=0,
                            step=1,
                            key="div_alto",
                        )

                        divs_alto = []
                        if cant_div_alto > 0:
                            for i in range(cant_div_alto):
                                col_j1, col_j2 = st.columns([1, 2])
                                with col_j1:
                                    seccion = st.text_input(
                                        f"División {i+1}:",
                                        value=chr(65 + i) if i < 26 else str(i + 1),
                                        key=f"sec_{ano_alto}_{i}",
                                    ).strip()
                                with col_j2:
                                    alumnos = st.number_input(
                                        f"Cantidad de alumnos en Div. {seccion}:",
                                        min_value=1,
                                        max_value=100,
                                        value=None,  # Campo sin inicializar en cero
                                        step=1,
                                        placeholder="Ingrese N° de alumnos (Obligatorio)",
                                        key=f"alu_{ano_alto}_{i}",
                                    )

                                if alumnos is None or alumnos == 0:
                                    hay_campos_alumnos_vacios_o_cero = True
                                else:
                                    total_alumnos_declarados += alumnos

                                divs_alto.append({"division": seccion, "alumnos": alumnos})
                        datos_cursos[ano_alto] = divs_alto

                        # Validaciones globales del Paso 3
                        if cant_div_bajo == 0 and cant_div_alto == 0:
                            st.warning("Debe indicar al menos 1 división en alguno de los dos años para continuar.")
                        elif hay_campos_alumnos_vacios_o_cero:
                            st.error("🚫 **Atención:** Debe completar obligatoriamente el número de alumnos para todas las divisiones indicadas.")
                        else:
                            estructura_valida_plan = True
                            st.success(f"✅ Relevamiento completado correctamente. Total de alumnos: **{total_alumnos_declarados}**.")

                    st.markdown("</div>", unsafe_allow_html=True)

                    # --- PASO 4: Selección de Turno (ESMERALDA) --- Solo si las divisiones están bien cargadas
                    if estructura_valida_plan:
                        st.markdown('<div class="step-card step-card-emerald">', unsafe_allow_html=True)
                        st.markdown('<div class="step-title title-emerald">📅 4. Selección de Turno Disponible</div>', unsafe_allow_html=True)

                        fecha_inicio = datetime.date(anio_actual, 8, 1)
                        fecha_limite = datetime.date(anio_actual, 11, 30)

                        lista_fechas_libres = generar_fechas_disponibles(
                            fecha_inicio, fecha_limite, feriados_arg, fechas_ocupadas
                        )

                        es_valida = False
                        fecha_seleccionada = None

                        if len(lista_fechas_libres) > 0:
                            opciones_combo = {formatear_fecha_espanol(f): f for f in lista_fechas_libres}

                            seleccion_usuario = st.selectbox(
                                "Seleccione una fecha disponible (Lunes, Martes, Jueves o Viernes):",
                                options=list(opciones_combo.keys()),
                                index=0,
                                key="combo_fechas_libres",
                            )

                            fecha_seleccionada = opciones_combo[seleccion_usuario]
                            es_valida = True
                            st.info(f"🎉 Elegiste el turno del día **{fecha_seleccionada.strftime('%d/%m/%Y')}**.")
                        else:
                            st.error("🔴 Lo sentimos, ya no quedan turnos disponibles en el rango de Agosto a Noviembre.")

                        st.divider()

                        formulario_listo = es_valida and bool(telefono_final.strip())

                        if st.button(" Confirmar y Registrar Agenda", disabled=not formulario_listo):
                            bajo_desc = ", ".join([f"Div {x['division']} ({x['alumnos']} al.)" for x in datos_cursos.get(ano_bajo, []) if x['alumnos']])
                            alto_desc = ", ".join([f"Div {x['division']} ({x['alumnos']} al.)" for x in datos_cursos.get(ano_alto, []) if x['alumnos']])
                            resumen_matricula = f"{ano_bajo}: [{bajo_desc}] | {ano_alto}: [{alto_desc}]"

                            datos_reserva = {
                                "CUE": normalizar_texto(cue_ingresado),
                                "Escuela": nombre_escuela,
                                "Modalidad_Oferta": modalidad,
                                "Departamento": departamento,
                                "Domicilio": domicilio,
                                "DNI_Director": normalizar_texto(dni_ingresado),
                                "Director": nombre_director,
                                "Telefono_Contacto": telefono_final.strip(),
                                "Estructura_Declarada": estructura_seleccionada,
                                "Detalle_Divisiones_Alumnos": resumen_matricula,
                                "Total_Alumnos": int(total_alumnos_declarados),
                                "Dia_Reservado": int(fecha_seleccionada.day),
                                "Mes_Reservado": int(fecha_seleccionada.month),
                                "Anio_Reservado": int(fecha_seleccionada.year),
                                "Fecha_Registro": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "Motivo_Cambio": "",
                            }

                            guardar_reserva(datos_reserva)
                            st.session_state.reserva_exitosa = datos_reserva
                            st.rerun()

                        if not telefono_final.strip():
                            st.warning("Debe ingresar un número telefónico de contacto para habilitar la confirmación.")
                            
                        st.markdown("</div>", unsafe_allow_html=True)
