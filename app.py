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

# Configuración de página de Streamlit
st.set_page_config(
    page_title="Calendario Excluyente - Gestión Escolar",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
        /* Estilo general y fondo neutro premium */
        .stApp {
            background-color: #f8fafc;
        }
        
        /* Contenedores estilizados tipo tarjetas */
        .custom-card {
            background-color: #ffffff;
            padding: 24px;
            border-radius: 12px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
            margin-bottom: 24px;
            border: 1px solid #e2e8f0;
        }
        
        /* Títulos limpios y profesionales */
        h1, h2, h3 {
            color: #1e293b !important;
            font-family: 'Inter', sans-serif;
            font-weight: 700 !important;
        }
        
        /* Barra lateral con paleta oscura elegante */
        section[data-testid="stSidebar"] {
            background-color: #0f172a !important;
        }
        section[data-testid="stSidebar"] * {
            color: #f1f5f9 !important;
        }
        
        /* Botones del sistema modernizados */
        .stButton>button {
            background-color: #0284c7 !important;
            color: white !important;
            border-radius: 8px !important;
            padding: 10px 24px !important;
            font-weight: 600 !important;
            border: none !important;
            transition: all 0.3s ease !important;
            box-shadow: 0 4px 6px -1px rgba(2, 132, 199, 0.3) !important;
        }
        
        .stButton>button:hover {
            background-color: #0369a1 !important;
            transform: translateY(-1px) !important;
            box-shadow: 0 6px 8px -1px rgba(2, 132, 199, 0.4) !important;
        }
        
        /* Estilo sutil para inputs bloqueados */
        input:disabled {
            background-color: #f1f5f9 !important;
            color: #475569 !important;
            opacity: 1 !important;
        }
        
        /* Tarjeta de información destacada */
        .info-pill-container {
            background-color: #f0fdf4;
            border: 1px solid #bbf7d0;
            padding: 16px;
            border-radius: 8px;
            margin-top: 10px;
        }
        .info-pill-title {
            color: #166534;
            font-weight: bold;
            font-size: 1.1rem;
            margin-bottom: 8px;
        }
        .info-pill-text {
            color: #14532d;
            font-size: 0.95rem;
            line-height: 1.5;
        }
        
        /* Alerta visual para campos requeridos del plan de estudios */
        .atencion-box {
            background-color: #fff7ed;
            border: 1px solid #fed7aa;
            padding: 12px;
            border-radius: 6px;
            color: #c2410c;
            font-size: 0.95rem;
            margin-bottom: 15px;
            font-weight: 500;
        }
    </style>
""", unsafe_allow_html=True)

if "escuelas_procesadas" not in st.session_state:
    st.session_state.escuelas_procesadas = None
if "personas_procesadas" not in st.session_state:
    st.session_state.personas_procesadas = None
if "admin_autenticado" not in st.session_state:
    st.session_state.admin_autenticado = False
if "reserva_exitosa" not in st.session_state:
    st.session_state.reserva_exitosa = None

# --- DETECCIÓN DE BASE DE DATOS ACTIVA ---
def usando_google_sheets():
    if not gsheets_librerias_listas:
        return False
    try:
        return "gcp_service_account" in st.secrets and "spreadsheet_url" in st.secrets
    except Exception:
        return False

def conectar_google_sheets():
    claves = dict(st.secrets["gcp_service_account"])
    if "private_key" in claves:
        claves["private_key"] = claves["private_key"].replace("\\n", "\n")
        
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    credenciales = Credentials.from_service_account_info(claves, scopes=scopes)
    cliente = gspread.authorize(credenciales)
    planilla = cliente.open_by_url(st.secrets["spreadsheet_url"])
    return planilla.sheet1

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
            df.columns = df.columns.str.strip().str.upper().str.replace(' ', '_')
            
            col_cue = [c for c in df.columns if 'CUE' in c]
            col_nombre = [c for c in df.columns if 'NOM' in c or 'ESC' in c]
            col_mod = [c for c in df.columns if 'MOD' in c or 'OFER' in c]
            col_depto = [c for c in df.columns if 'DEP' in c]
            col_dom = [c for c in df.columns if 'DOM' in c or 'DIR' in c]
            
            mapping = {}
            if col_cue: mapping[col_cue[0]] = 'CUE'
            if col_nombre: mapping[col_nombre[0]] = 'Nombre_Escuela'
            if col_mod: mapping[col_mod[0]] = 'Modalidad_Oferta'
            if col_depto: mapping[col_depto[0]] = 'Departamento'
            if col_dom: mapping[col_dom[0]] = 'Domicilio'
            
            df = df.rename(columns=mapping)
            df['CUE'] = df['CUE'].apply(normalizar_texto)
            
            for col in ['Modalidad_Oferta', 'Departamento', 'Domicilio']:
                if col not in df.columns:
                    df[col] = "No especificado"
                else:
                    df[col] = df[col].fillna("No especificado").astype(str).str.strip()
                    
            return df[['CUE', 'Nombre_Escuela', 'Modalidad_Oferta', 'Departamento', 'Domicilio']].drop_duplicates()
        except Exception as e:
            st.error(f"Error al leer la base de escuelas: {e}")
            return pd.DataFrame(columns=["CUE", "Nombre_Escuela", "Modalidad_Oferta", "Departamento", "Domicilio"])
    return pd.DataFrame(columns=["CUE", "Nombre_Escuela", "Modalidad_Oferta", "Departamento", "Domicilio"])

@st.cache_data
def cargar_base_personas():
    if os.path.exists(EXCEL_PERSONAS):
        try:
            df = pd.read_excel(EXCEL_PERSONAS)
            df.columns = df.columns.str.strip().str.upper().str.replace(' ', '_')
            
            col_dni = [c for c in df.columns if 'DNI' in c or 'DOC' in c]
            col_apellido = [c for c in df.columns if 'APE' in c]
            col_nombre = [c for c in df.columns if 'NOM' in c and 'ESC' not in c]
            col_tel = [c for c in df.columns if 'TEL' in c or 'CEL' in c]
            
            mapping = {}
            if col_dni: mapping[col_dni[0]] = 'DNI'
            if col_apellido: mapping[col_apellido[0]] = 'Apellido'
            if col_nombre: mapping[col_nombre[0]] = 'Nombre'
            if col_tel: mapping[col_tel[0]] = 'Telefono'
            
            df = df.rename(columns=mapping)
            
            if 'DNI' in df.columns and 'Apellido' in df.columns and 'Nombre' in df.columns:
                df['DNI'] = df['DNI'].apply(normalizar_texto)
                df['Apellido'] = df['Apellido'].fillna("").astype(str).str.strip()
                df['Nombre'] = df['Nombre'].fillna("").astype(str).str.strip()
                df['Apellido_Nombre'] = df['Apellido'] + ", " + df['Nombre']
                df['Apellido_Nombre'] = df['Apellido_Nombre'].str.strip(", ")
                
                if 'Telefono' not in df.columns:
                    df['Telefono'] = ""
                else:
                    df['Telefono'] = df['Telefono'].apply(normalizar_texto)
                    
                return df[['DNI', 'Apellido_Nombre', 'Telefono']].drop_duplicates()
            else:
                st.error("El Excel de personas debe tener columnas identificables para DNI, APELLIDO y NOMBRE.")
                return pd.DataFrame(columns=["DNI", "Apellido_Nombre", "Telefono"])
        except Exception as e:
            st.error(f"Error al leer la base de personas: {e}")
            return pd.DataFrame(columns=["DNI", "Apellido_Nombre", "Telefono"])
    return pd.DataFrame(columns=["DNI", "Apellido_Nombre", "Telefono"])

# Listado de columnas oficiales requeridas
COLUMNAS_SISTEMA = [
    "CUE", "Escuela", "Modalidad_Oferta", "Departamento", "Domicilio", 
    "DNI_Director", "Director", "Telefono_Contacto", "Estructura_Declarada", 
    "Detalle_Divisiones_Alumnos", "Total_Alumnos", "Dia_Reservado", 
    "Mes_Reservado", "Anio_Reservado", "Fecha_Registro"
]

def cargar_reservas_existentes():
    if usando_google_sheets():
        try:
            hoja = conectar_google_sheets()
            valores = hoja.get_all_values()
            if valores and len(valores) > 1:
                return pd.DataFrame(valores[1:], columns=valores[0])
        except Exception:
            pass
    else:
        if os.path.exists(EXCEL_RESERVAS_LOCAL):
            try:
                return pd.read_excel(EXCEL_RESERVAS_LOCAL)
            except Exception:
                pass
    return pd.DataFrame(columns=COLUMNAS_SISTEMA)

def obtener_fechas_ocupadas(df_reservas):
    fechas = []
    if not df_reservas.empty and 'Dia_Reservado' in df_reservas.columns:
        for _, row in df_reservas.iterrows():
            try:
                d = int(float(row['Dia_Reservado']))
                m = int(float(row['Mes_Reservado']))
                a = int(float(row['Anio_Reservado']))
                fechas.append(datetime.date(a, m, d))
            except Exception:
                continue
    return set(fechas)

def guardar_reserva(datos):
    if usando_google_sheets():
        try:
            hoja = conectar_google_sheets()
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

# Cargar configuración activa y datos generales
config_actual = cargar_configuracion_sistema()
registro_activo = config_actual.get("registro_habilitado", True)

anio_actual = datetime.date.today().year
feriados_arg = holidays.Argentina(years=[anio_actual, anio_actual + 1])

# Cargamos el histórico de reservas
df_reservas_historico = cargar_reservas_existentes()
fechas_ocupadas = obtener_fechas_ocupadas(df_reservas_historico)

df_escuelas = cargar_base_escuelas()
df_personas = cargar_base_personas()

def generar_fechas_disponibles(inicio, fin, feriados, ocupadas):
    libres = []
    dia_actual = inicio
    while dia_actual <= fin:
        if dia_actual.weekday() < 5:
            if dia_actual not in feriados and dia_actual not in ocupadas:
                libres.append(dia_actual)
        dia_actual += datetime.timedelta(days=1)
    return libres

def formatear_fecha_espanol(fecha):
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    meses = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    return f"{dias[fecha.weekday()]} {fecha.day} de {meses[fecha.month]}"

# Panel de administración oculto
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
    st.title("🔒 Panel de Control de Administración")
    st.write("Gestione los archivos cargados del sistema escolar y acceda al reporte de asignación de turnos.")
    
    st.markdown('<div class="custom-card">', unsafe_allow_html=True)
    st.subheader("🌐 Disponibilidad del Formulario en Internet")
    nuevo_estado = st.toggle("Habilitar Registro Público de Reservas", value=registro_activo)
    if nuevo_estado != registro_activo:
        config_actual["registro_habilitado"] = nuevo_estado
        guardar_configuracion_sistema(config_actual)
        st.success(f"Formulario de registro {'HABILITADO' if nuevo_estado else 'DESHABILITADO'} en internet con éxito.")
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        st.markdown('<div class="custom-card">', unsafe_allow_html=True)
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
        st.markdown('</div>', unsafe_allow_html=True)

    with col_u2:
        st.markdown('<div class="custom-card">', unsafe_allow_html=True)
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
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="custom-card">', unsafe_allow_html=True)
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
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("No se registran reservas agendadas todavía.")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="custom-card" style="border: 1px solid #fecaca; background-color: #fef2f2;">', unsafe_allow_html=True)
    st.subheader("⚠️ Zona de Peligro: Reiniciar Calendario")
    confirmar_reinicio = st.checkbox("Confirmo que deseo vaciar todo el registro de reservas", key="check_reinicio")
    if st.button("🗑️ Eliminar todas las reservas del Excel", disabled=not confirmar_reinicio):
        if usando_google_sheets():
            try:
                hoja = conectar_google_sheets()
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
    st.markdown('</div>', unsafe_allow_html=True)

# ================= VISTA DE DIRECTORES (PÚBLICA) =================
else:
    st.markdown('<h1 style="text-align: center; color: #0284c7 !important; margin-bottom: 5px;">📅 Sistema de Reserva de Turnos</h1>', unsafe_allow_html=True)
    st.markdown('<p style="text-align: center; color: #64748b; font-size: 1.1rem; margin-bottom: 25px;">Agende la jornada institucional de su establecimiento escolar sin superposiciones.</p>', unsafe_allow_html=True)

    if not registro_activo:
        st.error("⚠️ **Sistema Desactivado:** El período de agendamiento se encuentra inhabilitado en este momento.")
        
    elif st.session_state.reserva_exitosa is not None:
        r = st.session_state.reserva_exitosa
        st.markdown(f"""
            <div style="background-color: #ffffff; border: 2px solid #22c55e; border-radius: 12px; padding: 30px; max-width: 650px; margin: 0 auto 30px auto;">
                <div style="text-align: center; margin-bottom: 20px;">
                    <span style="font-size: 3.5rem;">🎉</span>
                    <h2 style="color: #15803d !important; margin-top: 10px;">¡Reserva Confirmada Exitosamente!</h2>
                </div>
                <hr style="border: 0; border-top: 1px solid #e2e8f0; margin-bottom: 20px;">
                <div style="font-size: 1rem; color: #1e293b; line-height: 1.6;">
                    <strong>Establecimiento:</strong> {r['Escuela']}<br>
                    <strong>CUE:</strong> {r['CUE']}<br>
                    <strong>Director Solicitante:</strong> {r['Director']}<br>
                    <strong>Teléfono Contacto:</strong> {r['Telefono_Contacto']}<br>
                    <strong>Total Alumnos Registrados:</strong> {r['Total_Alumnos']} alumnos.<br>
                    <hr style="border: 0; border-top: 1px dashed #cbd5e1; margin: 15px 0;">
                    <div style="background-color: #f0fdf4; border: 1px solid #bbf7d0; padding: 12px; border-radius: 8px; text-align: center; font-size: 1.1rem; color: #166534; font-weight: bold;">
                        📅 Día Reservado: {int(r['Dia_Reservado'])} / {int(r['Mes_Reservado'])} / {int(r['Anio_Reservado'])}
                    </div>
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        col_f1, col_f2, col_f3 = st.columns([1, 1.5, 1])
        with col_f2:
            if st.button("🏁 Finalizar y Cerrar Sesión", use_container_width=True):
                st.session_state.reserva_exitosa = None
                st.rerun()

    else:
        if df_escuelas.empty:
            st.warning("⚠️ No hay base de escuelas cargada en el sistema.")
        elif df_personas.empty:
            st.warning("⚠️ El padrón de autoridades no se encuentra cargado.")
        else:
            # CONTENEDOR 1: Validación del CUE de Escuelas y Duplicados
            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.subheader("📍 1. Identificación del Establecimiento Educativo")
            cue_ingresado = st.text_input("Ingrese el CUE de la institución:", key="cue_input_user", placeholder="Ej: 7000123").strip()
            
            nombre_escuela = ""
            modalidad = ""
            departamento = ""
            domicilio = ""
            escuela_valida = False
            
            if cue_ingresado:
                cue_limpio = normalizar_texto(cue_ingresado)
                
                cue_ya_reservado = False
                if 'CUE' in df_reservas_historico.columns:
                    cues_existentes = df_reservas_historico['CUE'].apply(normalizar_texto).values
                    if cue_limpio in cues_existentes:
                        cue_ya_reservado = True
                
                if cue_ya_reservado:
                    st.error("🚫 **Acceso Denegado:** Este CUE ya posee un turno asignado en el calendario institucional. No se permiten registros duplicados.")
                else:
                    coincidencia_esc = df_escuelas[df_escuelas['CUE'] == cue_limpio]
                    
                    if not coincidencia_esc.empty:
                        nombre_escuela = coincidencia_esc.iloc[0]['Nombre_Escuela']
                        modalidad = coincidencia_esc.iloc[0]['Modalidad_Oferta']
                        departamento = coincidencia_esc.iloc[0]['Departamento']
                        domicilio = coincidencia_esc.iloc[0]['Domicilio']
                        escuela_valida = True
                        
                        st.markdown(f"""
                            <div class="info-pill-container">
                                <div class="info-pill-title">🏫 Escuela Identificada</div>
                                <div class="info-pill-text">
                                    <strong>Nombre:</strong> {nombre_escuela}<br>
                                    <strong>Modalidad:</strong> {modalidad}<br>
                                    <strong>Departamento:</strong> {departamento}
                                </div>
                            </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.error("❌ El CUE ingresado no figura registrado en el sistema escolar.")
            st.markdown('</div>', unsafe_allow_html=True)

            # CONTENEDOR 2: Datos del Solicitante (Autoridad)
            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.subheader("👤 2. Datos del Solicitante (Autoridad)")
            dni_ingresado = st.text_input("Ingrese su DNI (sin puntos):", key="dni_input_user", placeholder="Ej: 22333444").strip()
            
            nombre_director = ""
            telefono_predicho = ""
            persona_valida = False
            
            if dni_ingresado and escuela_valida:
                dni_limpio = normalizar_texto(dni_ingresado)
                coincidencia_per = df_personas[df_personas['DNI'] == dni_limpio]
                
                if not coincidencia_per.empty:
                    nombre_director = coincidencia_per.iloc[0]['Apellido_Nombre']
                    telefono_predicho = coincidencia_per.iloc[0]['Telefono']
                    persona_valida = True
                    
                    st.markdown(f"""
                        <div class="info-pill-container" style="background-color: #f0fdf4; border: 1px solid #99f6e4;">
                            <div class="info-pill-title" style="color: #0f766e;">👤 Autoridad Verificada</div>
                            <div class="info-pill-text" style="color: #115e59;">
                                <strong>Apellido y Nombre:</strong> {nombre_director}
                            </div>
                        </div>
                        <br>
                    """, unsafe_allow_html=True)
                    telefono_final = st.text_input("Verifique o edite su Teléfono de Contacto:", value=telefono_predicho, placeholder="Ej: 2645551234")
                else:
                    st.warning("⚠️ El DNI ingresado no figura en el padrón precargado. Por favor, complete sus datos manualmente:")
                    
                    col_m1, col_m2 = st.columns(2)
                    with col_m1:
                        apellido_manual = st.text_input("Ingrese su/s Apellido/s:", placeholder="Ej: PÉREZ").strip().upper()
                    with col_m2:
                        nombre_manual = st.text_input("Ingrese su/s Nombre/s:", placeholder="Ej: Juan Carlos").strip()
                    
                    telefono_manual = st.text_input("Ingrese su Teléfono de Contacto:", placeholder="Ej: 2645551234").strip()
                    
                    if apellido_manual and nombre_manual:
                        nombre_director = f"{apellido_manual}, {nombre_manual}"
                        persona_valida = True
                    
                    telefono_final = telefono_manual
            else:
                telefono_final = ""
            st.markdown('</div>', unsafe_allow_html=True)

            # CONTENEDOR 3: REDISEÑADO - OBLIGATORIO Y VACÍO POR DEFECTO
            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.subheader("📊 3. Relevamiento de Cursos y Alumnos (Últimos 2 años)")
            
            # Caja de advertencia visual
            st.markdown('<div class="atencion-box">⚠️ PASO OBLIGATORIO: Debe seleccionar la estructura de su plan de estudios para poder declarar las divisiones y alumnos.</div>', unsafe_allow_html=True)
            
            # Selectbox con opción neutra/vacía inicial obligatoria
            estructura_opciones = [
                "Seleccione una opción...",
                "5° y 6° Año (Secundaria Orientada / Ciclo Superior Común)", 
                "6° y 7° Año (Escuelas Técnicas o de Modalidades Profesionales)"
            ]
            
            estructura_seleccionada = st.selectbox(
                "Estructura del plan de estudios de la institución:",
                options=estructura_opciones,
                index=0,
                key="estructura_plan_estudios"
            )
            
            datos_cursos = {}
            total_alumnos_declarados = 0
            estructura_valida_plan = False
            
            # Solo procesamos si el usuario elige una opción real distinta a la inicial
            if estructura_seleccionada != "Seleccione una opción...":
                estructura_valida_plan = True
                if "5° y 6°" in estructura_seleccionada:
                    ano_bajo, ano_alto = "5° Año", "6° Año"
                else:
                    ano_bajo, ano_alto = "6° Año", "7° Año"
                    
                col_a1, col_a2 = st.columns(2)
                
                with col_a1:
                    st.markdown(f"##### 📌 {ano_bajo}")
                    # Arranca en 0 para obligar a definir la cantidad real
                    cant_div_bajo = st.number_input(f"Cantidad de divisiones en {ano_bajo}:", min_value=0, max_value=15, value=0, step=1, key="div_bajo")
                    divs_bajo = []
                    if cant_div_bajo > 0:
                        for i in range(cant_div_bajo):
                            col_i1, col_i2 = st.columns([1, 2])
                            with col_i1:
                                seccion = st.text_input(f"Div. {i+1} ({ano_bajo}):", value=chr(65 + i) if i < 26 else str(i+1), key=f"sec_{ano_bajo}_{i}").strip()
                            with col_i2:
                                alumnos = st.number_input(f"Alumnos en {seccion}:", min_value=1, max_value=100, value=20, step=1, key=f"alu_{ano_bajo}_{i}")
                            divs_bajo.append({"division": seccion, "alumnos": alumnos})
                            total_alumnos_declarados += alumnos
                    datos_cursos[ano_bajo] = divs_bajo
                    
                with col_a2:
                    st.markdown(f"##### 📌 {ano_alto}")
                    # Arranca en 0 para obligar a definir la cantidad real
                    cant_div_alto = st.number_input(f"Cantidad de divisiones en {ano_alto}:", min_value=0, max_value=15, value=0, step=1, key="div_alto")
                    divs_alto = []
                    if cant_div_alto > 0:
                        for i in range(cant_div_alto):
                            col_j1, col_j2 = st.columns([1, 2])
                            with col_j1:
                                seccion = st.text_input(f"Div. {i+1} ({ano_alto}):", value=chr(65 + i) if i < 26 else str(i+1), key=f"sec_{ano_alto}_{i}").strip()
                            with col_j2:
                                alumnos = st.number_input(f"Alumnos en {seccion}:", min_value=1, max_value=100, value=20, step=1, key=f"alu_{ano_alto}_{i}")
                            divs_alto.append({"division": seccion, "alumnos": alumnos})
                            total_alumnos_declarados += alumnos
                    datos_cursos[ano_alto] = divs_alto
                    
                # Validación interna: Deben haber cargado al menos una división en total
                if cant_div_bajo == 0 and cant_div_alto == 0:
                    st.warning("Por favor, ingrese una cantidad de divisiones mayor a 0 para el año correspondiente.")
                    estructura_valida_plan = False
            else:
                st.info("💡 Por favor, despliegue el menú de arriba y elija la estructura de su plan de estudios para continuar.")
                
            st.markdown('</div>', unsafe_allow_html=True)

            # CONTENEDOR 4: Selección de Turno Disponible
            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.subheader("📅 4. Selección de Turno Disponible")
            
            fecha_inicio = datetime.date(anio_actual, 8, 1)
            fecha_limite = datetime.date(anio_actual, 11, 30)
            
            lista_fechas_libres = generar_fechas_disponibles(fecha_inicio, fecha_limite, feriados_arg, fechas_ocupadas)
            
            es_valida = False
            fecha_seleccionada = None
            
            # Agregamos la verificación de 'estructura_valida_plan' para habilitar el calendario
            if escuela_valida and persona_valida and estructura_valida_plan:
                if len(lista_fechas_libres) > 0:
                    opciones_combo = {formatear_fecha_espanol(f): f for f in lista_fechas_libres}
                    
                    seleccion_usuario = st.selectbox(
                        "Seleccione una de las fechas libres del sistema:",
                        options=list(opciones_combo.keys()),
                        index=0,
                        key="combo_fechas_libres"
                    )
                    
                    fecha_seleccionada = opciones_combo[seleccion_usuario]
                    es_valida = True
                    st.info(f"🎉 Elegiste el turno del día **{fecha_seleccionada.strftime('%d/%m/%Y')}**.")
                else:
                    st.error("🔴 Lo sentimos, ya no quedan turnos disponibles en el rango de Agosto a Noviembre.")
            else:
                st.info("Complete correctamente las secciones 1, 2 y 3 para poder calcular y seleccionar los turnos disponibles.")
                
            st.divider()
            
            # El formulario solo se habilita si pasó la nueva validación de estructura cargada
            formulario_listo = escuela_valida and persona_valida and estructura_valida_plan and es_valida and bool(telefono_final.strip())
            
            if st.button("Confirmar y Registrar Agenda", disabled=not formulario_listo):
                bajo_desc = ", ".join([f"Div {x['division']} ({x['alumnos']} al.)" for x in datos_cursos[ano_bajo]])
                alto_desc = ", ".join([f"Div {x['division']} ({x['alumnos']} al.)" for x in datos_cursos[ano_alto]])
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
                    "Fecha_Registro": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                
                guardar_reserva(datos_reserva)
                st.session_state.reserva_exitosa = datos_reserva
                st.rerun()
                
            if persona_valida and escuela_valida and es_valida and estructura_valida_plan and not telefono_final.strip():
                st.warning("Debe ingresar un número telefónico de contacto para habilitar la confirmación.")
            st.markdown('</div>', unsafe_allow_html=True)
