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

# --- DETECCIÓN DE BASE DE DATOS ACTIVA (GSheets o Local) ---
def usando_google_sheets():
    """Detecta si las credenciales de Google Sheets están configuradas en los Secrets."""
    if not gsheets_librerias_listas:
        return False
    try:
        return "gcp_service_account" in st.secrets and "spreadsheet_url" in st.secrets
    except Exception:
        return False

def conectar_google_sheets():
    """Autentica con la API de Google Sheets y devuelve la primera hoja de la planilla."""
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
    """Limpia cadenas de texto, CUEs o DNI removiendo formatos residuales de Excel."""
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
    """Carga y procesa la base de escuelas con dirección, modalidad y departamento."""
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
    """Carga la base de directivos/personas unificando Apellido, Nombre y Teléfono."""
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

def obtener_fechas_ocupadas():
    """Retorna un conjunto de objetos datetime.date ocupados. Inicializa la planilla de Google si está vacía."""
    if usando_google_sheets():
        try:
            hoja = conectar_google_sheets()
            valores = hoja.get_all_values()
            
            # --- AUTO-CURACIÓN: Si la planilla está 100% vacía, le escribimos las columnas ---
            if not valores or len(valores) == 0:
                hoja.append_row(COLUMNAS_SISTEMA)
                return set()
                
            # Si solo tiene cabeceras
            if len(valores) == 1:
                return set()
                
            # Procesamos con Pandas de forma segura a partir de los valores de las celdas
            df = pd.DataFrame(valores[1:], columns=valores[0])
            if 'Dia_Reservado' in df.columns and 'Mes_Reservado' in df.columns and 'Anio_Reservado' in df.columns:
                fechas = []
                for _, row in df.iterrows():
                    try:
                        d = int(float(row['Dia_Reservado']))
                        m = int(float(row['Mes_Reservado']))
                        a = int(float(row['Anio_Reservado']))
                        f = datetime.date(a, m, d)
                        fechas.append(f)
                    except Exception:
                        continue
                return set(fechas)
            return set()
        except gspread.exceptions.SpreadsheetNotFound:
            st.sidebar.error("❌ Error de Google: No se encuentra la planilla. Verificá que la `spreadsheet_url` en Secrets sea idéntica a tu navegador.")
            return set()
        except gspread.exceptions.APIError as e:
            error_data = e.response.text if hasattr(e, 'response') else str(e)
            if "permission" in error_data.lower() or "auth" in error_data.lower():
                st.sidebar.error("❌ Error de Google: Tu cuenta de servicio no tiene permiso para abrir esta planilla. Asegurate de haber compartido la planilla con el mail del operador como Editor.")
            else:
                st.sidebar.error(f"❌ Error de la API de Google: {e}")
            return set()
        except Exception as e:
            # --- MEJORA CRÍTICA DE DIAGNÓSTICO ---
            error_msg = str(e)
            tipo_error = type(e).__name__
            # Si str(e) está vacío, usamos repr(e) para que nunca se muestre en blanco
            detalles_error = str(e) if str(e).strip() else repr(e)
            
            if "private_key" in error_msg.lower() or "key" in error_msg.lower() or "pem" in error_msg.lower():
                st.sidebar.error("🔑 **Error de Clave Privada:** Tu `private_key` en los Secrets está mal copiada, le faltan comillas o tiene caracteres inválidos.")
            elif "email" in error_msg.lower():
                st.sidebar.error("📧 **Error de Correo:** Falta el `client_email` o está mal escrito en los Secrets.")
            elif "url" in error_msg.lower() or "http" in error_msg.lower():
                st.sidebar.error("🌐 **Error de URL:** La URL de la planilla en los Secrets está mal estructurada.")
            else:
                # Mostramos la clase de error de Python para saber exactamente qué está fallando
                st.sidebar.warning(f"⚠️ Error de Diagnóstico al leer Google Sheets: {tipo_error} ({detalles_error})")
            return set()
    else:
        if os.path.exists(EXCEL_RESERVAS_LOCAL):
            try:
                df = pd.read_excel(EXCEL_RESERVAS_LOCAL)
                if 'Dia_Reservado' in df.columns and 'Mes_Reservado' in df.columns and 'Anio_Reservado' in df.columns:
                    fechas = []
                    for _, row in df.iterrows():
                        try:
                            f = datetime.date(int(row['Anio_Reservado']), int(row['Mes_Reservado']), int(row['Dia_Reservado']))
                            fechas.append(f)
                        except Exception:
                            continue
                    return set(fechas)
            except Exception:
                return set()
        return set()

def guardar_reserva(datos):
    """Guarda la reserva en Google Sheets (producción) o en Excel Local (desarrollo/respaldo)."""
    if usando_google_sheets():
        try:
            hoja = conectar_google_sheets()
            valores = hoja.get_all_values()
            
            # Si está totalmente vacía la planilla, escribimos los encabezados
            if not valores or len(valores) == 0:
                hoja.append_row(COLUMNAS_SISTEMA)
                
            datos_lista = []
            for col in COLUMNAS_SISTEMA:
                val = datos.get(col, "")
                if isinstance(val, (datetime.date, datetime.datetime)):
                    datos_lista.append(str(val))
                else:
                    datos_lista.append(val)
                    
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
fechas_ocupadas = obtener_fechas_ocupadas()
df_escuelas = cargar_base_escuelas()
df_personas = cargar_base_personas()

# Panel de administración oculto discretamente en el Sidebar colapsado
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
        st.caption("Suba un archivo Excel (.xlsx) que contenga CUE, Nombre_Escuela, Modalidad_Oferta, Departamento y Domicilio.")
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
        st.caption("Suba un archivo Excel (.xlsx) con columnas DNI, Apellido, Nombre y Telefono.")
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
    
    base_de_datos_lista = False
    df_reservas = pd.DataFrame()
    
    if usando_google_sheets():
        try:
            hoja = conectar_google_sheets()
            valores = hoja.get_all_values()
            
            # Si está vacía o solo con cabeceras, no hay reservas
            if valores and len(valores) > 1:
                df_reservas = pd.DataFrame(valores[1:], columns=valores[0])
                base_de_datos_lista = True
                st.info("🟢 Los datos mostrados corresponden a la planilla de **Google Sheets** en tiempo real.")
            else:
                st.info("No se registran reservas agendadas en Google Sheets todavía.")
                # Si está 100% vacía, la inicializamos
                if not valores or len(valores) == 0:
                    hoja.append_row(COLUMNAS_SISTEMA)
        except gspread.exceptions.SpreadsheetNotFound:
            st.error("❌ Error de Sincronización: No se pudo localizar la planilla de Google Sheets. Revisá que hayas pegado la URL exacta.")
        except gspread.exceptions.APIError as e:
            error_data = e.response.text if hasattr(e, 'response') else str(e)
            if "permission" in error_data.lower() or "auth" in error_data.lower():
                st.error("❌ Error de Sincronización: Falta de permisos. Recordá compartir tu planilla de Drive con el correo del operador como Editor.")
            else:
                st.error(f"❌ Error de API de Google Sheets: {e}")
        except Exception as e:
            st.error(f"❌ Error al sincronizar con Google Sheets: {e}")
    else:
        if os.path.exists(EXCEL_RESERVAS_LOCAL):
            try:
                df_reservas = pd.read_excel(EXCEL_RESERVAS_LOCAL)
                base_de_datos_lista = True
                st.warning("⚠️ Los datos mostrados se encuentran guardados de forma **Local** en el servidor.")
            except Exception as e:
                st.error(f"Error al leer archivo local: {e}")
        else:
            st.info("No se registran reservas agendadas de manera local todavía.")
            
    if base_de_datos_lista and not df_reservas.empty:
        st.dataframe(df_reservas, use_container_width=True)
        
        buffer_excel = io.BytesIO()
        df_reservas.to_excel(buffer_excel, index=False)
        st.download_button(
            label="📥 Descargar Excel de Reservas Sincronizado",
            data=buffer_excel.getvalue(),
            file_name=f"registro_reservas_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="custom-card" style="border: 1px solid #fecaca; background-color: #fef2f2;">', unsafe_allow_html=True)
    st.subheader("⚠️ Zona de Peligro: Reiniciar Calendario")
    st.write("Si desea borrar permanentemente todas las reservas agendadas por los directores y comenzar desde cero, active la confirmación abajo.")
    
    confirmar_reinicio = st.checkbox("Confirmo que deseo vaciar todo el registro de reservas", key="check_reinicio")
    if st.button("🗑️ Eliminar todas las reservas del Excel", disabled=not confirmar_reinicio):
        if usando_google_sheets():
            try:
                hoja = conectar_google_sheets()
                hoja.clear()
                # Siempre mantenemos las cabeceras requeridas
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
            else:
                st.info("No hay un archivo de reservas registrado para eliminar.")
    st.markdown('</div>', unsafe_allow_html=True)

# ================= VISTA DE DIRECTORES (PÚBLICA) =================
else:
    st.markdown('<h1 style="text-align: center; color: #0284c7 !important; margin-bottom: 5px;">📅 Sistema de Reserva de Turnos</h1>', unsafe_allow_html=True)
    st.markdown('<p style="text-align: center; color: #64748b; font-size: 1.1rem; margin-bottom: 25px;">Agende la jornada institucional de su establecimiento escolar sin superposiciones.</p>', unsafe_allow_html=True)

    # Validamos si el registro está desactivado
    if not registro_activo:
        st.error("⚠️ **Sistema Desactivado:** El período de agendamiento se encuentra inhabilitado en este momento. Por favor comuníquese con el administrador del sistema para más información.")
        
    elif st.session_state.reserva_exitosa is not None:
        r = st.session_state.reserva_exitosa
        st.markdown(f"""
            <div style="background-color: #ffffff; border: 2px solid #22c55e; border-radius: 12px; padding: 30px; box-shadow: 0 10px 15px -3px rgba(34, 197, 94, 0.1); max-width: 650px; margin: 0 auto 30px auto;">
                <div style="text-align: center; margin-bottom: 20px;">
                    <span style="font-size: 3.5rem;">🎉</span>
                    <h2 style="color: #15803d !important; margin-top: 10px;">¡Reserva Confirmada Exitosamente!</h2>
                    <p style="color: #475569; font-size: 1rem;">Su establecimiento educativo ha reservado un turno exclusivo en la agenda.</p>
                </div>
                <hr style="border: 0; border-top: 1px solid #e2e8f0; margin-bottom: 20px;">
                <div style="font-size: 1rem; color: #1e293b; line-height: 1.6;">
                    <strong style="color: #0f172a;">Establecimiento:</strong> {r['Escuela']}<br>
                    <strong style="color: #0f172a;">CUE:</strong> {r['CUE']}<br>
                    <strong style="color: #0f172a;">Director Solicitante:</strong> {r['Director']}<br>
                    <strong style="color: #0f172a;">Teléfono Contacto:</strong> {r['Telefono_Contacto']}<br>
                    <strong style="color: #0f172a;">Estructura Declarada:</strong> {r['Estructura_Declarada']}<br>
                    <strong style="color: #0f172a;">Total Alumnos Registrados:</strong> {r['Total_Alumnos']} alumnos.<br>
                    <hr style="border: 0; border-top: 1px dashed #cbd5e1; margin: 15px 0;">
                    <div style="background-color: #f0fdf4; border: 1px solid #bbf7d0; padding: 12px; border-radius: 8px; text-align: center; font-size: 1.1rem; color: #166534; font-weight: bold;">
                        📅 Día Reservado: {r['Dia_Reservado']:02d} / {r['Mes_Reservado']:02d} / {r['Anio_Reservado']}
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
            st.warning("⚠️ No hay base de escuelas cargada en el sistema. Solicite asistencia al administrador.")
        elif df_personas.empty:
            st.warning("⚠️ El padrón de autoridades no se encuentra cargado. Solicite asistencia al administrador.")
        else:
            # CONTENEDOR 1: Validación del CUE de Escuelas
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
                                <strong>Departamento:</strong> {departamento}<br>
                                <strong>Domicilio:</strong> {domicilio}
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
                else:
                    st.error("❌ El CUE ingresado no figura registrado en el sistema escolar.")
                    
            st.markdown('</div>', unsafe_allow_html=True)

            # CONTENEDOR 2: Validación del DNI de Autoridades
            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.subheader("👤 2. Datos del Solicitante (Autoridad)")
            
            dni_ingresado = st.text_input("Ingrese su DNI (sin puntos):", key="dni_input_user", placeholder="Ej: 22333444").strip()
            
            nombre_director = ""
            telefono_predicho = ""
            persona_valida = False
            
            if dni_ingresado:
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
                    st.error("❌ El DNI ingresado no corresponde a un directivo habilitado en el padrón.")
                    telefono_final = ""
            else:
                telefono_final = ""
                
            st.markdown('</div>', unsafe_allow_html=True)

            # CONTENEDOR 3: Estructuras y Matrículas declaradas
            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.subheader("📊 3. Relevamiento de Cursos y Alumnos (Últimos 2 años)")
            st.write("Declare las divisiones y matrículas correspondientes únicamente a los dos últimos años de cursado de su plan de estudios.")
            
            estructura = st.radio(
                "Estructura del plan de estudios de la institución:",
                ["5° y 6° Año (Secundaria Orientada / Ciclo Superior Común)", 
                 "6° y 7° Año (Escuelas Técnicas o de Modalidades Profesionales)"],
                horizontal=True
            )
            
            if "5° y 6°" in estructura:
                ano_bajo, ano_alto = "5° Año", "6° Año"
            else:
                ano_bajo, ano_alto = "6° Año", "7° Año"
                
            datos_cursos = {}
            total_alumnos_declarados = 0
            
            col_a1, col_a2 = st.columns(2)
            
            with col_a1:
                st.markdown(f"##### 📌 {ano_bajo}")
                cant_div_bajo = st.number_input(f"Cantidad de divisiones en {ano_bajo}:", min_value=1, max_value=15, value=1, step=1, key="div_bajo")
                
                divs_bajo = []
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
                cant_div_alto = st.number_input(f"Cantidad de divisiones en {ano_alto}:", min_value=1, max_value=15, value=1, step=1, key="div_alto")
                
                divs_alto = []
                for i in range(cant_div_alto):
                    col_j1, col_j2 = st.columns([1, 2])
                    with col_j1:
                        seccion = st.text_input(f"Div. {i+1} ({ano_alto}):", value=chr(65 + i) if i < 26 else str(i+1), key=f"sec_{ano_alto}_{i}").strip()
                    with col_j2:
                        alumnos = st.number_input(f"Alumnos en {seccion}:", min_value=1, max_value=100, value=20, step=1, key=f"alu_{ano_alto}_{i}")
                    divs_alto.append({"division": seccion, "alumnos": alumnos})
                    total_alumnos_declarados += alumnos
                datos_cursos[ano_alto] = divs_alto
                
            st.markdown('</div>', unsafe_allow_html=True)

            # CONTENEDOR 4: Calendario de Reserva Excluyente
            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.subheader("📅 4. Selección de Turno Excluyente")
            st.write("Las fechas solo pueden reservarse de **Agosto a Noviembre** de lunes a viernes, exceptuando feriados.")
            
            fecha_minima = datetime.date(anio_actual, 8, 1)
            fecha_maxima = datetime.date(anio_actual, 11, 30)
            
            fecha_seleccionada = st.date_input(
                "Seleccione el día que reservará para la escuela:", 
                value=fecha_minima,
                min_value=fecha_minima,
                max_value=fecha_maxima,
                key="reserva_date"
            )
            
            es_valida = True
            motivo_invalido = ""
            
            if fecha_seleccionada.weekday() in [5, 6]:
                es_valida = False
                motivo_invalido = "La fecha seleccionada corresponde a un fin de semana (sábado/domingo)."
            elif fecha_seleccionada in feriados_arg:
                es_valida = False
                motivo_invalido = f"Feriado Nacional: {feriados_arg.get(fecha_seleccionada)}"
            elif fecha_seleccionada in fechas_ocupadas:
                es_valida = False
                motivo_invalido = "Esta fecha ya fue agendada por otra institución."
                
            if es_valida:
                st.info(f"🟢 La fecha **{fecha_seleccionada.strftime('%d/%m/%Y')}** está disponible para su asignación exclusiva.")
            else:
                st.error(f"🔴 No disponible: {motivo_invalido}")
                
            st.divider()
            
            formulario_listo = escuela_valida and persona_valida and es_valida and bool(telefono_final.strip())
            
            if st.button("Confirmar y Registrar Agenda", disabled=not formulario_listo):
                bajo_desc = ", ".join([f"Div {x['division']} ({x['alumnos']} al.)" for x in datos_cursos[ano_bajo]])
                alto_desc = ", ".join([f"Div {x['division']} ({x['alumnos']} al.)" for x in datos_cursos[ano_alto]])
                resumen_matricula = f"{ano_bajo}: [{bajo_desc}] | {ano_alto}: [{alto_desc}]"
                
                datos_reserva = {
                    "CUE": cue_ingresado,
                    "Escuela": nombre_escuela,
                    "Modalidad_Oferta": modalidad,
                    "Departamento": departamento,
                    "Domicilio": domicilio,
                    "DNI_Director": dni_ingresado,
                    "Director": nombre_director,
                    "Telefono_Contacto": telefono_final.strip(),
                    "Estructura_Declarada": f"{ano_bajo} y {ano_alto}",
                    "Detalle_Divisiones_Alumnos": resumen_matricula,
                    "Total_Alumnos": total_alumnos_declarados,
                    "Dia_Reservado": int(fecha_seleccionada.day),
                    "Mes_Reservado": int(fecha_seleccionada.month),
                    "Anio_Reservado": int(fecha_seleccionada.year),
                    "Fecha_Registro": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                guardar_reserva(datos_reserva)
                
                st.session_state.reserva_exitosa = datos_reserva
                st.cache_data.clear()
                st.rerun()
                
            if persona_valida and escuela_valida and es_valida and not telefono_final.strip():
                st.warning("Debe ingresar un número telefónico de contacto para habilitar la confirmación.")
            st.markdown('</div>', unsafe_allow_html=True)

