import datetime
from fpdf import FPDF
import pandas as pd
import streamlit as st

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Impulso Educativo - Capacitación en RCP",
    page_icon="📅",
    layout="centered",
)


# --- FUNCIÓN PARA GENERAR PDF DE COMPROBANTE ---
def crear_comprobante_pdf(
    establecimiento, cue, director, telefono, alumnos, fecha
):
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
        "Capacitación en RCP - Comprobante de Reserva",
        ln=True,
        align="C",
    )
    pdf.ln(5)

    # Línea separadora
    pdf.set_line_width(0.5)
    pdf.line(15, pdf.get_y(), 195, pdf.get_y())
    pdf.ln(8)

    # Título principal
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 10, "¡Reserva Confirmada Exitosamente!", ln=True, align="C")
    pdf.ln(6)

    # Datos
    pdf.set_font("Helvetica", "", 11)
    pdf.multi_cell(
        0,
        8,
        f"Establecimiento: {establecimiento}\n"
        f"CUE: {cue}\n"
        f"Director Solicitante: {director}\n"
        f"Teléfono Contacto: {telefono}\n"
        f"Total Alumnos Registrados: {alumnos} alumnos.",
    )
    pdf.ln(6)

    # Recuadro Fecha
    pdf.set_fill_color(240, 249, 240)
    pdf.set_draw_color(200, 230, 200)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(
        0,
        12,
        f"Día Reservado: {fecha}",
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
        "Este documento sirve como comprobante de la jornada reservada sin"
        " superposiciones.",
        ln=True,
        align="C",
    )

    return bytes(pdf.output())


# --- INICIALIZACIÓN DE ESTADO ---
if "paso" not in st.session_state:
    st.session_state.paso = "formulario"

if "datos_reserva" not in st.session_state:
    st.session_state.datos_reserva = {}


# --- PASO 1: FORMULARIO DE RESERVA ---
if st.session_state.paso == "formulario":
    st.title("Reserva de Jornada Institucional")
    st.caption(
        "Agende la jornada institucional de su establecimiento escolar sin"
        " superposiciones."
    )

    with st.form("form_reserva"):
        escuela = st.text_input(
            "Nombre del Establecimiento",
            value='ESCUELA DE EDUCACIÓN SECUNDARIA "JUSTO JOSÉ DE URQUIZA"',
        )
        cue = st.text_input("CUE", value="700094500")
        director = st.text_input(
            "Director Solicitante", value="GOMEZ GARCIA, MARCELO DANIEL"
        )
        telefono = st.text_input("Teléfono Contacto", value="2646687478")
        alumnos = st.number_input(
            "Total Alumnos Registrados", min_value=1, value=13, step=1
        )

        fecha_reserva = st.date_input(
            "Seleccionar Fecha de Jornada",
            value=datetime.date(2026, 11, 24),
            min_value=datetime.date.today(),
        )

        btn_confirmar = st.form_submit_button(
            "Confirmar Reserva", use_container_width=True, type="primary"
        )

        if btn_confirmar:
            # Guardar en session_state
            st.session_state.datos_reserva = {
                "establecimiento": escuela,
                "cue": cue,
                "director": director,
                "telefono": telefono,
                "alumnos": alumnos,
                "fecha": fecha_reserva.strftime("%d / %m / %Y"),
            }
            st.session_state.paso = "confirmacion"
            st.rerun()


# --- PASO 2: PANTALLA DE CONFIRMACIÓN Y DESCARGA ---
elif st.session_state.paso == "confirmacion":
    datos = st.session_state.datos_reserva

    st.caption(
        "Agende la jornada institucional de su establecimiento escolar sin"
        " superposiciones."
    )

    # Tarjeta de Confirmación
    with st.container(border=True):
        st.markdown(
            "<h1 style='text-align: center;'>🎉</h1>", unsafe_allow_html=True
        )
        st.markdown(
            "<h2 style='text-align: center;'>¡Reserva Confirmada"
            " Exitosamente!</h2>",
            unsafe_allow_html=True,
        )

        st.divider()

        st.write(f"**Establecimiento:** {datos.get('establecimiento', '')}")
        st.write(f"**CUE:** {datos.get('cue', '')}")
        st.write(f"**Director Solicitante:** {datos.get('director', '')}")
        st.write(f"**Teléfono Contacto:** {datos.get('telefono', '')}")
        st.write(
            f"**Total Alumnos Registrados:** {datos.get('alumnos', '')}"
            " alumnos."
        )

        st.success(f"📅 **Día Reservado: {datos.get('fecha', '')}**")

    st.write("")

    # Generación de Bytes en PDF
    pdf_bytes = crear_comprobante_pdf(
        establecimiento=datos.get("establecimiento", ""),
        cue=datos.get("cue", ""),
        director=datos.get("director", ""),
        telefono=datos.get("telefono", ""),
        alumnos=datos.get("alumnos", ""),
        fecha=datos.get("fecha", ""),
    )

    # Botón para descargar el PDF
    st.download_button(
        label="📥 Descargar Comprobante (PDF)",
        data=pdf_bytes,
        file_name=f"Comprobante_Reserva_{datos.get('cue', 'CUE')}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

    # Botón para reiniciar/cerrar
    if st.button(
        "🏁 Finalizar y Cerrar Sesión", use_container_width=True, type="primary"
    ):
        st.session_state.clear()
        st.rerun()
