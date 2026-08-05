from fpdf import FPDF
import streamlit as st

# Configuración de página
st.set_page_config(
    page_title="Reserva Confirmada - Impulso Educativo",
    page_icon="🎉",
    layout="centered",
)


# Función para generar el PDF del comprobante en memoria
def crear_comprobante_pdf(
    establecimiento, cue, director, telefono, alumnos, fecha
):
    pdf = FPDF()
    pdf.add_page()

    # Configuración de márgenes y tipografía
    pdf.set_margins(15, 15, 15)

    # Encabezado Institucional
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, "IMPULSO EDUCATIVO", ln=True, align="C")
    pdf.set_font("Helvetica", "", 12)
    pdf.cell(
        0,
        8,
        "Comprobante de Reserva de Jornada Institucional",
        ln=True,
        align="C",
    )
    pdf.ln(5)

    # Línea divisoria
    pdf.set_line_width(0.5)
    pdf.line(15, pdf.get_y(), 195, pdf.get_y())
    pdf.ln(10)

    # Título principal
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 10, "¡RESERVA CONFIRMADA EXITOSAMENTE!", ln=True, align="C")
    pdf.ln(8)

    # Detalle de la Reserva
    pdf.set_font("Helvetica", "", 11)

    pdf.set_font("Helvetica", "B", 11)
    pdf.write(7, "Establecimiento: ")
    pdf.set_font("Helvetica", "", 11)
    pdf.write(7, f"{establecimiento}\n")

    pdf.set_font("Helvetica", "B", 11)
    pdf.write(7, "CUE: ")
    pdf.set_font("Helvetica", "", 11)
    pdf.write(7, f"{cue}\n")

    pdf.set_font("Helvetica", "B", 11)
    pdf.write(7, "Director Solicitante: ")
    pdf.set_font("Helvetica", "", 11)
    pdf.write(7, f"{director}\n")

    pdf.set_font("Helvetica", "B", 11)
    pdf.write(7, "Teléfono Contacto: ")
    pdf.set_font("Helvetica", "", 11)
    pdf.write(7, f"{telefono}\n")

    pdf.set_font("Helvetica", "B", 11)
    pdf.write(7, "Total Alumnos Registrados: ")
    pdf.set_font("Helvetica", "", 11)
    pdf.write(7, f"{alumnos} alumnos.\n\n")

    pdf.ln(5)

    # Recuadro con la fecha confirmada
    pdf.set_fill_color(240, 249, 240)
    pdf.set_draw_color(200, 230, 200)
    pdf.set_font("Helvetica", "B", 13)
    pdf.cell(
        0,
        14,
        f"Día Reservado: {fecha}",
        border=1,
        ln=True,
        align="C",
        fill=True,
    )

    pdf.ln(15)
    pdf.set_font("Helvetica", "I", 9)
    pdf.cell(
        0,
        6,
        "Este documento sirve como comprobante oficial de la solicitud registrada en el sistema.",
        ln=True,
        align="C",
    )

    return bytes(pdf.output())


# --- ESTRUCTURA DE LA PANTALLA EN STREAMLIT ---

st.caption(
    "Agende la jornada institucional de su establecimiento escolar sin"
    " superposiciones."
)

# Datos de la reserva (se pueden tomar de st.session_state si vienen de un formulario previo)
datos_reserva = {
    "establecimiento": (
        st.session_state.get(
            "escuela",
            'ESCUELA DE EDUCACIÓN SECUNDARIA "JUSTO JOSÉ DE URQUIZA"',
        )
    ),
    "cue": st.session_state.get("cue", "700094500"),
    "director": st.session_state.get("director", "GOMEZ GARCIA, MARCELO DANIEL"),
    "telefono": st.session_state.get("telefono", "2646687478"),
    "alumnos": st.session_state.get("alumnos", 13),
    "fecha": st.session_state.get("fecha", "24 / 11 / 2026"),
}

# Tarjeta de confirmación estilo visual similar a la imagen
with st.container(border=True):
    st.markdown(
        "<h1 style='text-align: center; font-size: 50px;'>🎉</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<h2 style='text-align: center; color: #1b1b1b;'>¡Reserva Confirmada"
        " Exitosamente!</h2>",
        unsafe_allow_html=True,
    )

    st.divider()

    st.markdown(
        f"**Establecimiento:** {datos_reserva['establecimiento']}"
    )
    st.markdown(f"**CUE:** {datos_reserva['cue']}")
    st.markdown(
        f"**Director Solicitante:** {datos_reserva['director']}"
    )
    st.markdown(
        f"**Teléfono Contacto:** {datos_reserva['telefono']}"
    )
    st.markdown(
        f"**Total Alumnos Registrados:** {datos_reserva['alumnos']} alumnos."
    )

    st.success(f"📅 **Día Reservado: {datos_reserva['fecha']}**")

st.write("")

# Generación del archivo PDF
pdf_bytes = crear_comprobante_pdf(
    establecimiento=datos_reserva["establecimiento"],
    cue=datos_reserva["cue"],
    director=datos_reserva["director"],
    telefono=datos_reserva["telefono"],
    alumnos=datos_reserva["alumnos"],
    fecha=datos_reserva["fecha"],
)

# Botones de Acción
col_pdf, col_logout = st.columns(2)

with col_pdf:
    st.download_button(
        label="📄 Descargar Comprobante PDF",
        data=pdf_bytes,
        file_name=f"comprobante_CUE_{datos_reserva['cue']}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

with col_logout:
    if st.button(
        "🏁 Finalizar y Cerrar Sesión", use_container_width=True, type="primary"
    ):
        st.session_state.clear()
        st.rerun()
