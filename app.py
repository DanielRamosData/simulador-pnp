import io
import os
import pandas as pd
import streamlit as st
from parser import cargar_banco_preguntas, estructurar_15_modulos

# ReportLab para la generación del PDF
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Simulador PNP", layout="wide")

st.title("🛡️ Simulador de Examen PNP - 15 Módulos")

# RUTA DEL ARCHIVO DEFINIDO
NOMBRE_ARCHIVO_DOCX = "banco_preguntas.docx"

# ---------------------------------------------------------
# FUNCIÓN AUXILIAR DE COMPARACIÓN DE RESPUESTAS
# ---------------------------------------------------------
def es_igual(val1, val2):
    if not val1 or not val2:
        return False
    # Normaliza eliminando espacios extra, puntos finales y convirtiendo a mayúsculas
    n1 = str(val1).strip().rstrip('.').strip().upper()
    n2 = str(val2).strip().rstrip('.').strip().upper()
    return n1 == n2


# ---------------------------------------------------------
# FUNCIÓN PARA GENERAR EL REPORTE PDF DE ERRORES
# ---------------------------------------------------------
def generar_pdf_reporte(modulo_nombre, nota, correctas, incorrectas, sin_responder, preguntas, respuestas_usuario):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36
    )
    
    styles = getSampleStyleSheet()
    titulo_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=16, leading=20, textColor=colors.HexColor("#1A365D"))
    sub_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontSize=10, leading=14)
    pregunta_style = ParagraphStyle('PreguntaStyle', parent=styles['Normal'], fontSize=9, leading=12, textColor=colors.HexColor("#2D3748"))
    bad_ans_style = ParagraphStyle('BadAns', parent=styles['Normal'], fontSize=9, leading=12, textColor=colors.HexColor("#C53030"))
    good_ans_style = ParagraphStyle('GoodAns', parent=styles['Normal'], fontSize=9, leading=12, textColor=colors.HexColor("#2F855A"))
    
    story = []
    story.append(Paragraph(f"<b>REPORTE DE RETROALIMENTACIÓN - {modulo_nombre.upper()}</b>", titulo_style))
    story.append(Spacer(1, 8))
    
    resumen_text = f"<b>Nota Final:</b> {nota:.2f} / 20 &nbsp;&nbsp;|&nbsp;&nbsp; <b>Correctas:</b> {correctas} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Incorrectas:</b> {incorrectas} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Sin Responder:</b> {sin_responder}"
    story.append(Paragraph(resumen_text, sub_style))
    story.append(Spacer(1, 15))
    
    preguntas_revision = []
    for q in preguntas:
        num = q["num"]
        user_ans = respuestas_usuario.get(num)
        es_correcta = (user_ans is not None and es_igual(user_ans, q["correcta"]))
        if not es_correcta:
            preguntas_revision.append((q, user_ans))
            
    if not preguntas_revision:
        story.append(Paragraph("<b>¡Felicidades! Obtuviste un puntaje perfecto. No hay preguntas incorrectas para revisar.</b>", good_ans_style))
    else:
        story.append(Paragraph("<b>DETALLE DE PREGUNTAS A REVISAR (INCORRECTAS / OMITIDAS):</b>", styles['Heading2']))
        story.append(Spacer(1, 10))
        
        tabla_datos = [["#", "Enunciado de la Pregunta", "Tu Respuesta", "Rpta. Correcta"]]
        
        for q, user_ans in preguntas_revision:
            txt_num = str(q["num"])
            txt_preg = Paragraph(q["pregunta"], pregunta_style)
            txt_user = Paragraph(str(user_ans) if user_ans else "<i>Sin responder</i>", bad_ans_style)
            txt_correct = Paragraph(str(q["correcta"]), good_ans_style)
            
            tabla_datos.append([txt_num, txt_preg, txt_user, txt_correct])
            
        t = Table(tabla_datos, colWidths=[25, 330, 90, 90])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#EDF2F7")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#1A202C")),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        ]))
        story.append(t)
        
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------
# 1. CARGA AUTOMÁTICA Y CACHEADA DEL ARCHIVO LOCAL
# ---------------------------------------------------------
@st.cache_data(show_spinner="Cargando banco de preguntas definido...")
def obtener_modulos_definidos(ruta_docx):
    if not os.path.exists(ruta_docx):
        return None
    banco_total = cargar_banco_preguntas(ruta_docx)
    return estructurar_15_modulos(banco_total, preguntas_por_modulo=100)

modulos_15 = obtener_modulos_definidos(NOMBRE_ARCHIVO_DOCX)

if modulos_15 is None:
    st.error(f"❌ No se encontró el archivo **{NOMBRE_ARCHIVO_DOCX}** en la carpeta del proyecto. Por favor, renombra tu archivo Word como `{NOMBRE_ARCHIVO_DOCX}` y colócalo junto a `app.py`.")
    st.stop()

# ---------------------------------------------------------
# 2. SISTEMA DE NAVEGACIÓN Y EVALUACIÓN
# ---------------------------------------------------------
st.sidebar.subheader("📌 Módulos del Examen")
nombres_modulos = [f"Módulo {i}" for i in range(1, 16)]
modulo_seleccionado = st.sidebar.selectbox("Selecciona un Módulo:", nombres_modulos)

# Inicializar/Resetear variables de sesión al cambiar de módulo
if "modulo_activo" not in st.session_state or st.session_state.modulo_activo != modulo_seleccionado:
    st.session_state.modulo_activo = modulo_seleccionado
    st.session_state.idx_pregunta = 0
    st.session_state.respuestas_modulo = {}
    st.session_state.examen_finalizado = False

preguntas = modulos_15.get(modulo_seleccionado, [])
total_preguntas = len(preguntas)

st.header(f"📝 {modulo_seleccionado} ({total_preguntas} Preguntas)")

if total_preguntas == 0:
    st.warning("No hay preguntas disponibles para este módulo.")

elif st.session_state.examen_finalizado:
    # Cálculo de Resultados
    correctas = sum(1 for q in preguntas if es_igual(st.session_state.respuestas_modulo.get(q["num"]), q["correcta"]))
    sin_responder = sum(1 for q in preguntas if q["num"] not in st.session_state.respuestas_modulo)
    incorrectas = total_preguntas - correctas - sin_responder
    nota = (correctas / total_preguntas) * 20 if total_preguntas > 0 else 0
    
    st.success("🎉 ¡Examen Finalizado!")
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Correctas", correctas)
    col2.metric("Incorrectas", incorrectas)
    col3.metric("Sin responder", sin_responder)
    
    st.subheader(f"Nota Final: **{nota:.2f} / 20**")
    st.markdown("---")
    
    st.subheader("📋 DETALLE DE PREGUNTAS A REVISAR")
    filas_tabla = []
    for q in preguntas:
        user_ans = st.session_state.respuestas_modulo.get(q["num"], "Sin responder")
        rpta_correcta = q["correcta"]
        # Filtrar solo las incorrectas o sin responder usando la comparación flexible
        if not es_igual(user_ans, rpta_correcta):
            filas_tabla.append({
                "#": q["num"],
                "Enunciado de la Pregunta": q["pregunta"],
                "Tu Respuesta": user_ans,
                "Rpta. Correcta": rpta_correcta
            })
    
    if filas_tabla:
        st.dataframe(pd.DataFrame(filas_tabla), use_container_width=True)
    else:
        st.info("¡Respondiste todas las preguntas correctamente!")
    st.markdown("---")

    # PDF de retroalimentación
    pdf_bytes = generar_pdf_reporte(
        modulo_seleccionado,
        nota,
        correctas,
        incorrectas,
        sin_responder,
        preguntas,
        st.session_state.respuestas_modulo
    )
    
    col_pdf, col_reiniciar = st.columns(2)
    with col_pdf:
        st.download_button(
            label="📄 Descargar Reporte de Errores (PDF)",
            data=pdf_bytes,
            file_name=f"Reporte_{modulo_seleccionado.replace(' ', '_')}.pdf",
            mime="application/pdf",
            type="primary"
        )
        
    with col_reiniciar:
        if st.button("🔄 Volver a intentar este Módulo"):
            st.session_state.respuestas_modulo = {}
            st.session_state.idx_pregunta = 0
            st.session_state.examen_finalizado = False
            st.rerun()

else:
    idx = st.session_state.idx_pregunta
    q_actual = preguntas[idx]
    num_q = q_actual["num"]

    st.progress((idx + 1) / total_preguntas)
    st.caption(f"Pregunta {idx + 1} de {total_preguntas} | Respuestas guardadas: {len(st.session_state.respuestas_modulo)} / {total_preguntas}")

    card_slot = st.empty()
    with card_slot.container():
        st.markdown(f"### Pregunta {num_q}")
        st.write(q_actual['pregunta'])
        if q_actual.get("modulo"):
            st.caption(f"Tema: {q_actual['modulo']} | Código: {q_actual['codigo_id']}")

        opciones_raw = q_actual.get("opciones", [])
        opciones_validas = list(dict.fromkeys([str(opc).strip() for opc in opciones_raw if str(opc).strip()]))
        if not opciones_validas:
            opciones_validas = ["A", "B", "C", "D"]

        res_previa = st.session_state.respuestas_modulo.get(num_q)
        idx_opcion = opciones_validas.index(res_previa) if res_previa in opciones_validas else None

        eleccion = st.radio(
            "Selecciona una opción:",
            options=opciones_validas,
            index=idx_opcion,
            key=f"q_radio_{modulo_seleccionado}_{idx}"
        )

        if eleccion is not None:
            st.session_state.respuestas_modulo[num_q] = eleccion

    st.markdown("---")

    col_prev, col_center, col_next = st.columns([1, 2, 1])

    with col_prev:
        if st.button("⬅️ Anterior", disabled=(idx == 0)):
            st.session_state.idx_pregunta -= 1
            st.rerun()

    with col_center:
        opciones_salto = list(range(1, total_preguntas + 1))
        pregunta_ir = st.selectbox("Saltar a la pregunta:", opciones_salto, index=idx)
        if pregunta_ir - 1 != idx:
            st.session_state.idx_pregunta = pregunta_ir - 1
            st.rerun()

    with col_next:
        if idx < total_preguntas - 1:
            if st.button("Siguiente ➡️"):
                st.session_state.idx_pregunta += 1
                st.rerun()
        else:
            if st.button("🏁 Finalizar Módulo", type="primary"):
                st.session_state.examen_finalizado = True
                st.rerun()
