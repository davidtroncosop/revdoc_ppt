import os
import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from openai import OpenAI # Configurar la clave API de OpenAI
from dotenv import load_dotenv

# Cargar variables de entorno desde .env
load_dotenv()
api_key = os.getenv('OPENAI_API_KEY')

client = OpenAI(api_key=api_key)

def generar_propuesta_resolucion(filas):
    resultados = []
    for _, fila in filas.iterrows():
        prompt = f"""
A continuación se presenta la información de un estudiante. Con base en esta información, por favor genera una "Propuesta Resolución" que indique si se aprueba o rechaza la solicitud de beca, y los detalles de la resolución. Usa las siguientes condiciones para tomar la decisión:

1. Si el PPE es menor a 0.5, rechaza la solicitud porque no cumple con el requisito mínimo.
2. Si la deuda vencida en el sistema es 0, rechaza la solicitud porque no hay deuda a cubrir.
3. Si los documentos han sido validados por una Trabajadora Social y el estudiante tiene una deuda vencida mayor que 0, aprueba la solicitud con los detalles correspondientes.
4. Si el estudiante ha recibido beneficios anteriormente, verifica si hay algún incumplimiento relacionado y decide en consecuencia.

Información del Estudiante:
PPE: {fila['PPE']}
Nombre completo: {fila['Nombre completo']}
RUT: {fila['Folder']}
Sede: {fila['Sede']}
Carrera: {fila['Carrera']}
Vigencia con cursos inscritos: {fila['Vigencia con cursos inscritos']}
Año y Semestre de ingreso: {fila['Año y Semestre de ingreso']}
Motivo solicitud: {fila['Motivo solicitud.']}
¿Ha recibido beneficios anteriormente? ¿Cuál?: {fila['¿Ha recibido beneficios anteriormente? ¿Cuál?']}
Última fecha en que se entregó el Beneficio: {fila['Última fecha en que se entregó el Beneficio']}
Deuda vencida en sistema: {fila['Deuda vencida en sistema']}
Motivo, Breve explicación de la situación del estudiante: {fila['Motivo, Breve explicación de la situación del estudiante, por la cual se solicita la Beca.']}
Análisis_concatenado: {fila['Análisis_concatenado']}

Por favor, genera una respuesta en formato de tabla con las siguientes columnas, incluyendo el encabezado, usando guiones (-) como delimitadores:
Propuesta Resolución-RESOLUCIÓN-MONTO DE LA BECA-MOTIVO DEL CASO-DOCUMENTOS
Ejemplo:
Propuesta Resolución-RESOLUCIÓN-MONTO DE LA BECA-MOTIVO DEL CASO-DOCUMENTOS
Aprobada-La solicitud de beca se aprueba...-Monto a determinar según normativa-El estudiante solicita una beca porque...-Carta de solicitud de beca; Cartola Hogar; Certificado de remuneraciones; FICHA SOCIOECONOMICA
Rechazada-La solicitud de beca se rechaza porque el estudiante no cumple con el requisito mínimo de PPE.-N/A-El estudiante no cumple con el requisito mínimo de PPE.-N/A

**Asegúrate de que cada valor esté correctamente delimitado por guiones y de que no haya espacios adicionales antes o después de los guiones. Incluye solo las 5 columnas especificadas y usa punto y coma para separar múltiples documentos en la columna DOCUMENTOS.**
"""

        completion = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Eres un asistente social."},
                {"role": "user", "content": prompt}
            ]
        )

        respuesta_fila = completion.choices[0].message.content.strip()
        print(f"Raw GPT response: {respuesta_fila}")

        # Eliminar el encabezado de la respuesta
        if 'Propuesta Resolución' in respuesta_fila:
            respuesta_fila = respuesta_fila.split('\n')[1]

        # Dividir la fila en columnas usando el guion "-"
        columnas = respuesta_fila.split('-')

        # Si hay más de 5 columnas, combinar las columnas extra en la última columna
        if len(columnas) > 5:
            columnas[4] = '-'.join(columnas[4:])  # Combina columnas extra en la columna de DOCUMENTOS
            columnas = columnas[:5]  # Mantén solo las primeras 5 columnas

        # Si hay menos columnas de las esperadas, agregar columnas vacías
        while len(columnas) < 5:
            columnas.append('')

        # Convertir las columnas a string
        columnas = [str(valor).strip() for valor in columnas]

        # Agregar las columnas a la lista de resultados
        resultados.append(columnas)

    return resultados  # <-- Devuelve la lista de listas


def add_textbox(slide, left, top, width, height, text, font_size=Pt(14), bold=False, font_color=RGBColor(0, 0, 0), alignment=PP_ALIGN.LEFT):
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    p = text_frame.add_paragraph()
    p.text = text
    p.font.size = font_size
    p.font.bold = bold
    p.font.color.rgb = font_color
    p.alignment = alignment

def create_header_background(slide, left, top, width, height):
    background = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    background.fill.solid()
    background.fill.fore_color.rgb = RGBColor(255, 192, 0)  # Light blue
    background.line.color.rgb = RGBColor(142, 180, 227)  # Sky blue border

def create_card(slide, left, top, width, height, title, subtitles, contents, image_path=None):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(255, 255, 255)  # White
    card.line.color.rgb = RGBColor(200, 200, 200)  # Light gray border

    # Add title
    title_box = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.25), width - Inches(0.5), Inches(0.5))
    title_box.text_frame.text = title
    title_box.text_frame.paragraphs[0].font.size = Pt(18)
    title_box.text_frame.paragraphs[0].font.bold = True

    # Add subtitles and contents
    content_top = top + Inches(0.75)
    for subtitle, content in zip(subtitles, contents):
        subtitle_box = slide.shapes.add_textbox(left + Inches(0.25), content_top, width - Inches(0.5), Inches(0.3))
        subtitle_box.text_frame.text = subtitle
        subtitle_box.text_frame.paragraphs[0].font.size = Pt(14)
        subtitle_box.text_frame.paragraphs[0].font.bold = True
        content_top += Inches(0.3)

        content_box = slide.shapes.add_textbox(left + Inches(0.25), content_top, width - Inches(0.5), Inches(1))
        content_box.text_frame.word_wrap = True
        if isinstance(content, str):
            content_box.text_frame.text = content
        else:
            content_box.text_frame.text = str(content)
        for paragraph in content_box.text_frame.paragraphs:
            paragraph.font.size = Pt(12)
        content_top += Inches(1)

def create_button(slide, left, top, width, height, text, color):
    button = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    button.fill.solid()
    button.fill.fore_color.rgb = color
    button.line.color.rgb = color
    button.text_frame.text = text
    button.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)  # White text

def create_slide_from_row(prs, row):
    # Create slide
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout

    # Add header background
    create_header_background(slide, Inches(0.25), Inches(0.2), Inches(12.9), Inches(1.6))

    # Add header information
    add_textbox(slide, Inches(0.5), Inches(0.2), Inches(4), Inches(0.5),
                f"{row.get('Nombre completo', 'N/A')}\n{row.get('Folder', 'N/A')}",
                font_size=Pt(14), bold=True)

    add_textbox(slide, Inches(5), Inches(0.2), Inches(4), Inches(0.5),
                f"{row.get('Carrera', 'N/A')}\n{row.get('Sede', 'N/A')}",
                font_size=Pt(14), bold=True)

    add_textbox(slide, Inches(9.5), Inches(0.2), Inches(3.5), Inches(0.5),
                f"MATRÍCULA 2024-1\nCON CURSOS: {row.get('Vigencia con cursos inscritos', 'N/A')}",
                font_size=Pt(14), bold=True, alignment=PP_ALIGN.RIGHT)

    # Add timeline
    add_textbox(slide, Inches(0.7), Inches(1), Inches(2), Inches(0.3),
                f"Ingresa: {row.get('Año y Semestre de ingreso', 'N/A')}",
                font_size=Pt(10), font_color=RGBColor(255, 255, 255))
    
    # Add timeline
    add_textbox(slide, Inches(5), Inches(1), Inches(2), Inches(0.3),
                f"MOTIVO DEL CASO:\n {row.get('MOTIVO DEL CASO', 'N/A')}",
                font_size=Pt(12), font_color=RGBColor(255, 255, 255))

    add_textbox(slide, Inches(10.5), Inches(1), Inches(2.5), Inches(0.3),
                f"Envía Solicitud: {row.get('Hora de inicio', 'N/A')}",
                font_size=Pt(10), font_color=RGBColor(255, 255, 255), alignment=PP_ALIGN.RIGHT)

    # Create three cards
    card_width = Inches(4)
    card_height = Inches(5.3)
    spacing = Inches(0.5)

# First card content
    subtitles_list1 = ["MOTIVO", "DOCUMENTOS"]  # Agrega "MOTIVO DEL CASO"
    contents_list1 = [
        row.get('Motivo solicitud.', 'N/A'),
        row.get('Análisis_concatenado', 'No hay información disponible')
    ]

    card_left = Inches(0.1) + 0 * (card_width + spacing)
    card_top = Inches(2)
    create_card(slide, card_left, card_top, card_width, card_height,
                "SOLICITA", subtitles_list1, contents_list1)

    # Extract FUAS information safely
    analisis = str(row.get('Análisis_concatenado', ''))
    fuas_info = 'No especificado'
    if 'Postulación FUAS:' in analisis:
        fuas_info = analisis.split('Postulación FUAS:')[1].split('.')[0].strip()

    subtitles_list2 = ["ANTECEDENTES ECONÓMICOS", "ANTECEDENTES ACADÉMICOS"]
    contents_list2 = [
        f"Beneficio: {row.get('¿Ha recibido beneficios anteriormente? ¿Cuál?', 'N/A')}\n"
        f"Deuda: ${'{:,}'.format(row.get('Deuda vencida en sistema', 0.0))}\n"
        f"Postulación a FUAS: {fuas_info}\n"
        f"Arancel: ${'{:,}'.format(row.get('Monto cuota de Arancel', 0.0))}\n"
        f"Matrícula: ${'{:,}'.format(row.get('Monto valor de matrícula', 0.0))}",
        f"Avance Curricular: {row.get('Avance curricular (%)', 'N/A')}\n"
        f"PPS: {row.get('PPS', 'N/A')}\n"
        f"RSH: {row.get('Registro Social de Hogares (RSH) o Nivel Socioeconómico (NSE)', 'N/A')}\n"
        f"Promedio Ponderado Evaluación: {'{:.2f}'.format(row.get('PPE', 0.0))}"
    ]
    
    # Extract resolution information safely
    resolucion = row.get('RESOLUCIÓN', 'No hay información disponible')
    monto_beca = f"${'{:,}'.format(row.get('Plan de Retención', 0.0))}"
     

    subtitles_list3 = ["RESOLUCIÓN", "MONTO DE LA BECA"]
    contents_list3 = [resolucion, monto_beca]

    button_left = card_left + Inches(0.25)
    button_top = card_top + card_height - Inches(0.7)
    button_width = card_width - Inches(0.5)
    button_height = Inches(0.5)
    create_button(slide, button_left, button_top, button_width, button_height,
                  "Solicitud", RGBColor(13, 34, 60))  # Indigo color

    card_left = Inches(0.1) + 1 * (card_width + spacing)
    create_card(slide, card_left, card_top, card_width, card_height,
                "ANTECEDENTES", subtitles_list2, contents_list2)

    button_left = card_left + Inches(0.25)
    create_button(slide, button_left, button_top, button_width, button_height,
                  "Revisión", RGBColor(13, 34, 60))  # Teal color

    card_left = Inches(0.1) + 2 * (card_width + spacing)
    create_card(slide, card_left, card_top, card_width, card_height,
                "RESOLUCIÓN", subtitles_list3, contents_list3)

    button_left = card_left + Inches(0.25)
    create_button(slide, button_left, button_top, button_width, button_height,
                  "Resolución", RGBColor(13, 34, 60))   # Blue color

st.title('Generador de Propuestas de Resolución y Presentaciones')

uploaded_file = st.file_uploader("Carga un archivo Excel", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    st.write('Datos del archivo cargado:')
    st.dataframe(df)

    propuestas_resolucion = generar_propuesta_resolucion(df)

    # Agregar los resultados al dataframe original
    df['Propuesta Resolución'], df['RESOLUCIÓN'], df['MONTO DE LA BECA'], df['MOTIVO DEL CASO'], df['DOCUMENTOS'] = zip(*propuestas_resolucion)

    # Crear un nuevo archivo Excel con los resultados
    excel_output = 'propuestas_resolucion.xlsx'
    df.to_excel(excel_output, index=False)

    # Crea la presentación de PowerPoint
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    # Crea una diapositiva para cada fila
    for _, row in df.iterrows():
        create_slide_from_row(prs, row)

    ppt_output = 'propuestas_resolucion.pptx'
    prs.save(ppt_output)

    st.success('¡Propuestas de resolución generadas y guardadas!')

    with open(excel_output, 'rb') as excel_file:
        st.download_button('Descargar archivo Excel', data=excel_file, file_name=excel_output)

    with open(ppt_output, 'rb') as ppt_file:
        st.download_button('Descargar presentación PowerPoint', data=ppt_file, file_name=ppt_output)

