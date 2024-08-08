import os
import tempfile
import streamlit as st
from dotenv import load_dotenv
import zipfile
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image as PILImage
import io
import requests
import base64
from openai import OpenAI # Configurar la clave API de OpenAI


# Cargar variables de entorno desde .env
load_dotenv()
api_key = os.getenv('OPENAI_API_KEY')

if not api_key:
    st.error("API Key no encontrada. Asegúrate de que el archivo .env esté correctamente configurado.")
    st.stop()

def encode_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')

def analyze_image(base64_image, api_key):
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }

    payload = {
        "model": "gpt-4o-mini",
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": "Eres un asistente social virtual especializado en la revisión de documentos escaneados. Tu tarea es analizar imágenes de documentos como liquidaciones, finiquitos, licencias médicas, documentos de identidad y documentos firmados con sellos. Debes identificar el tipo de documento y extraer la siguiente información: nombres completos de las personas involucradas, fechas importantes (fechas de emisión, vencimiento, consultas médicas, etc.), diagnósticos médicos (en el caso de licencias médicas o certificados médicos), instituciones emisoras o relacionadas con el documento, firmas y sellos presentes en los documentos, y detalles adicionales relevantes (como medicamentos, detalles del tratamiento, etc.). Proporciona la información extraída en un formato estructurado y claro. Aquí tienes un ejemplo de cómo deberías presentar los resultados: - Tipo de documento: [Tipo de documento] - Nombre completo: [Nombre] - Fecha de emisión: [Fecha] - Institución emisora: [Institución] - Diagnóstico médico: [Diagnóstico] (si aplica) - Medicamentos y dosis: [Detalles] (si aplica) - Firmas y sellos: [Descripción] - Otros detalles relevantes: [Detalles] A continuación, adjunto una imagen del documento para que la revises: [Inserta aquí la imagen del documento]"
                    },
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/jpeg;base64,{base64_image}"
                        }
                    }
                ]
            }
        ],
        "max_tokens": 300
    }

    response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
    
    # Verificar y mostrar detalles del error
    if response.status_code != 200:
        st.error(f"Error: {response.status_code} - {response.text}")
        return None
    
    response_data = response.json()
    if 'error' in response_data:
        st.error(f"API Error: {response_data['error']['message']}")
    
    return response_data

def convert_pdf_page_to_image(pdf_path, page_num):
    pdf_document = fitz.open(pdf_path)
    page = pdf_document.load_page(page_num)
    pix = page.get_pixmap()
    img = PILImage.open(io.BytesIO(pix.tobytes()))
    return img

def process_pdfs_in_zip(zip_path, output_dir, api_key):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(output_dir)

    results = []

    for root, dirs, files in os.walk(output_dir):
        for file in files:
            if file.endswith('.pdf'):
                pdf_path = os.path.join(root, file)
                folder_name = os.path.basename(os.path.dirname(pdf_path))
                st.write(f"Processing {pdf_path}...")

                pdf_document = fitz.open(pdf_path)
                for page_num in range(len(pdf_document)):
                    img = convert_pdf_page_to_image(pdf_path, page_num)
                    temp_img_path = os.path.join(root, f"temp_image_page_{page_num}.png")
                    img.save(temp_img_path)

                    base64_image = encode_image(temp_img_path)
                    result = analyze_image(base64_image, api_key)
                    
                    if result and 'choices' in result and len(result['choices']) > 0:
                        analysis = result['choices'][0]['message']['content']
                    else:
                        analysis = "Analysis failed"

                    results.append({
                        "Folder": folder_name,
                        "File": file,
                        "Page": page_num,
                        "Analysis": analysis
                    })

    return results

def save_results_to_excel(results, output_excel_path):
    df = pd.DataFrame(results)
    df.to_excel(output_excel_path, index=False)

st.title("Análisis de Documentos PDF con OpenAI y Streamlit")
st.write(f"API Key: {api_key[:20]}...")

uploaded_file = st.file_uploader("Sube tu archivo ZIP", type=["zip"])

if uploaded_file is not None:
    temp_dir = tempfile.gettempdir()
    zip_path = os.path.join(temp_dir, uploaded_file.name)
    
    with open(zip_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    output_dir = os.path.join(temp_dir, "extracted_pdfs")
    output_excel_path = os.path.join(temp_dir, "results.xlsx")

    with st.spinner('Procesando PDFs...'):
        results = process_pdfs_in_zip(zip_path, output_dir, api_key)
        save_results_to_excel(results, output_excel_path)

    # Leer el archivo de resultados y concatenar análisis por folder
    data = pd.read_excel(output_excel_path, sheet_name='Sheet1')
    result = data.groupby('Folder')['Analysis'].apply(lambda x: ' '.join(x)).reset_index()
    result.columns = ['Folder', 'Análisis_concatenado']
    
    # Guardar los resultados finales en Excel
    final_output_path = os.path.join(temp_dir, "final_results.xlsx")
    result.to_excel(final_output_path, index=False)

    st.success(f"Análisis completado. Los resultados se han guardado en {final_output_path}.")
    st.download_button(
        label="Descargar resultados en Excel",
        data=open(final_output_path, "rb").read(),
        file_name="final_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


