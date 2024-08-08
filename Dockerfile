# Usa una imagen base de Python
FROM python:3.11-slim

# Establece el directorio de trabajo en /app
WORKDIR /app

# Copia los archivos de requerimientos y el código fuente en el contenedor
COPY requirements.txt requirements.txt
COPY app.py app.py

# Instala las dependencias
RUN pip install --no-cache-dir -r requirements.txt

# Expone el puerto en el que Streamlit se ejecuta por defecto
EXPOSE 8501

# Comando para ejecutar la aplicación
CMD ["sh", "-c", "streamlit run app.py"]