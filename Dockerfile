FROM python:3.11-slim

# tesseract-ocr: motor de OCR usado como respaldo cuando un PDF trae texto
# corrupto (fuente embebida sin mapa Unicode, común en algunos proveedores).
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr tesseract-ocr-spa libgl1 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app/ .

# Render/Railway inyectan la variable PORT; localmente usa 8000
ENV PORT=8000
CMD uvicorn main:app --host 0.0.0.0 --port ${PORT}
