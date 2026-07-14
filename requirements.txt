# OCR Dole API

API HTTP que envuelve la extracción de documentos Dole (BL + facturas + fitos)
y genera el Excel de control semanal. Pensada para ser llamada desde n8n con
un nodo HTTP Request, sin necesitar acceso SSH al servidor de n8n.

## Endpoints

- `GET /health` → chequeo rápido, responde `{"status":"ok"}`
- `POST /procesar` → recibe los archivos, devuelve el Excel

### POST /procesar

**Headers:**
- `X-API-Key: <tu clave>` (si configuraste `OCR_DOLE_API_KEY`)

**Body (multipart/form-data):**
- `files`: uno o más archivos — el PDF combinado de BLs + todas las
  facturas/fitos/manifiestos de la semana (se detecta solo cuál es el BL)
- `semana`: texto, ej. `"25"`
- `marchamo_file` (opcional): el Excel de LISTADO DE MARCHAMOS

**Respuesta:** el archivo `.xlsx` generado, más headers:
- `X-Rows-Count`: filas procesadas
- `X-Warnings-Count`: cantidad de advertencias (bultos que no cuadran, TIPO
  ambiguo, contenedores sin factura/fito, etc.)

## Desplegar en Render.com (sin necesitar SSH)

1. Sube esta carpeta a un repositorio de GitHub (puede ser privado).
2. En Render.com → **New + → Web Service** → conecta el repo.
3. Render detecta el `Dockerfile` automáticamente. Elige el plan gratuito o
   el más económico (esto procesa PDFs, no necesita mucha RAM).
4. En **Environment**, agrega la variable:
   - `OCR_DOLE_API_KEY` = una clave que inventes (ej. un password largo)
5. Deploy. Cuando termine, Render te da una URL tipo
   `https://ocr-dole-api.onrender.com`.
6. Prueba: `curl https://ocr-dole-api.onrender.com/health` → debe responder
   `{"status":"ok"}`.

## Probar localmente

```bash
pip install -r requirements.txt
cd app
uvicorn main:app --reload
# luego, en otra terminal:
curl -F "files=@BLS_DAC608.pdf" -F "semana=25" http://localhost:8000/procesar -o salida.xlsx
```

## Nodo HTTP Request en n8n

- **Method:** POST
- **URL:** `https://ocr-dole-api.onrender.com/procesar`
- **Authentication:** None (la API key va en un header, no en auth nativo)
- **Headers:** `X-API-Key` = tu clave
- **Body Content Type:** `multipart-form-data`
- **Body Parameters:**
  - `files` → tipo `n8n Binary File`, uno por cada archivo acumulado en el
    flujo (el PDF del BL + facturas + fitos)
  - `semana` → tipo texto, ej. `{{ $json.semana }}`
- **Response Format:** File (para que n8n reciba el binario del Excel)

Con la respuesta ya puedes conectar un nodo **Google Drive → Upload** o
**Write File**, y opcionalmente leer `X-Warnings-Count` de los headers de
respuesta para decidir si mandar una alerta.
