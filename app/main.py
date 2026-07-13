"""
API OCR Dole — envuelve extract_dole_core.run_extraction() en un endpoint HTTP.

Uso desde n8n (nodo HTTP Request):
  POST /procesar
  Headers: X-API-Key: <API_KEY>
  Body: multipart/form-data
    - bl_file: el PDF combinado de BLs (opcional si va incluido en 'files')
    - files: uno o más archivos (facturas, fitos, manifiestos)
    - semana: texto, ej. "25"
    - marchamo_file: opcional, el Excel de LISTADO DE MARCHAMOS

Respuesta: el archivo .xlsx generado, listo para descargar/reenviar.
"""

import os
import shutil
import tempfile
from typing import List, Optional

from fastapi import FastAPI, UploadFile, File, Form, Header, HTTPException
from fastapi.responses import FileResponse, JSONResponse

from extract_dole_core import run_extraction, classify_file

API_KEY = os.environ.get("OCR_DOLE_API_KEY", "")  # define esto en el hosting (variable de entorno)

app = FastAPI(title="OCR Dole API")


def check_api_key(x_api_key: Optional[str]):
    if API_KEY and x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="API key inválida o faltante")


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/procesar")
async def procesar(
    files: List[UploadFile] = File(..., description="Todos los archivos de la semana: el PDF de BLs + facturas + fitos + manifiestos"),
    semana: str = Form("?"),
    marchamo_file: Optional[UploadFile] = File(None),
    x_api_key: Optional[str] = Header(None),
):
    check_api_key(x_api_key)

    workdir = tempfile.mkdtemp(prefix="ocr_dole_")
    try:
        # Guardar todos los archivos recibidos en una carpeta plana
        saved_paths = []
        for f in files:
            dest = os.path.join(workdir, f.filename)
            with open(dest, "wb") as out:
                shutil.copyfileobj(f.file, out)
            saved_paths.append(dest)

        marchamo_path = None
        if marchamo_file is not None:
            marchamo_path = os.path.join(workdir, marchamo_file.filename)
            with open(marchamo_path, "wb") as out:
                shutil.copyfileobj(marchamo_file.file, out)

        # Detectar automáticamente cuál de los archivos subidos es el BL combinado
        bl_file = None
        for p in saved_paths:
            if classify_file(os.path.basename(p)) == 'BL':
                bl_file = p
                break

        output_path = os.path.join(workdir, f"DOLE-{semana}.xlsx")

        result = run_extraction(
            input_folder=workdir,
            output_path=output_path,
            semana=semana,
            bl_file=bl_file,
            marchamo_file=marchamo_path,
        )

        response = FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=os.path.basename(output_path),
        )
        # Adjuntar advertencias como header (n8n puede leerlas de la respuesta)
        response.headers["X-Warnings-Count"] = str(len(result["warnings"]))
        response.headers["X-Rows-Count"] = str(result["rows"])
        return response

    except ValueError as e:
        return JSONResponse(status_code=422, content={"error": str(e)})
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"Error inesperado: {e}"})
    # Nota: no borramos workdir aquí porque FileResponse necesita el archivo
    # después de retornar; el sistema operativo/hosting limpia /tmp periódicamente.
