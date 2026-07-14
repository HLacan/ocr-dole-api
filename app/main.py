"""
API OCR Dole — expone el motor de extracción como endpoints HTTP para n8n.

Flujo pensado (Python nunca toca Drive ni arma el Excel — eso es trabajo
de n8n):

1) POST /split-bl-pdf     → recibe el PDF combinado de ~100 BLs (1 correo,
                             1 sola vez por semana). Devuelve, por cada BL:
                             el PDF individual (base64) + los datos clave
                             (gyeprq, contenedor, destino, tipo). Con esto
                             n8n crea las carpetas en Drive y arranca el Excel.

2) POST /classify-document → recibe UN documento suelto (factura, fito,
                             manifiesto o marchamos) — se llama una vez por
                             archivo, a medida que van llegando en cada
                             correo (el primero o los siguientes). Devuelve
                             el tipo detectado + los campos extraídos, para
                             que n8n actualice las filas correspondientes
                             del Excel por contenedor.

Endpoints heredados (all-in-one, arman el Excel completo del lado de
Python): /procesar, /upload + /finalize. Se mantienen por si sirven para
otro flujo, pero el flujo recomendado con n8n es el de arriba.

Headers en todos los casos: X-API-Key: <API_KEY>
"""

import os
import re
import shutil
import tempfile
import time
import base64
from typing import List, Optional

from fastapi import FastAPI, UploadFile, File, Form, Header, HTTPException
from fastapi.responses import FileResponse, JSONResponse

from extract_dole_core import (
    run_extraction, classify_file,
    split_bl_pdf_pages, classify_single_document,
    flatten_zip_bytes,
)

API_KEY = os.environ.get("OCR_DOLE_API_KEY", "")  # define esto en el hosting (variable de entorno)

app = FastAPI(title="OCR Dole API")

BATCH_ROOT = os.path.join(tempfile.gettempdir(), "ocr_dole_batches")
os.makedirs(BATCH_ROOT, exist_ok=True)


def check_api_key(x_api_key: Optional[str]):
    if API_KEY and x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="API key inválida o faltante")


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/process-email")
async def process_email_endpoint(
    file: UploadFile = File(..., description="El adjunto crudo del correo (zip con o sin anidados, o un archivo suelto)"),
    x_api_key: Optional[str] = Header(None),
):
    """
    El endpoint 'todo en uno' para n8n: recibe el adjunto TAL CUAL viene del
    correo y hace todo el trabajo de clasificacion del lado de Python:

      1. Desempaca (flatten_zip_bytes) -- zips anidados a cualquier profundidad.
      2. Entre los archivos resultantes, identifica el PDF de ~100 BLs (si
         viene en este correo) y lo divide pagina por pagina.
      3. Cualquier otro archivo lo clasifica (factura/fito/manifiesto/marchamos)
         y le extrae sus datos.

    n8n ya no necesita decidir "es el PDF de BLs?" ni hacer una llamada HTTP
    por archivo -- solo llama esto UNA vez por correo, y con el JSON que
    regresa arma las carpetas, sube los PDFs y llena el Excel. Python nunca
    toca Drive ni Sheets.
    """
    check_api_key(x_api_key)
    content = await file.read()

    try:
        archivos = flatten_zip_bytes(file.filename, content)
    except Exception as e:
        return JSONResponse(status_code=422, content={"error": f"No se pudo procesar el adjunto: {e}"})

    bls = []
    documentos = []
    advertencias = []

    for archivo in archivos:
        fname = archivo["filename"]
        data = archivo["content"]

        if re.match(r'^BLS|VIA PQU', fname, re.IGNORECASE):
            try:
                paginas = split_bl_pdf_pages(data)
            except Exception as e:
                advertencias.append(f"{fname}: no se pudo dividir el PDF de BLs ({e})")
                continue

            for p in paginas:
                if p["warning"]:
                    advertencias.append(p["warning"])
                for campo in ("gyeprq", "contenedor", "destino", "tipo"):
                    if p[campo] == "???":
                        advertencias.append(f"{fname} pagina {p['pagina']}: no se encontro {campo}")

                nombre_base = p["gyeprq"] if p["gyeprq"] != "???" else f"SIN_CODIGO_pagina_{p['pagina']}"
                bls.append({
                    "gyeprq": p["gyeprq"],
                    "contenedor": p["contenedor"],
                    "destino": p["destino"],
                    "tipo": p["tipo"],
                    "filename": f"{nombre_base}.pdf",
                    "pdf_base64": base64.b64encode(p["pdf_bytes"]).decode("ascii"),
                })
            continue

        tmp_dir = tempfile.mkdtemp(prefix="ocr_dole_batch_")
        tmp_path = os.path.join(tmp_dir, fname)
        try:
            with open(tmp_path, "wb") as fh:
                fh.write(data)
            resultado = classify_single_document(tmp_path)
        except Exception as e:
            resultado = {"tipo": "DESCONOCIDO", "motivo": f"error inesperado: {e}"}
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

        resultado["filename"] = fname
        resultado["content_base64"] = base64.b64encode(data).decode("ascii")
        documentos.append(resultado)

    return {
        "total_bls": len(bls),
        "bls": bls,
        "total_documentos": len(documentos),
        "documentos": documentos,
        "advertencias": advertencias,
    }


@app.post("/flatten-attachment")
async def flatten_attachment_endpoint(
    file: UploadFile = File(..., description="El adjunto del correo -- puede ser un zip (con zips anidados adentro) o un archivo suelto"),
    x_api_key: Optional[str] = Header(None),
):
    """
    Reemplaza TODO el tramo de descompresion que antes vivia en n8n
    (Compression x2, IF "es zip anidado", Code "separar binarios",
    sub-workflow recursivo). Recibe el adjunto tal cual, y devuelve la
    lista PLANA de archivos reales que contiene, sin importar cuantos
    niveles de zip haya adentro. Ya viene filtrado (sin carpetas, sin
    extensiones que no nos sirven).
    """
    check_api_key(x_api_key)
    content = await file.read()

    try:
        archivos = flatten_zip_bytes(file.filename, content)
    except Exception as e:
        return JSONResponse(status_code=422, content={"error": f"No se pudo procesar el archivo: {e}"})

    return {
        "total_archivos": len(archivos),
        "archivos": [
            {
                "filename": a["filename"],
                "content_base64": base64.b64encode(a["content"]).decode("ascii"),
            }
            for a in archivos
        ],
    }


@app.post("/split-bl-pdf")
async def split_bl_pdf_endpoint(
    file: UploadFile = File(..., description="El PDF combinado de BLs (1 pagina = 1 BL, ej. ~100 paginas)"),
    x_api_key: Optional[str] = Header(None),
):
    """
    Divide el PDF de N hojas y devuelve, por cada pagina/BL:
    gyeprq, contenedor, destino, tipo, filename y el PDF de esa sola
    pagina en base64. No escribe nada a disco de forma persistente ni
    toca Drive — n8n decide que hacer con cada resultado (crear carpeta,
    subir el PDF, escribir la fila en el Excel).
    """
    check_api_key(x_api_key)
    content = await file.read()

    try:
        paginas = split_bl_pdf_pages(content)
    except Exception as e:
        return JSONResponse(status_code=422, content={"error": f"No se pudo procesar el PDF: {e}"})

    bls = []
    advertencias = []
    for p in paginas:
        if p["warning"]:
            advertencias.append(p["warning"])
        for campo in ("gyeprq", "contenedor", "destino", "tipo"):
            if p[campo] == "???":
                advertencias.append(f"pagina {p['pagina']}: no se encontro {campo}")

        nombre_base = p["gyeprq"] if p["gyeprq"] != "???" else f"SIN_CODIGO_pagina_{p['pagina']}"
        bls.append({
            "gyeprq": p["gyeprq"],
            "contenedor": p["contenedor"],
            "destino": p["destino"],
            "tipo": p["tipo"],
            "filename": f"{nombre_base}.pdf",
            "pdf_base64": base64.b64encode(p["pdf_bytes"]).decode("ascii"),
        })

    return {
        "total_paginas": len(bls),
        "bls": bls,
        "advertencias": advertencias,
    }


@app.post("/classify-document")
async def classify_document_endpoint(
    file: UploadFile = File(..., description="Un documento suelto: factura, fito, manifiesto o marchamos"),
    x_api_key: Optional[str] = Header(None),
):
    """
    Clasifica y extrae los datos de UN documento suelto (sin necesitar el
    resto de la semana). Pensado para llamarse una vez por archivo segun
    van llegando en cada correo. n8n usa el "tipo" devuelto para decidir
    como actualizar el Excel (por contenedor, o el manifiesto/marchamos
    que aplica a toda la semana).
    """
    check_api_key(x_api_key)
    content = await file.read()
    if len(content) == 0:
        return {"tipo": "DESCONOCIDO", "motivo": "archivo vacio"}

    tmp_path = os.path.join(tempfile.mkdtemp(prefix="ocr_dole_doc_"), file.filename)
    with open(tmp_path, "wb") as f:
        f.write(content)

    try:
        resultado = classify_single_document(tmp_path)
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"Error inesperado: {e}"})
    finally:
        shutil.rmtree(os.path.dirname(tmp_path), ignore_errors=True)

    return resultado


def _batch_dir(batch_id: str) -> str:
    safe_id = re.sub(r'[^A-Za-z0-9_\-]', '_', batch_id)[:100] or "default"
    d = os.path.join(BATCH_ROOT, safe_id)
    os.makedirs(d, exist_ok=True)
    return d


def _safe_join(workdir, relname):
    """
    Guarda un archivo preservando la subcarpeta que venga incrustada en su
    nombre (ej. 'OTROS EXPORTADORES/ASOAGRIBAL/factura.pdf'), porque el
    motor de extracción detecta el PROVEEDOR a partir de esa carpeta.
    Sanea contra path traversal (../, rutas absolutas).
    """
    relname = relname.replace("\\", "/").lstrip("/")
    parts = [p for p in relname.split("/") if p not in ("", ".", "..")]
    dest = os.path.join(workdir, *parts)
    os.makedirs(os.path.dirname(dest), exist_ok=True)
    return dest


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
        # Guardar todos los archivos recibidos, preservando la subcarpeta
        # que venga en el nombre (para que PROVEEDOR se detecte bien).
        # Se ignoran entradas de 0 bytes: son marcadores de carpeta que
        # algunos zips incluyen (ej. "OTROS EXPORTADORES/" como entrada
        # separada), no archivos reales — si se guardaran, chocarían con
        # la creación de subcarpetas de archivos que sí vienen adentro.
        saved_paths = []
        for f in files:
            content = await f.read()
            if len(content) == 0:
                continue
            dest = _safe_join(workdir, f.filename)
            with open(dest, "wb") as out:
                out.write(content)
            saved_paths.append(dest)

        marchamo_path = None
        if marchamo_file is not None:
            marchamo_path = _safe_join(workdir, marchamo_file.filename)
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


@app.post("/upload")
async def upload(
    batch_id: str = Form(..., description="Identificador del lote — usa el número de semana"),
    file: UploadFile = File(...),
    x_api_key: Optional[str] = Header(None),
):
    """
    Sube UN archivo al lote `batch_id`. Pensado para llamarse una vez por
    cada ítem que llega al nodo HTTP Request de n8n (así es como n8n
    ejecuta el nodo de forma natural — no hay que combinar los archivos
    en un solo ítem antes de mandarlos).
    """
    check_api_key(x_api_key)
    content = await file.read()
    if len(content) == 0:
        # Entradas de carpeta vacía que trae el zip (no son archivos reales)
        return {"status": "skipped_empty", "file": file.filename}
    dest = _safe_join(_batch_dir(batch_id), file.filename)
    with open(dest, "wb") as out:
        out.write(content)
    return {"status": "ok", "batch_id": batch_id, "file": file.filename}


@app.post("/finalize")
async def finalize(
    batch_id: str = Form(..., description="El mismo batch_id usado en las llamadas a /upload"),
    semana: str = Form("?"),
    marchamo_file: Optional[UploadFile] = File(None),
    cleanup: bool = Form(True, description="Si borra el lote después de generar el Excel"),
    x_api_key: Optional[str] = Header(None),
):
    """
    Procesa TODOS los archivos que se hayan subido a `batch_id` vía /upload
    y devuelve el Excel. Llamar UNA SOLA VEZ después de que todos los
    /upload de ese lote hayan terminado (en n8n: activa "Execute Once" en
    este nodo, ya que recibirá varios ítems pero solo debe correr una vez).
    """
    check_api_key(x_api_key)
    workdir = _batch_dir(batch_id)

    if not any(os.scandir(workdir)):
        return JSONResponse(status_code=422, content={
            "error": f"No hay archivos subidos todavía para el batch_id '{batch_id}'. "
                     f"Llama a /upload primero."
        })

    marchamo_path = None
    if marchamo_file is not None:
        content = await marchamo_file.read()
        if content:
            marchamo_path = _safe_join(workdir, marchamo_file.filename)
            with open(marchamo_path, "wb") as out:
                out.write(content)

    bl_file = None
    for root, _dirs, fnames in os.walk(workdir):
        for fname in fnames:
            if classify_file(fname) == 'BL':
                bl_file = os.path.join(root, fname)
                break
        if bl_file:
            break

    out_dir = tempfile.mkdtemp(prefix="ocr_dole_out_")
    output_path = os.path.join(out_dir, f"DOLE-{semana}.xlsx")

    try:
        result = run_extraction(
            input_folder=workdir,
            output_path=output_path,
            semana=semana,
            bl_file=bl_file,
            marchamo_file=marchamo_path,
        )
    except ValueError as e:
        return JSONResponse(status_code=422, content={"error": str(e)})
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"Error inesperado: {e}"})

    response = FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=os.path.basename(output_path),
    )
    response.headers["X-Warnings-Count"] = str(len(result["warnings"]))
    response.headers["X-Rows-Count"] = str(result["rows"])

    if cleanup:
        shutil.rmtree(workdir, ignore_errors=True)

    return response
