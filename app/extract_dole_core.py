#!/usr/bin/env python3
"""
extract_dole.py  —  OCR Dole: Excel de control semanal de embarques

Relaciones clave:
  1 BL (GYEPRQ...) → N facturas (por contenedor compartido)
  1 Factura        → N fitos (por contenedor + bultos que deben cuadrar)
  Todo se cruza por CONTENEDOR ([A-Z]{4}[0-9]{7}) — NO por nombre de carpeta.

Formato de entrada:
  --input-folder: TODOS los archivos de la semana (facturas, fitos,
    manifiestos). No importa si vienen sueltos o en subcarpetas — el
    script recorre todo recursivamente y cruza cada documento con su
    contenedor según el propio contenido del PDF.
  --bl-file: el PDF combinado de BLs (una página = un BL, ej. "BLS DAC608
    VIA PQU ORIG.pdf", ~100 páginas). Si se omite, el script busca
    automáticamente dentro de --input-folder un archivo con nombre de BL
    (ej. que empiece con "BLS").

Uso:
  python3 extract_dole.py \
    --input-folder  "ruta/DOLE SEMANA 6" \
    --bl-file       "ruta/BLS DAC608 VIA PQU ORIG.pdf" \
    --output        "ruta/DOLE-2026-N.xlsx" \
    --semana        "6" \
    [--marchamo-file "ruta/LISTADO DE MARCHAMOS.xlsx"]
"""

import os, re, sys, argparse, subprocess

def ensure_deps():
    try:
        import pdfplumber, openpyxl
    except ImportError:
        subprocess.run([sys.executable, "-m", "pip", "install",
                        "pdfplumber", "openpyxl", "--break-system-packages", "-q"], check=True)
ensure_deps()

import pdfplumber
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

CONTAINER_RE = re.compile(r'\b([A-Z]{4}[0-9]{7})\b')

# ─────────────────────────────────────────────
# Clasificación de archivos
# ─────────────────────────────────────────────

def classify_file(filename):
    name = filename.upper()
    base = os.path.splitext(name)[0]
    ext  = os.path.splitext(name)[1]
    if 'CONSULTAMANIFIESTO' in name or ('CONSULTA' in name and 'MANIFIESTO' in name):
        return 'MANIFIESTO'
    if name.startswith('BLS') or (name.startswith('BL') and 'DPC' in name):
        return 'BL'
    if 'FACT' in name and ext == '.PDF':
        return 'FACTURA'
    if 'INVOICE' in name and ext == '.PDF':
        return 'FACTURA'
    if re.match(r'^EC-UB-\d+', base) and ext == '.PDF':
        return 'FACTURA'
    if 'PHYTO' in name or 'FITO' in name:
        return 'FITO'
    if re.match(r'^(AA-)?[0-9]+P$', base) and ext == '.PDF':
        return 'FITO'
    if 'MANIFIESTO' in name and ext == '.PDF':
        return 'MANIFIESTO'
    return 'UNKNOWN'

def pdf_text(path):
    try:
        with pdfplumber.open(path) as pdf:
            return "\n".join(p.extract_text() or "" for p in pdf.pages)
    except Exception:
        return ""

def file_mtime(path):
    try: return os.path.getmtime(path)
    except: return 0

def newest_file(paths):
    return max(paths, key=file_mtime) if paths else None

# ─────────────────────────────────────────────
# Extracción BL — PDF combinado (1 página = 1 BL)
# ─────────────────────────────────────────────

def split_bl_pdf(bl_pdf_path):
    """
    El BL llega como un solo PDF de varias páginas (ej. 100 páginas).
    Cada página es un BL individual, identificado por su código
    GYEPRQ###### presente en el texto de esa página.

    Retorna: dict { "GYEPRQ356004": texto_pagina, ... }
    Si el mismo código aparece en más de una página (poco común), se
    concatena el texto de todas esas páginas.
    """
    pages_by_code = {}
    try:
        with pdfplumber.open(bl_pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                m = re.search(r'\b(GYEPRQ\d+)\b', text, re.IGNORECASE)
                if not m:
                    continue
                code = m.group(1).upper()
                if code in pages_by_code:
                    pages_by_code[code] += "\n" + text
                else:
                    pages_by_code[code] = text
    except Exception as e:
        print(f"  ERROR leyendo BL combinado {bl_pdf_path}: {e}")
    return pages_by_code


def extract_bl_from_text(text, bl_code):
    """
    Extrae de una página del BL combinado:
      BL, CONTENEDOR, DESTINO (Place of delivery), TIPO (ISO container type)

    Estructura clave:
      Línea "GUAYAQUIL {PLACE_OF_DELIVERY}"        → DESTINO
      Línea "1 x {TYPE} CONTAINERS"                → TIPO (texto libre)
      Tabla "XXXX1234567 SEAL 40 RF 960 BOX ..."   → CONTENEDOR + TIPO (tabla)

    El TIPO de texto libre y el de la tabla a veces no coinciden (ej. RC
    vs RF). En ese caso se usa el de la TABLA (específico del contenedor
    real) y se genera una advertencia para revisión manual.
    """
    lines = text.split('\n')

    bl = bl_code or "???"
    contenedor = "???"
    destino = "???"
    tipo_texto = None
    tipo_tabla = None
    warning = None

    # DESTINO: línea "GUAYAQUIL {Place of Delivery}" que contiene el país destino
    for line in lines:
        stripped = line.strip()
        if stripped.upper().startswith("GUAYAQUIL ") and "USA" in stripped.upper():
            destino_raw = stripped[len('GUAYAQUIL '):].strip()
            if destino_raw and not re.match(r'^\d', destino_raw):
                destino = destino_raw
                break

    # TIPO (texto libre): "1 x 40RV CONTAINERS"
    m_tipo_texto = re.search(r'\d+\s*x\s*(40[A-Z]{2})\s+CONTAINER', text, re.IGNORECASE)
    if m_tipo_texto:
        tipo_texto = m_tipo_texto.group(1).upper()

    # CONTENEDOR + TIPO (tabla): "CAAU4102550 132547E 40 RF 960 BOX 18941.76 KGS"
    m_row = re.search(
        r'\b([A-Z]{4}[0-9]{7})\s+\S+\s+(\d{2})\s*([A-Z]{2})\s+(\d+)\s*BOX',
        text, re.IGNORECASE
    )
    if m_row:
        contenedor = m_row.group(1).upper()
        tipo_tabla = (m_row.group(2) + m_row.group(3)).upper()
    else:
        containers = [c for c in CONTAINER_RE.findall(text)
                      if not c.startswith('GYE') and not c.startswith('AG1')]
        if containers:
            contenedor = containers[0]

    tipo = tipo_tabla or tipo_texto or "???"
    if tipo_texto and tipo_tabla and tipo_texto != tipo_tabla:
        warning = (f"TIPO no coincide en {bl} [{contenedor}]: "
                   f"texto libre={tipo_texto} vs tabla={tipo_tabla} "
                   f"→ se usó el de la tabla ({tipo_tabla}), revisar manualmente")

    return {
        "BL": bl,
        "CONTENEDOR": contenedor,
        "DESTINO": destino,
        "TIPO": tipo,
    }, warning

# ─────────────────────────────────────────────
# Extracción MANIFIESTO
# ─────────────────────────────────────────────

def extract_manifiesto(path):
    text = pdf_text(path)
    m = re.search(r'Numero de manifiesto\s+(AG\d+)', text)
    if m: return m.group(1)
    m = re.search(r'(AG\d{9,})', text)
    return m.group(1) if m else "???"

# ─────────────────────────────────────────────
# Extracción FITO
# ─────────────────────────────────────────────

def extract_fito(path):
    """Returns {number, bultos, contenedor}"""
    fname_base = os.path.splitext(os.path.basename(path))[0]
    text = pdf_text(path)

    # Número de certificado — siempre con prefijo AA-
    if re.match(r'^(AA-)?[0-9]+P$', fname_base, re.IGNORECASE):
        num = fname_base.upper()
        if not num.startswith('AA-'):
            num = 'AA-' + num
    else:
        m = re.search(r'N[°º°]\s*([0-9]{15,}P)', text)
        num = ('AA-' + m.group(1)) if m else fname_base

    # Bultos
    bultos = None
    m = re.search(r'(\d+)\s+Box(?:es)?', text, re.IGNORECASE)
    if m: bultos = int(m.group(1))

    # Contenedor mencionado en el fito
    containers = [c for c in CONTAINER_RE.findall(text)
                  if not c.startswith('GYE') and not c.startswith('AG1')]
    contenedor = containers[0] if containers else None

    return {"number": num, "bultos": bultos, "contenedor": contenedor}

# ─────────────────────────────────────────────
# Extracción FACTURA
# ─────────────────────────────────────────────

def extract_factura(path):
    text = pdf_text(path)
    fname = os.path.basename(path).upper()
    base  = os.path.splitext(fname)[0]

    result = {
        "proveedor": "???", "factura": "???", "item": "???",
        "contenedor": "???", "bultos_per_container": {},
    }

    # ── UBESA (EC-UB-XXXX.pdf) — puede listar varios contenedores ──
    if re.match(r'^EC-UB-', base):
        result["proveedor"] = "UBESA"
        result["factura"]   = base
        result["item"] = "PLATANO" if "PLANTAIN" in text.upper() else "BANANO"
        for line in text.split('\n'):
            tokens = line.split()
            if len(tokens) >= 3 and CONTAINER_RE.match(tokens[0]):
                cont = tokens[0]
                try:
                    boxes = int(tokens[1])
                    result["bultos_per_container"][cont] = boxes
                except ValueError:
                    pass
        return result

    # ── ASOPEQ / Tierra Fertil ──
    if ('TIERRA FERTIL' in text.upper()
            or 'ASOCIACION DE PEQUEÑOS' in text.upper()
            or 'ASOCIACIÓN' in text.upper()):
        result["proveedor"] = "ASOPEQ"
        m = re.search(r'FACTURA\s+([\d\-]+)', text)
        if m: result["factura"] = m.group(1)
        result["item"] = "PLATANO" if ("PLATANO" in text.upper() or "PLANTAIN" in text.upper()) else "BANANO"
        m2 = re.search(r'CONTENEDOR\(es\)\s+([A-Z]{4}[0-9]{7})', text)
        if m2:
            result["contenedor"] = m2.group(1)
        else:
            containers = [c for c in CONTAINER_RE.findall(text)
                          if not c.startswith('GYE') and not c.startswith('AG1')]
            if containers: result["contenedor"] = containers[0]
        m3 = re.search(r'CANTIDAD\s+(\d+)', text)
        if m3:
            boxes = int(m3.group(1))
            if result["contenedor"] != "???":
                result["bultos_per_container"][result["contenedor"]] = boxes
        return result

    # ── Marplantis ──
    if 'MARPLANTIS' in text.upper():
        result["proveedor"] = "MARPLANTIS"
        m = re.search(r'No\.\s*([\d\-]+)', text)
        if not m: m = re.search(r'(001-\d{3}-\d{8})', text)
        if m: result["factura"] = m.group(1)
        result["item"] = "PLATANO" if "PLANTAIN" in text.upper() else "BANANO"
        m2 = re.search(r'CONTAINER[:\s]+([A-Z]{4}[0-9]{7})', text)
        if m2: result["contenedor"] = m2.group(1)
        m3 = re.search(r'(?:Bananas?|Plantains?)\s+(\d+)\s+[\d\.,]+', text, re.IGNORECASE)
        if m3:
            boxes = int(m3.group(1))
            if result["contenedor"] != "???":
                result["bultos_per_container"][result["contenedor"]] = boxes
        return result

    # ── Fallback ──
    m = re.search(r'INVOICE\s+([\w\-]+)', text, re.IGNORECASE)
    if m: result["factura"] = m.group(1)
    containers = [c for c in CONTAINER_RE.findall(text)
                  if not c.startswith('GYE') and not c.startswith('AG1')]
    if containers: result["contenedor"] = containers[0]
    return result

# ─────────────────────────────────────────────
# Recorrer TODOS los archivos (sin importar estructura de carpetas)
# ─────────────────────────────────────────────

def classify_and_extract_all(input_folder, exclude_path=None):
    """
    Recorre input_folder recursivamente (sirve tanto si los archivos vienen
    sueltos como organizados en subcarpetas) y extrae cada factura/fito/
    manifiesto. exclude_path es el PDF combinado de BLs, que ya se procesó
    aparte y no debe reclasificarse.
    """
    exclude_abs = os.path.abspath(exclude_path) if exclude_path else None
    facturas, fitos, manifiestos = [], [], []
    for root, _dirs, files in os.walk(input_folder):
        for fname in files:
            fpath = os.path.join(root, fname)
            if exclude_abs and os.path.abspath(fpath) == exclude_abs:
                continue
            kind = classify_file(fname)
            if kind == 'FACTURA':
                fd = extract_factura(fpath)
                fd['_mtime'] = file_mtime(fpath)
                facturas.append(fd)
            elif kind == 'FITO':
                fi = extract_fito(fpath)
                fi['_mtime'] = file_mtime(fpath)
                fitos.append(fi)
            elif kind == 'MANIFIESTO':
                manifiestos.append({"number": extract_manifiesto(fpath), "_mtime": file_mtime(fpath)})
    return facturas, fitos, manifiestos


def dedup_by_key(items, key):
    """Dedup por valor de `key`, gana el más reciente (_mtime). Los items
    sin valor válido de `key` (None o '???') NO se colapsan entre sí — se
    conservan todos, porque agruparlos perdería datos distintos."""
    best = {}
    passthrough = []
    for it in items:
        k = it.get(key)
        if not k or k == "???":
            passthrough.append(it)
            continue
        mt = it.get('_mtime', 0)
        if k not in best or mt > best[k][1]:
            best[k] = (it, mt)
    return [v[0] for v in best.values()] + passthrough


# ─────────────────────────────────────────────
# Construir fila del Excel para un CONTENEDOR
# ─────────────────────────────────────────────

def build_row(container, bl_data, bl_warning, facturas, fitos, manifiestos):
    row = {
        "BL": bl_data["BL"] if bl_data else "???",
        "PROVEEDOR": "???", "CONTENEDOR": container,
        "TIPO": bl_data["TIPO"] if bl_data else "???",
        "DESTINO": bl_data["DESTINO"] if bl_data else "???",
        "FACTURA": "???", "ITEM": "???", "FITO": "???", "MANIFIESTO": "???",
    }
    warnings = []
    if bl_warning:
        warnings.append(bl_warning)
    if not bl_data:
        warnings.append(f"{container}: no se encontró BL para este contenedor")

    # FACTURAS que mencionan este contenedor (directo o en bultos_per_container)
    matched_facturas = [
        f for f in facturas
        if f.get("contenedor") == container or container in f.get("bultos_per_container", {})
    ]
    all_bultos_fact = {}
    if matched_facturas:
        row["PROVEEDOR"] = matched_facturas[0]["proveedor"]
        row["ITEM"]      = matched_facturas[0]["item"]
        nums = [f["factura"] for f in matched_facturas if f["factura"] != "???"]
        row["FACTURA"] = " / ".join(dict.fromkeys(nums)) if nums else "???"
        for fd in matched_facturas:
            all_bultos_fact.update(fd.get("bultos_per_container", {}))
    else:
        warnings.append(f"{container}: no se encontró factura para este contenedor")

    # FITOS de este contenedor
    matched_fitos = [fi for fi in fitos if fi.get("contenedor") == container]
    if matched_fitos:
        row["FITO"] = " / ".join(fi["number"] for fi in matched_fitos)
    else:
        warnings.append(f"{container}: no se encontró fito para este contenedor")

    # MANIFIESTO — normalmente uno solo cubre todo el embarque/semana
    if len(manifiestos) == 1:
        row["MANIFIESTO"] = manifiestos[0]["number"]
    elif len(manifiestos) > 1:
        row["MANIFIESTO"] = manifiestos[0]["number"]
        warnings.append(f"{container}: hay {len(manifiestos)} manifiestos distintos, "
                         f"se usó el primero — revisar manualmente")

    # Validación bultos factura vs fito
    bultos_fact = all_bultos_fact.get(container)
    bultos_fito = next((fi["bultos"] for fi in matched_fitos), None)
    if bultos_fact and bultos_fito and bultos_fact != bultos_fito:
        warnings.append(
            f"BULTOS NO CUADRAN en {container}: Factura={bultos_fact} Fito={bultos_fito}"
        )

    return row, warnings

# ─────────────────────────────────────────────
# LISTADO DE MARCHAMOS
# ─────────────────────────────────────────────

def load_marchamo_lookup(marchamo_file):
    """CONTENEDOR → 'SAT-GT-{numero}'"""
    lookup = {}
    if not marchamo_file or not os.path.isfile(marchamo_file):
        return lookup
    try:
        wb = openpyxl.load_workbook(marchamo_file)
        ws = wb.active
        for row in ws.iter_rows(min_row=8, values_only=True):
            if row[1] and row[5]:
                container    = str(row[1]).strip()
                marchamo_num = str(int(row[5])) if isinstance(row[5], float) else str(row[5]).strip()
                lookup[container] = f"SAT-GT-{marchamo_num}"
    except Exception as e:
        print(f"  ADVERTENCIA: No se pudo leer LISTADO DE MARCHAMOS: {e}")
    return lookup

# ─────────────────────────────────────────────
# Generación del Excel
# ─────────────────────────────────────────────

HEADERS = [
    "No.", "FILE", "PROVEEDOR", "BL", "CONTENEDOR", "TIPO", "DESTINO",
    "FACTURA", "ITEM", "FITO", "MANIFIESTO",
    "PILOTO", "LICENCIA", "PLACA", "FIANZA",
    "CODIGO DE TT", "TRANSPORTE ", "SELECTIVO", "DUCA",
    "RETENCION", "TRANSITO", "MARCHAMO"
]

# Siempre manuales (amarillo)
MANUAL_COLS = {
    "FILE", "PILOTO", "LICENCIA", "PLACA", "FIANZA",
    "CODIGO DE TT", "TRANSPORTE ", "SELECTIVO", "DUCA",
    "RETENCION", "TRANSITO",
}

YELLOW_FILL  = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
HEADER_FILL  = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
AUTO_FILL    = PatternFill(start_color="DEEAF1", end_color="DEEAF1", fill_type="solid")
MISSING_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
MARCH_FILL   = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
DEST_FILL    = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")  # naranja claro

HEADER_FONT  = Font(bold=True, color="FFFFFF", size=10)
DATA_FONT    = Font(size=9)
thin         = Side(style='thin', color='BBBBBB')
border       = Border(left=thin, right=thin, top=thin, bottom=thin)

COL_WIDTHS = {
    1: 5,   # No.
    2: 8,   # FILE
    3: 14,  # PROVEEDOR
    4: 16,  # BL
    5: 14,  # CONTENEDOR
    6: 7,   # TIPO
    7: 22,  # DESTINO
    8: 28,  # FACTURA
    9: 9,   # ITEM
    10: 36, # FITO
    11: 14, # MANIFIESTO
    12: 14, # PILOTO
    13: 12, # LICENCIA
    14: 10, # PLACA
    15: 10, # FIANZA
    16: 12, # CODIGO DE TT
    17: 12, # TRANSPORTE
    18: 10, # SELECTIVO
    19: 10, # DUCA
    20: 10, # RETENCION
    21: 10, # TRANSITO
    22: 16, # MARCHAMO
}


def generate_excel(rows, output_path, semana, marchamo_lookup=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"SEMANA {semana}"

    # Título
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
    tc = ws.cell(row=1, column=1, value=f"SEMANA {semana}")
    tc.font = Font(bold=True, size=14)
    tc.alignment = Alignment(horizontal='center')

    # Encabezados (fila 3)
    for ci, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=3, column=ci, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = border

    # Datos
    for ri, row_data in enumerate(rows, 4):
        cont = row_data.get("CONTENEDOR", "")
        for ci, h in enumerate(HEADERS, 1):
            cell = ws.cell(row=ri, column=ci)
            cell.border = border
            cell.font   = DATA_FONT
            cell.alignment = Alignment(wrap_text=True, vertical='top')

            if h == "No.":
                cell.value = ri - 3

            elif h in MANUAL_COLS:
                cell.value = None
                cell.fill  = YELLOW_FILL

            elif h == "MARCHAMO":
                if marchamo_lookup and cont in marchamo_lookup:
                    cell.value = marchamo_lookup[cont]
                    cell.fill  = MARCH_FILL
                else:
                    cell.value = None
                    cell.fill  = YELLOW_FILL

            elif h == "DESTINO":
                # Auto-llenado pero resaltado en naranja claro para que sea visible
                # como campo clave para asignación de pilotos
                val = row_data.get("DESTINO", "")
                cell.value = val if val != "???" else ""
                cell.fill  = DEST_FILL if val and val != "???" else MISSING_FILL

            else:
                val = row_data.get(h.strip(), "")
                if val == "???":
                    cell.value = ""
                    cell.fill  = MISSING_FILL
                else:
                    cell.value = val
                    cell.fill  = AUTO_FILL

    # Nota al pie
    fr = len(rows) + 5
    ws.merge_cells(start_row=fr, start_column=1, end_row=fr, end_column=len(HEADERS))
    nc = ws.cell(row=fr, column=1,
                 value="Azul=auto | Naranja=DESTINO (clave para asignación de piloto) | Amarillo=captura manual | Verde=marchamo auto | Rosado=no encontrado")
    nc.font = Font(italic=True, size=8, color="666666")

    # Anchos y alturas
    for ci, w in COL_WIDTHS.items():
        ws.column_dimensions[get_column_letter(ci)].width = w
    for r in range(4, len(rows) + 4):
        ws.row_dimensions[r].height = 30

    ws.freeze_panes = "C4"
    wb.save(output_path)

# ─────────────────────────────────────────────
# Función librería (usada por la API)
# ─────────────────────────────────────────────

def run_extraction(input_folder, output_path, semana="?", bl_file=None, marchamo_file=None):
    """
    Ejecuta el pipeline completo y escribe el Excel en output_path.
    Retorna: {"rows": N, "warnings": [str, ...]}
    Lanza ValueError si falta algo indispensable (sin sys.exit, para poder
    usarse dentro de un servicio web).
    """
    if not os.path.isdir(input_folder):
        raise ValueError(f"carpeta no encontrada: {input_folder}")

    marchamo_lookup = load_marchamo_lookup(marchamo_file)

    # ── Localizar el PDF combinado de BLs ──
    if not bl_file:
        candidates = []
        for root, _dirs, files in os.walk(input_folder):
            for fname in files:
                if classify_file(fname) == 'BL':
                    candidates.append(os.path.join(root, fname))
        if candidates:
            bl_file = newest_file(candidates)

    if not bl_file or not os.path.isfile(bl_file):
        raise ValueError("no se encontró el PDF combinado de BLs (pásalo explícito o colócalo en la carpeta)")

    pages_by_code = split_bl_pdf(bl_file)

    bl_by_container = {}  # CONTENEDOR -> (bl_data, warning)
    lead_warnings = []
    for code, text in pages_by_code.items():
        bl_data, bl_warning = extract_bl_from_text(text, code)
        cont = bl_data["CONTENEDOR"]
        if cont == "???":
            lead_warnings.append(f"No se pudo extraer el contenedor de {code}")
            continue
        bl_by_container[cont] = (bl_data, bl_warning)

    # ── Facturas / fitos / manifiestos (sin importar estructura de carpetas) ──
    facturas_raw, fitos_raw, manifiestos_raw = classify_and_extract_all(input_folder, exclude_path=bl_file)
    facturas    = dedup_by_key(facturas_raw, "factura")
    fitos       = dedup_by_key(fitos_raw, "number")
    manifiestos = dedup_by_key(manifiestos_raw, "number")

    # ── Universo de contenedores: BL ∪ facturas ∪ fitos ──
    all_containers = set(bl_by_container.keys())
    for f in facturas:
        if f.get("contenedor") not in (None, "???"):
            all_containers.add(f["contenedor"])
        all_containers.update(k for k in f.get("bultos_per_container", {}) if k)
    for fi in fitos:
        if fi.get("contenedor"):
            all_containers.add(fi["contenedor"])

    if not all_containers:
        raise ValueError("no se pudo determinar ningún contenedor a partir de los documentos")

    rows, all_warnings = [], list(lead_warnings)
    for cont in sorted(all_containers):
        bl_data, bl_warning = bl_by_container.get(cont, (None, None))
        try:
            row, warns = build_row(cont, bl_data, bl_warning, facturas, fitos, manifiestos)
            rows.append(row)
            all_warnings.extend(warns)
        except Exception as e:
            all_warnings.append(f"ERROR procesando {cont}: {e}")
            rows.append({"BL": "???", "PROVEEDOR": "???", "CONTENEDOR": cont,
                         "TIPO": "???", "DESTINO": "???", "FACTURA": "???",
                         "ITEM": "???", "FITO": "???", "MANIFIESTO": "???"})

    rows.sort(key=lambda r: (r.get("BL", "???"), r.get("CONTENEDOR", "")))

    generate_excel(rows, output_path, semana, marchamo_lookup)

    return {"rows": len(rows), "warnings": all_warnings}
