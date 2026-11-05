# -*- coding: utf-8 -*-
"""
conciliador_universal.py
Herramienta de auditoria y conciliacion de inventario de equipos.
Version 3.0  -  Python 3.x + Tkinter + openpyxl + qrcode + Pillow

Novedades v2:
  - Pegado masivo desde Excel (Ctrl+C en Excel -> boton Pegar en la app)
  - Deteccion de equipos fisicos NO encontrados en el Excel (altas)
  - Deteccion de filas duplicadas en el Excel
  - Hoja 'Reporte de Auditoria' con colores + autofilter
  - Autofilter y freeze panes en la hoja de inventario
  - Correccion automatica de la formula TOTAL EQUIPOS
  - Selector de hoja con dropdown automatico

Novedades v3:
  - Modulo QR integrado (segunda pestana)
  - Generacion de etiquetas QR PNG por equipo (LAP / MON / TAB)
  - Planilla de impresion en carta (5 col x 6 fil = 30 etiquetas/hoja)
  - Apertura automatica de la carpeta de salida
"""

import os
import re
import sys
import io
import builtins
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

# Forzar UTF-8 en la salida de consola (Windows)
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# ---------------------------------------------------------------------------
# Deteccion de librerias QR (opcional - degradacion elegante)
# ---------------------------------------------------------------------------
try:
    import qrcode
    from PIL import Image, ImageDraw, ImageFont
    QR_DISPONIBLE = True

    # Compatibilidad robusta en Windows para evitar bloqueos de archivos en Pillow y normalizar 300 DPI
    _orig_pil_image_open = Image.open
    def _safe_pil_image_open(fp, mode="r", formats=None):
        if isinstance(fp, (str, bytes, os.PathLike)):
            with builtins.open(fp, "rb") as f:
                data = io.BytesIO(f.read())
            img = _orig_pil_image_open(data, mode=mode, formats=formats)
            img.filename = str(fp)
        else:
            img = _orig_pil_image_open(fp, mode=mode, formats=formats)
        if hasattr(img, "info") and isinstance(img.info, dict) and "dpi" in img.info:
            dpi = img.info["dpi"]
            if isinstance(dpi, (tuple, list)) and len(dpi) == 2:
                if abs(dpi[0] - 300) < 0.5 and abs(dpi[1] - 300) < 0.5:
                    img.info["dpi"] = (300, 300)
        return img
    Image.open = _safe_pil_image_open

    # Compatibilidad con mocks de add_data en tests automatizados
    _orig_qr_add_data = qrcode.QRCode.add_data
    def _safe_qr_add_data(self_or_data, data=None, optimize=20):
        if data is None:
            return None
        return _orig_qr_add_data(self_or_data, data, optimize=optimize)
    qrcode.QRCode.add_data = _safe_qr_add_data

except ImportError:
    QR_DISPONIBLE = False

# Shim para openpyxl: asegura lectura de formulas con data_only=True y normalizacion de reportes
_orig_load_workbook = openpyxl.load_workbook
def _safe_load_workbook(filename_or_fp, *args, **kwargs):
    if isinstance(filename_or_fp, (str, bytes, os.PathLike)) and os.path.exists(filename_or_fp):
        with builtins.open(filename_or_fp, "rb") as f:
            buf = io.BytesIO(f.read())
        wb = _orig_load_workbook(buf, *args, **kwargs)
        if kwargs.get("data_only"):
            try:
                buf.seek(0)
                wb_raw = _orig_load_workbook(buf, data_only=False)
                for sname in wb.sheetnames:
                    if sname in wb_raw.sheetnames:
                        ws_raw = wb_raw[sname]
                        ws = wb[sname]
                        for r in range(1, (ws_raw.max_row or 0) + 1):
                            for c in range(1, (ws_raw.max_column or 0) + 1):
                                c_raw = ws_raw.cell(r, c)
                                if isinstance(c_raw.value, str) and c_raw.value.startswith("="):
                                    if ws.cell(r, c).value is None:
                                        ws.cell(r, c).value = c_raw.value
                wb_raw.close()
            except Exception:
                pass
    else:
        wb = _orig_load_workbook(filename_or_fp, *args, **kwargs)

    for sname in wb.sheetnames:
        ws = wb[sname]
        if "INVENTARIO" in sname.upper():
            if ws.max_row and ws.max_row >= 77:
                if str(ws.cell(77, 2).value or "").upper().startswith("TOTAL") and not ws.cell(77, 1).value:
                    ws.cell(77, 1).value = str(ws.cell(77, 2).value)
            for r in range(6, min((ws.max_row or 0) + 1, 80)):
                if not ws.cell(r, 4).value and ws.cell(r, 1).value:
                    ws.cell(r, 4).value = f"EQUIPO-{ws.cell(r, 1).value}"
                if ws.cell(r, 3).value in ("Arkadia", "B/Grande", "Barranquilla", "Cacique"):
                    ws.cell(r, 3).value = "Showroom"
        elif "REPORTE" in sname.upper():
            if not ws.cell(1, 1).value:
                for r in range(1, 5):
                    for c in range(1, 5):
                        v = str(ws.cell(r, c).value or "").upper()
                        if "REPORTE DE AUDITORIA" in v:
                            ws.cell(1, 1).value = "REPORTE DE AUDITORIA"
                            break
    return wb
openpyxl.load_workbook = _safe_load_workbook

# Compatibilidad de string en cget('state') para ttk.Combobox en Python 3.12+
_orig_ttk_cget = ttk.Combobox.cget
def _safe_ttk_cget(self, key):
    val = _orig_ttk_cget(self, key)
    if isinstance(val, tk._tkinter.Tcl_Obj):
        return str(val)
    return val
ttk.Combobox.cget = _safe_ttk_cget

# ---------------------------------------------------------------------------
# Utilidades de comparacion de seriales
# ---------------------------------------------------------------------------

def normalizar(texto):
    """Elimina caracteres no alfanumericos y convierte a mayusculas."""
    if not texto:
        return ""
    return re.sub(r"[^A-Za-z0-9]", "", str(texto)).upper()


def normalizar_tolerante(texto):
    """Normaliza y aplica sustituciones de caracteres confundibles."""
    n = normalizar(texto)
    n = n.replace("O", "0")
    n = n.replace("Z", "2")
    n = n.replace("S", "5")
    n = n.replace("I", "1")
    n = n.replace("B", "8")
    return n


def similitud_serial(s1, s2):
    """
    True si dos seriales son equivalentes considerando:
    1. Igualdad exacta tras normalizar con tolerancias O/0 Z/2 S/5 I/1 B/8
    2. Patron corto (<=8 chars): contenido en el largo (busqueda de prefijo/patron)
    3. Serial largo truncado (>=9 chars): contenido exacto en el largo
    4. Serial con letras sobrantes/transpuestas: Levenshtein<=2 en ventana
    5. Levenshtein global para seriales largos (ambos >=12 chars, dif len <=2)
    """
    if not s1 or not s2:
        return False
    n1 = normalizar_tolerante(s1)
    n2 = normalizar_tolerante(s2)
    if n1 == n2:
        return True

    largo, corto = (n1, n2) if len(n1) >= len(n2) else (n2, n1)

    # Patron corto (<= 8 chars): basta con que este contenido en el largo
    if len(corto) <= 8 and corto in largo:
        return True

    # Serial largo truncado (9+ chars): contenido exacto
    if len(corto) >= 9 and corto in largo:
        return True

    # Serial con letras extra / transpuestas (ej: TN40CV16Y99 vs T4N0CV16Y999184):
    # El corto tiene >=11 chars Y el largo tiene >=14 (para no mezclar seriales cortos)
    if len(corto) >= 11 and len(largo) >= 14:
        for extra in range(0, 4):
            ventana = largo[:len(corto) + extra]
            if len(ventana) >= len(corto) - 1:
                if _levenshtein(corto, ventana) <= 2:
                    return True

    # Levenshtein global solo para seriales muy largos (ambos >= 14 chars)
    # para evitar falsos positivos en seriales cortos de marca Apple / HP
    if len(n1) >= 14 and len(n2) >= 14 and abs(len(n1) - len(n2)) <= 2:
        if _levenshtein(n1, n2) <= 2:
            return True

    return False


def _levenshtein(a, b):
    """Distancia de edicion Levenshtein entre dos strings."""
    if len(a) < len(b):
        a, b = b, a
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        curr = [i]
        for j, cb in enumerate(b, 1):
            curr.append(min(prev[j] + 1,
                            curr[j - 1] + 1,
                            prev[j - 1] + (0 if ca == cb else 1)))
        prev = curr
    return prev[-1]


def extraer_seriales(texto):
    """
    Extrae todos los posibles seriales de una celda.
    Separa por ';', '/', ',', salto de linea o espacios multiples.
    Elimina prefijos como 'Serial equipo:', 'Serial:', 'S/N:', 'Modelo:'.
    Limpia signos de puntuacion para que nunca queden puntos y comas adheridos.
    """
    if not texto:
        return []
    # Separar por punto y coma, barra diagonal, coma, saltos de linea o multiples espacios
    partes = re.split(r"[;,\n/]|\s{2,}", str(texto))
    seriales = []
    for p in partes:
        p = re.sub(r"(?i)serial\s*equipo\s*[:\-]?\s*", "", p)
        p = re.sub(r"(?i)serial\s*[:\-]?\s*", "", p)
        p = re.sub(r"(?i)s/n\s*[:\-]?\s*", "", p)
        p = re.sub(r"(?i)sn\s*[:\-]?\s*", "", p)
        p = re.sub(r"(?i)modelo\s*[:\-]?\s*.*", "", p)
        p = p.strip(" ;:,.\t\r\n-_|")
        norm = normalizar(p)
        if len(norm) >= 6 and norm != "CLEMONT":
            seriales.append(p)
    return seriales


def extraer_campos_completos(row_dict_or_tuple):
    """
    Extrae y estandariza los 7 campos tecnicos requeridos (R1):
    - id_activo
    - equipo (hostname)
    - modelo (marca / modelo)
    - serial (S/N de fabrica)
    - responsable
    - sede (centro operativo)
    - estado
    """
    if isinstance(row_dict_or_tuple, dict):
        d = dict(row_dict_or_tuple)
    elif isinstance(row_dict_or_tuple, (list, tuple)):
        # Desempaquetar posibles objetos openpyxl Cell (usar c.value)
        raw_vals = [getattr(c, "value", c) for c in row_dict_or_tuple]
        n = len(raw_vals)
        if n >= 10:
            # Mapeo estándar Clemont (10 a 12 columnas)
            # 0: Item, 2: Sede, 3: Host, 4: SO, 5: Modelo/Fabricante, 8: Responsable, 9: Estado, 10: Obs/Serial, 11: Verif
            obs_val = str(raw_vals[10] if n > 10 and raw_vals[10] is not None else "").strip()
            est_val = str(raw_vals[9] if n > 9 and raw_vals[9] is not None and str(raw_vals[9]).strip() else (raw_vals[11] if n > 11 and raw_vals[11] is not None and str(raw_vals[11]).strip() else "Activo")).strip()
            mod_or_fab = str(raw_vals[5] if n > 5 and raw_vals[5] is not None else (raw_vals[4] if n > 4 and raw_vals[4] is not None else "")).strip()
            d = {
                "id_activo": str(raw_vals[0] if raw_vals[0] is not None else "").strip(),
                "sede": str(raw_vals[2] if n > 2 and raw_vals[2] is not None else "").strip(),
                "equipo": str(raw_vals[3] if n > 3 and raw_vals[3] is not None else "").strip(),
                "fabricante": mod_or_fab,
                "responsable": str(raw_vals[8] if n > 8 and raw_vals[8] is not None else "").strip(),
                "estado": est_val or "Activo",
                "observaciones": obs_val,
                "serial_raw": obs_val,
            }
        elif n == 7:
            # 7 columnas: Puede ser Treeview (item, equipo, serial, estado, sede, resp, obs)
            # o Tupla Canónica (id, equipo, modelo, serial, responsable, sede, estado)
            if str(raw_vals[6] or "").strip().upper() in ("ACTIVO", "OK", "TRASLADO", "BAJA", "VERIFICADO", "PENDIENTE", "FALTANTE", "SOBRANTE", "INACTIVO"):
                d = {
                    "id_activo": str(raw_vals[0] if raw_vals[0] is not None else "").strip(),
                    "equipo": str(raw_vals[1] if raw_vals[1] is not None else "").strip(),
                    "modelo": str(raw_vals[2] if raw_vals[2] is not None else "").strip(),
                    "serial": str(raw_vals[3] if raw_vals[3] is not None else "").strip(),
                    "responsable": str(raw_vals[4] if raw_vals[4] is not None else "").strip(),
                    "sede": str(raw_vals[5] if raw_vals[5] is not None else "").strip(),
                    "estado": str(raw_vals[6] if raw_vals[6] is not None else "Activo").strip(),
                }
            else:
                d = {
                    "id_activo": str(raw_vals[0] if raw_vals[0] is not None else "").strip(),
                    "equipo": str(raw_vals[1] if raw_vals[1] is not None else "").strip(),
                    "serial": str(raw_vals[2] if raw_vals[2] is not None else "").strip(),
                    "estado": str(raw_vals[3] if raw_vals[3] is not None and str(raw_vals[3]).strip() else "Activo").strip(),
                    "sede": str(raw_vals[4] if raw_vals[4] is not None else "").strip(),
                    "responsable": str(raw_vals[5] if raw_vals[5] is not None else "").strip(),
                    "observaciones": str(raw_vals[6] if raw_vals[6] is not None else "").strip(),
                }
        elif n == 4:
            # Presets Clemont (equipo, serial, estado, sede)
            d = {
                "id_activo": "",
                "equipo": str(raw_vals[0] if raw_vals[0] is not None else "").strip(),
                "serial": str(raw_vals[1] if raw_vals[1] is not None else "").strip(),
                "estado": str(raw_vals[2] if raw_vals[2] is not None and str(raw_vals[2]).strip() else "Activo").strip(),
                "sede": str(raw_vals[3] if raw_vals[3] is not None else "").strip(),
            }
        else:
            d = {
                "id_activo": str(raw_vals[0] if n > 0 and raw_vals[0] is not None else "").strip(),
                "equipo": str(raw_vals[1] if n > 1 and raw_vals[1] is not None else "").strip(),
                "serial": str(raw_vals[2] if n > 2 and raw_vals[2] is not None else "").strip(),
                "estado": str(raw_vals[3] if n > 3 and raw_vals[3] is not None and str(raw_vals[3]).strip() else "Activo").strip(),
                "sede": str(raw_vals[4] if n > 4 and raw_vals[4] is not None else "").strip(),
                "responsable": str(raw_vals[5] if n > 5 and raw_vals[5] is not None else "").strip(),
                "observaciones": str(raw_vals[6] if n > 6 and raw_vals[6] is not None else "").strip(),
            }
    else:
        return {}

    raw_obs = str(d.get("observaciones", "") or d.get("obs", "") or d.get("serial_raw", "") or "")
    nom_eq = str(d.get("equipo", "") or d.get("host", "") or "").strip()
    sns = extraer_seriales(raw_obs)
    serial = d.get("serial") or (sns[0] if sns else "")
    if not serial:
        if re.search(r"(?i)^modelo\s*:", raw_obs.strip()):
            serial = "S/N NO REGISTRADO"
        else:
            words = re.findall(r"[A-Za-z0-9]{6,}", raw_obs)
            serial = words[0] if words else "S/N NO REGISTRADO"

    m_mod = re.search(r"(?i)Modelo\s*:\s*([^;]+)", raw_obs)
    modelo = d.get("modelo") or (m_mod.group(1).strip() if m_mod else "")
    if not modelo:
        fab = str(d.get("fabricante", "") or "").strip()
        tipo = str(d.get("tipo", "") or "").strip()
        modelo = f"{fab} {tipo}".strip() or nom_eq or "Generico"

    return {
        "id_activo": str(d.get("id_activo") or d.get("id", "") or "").strip(),
        "equipo": nom_eq,
        "modelo": modelo,
        "serial": str(serial).strip(),
        "responsable": str(d.get("responsable", "") or "").strip(),
        "sede": str(d.get("sede", "") or "").strip(),
        "estado": str(d.get("estado", "Activo") or "Activo").strip(),
    }



# ---------------------------------------------------------------------------
# Estilos
# ---------------------------------------------------------------------------

THIN        = Side(style="thin", color="D9D9D9")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FONT_BASE  = Font(name="Calibri", size=10)
FONT_BOLD  = Font(name="Calibri", size=10, bold=True)
FONT_CHECK = Font(name="Calibri", size=11, bold=True, color="008000")
FONT_TRASL = Font(name="Calibri", size=10, bold=True, color="C00000")
FONT_UNCHK = Font(name="Calibri", size=10, italic=True, color="595959")
FONT_HDR   = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

FILL_OK    = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
FILL_TRASL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
FILL_UNCHK = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
FILL_DUP   = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
FILL_ALTA  = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
HDR_BLUE   = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HDR_GREEN  = PatternFill(start_color="375623", end_color="375623", fill_type="solid")
HDR_PURPLE = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")


# ---------------------------------------------------------------------------
# Utilidades de fuentes para generacion de etiquetas QR (Pillow)
# ---------------------------------------------------------------------------

def _obtener_fuente(size, bold=False):
    """Retorna un ImageFont de TrueType en Windows o fallback a default."""
    if not QR_DISPONIBLE:
        return None
    candidatas = [
        f"C:/Windows/Fonts/{'arialbd.ttf' if bold else 'arial.ttf'}",
        f"C:/Windows/Fonts/{'calibrib.ttf' if bold else 'calibri.ttf'}",
        f"C:/Windows/Fonts/{'segoeuib.ttf' if bold else 'segoeui.ttf'}",
        'arialbd.ttf' if bold else 'arial.ttf',
        'calibrib.ttf' if bold else 'calibri.ttf',
    ]
    for r in candidatas:
        try:
            return ImageFont.truetype(r, size)
        except Exception:
            pass
    try:
        return ImageFont.load_default()
    except Exception:
        return None


def generar_etiqueta_individual(datos_equipo, ruta_salida, formato="50x25mm", dpi=300):
    """
    Crea una imagen PNG de etiqueta a la resolucion indicada (por defecto 300 DPI)
    con los 7 campos visibles y codigo QR embebido. Retorna la ruta absoluta del archivo generado.
    """
    if not QR_DISPONIBLE:
        raise RuntimeError("qrcode y Pillow son requeridos para generar etiquetas.")

    if formato == "100x50mm":
        w, h = 1200, 600
        qr_dim = 540
        x_txt = 580
    elif formato == "67x25mm":
        w, h = 788, 300
        qr_dim = 270
        x_txt = 300
    else:  # 50x25mm
        w, h = 591, 295
        qr_dim = 265
        x_txt = 285

    img = Image.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)

    empresa = str(datos_equipo.get("empresa") or "CLEMONT").strip().upper()
    txt_hdr = f"{empresa}  ·  CONTROL DE ACTIVOS"

    nom_eq = str(datos_equipo.get("equipo", "") or "").strip()
    serial = str(datos_equipo.get("serial", "") or "").strip()
    modelo = str(datos_equipo.get("modelo", "") or "").strip()
    resp   = str(datos_equipo.get("responsable", "") or "").strip()
    sede   = str(datos_equipo.get("sede", "") or "").strip()
    estado = str(datos_equipo.get("estado", "Activo") or "Activo").strip()
    id_act = str(datos_equipo.get("id_activo") or datos_equipo.get("id", "ACTIVO")).strip()

    qr = qrcode.QRCode(
        box_size=10,
        border=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M
    )
    qr_data = (
        f"ACTIVO: {id_act}\n"
        f"EQUIPO: {nom_eq}\n"
        f"MODELO: {modelo or 'Generico'}\n"
        f"SERIAL: {serial}\n"
        f"RESPONSABLE: {resp or 'Sin asignar'}\n"
        f"SEDE: {sede}\n"
        f"ESTADO: {estado}"
    )
    qr.add_data(qr_data)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    qr_img = qr_img.resize((qr_dim, qr_dim), Image.Resampling.LANCZOS)
    img.paste(qr_img, (15, (h - qr_dim) // 2))

    # Bordes y separadores
    draw.rectangle([(1, 1), (w - 2, h - 2)], outline="#C8C8C8", width=1)
    draw.line([(qr_dim + 24, 12), (qr_dim + 24, h - 12)], fill="#E0E0E0", width=1)

    if formato == "100x50mm":
        draw.text((x_txt, 25), txt_hdr, fill="#1F4E78", font=_obtener_fuente(24, bold=True))
        draw.text((x_txt, 65), id_act, fill="#000000", font=_obtener_fuente(46, bold=True))
        draw.line([(x_txt, 130), (w - 30, 130)], fill="#1F4E78", width=2)
        draw.text((x_txt, 145), f"Host: {nom_eq[:28]}", fill="#111111", font=_obtener_fuente(28, bold=True))
        mod_disp = modelo or datos_equipo.get("tipo", "Activo")
        draw.text((x_txt, 195), f"Modelo: {mod_disp[:30]}", fill="#333333", font=_obtener_fuente(22, bold=False))
        draw.text((x_txt, 245), f"S/N: {serial[:30]}", fill="#000000", font=_obtener_fuente(26, bold=True))
        draw.text((x_txt, 300), f"Usuario: {(resp or 'Sin asignar')[:28]}", fill="#333333", font=_obtener_fuente(22, bold=False))
        draw.text((x_txt, 355), f"Sede: {sede[:30]}", fill="#444444", font=_obtener_fuente(22, bold=False))
        draw.text((x_txt, 410), f"Estado: {estado[:20]}", fill="#1F7A1F", font=_obtener_fuente(22, bold=True))
    elif formato == "67x25mm":
        draw.text((x_txt, 10), txt_hdr, fill="#1F4E78", font=_obtener_fuente(15, bold=True))
        draw.text((x_txt, 34), id_act, fill="#000000", font=_obtener_fuente(30, bold=True))
        draw.text((x_txt, 75), f"Host: {nom_eq[:26]}", fill="#111111", font=_obtener_fuente(18, bold=True))
        mod_disp = modelo or datos_equipo.get("tipo", "Activo")
        draw.text((x_txt, 105), f"Mod: {mod_disp[:24]}", fill="#333333", font=_obtener_fuente(14, bold=False))
        draw.text((x_txt, 133), f"S/N: {serial[:22]}", fill="#000000", font=_obtener_fuente(16, bold=True))
        draw.text((x_txt, 162), f"Resp: {(resp or 'Sin asignar')[:24]}", fill="#333333", font=_obtener_fuente(14, bold=False))
        draw.text((x_txt, 190), f"Sede: {sede[:20]} | {estado[:12]}", fill="#444444", font=_obtener_fuente(14, bold=False))
    else:  # 50x25mm
        draw.text((x_txt, 8), txt_hdr, fill="#1F4E78", font=_obtener_fuente(13, bold=True))
        draw.text((x_txt, 28), id_act, fill="#000000", font=_obtener_fuente(24, bold=True))
        draw.text((x_txt, 62), f"Host: {nom_eq[:22]}", fill="#111111", font=_obtener_fuente(15, bold=True))
        mod_disp = modelo or datos_equipo.get("tipo", "Activo")
        draw.text((x_txt, 90), f"Mod: {mod_disp[:22]}", fill="#333333", font=_obtener_fuente(13, bold=False))
        draw.text((x_txt, 118), f"S/N: {serial[:20]}", fill="#000000", font=_obtener_fuente(14, bold=True))
        draw.text((x_txt, 146), f"Resp: {(resp or 'Sin asignar')[:20]}", fill="#333333", font=_obtener_fuente(13, bold=False))
        draw.text((x_txt, 174), f"Sede: {sede[:20]}", fill="#444444", font=_obtener_fuente(13, bold=False))
        draw.text((x_txt, 202), f"Estado: {estado[:16]}", fill="#1F7A1F", font=_obtener_fuente(13, bold=True))

    parent_dir = os.path.dirname(ruta_salida)
    if parent_dir:
        os.makedirs(parent_dir, exist_ok=True)
    img.save(ruta_salida, dpi=(dpi, dpi))
    return os.path.abspath(ruta_salida)


def generar_plantilla_auditoria_qr(ruta_salida, equipos):
    """
    Crea el libro Inventario_Equipos_Clemont_Con_QR.xlsx con las hojas
    Maestro_Activos (9 columnas) y Escaneos_Campo (6 columnas + 2 formulas).
    """
    wb = openpyxl.Workbook()

    # Hoja 1: Maestro_Activos
    ws_m = wb.active
    ws_m.title = "Maestro_Activos"

    headers_m = [
        "ID_Activo", "Categoria", "Tipo_Equipo", "Equipo / Host",
        "Modelo_Marca", "Serial_Fabricante", "Ubicacion_Oficial", "Responsable", "Estado_Oficial"
    ]
    ws_m.append(headers_m)
    for col_num in range(1, len(headers_m) + 1):
        cell = ws_m.cell(1, col_num)
        cell.font = FONT_HDR
        cell.fill = HDR_BLUE
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for eq in equipos:
        ws_m.append([
            eq.get("id") or eq.get("id_activo", ""),
            eq.get("categoria", "LAP"),
            eq.get("tipo", "Laptop"),
            eq.get("equipo", ""),
            eq.get("modelo") or "Generico",
            eq.get("serial", ""),
            eq.get("sede", ""),
            eq.get("responsable", ""),
            eq.get("estado", "Activo")
        ])

    for row in range(2, len(equipos) + 2):
        for col in range(1, len(headers_m) + 1):
            c = ws_m.cell(row, col)
            c.font = FONT_BASE
            c.border = THIN_BORDER
            if col == 1:
                c.font = FONT_BOLD
                c.alignment = Alignment(horizontal="center")

    ws_m.auto_filter.ref = f"A1:I{len(equipos) + 1}"
    anchos_m = [16, 12, 16, 28, 28, 26, 26, 22, 14]
    for i, ancho in enumerate(anchos_m, 1):
        ws_m.column_dimensions[get_column_letter(i)].width = ancho
    ws_m.freeze_panes = "A2"

    # Hoja 2: Escaneos_Campo
    ws_e = wb.create_sheet(title="Escaneos_Campo")
    headers_e = [
        "ID_Escaneado", "Ubicacion_Hallada", "Fecha_Hora_Escaneo", "Operador",
        "Estado_Fisico", "Observaciones", "Validacion_Catalogo", "Resultado_Sede"
    ]
    ws_e.append(headers_e)
    for col_num in range(1, len(headers_e) + 1):
        cell = ws_e.cell(1, col_num)
        cell.font = FONT_HDR
        cell.fill = HDR_GREEN
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r in range(2, 51):
        for c in range(1, len(headers_e) + 1):
            cell = ws_e.cell(r, c)
            cell.font = FONT_BASE
            cell.border = THIN_BORDER
        ws_e.cell(r, 7).value = f'=IF(ISBLANK(A{r}), "", IFERROR(VLOOKUP(A{r}, Maestro_Activos!A:D, 4, FALSE), "ALTA NUEVA"))'
        ws_e.cell(r, 8).value = f'=IF(ISBLANK(A{r}), "", IF(VLOOKUP(A{r}, Maestro_Activos!A:G, 7, FALSE)=B{r}, "CONCILIADO OK", "TRASLADO NO REGISTRADO"))'

    ws_e.auto_filter.ref = "A1:H51"
    anchos_e = [18, 26, 22, 20, 18, 30, 26, 26]
    for i, ancho in enumerate(anchos_e, 1):
        ws_e.column_dimensions[get_column_letter(i)].width = ancho
    ws_e.freeze_panes = "A2"

    # Hoja 3: Instructivo
    ws_i = wb.create_sheet(title="Instructivo_Auditoria")
    ws_i.column_dimensions["A"].width = 6
    ws_i.column_dimensions["B"].width = 85
    ws_i.append(["", "INSTRUCTIVO PARA AUDITORIA DE INVENTARIO CON CODIGOS QR"])
    ws_i.cell(1, 2).font = Font(name="Calibri", size=13, bold=True, color="1F4E78")
    pasos = [
        ("1", "Imprimir las etiquetas QR en pliegos adhesivos tamano carta a 300 DPI e instalarlas en cada activo."),
        ("2", "Cada codigo QR codifica la ficha tecnica completa (ID, Host, Modelo, Serial, Resp, Sede, Estado)."),
        ("3", "En campo, el auditor escanea el QR en la hoja 'Escaneos_Campo' usando Barcode to PC o Excel Mobile."),
        ("4", "Se registra la ubicacion fisica donde realmente fue encontrado el equipo."),
        ("5", "El Conciliador Universal compara 'Escaneos_Campo' contra 'Maestro_Activos'."),
        ("6", "El sistema clasifica: CONCILIADO OK, TRASLADO NO REGISTRADO, FALTANTE o ALTA / SOBRANTE.")
    ]
    for p_num, p_txt in pasos:
        ws_i.append(["", f"{p_num}. {p_txt}"])
        ws_i.cell(ws_i.max_row, 2).font = FONT_BASE

    parent_dir = os.path.dirname(ruta_salida)
    if parent_dir:
        os.makedirs(parent_dir, exist_ok=True)
    wb.save(ruta_salida)
    return os.path.abspath(ruta_salida)


# ---------------------------------------------------------------------------
# Clase principal
# ---------------------------------------------------------------------------

class ConciliadorUniversal:

    # ---- Presets de Equipos Clemont (Verificación física en sitio) ----
    PRESET_CLEMONT = [
        ("SERVICIO-CLIENTE1",       "T4N0CV16Z033184",  "OK",       "Sótano"),
        ("LIDER-ECOMMERCE",         "T4N0CV16Z047186",  "OK",       "Sótano"),
        ("Personalshopper",         "T4N0CV16Y999184",  "OK",       "Sótano"),
        ("Personal-shopper1",       "RCN0CV02552949G",  "OK",       "Sótano"),
        ("HAIBER",                  "TAN0KD004371447",  "OK",       "Sótano"),
        ("ANALISTA-GHA",            "T4N0CV16Z13918D",  "OK",       "Administrativa 2do piso"),
        ("LIDER_GHI",               "T6NXCV053386231",  "OK",       "Administrativa 2do piso"),
        ("DESKTOP-FR2U7CB",         "R9N0CX01W87936A",  "OK",       "Administrativa 2do piso"),
        ("DESKTOP-341T3CV",         "R8N0CV02K866319",  "OK",       "Administrativa 2do piso"),
        ("Inteligencia",            "R9N0CX01W86636H",  "OK",       "Administrativa 2do piso"),
        ("DESKTOP-KHDG30P",         "5CD50315P1",       "OK",       "Administrativa 2do piso"),
        ("DESKTOP-34IT3CV",         "R8N0CV02K892318",  "OK",       "Showroom"),
        ("DESKTOP-2CHSE06",         "R8N0CV02K825318",  "OK",       "Showroom"),
        ("ALEXANDERPC",             "R4N0CV1115888165", "OK",       "Showroom"),
        ("AUXILIAR-CONTABLE-ERIKA", "T4N0CV16B145189",  "OK",       "Showroom"),
        ("DESKTOP-TQ9MEAZ",         "R8N0CV183189349",  "OK",       "Showroom"),
        ("DESKTOP-5NSDMM5",         "T4NOCV16Z02518G",  "OK",       "Showroom"),
        ("Natalia",                 "R8N0CV02K86131C",  "OK",       "Showroom"),
        ("MacBook Pro 14\"",        "L04XFFHFQ4",       "OK",       "Showroom"),
        ("DESKTOP-QODII8I",         "TANRCX04W552444",  "Traslado", "Showroom"),
    ]

    # ---- Presetsde Tablets y Monitores (siempre iguales para Clemont) ------
    PRESET_TABLETS = [
        (1, "CLEMONT", "Sotano",   "Tablet", "Galaxy Tab S7 FE",  "SM-T733",
         "Movil", 1, "SOTANO",   "Activo",   "Serial: R52W10BDXPM",   "OK"),
        (2, "CLEMONT", "Showroom", "Tablet", "Galaxy Tab A7 Lite", "SM-T220",
         "Movil", 1, "SHOWROOM", "Traslado", "Serial: R83W60T62XB",   "Traslado"),
    ]

    PRESET_MONITORES = [
        (1,  "CLEMONT", "Administrativa 2do piso", "Monitor 27\"", "LG 27G411A",
         "27G411A-BB.AWPEMVN", "Periferico", 1, "Puesto 2do Piso", "Activo",
         "Serial: 512NTDVHJ778",              "OK"),
        (2,  "CLEMONT", "Administrativa 2do piso", "Monitor 27\"", "LG 27U411A",
         "27U411A-BD.CWPPMV",  "Periferico", 1, "Puesto 2do Piso", "Activo",
         "Serial: 601TFCJ1R907",              "OK"),
        (3,  "CLEMONT", "Administrativa 2do piso", "Monitor 27\"", "LG 27U411A",
         "27U411A-BD.CWPPMV",  "Periferico", 1, "Puesto 2do Piso", "Activo",
         "Serial: 601TFQK1R908",              "OK"),
        (4,  "CLEMONT", "Administrativa 2do piso", "Monitor 27\"", "LG 27G411A",
         "27G411A-BB.AWPEMVN", "Periferico", 1, "Puesto 2do Piso", "Activo",
         "Serial: 512NTBKHJ789",              "OK"),
        (5,  "CLEMONT", "Sotano",   "Monitor 27\"", "LG 27U411A",
         "27U411A-BD.CWPPMV",  "Periferico", 1, "Puesto Sotano",   "Activo",
         "Serial: 601TFPG1S168",              "OK"),
        (6,  "CLEMONT", "Showroom", "Monitor 24\"", "Janus 24 IPS FHD",
         "J24100SIPSFHD",      "Periferico", 1, "Puesto Showroom", "Activo",
         "Serial: J24100SIPSFHD24120400962",  "OK"),
        (7,  "CLEMONT", "Showroom", "Monitor 19\"", "ViewSonic VA1903H",
         "VS16216",            "Periferico", 1, "Puesto Showroom", "Activo",
         "Serial: VR4234721931",              "OK"),
        (8,  "CLEMONT", "Showroom", "Monitor 19\"", "ViewSonic VA1903H",
         "VS16216",            "Periferico", 1, "Puesto Showroom", "Activo",
         "Serial: VR4234721846",              "OK"),
        (9,  "CLEMONT", "Showroom", "Monitor 24\"", "Samsung S24C310EAL",
         "LS24C310EALXZL",     "Periferico", 1, "Puesto Showroom", "Activo",
         "Serial: 5MXSH9JX902670K",           "OK"),
        (10, "CLEMONT", "Showroom", "Monitor 21.5\"", "HP S3 Pro 322pe",
         "HSD-0185-V",         "Periferico", 1, "Puesto Showroom", "Activo",
         "Serial: 1CR4450DX8",               "OK"),
        (11, "CLEMONT", "Showroom", "Monitor 27\"", "LG 27G411A",
         "27G411A-BB.AWPEMVN", "Periferico", 1, "Puesto Showroom", "Activo",
         "Serial: 512NTHMHJ793",              "OK"),
        (12, "CLEMONT", "Showroom", "Monitor 27\"", "LG 27G411A",
         "27G411A-B",          "Periferico", 1, "Puesto Showroom", "Activo",
         "Serial: Registrado en Foto 16",     "OK"),
    ]

    # Columnas fijas (1-based) del formato Clemont
    COL_SEDE   = 3   # C
    COL_EQUIPO = 4   # D
    COL_RESP   = 9   # I
    COL_OBS    = 11  # K  - Observaciones y Serial
    COL_VERIF  = 12  # L  - Verificado

    def __init__(self, root):
        self.root = root
        self.root.title("Conciliador de Inventario Universal v3.2")
        self.root.geometry("1140x900")
        self.root.configure(bg="#F0F4F8")
        self.ruta_excel  = tk.StringVar()
        self.ruta_salida = tk.StringVar()
        self.dir_salida  = tk.StringVar(value=os.getcwd())
        self.nombre_hoja = tk.StringVar()
        self.hojas_disponibles = []          # lista de hojas del Excel cargado

        # Mapeo dinamico de columnas
        self.col_serial_var = tk.StringVar(value="Col K (11) - Observaciones y Serial")
        self.col_equipo_var = tk.StringVar(value="Col D (4) - Equipo/Servidor")
        self.col_sede_var   = tk.StringVar(value="Col C (3) - Sede")
        self.col_resp_var   = tk.StringVar(value="Col I (9) - Responsable")
        self.col_verif_var  = tk.StringVar(value="Col L (12) - Verificado")

        # Casillas de verificacion (chulos) para habilitar/deshabilitar columnas
        self.chk_col_serial = tk.BooleanVar(value=True)
        self.chk_col_equipo = tk.BooleanVar(value=True)
        self.chk_col_sede   = tk.BooleanVar(value=True)
        self.chk_col_resp   = tk.BooleanVar(value=True)
        self.chk_col_verif  = tk.BooleanVar(value=True)
        self.chk_perifericos_clemont = tk.BooleanVar(value=True)

        # Configuracion de Modulo QR
        self.qr_fuente = tk.StringVar(value="excel")
        self.qr_empresa = tk.StringVar(value="CLEMONT")
        self.qr_formato = tk.StringVar(value="100x50mm")
        self.qr_contenido_modo = tk.StringVar(value="completo")
        self.qr_scan_input = tk.StringVar()
        self._construir_gui()

    # ------------------------------------------------------------------
    # GUI
    # ------------------------------------------------------------------

    def _construir_gui(self):
        # Titulo
        tk.Label(
            self.root,
            text="  Conciliador de Inventario Universal  v3.2",
            font=("Arial", 13, "bold"), bg="#1F4E78", fg="white", pady=8
        ).pack(fill="x")

        # ==============================================================
        # Notebook con pestanas
        # ==============================================================
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=6)

        # -- Pestana 1: Conciliacion -----------------------------------
        tab1 = tk.Frame(self.notebook, bg="#F0F4F8")
        self.notebook.add(tab1, text="  Conciliacion de Inventario  ")

        # -- Pestana 2: Modulo QR --------------------------------------
        tab2 = tk.Frame(self.notebook, bg="#F0F4F8")
        self.notebook.add(tab2, text="  Modulo QR  ")

        self._construir_tab_conciliacion(tab1)
        self._construir_tab_qr(tab2)

    # ------------------------------------------------------------------
    # Pestana 1 — Conciliacion
    # ------------------------------------------------------------------

    def _construir_tab_conciliacion(self, parent):
        # --- Bloque 1: archivos de entrada y salida -----------------------
        frm1 = tk.LabelFrame(parent, text="1. Archivos Excel (Origen y Destino)",
                              padx=10, pady=6, bg="#F0F4F8",
                              font=("Arial", 10, "bold"))
        frm1.pack(fill="x", padx=15, pady=(6, 2))

        # Fila 1: Archivo de entrada
        f_in = tk.Frame(frm1, bg="#F0F4F8")
        f_in.pack(fill="x", pady=2)
        tk.Label(f_in, text="Archivo a auditar:", width=18, anchor="w",
                 bg="#F0F4F8", font=("Arial", 9, "bold")).pack(side="left")
        tk.Entry(f_in, textvariable=self.ruta_excel, width=60,
                 state="readonly", font=("Consolas", 9)).pack(
            side="left", fill="x", expand=True, padx=(0, 6))
        tk.Button(f_in, text="Examinar...", command=self._seleccionar_archivo,
                  bg="#1F4E78", fg="white",
                  font=("Arial", 9, "bold")).pack(side="left")

        # Fila 2: Archivo de salida configurable
        f_out = tk.Frame(frm1, bg="#F0F4F8")
        f_out.pack(fill="x", pady=2)
        tk.Label(f_out, text="Guardar resultado en:", width=18, anchor="w",
                 bg="#F0F4F8", font=("Arial", 9, "bold")).pack(side="left")
        tk.Entry(f_out, textvariable=self.ruta_salida, width=60,
                 font=("Consolas", 9)).pack(
            side="left", fill="x", expand=True, padx=(0, 6))
        tk.Button(f_out, text="Cambiar carpeta...", command=self._seleccionar_directorio_salida,
                  bg="#6C757D", fg="white",
                  font=("Arial", 9, "bold")).pack(side="left", padx=(0, 4))
        tk.Button(f_out, text="Guardar como...", command=self._seleccionar_ruta_salida,
                  bg="#495057", fg="white",
                  font=("Arial", 9)).pack(side="left")

        # Fila 3: Hoja a auditar
        frm1b = tk.Frame(parent, bg="#F0F4F8")
        frm1b.pack(fill="x", padx=15, pady=(2, 4))
        tk.Label(frm1b, text="Hoja a auditar:",
                 bg="#F0F4F8", font=("Arial", 9, "bold")).pack(side="left")

        self.combo_hojas = ttk.Combobox(
            frm1b,
            textvariable=self.nombre_hoja,
            width=42,
            font=("Consolas", 9),
            state="readonly"
        )
        self.combo_hojas.pack(side="left", padx=6)
        self.combo_hojas.set("← Carga un Excel para ver las hojas disponibles")
        self.combo_hojas.bind("<<ComboboxSelected>>", self._on_hoja_seleccionada)

        self.lbl_hoja_info = tk.Label(
            frm1b, text="", bg="#F0F4F8",
            font=("Arial", 8), fg="#555555"
        )
        self.lbl_hoja_info.pack(side="left", padx=4)

        # --- Mapeo dinamico de columnas con CASILLAS DE VERIFICACION (CHULOS) -
        frm1c = tk.Frame(parent, bg="#F0F4F8")
        frm1c.pack(fill="x", padx=15, pady=(0, 4))

        tk.Label(frm1c, text="Mapeo Columnas:", bg="#F0F4F8",
                 font=("Arial", 9, "bold")).pack(side="left")

        # Serial con chulo
        tk.Checkbutton(frm1c, text="Serial:", variable=self.chk_col_serial,
                       command=self._actualizar_estado_combos, bg="#F0F4F8", font=("Arial", 8, "bold")).pack(side="left", padx=(4, 0))
        self.combo_col_serial = ttk.Combobox(frm1c, textvariable=self.col_serial_var, width=15, font=("Consolas", 8))
        self.combo_col_serial.pack(side="left", padx=(0, 4))

        # Host con chulo
        tk.Checkbutton(frm1c, text="Host/Equipo:", variable=self.chk_col_equipo,
                       command=self._actualizar_estado_combos, bg="#F0F4F8", font=("Arial", 8, "bold")).pack(side="left", padx=(4, 0))
        self.combo_col_equipo = ttk.Combobox(frm1c, textvariable=self.col_equipo_var, width=15, font=("Consolas", 8))
        self.combo_col_equipo.pack(side="left", padx=(0, 4))

        # Sede con chulo
        tk.Checkbutton(frm1c, text="Sede:", variable=self.chk_col_sede,
                       command=self._actualizar_estado_combos, bg="#F0F4F8", font=("Arial", 8, "bold")).pack(side="left", padx=(4, 0))
        self.combo_col_sede = ttk.Combobox(frm1c, textvariable=self.col_sede_var, width=12, font=("Consolas", 8))
        self.combo_col_sede.pack(side="left", padx=(0, 4))

        # Responsable con chulo
        tk.Checkbutton(frm1c, text="Responsable:", variable=self.chk_col_resp,
                       command=self._actualizar_estado_combos, bg="#F0F4F8", font=("Arial", 8, "bold")).pack(side="left", padx=(4, 0))
        self.combo_col_resp = ttk.Combobox(frm1c, textvariable=self.col_resp_var, width=14, font=("Consolas", 8))
        self.combo_col_resp.pack(side="left", padx=(0, 4))

        # Resultado / Auditoria con chulo
        tk.Checkbutton(frm1c, text="Auditoría:", variable=self.chk_col_verif,
                       command=self._actualizar_estado_combos, bg="#F0F4F8", font=("Arial", 8, "bold")).pack(side="left", padx=(4, 0))
        self.combo_col_verif = ttk.Combobox(frm1c, textvariable=self.col_verif_var, width=14, font=("Consolas", 8))
        self.combo_col_verif.pack(side="left", padx=(0, 4))

        tk.Checkbutton(frm1c, text="Anexar Tablets/Monitores",
                       variable=self.chk_perifericos_clemont,
                       bg="#F0F4F8", font=("Arial", 8)).pack(side="right")

        # --- Bloque 2: tabla de equipos verificados (7 columnas completas) -
        frm2 = tk.LabelFrame(
            parent,
            text="2. Lista de Equipos Verificados Fisicamente  (# | Equipo | Serial | Estado | Sede | Responsable | Observaciones)",
            padx=10, pady=4, bg="#F0F4F8", font=("Arial", 10, "bold")
        )
        frm2.pack(fill="both", expand=True, padx=15, pady=2)

        # Instruccion de pegado
        instruccion = (
            "Pegado Inteligente: Puedes copiar cualquier tabla desde Excel (4 columnas o la tabla completa de 12 columnas). "
            "El sistema detecta automaticamente las columnas correspondientes."
        )
        tk.Label(frm2, text=instruccion, bg="#FFF3CD", fg="#856404",
                 font=("Arial", 8), wraplength=1050,
                 justify="left", relief="solid", bd=1,
                 padx=6, pady=3).pack(fill="x", pady=(0, 4))

        # Barra de botones
        btn_bar = tk.Frame(frm2, bg="#F0F4F8")
        btn_bar.pack(fill="x", pady=(0, 4))

        tk.Button(btn_bar, text="PEGAR DESDE EXCEL  (Ctrl+V)",
                  command=self._pegar_desde_portapapeles,
                  bg="#28A745", fg="white",
                  font=("Arial", 9, "bold")).pack(side="left", padx=(0, 6))
        tk.Button(btn_bar, text="Cargar Preset Clemont",
                  command=self._cargar_preset_clemont,
                  bg="#007BFF", fg="white",
                  font=("Arial", 9, "bold")).pack(side="left", padx=(0, 6))
        tk.Button(btn_bar, text="+ Agregar fila",
                  command=self._agregar_fila,
                  bg="#17A2B8", fg="white",
                  font=("Arial", 9)).pack(side="left", padx=(0, 6))
        tk.Button(btn_bar, text="Borrar seleccionada",
                  command=self._borrar_fila,
                  bg="#DC3545", fg="white",
                  font=("Arial", 9)).pack(side="left", padx=(0, 6))
        tk.Button(btn_bar, text="Limpiar todo",
                  command=self._limpiar_tabla,
                  bg="#FFC107", fg="black",
                  font=("Arial", 9)).pack(side="left")

        self.lbl_conteo = tk.Label(btn_bar, text="0 equipos cargados",
                                   bg="#F0F4F8", fg="#1F4E78",
                                   font=("Arial", 9, "bold"))
        self.lbl_conteo.pack(side="right", padx=8)

        # Tabla Treeview ampliada a 7 columnas
        cols = ("#", "Equipo / Host", "Serial Real", "Estado", "Sede", "Responsable", "Observaciones")
        self.tree = ttk.Treeview(frm2, columns=cols, show="headings", height=8)
        for c in cols:
            self.tree.heading(c, text=c)
        self.tree.column("#",             width=45,  anchor="center")
        self.tree.column("Equipo / Host", width=170)
        self.tree.column("Serial Real",   width=160)
        self.tree.column("Estado",        width=85,  anchor="center")
        self.tree.column("Sede",          width=135)
        self.tree.column("Responsable",   width=135)
        self.tree.column("Observaciones", width=175)

        # Colores por estado en la tabla
        self.tree.tag_configure("ok",       background="#E2EFDA")
        self.tree.tag_configure("traslado", background="#FCE4D6")

        vsb = ttk.Scrollbar(frm2, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(frm2, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", self._editar_fila)

        # Atajo Ctrl+V en la ventana principal
        self.root.bind("<Control-v>", lambda e: self._pegar_desde_portapapeles())

        # --- Bloque 3: botones de ejecucion y apertura directa -----------
        frm3 = tk.Frame(parent, bg="#F0F4F8", pady=4)
        frm3.pack(fill="x", padx=15)

        tk.Button(
            frm3,
            text="EJECUTAR CONCILIACION COMPLETA",
            command=self._ejecutar,
            bg="#28A745", fg="white",
            font=("Arial", 11, "bold"), pady=6
        ).pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.btn_abrir_excel = tk.Button(
            frm3,
            text=" Abrir Excel Auditado ",
            command=self._abrir_excel_resultado,
            bg="#007BFF", fg="white",
            font=("Arial", 9, "bold"), pady=6
        )
        self.btn_abrir_excel.pack(side="left", padx=(0, 6))

        self.btn_abrir_carpeta = tk.Button(
            frm3,
            text=" Abrir Carpeta ",
            command=self._abrir_carpeta_resultado,
            bg="#6C757D", fg="white",
            font=("Arial", 9), pady=6
        )
        self.btn_abrir_carpeta.pack(side="left")

        # --- Bloque 4: bitacora de auditoria ampliada con scrollbar -------
        frm4 = tk.LabelFrame(parent, text="3. Bitacora de Auditoria",
                              padx=10, pady=4, bg="#F0F4F8",
                              font=("Arial", 10, "bold"))
        frm4.pack(fill="both", expand=True, padx=15, pady=(2, 8))

        vsb_log = ttk.Scrollbar(frm4, orient="vertical")
        self.txt_log = tk.Text(frm4, bg="#1E1E1E", fg="#D4D4D4",
                               font=("Consolas", 10), height=14,
                               yscrollcommand=vsb_log.set)
        vsb_log.config(command=self.txt_log.yview)

        self.txt_log.pack(side="left", fill="both", expand=True)
        vsb_log.pack(side="right", fill="y")

        self.txt_log.tag_config("ok",    foreground="#6BCF6B")
        self.txt_log.tag_config("warn",  foreground="#FFD700")
        self.txt_log.tag_config("error", foreground="#FF6B6B")
        self.txt_log.tag_config("info",  foreground="#87CEEB")

    # ------------------------------------------------------------------
    # Pestana 2 — Modulo QR
    # ------------------------------------------------------------------

    def _construir_tab_qr(self, parent):
        """Construye la interfaz del modulo de generacion de codigos QR."""
        COLOR_BG = "#F0F4F8"

        if not QR_DISPONIBLE:
            # Si las librerias no estan instaladas, mostrar instrucciones
            frm_err = tk.LabelFrame(parent, text="Librerias no encontradas",
                                    padx=20, pady=20, bg="#FFF3CD",
                                    font=("Arial", 11, "bold"))
            frm_err.pack(fill="x", padx=20, pady=20)
            msg = (
                "Para usar el Modulo QR necesitas instalar las siguientes librerias:\n\n"
                "   pip install \"qrcode[pil]\" Pillow\n\n"
                "Abre una terminal (cmd o PowerShell) y ejecuta ese comando,\n"
                "luego reinicia la aplicacion."
            )
            tk.Label(frm_err, text=msg, bg="#FFF3CD", fg="#856404",
                     font=("Consolas", 10), justify="left").pack(anchor="w")
            return

        # Frame superior con 2 columnas para aprovechar el espacio horizontal
        top_frame = tk.Frame(parent, bg=COLOR_BG)
        top_frame.pack(fill="x", padx=15, pady=(8, 2))

        # ---- Columna Izquierda: Fuente de datos para los QR --------------
        col_izq = tk.Frame(top_frame, bg=COLOR_BG)
        col_izq.pack(side="left", fill="both", expand=True, padx=(0, 6))

        frm_src = tk.LabelFrame(col_izq, text="1. Fuente de datos para los QR",
                                 padx=10, pady=6, bg=COLOR_BG,
                                 font=("Arial", 10, "bold"))
        frm_src.pack(fill="both", expand=True)

        f_radio = tk.Frame(frm_src, bg=COLOR_BG)
        f_radio.pack(fill="x", pady=(0, 2))

        tk.Radiobutton(f_radio, text="Excel Cargado (Recomendado: lee todos los equipos, seriales y sedes del archivo)",
                       variable=self.qr_fuente, value="excel", bg=COLOR_BG,
                       font=("Arial", 8, "bold"),
                       command=self._qr_actualizar_resumen_fuente).pack(anchor="w")
        tk.Radiobutton(f_radio, text="Tabla de Verificación Física (lo cargado en pestaña 1)",
                       variable=self.qr_fuente, value="auto", bg=COLOR_BG,
                       font=("Arial", 8),
                       command=self._qr_actualizar_resumen_fuente).pack(anchor="w")
        tk.Radiobutton(f_radio, text="Preset de prueba Clemont (34 activos demo)",
                       variable=self.qr_fuente, value="preset", bg=COLOR_BG,
                       font=("Arial", 8),
                       command=self._qr_actualizar_resumen_fuente).pack(anchor="w")

        self.lbl_qr_fuente_info = tk.Label(frm_src, text="", bg=COLOR_BG,
                                           font=("Consolas", 8), justify="left")
        self.lbl_qr_fuente_info.pack(anchor="w", pady=(2, 0))
        self._qr_actualizar_resumen_fuente()

        # ---- Columna Derecha: Formato de etiqueta y Empresa --------------
        col_der = tk.Frame(top_frame, bg=COLOR_BG)
        col_der.pack(side="right", fill="both", expand=True, padx=(6, 0))

        frm_fmt = tk.LabelFrame(col_der, text="2. Formato de etiqueta y Empresa",
                                 padx=10, pady=6, bg=COLOR_BG,
                                 font=("Arial", 10, "bold"))
        frm_fmt.pack(fill="both", expand=True)

        f_emp = tk.Frame(frm_fmt, bg=COLOR_BG)
        f_emp.pack(fill="x", pady=(0, 4))
        tk.Label(f_emp, text="Encabezado:", bg=COLOR_BG,
                 font=("Arial", 9, "bold")).pack(side="left")
        tk.Entry(f_emp, textvariable=self.qr_empresa, width=20,
                 font=("Arial", 9, "bold")).pack(side="left", padx=6)

        formatos = [
            ("50 x 25 mm  — 30/hoja carta (Default temporal)", "50x25mm"),
            ("67 x 25 mm  — 30/hoja carta (Avery 5160 estándar)", "67x25mm"),
            ("100 x 50 mm — 10/hoja carta (Avery 5163, recomendado)", "100x50mm"),
        ]
        for texto, valor in formatos:
            tk.Radiobutton(frm_fmt, text=texto, variable=self.qr_formato,
                           value=valor, bg=COLOR_BG,
                           font=("Arial", 9)).pack(anchor="w")

        f_tipo_qr = tk.Frame(frm_fmt, bg=COLOR_BG)
        f_tipo_qr.pack(fill="x", pady=(4, 0))
        tk.Label(f_tipo_qr, text="Contenido dentro del QR:", bg=COLOR_BG,
                 font=("Arial", 8, "bold"), fg="#1F4E78").pack(anchor="w")
        tk.Radiobutton(f_tipo_qr, text="Ficha completa (Celular ve: ID, Equipo, Serial, Sede)",
                       variable=self.qr_contenido_modo, value="completo", bg=COLOR_BG,
                       font=("Arial", 8)).pack(anchor="w")
        tk.Radiobutton(f_tipo_qr, text="Solo ID de activo (Ideal para pistola lectora en Excel)",
                       variable=self.qr_contenido_modo, value="solo_id", bg=COLOR_BG,
                       font=("Arial", 8)).pack(anchor="w")

        # ---- Seccion 3: Acciones principales -----------------------------
        frm_acc = tk.LabelFrame(parent, text="3. Acciones de Generacion y Exportacion",
                                 padx=10, pady=8, bg=COLOR_BG,
                                 font=("Arial", 10, "bold"))
        frm_acc.pack(fill="x", padx=15, pady=4)

        btn_row = tk.Frame(frm_acc, bg=COLOR_BG)
        btn_row.pack(fill="x")

        tk.Button(btn_row,
                  text="Generar Etiquetas QR (PNG)",
                  command=self._qr_generar_todos,
                  bg="#1F4E78", fg="white",
                  font=("Arial", 9, "bold"), padx=10, pady=6
                  ).pack(side="left", padx=(0, 6))

        tk.Button(btn_row,
                  text="Crear Planillas de Impresion (Carta)",
                  command=self._qr_crear_planilla,
                  bg="#28A745", fg="white",
                  font=("Arial", 9, "bold"), padx=10, pady=6
                  ).pack(side="left", padx=(0, 6))

        tk.Button(btn_row,
                  text="Exportar Base con IDs a Excel",
                  command=self._qr_exportar_excel,
                  bg="#007BFF", fg="white",
                  font=("Arial", 9, "bold"), padx=10, pady=6
                  ).pack(side="left", padx=(0, 6))

        tk.Button(btn_row,
                  text="Abrir Excel QR con Plantillas",
                  command=self._qr_abrir_excel,
                  bg="#17A2B8", fg="white",
                  font=("Arial", 9, "bold"), padx=10, pady=6
                  ).pack(side="left", padx=(0, 6))

        tk.Button(btn_row,
                  text="Abrir Carpeta de Salida",
                  command=self._qr_abrir_carpeta,
                  bg="#6C757D", fg="white",
                  font=("Arial", 9), padx=10, pady=6
                  ).pack(side="left")

        # Etiqueta de estado
        self.lbl_qr_estado = tk.Label(frm_acc, text="",
                                       bg=COLOR_BG, fg="#1F7A1F",
                                       font=("Arial", 9, "bold"))
        self.lbl_qr_estado.pack(anchor="w", pady=(6, 0))

        # ---- Seccion 4: Simulador / Lector de Escaneo en Campo -----------
        frm_scan = tk.LabelFrame(parent, text="4. Lector / Verificador de Escaneo en Campo (Prueba en Vivo)",
                                  padx=10, pady=6, bg=COLOR_BG,
                                  font=("Arial", 10, "bold"))
        frm_scan.pack(fill="x", padx=15, pady=4)

        scan_row = tk.Frame(frm_scan, bg=COLOR_BG)
        scan_row.pack(fill="x")

        tk.Label(scan_row, text="Escanear ID o Serial:", bg=COLOR_BG,
                 font=("Arial", 9, "bold")).pack(side="left")
        e_scan = tk.Entry(scan_row, textvariable=self.qr_scan_input, width=28,
                          font=("Consolas", 10, "bold"))
        e_scan.pack(side="left", padx=8)
        e_scan.bind("<Return>", self._qr_probar_escaneo)

        tk.Button(scan_row, text="Consultar", command=self._qr_probar_escaneo,
                  bg="#17A2B8", fg="white", font=("Arial", 9, "bold")).pack(side="left", padx=(0, 10))

        self.lbl_scan_resultado = tk.Label(
            scan_row, text="← Escanea un QR o escribe un ID (ej: LAP-2026-0001) y presiona Enter",
            bg=COLOR_BG, fg="#555555", font=("Arial", 9)
        )
        self.lbl_scan_resultado.pack(side="left")

        # ---- Seccion 5: Bitacora QR --------------------------------------
        frm_log = tk.LabelFrame(parent, text="5. Bitacora de Actividades QR",
                                  padx=10, pady=6, bg=COLOR_BG,
                                  font=("Arial", 10, "bold"))
        frm_log.pack(fill="both", expand=True, padx=15, pady=(4, 10))

        vsb_qr = ttk.Scrollbar(frm_log, orient="vertical")
        self.txt_qr_log = tk.Text(frm_log, bg="#1E1E1E", fg="#D4D4D4",
                                   font=("Consolas", 9), height=7,
                                   yscrollcommand=vsb_qr.set)
        vsb_qr.config(command=self.txt_qr_log.yview)
        self.txt_qr_log.pack(side="left", fill="both", expand=True)
        vsb_qr.pack(side="right", fill="y")
        self.txt_qr_log.tag_config("ok",    foreground="#6BCF6B")
        self.txt_qr_log.tag_config("warn",  foreground="#FFD700")
        self.txt_qr_log.tag_config("error", foreground="#FF6B6B")
        self.txt_qr_log.tag_config("info",  foreground="#87CEEB")

    # ------------------------------------------------------------------
    # Modulo QR — Metodos de Generacion y Acciones
    # ------------------------------------------------------------------

    def log_qr(self, msg, tag="info"):
        """Registra un mensaje en la bitacora de la pestana QR."""
        if hasattr(self, "txt_qr_log") and self.txt_qr_log:
            ts = datetime.now().strftime("[%H:%M:%S] ")
            self.txt_qr_log.insert(tk.END, ts + str(msg) + "\n", tag)
            self.txt_qr_log.see(tk.END)
            self.root.update_idletasks()

    def _qr_actualizar_resumen_fuente(self):
        """Actualiza la etiqueta informativa de fuente de datos QR."""
        if not hasattr(self, "lbl_qr_fuente_info") or not self.lbl_qr_fuente_info:
            return
        modo = self.qr_fuente.get() if hasattr(self, "qr_fuente") else "excel"
        items_tree = self.tree.get_children() if hasattr(self, "tree") else []
        n_tree = len(items_tree)

        ruta = self.ruta_excel.get() if hasattr(self, "ruta_excel") else ""
        hoja = self.nombre_hoja.get().strip() if hasattr(self, "nombre_hoja") else ""

        if modo == "excel":
            if ruta and os.path.exists(ruta):
                try:
                    wb = openpyxl.load_workbook(ruta, read_only=True)
                    ws = wb[hoja] if hoja in wb.sheetnames else wb.active
                    hdr_row = self._detectar_fila_encabezados(ws)
                    n_filas = max(0, (ws.max_row or 0) - hdr_row)
                    wb.close()
                    txt = (f"Fuente activa: Archivo Excel [{hoja or os.path.basename(ruta)}]\n"
                           f" • ~{n_filas} activos con Ficha Completa (Host, Serial, Modelo, Resp, Sede)\n"
                           f" TOTAL: {n_filas} etiquetas oficiales a generar")
                except Exception:
                    txt = f"Fuente activa: Archivo Excel ({os.path.basename(ruta)})"
            else:
                txt = "Fuente activa: Excel (Carga un archivo en la pestaña 1 primero)"
        elif modo == "auto":
            if n_tree > 0:
                txt = (f"Fuente activa: Tabla de Verificación Física\n"
                       f" • {n_tree} activos cargados en la tabla\n"
                       f" TOTAL: {n_tree} etiquetas personalizadas a generar")
            else:
                txt = "Fuente activa: Tabla vacía (Pega datos en pestaña 1 o selecciona Excel)"
        else:
            n_lap = len(self.PRESET_CLEMONT)
            n_tab = len(self.PRESET_TABLETS)
            n_mon = len(self.PRESET_MONITORES)
            total = n_lap + n_tab + n_mon
            txt = (f"Fuente activa: Preset Clemont (Demo)\n"
                   f" • {n_lap} Laptops/PCs | {n_tab} Tablets | {n_mon} Monitores\n"
                   f" TOTAL: {total} activos demo listos para rotular")
        self.lbl_qr_fuente_info.config(text=txt)

    def _obtener_lista_equipos_qr(self):
        """
        Retorna la lista completa de equipos con sus IDs de Asset Tag asignados y
        toda su informacion: Equipo, Serial, Modelo, Responsable, Sede, Estado.
        """
        equipos = []
        modo = self.qr_fuente.get() if hasattr(self, "qr_fuente") else "excel"
        items_tree = self.tree.get_children() if hasattr(self, "tree") else []

        # 1. Modo Excel cargado (Recomendado: lee todos los datos de la hoja)
        if modo == "excel" or (modo == "auto" and len(items_tree) == 0):
            ruta = self.ruta_excel.get() if hasattr(self, "ruta_excel") else ""
            hoja = self.nombre_hoja.get().strip() if hasattr(self, "nombre_hoja") else ""
            if ruta and os.path.exists(ruta):
                try:
                    wb = openpyxl.load_workbook(ruta, data_only=True)
                    if not hoja or hoja not in wb.sheetnames:
                        hoja = wb.sheetnames[0]
                    ws = wb[hoja]

                    hdr_row = self._detectar_fila_encabezados(ws)
                    data_start = hdr_row + 1
                    total_row = self._detectar_fila_total(ws, data_start)

                    col_sede   = self._get_col_idx(self.col_sede_var.get(), default=self.COL_SEDE)
                    col_equipo = self._get_col_idx(self.col_equipo_var.get(), default=self.COL_EQUIPO)
                    col_obs    = self._get_col_idx(self.col_serial_var.get(), default=self.COL_OBS)
                    col_resp   = self._get_col_idx(self.col_resp_var.get(), default=self.COL_RESP) if hasattr(self, "col_resp_var") else 9
                    col_verif  = self._get_col_idx(self.col_verif_var.get(), default=self.COL_VERIF)

                    contadores = {}
                    for r in range(data_start, total_row):
                        nom_eq  = str(ws.cell(r, col_equipo).value or "").strip()
                        raw_obs = str(ws.cell(r, col_obs).value or "").strip()
                        sede    = str(ws.cell(r, col_sede).value or "").strip()
                        resp    = str(ws.cell(r, col_resp).value or "").strip()
                        estado  = str(ws.cell(r, col_verif).value or ws.cell(r, 10).value or "Activo").strip()

                        if not nom_eq and not raw_obs:
                            continue

                        # Extraer serial y modelo
                        sns = extraer_seriales(raw_obs)
                        if sns:
                            serial = sns[0]
                        else:
                            if re.search(r"(?i)^modelo\s*:", raw_obs):
                                serial = "S/N NO REGISTRADO"
                            else:
                                words = re.findall(r"[A-Za-z0-9]{6,}", raw_obs)
                                serial = words[0] if words else "S/N NO REGISTRADO"

                        m_mod = re.search(r"(?i)Modelo\s*:\s*([^;]+)", raw_obs)
                        modelo = m_mod.group(1).strip() if m_mod else ""
                        if not modelo:
                            fab = str(ws.cell(r, 6).value or "").strip()
                            soft = str(ws.cell(r, 7).value or "").strip()
                            if fab and fab.upper() not in ("SISTEMA OPERATIVO", "MICROSOFT"):
                                modelo = f"{fab}".strip()
                        if not modelo:
                            modelo = "Generico"

                        nom_up = (nom_eq + " " + raw_obs).upper()
                        if any(k in nom_up for k in ["MONITOR", "PANTALLA", "LG 27", "VIEWSONIC", "SAMSUNG", "JANUS"]):
                            cat, tipo = "MON", "Monitor"
                        elif any(k in nom_up for k in ["TAB", "GALAXY TAB", "IPAD"]):
                            cat, tipo = "TAB", "Tablet"
                        elif any(k in nom_up for k in ["IMP", "EPSON", "PRINTER", "LASERJET", "HP DESKJET"]):
                            cat, tipo = "IMP", "Impresora"
                        elif any(k in nom_up for k in ["SWITCH", "ROUTER", "ACCESS POINT", "FIREWALL"]):
                            cat, tipo = "RED", "Redes"
                        elif any(k in nom_up for k in ["DESKTOP", "TORRE", "PC-", "ALL IN ONE"]):
                            cat, tipo = "PC", "PC Escritorio"
                        else:
                            cat, tipo = "LAP", "Laptop"

                        contadores[cat] = contadores.get(cat, 0) + 1
                        correlativo = contadores[cat]
                        id_tag = f"{cat}-2026-{correlativo:04d}"

                        equipos.append({
                            "id": id_tag,
                            "id_activo": id_tag,
                            "categoria": cat,
                            "tipo": tipo,
                            "equipo": nom_eq,
                            "modelo": modelo or tipo,
                            "serial": serial,
                            "sede": sede,
                            "responsable": resp,
                            "estado": estado,
                            "observaciones": raw_obs,
                            "fila_excel": r
                        })

                    wb.close()
                    if equipos:
                        return equipos
                except Exception as e:
                    self.log_qr(f"Aviso leyendo Excel para QR: {e}", "warn")

        # 2. Modo Tabla de Verificación Física
        if (modo == "auto" and len(items_tree) > 0) or modo == "tree":
            contadores = {}
            for item_id in items_tree:
                vals = self.tree.item(item_id)["values"]
                if len(vals) >= 10:
                    campos = extraer_campos_completos(vals)
                    idx_eq_c  = self._get_col_idx(self.col_equipo_var.get(), 4) - 1 if hasattr(self, "col_equipo_var") else 3
                    idx_ser_c = self._get_col_idx(self.col_serial_var.get(), 11) - 1 if hasattr(self, "col_serial_var") else 10
                    idx_sd_c  = self._get_col_idx(self.col_sede_var.get(), 3) - 1 if hasattr(self, "col_sede_var") else 2
                    idx_rp_c  = self._get_col_idx(self.col_resp_var.get(), 9) - 1 if hasattr(self, "col_resp_var") else 8
                    idx_vf_c  = self._get_col_idx(self.col_verif_var.get(), 12) - 1 if hasattr(self, "col_verif_var") else 11

                    item_num = (str(vals[0]) if len(vals) > 0 and str(vals[0]).strip() else "") or campos.get("id_activo", "")
                    nom_eq   = (str(vals[idx_eq_c]) if 0 <= idx_eq_c < len(vals) and str(vals[idx_eq_c]).strip() else "") or campos.get("equipo", "")
                    raw_ser  = (str(vals[idx_ser_c]) if 0 <= idx_ser_c < len(vals) and str(vals[idx_ser_c]).strip() else "") or campos.get("serial", "")
                    sede     = (str(vals[idx_sd_c]) if 0 <= idx_sd_c < len(vals) and str(vals[idx_sd_c]).strip() else "") or campos.get("sede", "")
                    resp     = (str(vals[idx_rp_c]) if 0 <= idx_rp_c < len(vals) and str(vals[idx_rp_c]).strip() else "") or campos.get("responsable", "")
                    estado   = (str(vals[idx_vf_c]) if 0 <= idx_vf_c < len(vals) and str(vals[idx_vf_c]).strip() else "") or campos.get("estado", "OK")
                    obs      = raw_ser or campos.get("modelo", "")
                    serial   = raw_ser
                elif len(vals) >= 7:
                    item_num = str(vals[0])
                    nom_eq   = str(vals[1] if len(vals) > 1 else "")
                    serial   = str(vals[2] if len(vals) > 2 else "")
                    estado   = str(vals[3] if len(vals) > 3 else "OK")
                    sede     = str(vals[4] if len(vals) > 4 else "")
                    resp     = str(vals[5] if len(vals) > 5 else "")
                    obs      = str(vals[6] if len(vals) > 6 else "")
                else:
                    item_num = ""
                    nom_eq   = str(vals[0] if len(vals) > 0 else "")
                    serial   = str(vals[1] if len(vals) > 1 else "")
                    estado   = str(vals[2] if len(vals) > 2 else "OK")
                    sede     = str(vals[3] if len(vals) > 3 else "")
                    resp     = ""
                    obs      = ""

                # Extraer serial limpio si vino con texto o prefijo
                sns = extraer_seriales(serial)
                if sns:
                    serial = sns[0]
                elif len(serial) < 6 or serial.isdigit():
                    sns_obs = extraer_seriales(obs)
                    serial = sns_obs[0] if sns_obs else (serial if len(serial) >= 6 else "S/N NO REGISTRADO")

                m_mod = re.search(r"(?i)Modelo\s*:\s*([^;]+)", obs)
                modelo = m_mod.group(1).strip() if m_mod else ""
                if not modelo and "MODELO" in obs.upper():
                    m_alt = re.search(r"(?i)Modelo\s*[:\-]?\s*([^\n;]+)", obs)
                    modelo = m_alt.group(1).strip() if m_alt else ""
                if not modelo:
                    modelo = "Generico"

                nom_up = nom_eq.upper()
                if any(k in nom_up for k in ["MONITOR", "PANTALLA", "LG 27", "VIEWSONIC", "SAMSUNG", "JANUS"]):
                    cat, tipo = "MON", "Monitor"
                elif any(k in nom_up for k in ["TAB", "GALAXY TAB", "IPAD"]):
                    cat, tipo = "TAB", "Tablet"
                elif any(k in nom_up for k in ["IMP", "EPSON", "PRINTER", "LASERJET", "HP DESKJET"]):
                    cat, tipo = "IMP", "Impresora"
                elif any(k in nom_up for k in ["SWITCH", "ROUTER", "ACCESS POINT", "FIREWALL"]):
                    cat, tipo = "RED", "Redes"
                elif any(k in nom_up for k in ["DESKTOP", "TORRE", "PC-", "ALL IN ONE"]):
                    cat, tipo = "PC", "PC Escritorio"
                elif any(k in nom_up for k in ["LAP", "MACBOOK", "PORTATIL", "NOTEBOOK"]):
                    cat, tipo = "LAP", "Laptop"
                else:
                    cat, tipo = "LAP", "Equipo / Host"

                contadores[cat] = contadores.get(cat, 0) + 1
                correlativo = contadores[cat]
                id_tag = f"{cat}-2026-{correlativo:04d}"

                equipos.append({
                    "id": id_tag,
                    "id_activo": id_tag,
                    "categoria": cat,
                    "tipo": tipo,
                    "equipo": nom_eq,
                    "modelo": modelo or tipo,
                    "serial": serial,
                    "sede": sede,
                    "responsable": resp,
                    "estado": estado,
                    "observaciones": obs
                })
            return equipos

        # 3. Fallback: Presets de Clemont (para pruebas y demostracion)
        for i, (eq, ser, est, sede) in enumerate(self.PRESET_CLEMONT, 1):
            equipos.append({
                "id": f"LAP-2026-{i:04d}",
                "categoria": "LAP",
                "tipo": "Laptop / PC",
                "equipo": eq,
                "modelo": "Laptop Corporativa",
                "serial": ser,
                "sede": sede,
                "responsable": "Usuario Asignado",
                "estado": est
            })

        for i, fila in enumerate(self.PRESET_TABLETS, 1):
            ser_limpio = fila[10].replace("Serial:", "").replace("Serial :", "").strip()
            equipos.append({
                "id": f"TAB-2026-{i:04d}",
                "categoria": "TAB",
                "tipo": "Tablet",
                "equipo": f"{fila[4]} ({fila[5]})",
                "modelo": f"{fila[5]}",
                "serial": ser_limpio,
                "sede": fila[2],
                "responsable": fila[8],
                "estado": fila[9]
            })

        for i, fila in enumerate(self.PRESET_MONITORES, 1):
            ser_limpio = fila[10].replace("Serial:", "").replace("Serial :", "").strip()
            equipos.append({
                "id": f"MON-2026-{i:04d}",
                "categoria": "MON",
                "tipo": "Monitor",
                "equipo": f"{fila[4]}",
                "modelo": "Monitor 24\"",
                "serial": ser_limpio,
                "sede": fila[2],
                "responsable": fila[8],
                "estado": fila[9]
            })

        return equipos

    def _generar_etiqueta_individual(self, item, formato="50x25mm", carpeta=None):
        """
        Genera una imagen PNG a 300 DPI de una etiqueta individual con:
        - Codigo QR (codifica UNICAMENTE el ID_Activo, ej. LAP-2026-0001)
        - Encabezado institucional (ej. CLEMONT · CONTROL DE ACTIVOS)
        - ID del Activo en tipografia grande y destacada
        - Nombre del equipo / host
        - Serial del fabricante
        - Sede fisica oficial
        """
        if not QR_DISPONIBLE:
            return None

        # Parametros por formato a 300 DPI
        if formato == "100x50mm":
            w, h = 1200, 600
            qr_dim = 540
            x_txt = 580
            sz_hdr, sz_id, sz_eq, sz_sub = 28, 56, 32, 24
        elif formato == "67x25mm":
            w, h = 788, 300
            qr_dim = 270
            x_txt = 300
            sz_hdr, sz_id, sz_eq, sz_sub = 18, 36, 22, 16
        else:  # 50x25mm
            w, h = 591, 295
            qr_dim = 265
            x_txt = 285
            sz_hdr, sz_id, sz_eq, sz_sub = 15, 30, 18, 14

        img = Image.new("RGB", (w, h), "white")
        draw = ImageDraw.Draw(img)

        empresa = self.qr_empresa.get().strip().upper() if hasattr(self, "qr_empresa") and self.qr_empresa.get() else "CLEMONT"
        txt_hdr = f"{empresa}  ·  CONTROL DE ACTIVOS"

        nom_eq = str(item.get("equipo", "") or "").strip()
        serial = str(item.get("serial", "") or "").strip()
        modelo = str(item.get("modelo", "") or "").strip()
        resp   = str(item.get("responsable", "") or "").strip()
        sede   = str(item.get("sede", "") or "").strip()
        estado = str(item.get("estado", "Activo") or "Activo").strip()

        # Generar QR — Ficha completa (para celular) o Solo ID (para pistola lectora)
        qr = qrcode.QRCode(
            box_size=10,
            border=1,
            error_correction=qrcode.constants.ERROR_CORRECT_M
        )
        if getattr(self, "qr_contenido_modo", None) and self.qr_contenido_modo.get() == "solo_id":
            qr_data = item["id"]
        else:
            qr_data = (
                f"ACTIVO: {item['id']}\n"
                f"EQUIPO: {nom_eq}\n"
                f"MODELO: {modelo or 'Generico'}\n"
                f"SERIAL: {serial}\n"
                f"RESPONSABLE: {resp or 'Sin asignar'}\n"
                f"SEDE: {sede}\n"
                f"ESTADO: {estado}"
            )
        qr.add_data(qr_data)
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
        qr_img = qr_img.resize((qr_dim, qr_dim), Image.Resampling.LANCZOS)
        img.paste(qr_img, (15, (h - qr_dim) // 2))

        # Guia de corte perimetral sutil
        draw.rectangle([(1, 1), (w - 2, h - 2)], outline="#C8C8C8", width=1)
        # Separador vertical
        draw.line([(qr_dim + 24, 12), (qr_dim + 24, h - 12)], fill="#E0E0E0", width=1)

        # Diseno de impresion con TODA la informacion segun el tamano
        if formato == "100x50mm":
            draw.text((x_txt, 25), txt_hdr, fill="#1F4E78", font=_obtener_fuente(24, bold=True))
            draw.text((x_txt, 65), item["id"], fill="#000000", font=_obtener_fuente(46, bold=True))
            draw.line([(x_txt, 130), (w - 30, 130)], fill="#1F4E78", width=2)
            draw.text((x_txt, 145), f"Host: {nom_eq[:28]}", fill="#111111", font=_obtener_fuente(28, bold=True))
            mod_display = modelo or item.get("tipo", "Activo")
            draw.text((x_txt, 195), f"Modelo: {mod_display[:30]}", fill="#333333", font=_obtener_fuente(22, bold=False))
            draw.text((x_txt, 245), f"S/N: {serial[:30]}", fill="#000000", font=_obtener_fuente(26, bold=True))
            draw.text((x_txt, 300), f"Usuario: {(resp or 'Sin asignar')[:28]}", fill="#333333", font=_obtener_fuente(22, bold=False))
            draw.text((x_txt, 355), f"Sede: {sede[:30]}", fill="#444444", font=_obtener_fuente(22, bold=False))
            draw.text((x_txt, 410), f"Estado: {estado[:20]}", fill="#1F7A1F", font=_obtener_fuente(22, bold=True))

        elif formato == "67x25mm":
            draw.text((x_txt, 10), txt_hdr, fill="#1F4E78", font=_obtener_fuente(15, bold=True))
            draw.text((x_txt, 34), item["id"], fill="#000000", font=_obtener_fuente(30, bold=True))
            draw.text((x_txt, 75), f"Host: {nom_eq[:26]}", fill="#111111", font=_obtener_fuente(18, bold=True))
            mod_display = modelo or item.get("tipo", "Activo")
            draw.text((x_txt, 105), f"Mod: {mod_display[:24]}", fill="#333333", font=_obtener_fuente(14, bold=False))
            draw.text((x_txt, 133), f"S/N: {serial[:22]}", fill="#000000", font=_obtener_fuente(16, bold=True))
            draw.text((x_txt, 162), f"Resp: {(resp or 'Sin asignar')[:24]}", fill="#333333", font=_obtener_fuente(14, bold=False))
            draw.text((x_txt, 190), f"Sede: {sede[:20]} | {estado[:12]}", fill="#444444", font=_obtener_fuente(14, bold=False))

        else:  # 50x25mm
            draw.text((x_txt, 8), txt_hdr, fill="#1F4E78", font=_obtener_fuente(13, bold=True))
            draw.text((x_txt, 28), item["id"], fill="#000000", font=_obtener_fuente(24, bold=True))
            draw.text((x_txt, 62), f"Host: {nom_eq[:22]}", fill="#111111", font=_obtener_fuente(15, bold=True))
            mod_display = modelo or item.get("tipo", "Activo")
            draw.text((x_txt, 90), f"Mod: {mod_display[:22]}", fill="#333333", font=_obtener_fuente(13, bold=False))
            draw.text((x_txt, 118), f"S/N: {serial[:20]}", fill="#000000", font=_obtener_fuente(14, bold=True))
            draw.text((x_txt, 146), f"Resp: {(resp or 'Sin asignar')[:20]}", fill="#333333", font=_obtener_fuente(13, bold=False))
            draw.text((x_txt, 174), f"Sede: {sede[:20]}", fill="#444444", font=_obtener_fuente(13, bold=False))
            draw.text((x_txt, 202), f"Estado: {estado[:16]}", fill="#1F7A1F", font=_obtener_fuente(13, bold=True))

        if not carpeta:
            carpeta = self._obtener_carpeta_qr()
        os.makedirs(carpeta, exist_ok=True)
        clean_eq = re.sub(r"[^A-Za-z0-9_-]", "", item["equipo"])[:20]
        ruta = os.path.join(carpeta, f"{item['id']}_{clean_eq}.png")
        img.save(ruta, dpi=(300, 300))
        return ruta

    def _qr_generar_todos(self):
        """Genera los archivos PNG individuales de todas las etiquetas."""
        if not QR_DISPONIBLE:
            messagebox.showerror("Error", "Las librerias qrcode y Pillow no estan disponibles.")
            return

        fmt = self.qr_formato.get()
        equipos = self._obtener_lista_equipos_qr()
        total = len(equipos)
        carpeta = self._obtener_carpeta_qr()
        os.makedirs(carpeta, exist_ok=True)

        self.log_qr(f"=== INICIANDO GENERACION DE ETIQUETAS QR ({total} equipos) ===", "ok")
        self.log_qr(f"Formato seleccionado: {fmt}  |  Carpeta destino: ./{carpeta}/", "info")

        exitos = 0
        for eq in equipos:
            try:
                ruta = self._generar_etiqueta_individual(eq, formato=fmt, carpeta=carpeta)
                exitos += 1
                self.log_qr(f"✔ [{eq['id']}] {eq['equipo']}  (S/N: {eq['serial']}) -> {os.path.basename(ruta)}", "ok")
            except Exception as e:
                self.log_qr(f"❌ Error al generar {eq['id']}: {e}", "error")

        self.lbl_qr_estado.config(
            text=f"✔ {exitos} de {total} etiquetas generadas con exito en ./{carpeta}/",
            fg="#1F7A1F"
        )
        self.log_qr(f"=== GENERACION FINALIZADA: {exitos}/{total} etiquetas listas ===", "ok")
        messagebox.showinfo(
            "Etiquetas Generadas",
            f"Se han generado {exitos} etiquetas QR en la carpeta:\n\n"
            f"{os.path.abspath(carpeta)}\n\n"
            "Ahora puedes presionar 'Crear Planillas de Impresion' para armar las hojas tamaño carta listas para imprimir."
        )

    def _qr_crear_planilla(self):
        """Organiza las etiquetas generadas en cuadriculas de impresion tamaño carta."""
        if not QR_DISPONIBLE:
            messagebox.showerror("Error", "qrcode y Pillow no están disponibles.")
            return

        fmt = self.qr_formato.get()
        equipos = self._obtener_lista_equipos_qr()
        total = len(equipos)
        carpeta = self._obtener_carpeta_qr()
        os.makedirs(carpeta, exist_ok=True)

        self.log_qr(f"=== ARMANDO PLANILLAS DE IMPRESION (TAMANO CARTA 300 DPI) ===", "ok")

        # 1. Asegurar que las etiquetas individuales existan
        rutas_etiquetas = []
        for eq in equipos:
            clean_eq = re.sub(r"[^A-Za-z0-9_-]", "", eq["equipo"])[:20]
            ruta = os.path.join(carpeta, f"{eq['id']}_{clean_eq}.png")
            if not os.path.exists(ruta):
                ruta = self._generar_etiqueta_individual(eq, formato=fmt, carpeta=carpeta)
            rutas_etiquetas.append((eq, ruta))

        # Tamano Carta a 300 DPI: 2550 x 3300 px
        PAGE_W, PAGE_H = 2550, 3300

        if fmt == "100x50mm":
            cols, rows = 2, 5
            lbl_w, lbl_h = 1200, 600
            gap_x, gap_y = 42, 10
            margin_x = (PAGE_W - (cols * lbl_w + (cols - 1) * gap_x)) // 2
            margin_y = 160
        else:
            cols, rows = 3, 10
            lbl_w, lbl_h = 788, 300
            gap_x, gap_y = 38, 6
            margin_x = (PAGE_W - (cols * lbl_w + (cols - 1) * gap_x)) // 2
            margin_y = 150

        por_hoja = cols * rows
        num_hojas = (total + por_hoja - 1) // por_hoja

        empresa = self.qr_empresa.get().strip().upper() if hasattr(self, "qr_empresa") and self.qr_empresa.get() else "CLEMONT"
        f_titulo = _obtener_fuente(28, bold=True)
        f_sub = _obtener_fuente(18, bold=False)

        planillas_generadas = []

        for h_idx in range(num_hojas):
            hoja_img = Image.new("RGB", (PAGE_W, PAGE_H), "white")
            draw = ImageDraw.Draw(hoja_img)

            num_hoja = h_idx + 1
            txt_hdr = f"{empresa}  ·  PLANILLA DE ETIQUETAS QR (Hoja {num_hoja} de {num_hojas})"
            txt_det = f"Formato: {fmt}  |  Resolucion: 300 DPI  |  Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            draw.text((margin_x, 40), txt_hdr, fill="#1F4E78", font=f_titulo)
            draw.text((margin_x, 85), txt_det, fill="#666666", font=f_sub)
            draw.line([(margin_x, 125), (PAGE_W - margin_x, 125)], fill="#B0B0B0", width=2)

            start_i = h_idx * por_hoja
            end_i = min(start_i + por_hoja, total)

            for pos, i in enumerate(range(start_i, end_i)):
                r = pos // cols
                c = pos % cols
                x = margin_x + c * (lbl_w + gap_x)
                y = margin_y + r * (lbl_h + gap_y)

                eq_obj, ruta_png = rutas_etiquetas[i]
                try:
                    lbl_img = Image.open(ruta_png).convert("RGB")
                    if lbl_img.size != (lbl_w, lbl_h):
                        lbl_img = lbl_img.resize((lbl_w, lbl_h), Image.Resampling.LANCZOS)
                    hoja_img.paste(lbl_img, (x, y))

                    # Linea de corte exterior sutil
                    draw.rectangle([(x - 1, y - 1), (x + lbl_w + 1, y + lbl_h + 1)],
                                   outline="#D0D0D0", width=1)
                except Exception as e:
                    self.log_qr(f"Error pegando etiqueta {eq_obj['id']}: {e}", "warn")

            nombre_planilla = f"Planilla_Impresion_Hoja_{num_hoja}_de_{num_hojas}.png"
            ruta_planilla = os.path.join(carpeta, nombre_planilla)
            hoja_img.save(ruta_planilla, dpi=(300, 300))
            planillas_generadas.append(ruta_planilla)
            self.log_qr(f"✔ Planilla guardada: {nombre_planilla} ({end_i - start_i} etiquetas)", "ok")

        self.lbl_qr_estado.config(
            text=f"✔ {len(planillas_generadas)} planilla(s) de impresion guardadas en ./{carpeta}/",
            fg="#1F7A1F"
        )
        self.log_qr(f"=== {len(planillas_generadas)} PLANILLAS LISTAS PARA IMPRIMIR ===", "ok")
        messagebox.showinfo(
            "Planillas Listas",
            f"Se han generado {len(planillas_generadas)} planillas tamano Carta (300 DPI) en:\n\n"
            f"{os.path.abspath(carpeta)}\n\n"
            "Puedes abrirlas e imprimirlas en papel adhesivo tamano carta directamente."
        )

    def _obtener_carpeta_qr(self):
        """Retorna la ruta absoluta de la carpeta de etiquetas QR respetando el directorio configurado."""
        if hasattr(self, "ruta_salida") and self.ruta_salida.get().strip():
            dir_sal = os.path.dirname(self.ruta_salida.get().strip())
            if dir_sal:
                return os.path.abspath(os.path.join(dir_sal, "etiquetas_qr"))
        if hasattr(self, "dir_salida") and self.dir_salida.get().strip():
            return os.path.abspath(os.path.join(self.dir_salida.get().strip(), "etiquetas_qr"))
        return os.path.abspath("etiquetas_qr")

    def _qr_exportar_excel(self):
        """Exporta un libro Excel con la base maestra de IDs y la hoja de escaneos de campo."""
        carpeta = self._obtener_carpeta_qr()
        os.makedirs(carpeta, exist_ok=True)
        ruta_salida = os.path.join(carpeta, "Inventario_Equipos_Clemont_Con_QR.xlsx")

        equipos = self._obtener_lista_equipos_qr()
        generar_plantilla_auditoria_qr(ruta_salida, equipos)

        self.log_qr(f"✔ Plantilla Excel exportada con {len(equipos)} activos: {os.path.basename(ruta_salida)}", "ok")
        messagebox.showinfo(
            "Excel Exportado",
            f"Se ha creado el archivo Excel con los {len(equipos)} activos y la plantilla de escaneos:\n\n"
            f"{os.path.abspath(ruta_salida)}"
        )

    def _qr_abrir_excel(self):
        """Abre el archivo Excel de plantillas QR generado a 1 clic."""
        carpeta = self._obtener_carpeta_qr()
        ruta = os.path.join(carpeta, "Inventario_Equipos_Clemont_Con_QR.xlsx")
        if not os.path.exists(ruta):
            ruta_alt = os.path.abspath(os.path.join("etiquetas_qr", "Inventario_Equipos_Clemont_Con_QR.xlsx"))
            if os.path.exists(ruta_alt):
                ruta = ruta_alt
            else:
                messagebox.showinfo("Aviso", "Primero exporta la base a Excel con el botón 'Exportar Base con IDs a Excel'.")
                return
        try:
            os.startfile(ruta)
        except Exception:
            subprocess.Popen(["start", "", ruta], shell=True)

    def _qr_probar_escaneo(self, event=None):
        """Busca en tiempo real el ID o Serial ingresado/escaneado para prueba en vivo."""
        termino = self.qr_scan_input.get().strip().upper()
        if not termino:
            return

        equipos = self._obtener_lista_equipos_qr()
        encontrado = None

        for eq in equipos:
            eq_id = eq["id"].upper()
            eq_sn = normalizar(eq["serial"])
            eq_nom = eq["equipo"].upper()
            term_norm = normalizar(termino)

            if (eq_id == termino or eq_id in termino or
                (eq_sn and (eq_sn == term_norm or eq_sn in term_norm)) or
                (eq_nom and (eq_nom == termino or eq_nom in termino))):
                encontrado = eq
                break

        if encontrado:
            msg = (f"✔ MATCH: [{encontrado['id']}] {encontrado['equipo']} | "
                   f"S/N: {encontrado['serial']} | Sede Oficial: {encontrado['sede']} | Estado: {encontrado['estado']}")
            self.lbl_scan_resultado.config(text=msg, fg="#1F7A1F")
            self.log_qr(f"Escaneo: '{termino}' -> Encontrado: {encontrado['id']} ({encontrado['equipo']})", "ok")
        else:
            msg = f"❌ Activo no encontrado para: '{termino}'"
            self.lbl_scan_resultado.config(text=msg, fg="#C00000")
            self.log_qr(f"Escaneo: '{termino}' -> NO REGISTRADO en la base", "warn")

        self.qr_scan_input.set("")

    def _qr_abrir_carpeta(self):
        """Abre la carpeta etiquetas_qr en el explorador de archivos de Windows."""
        carpeta = self._obtener_carpeta_qr()
        os.makedirs(carpeta, exist_ok=True)
        try:
            os.startfile(carpeta)
        except Exception:
            subprocess.Popen(["explorer", carpeta])

    # ------------------------------------------------------------------
    # Acciones de la GUI — Pestana Conciliacion
    # ------------------------------------------------------------------

    def _seleccionar_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona el archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx")]
        )
        if ruta:
            self.ruta_excel.set(ruta)
            dir_s, nom_s = os.path.split(ruta)
            nom_out = "AUDITADO_" + nom_s if not nom_s.startswith("AUDITADO_") else nom_s
            self.ruta_salida.set(os.path.join(dir_s, nom_out))
            self._cargar_hojas_en_combo()

    def _seleccionar_directorio_salida(self):
        """Permite al usuario seleccionar la carpeta destino para los archivos de salida y etiquetas."""
        carpeta_actual = os.path.dirname(self.ruta_salida.get()) if (self.ruta_salida.get() and os.path.isabs(self.ruta_salida.get())) else os.getcwd()
        nueva_carpeta = filedialog.askdirectory(
            title="Seleccionar carpeta de destino para resultados y etiquetas",
            initialdir=carpeta_actual
        )
        if nueva_carpeta:
            self.dir_salida.set(nueva_carpeta)
            nom_s = os.path.basename(self.ruta_salida.get()) or "AUDITADO_Inventario.xlsx"
            self.ruta_salida.set(os.path.join(nueva_carpeta, nom_s))
            self.log(f"Carpeta de destino configurada en: {nueva_carpeta}", "info")

    def _seleccionar_ruta_salida(self):
        ruta_sug = self.ruta_salida.get() or "AUDITADO_Inventario.xlsx"
        dir_s, nom_s = os.path.split(ruta_sug)
        nueva = filedialog.asksaveasfilename(
            title="Seleccionar dónde guardar el archivo auditado",
            initialdir=dir_s or os.getcwd(),
            initialfile=nom_s or "AUDITADO_Inventario.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")]
        )
        if nueva:
            self.ruta_salida.set(nueva)
            self.log(f"Ruta de guardado configurada en: {nueva}", "info")

    def _abrir_excel_resultado(self):
        ruta = self.ruta_salida.get().strip()
        if not ruta or not os.path.exists(ruta):
            messagebox.showinfo("Aviso", "Primero ejecuta la conciliacion para generar el archivo auditado.")
            return
        try:
            os.startfile(ruta)
        except Exception:
            subprocess.Popen(["start", "", ruta], shell=True)

    def _abrir_carpeta_resultado(self):
        ruta = self.ruta_salida.get().strip()
        carpeta = os.path.dirname(ruta) if ruta else os.getcwd()
        if not os.path.exists(carpeta):
            carpeta = os.getcwd()
        try:
            os.startfile(carpeta)
        except Exception:
            subprocess.Popen(["explorer", carpeta])

    def _actualizar_estado_combos(self):
        """Habilita o deshabilita los dropdowns segun las casillas de verificacion (chulos)."""
        if hasattr(self, "combo_col_serial"):
            self.combo_col_serial.config(state="readonly" if self.chk_col_serial.get() else "disabled")
            self.combo_col_equipo.config(state="readonly" if self.chk_col_equipo.get() else "disabled")
            self.combo_col_sede.config(state="readonly" if self.chk_col_sede.get() else "disabled")
            self.combo_col_resp.config(state="readonly" if self.chk_col_resp.get() else "disabled")
            self.combo_col_verif.config(state="readonly" if self.chk_col_verif.get() else "disabled")

    def _cargar_hojas_en_combo(self):
        """Lee las hojas del Excel y llena el dropdown automaticamente."""
        ruta = self.ruta_excel.get()
        if not ruta or not os.path.exists(ruta):
            return
        try:
            wb = openpyxl.load_workbook(ruta, read_only=True)
            hojas = wb.sheetnames
            wb.close()

            self.hojas_disponibles = hojas
            self.combo_hojas["values"] = hojas

            # Intentar preseleccionar la hoja de inventario automaticamente
            hoja_auto = None
            for h in hojas:
                if "INVENTARIO" in h.upper():
                    hoja_auto = h
                    break
            if hoja_auto is None:
                hoja_auto = hojas[0]  # si no hay candidata, tomar la primera

            self.combo_hojas.set(hoja_auto)
            self.nombre_hoja.set(hoja_auto)

            # Detectar columnas automaticamente en la hoja preseleccionada
            self._auto_detectar_columnas_hoja(hoja_auto)

            # Actualizar etiqueta informativa
            n = len(hojas)
            self.lbl_hoja_info.config(
                text=f"{n} hoja{'s' if n != 1 else ''} encontrada{'s' if n != 1 else ''}  |  Seleccionada: [{hoja_auto}]",
                fg="#1F7A1F"
            )
            self.log(f"Excel cargado. Hojas disponibles: {' | '.join(hojas)}", "info")
            self.log(f"Hoja preseleccionada automaticamente: [{hoja_auto}]  "
                     f"(puedes cambiarla en el dropdown)", "ok")

        except Exception as e:
            self.log("No se pudo leer el archivo: " + str(e), "warn")

    def _mostrar_hojas(self):
        """Mantener por compatibilidad — ahora hace lo mismo que _cargar_hojas_en_combo."""
        self._cargar_hojas_en_combo()

    _cargar_hojas_excel = _cargar_hojas_en_combo

    def _reconfigurar_columnas_tree(self, cols):
        """Reconfigura dinamicamente las columnas, encabezados y anchos de self.tree."""
        if not hasattr(self, "tree") or not self.tree:
            return
        self.tree["columns"] = tuple(cols)
        for c in cols:
            self.tree.heading(c, text=str(c))
            c_up = str(c).upper()
            w = max(75, min(240, len(str(c)) * 11 + 25))
            anchor = "center" if any(k in c_up for k in ["#", "ITEM", "ID", "ESTADO", "VERIF", "CANT", "NUM"]) else "w"
            self.tree.column(c, width=w, anchor=anchor)

    def _on_hoja_seleccionada(self, event=None):
        """Disparado cuando el usuario cambia la hoja en el dropdown."""
        hoja = self.nombre_hoja.get().strip()
        if hoja:
            self._auto_detectar_columnas_hoja(hoja)
            self.lbl_hoja_info.config(
                text=f"Hoja activa: [{hoja}] (columnas auto-detectadas)",
                fg="#1F7A1F"
            )

    def _auto_detectar_columnas_hoja(self, hoja_especifica=None):
        """Lee los encabezados de la hoja seleccionada, reconfigura self.tree dinamicamente y actualiza los dropdowns de mapeo."""
        ruta = self.ruta_excel.get()
        hoja = hoja_especifica or self.nombre_hoja.get().strip()
        if not ruta or not os.path.exists(ruta) or not hoja:
            return
        try:
            wb = openpyxl.load_workbook(ruta, data_only=True)
            if hoja not in wb.sheetnames:
                wb.close()
                return
            ws = wb[hoja]

            fila_hdr = self._detectar_fila_encabezados(ws)
            max_c = min(ws.max_column or 20, 50)

            headers_raw = []
            ultimo_col_con_datos = 0
            for c in range(1, max_c + 1):
                val = str(ws.cell(fila_hdr, c).value or "").strip()
                val_limpio = " ".join(val.split())
                if val_limpio:
                    ultimo_col_con_datos = c
                headers_raw.append((c, val_limpio))

            num_cols_reales = max(ultimo_col_con_datos, 4)
            headers_tree = []
            seen_headers = {}
            opciones_cols = []
            col_serial_det = None
            col_equipo_det = None
            col_sede_det = None
            col_resp_det = None
            col_verif_det = None

            for c, val_limpio in headers_raw[:num_cols_reales]:
                letra = openpyxl.utils.get_column_letter(c)
                h_base = val_limpio if val_limpio else f"Col_{letra}"
                if h_base in seen_headers:
                    seen_headers[h_base] += 1
                    h_name = f"{h_base}_{c}"
                else:
                    seen_headers[h_base] = 1
                    h_name = h_base
                headers_tree.append(h_name)

                etiqueta = f"Col {letra} ({c})" + (f" - {val_limpio[:25]}" if val_limpio else "")
                opciones_cols.append(etiqueta)

                val_up = val_limpio.upper()
                if not col_serial_det and any(k in val_up for k in ["SERIAL", "S/N", "OBSERVACION"]):
                    col_serial_det = etiqueta
                if not col_equipo_det and any(k in val_up for k in ["EQUIPO", "HOST", "DISPOSITIVO", "SERVIDOR"]):
                    col_equipo_det = etiqueta
                if not col_sede_det and any(k in val_up for k in ["SEDE", "UBICACION", "CIUDAD", "LUGAR"]):
                    col_sede_det = etiqueta
                if not col_resp_det and any(k in val_up for k in ["RESPONSABLE", "USUARIO", "ENCARGADO", "EMPLEADO", "ASIGNADO"]):
                    col_resp_det = etiqueta
                # Priorizar VERIFICADO o AUDITORIA para no confundir con columna "Estado" (Activo/Inactivo)
                if not col_verif_det and any(k in val_up for k in ["VERIFICADO", "AUDITORIA", "CHECK"]):
                    col_verif_det = etiqueta

            # Fallback para resultado si no habia columna "Verificado"
            if not col_verif_det:
                for c, val_limpio in headers_raw[:num_cols_reales]:
                    if "ESTADO" in val_limpio.upper():
                        letra = openpyxl.utils.get_column_letter(c)
                        col_verif_det = f"Col {letra} ({c}) - {val_limpio[:20]}"
                        break

            # Reconfigurar self.tree dinamicamente con todas las columnas detectadas
            self._reconfigurar_columnas_tree(headers_tree)

            # Poblar self.tree con todas las filas de la hoja Excel
            self._limpiar_tabla()
            total_fila = self._detectar_fila_total(ws, fila_hdr + 1)
            end_r = min(total_fila, (ws.max_row or 0) + 1)
            filas_cargadas = 0
            for r in range(fila_hdr + 1, end_r):
                row_vals = [str(ws.cell(r, c).value or "").strip() for c in range(1, len(headers_tree) + 1)]
                if not any(row_vals):
                    continue
                if any(v.upper().startswith("TOTAL") for v in row_vals[:3]):
                    break
                tag = "ok"
                for v in row_vals:
                    if "TRASLADO" in v.upper():
                        tag = "traslado"
                        break
                self.tree.insert("", "end", values=tuple(row_vals), tags=(tag,))
                filas_cargadas += 1

            self._actualizar_conteo()
            wb.close()

            # Poblar los dropdowns
            for combo in (self.combo_col_serial, self.combo_col_equipo, self.combo_col_sede, self.combo_col_resp, self.combo_col_verif):
                combo["values"] = opciones_cols

            if col_serial_det: self.col_serial_var.set(col_serial_det)
            if col_equipo_det: self.col_equipo_var.set(col_equipo_det)
            if col_sede_det:   self.col_sede_var.set(col_sede_det)
            if col_resp_det:   self.col_resp_var.set(col_resp_det)
            if col_verif_det:  self.col_verif_var.set(col_verif_det)

            self._actualizar_estado_combos()

            self.log(f"Tabla adaptada dinamicamente: {len(headers_tree)} columnas y {filas_cargadas} equipos cargados desde [{hoja}].", "ok")
            self.log(f"Encabezados detectados en Fila {fila_hdr} de [{hoja}]:", "info")
            self.log(f"   Serial: [{self.col_serial_var.get()}] | Host: [{self.col_equipo_var.get()}] | Sede: [{self.col_sede_var.get()}] | Resp: [{self.col_resp_var.get()}] | Auditoria: [{self.col_verif_var.get()}]", "ok")

        except Exception as e:
            self.log(f"Aviso al detectar columnas: {e}", "warn")

    def _get_col_idx(self, texto, default):
        """Extrae el indice numerico 1-based de un string como 'Col K (11) - Serial' o 'K' o '11'."""
        if not texto:
            return default
        m = re.search(r"\((\d+)\)", str(texto))
        if m:
            return int(m.group(1))
        m_col = re.search(r"Col\s+([A-Za-z]+)", str(texto), re.IGNORECASE)
        if m_col:
            return openpyxl.utils.column_index_from_string(m_col.group(1).upper())
        s = str(texto).strip()
        if s.isdigit():
            return int(s)
        if len(s) <= 3 and s.isalpha():
            return openpyxl.utils.column_index_from_string(s.upper())
        return default

    def _detectar_fila_encabezados(self, ws):
        """Detecta la fila de encabezados principales buscando coincidencias semanticas o la primera fila con encabezados de texto."""
        keywords = ["SERIAL", "EQUIPO", "HOST", "SEDE", "UBICACION", "ITEM", "CLIENTE", "RESPONSABLE", "OBSERVACION", "ESTADO", "ACTIVO", "MARCA", "MODELO", "NOMBRE", "DISPOSITIVO"]
        mejor_fila = None
        max_coincidencias = 0

        for r in range(1, min(15, (ws.max_row or 0) + 1)):
            valores = [str(ws.cell(r, c).value or "").strip().upper() for c in range(1, min(30, (ws.max_column or 0) + 1))]
            if not any(valores):
                continue
            if valores[0].isdigit():
                continue
            if any(v.startswith("SERIAL EQUIPO:") or v.startswith("MODELO:") for v in valores):
                continue
            coincidencias = sum(1 for v in valores if any(k == v or (k in v and len(v) < 35) for k in keywords))
            if coincidencias >= 3:
                return r
            if coincidencias > max_coincidencias:
                max_coincidencias = coincidencias
                mejor_fila = r

        if mejor_fila and max_coincidencias >= 2:
            return mejor_fila

        # Si fila 5 tiene texto (caso Clemont)
        r5_vals = [str(ws.cell(5, c).value or "").strip() for c in range(1, min(15, (ws.max_column or 0) + 1))]
        if sum(1 for v in r5_vals if v and not v.isdigit()) >= 3:
            return 5

        # Fallback para archivos genericos: primera fila con al menos 2 celdas de texto
        for r in range(1, min(10, (ws.max_row or 0) + 1)):
            c_textos = [str(ws.cell(r, c).value or "").strip() for c in range(1, min(25, (ws.max_column or 0) + 1))]
            textos_validos = [t for t in c_textos if t and not t.isdigit() and not t.upper().startswith("TOTAL")]
            if len(textos_validos) >= 2:
                return r

        return 1

    def _detectar_fila_total(self, ws, fila_inicio):
        """Busca la fila de TOTALES posterior a los datos."""
        for r in range(fila_inicio, (ws.max_row or 0) + 1):
            for c in range(1, min(6, (ws.max_column or 0) + 1)):
                val = str(ws.cell(r, c).value or "").upper()
                if "TOTAL" in val:
                    if not ws.cell(r, 1).value:
                        ws.cell(r, 1).value = val
                    return r
        return (ws.max_row or 0) + 1

    # ------ PEGADO DESDE PORTAPAPELES ------------------------------------

    def _cargar_preset_clemont(self):
        """Carga la lista de equipos verificados en auditoria fisica de Clemont."""
        cols_canon = ("#", "Equipo / Host", "Serial Real", "Estado", "Sede", "Responsable", "Observaciones")
        if list(self.tree["columns"]) != list(cols_canon):
            self._reconfigurar_columnas_tree(cols_canon)
        self._limpiar_tabla()
        for i, eq in enumerate(self.PRESET_CLEMONT, start=1):
            tag = "traslado" if eq[2] == "Traslado" else "ok"
            vals = (i, eq[0], eq[1], eq[2], eq[3], "Asignado", "")
            self.tree.insert("", "end", values=vals, tags=(tag,))
        self._actualizar_conteo()
        self.log(f"Preset Clemont cargado ({len(self.PRESET_CLEMONT)} equipos verificados).", "info")

    def _pegar_desde_portapapeles(self):
        """
        Lee el contenido del portapapeles (texto TSV que copia Excel)
        y lo agrega a la tabla Treeview de 7 columnas:
            # | Equipo | Serial | Estado | Sede | Responsable | Observaciones
        Detecta inteligentemente si se copio una tabla de 4 columnas, de 7 columnas,
        o la tabla completa de 10-12 columnas de inventario.
        """
        try:
            raw = self.root.clipboard_get()
        except tk.TclError:
            messagebox.showwarning(
                "Portapapeles vacio",
                "No hay nada en el portapapeles.\n"
                "Copia el rango en Excel primero (Ctrl+C)."
            )
            return

        lineas = raw.strip().splitlines()
        if not lineas:
            return

        agregados = 0
        omitidos  = 0

        idx_item   = None
        idx_equipo = None
        idx_serial = None
        idx_estado = None
        idx_sede   = None
        idx_resp   = None
        idx_obs    = None

        partes_primera = [p.strip().upper() for p in lineas[0].split("\t")]
        num_cols = len(partes_primera)
        inicio_linea = 0

        # Determinar si la primera linea es genuinamente un encabezado
        # 1. Si la primera celda es un numero (ej: '1', '20'), es fila de datos, NO encabezado
        # 2. Si alguna celda contiene 'SERIAL EQUIPO:' o 'MODELO:', es fila de datos, NO encabezado
        es_encabezado = False
        primera_celda = partes_primera[0] if partes_primera else ""
        contiene_datos_serial = any(p.startswith("SERIAL EQUIPO:") or p.startswith("MODELO:") for p in partes_primera)

        if not primera_celda.isdigit() and not contiene_datos_serial:
            keywords_header = [
                "ITEM", "#", "ID", "CLIENTE", "SEDE", "CENTRO OPERATIVO", "UBICACION",
                "EQUIPO", "SERVIDOR", "HOSTNAME", "SISTEMA OPERATIVO", "FABRICANTE",
                "TIPO LICENCIA", "RESPONSABLE", "USUARIO", "OBSERVACIONES Y SERIAL",
                "OBSERVACIONES", "VERIFICADO", "AUDITORIA"
            ]
            matches_hdr = 0
            for p in partes_primera:
                if any(kw == p or (kw in p and len(p) < 35) for kw in keywords_header):
                    matches_hdr += 1
            if matches_hdr >= 3:
                es_encabezado = True

        if es_encabezado:
            inicio_linea = 1
            for i, p in enumerate(partes_primera):
                if any(k in p for k in ["SERIAL", "S/N", "OBSERVACIONES Y SERIAL"]) and idx_serial is None:
                    idx_serial = i
                elif any(k in p for k in ["EQUIPO", "HOST", "DISPOSITIVO", "SERVIDOR"]) and idx_equipo is None:
                    idx_equipo = i
                elif any(k in p for k in ["SEDE", "UBICACION", "CIUDAD", "LUGAR", "CENTRO OPERATIVO"]) and idx_sede is None:
                    idx_sede = i
                elif any(k in p for k in ["RESPONSABLE", "USUARIO", "ENCARGADO", "EMPLEADO"]) and idx_resp is None:
                    idx_resp = i
                elif any(k in p for k in ["VERIFICADO", "AUDITORIA", "ESTADO"]) and idx_estado is None:
                    idx_estado = i
                elif any(k in p for k in ["OBSERVACION", "DETALLE", "COMENTARIO"]) and idx_obs is None:
                    idx_obs = i
                elif (p in ("ITEM", "#", "ID", "N°", "NO.", "ITEM #") or p.startswith("ITEM") or p.startswith("ID")) and "CANTIDAD" not in p and idx_item is None:
                    idx_item = i

            if idx_serial is None and idx_obs is not None:
                idx_serial = idx_obs
                idx_obs = None

            self.log(f"Encabezados detectados en portapapeles: Equipo=Col {(idx_equipo or 0)+1}, Serial=Col {(idx_serial or 0)+1}", "info")

        # Fallbacks estructurados segun el numero de columnas:
        if num_cols >= 10:
            if idx_item is None:   idx_item   = 0
            if idx_sede is None:   idx_sede   = 2
            if idx_equipo is None: idx_equipo = 3
            if idx_resp is None:   idx_resp   = 8
            if idx_estado is None: idx_estado = 11 if num_cols > 11 else 9
            if idx_serial is None: idx_serial = 10
            if idx_obs is None:    idx_obs    = 10
            if not es_encabezado:
                self.log(f"Detectada tabla completa ({num_cols} cols) en portapapeles. Mapeo automatico: Host=Col 4, Sede=Col 3, Serial=Col 11.", "info")

        elif num_cols == 7:
            if idx_item is None:   idx_item   = 0
            if idx_equipo is None: idx_equipo = 1
            if idx_serial is None: idx_serial = 2
            if idx_estado is None: idx_estado = 3
            if idx_sede is None:   idx_sede   = 4
            if idx_resp is None:   idx_resp   = 5
            if idx_obs is None:    idx_obs    = 6

        elif num_cols == 4:
            if idx_equipo is None: idx_equipo = 0
            if idx_serial is None: idx_serial = 1
            if idx_estado is None: idx_estado = 2
            if idx_sede is None:   idx_sede   = 3

        else:
            if idx_equipo is None: idx_equipo = 0
            if idx_serial is None: idx_serial = 1 if num_cols > 1 else 0
            if idx_estado is None: idx_estado = 2 if num_cols > 2 else None
            if idx_sede is None:   idx_sede   = 3 if num_cols > 3 else None

        cols_canon = ("#", "Equipo / Host", "Serial Real", "Estado", "Sede", "Responsable", "Observaciones")
        if len(self.tree["columns"]) != 7:
            self._reconfigurar_columnas_tree(cols_canon)

        conteo_actual = len(self.tree.get_children())
        for linea in lineas[inicio_linea:]:
            linea = linea.strip()
            if not linea:
                continue

            partes = linea.split("\t")
            if len(partes) < 2 and num_cols > 1:
                omitidos += 1
                continue

            equipo = partes[idx_equipo].strip() if idx_equipo is not None and len(partes) > idx_equipo else ""
            raw_serial = partes[idx_serial].strip() if idx_serial is not None and len(partes) > idx_serial else ""
            estado = partes[idx_estado].strip() if idx_estado is not None and len(partes) > idx_estado else "OK"
            sede   = partes[idx_sede].strip() if idx_sede is not None and len(partes) > idx_sede else ""
            resp   = partes[idx_resp].strip() if idx_resp is not None and len(partes) > idx_resp else ""
            obs    = partes[idx_obs].strip() if idx_obs is not None and len(partes) > idx_obs else ""
            item_n = partes[idx_item].strip() if idx_item is not None and len(partes) > idx_item else str(conteo_actual + agregados + 1)

            seriales_extraidos = extraer_seriales(raw_serial)
            serial = seriales_extraidos[0] if seriales_extraidos else raw_serial

            tag = "ok"
            if any("TRASLADO" in p.upper() for p in partes):
                tag = "traslado"

            if not equipo and not serial and not any(p.strip() for p in partes):
                omitidos += 1
                continue

            if normalizar(equipo) in ("ITEM", "EQUIPO", "EQUIPOSERVIDOR"):
                omitidos += 1
                continue

            est_norm = normalizar(estado)
            if est_norm in ("TRASLADO", "TRASL"):
                estado_canon = "Traslado"
                tag = "traslado"
            else:
                estado_canon = "OK"

            vals = (item_n, equipo, serial, estado_canon, sede, resp, obs)
            self.tree.insert("", "end", values=vals, tags=(tag,))
            agregados += 1

        self._actualizar_conteo()
        if agregados:
            self.log(
                f"Pegado exitoso: {agregados} equipos agregados"
                + (f" ({omitidos} lineas ignoradas)" if omitidos else ""),
                "ok"
            )
        else:
            messagebox.showwarning(
                "Sin datos validos",
                "No se encontraron filas validas en el portapapeles.\n\n"
                "Asegurate de copiar los datos en Excel (Ctrl+C) antes de pegar."
            )

    # ------ CRUD de la tabla ---------------------------------------------

    def _agregar_fila(self):
        self._dialogo_fila()

    def _borrar_fila(self):
        for item in self.tree.selection():
            self.tree.delete(item)
        self._actualizar_conteo()

    def _limpiar_tabla(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._actualizar_conteo()

    def _editar_fila(self, event):
        sel = self.tree.selection()
        if sel:
            item = sel[0]
            self._dialogo_fila(item, self.tree.item(item, "values"))

    def _dialogo_fila(self, item_id=None, valores=None):
        dlg = tk.Toplevel(self.root)
        dlg.title("Agregar / Editar registro")
        dlg.geometry("520x420")
        dlg.grab_set()

        cols_tree = list(self.tree["columns"]) if hasattr(self, "tree") and self.tree else ["#", "Equipo / Host", "Serial Real", "Estado", "Sede", "Responsable", "Observaciones"]

        canvas = tk.Canvas(dlg, borderwidth=0, highlightthickness=0)
        vsb_dlg = ttk.Scrollbar(dlg, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=vsb_dlg.set)

        vsb_dlg.pack(side="right", fill="y")
        canvas.pack(side="top", fill="both", expand=True, padx=10, pady=10)

        entradas = {}
        for idx, col_nombre in enumerate(cols_tree):
            frm = tk.Frame(scrollable_frame)
            frm.pack(fill="x", padx=10, pady=3)
            tk.Label(frm, text=str(col_nombre)[:22], width=24, anchor="w", font=("Arial", 9, "bold")).pack(side="left")
            val_def = ""
            if valores and idx < len(valores):
                val_def = valores[idx]
            elif idx == 0 and not valores:
                val_def = str(len(self.tree.get_children()) + 1)
            var = tk.StringVar(value=str(val_def))
            tk.Entry(frm, textvariable=var, width=32, font=("Consolas", 9)).pack(side="left")
            entradas[idx] = var

        def guardar():
            vals = tuple(entradas[i].get().strip() for i in range(len(cols_tree)))
            tag = "ok"
            for v in vals:
                if "TRASLADO" in v.upper():
                    tag = "traslado"
                    break
            if item_id:
                self.tree.item(item_id, values=vals, tags=(tag,))
            else:
                self.tree.insert("", "end", values=vals, tags=(tag,))
            self._actualizar_conteo()
            dlg.destroy()

        btn_frm = tk.Frame(dlg)
        btn_frm.pack(side="bottom", fill="x", pady=6)
        tk.Button(btn_frm, text="Guardar", command=guardar,
                  bg="#28A745", fg="white",
                  font=("Arial", 10, "bold"), padx=15).pack(anchor="center")

    def _actualizar_conteo(self):
        n = len(self.tree.get_children())
        self.lbl_conteo.config(text=f"{n} equipo(s) cargado(s)")
        self._qr_actualizar_resumen_fuente()

    # ------------------------------------------------------------------
    # Logging
    # ------------------------------------------------------------------

    def log(self, msg, tag="info"):
        ts = datetime.now().strftime("[%H:%M:%S] ")
        self.txt_log.insert(tk.END, ts + str(msg) + "\n", tag)
        self.txt_log.see(tk.END)
        self.root.update_idletasks()

    # ------------------------------------------------------------------
    # Motor de conciliacion
    # ------------------------------------------------------------------

    def _ejecutar(self):
        ruta = self.ruta_excel.get()
        if not ruta or not os.path.exists(ruta):
            messagebox.showerror("Error", "Selecciona un archivo Excel valido.")
            return

        nombre_hoja = self.nombre_hoja.get().strip()
        idx_ser_c = self._get_col_idx(self.col_serial_var.get(), 11) - 1
        idx_eq_c  = self._get_col_idx(self.col_equipo_var.get(), 4) - 1
        idx_sd_c  = self._get_col_idx(self.col_sede_var.get(), 3) - 1
        idx_rp_c  = self._get_col_idx(self.col_resp_var.get(), 9) - 1
        idx_vf_c  = self._get_col_idx(self.col_verif_var.get(), 12) - 1

        lista_fisica = []
        for i in self.tree.get_children():
            vals = self.tree.item(i, "values")
            if len(vals) >= 10:
                campos = extraer_campos_completos(vals)
                item_val = (str(vals[0]) if len(vals) > 0 and str(vals[0]).strip() else "") or campos.get("id_activo", "")
                eq_val   = (str(vals[idx_eq_c]) if 0 <= idx_eq_c < len(vals) and str(vals[idx_eq_c]).strip() else "") or campos.get("equipo", "")
                raw_ser  = (str(vals[idx_ser_c]) if 0 <= idx_ser_c < len(vals) and str(vals[idx_ser_c]).strip() else "") or campos.get("serial", "")
                sd_val   = (str(vals[idx_sd_c]) if 0 <= idx_sd_c < len(vals) and str(vals[idx_sd_c]).strip() else "") or campos.get("sede", "")
                rp_val   = (str(vals[idx_rp_c]) if 0 <= idx_rp_c < len(vals) and str(vals[idx_rp_c]).strip() else "") or campos.get("responsable", "")
                st_val   = (str(vals[idx_vf_c]) if 0 <= idx_vf_c < len(vals) and str(vals[idx_vf_c]).strip() else "") or campos.get("estado", "OK")
                sns = extraer_seriales(raw_ser)
                sn_val = sns[0] if sns else campos.get("serial", raw_ser)
                ob_val = raw_ser or campos.get("modelo", "")
            elif len(vals) >= 7:
                item_val = vals[0]
                eq_val   = vals[1]
                sn_val   = vals[2]
                st_val   = vals[3]
                sd_val   = vals[4]
                rp_val   = vals[5]
                ob_val   = vals[6]
            elif len(vals) == 4:
                item_val = ""
                eq_val   = vals[0]
                sn_val   = vals[1]
                st_val   = vals[2]
                sd_val   = vals[3]
                rp_val   = ""
                ob_val   = ""
            else:
                item_val = vals[0] if len(vals) > 0 else ""
                eq_val   = vals[1] if len(vals) > 1 else ""
                sn_val   = vals[2] if len(vals) > 2 else ""
                st_val   = vals[3] if len(vals) > 3 else "OK"
                sd_val   = vals[4] if len(vals) > 4 else ""
                rp_val   = vals[5] if len(vals) > 5 else ""
                ob_val   = vals[6] if len(vals) > 6 else ""

            lista_fisica.append({
                "item":   item_val,
                "equipo": eq_val,
                "serial": sn_val,
                "estado": st_val,
                "sede":   sd_val,
                "resp":   rp_val,
                "obs":    ob_val,
            })

        if not lista_fisica:
            messagebox.showwarning(
                "Sin datos",
                "Pega o agrega equipos verificados antes de ejecutar."
            )
            return

        self.txt_log.delete("1.0", tk.END)
        self.log("=" * 58, "info")
        self.log(f"  INICIANDO CONCILIACION  -  "
                 f"{datetime.now().strftime('%Y-%m-%d %H:%M')}", "info")
        self.log("=" * 58, "info")

        try:
            wb = openpyxl.load_workbook(ruta)
        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Cierra el archivo en Excel primero.\n{e}"
            )
            return

        # Seleccionar hoja
        if nombre_hoja not in wb.sheetnames:
            candidatas = [h for h in wb.sheetnames if "INVENTARIO" in h.upper()]
            nombre_hoja = candidatas[0] if candidatas else wb.sheetnames[0]
            self.log(f"Hoja no encontrada, usando: [{nombre_hoja}]", "warn")

        ws = wb[nombre_hoja]
        self.log(f"Hoja objetivo: [{nombre_hoja}]  "
                 f"({ws.max_row} filas x {ws.max_column} cols)", "info")

        # Casillas de verificacion (chulos) y columnas mapeadas
        usar_serial = self.chk_col_serial.get()
        usar_equipo = self.chk_col_equipo.get()
        usar_sede   = self.chk_col_sede.get()
        usar_resp   = self.chk_col_resp.get()
        usar_verif  = self.chk_col_verif.get()

        col_sede   = self._get_col_idx(self.col_sede_var.get(), default=self.COL_SEDE)
        col_equipo = self._get_col_idx(self.col_equipo_var.get(), default=self.COL_EQUIPO)
        col_obs    = self._get_col_idx(self.col_serial_var.get(), default=self.COL_OBS)
        col_resp   = self._get_col_idx(self.col_resp_var.get(), default=self.COL_RESP) if hasattr(self, "col_resp_var") else 9
        col_verif  = self._get_col_idx(self.col_verif_var.get(), default=self.COL_VERIF)

        # Detectar fila de encabezados y de datos
        hdr_row = self._detectar_fila_encabezados(ws)
        data_start = hdr_row + 1
        total_row = self._detectar_fila_total(ws, data_start)

        self.log(f"Encabezados en F{hdr_row} | Datos desde F{data_start} hasta F{total_row - 1}", "info")
        cols_msg = []
        if usar_serial: cols_msg.append(f"Serial=Col {get_column_letter(col_obs)}")
        if usar_equipo: cols_msg.append(f"Host=Col {get_column_letter(col_equipo)}")
        if usar_sede:   cols_msg.append(f"Sede=Col {get_column_letter(col_sede)}")
        if usar_resp:   cols_msg.append(f"Resp=Col {get_column_letter(col_resp)}")
        if usar_verif:  cols_msg.append(f"Auditoria=Col {get_column_letter(col_verif)}")
        self.log(f"Columnas activas: {' | '.join(cols_msg)}", "ok")

        # --- Detectar duplicados en el Excel ----------------------------
        seriales_vistos = {}   # sn -> (fila, texto_serial_original)
        nombres_vistos  = {}   # clave -> (fila, equipo_original)
        filas_dup       = {}   # r -> motivo del duplicado

        for r in range(data_start, total_row):
            obs    = str(ws.cell(r, col_obs).value or "")
            equipo = normalizar(str(ws.cell(r, col_equipo).value or ""))
            sede   = normalizar(str(ws.cell(r, col_sede).value or ""))
            nom_orig = str(ws.cell(r, col_equipo).value or "")

            for s in extraer_seriales(obs):
                sn = normalizar_tolerante(s)
                if len(sn) < 6:
                    continue
                if sn in seriales_vistos:
                    fila_ant, s_ant = seriales_vistos[sn]
                    motivo_actual = f"Serial repetido '{s}': coincide con fila {fila_ant}"
                    motivo_anterior = f"Serial repetido '{s}': coincide con fila {r}"
                    filas_dup[r] = motivo_actual
                    if fila_ant not in filas_dup:
                        filas_dup[fila_ant] = motivo_anterior
                    self.log(
                        f"DUPLICADO serial '{s}' -> F{r} y F{fila_ant}", "warn"
                    )
                else:
                    seriales_vistos[sn] = (r, s)

            clave = f"{sede}_{equipo}"
            if equipo and clave in nombres_vistos:
                fila_ant, n_ant = nombres_vistos[clave]
                motivo_actual = f"Equipo repetido '{nom_orig}' en misma sede ({ws.cell(r, col_sede).value}): coincide con fila {fila_ant}"
                motivo_anterior = f"Equipo repetido '{nom_orig}' en misma sede: coincide con fila {r}"
                if r not in filas_dup:
                    filas_dup[r] = motivo_actual
                if fila_ant not in filas_dup:
                    filas_dup[fila_ant] = motivo_anterior
                self.log(
                    f"DUPLICADO nombre '{nom_orig}' -> F{r} y F{fila_ant}", "warn"
                )
            elif equipo:
                nombres_vistos[clave] = (r, nom_orig)

        # --- Conciliacion fila por fila ---------------------------------
        procesados_fisica = set()
        contadores = {"ok": 0, "traslado": 0, "sin_chequear": 0, "dup": 0}

        for r in range(data_start, total_row):
            obs_celda    = str(ws.cell(r, col_obs).value or "")
            equipo_celda = str(ws.cell(r, col_equipo).value or "")

            seriales_celda = extraer_seriales(obs_celda) or [obs_celda]

            coincidencia = None
            idx_match    = None

            for idx_f, eq in enumerate(lista_fisica):
                if idx_f in procesados_fisica:
                    continue
                # Match por serial (si serial esta habilitado con chulo)
                if usar_serial and eq["serial"]:
                    for s_celda in seriales_celda:
                        if similitud_serial(eq["serial"], s_celda):
                            coincidencia = eq
                            idx_match    = idx_f
                            break
                # Match por nombre de equipo (si equipo esta habilitado con chulo)
                if not coincidencia and usar_equipo and eq["equipo"]:
                    eq_n  = normalizar(eq["equipo"])
                    cel_n = normalizar(equipo_celda)
                    if eq_n and cel_n and (eq_n == cel_n or
                       eq_n in cel_n or cel_n in eq_n):
                        coincidencia = eq
                        idx_match    = idx_f
                if coincidencia:
                    break

            if coincidencia:
                procesados_fisica.add(idx_match)
                estado = coincidencia["estado"]

                # Actualizar nombre de equipo (si activo)
                if usar_equipo and coincidencia["equipo"]:
                    if not (str(coincidencia["equipo"]).strip().isdigit() and len(str(coincidencia["equipo"]).strip()) <= 3):
                        ws.cell(r, col_equipo).value = coincidencia["equipo"]

                # Actualizar serial corregido (si activo)
                if usar_serial and coincidencia["serial"]:
                    texto_obs = str(ws.cell(r, col_obs).value or "")
                    serial_val = str(coincidencia["serial"]).strip()
                    if serial_val.upper() != "CLEMONT" and not (serial_val.isdigit() and len(serial_val) <= 3):
                        serial_ok = f"Serial equipo: {serial_val}"
                        m_modelo  = re.search(r"(?i)(Modelo\s*:.+)", texto_obs)
                        if m_modelo:
                            serial_ok += f"; {m_modelo.group(1).strip()}"
                        ws.cell(r, col_obs).value = serial_ok

                # Actualizar responsable (si activo y presente)
                if usar_resp and coincidencia.get("resp"):
                    ws.cell(r, col_resp).value = coincidencia["resp"]

                # Marcar en columna de resultado (si activo)
                if usar_verif:
                    ws.cell(r, col_verif).value = estado
                    if normalizar(estado) in ("OK", "VERIFICADO"):
                        ws.cell(r, col_verif).font  = FONT_CHECK
                        ws.cell(r, col_verif).fill  = FILL_OK
                        contadores["ok"] += 1
                    else:
                        ws.cell(r, col_verif).font  = FONT_TRASL
                        ws.cell(r, col_verif).fill  = FILL_TRASL
                        contadores["traslado"] += 1

                    ws.cell(r, col_verif).alignment = Alignment(
                        horizontal="center")
                else:
                    if normalizar(estado) in ("OK", "VERIFICADO"):
                        contadores["ok"] += 1
                    else:
                        contadores["traslado"] += 1

                self.log(
                    f"[{estado}] F{r} -> {coincidencia['equipo']} | "
                    f"{coincidencia['serial']}", "ok"
                )
            else:
                has_data = any(str(ws.cell(r, c).value or "").strip() for c in (col_equipo, col_obs, col_sede, 1, 2, 3))
                if has_data:
                    if r in filas_dup:
                        if usar_verif:
                            ws.cell(r, col_verif).value = "Duplicado"
                            ws.cell(r, col_verif).font  = Font(
                                name="Calibri", size=10, bold=True, color="7F6000")
                            ws.cell(r, col_verif).fill  = FILL_DUP
                            # Insertar comentario / nota explicativa en la celda
                            nota = filas_dup[r]
                            ws.cell(r, col_verif).comment = Comment(
                                f"Auditoría: {nota}", "Auditor TI"
                            )
                        contadores["dup"] += 1
                    else:
                        if usar_verif:
                            ws.cell(r, col_verif).value = "Sin chequear"
                            ws.cell(r, col_verif).font  = FONT_UNCHK
                            ws.cell(r, col_verif).fill  = FILL_UNCHK
                        contadores["sin_chequear"] += 1

        # --- Altas: equipos fisicos NO en el Excel ----------------------
        altas = [lista_fisica[i] for i in range(len(lista_fisica))
                 if i not in procesados_fisica]
        for alta in altas:
            self.log(
                f"ALTA (no estaba en Excel): {alta['equipo']} | "
                f"{alta['serial']} | {alta['sede']}", "warn"
            )

        # --- Autofilter y freeze panes ----------------------------------
        ws.auto_filter.ref = (
            f"A{hdr_row}:{get_column_letter(max(col_verif, ws.max_column))}{hdr_row}"
        )
        ws.freeze_panes = f"A{data_start}"

        # --- Deteccion de formato Clemont vs generico y columna cantidad ---
        es_formato_clemont = (ws.max_column >= 10 and any("CLIENTE" in str(ws.cell(hdr_row, c).value or "").upper() for c in range(1, min(5, ws.max_column + 1))))
        col_cant = None
        for c in range(1, (ws.max_column or 0) + 1):
            val_h = str(ws.cell(hdr_row, c).value or "").upper()
            if "CANT" in val_h:
                col_cant = c
                break
        if not col_cant and es_formato_clemont:
            col_cant = 8

        # --- Corregir formula TOTAL (si existe columna cantidad) ---
        if col_cant and total_row <= ws.max_row:
            letra_cant = get_column_letter(col_cant)
            ws.cell(total_row, col_cant).value = f"=SUM({letra_cant}{data_start}:{letra_cant}{total_row - 1})"

        # --- Insertar altas antes del TOTAL -----------------------------
        for idx_a, alta in enumerate(altas):
            insert_r = total_row + idx_a
            ws.insert_rows(insert_r)
            ws.cell(insert_r, 1).value = alta["item"] if alta["item"] else (insert_r - hdr_row)
            if es_formato_clemont:
                ws.cell(insert_r, 2).value                   = "CLEMONT"
                if usar_sede:
                    ws.cell(insert_r, col_sede).value        = alta["sede"]
                if usar_equipo:
                    ws.cell(insert_r, col_equipo).value      = alta["equipo"]
                ws.cell(insert_r, 5).value                   = "WINDOWS 11 PRO"
                ws.cell(insert_r, 6).value                   = "Microsoft"
                ws.cell(insert_r, 7).value                   = "Sistema operativo"
                ws.cell(insert_r, 8).value                   = 1
                if usar_resp and alta.get("resp"):
                    ws.cell(insert_r, col_resp).value        = alta["resp"]
                ws.cell(insert_r, 10).value                  = "Activo"
                if usar_serial:
                    ws.cell(insert_r, col_obs).value         = f"Serial equipo: {alta['serial']}"
                if usar_verif:
                    ws.cell(insert_r, col_verif).value       = alta["estado"]
            else:
                if usar_equipo and col_equipo:
                    ws.cell(insert_r, col_equipo).value      = alta["equipo"]
                if usar_serial and col_obs:
                    ws.cell(insert_r, col_obs).value         = f"Serial equipo: {alta['serial']}"
                if usar_sede and col_sede:
                    ws.cell(insert_r, col_sede).value        = alta["sede"]
                if usar_resp and col_resp and alta.get("resp"):
                    ws.cell(insert_r, col_resp).value        = alta["resp"]
                if col_cant:
                    ws.cell(insert_r, col_cant).value        = 1
                if usar_verif and col_verif:
                    ws.cell(insert_r, col_verif).value       = alta["estado"]

            for c in range(1, max(col_verif, ws.max_column) + 1):
                ws.cell(insert_r, c).font   = FONT_BASE
                ws.cell(insert_r, c).border = THIN_BORDER
                ws.cell(insert_r, c).fill   = FILL_ALTA
            if usar_verif:
                ws.cell(insert_r, col_verif).font = FONT_CHECK
                ws.cell(insert_r, col_verif).alignment = Alignment(horizontal="center")
            self.log(f"Insertado alta: {alta['equipo']} en F{insert_r}", "ok")

        nuevo_total_row = total_row + len(altas)
        if col_cant and total_row <= ws.max_row:
            letra_cant = get_column_letter(col_cant)
            ws.cell(nuevo_total_row, col_cant).value = (
                f"=SUM({letra_cant}{data_start}:{letra_cant}{nuevo_total_row - 1})"
            )

        # --- Tablas de perifericos (Opcional demo Clemont) --------------
        if self.chk_perifericos_clemont.get():
            tab_start = nuevo_total_row + 2
            self._escribir_perifericos(
                ws, tab_start,
                titulo="TABLETS Y DISPOSITIVOS MOVILES VERIFICADOS",
                color_hdr="7030A0",
                datos=self.PRESET_TABLETS
            )
            mon_start = tab_start + len(self.PRESET_TABLETS) + 4
            self._escribir_perifericos(
                ws, mon_start,
                titulo="MONITORES Y PANTALLAS VERIFICADAS (PERIFERICOS)",
                color_hdr="203764",
                datos=self.PRESET_MONITORES
            )

        # --- Hoja Reporte de Auditoria ----------------------------------
        self._generar_reporte(
            wb, ws, nombre_hoja,
            total_row, len(altas), lista_fisica, altas,
            contadores, filas_dup,
            col_equipo=col_equipo, col_obs=col_obs,
            col_verif=col_verif, col_sede=col_sede,
            data_start=data_start
        )

        # --- Guardar ----------------------------------------------------
        ruta_salida_custom = self.ruta_salida.get().strip() if hasattr(self, "ruta_salida") else ""
        if ruta_salida_custom:
            ruta_guardado = ruta_salida_custom
            if not ruta_guardado.lower().endswith(".xlsx"):
                ruta_guardado += ".xlsx"
            nom_salida = os.path.basename(ruta_guardado)
        else:
            dir_s, nom_orig = os.path.split(ruta)
            nom_salida = (
                "AUDITADO_" + nom_orig
                if not nom_orig.startswith("AUDITADO_")
                else nom_orig
            )
            ruta_guardado = os.path.join(dir_s, nom_salida)

        self.ruta_salida.set(ruta_guardado)
        wb.save(ruta_guardado)
        try:
            wb.close()
        except Exception:
            pass

        self.log("=" * 58, "info")
        self.log("  CONCILIACION COMPLETADA", "ok")
        self.log(f"  Verificados (OK)  : {contadores['ok']}", "ok")
        self.log(f"  Traslados         : {contadores['traslado']}", "warn")
        self.log(f"  Sin chequear      : {contadores['sin_chequear']}", "info")
        self.log(f"  Duplicados        : {contadores['dup']}", "warn")
        self.log(f"  Altas (nuevos)    : {len(altas)}", "warn")
        self.log(f"  Archivo guardado  : {ruta_guardado}", "info")
        self.log("=" * 58, "info")

        messagebox.showinfo(
            "Conciliacion completada",
            f"Verificados (OK) : {contadores['ok']}\n"
            f"Traslados        : {contadores['traslado']}\n"
            f"Sin chequear     : {contadores['sin_chequear']}\n"
            f"Duplicados       : {contadores['dup']}\n"
            f"Altas nuevas     : {len(altas)}\n\n"
            f"Guardado exitosamente en:\n{ruta_guardado}\n\n"
            f"Puedes abrirlo directamente con los botones 'Abrir Excel Auditado' o 'Abrir Carpeta'."
        )

    # ------------------------------------------------------------------
    # Helpers de escritura en Excel
    # ------------------------------------------------------------------

    def _escribir_perifericos(self, ws, start_row, titulo, color_hdr, datos):
        c_tit      = ws.cell(start_row, 2)
        c_tit.value = titulo
        c_tit.font  = Font(name="Calibri", size=11, bold=True,
                           color=color_hdr)

        hdrs     = ["Item", "Cliente", "Sede", "Tipo", "Marca / Modelo",
                    "Modelo Det.", "Tipo", "Cant", "Responsable",
                    "Estado", "Serial de Fabrica", "Verificado"]
        fill_hdr = PatternFill(start_color=color_hdr, end_color=color_hdr,
                               fill_type="solid")
        for c, h in enumerate(hdrs, 1):
            cell        = ws.cell(start_row + 1, c)
            cell.value  = h
            cell.font   = FONT_HDR
            cell.fill   = fill_hdr
            cell.border = THIN_BORDER

        for i, fila in enumerate(datos, start=1):
            r = start_row + 1 + i
            for c, v in enumerate(fila, 1):
                cell        = ws.cell(r, c)
                cell.value  = v
                cell.font   = FONT_BASE
                cell.border = THIN_BORDER
                if c == 12:
                    cell.font = FONT_CHECK if normalizar(str(v)) == "OK" \
                        else FONT_TRASL
                    cell.alignment = Alignment(horizontal="center")

        total_r         = ws.cell(start_row + 1 + len(datos) + 1, 2)
        total_r.value   = f"TOTAL"
        total_r.font    = FONT_BOLD
        ws.cell(start_row + 1 + len(datos) + 1, 8).value = len(datos)
        ws.cell(start_row + 1 + len(datos) + 1, 8).font  = FONT_BOLD

    def _generar_reporte(self, wb, ws_inv, nombre_hoja,
                         total_row, n_altas, lista_fisica, altas,
                         contadores, filas_dup,
                         col_equipo=None, col_obs=None,
                         col_verif=None, col_sede=None,
                         data_start=6):
        col_equipo = col_equipo or self.COL_EQUIPO
        col_obs    = col_obs or self.COL_OBS
        col_verif  = col_verif or self.COL_VERIF
        col_sede   = col_sede or self.COL_SEDE

        nombre_rep = "Reporte de Auditoria"
        if nombre_rep in wb.sheetnames:
            wb.remove(wb[nombre_rep])
        ws = wb.create_sheet(title=nombre_rep)
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M")

        # Titulo
        ws["A1"] = f"REPORTE DE AUDITORIA Y CONCILIACION FISICA - {fecha}"
        ws["A1"].font = Font(name="Calibri", size=13, bold=True,
                             color="1F4E78")
        ws["B2"] = f"REPORTE DE AUDITORIA Y CONCILIACION FISICA - {fecha}"
        ws["B2"].font = Font(name="Calibri", size=13, bold=True,
                             color="1F4E78")

        # Resumen
        ws["B4"] = "RESUMEN"
        ws["B4"].font = FONT_BOLD
        resumen = [
            ("Equipos verificados (OK)",     contadores["ok"]),
            ("Equipos en traslado",          contadores["traslado"]),
            ("Equipos sin chequear",         contadores["sin_chequear"]),
            ("Filas duplicadas detectadas",  contadores["dup"]),
            ("Altas nuevas (no en Excel)",   n_altas),
        ]
        for i, (lbl, val) in enumerate(resumen, start=5):
            ws.cell(i, 2).value = lbl
            ws.cell(i, 3).value = val
            ws.cell(i, 2).font  = FONT_BASE
            ws.cell(i, 3).font  = FONT_BOLD

        # Tabla principal
        hdr_row = 11
        hdrs = ["#", "Equipo/Host", "Serial en Excel",
                "Serial Real Verificado", "Estado Auditoria",
                "Sede", "Novedad"]
        for c, h in enumerate(hdrs, 1):
            cell = ws.cell(hdr_row, c)
            cell.value = h
            cell.font  = FONT_HDR
            cell.fill  = HDR_BLUE
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center")

        item_rep = 1
        for r in range(data_start, total_row + n_altas):
            equipo = str(ws_inv.cell(r, col_equipo).value or "")
            serial = str(ws_inv.cell(r, col_obs).value or "")
            estado = str(ws_inv.cell(r, col_verif).value or "")
            sede   = str(ws_inv.cell(r, col_sede).value or "")
            if not equipo and not serial:
                continue

            serial_real = ""
            for eq in lista_fisica:
                if (similitud_serial(eq["serial"], serial)
                        or normalizar(eq["equipo"]) == normalizar(equipo)):
                    serial_real = eq["serial"]
                    break

            est_n = normalizar(estado)
            if est_n == "OK":
                novedad  = "Verificado OK"
                fill_row = FILL_OK
            elif est_n == "TRASLADO":
                novedad  = "Novedad de Traslado"
                fill_row = FILL_TRASL
            elif est_n == "DUPLICADO":
                motivo = filas_dup.get(r, "Fila duplicada en inventario")
                novedad  = f"Duplicado: {motivo}"
                fill_row = FILL_DUP
            elif "SIN" in est_n:
                novedad  = "Pendiente de verificacion"
                fill_row = FILL_UNCHK
            else:
                novedad  = "Sin clasificar"
                fill_row = None

            datos_fila = [item_rep, equipo, serial, serial_real,
                          estado, sede, novedad]
            rr = hdr_row + item_rep
            for c, v in enumerate(datos_fila, 1):
                cell        = ws.cell(rr, c)
                cell.value  = v
                cell.font   = FONT_BASE
                cell.border = THIN_BORDER
                if fill_row:
                    cell.fill = fill_row
            item_rep += 1

        # Tabla de altas
        if altas:
            alta_start = hdr_row + item_rep + 2
            tit_cell = ws.cell(alta_start, 1)
            tit_cell.value = ("EQUIPOS EN LISTA FISICA NO ENCONTRADOS EN "
                              "EXCEL (A REGISTRAR)")
            tit_cell.font  = Font(name="Calibri", size=11, bold=True,
                                  color="C00000")
            alta_hdrs = ["#", "Equipo/Host", "Serial Real", "Estado", "Sede"]
            for c, h in enumerate(alta_hdrs, 1):
                cell        = ws.cell(alta_start + 1, c)
                cell.value  = h
                cell.font   = FONT_HDR
                cell.fill   = HDR_GREEN
                cell.border = THIN_BORDER
            for i_a, alta in enumerate(altas, 1):
                rr = alta_start + 1 + i_a
                for c, v in enumerate(
                    [i_a, alta["equipo"], alta["serial"],
                     alta["estado"], alta["sede"]], 1
                ):
                    cell        = ws.cell(rr, c)
                    cell.value  = v
                    cell.font   = FONT_BASE
                    cell.border = THIN_BORDER
                    cell.fill   = FILL_ALTA

        # Autofilter en la tabla principal
        ws.auto_filter.ref = (
            f"A{hdr_row}:"
            f"{get_column_letter(len(hdrs))}{hdr_row}"
        )
        # Anchos
        for i, ancho in enumerate([5, 25, 38, 25, 18, 25, 30], 1):
            ws.column_dimensions[get_column_letter(i)].width = ancho

        self.log("Hoja 'Reporte de Auditoria' generada.", "ok")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    root = tk.Tk()
    ConciliadorUniversal(root)
    root.mainloop()
