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
except ImportError:
    QR_DISPONIBLE = False

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
    Separa por '/', ',', salto de linea o doble espacio.
    """
    if not texto:
        return []
    partes = re.split(r"[/,\n]|\s{2,}", str(texto))
    seriales = []
    for p in partes:
        p = re.sub(r"(?i)serial\s*equipo\s*[:\-]?\s*", "", p)
        p = re.sub(r"(?i)modelo\s*[:\-]?\s*.*", "", p).strip()
        norm = normalizar(p)
        if len(norm) >= 6:
            seriales.append(p.strip())
    return seriales


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
        self.root.title("Conciliador de Inventario Universal v3.1")
        self.root.geometry("1040x840")
        self.root.configure(bg="#F0F4F8")
        self.ruta_excel  = tk.StringVar()
        self.nombre_hoja = tk.StringVar()
        self.hojas_disponibles = []          # lista de hojas del Excel cargado

        # Mapeo dinamico de columnas (Universal para cualquier cliente)
        self.col_serial_var = tk.StringVar(value="Col K (11) - Observaciones y Serial")
        self.col_equipo_var = tk.StringVar(value="Col D (4) - Equipo/Servidor")
        self.col_sede_var   = tk.StringVar(value="Col C (3) - Sede")
        self.col_verif_var  = tk.StringVar(value="Col L (12) - Verificado")
        self.chk_perifericos_clemont = tk.BooleanVar(value=True)

        # Configuracion de Modulo QR
        self.qr_fuente = tk.StringVar(value="auto")
        self.qr_empresa = tk.StringVar(value="CLEMONT")
        self.qr_formato = tk.StringVar(value="50x25mm")
        self.qr_scan_input = tk.StringVar()
        self._construir_gui()

    # ------------------------------------------------------------------
    # GUI
    # ------------------------------------------------------------------

    def _construir_gui(self):
        # Titulo
        tk.Label(
            self.root,
            text="  Conciliador de Inventario - Clemont  v3.0",
            font=("Arial", 13, "bold"), bg="#1F4E78", fg="white", pady=10
        ).pack(fill="x")

        # ==============================================================
        # Notebook con pestanas
        # ==============================================================
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=8)

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
        # --- Bloque 1: archivo -------------------------------------------
        frm1 = tk.LabelFrame(parent, text="1. Archivo Excel",
                              padx=10, pady=8, bg="#F0F4F8",
                              font=("Arial", 10, "bold"))
        frm1.pack(fill="x", padx=15, pady=(10, 2))

        tk.Entry(frm1, textvariable=self.ruta_excel, width=65,
                 state="readonly", font=("Consolas", 9)).pack(
            side="left", fill="x", expand=True, padx=(0, 8))
        tk.Button(frm1, text="Examinar...", command=self._seleccionar_archivo,
                  bg="#1F4E78", fg="white",
                  font=("Arial", 9, "bold")).pack(side="left")

        frm1b = tk.Frame(parent, bg="#F0F4F8")
        frm1b.pack(fill="x", padx=15, pady=(2, 6))
        tk.Label(frm1b, text="Hoja a auditar:",
                 bg="#F0F4F8", font=("Arial", 9, "bold")).pack(side="left")

        # Dropdown (Combobox) — se llena automaticamente al cargar el Excel
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

        # Etiqueta de ayuda dinamica
        self.lbl_hoja_info = tk.Label(
            frm1b, text="", bg="#F0F4F8",
            font=("Arial", 8), fg="#555555"
        )
        self.lbl_hoja_info.pack(side="left", padx=4)

        # --- Mapeo dinamico de columnas (Universal para cualquier cliente) --
        frm1c = tk.Frame(parent, bg="#F0F4F8")
        frm1c.pack(fill="x", padx=15, pady=(0, 6))

        tk.Label(frm1c, text="Mapeo de Columnas:", bg="#F0F4F8",
                 font=("Arial", 9, "bold")).pack(side="left")

        tk.Label(frm1c, text="Serial:", bg="#F0F4F8", font=("Arial", 8)).pack(side="left", padx=(8, 2))
        self.combo_col_serial = ttk.Combobox(frm1c, textvariable=self.col_serial_var, width=17, font=("Consolas", 8))
        self.combo_col_serial.pack(side="left")

        tk.Label(frm1c, text="Host/Equipo:", bg="#F0F4F8", font=("Arial", 8)).pack(side="left", padx=(8, 2))
        self.combo_col_equipo = ttk.Combobox(frm1c, textvariable=self.col_equipo_var, width=17, font=("Consolas", 8))
        self.combo_col_equipo.pack(side="left")

        tk.Label(frm1c, text="Sede:", bg="#F0F4F8", font=("Arial", 8)).pack(side="left", padx=(8, 2))
        self.combo_col_sede = ttk.Combobox(frm1c, textvariable=self.col_sede_var, width=14, font=("Consolas", 8))
        self.combo_col_sede.pack(side="left")

        tk.Label(frm1c, text="Resultado:", bg="#F0F4F8", font=("Arial", 8)).pack(side="left", padx=(8, 2))
        self.combo_col_verif = ttk.Combobox(frm1c, textvariable=self.col_verif_var, width=15, font=("Consolas", 8))
        self.combo_col_verif.pack(side="left")

        tk.Checkbutton(frm1c, text="Anexar Tablets/Monitores Clemont",
                       variable=self.chk_perifericos_clemont,
                       bg="#F0F4F8", font=("Arial", 8)).pack(side="right")

        # --- Bloque 2: tabla de equipos verificados ----------------------
        frm2 = tk.LabelFrame(
            parent,
            text="2. Lista de Equipos Verificados Fisicamente  "
                 "(Equipo | Serial Real | Estado | Sede)",
            padx=10, pady=6, bg="#F0F4F8", font=("Arial", 10, "bold")
        )
        frm2.pack(fill="both", expand=True, padx=15, pady=2)

        # Instruccion de pegado
        instruccion = (
            "Como pegar desde Excel:  Selecciona el rango en tu hoja de "
            "verificacion (4 columnas: Equipo, Serial, Estado, Sede)  -->  "
            "Ctrl+C  -->  Haz clic en 'PEGAR DESDE EXCEL'"
        )
        tk.Label(frm2, text=instruccion, bg="#FFF3CD", fg="#856404",
                 font=("Arial", 8), wraplength=950,
                 justify="left", relief="solid", bd=1,
                 padx=6, pady=4).pack(fill="x", pady=(0, 6))

        # Barra de botones
        btn_bar = tk.Frame(frm2, bg="#F0F4F8")
        btn_bar.pack(fill="x", pady=(0, 4))

        tk.Button(btn_bar, text="PEGAR DESDE EXCEL  (Ctrl+V)",
                  command=self._pegar_desde_portapapeles,
                  bg="#28A745", fg="white",
                  font=("Arial", 10, "bold")).pack(side="left", padx=(0, 6))
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

        # Tabla Treeview
        cols = ("Equipo / Host", "Serial Real", "Estado", "Sede")
        self.tree = ttk.Treeview(frm2, columns=cols, show="headings", height=10)
        for c in cols:
            self.tree.heading(c, text=c)
        self.tree.column("Equipo / Host", width=220)
        self.tree.column("Serial Real",   width=200)
        self.tree.column("Estado",        width=100)
        self.tree.column("Sede",          width=210)

        # Colores por estado en la tabla
        self.tree.tag_configure("ok",       background="#E2EFDA")
        self.tree.tag_configure("traslado", background="#FCE4D6")

        vsb = ttk.Scrollbar(frm2, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._editar_fila)

        # Atajo Ctrl+V en la ventana principal
        self.root.bind("<Control-v>", lambda e: self._pegar_desde_portapapeles())

        # --- Bloque 3: boton ejecutar ------------------------------------
        frm3 = tk.Frame(parent, bg="#F0F4F8", pady=6)
        frm3.pack(fill="x", padx=15)
        tk.Button(
            frm3,
            text="EJECUTAR CONCILIACION COMPLETA",
            command=self._ejecutar,
            bg="#28A745", fg="white",
            font=("Arial", 12, "bold"), pady=10
        ).pack(fill="x")

        # --- Bloque 4: bitacora ------------------------------------------
        frm4 = tk.LabelFrame(parent, text="3. Bitacora",
                              padx=10, pady=6, bg="#F0F4F8",
                              font=("Arial", 10, "bold"))
        frm4.pack(fill="both", expand=True, padx=15, pady=(2, 10))

        self.txt_log = tk.Text(frm4, bg="#1E1E1E", fg="#D4D4D4",
                               font=("Consolas", 9), height=8)
        self.txt_log.pack(fill="both", expand=True)
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

        tk.Radiobutton(f_radio, text="Automático (Usa tabla de verificación física)",
                       variable=self.qr_fuente, value="auto", bg=COLOR_BG,
                       font=("Arial", 8, "bold"),
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

        # ---- Seccion 3: Acciones principales -----------------------------
        frm_acc = tk.LabelFrame(parent, text="3. Acciones de Generacion y Exportacion",
                                 padx=10, pady=8, bg=COLOR_BG,
                                 font=("Arial", 10, "bold"))
        frm_acc.pack(fill="x", padx=15, pady=4)

        btn_row = tk.Frame(frm_acc, bg=COLOR_BG)
        btn_row.pack(fill="x")

        tk.Button(btn_row,
                  text="Generar 34 Etiquetas QR (PNG)",
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

        self.txt_qr_log = tk.Text(frm_log, bg="#1E1E1E", fg="#D4D4D4",
                                   font=("Consolas", 9), height=7)
        self.txt_qr_log.pack(fill="both", expand=True)
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
            self.txt_qr_log.insert(tk.END, msg + "\n", tag)
            self.txt_qr_log.see(tk.END)
            self.root.update_idletasks()

    def _qr_actualizar_resumen_fuente(self):
        """Actualiza la etiqueta informativa de fuente de datos QR."""
        if not hasattr(self, "lbl_qr_fuente_info") or not self.lbl_qr_fuente_info:
            return
        modo = self.qr_fuente.get() if hasattr(self, "qr_fuente") else "auto"
        items_tree = self.tree.get_children() if hasattr(self, "tree") else []
        n_tree = len(items_tree)

        if modo == "preset" or (modo == "auto" and n_tree == 0):
            n_lap = len(self.PRESET_CLEMONT)
            n_tab = len(self.PRESET_TABLETS)
            n_mon = len(self.PRESET_MONITORES)
            total = n_lap + n_tab + n_mon
            txt = (f"Fuente activa: Preset Clemont (Demo)\n"
                   f" • {n_lap} Laptops/PCs | {n_tab} Tablets | {n_mon} Monitores\n"
                   f" TOTAL: {total} activos listos para rotular")
            if n_tree == 0 and modo == "auto":
                txt += " (Tabla vacia, usando demo)"
        else:
            txt = (f"Fuente activa: Tabla de Verificación Física\n"
                   f" • {n_tree} activos cargados en la tabla\n"
                   f" TOTAL: {n_tree} etiquetas personalizadas a generar")
        self.lbl_qr_fuente_info.config(text=txt)

    def _obtener_lista_equipos_qr(self):
        """
        Retorna la lista completa de equipos con sus IDs de Asset Tag asignados.
        Universal:
        - Si hay datos en la tabla (pegados desde cualquier Excel), los usa directamente.
        - Si no hay datos o se selecciona Preset, usa los 34 activos Clemont demo.
        """
        equipos = []
        modo = self.qr_fuente.get() if hasattr(self, "qr_fuente") else "auto"
        items_tree = self.tree.get_children() if hasattr(self, "tree") else []

        if (modo == "auto" and len(items_tree) > 0) or modo == "tree":
            contadores = {}
            for item_id in items_tree:
                vals = self.tree.item(item_id)["values"]
                nom_eq = str(vals[0] if len(vals) > 0 else "")
                serial = str(vals[1] if len(vals) > 1 else "")
                estado = str(vals[2] if len(vals) > 2 else "OK")
                sede   = str(vals[3] if len(vals) > 3 else "")

                nom_up = nom_eq.upper()
                if any(k in nom_up for k in ["MONITOR", "PANTALLA", "LG 27", "VIEWSONIC", "SAMSUNG", "JANUS"]):
                    cat = "MON"
                    tipo = "Monitor"
                elif any(k in nom_up for k in ["TAB", "GALAXY TAB", "IPAD"]):
                    cat = "TAB"
                    tipo = "Tablet"
                elif any(k in nom_up for k in ["IMP", "EPSON", "PRINTER", "LASERJET", "HP DESKJET"]):
                    cat = "IMP"
                    tipo = "Impresora"
                elif any(k in nom_up for k in ["SWITCH", "ROUTER", "ACCESS POINT", "FIREWALL"]):
                    cat = "RED"
                    tipo = "Redes"
                elif any(k in nom_up for k in ["DESKTOP", "TORRE", "PC-", "ALL IN ONE"]):
                    cat = "PC"
                    tipo = "PC Escritorio"
                elif any(k in nom_up for k in ["LAP", "MACBOOK", "PORTATIL", "NOTEBOOK"]):
                    cat = "LAP"
                    tipo = "Laptop"
                else:
                    cat = "LAP"
                    tipo = "Equipo / Host"

                contadores[cat] = contadores.get(cat, 0) + 1
                correlativo = contadores[cat]
                id_tag = f"{cat}-2026-{correlativo:04d}"

                equipos.append({
                    "id": id_tag,
                    "categoria": cat,
                    "tipo": tipo,
                    "equipo": nom_eq,
                    "serial": serial,
                    "sede": sede,
                    "responsable": "",
                    "estado": estado
                })
            return equipos

        # Fallback de conveniencia: Presets de Clemont (para pruebas y demostracion)
        for i, (eq, ser, est, sede) in enumerate(self.PRESET_CLEMONT, 1):
            equipos.append({
                "id": f"LAP-2026-{i:04d}",
                "categoria": "LAP",
                "tipo": "Laptop / PC",
                "equipo": eq,
                "serial": ser,
                "sede": sede,
                "responsable": "",
                "estado": est
            })

        for i, fila in enumerate(self.PRESET_TABLETS, 1):
            ser_limpio = fila[10].replace("Serial:", "").replace("Serial :", "").strip()
            equipos.append({
                "id": f"TAB-2026-{i:04d}",
                "categoria": "TAB",
                "tipo": "Tablet",
                "equipo": f"{fila[4]} ({fila[5]})",
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
                "serial": ser_limpio,
                "sede": fila[2],
                "responsable": fila[8],
                "estado": fila[9]
            })

        return equipos

    def _generar_etiqueta_individual(self, item, formato="50x25mm", carpeta="etiquetas_qr"):
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

        # Generar QR — SOLO contiene el ID del activo
        qr = qrcode.QRCode(
            box_size=10,
            border=1,
            error_correction=qrcode.constants.ERROR_CORRECT_M
        )
        qr.add_data(item["id"])
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
        qr_img = qr_img.resize((qr_dim, qr_dim), Image.Resampling.LANCZOS)
        img.paste(qr_img, (15, (h - qr_dim) // 2))

        # Guia de corte perimetral sutil
        draw.rectangle([(1, 1), (w - 2, h - 2)], outline="#C8C8C8", width=1)
        # Separador vertical
        draw.line([(qr_dim + 24, 12), (qr_dim + 24, h - 12)], fill="#E0E0E0", width=1)

        f_hdr = _obtener_fuente(sz_hdr, bold=True)
        f_id  = _obtener_fuente(sz_id, bold=True)
        f_eq  = _obtener_fuente(sz_eq, bold=True)
        f_sub = _obtener_fuente(sz_sub, bold=False)

        empresa = self.qr_empresa.get().strip().upper() if hasattr(self, "qr_empresa") and self.qr_empresa.get() else "CLEMONT"
        txt_hdr = f"{empresa}  ·  CONTROL DE ACTIVOS"

        draw.text((x_txt, int(h * 0.08)), txt_hdr, fill="#1F4E78", font=f_hdr)
        draw.text((x_txt, int(h * 0.22)), item["id"], fill="#000000", font=f_id)

        eq_txt = item["equipo"]
        if len(eq_txt) > 28:
            eq_txt = eq_txt[:26] + ".."
        draw.text((x_txt, int(h * 0.48)), eq_txt, fill="#222222", font=f_eq)

        ser_txt = f"S/N: {item['serial']}"
        if len(ser_txt) > 32:
            ser_txt = ser_txt[:30] + ".."
        draw.text((x_txt, int(h * 0.68)), ser_txt, fill="#444444", font=f_sub)

        sede_txt = f"Sede: {item['sede']}"
        draw.text((x_txt, int(h * 0.82)), sede_txt, fill="#666666", font=f_sub)

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
        carpeta = "etiquetas_qr"
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
        carpeta = "etiquetas_qr"
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

    def _qr_exportar_excel(self):
        """Exporta un libro Excel con la base maestra de IDs y la hoja de escaneos de campo."""
        carpeta = "etiquetas_qr"
        os.makedirs(carpeta, exist_ok=True)
        ruta_salida = os.path.join(carpeta, "Inventario_Equipos_Clemont_Con_QR.xlsx")

        wb = openpyxl.Workbook()

        # Hoja 1: Maestro_Activos
        ws_m = wb.active
        ws_m.title = "Maestro_Activos"

        headers_m = [
            "ID_Activo", "Categoria", "Tipo_Equipo", "Equipo / Host",
            "Serial_Fabricante", "Ubicacion_Oficial", "Responsable", "Estado_Oficial"
        ]
        ws_m.append(headers_m)
        for col_num in range(1, len(headers_m) + 1):
            cell = ws_m.cell(1, col_num)
            cell.font = FONT_HDR
            cell.fill = HDR_BLUE
            cell.alignment = Alignment(horizontal="center", vertical="center")

        equipos = self._obtener_lista_equipos_qr()
        for eq in equipos:
            ws_m.append([
                eq["id"],
                eq["categoria"],
                eq["tipo"],
                eq["equipo"],
                eq["serial"],
                eq["sede"],
                eq["responsable"],
                eq["estado"]
            ])

        for row in range(2, len(equipos) + 2):
            for col in range(1, len(headers_m) + 1):
                c = ws_m.cell(row, col)
                c.font = FONT_BASE
                c.border = THIN_BORDER
                if col == 1:
                    c.font = FONT_BOLD
                    c.alignment = Alignment(horizontal="center")

        ws_m.auto_filter.ref = f"A1:H{len(equipos) + 1}"
        anchos_m = [16, 12, 16, 28, 26, 26, 22, 14]
        for i, ancho in enumerate(anchos_m, 1):
            ws_m.column_dimensions[get_column_letter(i)].width = ancho

        # Hoja 2: Escaneos_Campo
        ws_e = wb.create_sheet(title="Escaneos_Campo")
        headers_e = [
            "ID_Escaneado", "Ubicacion_Hallada", "Fecha_Hora_Escaneo", "Operador", "Observaciones"
        ]
        ws_e.append(headers_e)
        for col_num in range(1, len(headers_e) + 1):
            cell = ws_e.cell(1, col_num)
            cell.font = FONT_HDR
            cell.fill = HDR_GREEN
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for r in range(2, 36):
            for c in range(1, len(headers_e) + 1):
                cell = ws_e.cell(r, c)
                cell.font = FONT_BASE
                cell.border = THIN_BORDER

        ws_e.auto_filter.ref = "A1:E36"
        anchos_e = [18, 26, 22, 20, 30]
        for i, ancho in enumerate(anchos_e, 1):
            ws_e.column_dimensions[get_column_letter(i)].width = ancho

        # Hoja 3: Instructivo
        ws_i = wb.create_sheet(title="Instructivo_Auditoria")
        ws_i.column_dimensions["A"].width = 6
        ws_i.column_dimensions["B"].width = 85
        ws_i.append(["", "INSTRUCTIVO PARA AUDITORIA DE INVENTARIO CON CODIGOS QR"])
        ws_i.cell(1, 2).font = Font(name="Calibri", size=13, bold=True, color="1F4E78")
        pasos = [
            ("1", "Imprimir las etiquetas QR en pliegos adhesivos tamano carta e instalarlas en cada equipo."),
            ("2", "Cada codigo QR codifica exclusivamente el ID alfanumerico (ej: LAP-2026-0001, MON-2026-0001)."),
            ("3", "En campo, el auditor escanea el QR en la hoja 'Escaneos_Campo' usando lector de codigo de barras."),
            ("4", "Se registra la ubicacion fisica donde realmente fue encontrado el equipo."),
            ("5", "El Conciliador Universal compara 'Escaneos_Campo' contra 'Maestro_Activos'."),
            ("6", "El sistema clasifica: CONCILIADO OK, TRASLADO NO REGISTRADO, FALTANTE o SOBRANTE.")
        ]
        for p_num, p_txt in pasos:
            ws_i.append(["", f"{p_num}. {p_txt}"])
            ws_i.cell(ws_i.max_row, 2).font = FONT_BASE

        wb.save(ruta_salida)
        self.log_qr(f"✔ Plantilla Excel exportada con 34 activos: {os.path.basename(ruta_salida)}", "ok")
        messagebox.showinfo(
            "Excel Exportado",
            f"Se ha creado el archivo Excel con los 34 activos y la plantilla de escaneos:\n\n"
            f"{os.path.abspath(ruta_salida)}"
        )

    def _qr_probar_escaneo(self, event=None):
        """Busca en tiempo real el ID o Serial ingresado/escaneado para prueba en vivo."""
        termino = self.qr_scan_input.get().strip().upper()
        if not termino:
            return

        equipos = self._obtener_lista_equipos_qr()
        encontrado = None

        for eq in equipos:
            if (eq["id"].upper() == termino or
                normalizar(eq["serial"]) == normalizar(termino) or
                eq["equipo"].upper() == termino):
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
        carpeta = os.path.abspath("etiquetas_qr")
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
            self._cargar_hojas_en_combo()

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
        """Lee los encabezados de la hoja seleccionada y actualiza los dropdowns de mapeo."""
        ruta = self.ruta_excel.get()
        hoja = hoja_especifica or self.nombre_hoja.get().strip()
        if not ruta or not os.path.exists(ruta) or not hoja:
            return
        try:
            wb = openpyxl.load_workbook(ruta, read_only=True)
            if hoja not in wb.sheetnames:
                wb.close()
                return
            ws = wb[hoja]

            fila_hdr = self._detectar_fila_encabezados(ws)
            max_c = min(ws.max_column or 20, 30)

            opciones_cols = []
            col_serial_det = None
            col_equipo_det = None
            col_sede_det = None
            col_verif_det = None

            for c in range(1, max_c + 1):
                letra = openpyxl.utils.get_column_letter(c)
                val = str(ws.cell(fila_hdr, c).value or "").strip()
                val_limpio = " ".join(val.split())
                etiqueta = f"Col {letra} ({c})" + (f" - {val_limpio[:25]}" if val_limpio else "")
                opciones_cols.append(etiqueta)

                val_up = val_limpio.upper()
                if not col_serial_det and any(k in val_up for k in ["SERIAL", "S/N", "OBSERVACION"]):
                    col_serial_det = etiqueta
                if not col_equipo_det and any(k in val_up for k in ["EQUIPO", "HOST", "DISPOSITIVO", "SERVIDOR"]):
                    col_equipo_det = etiqueta
                if not col_sede_det and any(k in val_up for k in ["SEDE", "UBICACION", "CIUDAD", "LUGAR"]):
                    col_sede_det = etiqueta
                if not col_verif_det and any(k in val_up for k in ["VERIFICADO", "ESTADO", "AUDITORIA", "CHECK"]):
                    col_verif_det = etiqueta

            wb.close()

            # Poblar los dropdowns
            for combo in (self.combo_col_serial, self.combo_col_equipo, self.combo_col_sede, self.combo_col_verif):
                combo["values"] = opciones_cols

            if col_serial_det: self.col_serial_var.set(col_serial_det)
            if col_equipo_det: self.col_equipo_var.set(col_equipo_det)
            if col_sede_det:   self.col_sede_var.set(col_sede_det)
            if col_verif_det:  self.col_verif_var.set(col_verif_det)

            self.log(f"Encabezados detectados en Fila {fila_hdr} de [{hoja}]:", "info")
            self.log(f"   Serial: [{self.col_serial_var.get()}] | Host: [{self.col_equipo_var.get()}] | Sede: [{self.col_sede_var.get()}] | Auditoria: [{self.col_verif_var.get()}]", "ok")

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
        """Detecta la fila de encabezados principales buscando coincidencias semanticas."""
        for r in range(1, min(15, ws.max_row + 1)):
            valores = [str(ws.cell(r, c).value or "").strip().upper() for c in range(1, min(25, ws.max_column + 1))]
            coincidencias = sum(1 for v in valores if any(k in v for k in ["SERIAL", "EQUIPO", "HOST", "SEDE", "UBICACION", "ITEM", "CLIENTE", "RESPONSABLE", "OBSERVACION"]))
            if coincidencias >= 2:
                return r
        return 5 # fallback Clemont

    def _detectar_fila_total(self, ws, fila_inicio):
        """Busca la fila de TOTALES posterior a los datos."""
        for r in range(fila_inicio, ws.max_row + 1):
            for c in range(1, min(6, ws.max_column + 1)):
                val = str(ws.cell(r, c).value or "").upper()
                if "TOTAL" in val:
                    return r
        return ws.max_row + 1

    # ------ PEGADO DESDE PORTAPAPELES ------------------------------------

    def _cargar_preset_clemont(self):
        """Carga la lista de equipos verificados en auditoria fisica de Clemont."""
        self._limpiar_tabla()
        for eq in self.PRESET_CLEMONT:
            tag = "traslado" if eq[2] == "Traslado" else "ok"
            self.tree.insert("", "end", values=eq, tags=(tag,))
        self._actualizar_conteo()
        self.log(f"Preset Clemont cargado ({len(self.PRESET_CLEMONT)} equipos verificados).", "info")

    def _pegar_desde_portapapeles(self):
        """
        Lee el contenido del portapapeles (texto TSV que copia Excel)
        y lo agrega a la tabla. Formato esperado de cada fila:
            Equipo <TAB> Serial <TAB> Estado <TAB> Sede
        Si el usuario pega filas con mas o menos columnas se hace lo posible.
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
        agregados = 0
        omitidos  = 0

        for linea in lineas:
            linea = linea.strip()
            if not linea:
                continue

            # Excel separa columnas con TAB
            partes = linea.split("\t")

            # Necesitamos al menos Equipo y Serial (cols 0 y 1)
            if len(partes) < 2:
                omitidos += 1
                continue

            equipo = partes[0].strip() if len(partes) > 0 else ""
            serial = partes[1].strip() if len(partes) > 1 else ""
            estado = partes[2].strip() if len(partes) > 2 else "OK"
            sede   = partes[3].strip() if len(partes) > 3 else ""

            # Normalizar estado
            est_norm = normalizar(estado)
            if est_norm in ("TRASLADO", "TRASL"):
                estado = "Traslado"
                tag = "traslado"
            else:
                estado = "OK"
                tag = "ok"

            if not equipo and not serial:
                omitidos += 1
                continue

            self.tree.insert("", "end", values=(equipo, serial, estado, sede),
                             tags=(tag,))
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
                "Asegurate de copiar exactamente 4 columnas desde Excel:\n"
                "  Equipo | Serial | Estado | Sede"
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
        dlg.title("Agregar / Editar equipo")
        dlg.geometry("440x230")
        dlg.grab_set()

        campos = [
            ("Equipo / Host",        0),
            ("Serial Real",          1),
            ("Estado (OK / Traslado)", 2),
            ("Sede",                 3),
        ]
        entradas = {}
        for lbl, idx in campos:
            frm = tk.Frame(dlg)
            frm.pack(fill="x", padx=16, pady=5)
            tk.Label(frm, text=lbl, width=24, anchor="w").pack(side="left")
            var = tk.StringVar(value=(valores[idx] if valores else ""))
            tk.Entry(frm, textvariable=var, width=28).pack(side="left")
            entradas[idx] = var

        def guardar():
            vals = tuple(entradas[i].get().strip() for i in range(4))
            tag  = "traslado" if normalizar(vals[2]) == "TRASLADO" else "ok"
            if item_id:
                self.tree.item(item_id, values=vals, tags=(tag,))
            else:
                self.tree.insert("", "end", values=vals, tags=(tag,))
            self._actualizar_conteo()
            dlg.destroy()

        tk.Button(dlg, text="Guardar", command=guardar,
                  bg="#28A745", fg="white",
                  font=("Arial", 10, "bold")).pack(pady=10)

    def _actualizar_conteo(self):
        n = len(self.tree.get_children())
        self.lbl_conteo.config(text=f"{n} equipo(s) cargado(s)")
        self._qr_actualizar_resumen_fuente()

    # ------------------------------------------------------------------
    # Logging
    # ------------------------------------------------------------------

    def log(self, msg, tag="info"):
        self.txt_log.insert(tk.END, msg + "\n", tag)
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
        lista_fisica = [
            {
                "equipo": self.tree.item(i, "values")[0],
                "serial": self.tree.item(i, "values")[1],
                "estado": self.tree.item(i, "values")[2],
                "sede":   self.tree.item(i, "values")[3],
            }
            for i in self.tree.get_children()
        ]

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

        # Obtener columnas mapeadas dinamicamente
        col_sede   = self._get_col_idx(self.col_sede_var.get(), default=self.COL_SEDE)
        col_equipo = self._get_col_idx(self.col_equipo_var.get(), default=self.COL_EQUIPO)
        col_obs    = self._get_col_idx(self.col_serial_var.get(), default=self.COL_OBS)
        col_verif  = self._get_col_idx(self.col_verif_var.get(), default=self.COL_VERIF)

        # Detectar fila de encabezados y de datos
        hdr_row = self._detectar_fila_encabezados(ws)
        data_start = hdr_row + 1
        total_row = self._detectar_fila_total(ws, data_start)

        self.log(f"Encabezados en F{hdr_row} | Datos desde F{data_start} hasta F{total_row - 1}", "info")
        self.log(f"Columnas activas: Serial=Col {get_column_letter(col_obs)}, Host=Col {get_column_letter(col_equipo)}, Sede=Col {get_column_letter(col_sede)}, Auditoria=Col {get_column_letter(col_verif)}", "ok")

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
                # Match por serial
                for s_celda in seriales_celda:
                    if similitud_serial(eq["serial"], s_celda):
                        coincidencia = eq
                        idx_match    = idx_f
                        break
                # Match por nombre de equipo (fallback)
                if not coincidencia:
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

                # Actualizar nombre de equipo
                ws.cell(r, col_equipo).value = coincidencia["equipo"]

                # Actualizar serial corregido
                texto_obs = str(ws.cell(r, col_obs).value or "")
                serial_ok = f"Serial equipo: {coincidencia['serial']}"
                m_modelo  = re.search(r"(?i)(Modelo\s*:.+)", texto_obs)
                if m_modelo:
                    serial_ok += f"; {m_modelo.group(1).strip()}"
                ws.cell(r, col_obs).value = serial_ok

                # Marcar en columna de resultado
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
                self.log(
                    f"[{estado}] F{r} -> {coincidencia['equipo']} | "
                    f"{coincidencia['serial']}", "ok"
                )
            else:
                if ws.cell(r, 2).value:
                    if r in filas_dup:
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

        # --- Corregir formula TOTAL (si existe columna cantidad col 8) ---
        if total_row <= ws.max_row:
            ws.cell(total_row, 8).value = f"=SUM(H{data_start}:H{total_row - 1})"

        # --- Insertar altas antes del TOTAL -----------------------------
        for idx_a, alta in enumerate(altas):
            insert_r = total_row + idx_a
            ws.insert_rows(insert_r)
            ws.cell(insert_r, 1).value                   = insert_r - hdr_row
            ws.cell(insert_r, 2).value                   = "CLEMONT"
            ws.cell(insert_r, col_sede).value            = alta["sede"]
            ws.cell(insert_r, col_equipo).value          = alta["equipo"]
            ws.cell(insert_r, 5).value                   = "WINDOWS 11 PRO"
            ws.cell(insert_r, 6).value                   = "Microsoft"
            ws.cell(insert_r, 7).value                   = "Sistema operativo"
            ws.cell(insert_r, 8).value                   = 1
            ws.cell(insert_r, 10).value                  = "Activo"
            ws.cell(insert_r, col_obs).value             = (
                f"Serial equipo: {alta['serial']}")
            ws.cell(insert_r, col_verif).value           = alta["estado"]
            for c in range(1, col_verif + 1):
                ws.cell(insert_r, c).font   = FONT_BASE
                ws.cell(insert_r, c).border = THIN_BORDER
                ws.cell(insert_r, c).fill   = FILL_ALTA
            ws.cell(insert_r, col_verif).font = FONT_CHECK
            ws.cell(insert_r, col_verif).alignment = Alignment(
                horizontal="center")
            self.log(f"Insertado alta: {alta['equipo']} en F{insert_r}", "ok")

        nuevo_total_row = total_row + len(altas)
        if total_row <= ws.max_row:
            ws.cell(nuevo_total_row, 8).value = (
                f"=SUM(H{data_start}:H{nuevo_total_row - 1})"
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
        dir_s, nom_orig = os.path.split(ruta)
        nom_salida = (
            "AUDITADO_" + nom_orig
            if not nom_orig.startswith("AUDITADO_")
            else nom_orig
        )
        ruta_guardado = os.path.join(dir_s, nom_salida)
        wb.save(ruta_guardado)

        self.log("=" * 58, "info")
        self.log("  CONCILIACION COMPLETADA", "ok")
        self.log(f"  Verificados (OK)  : {contadores['ok']}", "ok")
        self.log(f"  Traslados         : {contadores['traslado']}", "warn")
        self.log(f"  Sin chequear      : {contadores['sin_chequear']}", "info")
        self.log(f"  Duplicados        : {contadores['dup']}", "warn")
        self.log(f"  Altas (nuevos)    : {len(altas)}", "warn")
        self.log(f"  Archivo guardado  : {nom_salida}", "info")
        self.log("=" * 58, "info")

        messagebox.showinfo(
            "Conciliacion completada",
            f"Verificados (OK) : {contadores['ok']}\n"
            f"Traslados        : {contadores['traslado']}\n"
            f"Sin chequear     : {contadores['sin_chequear']}\n"
            f"Duplicados       : {contadores['dup']}\n"
            f"Altas nuevas     : {len(altas)}\n\n"
            f"Guardado como: {nom_salida}"
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
