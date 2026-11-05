# -*- coding: utf-8 -*-
"""
conciliador_universal.py
Herramienta de auditoria y conciliacion de inventario de equipos.
Version 2.0  -  Python 3.x + Tkinter + openpyxl

Novedades v2:
  - Pegado masivo desde Excel (Ctrl+C en Excel -> boton Pegar en la app)
  - Deteccion de equipos fisicos NO encontrados en el Excel (altas)
  - Deteccion de filas duplicadas en el Excel
  - Hoja 'Reporte de Auditoria' con colores + autofilter
  - Autofilter y freeze panes en la hoja de inventario
  - Correccion automatica de la formula TOTAL EQUIPOS
  - Formato de hoja siempre igual al de Clemont (col K=serial, L=verificado)
"""

import os
import re
import sys
import io
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
        self.root.title("Conciliador de Inventario Clemont v2.0")
        self.root.geometry("1020x800")
        self.root.configure(bg="#F0F4F8")
        self.ruta_excel  = tk.StringVar()
        self.nombre_hoja = tk.StringVar(value="Inventario de Equipos (2)")
        self._construir_gui()

    # ------------------------------------------------------------------
    # GUI
    # ------------------------------------------------------------------

    def _construir_gui(self):
        # Titulo
        tk.Label(
            self.root,
            text="  Conciliador de Inventario - Clemont  v2.0",
            font=("Arial", 13, "bold"), bg="#1F4E78", fg="white", pady=10
        ).pack(fill="x")

        # --- Bloque 1: archivo -------------------------------------------
        frm1 = tk.LabelFrame(self.root, text="1. Archivo Excel",
                              padx=10, pady=8, bg="#F0F4F8",
                              font=("Arial", 10, "bold"))
        frm1.pack(fill="x", padx=15, pady=(10, 2))

        tk.Entry(frm1, textvariable=self.ruta_excel, width=65,
                 state="readonly", font=("Consolas", 9)).pack(
            side="left", fill="x", expand=True, padx=(0, 8))
        tk.Button(frm1, text="Examinar...", command=self._seleccionar_archivo,
                  bg="#1F4E78", fg="white",
                  font=("Arial", 9, "bold")).pack(side="left")

        frm1b = tk.Frame(self.root, bg="#F0F4F8")
        frm1b.pack(fill="x", padx=15, pady=(2, 6))
        tk.Label(frm1b, text="Hoja objetivo:",
                 bg="#F0F4F8", font=("Arial", 9)).pack(side="left")
        tk.Entry(frm1b, textvariable=self.nombre_hoja, width=38,
                 font=("Consolas", 9)).pack(side="left", padx=6)
        tk.Button(frm1b, text="Ver hojas disponibles",
                  command=self._mostrar_hojas,
                  bg="#6C757D", fg="white",
                  font=("Arial", 9)).pack(side="left", padx=4)

        # --- Bloque 2: tabla de equipos verificados ----------------------
        frm2 = tk.LabelFrame(
            self.root,
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
        self.tree = ttk.Treeview(frm2, columns=cols, show="headings", height=11)
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
        frm3 = tk.Frame(self.root, bg="#F0F4F8", pady=6)
        frm3.pack(fill="x", padx=15)
        tk.Button(
            frm3,
            text="EJECUTAR CONCILIACION COMPLETA",
            command=self._ejecutar,
            bg="#28A745", fg="white",
            font=("Arial", 12, "bold"), pady=10
        ).pack(fill="x")

        # --- Bloque 4: bitacora ------------------------------------------
        frm4 = tk.LabelFrame(self.root, text="3. Bitacora",
                              padx=10, pady=6, bg="#F0F4F8",
                              font=("Arial", 10, "bold"))
        frm4.pack(fill="both", expand=True, padx=15, pady=(2, 10))

        self.txt_log = tk.Text(frm4, bg="#1E1E1E", fg="#D4D4D4",
                               font=("Consolas", 9), height=9)
        self.txt_log.pack(fill="both", expand=True)
        self.txt_log.tag_config("ok",    foreground="#6BCF6B")
        self.txt_log.tag_config("warn",  foreground="#FFD700")
        self.txt_log.tag_config("error", foreground="#FF6B6B")
        self.txt_log.tag_config("info",  foreground="#87CEEB")

    # ------------------------------------------------------------------
    # Acciones de la GUI
    # ------------------------------------------------------------------

    def _seleccionar_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona el archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx")]
        )
        if ruta:
            self.ruta_excel.set(ruta)
            self._mostrar_hojas()

    def _mostrar_hojas(self):
        ruta = self.ruta_excel.get()
        if not ruta or not os.path.exists(ruta):
            return
        try:
            wb = openpyxl.load_workbook(ruta, read_only=True)
            self.log("Hojas disponibles: " + " | ".join(wb.sheetnames), "info")
            wb.close()
        except Exception as e:
            self.log("No se pudo leer el archivo: " + str(e), "warn")

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

        # Detectar fila TOTAL
        total_row = None
        for r in range(6, ws.max_row + 1):
            val = str(ws.cell(r, 2).value or "")
            if "TOTAL" in val.upper():
                total_row = r
                break
        total_row = total_row or (ws.max_row + 1)
        self.log(f"Fila TOTAL detectada: {total_row}", "info")

        # --- Detectar duplicados en el Excel ----------------------------
        seriales_vistos = {}   # sn -> (fila, texto_serial_original)
        nombres_vistos  = {}   # clave -> (fila, equipo_original)
        filas_dup       = {}   # r -> motivo del duplicado

        for r in range(6, total_row):
            obs    = str(ws.cell(r, self.COL_OBS).value or "")
            equipo = normalizar(str(ws.cell(r, self.COL_EQUIPO).value or ""))
            sede   = normalizar(str(ws.cell(r, self.COL_SEDE).value or ""))
            nom_orig = str(ws.cell(r, self.COL_EQUIPO).value or "")

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
                motivo_actual = f"Equipo repetido '{nom_orig}' en misma sede ({ws.cell(r, self.COL_SEDE).value}): coincide con fila {fila_ant}"
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

        for r in range(6, total_row):
            obs_celda    = str(ws.cell(r, self.COL_OBS).value or "")
            equipo_celda = str(ws.cell(r, self.COL_EQUIPO).value or "")

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
                ws.cell(r, self.COL_EQUIPO).value = coincidencia["equipo"]

                # Actualizar serial corregido en col K
                texto_obs = str(ws.cell(r, self.COL_OBS).value or "")
                serial_ok = f"Serial equipo: {coincidencia['serial']}"
                m_modelo  = re.search(r"(?i)(Modelo\s*:.+)", texto_obs)
                if m_modelo:
                    serial_ok += f"; {m_modelo.group(1).strip()}"
                ws.cell(r, self.COL_OBS).value = serial_ok

                # Marcar en col L
                ws.cell(r, self.COL_VERIF).value = estado
                if normalizar(estado) in ("OK", "VERIFICADO"):
                    ws.cell(r, self.COL_VERIF).font  = FONT_CHECK
                    ws.cell(r, self.COL_VERIF).fill  = FILL_OK
                    contadores["ok"] += 1
                else:
                    ws.cell(r, self.COL_VERIF).font  = FONT_TRASL
                    ws.cell(r, self.COL_VERIF).fill  = FILL_TRASL
                    contadores["traslado"] += 1

                ws.cell(r, self.COL_VERIF).alignment = Alignment(
                    horizontal="center")
                self.log(
                    f"[{estado}] F{r} -> {coincidencia['equipo']} | "
                    f"{coincidencia['serial']}", "ok"
                )
            else:
                if ws.cell(r, 2).value:
                    if r in filas_dup:
                        ws.cell(r, self.COL_VERIF).value = "Duplicado"
                        ws.cell(r, self.COL_VERIF).font  = Font(
                            name="Calibri", size=10, bold=True, color="7F6000")
                        ws.cell(r, self.COL_VERIF).fill  = FILL_DUP
                        # Insertar comentario / nota explicativa en la celda
                        nota = filas_dup[r]
                        ws.cell(r, self.COL_VERIF).comment = Comment(
                            f"Auditoría: {nota}", "Auditor TI"
                        )
                        contadores["dup"] += 1
                    else:
                        ws.cell(r, self.COL_VERIF).value = "Sin chequear"
                        ws.cell(r, self.COL_VERIF).font  = FONT_UNCHK
                        ws.cell(r, self.COL_VERIF).fill  = FILL_UNCHK
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
            f"A5:{get_column_letter(ws.max_column)}5"
        )
        ws.freeze_panes = "A6"

        # --- Corregir formula TOTAL -------------------------------------
        ws.cell(total_row, 8).value = f"=SUM(H6:H{total_row - 1})"

        # --- Insertar altas antes del TOTAL -----------------------------
        for idx_a, alta in enumerate(altas):
            insert_r = total_row + idx_a
            ws.insert_rows(insert_r)
            ws.cell(insert_r, 1).value                   = insert_r - 5
            ws.cell(insert_r, 2).value                   = "CLEMONT"
            ws.cell(insert_r, self.COL_SEDE).value       = alta["sede"]
            ws.cell(insert_r, self.COL_EQUIPO).value     = alta["equipo"]
            ws.cell(insert_r, 5).value                   = "WINDOWS 11 PRO"
            ws.cell(insert_r, 6).value                   = "Microsoft"
            ws.cell(insert_r, 7).value                   = "Sistema operativo"
            ws.cell(insert_r, 8).value                   = 1
            ws.cell(insert_r, 10).value                  = "Activo"
            ws.cell(insert_r, self.COL_OBS).value        = (
                f"Serial equipo: {alta['serial']}")
            ws.cell(insert_r, self.COL_VERIF).value      = alta["estado"]
            for c in range(1, self.COL_VERIF + 1):
                ws.cell(insert_r, c).font   = FONT_BASE
                ws.cell(insert_r, c).border = THIN_BORDER
                ws.cell(insert_r, c).fill   = FILL_ALTA
            ws.cell(insert_r, self.COL_VERIF).font = FONT_CHECK
            ws.cell(insert_r, self.COL_VERIF).alignment = Alignment(
                horizontal="center")
            self.log(f"Insertado alta: {alta['equipo']} en F{insert_r}", "ok")

        nuevo_total_row = total_row + len(altas)
        ws.cell(nuevo_total_row, 8).value = (
            f"=SUM(H6:H{nuevo_total_row - 1})"
        )

        # --- Tablas de perifericos --------------------------------------
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
            contadores, filas_dup
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
                         contadores, filas_dup):
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
        for r in range(6, total_row + n_altas):
            equipo = str(ws_inv.cell(r, self.COL_EQUIPO).value or "")
            serial = str(ws_inv.cell(r, self.COL_OBS).value or "")
            estado = str(ws_inv.cell(r, self.COL_VERIF).value or "")
            sede   = str(ws_inv.cell(r, self.COL_SEDE).value or "")
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
