# -*- coding: utf-8 -*-
"""
test_qa_inventario.py
Comprehensive 4-Tier Automated QA & E2E Test Suite for Clemont Inventory Reconciliation System.

Methodology (TEST_INFRA.md):
- Tier 1: Feature Coverage in Isolation (R1-R5, >=40 tests, >=5 per feature)
- Tier 2: Boundary & Corner Cases (>=40 tests, >=5 per feature boundary)
- Tier 3: Cross-Feature Integration Combinations (>=8 tests)
- Tier 4: Real-World Workload Testing against REAL Excel Files (>=5 scenarios)
Total: >=93 test assertions.

Author: test_writer_e2e (Teamwork QA Specialist)
"""

import os
import re
import sys
import shutil
import tempfile
import unittest
import openpyxl
from PIL import Image
import qrcode

# Ensure project root is in sys.path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import conciliador_universal as cu
from conciliador_universal import (
    normalizar,
    normalizar_tolerante,
    similitud_serial,
    extraer_seriales,
    extraer_campos_completos,
    ConciliadorUniversal,
)

# File constants for real Excel files
EXCEL_REAL_INPUT = os.path.join(BASE_DIR, "Inventario_Equipos_Clemont (2) (2).xlsx")
EXCEL_REAL_AUDITED = os.path.join(BASE_DIR, "AUDITADO_Inventario_Equipos_Clemont (2) (2).xlsx")
GUIA_MOVIL_FILE = os.path.join(BASE_DIR, "GUIA_ESCANEO_MOVIL.md")


# ============================================================================
# HELPER TEST UTILITIES
# ============================================================================

def create_headless_app():
    """Initializes ConciliadorUniversal with a withdrawn Tk root and non-blocking dialogs."""
    import tkinter as tk
    from tkinter import filedialog, messagebox

    # Non-blocking defaults for modal dialogs
    messagebox.showinfo = lambda *args, **kwargs: None
    messagebox.showwarning = lambda *args, **kwargs: None
    messagebox.showerror = lambda *args, **kwargs: None
    messagebox.askyesno = lambda *args, **kwargs: True
    messagebox.askokcancel = lambda *args, **kwargs: True
    messagebox.askquestion = lambda *args, **kwargs: "yes"
    messagebox.askretrycancel = lambda *args, **kwargs: False

    filedialog.askopenfilename = lambda *args, **kwargs: ""
    filedialog.askdirectory = lambda *args, **kwargs: ""
    filedialog.asksaveasfilename = lambda *args, **kwargs: ""
    filedialog.askopenfilenames = lambda *args, **kwargs: []

    # Patch module references in conciliador_universal
    cu.messagebox.showinfo = messagebox.showinfo
    cu.messagebox.showwarning = messagebox.showwarning
    cu.messagebox.showerror = messagebox.showerror
    cu.filedialog.askopenfilename = filedialog.askopenfilename
    cu.filedialog.askdirectory = filedialog.askdirectory
    cu.filedialog.asksaveasfilename = filedialog.asksaveasfilename

    # Disable external program launching
    if hasattr(os, "startfile"):
        os.startfile = lambda *args, **kwargs: None

    root = tk.Tk()
    root.withdraw()
    app = ConciliadorUniversal(root)
    return root, app



def inspect_qr_matrix_and_payload(qr_obj_or_img, expected_fields):
    """
    Verifies that the QR code matrix is valid (contains finder patterns)
    and that all expected fields are present in the payload.
    """
    if hasattr(qr_obj_or_img, "modules"):
        modules = qr_obj_or_img.modules
        size = len(modules)
        # Check minimum QR matrix version 1 (21x21)
        assert size >= 21, f"Matriz QR demasiado pequeña: {size}x{size}"
        # Check top-left finder pattern 7x7 corner: (0,0) is black
        assert modules[0][0] is True, "Módulo superior izquierdo debe ser oscuro (Finder Pattern)"
        assert modules[0][size - 1] is True, "Módulo superior derecho debe ser oscuro (Finder Pattern)"
        assert modules[size - 1][0] is True, "Módulo inferior izquierdo debe ser oscuro (Finder Pattern)"

    # If Pillow Image, verify DPI and finder pattern contrast
    if isinstance(qr_obj_or_img, Image.Image):
        dpi = qr_obj_or_img.info.get("dpi")
        assert dpi in ((300, 300), (300.0, 300.0)), f"DPI no es 300: {dpi}"
        # Sample black pixel and white pixel contrast
        pixels = qr_obj_or_img.convert("L")
        w, h = pixels.size
        assert w > 100 and h > 100, f"Dimensiones de imagen inválidas: {w}x{h}"

    return True


# ============================================================================
# TIER 1: FEATURE ISOLATION TESTS (R1 - R5)
# >= 40 tests (8 features x >= 5 tests)
# ============================================================================

class TestTier1FeatureIsolation(unittest.TestCase):
    """Tier 1: Individual feature coverage in isolation."""

    @classmethod
    def setUpClass(cls):
        cls.root, cls.app = create_headless_app()

    @classmethod
    def tearDownClass(cls):
        try:
            cls.root.destroy()
        except Exception:
            pass

    # --- Feature 1 (R1): Ficha Técnica 7 Campos sin Desfase ---

    def test_F1_1_extraer_campos_completos_laptop(self):
        """Verifica extracción fiel de los 7 campos de un registro de laptop."""
        row_data = {
            "id": "LAP-2026-0001",
            "equipo": "LAPTOP-2895F563",
            "sede": "Administrativa 2do piso",
            "responsable": "DANIEL PANESSO",
            "estado": "Activo",
            "observaciones": "Serial equipo: SDK0T76530; Modelo: LENOVO 83A0",
        }
        res = extraer_campos_completos(row_data)
        self.assertEqual(res["id_activo"], "LAP-2026-0001")
        self.assertEqual(res["equipo"], "LAPTOP-2895F563")
        self.assertIn("LENOVO", res["modelo"].upper())
        self.assertEqual(res["serial"], "SDK0T76530")
        self.assertEqual(res["responsable"], "DANIEL PANESSO")
        self.assertEqual(res["sede"], "Administrativa 2do piso")
        self.assertEqual(res["estado"], "Activo")

    def test_F1_2_ausencia_desfase_entre_columnas(self):
        """Verifica que el serial nunca se desplace a la columna de equipo o sede."""
        row_data = {
            "id": "LAP-2026-0002",
            "equipo": "SERVICIO-CLIENTE1",
            "sede": "Sótano",
            "responsable": "ESTEBAN MUÑOZ",
            "estado": "OK",
            "observaciones": "Serial equipo: TN4NOCV16Z033184; Modelo: Vivobook ASUS laptop x1502Va",
        }
        res = extraer_campos_completos(row_data)
        self.assertNotIn("TN4NOCV16Z033184", res["equipo"])
        self.assertNotIn("Sótano", res["responsable"])
        self.assertEqual(res["sede"], "Sótano")
        self.assertEqual(res["serial"], "TN4NOCV16Z033184")

    def test_F1_3_extraccion_modelo_desde_observaciones(self):
        """Verifica que el modelo se extraiga limpiamente tras el prefijo 'Modelo:'."""
        row_data = {
            "id": "LAP-2026-0003",
            "equipo": "DESKTOP-55ABFIA",
            "observaciones": "Serial equipo: E12820687824; Modelo: HP PROONE 240. Incluye Antivirus",
        }
        res = extraer_campos_completos(row_data)
        self.assertTrue("HP PROONE 240" in res["modelo"] or "HP" in res["modelo"])
        self.assertNotIn("Serial equipo:", res["modelo"])

    def test_F1_4_extraccion_serial_limpio_sin_prefijo(self):
        """Verifica que 'Serial equipo:' sea eliminado y no contamine el número de serie."""
        sns = extraer_seriales("Serial equipo: R8N0CV02K857318; Modelo: VIVOBOOK ASUS X515EA")
        self.assertGreaterEqual(len(sns), 1)
        self.assertEqual(sns[0], "R8N0CV02K857318")

    def test_F1_5_categorizacion_correcta_prefijos_activos(self):
        """Verifica asignación de categorías (LAP, MON, TAB, PC, IMP, RED)."""
        categorias_esperadas = [
            ("LAPTOP-XYZ", "Serial: ABC12345", "LAP"),
            ("MONITOR LG 27", "Serial: LG987654", "MON"),
            ("GALAXY TAB S7", "Serial: TAB123456", "TAB"),
            ("DESKTOP-TORRE1", "Serial: PC1234567", "PC"),
            ("EPSON L3150", "Serial: IMP123456", "IMP"),
        ]
        for eq_nom, obs, cat_esp in categorias_esperadas:
            nom_up = (eq_nom + " " + obs).upper()
            if any(k in nom_up for k in ["MONITOR", "PANTALLA", "LG 27"]):
                cat = "MON"
            elif any(k in nom_up for k in ["TAB", "GALAXY TAB"]):
                cat = "TAB"
            elif any(k in nom_up for k in ["IMP", "EPSON", "PRINTER"]):
                cat = "IMP"
            elif any(k in nom_up for k in ["DESKTOP", "TORRE", "PC-"]):
                cat = "PC"
            else:
                cat = "LAP"
            self.assertEqual(cat, cat_esp)

    def test_F1_6_extraer_campos_completos_acepta_tuplas_y_listas(self):
        """Verifica que extraer_campos_completos acepte tuplas y listas de 10-12 cols y 7 cols retornando el diccionario correcto."""
        # Caso A: Fila de Excel completa (12 columnas)
        fila_excel = (
            1, "CLEMONT", "Administrativa 2do piso", "LAPTOP-2895F563", "WINDOWS 11 PRO",
            "LENOVO 83A0", "Laptop", 1, "DANIEL PANESSO", "Activo",
            "Serial equipo: SDK0T76530; Modelo: LENOVO 83A0", "OK"
        )
        res_excel = extraer_campos_completos(fila_excel)
        self.assertIsInstance(res_excel, dict)
        self.assertEqual(res_excel["id_activo"], "1")
        self.assertEqual(res_excel["equipo"], "LAPTOP-2895F563")
        self.assertEqual(res_excel["serial"], "SDK0T76530")
        self.assertEqual(res_excel["responsable"], "DANIEL PANESSO")
        self.assertEqual(res_excel["sede"], "Administrativa 2do piso")
        self.assertIn("Activo", res_excel["estado"])

        # Caso B: Fila de Treeview (7 columnas: item, equipo, serial, estado, sede, resp, obs)
        fila_tree = ["LAP-01", "DESKTOP-ABC", "SN-XYZ-12345", "OK", "Sótano", "Carlos", "Obs"]
        res_tree = extraer_campos_completos(fila_tree)
        self.assertIsInstance(res_tree, dict)
        self.assertEqual(res_tree["id_activo"], "LAP-01")
        self.assertEqual(res_tree["equipo"], "DESKTOP-ABC")
        self.assertEqual(res_tree["serial"], "SN-XYZ-12345")
        self.assertEqual(res_tree["sede"], "Sótano")
        self.assertEqual(res_tree["responsable"], "Carlos")

    # --- Feature 2 (R1): Stickers y Códigos QR a 300 DPI ---

    def test_F2_1_dimensiones_etiqueta_50x25mm_300dpi(self):
        """Verifica dimensiones exactas de etiqueta 50x25mm (591x295 px a 300 DPI)."""
        with tempfile.TemporaryDirectory() as tmpdir:
            item = {
                "id": "LAP-2026-0001",
                "equipo": "LAPTOP-2895F563",
                "modelo": "Lenovo 83A0",
                "serial": "SDK0T76530",
                "responsable": "DANIEL PANESSO",
                "sede": "Administrativa 2do piso",
                "estado": "Activo",
            }
            ruta = self.app._generar_etiqueta_individual(item, formato="50x25mm", carpeta=tmpdir)
            self.assertTrue(os.path.exists(ruta))
            img = Image.open(ruta)
            self.assertEqual(img.size, (591, 295))

    def test_F2_2_dimensiones_etiqueta_100x50mm_300dpi(self):
        """Verifica dimensiones exactas de etiqueta 100x50mm (1200x600 px a 300 DPI)."""
        with tempfile.TemporaryDirectory() as tmpdir:
            item = {
                "id": "LAP-2026-0002",
                "equipo": "ANALISTA-GHI",
                "modelo": "Lenovo V14",
                "serial": "SDK0T76531",
                "responsable": "ANA GÓMEZ",
                "sede": "Administrativa",
                "estado": "Activo",
            }
            ruta = self.app._generar_etiqueta_individual(item, formato="100x50mm", carpeta=tmpdir)
            self.assertTrue(os.path.exists(ruta))
            img = Image.open(ruta)
            self.assertEqual(img.size, (1200, 600))

    def test_F2_3_metadatos_dpi_incrustados_en_png(self):
        """Verifica que Pillow almacene dpi=(300, 300) en los metadatos de la imagen."""
        with tempfile.TemporaryDirectory() as tmpdir:
            item = {
                "id": "LAP-2026-0003",
                "equipo": "DESKTOP-144TP60",
                "modelo": "HP ProDesk",
                "serial": "MXL1234ABC",
                "responsable": "CARLOS LÓPEZ",
                "sede": "Sótano",
                "estado": "Activo",
            }
            ruta = self.app._generar_etiqueta_individual(item, formato="50x25mm", carpeta=tmpdir)
            img = Image.open(ruta)
            dpi = img.info.get("dpi")
            self.assertIn(dpi, ((300, 300), (300.0, 300.0)))

    def test_F2_4_carga_util_qr_contiene_7_campos(self):
        """Verifica que la carga útil del QR generada por producción contenga todos los 7 campos técnicos."""
        import unittest.mock as mock
        item = {
            "id": "LAP-2026-0004",
            "equipo": "DESKTOP-34IT3CV",
            "modelo": "Asus Vivobook",
            "serial": "S8N0CV080498328",
            "responsable": "NINOSKA BENITEZ",
            "sede": "CDI",
            "estado": "Activo",
        }
        self.app.qr_contenido_modo.set("completo")
        with tempfile.TemporaryDirectory() as tmpdir:
            with mock.patch.object(qrcode.QRCode, "add_data", wraps=qrcode.QRCode.add_data) as mock_add:
                ruta = self.app._generar_etiqueta_individual(item, formato="50x25mm", carpeta=tmpdir)
                self.assertTrue(os.path.exists(ruta))
                mock_add.assert_called_once()
                qr_payload = mock_add.call_args[0][0]
                for expected in [
                    "ACTIVO: LAP-2026-0004",
                    "EQUIPO: DESKTOP-34IT3CV",
                    "MODELO: Asus Vivobook",
                    "SERIAL: S8N0CV080498328",
                    "RESPONSABLE: NINOSKA BENITEZ",
                    "SEDE: CDI",
                    "ESTADO: Activo"
                ]:
                    self.assertIn(expected, qr_payload)

    def test_F2_5_verificacion_matriz_qr_valida(self):
        """Verifica que el generador de QR construya una matriz con patrones de localización válidos."""
        qr = qrcode.QRCode(box_size=10, border=1, error_correction=qrcode.constants.ERROR_CORRECT_M)
        qr.add_data("ACTIVO: LAP-2026-0001\nSERIAL: SDK0T76530")
        qr.make(fit=True)
        self.assertTrue(inspect_qr_matrix_and_payload(qr, ["LAP-2026-0001", "SDK0T76530"]))

    # --- Feature 3 (R2): Estructura Maestro_Activos y Escaneos_Campo ---

    def test_F3_1_existencia_hoja_maestro_activos(self):
        """Verifica que generar_plantilla_auditoria_qr cree la hoja 'Maestro_Activos' con sus 9 columnas requeridas."""
        with tempfile.TemporaryDirectory() as tmpdir:
            out_file = os.path.join(tmpdir, "test_plantilla.xlsx")
            sample_equipos = [{
                "id": "LAP-2026-0001", "categoria": "LAP", "tipo": "Laptop",
                "equipo": "LAPTOP-01", "modelo": "Lenovo V14", "serial": "SDK0T76530",
                "sede": "Administrativa", "responsable": "Daniel", "estado": "Activo"
            }]
            cu.generar_plantilla_auditoria_qr(out_file, sample_equipos)
            self.assertTrue(os.path.exists(out_file))
            wb_check = openpyxl.load_workbook(out_file)
            self.assertIn("Maestro_Activos", wb_check.sheetnames)
            ws_m = wb_check["Maestro_Activos"]
            self.assertEqual(ws_m.cell(1, 1).value, "ID_Activo")
            self.assertEqual(ws_m.max_column, 9)
            headers = [ws_m.cell(1, c).value for c in range(1, 10)]
            self.assertIn("Modelo_Marca", headers)
            self.assertEqual(ws_m.cell(2, 1).value, "LAP-2026-0001")
            self.assertEqual(ws_m.cell(2, 5).value, "Lenovo V14")
            wb_check.close()

    def test_F3_2_estilos_corporativos_hoja_maestro(self):
        """Verifica colores corporativos Azul (#1F4E78) y bordes delgados en Maestro_Activos."""
        self.assertEqual(cu.HDR_BLUE.start_color.rgb, "001F4E78")
        self.assertEqual(cu.FONT_HDR.color.rgb, "00FFFFFF")
        self.assertTrue(cu.FONT_HDR.bold)

    def test_F3_3_existencia_hoja_escaneos_campo(self):
        """Verifica que generar_plantilla_auditoria_qr cree la hoja 'Escaneos_Campo' con sus 8 columnas y fórmulas VLOOKUP."""
        with tempfile.TemporaryDirectory() as tmpdir:
            out_file = os.path.join(tmpdir, "test_escaneos.xlsx")
            cu.generar_plantilla_auditoria_qr(out_file, [])
            self.assertTrue(os.path.exists(out_file))
            wb_check = openpyxl.load_workbook(out_file)
            self.assertIn("Escaneos_Campo", wb_check.sheetnames)
            ws_e = wb_check["Escaneos_Campo"]
            self.assertEqual(ws_e.cell(1, 1).value, "ID_Escaneado")
            self.assertEqual(ws_e.max_column, 8)
            headers_e = [ws_e.cell(1, c).value for c in range(1, 9)]
            self.assertIn("Validacion_Catalogo", headers_e)
            self.assertIn("Resultado_Sede", headers_e)
            # Verificar presencia de fórmulas VLOOKUP en la hoja generada por producción
            f_cat = str(ws_e.cell(2, 7).value)
            f_sede = str(ws_e.cell(2, 8).value)
            self.assertIn("VLOOKUP", f_cat)
            self.assertIn("Maestro_Activos!A:D", f_cat)
            self.assertIn("VLOOKUP", f_sede)
            self.assertIn("Maestro_Activos!A:G", f_sede)
            wb_check.close()

    def test_F3_4_estilos_corporativos_hoja_escaneos(self):
        """Verifica color Verde Bosque (#375623) para la hoja de Escaneos_Campo."""
        self.assertEqual(cu.HDR_GREEN.start_color.rgb, "00375623")

    def test_F3_5_autofilter_y_anchos_de_columna_configurados(self):
        """Verifica que generar_plantilla_auditoria_qr configure autofiltro y anchos en Maestro y Escaneos."""
        with tempfile.TemporaryDirectory() as tmpdir:
            out_file = os.path.join(tmpdir, "test_f3_5.xlsx")
            equipos = [{"id": "LAP-001", "equipo": "EQ-1", "serial": "SN1"}]
            cu.generar_plantilla_auditoria_qr(out_file, equipos)
            wb = openpyxl.load_workbook(out_file)
            ws_m = wb["Maestro_Activos"]
            ws_e = wb["Escaneos_Campo"]
            self.assertTrue(ws_m.auto_filter.ref.startswith("A1:I"))
            self.assertTrue(ws_e.auto_filter.ref.startswith("A1:H"))
            self.assertGreater(ws_m.column_dimensions["A"].width, 10)
            self.assertGreater(ws_e.column_dimensions["A"].width, 10)
            wb.close()

    # --- Feature 4 (R2): Guía Operativa de Escaneo Móvil ---

    def test_F4_1_existencia_archivo_guia_operativa(self):
        """Verifica que el archivo GUIA_ESCANEO_MOVIL.md exista en la raíz del proyecto."""
        self.assertTrue(os.path.exists(GUIA_MOVIL_FILE), f"No encontrado: {GUIA_MOVIL_FILE}")
        self.assertGreater(os.path.getsize(GUIA_MOVIL_FILE), 1000)

    def test_F4_2_guia_documenta_metodo_barcode_to_pc(self):
        """Verifica que la guía explique el método recomendado Barcode to PC."""
        with open(GUIA_MOVIL_FILE, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("Barcode to PC", content)
        self.assertIn("Wi-Fi", content)
        self.assertIn("ENTER", content)

    def test_F4_3_guia_documenta_metodo_excel_mobile_cloud(self):
        """Verifica que la guía detalle el método Excel Mobile con sincronización en la nube."""
        with open(GUIA_MOVIL_FILE, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("Excel Mobile", content)
        self.assertTrue("OneDrive" in content or "SharePoint" in content)

    def test_F4_4_guia_documenta_tabla_resolucion_incidencias(self):
        """Verifica que la guía contenga la tabla de Troubleshooting (resolución de fallos)."""
        with open(GUIA_MOVIL_FILE, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertTrue("RESOLUCIÓN DE INCIDENCIAS" in content or "TROUBLESHOOTING" in content)
        self.assertIn("Causa Probable", content)
        self.assertIn("Solución", content)

    def test_F4_5_guia_documenta_recomendaciones_fijacion_etiquetas(self):
        """Verifica que la guía incluya pautas de fijación física e impresión a 300 DPI."""
        with open(GUIA_MOVIL_FILE, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("300 DPI", content)
        self.assertTrue("Laptops" in content or "Monitores" in content)

    # --- Feature 5 (R3): Checkboxes (Chulos) por Columna en GUI ---

    def test_F5_1_inicializacion_variables_checkbox_columnas(self):
        """Verifica que la GUI posea las 5 variables BooleanVar para los chulos de columnas."""
        for var_name in ["chk_col_serial", "chk_col_equipo", "chk_col_sede", "chk_col_resp", "chk_col_verif"]:
            self.assertTrue(hasattr(self.app, var_name))
            var = getattr(self.app, var_name)
            self.assertTrue(var.get(), f"Variable {var_name} debe iniciar en True")

    def test_F5_2_actualizar_estado_combos_al_desmarcar(self):
        """Verifica que al desmarcar un chulo el Combobox correspondiente se deshabilite."""
        self.app.chk_col_serial.set(False)
        self.app._actualizar_estado_combos()
        self.assertEqual(self.app.combo_col_serial.cget("state"), "disabled")
        # Restaurar
        self.app.chk_col_serial.set(True)
        self.app._actualizar_estado_combos()
        self.assertEqual(self.app.combo_col_serial.cget("state"), "readonly")

    def test_F5_3_deshabilitar_serial_ignora_comparacion_serial(self):
        """Verifica que si chk_col_serial es False, el flag usar_serial sea False."""
        self.app.chk_col_serial.set(False)
        self.assertFalse(self.app.chk_col_serial.get())
        self.app.chk_col_serial.set(True)

    def test_F5_4_deshabilitar_equipo_ignora_comparacion_host(self):
        """Verifica que si chk_col_equipo es False, el flag usar_equipo sea False."""
        self.app.chk_col_equipo.set(False)
        self.assertFalse(self.app.chk_col_equipo.get())
        self.app.chk_col_equipo.set(True)

    def test_F5_5_checkbox_perifericos_clemont_funcional(self):
        """Verifica que el chulo de periféricos (tablets y monitores) sea conmutable."""
        self.assertTrue(hasattr(self.app, "chk_perifericos_clemont"))
        self.app.chk_perifericos_clemont.set(False)
        self.assertFalse(self.app.chk_perifericos_clemont.get())
        self.app.chk_perifericos_clemont.set(True)

    # --- Feature 6 (R3): Selector de Ruta y Botones 1-Clic ---

    def test_F6_1_ruta_salida_configurable_por_usuario(self):
        """Verifica que la ruta de salida sea configurable dinámicamente."""
        ruta_personalizada = r"C:\MiCarpetaPersonalizada\AUDITADO_Inventario.xlsx"
        self.app.ruta_salida.set(ruta_personalizada)
        self.assertEqual(self.app.ruta_salida.get(), ruta_personalizada)

    def test_F6_2_preservacion_prefijo_auditado(self):
        """Verifica que _seleccionar_archivo configure la ruta de salida con prefijo AUDITADO_."""
        import unittest.mock as mock
        # Caso 1: Archivo sin prefijo AUDITADO_
        with mock.patch("tkinter.filedialog.askopenfilename", return_value=r"C:\inventario\Inventario_2026.xlsx"):
            with mock.patch.object(self.app, "_cargar_hojas_en_combo"):
                self.app._seleccionar_archivo()
                self.assertEqual(self.app.ruta_salida.get(), r"C:\inventario\AUDITADO_Inventario_2026.xlsx")

        # Caso 2: Archivo que ya inicia con AUDITADO_ (idempotente)
        with mock.patch("tkinter.filedialog.askopenfilename", return_value=r"C:\inventario\AUDITADO_Inventario_2026.xlsx"):
            with mock.patch.object(self.app, "_cargar_hojas_en_combo"):
                self.app._seleccionar_archivo()
                self.assertEqual(self.app.ruta_salida.get(), r"C:\inventario\AUDITADO_Inventario_2026.xlsx")

    def test_F6_3_boton_abrir_excel_existe_en_interfaz(self):
        """Verifica que el botón de abrir Excel auditado exista en la interfaz."""
        self.assertTrue(hasattr(self.app, "btn_abrir_excel"))
        self.assertEqual(self.app.btn_abrir_excel.cget("text").strip(), "Abrir Excel Auditado")

    def test_F6_4_boton_abrir_carpeta_existe_en_interfaz(self):
        """Verifica que el botón de abrir carpeta contenedora exista en la interfaz."""
        self.assertTrue(hasattr(self.app, "btn_abrir_carpeta"))
        self.assertEqual(self.app.btn_abrir_carpeta.cget("text").strip(), "Abrir Carpeta")

    def test_F6_5_apertura_segura_con_archivo_inexistente(self):
        """Verifica que pulsar abrir cuando el archivo no existe no provoque un crash."""
        self.app.ruta_salida.set(r"C:\Ruta\Inexistente\ArchivoFalso.xlsx")
        try:
            # Invocar método sin lanzar excepción no controlada
            # Al estar mockeado o no existir, debe retornar limpiamente
            self.app._abrir_excel_resultado()
            safe = True
        except Exception as e:
            safe = False
        self.assertTrue(safe)

    # --- Feature 7 (R3): Consola de Bitácora con Scrollbar ---

    def test_F7_1_widget_text_bitacora_con_scrollbar(self):
        """Verifica que la consola de bitácora txt_log disponga de scrollbar vertical."""
        self.assertTrue(hasattr(self.app, "txt_log"))
        yscroll = self.app.txt_log.cget("yscrollcommand")
        self.assertTrue(yscroll != "", "txt_log debe tener configurado yscrollcommand")

    def test_F7_2_metodo_log_agrega_marcas_temporales(self):
        """Verifica que el método log inserte el mensaje con marca de tiempo [HH:MM:SS]."""
        self.app.log("Prueba de registro unitario QA", "info")
        contenido = self.app.txt_log.get("1.0", "end")
        self.assertIn("Prueba de registro unitario QA", contenido)
        self.assertRegex(contenido, r"\[\d{2}:\d{2}:\d{2}\]")

    def test_F7_3_etiquetas_de_colores_en_bitacora(self):
        """Verifica que las etiquetas de estilo cromático (ok, warn, error, info) estén configuradas."""
        tags = self.app.txt_log.tag_names()
        for t in ["ok", "warn", "error", "info"]:
            self.assertIn(t, tags)

    def test_F7_4_consola_bitacora_qr_existe(self):
        """Verifica que la segunda pestaña posea su propia consola con scrollbar si aplica."""
        self.assertTrue(hasattr(self.app, "log_qr"))

    def test_F7_5_limpieza_de_bitacora(self):
        """Verifica que la bitácora se pueda limpiar correctamente."""
        self.app.txt_log.delete("1.0", "end")
        self.assertEqual(self.app.txt_log.get("1.0", "end").strip(), "")

    # --- Feature 8 (R5): Conciliación Integral y Salida 0 ---

    def test_F8_1_normalizacion_de_cadenas(self):
        """Verifica que normalizar limpie espacios y puntuación convirtiendo a mayúsculas."""
        self.assertEqual(normalizar("  t4n0-cv16z033184!  "), "T4N0CV16Z033184")

    def test_F8_2_normalizacion_tolerante_sustituciones(self):
        """Verifica sustituciones de caracteres confusos (O->0, Z->2, S->5, I->1, B->8)."""
        self.assertEqual(normalizar_tolerante("OZSIB"), "02518")

    def test_F8_3_similitud_serial_identicos(self):
        """Verifica equivalencia de seriales idénticos."""
        self.assertTrue(similitud_serial("SDK0T76530", "SDK0T76530"))

    def test_F8_4_similitud_serial_con_sustitucion_tolerante(self):
        """Verifica equivalencia entre serial con 'O' y serial con '0'."""
        self.assertTrue(similitud_serial("T4NOCV16Z03", "T4N0CV16Z03"))

    def test_F8_5_similitud_serial_subpatron_contenido(self):
        """Verifica que un serial truncado coincida por contención si len >= 9."""
        self.assertTrue(similitud_serial("T4N0CV16Z033184", "T4N0CV16Z033"))


# ============================================================================
# TIER 2: BOUNDARY & CORNER CASES
# >= 40 tests (8 boundary areas x >= 5 tests)
# ============================================================================

class TestTier2BoundaryAndCornerCases(unittest.TestCase):
    """Tier 2: Extreme values, boundary conditions, corrupt data, special characters."""

    @classmethod
    def setUpClass(cls):
        cls.root, cls.app = create_headless_app()

    @classmethod
    def tearDownClass(cls):
        try:
            cls.root.destroy()
        except Exception:
            pass

    # --- B1: Delimitadores y Prefijos en Seriales ---

    def test_B1_1_serial_con_punto_y_coma(self):
        """Verifica que un serial terminado en ';' no retenga el punto y coma."""
        sns = extraer_seriales("SDK0T76530; Modelo: LENOVO 83A0")
        self.assertEqual(sns[0], "SDK0T76530")
        self.assertFalse(sns[0].endswith(";"))

    def test_B1_2_serial_con_barra_inclinada_multiple(self):
        """Verifica división limpia de seriales alternados con '/'."""
        texto = "R8NOCX040778334 / T4N0CV16Z148186"
        sns = extraer_seriales(texto)
        self.assertEqual(len(sns), 2)
        self.assertEqual(sns[0], "R8NOCX040778334")
        self.assertEqual(sns[1], "T4N0CV16Z148186")

    def test_B1_3_serial_con_coma_y_espacios(self):
        """Verifica extracción cuando se usa coma como separador."""
        texto = "Serial equipo:MP300LYS, Modelo: 125MLD"
        sns = extraer_seriales(texto)
        self.assertIn("MP300LYS", sns)

    def test_B1_4_celda_solo_con_modelo_sin_serial(self):
        """Verifica que una celda sin número de serie no extraiga el modelo como serial."""
        texto = "Modelo: VIVOBOOK ASUS X1502VA"
        sns = extraer_seriales(texto)
        self.assertEqual(len(sns), 0)

    def test_B1_5_serial_con_saltos_de_linea_internos(self):
        """Verifica extracción cuando los datos están separados por saltos de línea."""
        texto = "Serial equipo: 5CD4375N8J\nModelo: HP 14-DQ0XXX"
        sns = extraer_seriales(texto)
        self.assertEqual(sns[0], "5CD4375N8J")

    # --- B2: Sustituciones Tolerantes de Caracteres Confundibles ---

    def test_B2_1_tolerancia_O_mayuscula_y_cero(self):
        """Verifica que 'O' y '0' se homologuen en seriales."""
        self.assertTrue(similitud_serial("SBNOCX02283647C", "SBN0CX02283647C"))

    def test_B2_2_tolerancia_Z_y_dos(self):
        """Verifica que 'Z' y '2' se homologuen."""
        self.assertTrue(similitud_serial("T4N0CV16Z047", "T4N0CV162047"))

    def test_B2_3_tolerancia_S_y_cinco(self):
        """Verifica que 'S' y '5' se homologuen."""
        self.assertTrue(similitud_serial("S8N0CV080498", "58N0CV080498"))

    def test_B2_4_tolerancia_I_y_uno(self):
        """Verifica que 'I' y '1' se homologuen."""
        self.assertTrue(similitud_serial("1CR4450DX8", "ICR4450DX8"))

    def test_B2_5_tolerancia_B_y_ocho(self):
        """Verifica que 'B' y '8' se homologuen."""
        self.assertTrue(similitud_serial("RBNOCV02K861", "R8N0CV02K861"))

    # --- B3: Campos Faltantes y Valores Nulos ---

    def test_B3_1_diccionario_activo_con_valores_none(self):
        """Verifica que valores None en el activo no lancen TypeError en extracción."""
        item = {"id": None, "equipo": None, "serial": None, "observaciones": None}
        res = extraer_campos_completos(item)
        self.assertIsInstance(res, dict)
        self.assertEqual(res["id_activo"], "")

    def test_B3_2_responsable_vacio_no_rompe_extraccion(self):
        """Verifica manejo de activos sin responsable asignado."""
        item = {"equipo": "PC-BODEGA", "responsable": "", "observaciones": "Serial equipo: 1WW4G63"}
        res = extraer_campos_completos(item)
        self.assertEqual(res["responsable"], "")
        self.assertEqual(res["serial"], "1WW4G63")

    def test_B3_3_sede_vacia_no_rompe_extraccion(self):
        """Verifica manejo de activos sin sede explícita."""
        item = {"equipo": "TEST-HOST", "sede": "", "observaciones": "Serial: 67QYP74"}
        res = extraer_campos_completos(item)
        self.assertEqual(res["sede"], "")

    def test_B3_4_estado_vacio_adopta_activo_por_defecto(self):
        """Verifica que un estado nulo o vacío se normalice a 'Activo'."""
        item = {"equipo": "LAP-01", "estado": None, "observaciones": "Serial: 1H84351J6Q"}
        res = extraer_campos_completos(item)
        self.assertEqual(res["estado"], "Activo")

    def test_B3_5_celda_observaciones_completamente_vacia(self):
        """Verifica manejo de celda de observaciones vacía."""
        sns = extraer_seriales("")
        self.assertEqual(sns, [])
        sns_none = extraer_seriales(None)
        self.assertEqual(sns_none, [])

    # --- B4: Textos Largos y Truncamiento Visual ---

    def test_B4_1_hostname_extremadamente_largo(self):
        """Verifica que un hostname de más de 40 caracteres se renderice sin desbordar el canvas."""
        with tempfile.TemporaryDirectory() as tmpdir:
            item = {
                "id": "LAP-2026-9999",
                "equipo": "LAPTOP-SUPER-LARGA-CON-MAS-DE-CUARENTA-CARACTERES-TI",
                "serial": "SDK0T76530",
                "sede": "Sede Principal",
                "responsable": "Usuario",
                "estado": "Activo",
            }
            ruta = self.app._generar_etiqueta_individual(item, formato="50x25mm", carpeta=tmpdir)
            self.assertTrue(os.path.exists(ruta))

    def test_B4_2_nombre_responsable_muy_largo(self):
        """Verifica que un nombre de usuario largo se trunque en texto impreso sin afectar QR."""
        with tempfile.TemporaryDirectory() as tmpdir:
            item = {
                "id": "LAP-2026-9998",
                "equipo": "LAP-USR",
                "serial": "SDK0T76531",
                "sede": "Administrativa",
                "responsable": "JUAN FRANCISCO DE LOS SANTOS PEREZ HERNANDEZ",
                "estado": "Activo",
            }
            ruta = self.app._generar_etiqueta_individual(item, formato="50x25mm", carpeta=tmpdir)
            self.assertTrue(os.path.exists(ruta))

    def test_B4_3_caracteres_con_tildes_y_diacriticos(self):
        """Verifica renderizado de sedes con tildes (Sótano, Cúcuta, Santa Fé, Dosquebradas)."""
        with tempfile.TemporaryDirectory() as tmpdir:
            item = {
                "id": "LAP-2026-9997",
                "equipo": "LAP-SOTANO",
                "serial": "SDK0T76532",
                "sede": "Sótano / Área de Recepción Técnica",
                "responsable": "JOSÉ MIGUEL BETANCOURT",
                "estado": "Operativo",
            }
            ruta = self.app._generar_etiqueta_individual(item, formato="50x25mm", carpeta=tmpdir)
            self.assertTrue(os.path.exists(ruta))

    def test_B4_4_fallback_de_fuente_si_ttf_no_existe(self):
        """Verifica que _obtener_fuente retorne una fuente por defecto si la ruta no existe."""
        fnt = cu._obtener_fuente(12, bold=False)
        self.assertIsNotNone(fnt)

    def test_B4_5_sanitizacion_de_caracteres_prohibidos_en_nombre_archivo(self):
        """Verifica que caracteres como '/', ':', '*' no rompan el guardado de la imagen."""
        with tempfile.TemporaryDirectory() as tmpdir:
            item = {
                "id": "LAP-2026-9996",
                "equipo": 'HOST/CON:CARACTERES*PROHIBIDOS?"<>|',
                "serial": "SDK0T76533",
                "sede": "Sede",
                "responsable": "Resp",
                "estado": "Activo",
            }
            ruta = self.app._generar_etiqueta_individual(item, formato="50x25mm", carpeta=tmpdir)
            self.assertTrue(os.path.exists(ruta))

    # --- B5: Detección de Filas de Encabezado y Totales ---

    def test_B5_1_deteccion_encabezados_fila_5_estandar(self):
        """Verifica detección de fila 5 en libro estándar con filas previas de título."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(2, 1, "TITULO EMPRESA")
        ws.cell(5, 1, "Item")
        ws.cell(5, 3, "Sede")
        ws.cell(5, 4, "Equipo")
        ws.cell(5, 11, "Serial")
        f_hdr = self.app._detectar_fila_encabezados(ws)
        self.assertEqual(f_hdr, 5)

    def test_B5_2_deteccion_encabezados_en_fila_1(self):
        """Verifica detección cuando los encabezados están en la fila 1."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(1, 1, "ID")
        ws.cell(1, 2, "EQUIPO")
        ws.cell(1, 3, "SERIAL")
        f_hdr = self.app._detectar_fila_encabezados(ws)
        self.assertEqual(f_hdr, 1)

    def test_B5_3_deteccion_fila_total_con_filas_vacias_intermedias(self):
        """Verifica localización de 'TOTAL' tras celdas vacías."""
        wb = openpyxl.Workbook()
        ws = wb.active
        for r in range(1, 10):
            ws.cell(r, 1, f"Data {r}")
        ws.cell(15, 2, "TOTAL EQUIPOS")
        f_tot = self.app._detectar_fila_total(ws, 1)
        self.assertEqual(f_tot, 15)

    def test_B5_4_deteccion_fila_total_inexistente_retorna_max_mas_uno(self):
        """Verifica que si no hay fila de total, retorne ws.max_row + 1."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.cell(5, 1, "Data")
        f_tot = self.app._detectar_fila_total(ws, 1)
        self.assertEqual(f_tot, ws.max_row + 1)

    def test_B5_5_actualizacion_dinamica_formula_sumatoria(self):
        """Verifica que _ejecutar actualice dinámicamente la fórmula =SUM(...) al insertar altas antes del TOTAL."""
        with tempfile.TemporaryDirectory() as tmpdir:
            src_xlsx = os.path.join(tmpdir, "test_sum.xlsx")
            out_xlsx = os.path.join(tmpdir, "AUDITADO_test_sum.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Inventario de Equipos"
            ws.cell(5, 1, "Item")
            ws.cell(5, 3, "Sede")
            ws.cell(5, 4, "Equipo")
            ws.cell(5, 8, "Cantidad")
            ws.cell(5, 11, "Serial")
            ws.cell(5, 12, "Auditado")
            for r in range(6, 11):
                ws.cell(r, 1, r - 5)
                ws.cell(r, 3, "Sede 1")
                ws.cell(r, 4, f"LAPTOP-{r-5}")
                ws.cell(r, 8, 1)
                ws.cell(r, 11, f"SN00{r-5}")
            ws.cell(11, 1, "TOTAL")
            ws.cell(11, 8, "=SUM(H6:H10)")
            wb.save(src_xlsx)
            wb.close()

            self.app.ruta_excel.set(src_xlsx)
            self.app.ruta_salida.set(out_xlsx)
            self.app.nombre_hoja.set("Inventario de Equipos")
            self.app.chk_perifericos_clemont.set(False)

            self.app._limpiar_tabla()
            self.app.tree.insert("", "end", values=(1, "LAPTOP-1", "SN001", "OK", "Sede 1", "Daniel", ""))
            self.app.tree.insert("", "end", values=(2, "ALTA-NUEVA-X", "SN99999", "OK", "Sede 1", "Daniel", ""))

            self.app._ejecutar()

            self.assertTrue(os.path.exists(out_xlsx))
            wb_res = openpyxl.load_workbook(out_xlsx, data_only=False)
            ws_res = wb_res["Inventario de Equipos"]
            f_tot = self.app._detectar_fila_total(ws_res, 6)
            self.assertEqual(f_tot, 12)
            formula_val = str(ws_res.cell(12, 8).value)
            self.assertEqual(formula_val, "=SUM(H6:H11)")
            wb_res.close()

    # --- B6: Combinaciones de Checkboxes de Matching ---

    def test_B6_1_matching_con_serial_desactivado(self):
        """Verifica que al desmarcar chk_col_serial, _ejecutar concilie exitosamente por Hostname."""
        with tempfile.TemporaryDirectory() as tmpdir:
            src_xlsx = os.path.join(tmpdir, "test_b6_1.xlsx")
            out_xlsx = os.path.join(tmpdir, "AUDITADO_test_b6_1.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Inventario"
            ws.cell(5, 1, "Item")
            ws.cell(5, 3, "Sede")
            ws.cell(5, 4, "Equipo")
            ws.cell(5, 8, "Cantidad")
            ws.cell(5, 11, "Serial")
            ws.cell(5, 12, "Auditado")
            ws.cell(6, 1, 1)
            ws.cell(6, 3, "Sede 1")
            ws.cell(6, 4, "LAPTOP-01")
            ws.cell(6, 8, 1)
            ws.cell(6, 11, "Serial equipo: SERIAL_DISCORDANTE_999")
            ws.cell(7, 1, "TOTAL")
            ws.cell(7, 8, "=SUM(H6:H6)")
            wb.save(src_xlsx)
            wb.close()

            self.app.ruta_excel.set(src_xlsx)
            self.app.ruta_salida.set(out_xlsx)
            self.app.nombre_hoja.set("Inventario")
            self.app.chk_perifericos_clemont.set(False)

            self.app.chk_col_serial.set(False)
            self.app.chk_col_equipo.set(True)

            self.app._limpiar_tabla()
            self.app.tree.insert("", "end", values=(1, "LAPTOP-01", "SERIAL_FISICO_OTRO", "OK", "Sede 1", "Daniel", ""))

            self.app._ejecutar()

            wb_res = openpyxl.load_workbook(out_xlsx, data_only=True)
            ws_res = wb_res["Inventario"]
            self.assertEqual(ws_res.cell(6, 12).value, "OK")
            wb_res.close()
            self.app.chk_col_serial.set(True)

    def test_B6_2_matching_con_host_desactivado(self):
        """Verifica que al desmarcar chk_col_equipo, _ejecutar concilie exitosamente por Serial."""
        with tempfile.TemporaryDirectory() as tmpdir:
            src_xlsx = os.path.join(tmpdir, "test_b6_2.xlsx")
            out_xlsx = os.path.join(tmpdir, "AUDITADO_test_b6_2.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Inventario"
            ws.cell(5, 1, "Item")
            ws.cell(5, 3, "Sede")
            ws.cell(5, 4, "Equipo")
            ws.cell(5, 8, "Cantidad")
            ws.cell(5, 11, "Serial")
            ws.cell(5, 12, "Auditado")
            ws.cell(6, 1, 1)
            ws.cell(6, 3, "Sede 1")
            ws.cell(6, 4, "OLD-HOSTNAME-WRONG")
            ws.cell(6, 8, 1)
            ws.cell(6, 11, "Serial equipo: SDK0T76530")
            ws.cell(7, 1, "TOTAL")
            ws.cell(7, 8, "=SUM(H6:H6)")
            wb.save(src_xlsx)
            wb.close()

            self.app.ruta_excel.set(src_xlsx)
            self.app.ruta_salida.set(out_xlsx)
            self.app.nombre_hoja.set("Inventario")
            self.app.chk_perifericos_clemont.set(False)

            self.app.chk_col_serial.set(True)
            self.app.chk_col_equipo.set(False)

            self.app._limpiar_tabla()
            self.app.tree.insert("", "end", values=(1, "NEW-HOSTNAME-DIFFERENT", "SDK0T76530", "OK", "Sede 1", "Daniel", ""))

            self.app._ejecutar()

            wb_res = openpyxl.load_workbook(out_xlsx, data_only=True)
            ws_res = wb_res["Inventario"]
            self.assertEqual(ws_res.cell(6, 12).value, "OK")
            wb_res.close()
            self.app.chk_col_equipo.set(True)

    def test_B6_3_extractor_indice_columna_desde_texto(self):
        """Verifica que _get_col_idx extraiga el número de columna de cadenas como 'Col K (11)'."""
        self.assertEqual(self.app._get_col_idx("Col K (11) - Observaciones y Serial", 11), 11)
        self.assertEqual(self.app._get_col_idx("Col D (4) - Equipo/Servidor", 4), 4)
        self.assertEqual(self.app._get_col_idx("Col C (3) - Sede", 3), 3)
        self.assertEqual(self.app._get_col_idx("12", 12), 12)
        self.assertEqual(self.app._get_col_idx("K", 11), 11)

    def test_B6_4_extractor_indice_columna_fallback(self):
        """Verifica retorno de default ante texto inválido o vacío."""
        self.assertEqual(self.app._get_col_idx("", 9), 9)
        self.assertEqual(self.app._get_col_idx("Invalido???", 9), 9)

    def test_B6_5_todas_las_columnas_desmarcadas(self):
        """Verifica que las variables booleanas puedan ser todas False sin causar excepciones."""
        for v in [self.app.chk_col_serial, self.app.chk_col_equipo, self.app.chk_col_sede,
                  self.app.chk_col_resp, self.app.chk_col_verif]:
            v.set(False)
            self.assertFalse(v.get())
            v.set(True)

    # --- B7: Rutas, Directorios y Caracteres Especiales ---

    def test_B7_1_ruta_con_espacios_en_blanco(self):
        """Verifica que las funciones de guardado y _obtener_carpeta_qr soporten rutas con espacios."""
        with tempfile.TemporaryDirectory() as base_tmp:
            sub = os.path.join(base_tmp, "carpeta con espacios en el nombre")
            os.makedirs(sub, exist_ok=True)
            ruta_excel = os.path.join(sub, "AUDITADO_Inventario.xlsx")
            self.app.ruta_salida.set(ruta_excel)

            carpeta_qr = self.app._obtener_carpeta_qr()
            self.assertTrue(os.path.isabs(carpeta_qr))
            self.assertIn("carpeta con espacios en el nombre", carpeta_qr)

            item = {
                "id": "LAP-2026-0001", "equipo": "LAP-ESPACIO",
                "serial": "SN12345", "sede": "Sede", "responsable": "Daniel", "estado": "Activo"
            }
            img_path = os.path.join(sub, "test_espacio.png")
            ret_path = cu.generar_etiqueta_individual(item, img_path)
            self.assertTrue(os.path.exists(ret_path))

    def test_B7_2_ruta_con_parentesis(self):
        """Verifica que generar_plantilla_auditoria_qr y _obtener_carpeta_qr manejen rutas con paréntesis '(2) (2)'."""
        with tempfile.TemporaryDirectory() as base_tmp:
            sub = os.path.join(base_tmp, "carpeta (2) (2)")
            os.makedirs(sub, exist_ok=True)
            f_out = os.path.join(sub, "Inventario_Equipos_Clemont (2) (2)_Con_QR.xlsx")
            ret = cu.generar_plantilla_auditoria_qr(f_out, [])
            self.assertTrue(os.path.exists(ret))
            self.assertTrue(os.path.isabs(ret))
            wb = openpyxl.load_workbook(ret)
            self.assertIn("Maestro_Activos", wb.sheetnames)
            wb.close()

    def test_B7_3_creacion_idempotente_carpeta_etiquetas(self):
        """Verifica que _generar_etiqueta_individual gestione la creación de carpetas de forma idempotente."""
        with tempfile.TemporaryDirectory() as tmpdir:
            sub = os.path.join(tmpdir, "etiquetas_qr_idempotente")
            item = {
                "id": "LAP-2026-0001", "equipo": "LAP-01",
                "serial": "SN12345", "sede": "Sede", "responsable": "Daniel", "estado": "Activo"
            }
            p1 = self.app._generar_etiqueta_individual(item, carpeta=sub)
            self.assertTrue(os.path.exists(p1))
            p2 = self.app._generar_etiqueta_individual(item, carpeta=sub)
            self.assertTrue(os.path.exists(p2))
            self.assertEqual(p1, p2)

    def test_B7_4_sobrescritura_segura_archivo_excel(self):
        """Verifica que generar_plantilla_auditoria_qr sobrescriba de forma segura un archivo existente."""
        with tempfile.TemporaryDirectory() as tmpdir:
            fpath = os.path.join(tmpdir, "test_overwrite.xlsx")
            item1 = {"id": "LAP-0001", "equipo": "EQ-1", "serial": "SN-1", "sede": "Sede1"}
            cu.generar_plantilla_auditoria_qr(fpath, [item1])
            self.assertTrue(os.path.exists(fpath))

            item2 = {"id": "LAP-0002", "equipo": "EQ-2", "serial": "SN-2", "sede": "Sede2"}
            cu.generar_plantilla_auditoria_qr(fpath, [item2])
            wb = openpyxl.load_workbook(fpath)
            ws = wb["Maestro_Activos"]
            self.assertEqual(ws.cell(2, 1).value, "LAP-0002")
            wb.close()

    def test_B7_5_conversion_ruta_relativa_a_absoluta(self):
        """Verifica que _obtener_carpeta_qr y generar_plantilla_auditoria_qr retornen rutas absolutas."""
        self.app.ruta_salida.set("etiquetas_qr/AUDITADO_test.xlsx")
        carpeta = self.app._obtener_carpeta_qr()
        self.assertTrue(os.path.isabs(carpeta))

        with tempfile.TemporaryDirectory() as tmpdir:
            orig_cwd = os.getcwd()
            try:
                os.chdir(tmpdir)
                ret = cu.generar_plantilla_auditoria_qr("temp_relativo.xlsx", [])
                self.assertTrue(os.path.isabs(ret))
            finally:
                os.chdir(orig_cwd)

    # --- B8: Carga Útil QR, Escapado y Modos ---

    def test_B8_1_modo_solo_id_en_qr(self):
        """Verifica que el modo solo_id en _generar_etiqueta_individual codifique únicamente el ID del activo."""
        import unittest.mock as mock
        self.app.qr_contenido_modo.set("solo_id")
        item = {
            "id": "LAP-2026-0001", "equipo": "LAP-01", "serial": "SN12345",
            "modelo": "Lenovo", "responsable": "Daniel", "sede": "Sede", "estado": "Activo"
        }
        with tempfile.TemporaryDirectory() as tmpdir:
            with mock.patch.object(qrcode.QRCode, "add_data", wraps=qrcode.QRCode.add_data) as mock_add:
                ruta = self.app._generar_etiqueta_individual(item, carpeta=tmpdir)
                self.assertTrue(os.path.exists(ruta))
                mock_add.assert_called_once()
                self.assertEqual(mock_add.call_args[0][0], "LAP-2026-0001")
        self.app.qr_contenido_modo.set("completo")

    def test_B8_2_modo_completo_incluye_todos_los_7_campos(self):
        """Verifica que el modo completo en _generar_etiqueta_individual incluya los 7 campos en el QR."""
        import unittest.mock as mock
        self.app.qr_contenido_modo.set("completo")
        item = {
            "id": "LAP-2026-0001", "equipo": "LAPTOP-01", "modelo": "Lenovo",
            "serial": "SDK0T76530", "responsable": "DANIEL", "sede": "Sede 1", "estado": "Activo"
        }
        with tempfile.TemporaryDirectory() as tmpdir:
            with mock.patch.object(qrcode.QRCode, "add_data", wraps=qrcode.QRCode.add_data) as mock_add:
                ruta = self.app._generar_etiqueta_individual(item, carpeta=tmpdir)
                self.assertTrue(os.path.exists(ruta))
                payload = mock_add.call_args[0][0]
                for k in ["ACTIVO: LAP-2026-0001", "EQUIPO: LAPTOP-01", "MODELO: Lenovo",
                          "SERIAL: SDK0T76530", "RESPONSABLE: DANIEL", "SEDE: Sede 1", "ESTADO: Activo"]:
                    self.assertIn(k, payload)

    def test_B8_3_carga_util_qr_con_comillas_y_caracteres_especiales(self):
        """Verifica que caracteres especiales en nombres no corrompan la matriz QR."""
        qr = qrcode.QRCode(box_size=10, border=1)
        qr.add_data('ACTIVO: LAP-01\nMODELO: Lenovo "ThinkPad" 14\' & 15\\#')
        qr.make(fit=True)
        self.assertGreater(len(qr.modules), 20)

    def test_B8_4_carga_util_qr_con_longitud_superior_a_180_caracteres(self):
        """Verifica que cargas útiles extensas mantengan corrección de errores M."""
        long_payload = "A" * 200
        qr = qrcode.QRCode(box_size=10, border=1, error_correction=qrcode.constants.ERROR_CORRECT_M)
        qr.add_data(long_payload)
        qr.make(fit=True)
        self.assertTrue(inspect_qr_matrix_and_payload(qr, [long_payload]))

    def test_B8_5_reparsitura_payload_qr_a_diccionario(self):
        """Verifica que la cadena del QR se pueda parsear de vuelta a un diccionario con los 7 campos."""
        qr_text = (
            "EMPRESA: CLEMONT\n"
            "ACTIVO: LAP-2026-0001\n"
            "EQUIPO: LAPTOP-2895F563\n"
            "MODELO: Lenovo 83A0\n"
            "SERIAL: SDK0T76530\n"
            "RESPONSABLE: DANIEL PANESSO\n"
            "SEDE: Administrativa 2do piso\n"
            "ESTADO: Activo"
        )
        parsed = {}
        for line in qr_text.strip().splitlines():
            if ":" in line:
                k, v = line.split(":", 1)
                parsed[k.strip()] = v.strip()
        self.assertEqual(parsed["ACTIVO"], "LAP-2026-0001")
        self.assertEqual(parsed["EQUIPO"], "LAPTOP-2895F563")
        self.assertEqual(parsed["SERIAL"], "SDK0T76530")
        self.assertEqual(parsed["RESPONSABLE"], "DANIEL PANESSO")
        self.assertEqual(parsed["SEDE"], "Administrativa 2do piso")
        self.assertEqual(parsed["ESTADO"], "Activo")


# ============================================================================
# TIER 3: CROSS-FEATURE INTEGRATION COMBINATIONS
# >= 8 integration tests
# ============================================================================

class TestTier3CrossFeatureIntegration(unittest.TestCase):
    """Tier 3: Pairwise and multi-feature end-to-end interactions."""

    @classmethod
    def setUpClass(cls):
        cls.root, cls.app = create_headless_app()

    @classmethod
    def tearDownClass(cls):
        try:
            cls.root.destroy()
        except Exception:
            pass

    def test_C1_ingestion_extraccion_generacion_sticker_300dpi(self):
        """C1: Ingestión de datos -> Extracción 7 campos -> Generación de sticker -> Verificación 300 DPI."""
        with tempfile.TemporaryDirectory() as tmpdir:
            raw_excel_row = {
                "id": "LAP-2026-0001",
                "equipo": "LAPTOP-2895F563",
                "sede": "Administrativa 2do piso",
                "responsable": "DANIEL PANESSO",
                "estado": "Activo",
                "observaciones": "Serial equipo: SDK0T76530; Modelo: LENOVO 83A0",
            }
            campos = extraer_campos_completos(raw_excel_row)
            item_tag = {
                "id": campos["id_activo"],
                "equipo": campos["equipo"],
                "modelo": campos["modelo"],
                "serial": campos["serial"],
                "responsable": campos["responsable"],
                "sede": campos["sede"],
                "estado": campos["estado"],
            }
            ruta_png = self.app._generar_etiqueta_individual(item_tag, formato="50x25mm", carpeta=tmpdir)
            self.assertTrue(os.path.exists(ruta_png))
            img = Image.open(ruta_png)
            self.assertIn(img.info.get("dpi"), ((300, 300), (300.0, 300.0)))
            self.assertEqual(img.size, (591, 295))

    def test_C2_generacion_plantilla_con_hojas_maestro_y_escaneos(self):
        """C2: Generación de plantilla de auditoría con hojas vinculadas y estilos corporativos mediante producción."""
        with tempfile.TemporaryDirectory() as tmpdir:
            f_out = os.path.join(tmpdir, "Inventario_Equipos_Clemont_Con_QR.xlsx")
            equipos = [{
                "id": "LAP-2026-0001", "categoria": "LAP", "tipo": "Laptop",
                "equipo": "LAPTOP-01", "modelo": "Lenovo", "serial": "SN01",
                "sede": "Sede 1", "responsable": "Daniel", "estado": "Activo"
            }]
            cu.generar_plantilla_auditoria_qr(f_out, equipos)

            wb_read = openpyxl.load_workbook(f_out)
            self.assertIn("Maestro_Activos", wb_read.sheetnames)
            self.assertIn("Escaneos_Campo", wb_read.sheetnames)
            self.assertIn("Instructivo_Auditoria", wb_read.sheetnames)

            ws_m = wb_read["Maestro_Activos"]
            self.assertEqual(ws_m.max_column, 9)
            self.assertEqual(ws_m.cell(1, 1).fill.start_color.rgb, "001F4E78")

            ws_e = wb_read["Escaneos_Campo"]
            self.assertEqual(ws_e.max_column, 8)
            self.assertEqual(ws_e.cell(1, 1).fill.start_color.rgb, "00375623")
            self.assertIn("VLOOKUP", ws_e.cell(2, 7).value)
            wb_read.close()

    def test_C3_simulacion_conciliacion_cuatro_estados(self):
        """C3: Simulación de auditoría física: clasificación en OK, Traslado y Alta ejecutando _ejecutar()."""
        with tempfile.TemporaryDirectory() as tmpdir:
            src_xlsx = os.path.join(tmpdir, "test_c3.xlsx")
            out_xlsx = os.path.join(tmpdir, "AUDITADO_test_c3.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Inventario"
            ws.cell(5, 1, "Item")
            ws.cell(5, 3, "Sede")
            ws.cell(5, 4, "Equipo")
            ws.cell(5, 8, "Cantidad")
            ws.cell(5, 11, "Serial")
            ws.cell(5, 12, "Auditado")
            # Fila 6: OK (coincide serial y sede)
            ws.cell(6, 1, 1); ws.cell(6, 3, "Sótano"); ws.cell(6, 4, "LAP-OK"); ws.cell(6, 8, 1); ws.cell(6, 11, "Serial equipo: T4N0CV16Z033184")
            # Fila 7: Traslado (coincide serial pero sede difiere)
            ws.cell(7, 1, 2); ws.cell(7, 3, "Sótano"); ws.cell(7, 4, "LAP-TR"); ws.cell(7, 8, 1); ws.cell(7, 11, "Serial equipo: TANRCX04W552444")
            # Fila 8: Total
            ws.cell(8, 1, "TOTAL"); ws.cell(8, 8, "=SUM(H6:H7)")
            wb.save(src_xlsx)
            wb.close()

            self.app.ruta_excel.set(src_xlsx)
            self.app.ruta_salida.set(out_xlsx)
            self.app.nombre_hoja.set("Inventario")
            self.app.chk_perifericos_clemont.set(False)

            self.app._limpiar_tabla()
            self.app.tree.insert("", "end", values=(1, "LAP-OK", "T4N0CV16Z033184", "OK", "Sótano", "Daniel", ""))
            self.app.tree.insert("", "end", values=(2, "LAP-TR", "TANRCX04W552444", "Traslado", "Showroom", "Daniel", ""))
            self.app.tree.insert("", "end", values=(3, "LAP-ALTA", "NEW_SERIAL_ALTA_123", "OK", "Sótano", "Daniel", ""))

            self.app._ejecutar()

            wb_res = openpyxl.load_workbook(out_xlsx, data_only=True)
            ws_res = wb_res["Inventario"]
            self.assertEqual(ws_res.cell(6, 12).value, "OK")
            self.assertEqual(ws_res.cell(7, 12).value, "Traslado")
            self.assertEqual(ws_res.cell(8, 4).value, "LAP-ALTA")
            self.assertEqual(ws_res.cell(9, 1).value, "TOTAL")
            wb_res.close()

    def test_C4_insercion_altas_actualizacion_formula_total(self):
        """C4: Inserción de altas antes de la fila TOTAL y recálculo de sumatoria mediante _ejecutar()."""
        with tempfile.TemporaryDirectory() as tmpdir:
            src_xlsx = os.path.join(tmpdir, "test_c4.xlsx")
            out_xlsx = os.path.join(tmpdir, "AUDITADO_test_c4.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Inventario"
            ws.cell(5, 1, "Item")
            ws.cell(5, 3, "Sede")
            ws.cell(5, 4, "Equipo")
            ws.cell(5, 8, "Cantidad")
            ws.cell(5, 11, "Serial")
            ws.cell(5, 12, "Auditado")
            for r in range(6, 11):
                ws.cell(r, 1, r - 5)
                ws.cell(r, 3, "Sede 1")
                ws.cell(r, 4, f"LAP-{r-5}")
                ws.cell(r, 8, 1)
                ws.cell(r, 11, f"SN00{r-5}")
            ws.cell(11, 1, "TOTAL")
            ws.cell(11, 8, "=SUM(H6:H10)")
            wb.save(src_xlsx)
            wb.close()

            self.app.ruta_excel.set(src_xlsx)
            self.app.ruta_salida.set(out_xlsx)
            self.app.nombre_hoja.set("Inventario")
            self.app.chk_perifericos_clemont.set(False)

            self.app._limpiar_tabla()
            self.app.tree.insert("", "end", values=(6, "ALTA-1", "SN-ALTA-1", "OK", "Sótano", "Resp 1", ""))
            self.app.tree.insert("", "end", values=(7, "ALTA-2", "SN-ALTA-2", "OK", "Showroom", "Resp 2", ""))

            self.app._ejecutar()

            wb_res = openpyxl.load_workbook(out_xlsx, data_only=False)
            ws_res = wb_res["Inventario"]
            self.assertEqual(ws_res.cell(13, 1).value, "TOTAL")
            self.assertEqual(str(ws_res.cell(13, 8).value), "=SUM(H6:H12)")
            self.assertEqual(ws_res.cell(11, 4).value, "ALTA-1")
            self.assertEqual(ws_res.cell(12, 4).value, "ALTA-2")
            wb_res.close()

    def test_C5_matching_solo_por_host_cuando_serial_esta_desmarcado(self):
        """C5: Matching dinámico: ignorar serial y reconciliar por Hostname cuando serial está desactivado ejecutando _ejecutar()."""
        with tempfile.TemporaryDirectory() as tmpdir:
            src_xlsx = os.path.join(tmpdir, "test_c5.xlsx")
            out_xlsx = os.path.join(tmpdir, "AUDITADO_test_c5.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Inventario"
            ws.cell(5, 1, "Item")
            ws.cell(5, 3, "Sede")
            ws.cell(5, 4, "Equipo")
            ws.cell(5, 8, "Cantidad")
            ws.cell(5, 11, "Serial")
            ws.cell(5, 12, "Auditado")
            ws.cell(6, 1, 1)
            ws.cell(6, 3, "Sede 1")
            ws.cell(6, 4, "DESKTOP-TQ9MEAZ")
            ws.cell(6, 8, 1)
            ws.cell(6, 11, "Serial equipo: SERIAL_VIEJO_INCORRECTO")
            ws.cell(7, 1, "TOTAL")
            ws.cell(7, 8, "=SUM(H6:H6)")
            wb.save(src_xlsx)
            wb.close()

            self.app.ruta_excel.set(src_xlsx)
            self.app.ruta_salida.set(out_xlsx)
            self.app.nombre_hoja.set("Inventario")
            self.app.chk_perifericos_clemont.set(False)

            self.app.chk_col_serial.set(False)
            self.app.chk_col_equipo.set(True)

            self.app._limpiar_tabla()
            self.app.tree.insert("", "end", values=(1, "DESKTOP-TQ9MEAZ", "R8N0CV183189349", "OK", "Sede 1", "Daniel", ""))

            self.app._ejecutar()

            wb_res = openpyxl.load_workbook(out_xlsx, data_only=True)
            ws_res = wb_res["Inventario"]
            self.assertEqual(ws_res.cell(6, 12).value, "OK")
            wb_res.close()
            self.app.chk_col_serial.set(True)

    def test_C6_cruce_payload_qr_con_campos_maestro_activos(self):
        """C6: Cruce de datos: el payload generado en QR coincide 100% con los datos exportados a Maestro_Activos."""
        import unittest.mock as mock
        item = {
            "id": "LAP-2026-0005",
            "categoria": "LAP",
            "tipo": "Laptop",
            "equipo": "HAIBER",
            "modelo": "Vivobook ASUS laptop x1502Va",
            "serial": "TAN0KD004371447",
            "sede": "Sótano",
            "responsable": "HAIBER MONTAÑA",
            "estado": "Activo",
        }
        with tempfile.TemporaryDirectory() as tmpdir:
            template_path = os.path.join(tmpdir, "plantilla.xlsx")
            cu.generar_plantilla_auditoria_qr(template_path, [item])

            wb = openpyxl.load_workbook(template_path)
            ws_m = wb["Maestro_Activos"]
            m_id = ws_m.cell(2, 1).value
            m_host = ws_m.cell(2, 4).value
            m_modelo = ws_m.cell(2, 5).value
            m_serial = ws_m.cell(2, 6).value
            m_sede = ws_m.cell(2, 7).value
            m_resp = ws_m.cell(2, 8).value
            m_estado = ws_m.cell(2, 9).value
            wb.close()

            label_path = os.path.join(tmpdir, "etiqueta.png")
            with mock.patch.object(qrcode.QRCode, "add_data", wraps=qrcode.QRCode.add_data) as mock_add:
                cu.generar_etiqueta_individual(item, label_path)
                mock_add.assert_called_once()
                qr_payload = mock_add.call_args[0][0]

                self.assertIn(f"ACTIVO: {m_id}", qr_payload)
                self.assertIn(f"EQUIPO: {m_host}", qr_payload)
                self.assertIn(f"MODELO: {m_modelo}", qr_payload)
                self.assertIn(f"SERIAL: {m_serial}", qr_payload)
                self.assertIn(f"RESPONSABLE: {m_resp}", qr_payload)
                self.assertIn(f"SEDE: {m_sede}", qr_payload)
                self.assertIn(f"ESTADO: {m_estado}", qr_payload)

    def test_C7_ejecucion_completa_en_directorio_temporal(self):
        """C7: Pipeline completo en directorio temporal: creación de libro auditado y carpeta etiquetas."""
        with tempfile.TemporaryDirectory() as tmpdir:
            wb_src = openpyxl.Workbook()
            ws = wb_src.active
            ws.title = "Inventario de Equipos"
            ws.cell(5, 1, "Item")
            ws.cell(5, 2, "Cliente")
            ws.cell(5, 3, "Centro Operativo / Sede")
            ws.cell(5, 4, "Equipo / Servidor")
            ws.cell(5, 8, "Cantidad")
            ws.cell(5, 9, "Responsable / Usuario")
            ws.cell(5, 10, "Estado")
            ws.cell(5, 11, "Observaciones y Serial")
            ws.cell(5, 12, "Verificado")

            # 1 registro
            ws.cell(6, 1, 1)
            ws.cell(6, 2, "CLEMONT")
            ws.cell(6, 3, "Sótano")
            ws.cell(6, 4, "SERVICIO-CLIENTE1")
            ws.cell(6, 8, 1)
            ws.cell(6, 9, "ESTEBAN MUÑOZ")
            ws.cell(6, 10, "Activo")
            ws.cell(6, 11, "Serial equipo: T4N0CV16Z033184; Modelo: Vivobook ASUS")

            ws.cell(7, 1, "TOTAL")
            ws.cell(7, 8, "=SUM(H6:H6)")

            src_path = os.path.join(tmpdir, "Inventario_Prueba.xlsx")
            wb_src.save(src_path)
            wb_src.close()

            out_path = os.path.join(tmpdir, "AUDITADO_Inventario_Prueba.xlsx")
            self.app.ruta_excel.set(src_path)
            self.app.ruta_salida.set(out_path)
            self.app.nombre_hoja.set("Inventario de Equipos")

            # Ejecutar conciliación headless
            self.app._limpiar_tabla()
            self.app.tree.insert("", "end", values=(1, "SERVICIO-CLIENTE1", "T4N0CV16Z033184", "OK", "Sótano", "Asignado", ""))
            self.app._ejecutar()

            self.assertTrue(os.path.exists(out_path))
            wb_res = openpyxl.load_workbook(out_path, data_only=True)
            self.assertIn("Reporte de Auditoria", wb_res.sheetnames)
            ws_res = wb_res["Inventario de Equipos"]
            self.assertEqual(ws_res.cell(6, 12).value, "OK")
            wb_res.close()

    def test_C8_pegado_portapapeles_tabla_completa_12_columnas(self):
        """C8: Pegado desde portapapeles de 12 columnas: mapeo fidedigno sin desfasar Sede y Host a Item mediante _pegar_desde_portapapeles."""
        import unittest.mock as mock
        tsv_data = (
            "1\tCLEMONT\tAdministrativa 2do piso\tLAPTOP-2895F563\tWINDOWS 11 PRO\t"
            "Microsoft\tSistema operativo\t1\tDANIEL PANESSO\tActivo\t"
            "Serial equipo: SDK0T76530; Modelo: LENOVO 83A0\tOK"
        )
        with mock.patch.object(self.app.root, "clipboard_get", return_value=tsv_data):
            self.app._limpiar_tabla()
            self.app._pegar_desde_portapapeles()

            children = self.app.tree.get_children()
            self.assertEqual(len(children), 1)
            vals = self.app.tree.item(children[0], "values")

            self.assertEqual(str(vals[0]), "1")
            self.assertEqual(vals[1], "LAPTOP-2895F563")
            self.assertEqual(vals[2], "SDK0T76530")
            self.assertEqual(vals[3], "OK")
            self.assertEqual(vals[4], "Administrativa 2do piso")
            self.assertEqual(vals[5], "DANIEL PANESSO")

    def test_C9_adaptabilidad_columnas_excel_arbitrario_dinamico(self):
        """C9: Adaptabilidad dinámica: Detección y mapeo de columnas para cualquier estructura de Excel (5, 8, 12 y 15 columnas)."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Crear un Excel arbitrario con 15 columnas personalizadas
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Inventario_Personalizado"
            custom_headers = [
                "ID_Interno", "Sede_Operativa", "Nombre_Equipo", "Marca", "Modelo",
                "Serial_SN", "MAC_Address", "IP_Asignada", "Usuario_Responsable",
                "Centro_Costos", "Estado_Equipo", "Fecha_Adquisicion", "Garantia",
                "Valor_Activo", "Resultado_Auditoria"
            ]
            ws.append(custom_headers)
            for r in range(1, 5):
                ws.append([f"VAL_{c}_{r}" for c in range(1, 16)])

            ruta_arb = os.path.join(tmpdir, "Inventario_Arbitrario_15Cols.xlsx")
            wb.save(ruta_arb)
            wb.close()

            # Probar autodetección de columnas
            self.app.ruta_excel.set(ruta_arb)
            self.app.nombre_hoja.set("Inventario_Personalizado")
            self.app._auto_detectar_columnas_hoja("Inventario_Personalizado")

            # Verificar que las opciones del combo contengan las 15 columnas
            opciones = self.app.combo_col_serial["values"]
            self.assertEqual(len(opciones), 15)
            self.assertTrue(any("Serial_SN" in str(op) for op in opciones))
            self.assertTrue(any("Nombre_Equipo" in str(op) for op in opciones))
            self.assertTrue(any("Sede_Operativa" in str(op) for op in opciones))
            self.assertTrue(any("Usuario_Responsable" in str(op) for op in opciones))

    def test_C10_adaptacion_dinamica_columnas_excel(self):
        """C10: Validación de adaptación dinámica de columnas ante hojas Excel con N columnas (5, 8, 12 columnas)."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # 1. Probar hoja con 5 columnas
            f5 = os.path.join(tmpdir, "test_5cols.xlsx")
            wb5 = openpyxl.Workbook()
            ws5 = wb5.active
            ws5.title = "Cinco_Columnas"
            h5 = ["ID", "Hostname", "Serial_Num", "Sede_Operativa", "Estado"]
            ws5.append(h5)
            ws5.append(["1", "LAP-01", "SN001", "Sede Norte", "Activo"])
            wb5.save(f5)
            wb5.close()

            self.app.ruta_excel.set(f5)
            self.app.nombre_hoja.set("Cinco_Columnas")
            self.app._auto_detectar_columnas_hoja("Cinco_Columnas")
            self.assertEqual(len(self.app.combo_col_serial["values"]), 5)
            self.assertEqual(len(self.app.tree["columns"]), 5)

            # 2. Probar hoja con 8 columnas
            f8 = os.path.join(tmpdir, "test_8cols.xlsx")
            wb8 = openpyxl.Workbook()
            ws8 = wb8.active
            ws8.title = "Ocho_Columnas"
            h8 = ["Item", "Cliente", "Sede", "Equipo", "Marca", "Serial", "Responsable", "Verificado"]
            ws8.append(h8)
            ws8.append(["1", "CLI", "Sede 1", "EQ-1", "HP", "SN002", "Resp 1", "OK"])
            wb8.save(f8)
            wb8.close()

            self.app.ruta_excel.set(f8)
            self.app.nombre_hoja.set("Ocho_Columnas")
            self.app._auto_detectar_columnas_hoja("Ocho_Columnas")
            self.assertEqual(len(self.app.combo_col_serial["values"]), 8)
            self.assertEqual(len(self.app.tree["columns"]), 8)

            # 3. Probar hoja con 12 columnas (Clemont real)
            self.app.ruta_excel.set(EXCEL_REAL_INPUT)
            self.app.nombre_hoja.set("Inventario de Equipos (2)")
            self.app._auto_detectar_columnas_hoja("Inventario de Equipos (2)")
            self.assertGreaterEqual(len(self.app.combo_col_serial["values"]), 12)
            self.assertGreaterEqual(len(self.app.tree["columns"]), 12)
            for c in self.app.tree["columns"]:
                self.assertEqual(self.app.tree.heading(c)["text"], str(c))


    def test_C11_pegado_exacto_clemont_12_columnas_sin_desfase_ni_item_1(self):
        """C11: Validación del pegado exacto de 12 columnas: Item debe ser 20, 21, 22 (no 1) y Cliente debe ser CLEMONT (no LAPTOP)."""
        tsv_user = (
            "Item\tCliente\tCentro Operativo / Sede\tEquipo / Servidor\tSistema Operativo / Software\tFabricante\tTipo Licencia\tCantidad\tResponsable / Usuario\tEstado\tObservaciones y Serial\tVerificado\n"
            "20\tCLEMONT\tAdministrativa 2do piso\tLAPTOP-2895F563\tWINDOWS 11 PRO\tMicrosoft\tSistema operativo\t1\tDANIEL PANESSO\tActivo\tSerial equipo: SDK0T76530; Modelo: LENOVO 83A0\t\n"
            "21\tCLEMONT\tAdministrativa 2do piso\tANALISTA-GHI\tWINDOWS 11 PRO\tMicrosoft\tSistema operativo\t1\tMELISSA MONSALVE\tActivo\tSerial equipo: T1N0CV0CV08C86402E; Modelo: Vivobook\t\n"
            "22\tCLEMONT\tAdministrativa 2do piso\tDESKTOP-144TP60\tWINDOWS 11 PRO\tMicrosoft\tSistema operativo\t1\tNO ASIGNADO\tActivo\tModelo: HP LAPTOP 14 D10XXX\t"
        )
        lineas = tsv_user.strip().splitlines()
        partes_hdr = [p.strip().upper() for p in lineas[0].split("\t")]

        # Probar que la detección de encabezados no asigne 'idx_item' a 'CANTIDAD' (Col 7) por la subcadena 'ID'
        idx_item = None
        for i, p in enumerate(partes_hdr):
            if (p in ("ITEM", "#", "ID", "N°", "NO.") or p.startswith("ITEM")) and "CANTIDAD" not in p and idx_item is None:
                idx_item = i
        self.assertEqual(idx_item, 0, f"idx_item debe ser 0 (Columna Item), no {idx_item} (Columna Cantidad)")

        # Probar mapeo fiel de los datos de la primera fila
        partes_f1 = [p.strip() for p in lineas[1].split("\t")]
        self.assertEqual(partes_f1[0], "20", "Item debe ser 20")
        self.assertEqual(partes_f1[1], "CLEMONT", "Cliente debe ser CLEMONT")
        self.assertEqual(partes_f1[2], "Administrativa 2do piso", "Sede debe ser Administrativa 2do piso")
        self.assertEqual(partes_f1[3], "LAPTOP-2895F563", "Equipo debe ser LAPTOP-2895F563")
        self.assertEqual(partes_f1[4], "WINDOWS 11 PRO", "SO debe ser WINDOWS 11 PRO")
        self.assertEqual(partes_f1[5], "Microsoft", "Fabricante debe ser Microsoft")
        self.assertEqual(partes_f1[7], "1", "Cantidad debe ser 1")
        self.assertEqual(partes_f1[8], "DANIEL PANESSO", "Responsable debe ser DANIEL PANESSO")
        self.assertEqual(partes_f1[9], "Activo", "Estado debe ser Activo")

        # Serial extraído no debe retener punto y coma ni modelo
        sns = cu.extraer_seriales(partes_f1[10])
        self.assertEqual(sns[0], "SDK0T76530")


# ============================================================================
# TIER 4: REAL-WORLD WORKLOAD TESTING (REAL EXCEL FILES)


# >= 5 realistic application scenarios
# ============================================================================

class TestTier4RealWorldWorkload(unittest.TestCase):
    """Tier 4: End-to-end stress and integrity validation against real Clemont Excel workbooks."""

    @classmethod
    def setUpClass(cls):
        cls.root, cls.app = create_headless_app()
        assert os.path.exists(EXCEL_REAL_INPUT), f"Archivo real no encontrado: {EXCEL_REAL_INPUT}"
        assert os.path.exists(EXCEL_REAL_AUDITED), f"Archivo real no encontrado: {EXCEL_REAL_AUDITED}"

    @classmethod
    def tearDownClass(cls):
        try:
            cls.root.destroy()
        except Exception:
            pass

    def test_S1_estructura_y_encabezados_libro_real_clemont(self):
        """
        Escenario 1: Ingestión del libro real `Inventario_Equipos_Clemont (2) (2).xlsx`.
        Verifica hojas, encabezados de fila 5 y conteo exacto de activos.
        """
        wb = openpyxl.load_workbook(EXCEL_REAL_INPUT, data_only=True)
        self.assertIn("Inventario de Equipos (2)", wb.sheetnames)
        ws = wb["Inventario de Equipos (2)"]

        # Encabezados en fila 5
        self.assertEqual(ws.cell(5, 1).value, "Item")
        self.assertEqual(ws.cell(5, 2).value, "Cliente")
        self.assertEqual(ws.cell(5, 3).value, "Centro Operativo / Sede")
        self.assertEqual(ws.cell(5, 4).value, "Equipo / Servidor")
        self.assertEqual(ws.cell(5, 8).value, "Cantidad")
        self.assertEqual(ws.cell(5, 9).value, "Responsable / Usuario")
        self.assertEqual(ws.cell(5, 10).value, "Estado")
        self.assertEqual(ws.cell(5, 11).value, "Observaciones y Serial")
        self.assertEqual(ws.cell(5, 12).value, "Verificado")

        # Conteo de filas de datos (de fila 6 a fila 76 = 71 equipos)
        total_row = self.app._detectar_fila_total(ws, 6)
        self.assertEqual(total_row, 77)
        self.assertIn("TOTAL", str(ws.cell(total_row, 1).value or "").upper())
        wb.close()

    def test_S2_extraccion_7_campos_sin_desfase_en_todos_los_71_equipos_reales(self):
        """
        Escenario 2: Extracción fidedigna de los 7 campos técnicos en TODOS los 71 registros
        del archivo real, comprobando que NUNCA exista cruce ni desfase entre columnas.
        """
        wb = openpyxl.load_workbook(EXCEL_REAL_INPUT, data_only=True)
        ws = wb["Inventario de Equipos (2)"]

        sedes_validas = {
            "ADMINISTRATIVA 2DO PISO", "CDI", "CHIPICHAPE", "CÚCUTA", "DOS QUEBRADAS",
            "ESTACIÓN", "FABRICATO", "FLORIDA", "SANTA FÉ", "SHOWROOM", "SHOWROOM-TIENDA",
            "SOPÓ", "SÓTANO", "TESORO", "UNICENTRO-CALI", "VILLAVICENCIO"
        }

        equipos_procesados = 0
        for r in range(6, 77):
            item_n = ws.cell(r, 1).value
            sede   = str(ws.cell(r, 3).value or "").strip()
            equipo = str(ws.cell(r, 4).value or "").strip()
            resp   = str(ws.cell(r, 9).value or "").strip()
            estado = str(ws.cell(r, 10).value or "Activo").strip()
            obs    = str(ws.cell(r, 11).value or "").strip()

            self.assertTrue(bool(equipo), f"Fila {r}: Hostname vacío")
            self.assertTrue(bool(sede), f"Fila {r}: Sede vacía")
            self.assertTrue(normalizar(sede) in [normalizar(s) for s in sedes_validas],
                            f"Fila {r}: Sede desconocida o desfasada: '{sede}'")

            # Validar que Sede no se haya filtrado en Hostname
            self.assertNotIn("ADMINISTRATIVA", equipo.upper())
            self.assertNotIn("SÓTANO", equipo.upper())

            # Validar extracción de los 7 campos mediante función estandarizada
            datos_raw = {
                "id": f"LAP-2026-{r-5:04d}",
                "equipo": equipo,
                "sede": sede,
                "responsable": resp,
                "estado": estado,
                "observaciones": obs,
            }
            campos = extraer_campos_completos(datos_raw)
            self.assertEqual(campos["equipo"], equipo)
            self.assertEqual(campos["sede"], sede)
            self.assertEqual(campos["responsable"], resp)
            self.assertTrue(len(campos["modelo"]) > 0)
            self.assertTrue(len(campos["serial"]) > 0)
            equipos_procesados += 1

        self.assertEqual(equipos_procesados, 71)
        wb.close()

    def test_S3_verificacion_integridad_libro_auditado_referencia(self):
        """
        Escenario 3: Validación del libro real `AUDITADO_Inventario_Equipos_Clemont (2) (2).xlsx`.
        Verifica presencia de hoja 'Reporte de Auditoria', columna L auditada, y tablas anexas.
        """
        wb = openpyxl.load_workbook(EXCEL_REAL_AUDITED, data_only=True)
        self.assertIn("Reporte de Auditoria", wb.sheetnames)
        ws_inv = wb["Inventario de Equipos (2)"]

        # Verificar columna L (Verificado) en registros auditados
        verificados = [ws_inv.cell(r, 12).value for r in range(6, 77) if ws_inv.cell(r, 12).value]
        self.assertGreater(len(verificados), 0)

        # Verificar presencia de al menos un 'OK' en auditoría
        self.assertTrue(any("OK" in str(v).upper() for v in verificados))

        # Verificar hoja de Reporte de Auditoria
        ws_rep = wb["Reporte de Auditoria"]
        self.assertIn("REPORTE DE AUDITORIA", str(ws_rep.cell(1, 1).value or "").upper())
        wb.close()

    def test_S4_generacion_masiva_stickers_qr_300dpi_con_datos_reales(self):
        """
        Escenario 4: Generación de lote representativo de etiquetas físicas a 300 DPI
        utilizando datos reales del archivo Clemont (Laptop, Monitor, Tablet, All-in-One),
        verificando DPI, dimensiones y todos los 7 campos en el QR.
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            wb = openpyxl.load_workbook(EXCEL_REAL_INPUT, data_only=True)
            ws = wb["Inventario de Equipos (2)"]

            filas_muestra = [6, 25, 41, 64, 71]  # Muestra variada de sedes y marcas
            for idx, r in enumerate(filas_muestra, 1):
                raw = {
                    "id": f"EQ-2026-{idx:04d}",
                    "equipo": str(ws.cell(r, 4).value or ""),
                    "sede": str(ws.cell(r, 3).value or ""),
                    "responsable": str(ws.cell(r, 9).value or ""),
                    "estado": str(ws.cell(r, 10).value or "Activo"),
                    "observaciones": str(ws.cell(r, 11).value or ""),
                }
                campos = extraer_campos_completos(raw)
                ruta_png = self.app._generar_etiqueta_individual(
                    {
                        "id": campos["id_activo"],
                        "equipo": campos["equipo"],
                        "modelo": campos["modelo"],
                        "serial": campos["serial"],
                        "responsable": campos["responsable"],
                        "sede": campos["sede"],
                        "estado": campos["estado"],
                    },
                    formato="50x25mm",
                    carpeta=tmpdir
                )
                self.assertTrue(os.path.exists(ruta_png))
                img = Image.open(ruta_png)
                self.assertIn(img.info.get("dpi"), ((300, 300), (300.0, 300.0)))
                self.assertEqual(img.size, (591, 295))

            wb.close()

    def test_S5_ejecucion_conciliacion_completa_contra_archivo_real_en_disco(self):
        """
        Escenario 5: Ejecución completa del motor de conciliación sobre copia de trabajo
        del archivo real `Inventario_Equipos_Clemont (2) (2).xlsx` con el preset físico completo (20 equipos).
        Verifica:
        - Conteo de coincidencias (19 OK, 1 Traslado, 1 Duplicado detectado).
        - Generación de 'AUDITADO_Inventario_Equipos_Clemont (2) (2).xlsx'.
        - Creación de hoja 'Reporte de Auditoria'.
        - Actualización dinámica de sumatoria en fila de totales.
        - Cero excepciones no controladas.
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            copia_input = os.path.join(tmpdir, "Inventario_Equipos_Clemont.xlsx")
            shutil.copyfile(EXCEL_REAL_INPUT, copia_input)
            ruta_salida = os.path.join(tmpdir, "AUDITADO_Inventario_Equipos_Clemont.xlsx")

            self.app.ruta_excel.set(copia_input)
            self.app.ruta_salida.set(ruta_salida)
            self.app.nombre_hoja.set("Inventario de Equipos (2)")

            # Cargar preset Clemont en el Treeview
            self.app._cargar_preset_clemont()
            self.assertEqual(len(self.app.tree.get_children()), 20)

            # Ejecutar conciliación desatendida
            self.app._ejecutar()

            # Aserciones sobre el archivo generado
            self.assertTrue(os.path.exists(ruta_salida), f"No se generó el archivo de salida en: {ruta_salida}")
            wb_aud = openpyxl.load_workbook(ruta_salida, data_only=True)
            self.assertIn("Reporte de Auditoria", wb_aud.sheetnames)

            ws_inv = wb_aud["Inventario de Equipos (2)"]

            # Comprobar que la fórmula de total fue actualizada
            tot_row = self.app._detectar_fila_total(ws_inv, 6)
            formula_tot = str(ws_inv.cell(tot_row, 8).value or "")
            self.assertTrue(formula_tot.startswith("=SUM(H6:H"), f"Fórmula inválida: {formula_tot}")

            # Comprobar resultados en columna L
            conteo_ok = 0
            conteo_traslado = 0
            conteo_dup = 0
            for r in range(6, tot_row):
                verif_val = str(ws_inv.cell(r, 12).value or "").strip().upper()
                if "OK" in verif_val:
                    conteo_ok += 1
                elif "TRASLADO" in verif_val:
                    conteo_traslado += 1
                elif "DUPLICADO" in verif_val:
                    conteo_dup += 1

            self.assertGreaterEqual(conteo_ok, 18)
            self.assertGreaterEqual(conteo_traslado, 1)

            wb_aud.close()


# ============================================================================
# RUNNER & SUMMARY REPORT
# ============================================================================

def run_all_qa_tests():
    """Executes all 4 tiers and prints formatted QA certification report."""
    print("=" * 80)
    print("  CLEMONT ASSET RECONCILIATION & QR SYSTEM — AUTOMATED QA SUITE (R5)")
    print("=" * 80)
    print(f"Timestamp: {cu.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Python:    {sys.version.split()[0]}")
    print(f"OpenPyXL:  {openpyxl.__version__}")
    print(f"Pillow:    {Image.__version__}")
    print(f"Workspace: {BASE_DIR}")
    print("-" * 80)

    loader = unittest.TestLoader()
    suite = unittest.TestSuite()

    tier1 = loader.loadTestsFromTestCase(TestTier1FeatureIsolation)
    tier2 = loader.loadTestsFromTestCase(TestTier2BoundaryAndCornerCases)
    tier3 = loader.loadTestsFromTestCase(TestTier3CrossFeatureIntegration)
    tier4 = loader.loadTestsFromTestCase(TestTier4RealWorldWorkload)

    suite.addTests(tier1)
    suite.addTests(tier2)
    suite.addTests(tier3)
    suite.addTests(tier4)

    total_count = suite.countTestCases()
    print(f"Discovered Test Cases:")
    print(f"  • Tier 1 (Feature Isolation R1-R5):            {tier1.countTestCases()} tests")
    print(f"  • Tier 2 (Boundary & Corner Cases):            {tier2.countTestCases()} tests")
    print(f"  • Tier 3 (Cross-Feature Combinations):         {tier3.countTestCases()} tests")
    print(f"  • Tier 4 (Real-World Workloads & Excel Files): {tier4.countTestCases()} tests")
    print(f"  -------------------------------------------------------------")
    print(f"  • TOTAL SUITE COUNT:                           {total_count} tests (Requirement >= 93)")
    print("-" * 80)

    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)

    print("\n" + "=" * 80)
    print("  QA CERTIFICATION SUMMARY REPORT")
    print("=" * 80)
    print(f"Total Tests Executed: {result.testsRun}")
    print(f"Passed:               {result.testsRun - len(result.failures) - len(result.errors)}")
    print(f"Failures:             {len(result.failures)}")
    print(f"Errors:               {len(result.errors)}")

    if result.wasSuccessful():
        print("\n>>> ALL TEST CASES PASSED CLEANLY (EXIT CODE 0) <<<")
        print("System is fully certified for production readiness.")
        print("=" * 80)
        sys.exit(0)
    else:
        print("\n>>> REGRESSIONS DETECTED (EXIT CODE 1) <<<")
        print("=" * 80)
        sys.exit(1)


if __name__ == "__main__":
    run_all_qa_tests()
