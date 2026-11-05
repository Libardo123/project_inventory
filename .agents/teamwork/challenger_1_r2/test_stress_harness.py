# -*- coding: utf-8 -*-
"""
test_stress_harness.py — Adversarial Stress Harness (Round 2)
Agent: challenger_1_r2 (Empirical Challenger / Stress Challenger)

Scope of Adversarial Testing:
1. Multi-column clipboard paste (10-12 columns) using real Clemont data:
   `20\tCLEMONT\tAdministrativa 2do piso\tLAPTOP-2895F563...`
   - Assert Item # is 20 (NOT 1).
   - Assert columns align 1:1 (Host, Serial, Estado, Sede, Resp, Obs).
   - Also test with and without headers, mixed rows, and 10 vs 11 vs 12 cols.
2. Delimiter handling:
   - Semicolons, colons, commas, slashes, multiple spaces, newlines.
   - Multiple serials in observations (splitting cleanly without punctuation).
   - Observations with only models vs observations with models + serials.
   - Whitespace padding, trailing delimiters, tabs.
   - Missing values (None, empty string, all-whitespace, numeric values).
3. 300 DPI sticker image sizing and Pillow metadata:
   - Formats: 50x25mm (591x295), 100x50mm (1200x600), 67x25mm (788x300).
   - PIL Image open verification: info.get('dpi') == (300, 300).
   - High resolution rendering and physical mm/inch ratio check.
4. Dynamic column adaptation for 5, 8, 12, 15 column sheets:
   - 5-column custom asset sheet.
   - 8-column branch audit sheet.
   - 12-column Clemont master sheet.
   - 15-column enterprise datacenter inventory sheet.
   - Verifying Treeview columns reconfigure dynamically, mapping dropdowns populate,
     and headless execution completes cleanly.
"""

import os
import sys
import tempfile
import unittest
import openpyxl
from PIL import Image

# Ensure project root is in sys.path
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

import conciliador_universal as cu
from conciliador_universal import (
    normalizar,
    normalizar_tolerante,
    similitud_serial,
    extraer_seriales,
    extraer_campos_completos,
    generar_etiqueta_individual,
    generar_plantilla_auditoria_qr,
    ConciliadorUniversal,
)

def create_headless_app():
    """Initializes ConciliadorUniversal with hidden Tkinter window and mocked dialogs."""
    import tkinter as tk
    from tkinter import filedialog, messagebox

    messagebox.showinfo = lambda *args, **kwargs: None
    messagebox.showwarning = lambda *args, **kwargs: None
    messagebox.showerror = lambda *args, **kwargs: None
    messagebox.askyesno = lambda *args, **kwargs: True
    messagebox.askokcancel = lambda *args, **kwargs: True

    filedialog.askopenfilename = lambda *args, **kwargs: ""
    filedialog.askdirectory = lambda *args, **kwargs: ""
    filedialog.asksaveasfilename = lambda *args, **kwargs: ""

    cu.messagebox.showinfo = messagebox.showinfo
    cu.messagebox.showwarning = messagebox.showwarning
    cu.messagebox.showerror = messagebox.showerror
    cu.filedialog.askopenfilename = filedialog.askopenfilename
    cu.filedialog.askdirectory = filedialog.askdirectory
    cu.filedialog.asksaveasfilename = filedialog.asksaveasfilename

    if hasattr(os, "startfile"):
        os.startfile = lambda *args, **kwargs: None

    root = tk.Tk()
    root.withdraw()
    app = ConciliadorUniversal(root)
    return root, app


class TestAdversarialScope1ClipboardPaste(unittest.TestCase):
    """Scope 1: Multi-column clipboard paste with real Clemont data."""

    def setUp(self):
        self.root, self.app = create_headless_app()

    def tearDown(self):
        try:
            self.root.destroy()
        except Exception:
            pass

    def test_01_paste_real_clemont_row_item_20_assert_not_1(self):
        """Pasting row 20 from Clemont data must keep Item # as '20', NOT '1'."""
        # Exact real Clemont data for item 20
        raw_clipboard = (
            "20\tCLEMONT\tAdministrativa 2do piso\tLAPTOP-2895F563\tWindows 11 Pro\t"
            "OEM\tLenovo\tLenovo 83A0\tDANIEL PANESSO\tActivo\t"
            "Serial equipo: SDK0T76530; Modelo: 83A0\tOK"
        )
        self.app.root.clipboard_get = lambda: raw_clipboard
        self.app._pegar_desde_portapapeles()

        children = self.app.tree.get_children()
        self.assertEqual(len(children), 1, "Should have inserted exactly 1 row")
        item_vals = self.app.tree.item(children[0], "values")

        # Treeview columns: ("#", "Equipo / Host", "Serial Real", "Estado", "Sede", "Responsable", "Observaciones")
        item_num = item_vals[0]
        equipo = item_vals[1]
        serial = item_vals[2]
        estado = item_vals[3]
        sede = item_vals[4]
        resp = item_vals[5]
        obs = item_vals[6]

        # CORE ASSERTION 1: Item # is 20 (NOT 1)
        self.assertEqual(item_num, "20", f"CRITICAL BUG: Expected Item # to be '20', got '{item_num}'")
        self.assertNotEqual(item_num, "1", "CRITICAL BUG: Item # defaulted to 1 instead of preserving 20")

        # CORE ASSERTION 2: Columns align 1:1 without shift
        self.assertEqual(equipo, "LAPTOP-2895F563", f"Host shifted: {equipo}")
        self.assertEqual(serial, "SDK0T76530", f"Serial extraction shifted: {serial}")
        self.assertEqual(estado, "OK", f"Estado shifted: {estado}")
        self.assertEqual(sede, "Administrativa 2do piso", f"Sede shifted: {sede}")
        self.assertEqual(resp, "DANIEL PANESSO", f"Responsable shifted: {resp}")
        self.assertIn("SDK0T76530", obs, f"Obs shifted: {obs}")

    def test_02_paste_multi_row_with_headers_preserves_item_numbers(self):
        """Pasting multiple rows including header must parse header and keep item numbers."""
        raw_clipboard = (
            "ITEM\tCLIENTE\tSEDE\tEQUIPO\tSO\tLICENCIA\tFABRICANTE\tMODELO\tRESPONSABLE\tESTADO\tOBSERVACIONES Y SERIAL\tAUDITORIA\n"
            "45\tCLEMONT\tPlanta 1\tPC-CONTROL-01\tWin 10\tOEM\tDell\tOptiplex 7050\tJUAN PEREZ\tActivo\tSerial equipo: 9X8Y7Z1; Modelo: Optiplex\tOK\n"
            "60\tCLEMONT\tPlanta 2\tPC-CONTROL-02\tWin 10\tOEM\tDell\tOptiplex 7050\tANA GOMEZ\tActivo\tSerial equipo: 1A2B3C4; Modelo: Optiplex\tOK"
        )
        self.app.root.clipboard_get = lambda: raw_clipboard
        self.app._pegar_desde_portapapeles()

        children = self.app.tree.get_children()
        self.assertEqual(len(children), 2, "Header should be skipped and 2 rows added")

        vals_1 = self.app.tree.item(children[0], "values")
        self.assertEqual(vals_1[0], "45", f"Item # expected 45, got {vals_1[0]}")
        self.assertEqual(vals_1[1], "PC-CONTROL-01")
        self.assertEqual(vals_1[2], "9X8Y7Z1")
        self.assertEqual(vals_1[4], "Planta 1")
        self.assertEqual(vals_1[5], "JUAN PEREZ")

        vals_2 = self.app.tree.item(children[1], "values")
        self.assertEqual(vals_2[0], "60", f"Item # expected 60, got {vals_2[1]}")
        self.assertEqual(vals_2[1], "PC-CONTROL-02")
        self.assertEqual(vals_2[2], "1A2B3C4")

    def test_03_paste_10_columns_no_header_item_preserved(self):
        """Pasting 10-column table without header preserves Item # and 1:1 columns."""
        raw_clipboard = (
            "33\tCLEMONT\tSede Principal\tSRV-CLEMONT-01\tWindows Server\tVolume\tHP\tDL380\tMARIO LOPEZ\tActivo"
        )
        self.app.root.clipboard_get = lambda: raw_clipboard
        self.app._pegar_desde_portapapeles()

        children = self.app.tree.get_children()
        self.assertEqual(len(children), 1)
        vals = self.app.tree.item(children[0], "values")
        self.assertEqual(vals[0], "33", f"Expected 33, got {vals[0]}")
        self.assertEqual(vals[1], "SRV-CLEMONT-01")
        self.assertEqual(vals[4], "Sede Principal")
        self.assertEqual(vals[5], "MARIO LOPEZ")


class TestAdversarialScope2DelimiterHandling(unittest.TestCase):
    """Scope 2: Delimiter handling, multiple serials, whitespace, and missing values."""

    def test_01_multiple_serials_semicolon_delimited(self):
        """Extract multiple serials separated by semicolons cleanly."""
        text = "Serial equipo: SDK0T76530; Serial: PF3A1B2C; S/N: 88776655"
        serials = extraer_seriales(text)
        self.assertEqual(serials, ["SDK0T76530", "PF3A1B2C", "88776655"])

    def test_02_multiple_serials_mixed_delimiters(self):
        """Extract serials split by semicolons, commas, slashes, newlines, multiple spaces."""
        text = "SN: SER111111, SN: SER222222 / SER333333\nSER444444   SER555555"
        serials = extraer_seriales(text)
        self.assertEqual(serials, ["SER111111", "SER222222", "SER333333", "SER444444", "SER555555"])

    def test_03_trailing_punctuation_and_semicolons(self):
        """Strip trailing semicolons, commas, periods, colons without corrupting serial."""
        text = "Serial equipo: 5CD2345678;;;,,,  "
        serials = extraer_seriales(text)
        self.assertEqual(serials, ["5CD2345678"])

    def test_04_observation_with_serial_and_model(self):
        """Ensure Model part does NOT swallow serial and is not extracted as serial."""
        text = "Serial equipo: SDK0T76530; Modelo: Lenovo 83A0"
        serials = extraer_seriales(text)
        self.assertEqual(serials, ["SDK0T76530"])

    def test_05_observation_with_only_model_no_serial(self):
        """When observation only has model description, serial must be empty or 'S/N NO REGISTRADO'."""
        text = "Modelo: Asus Vivobook Pro 15"
        serials = extraer_seriales(text)
        self.assertEqual(serials, [], "Should not extract model text as serial")

        # Test via extraer_campos_completos
        row_dict = {
            "id_activo": "LAP-01",
            "equipo": "ASUS-VIVO",
            "observaciones": "Modelo: Asus Vivobook Pro 15",
            "responsable": "CARLOS",
            "sede": "Norte",
            "estado": "Activo"
        }
        campos = extraer_campos_completos(row_dict)
        self.assertEqual(campos["serial"], "S/N NO REGISTRADO")
        self.assertIn("Asus Vivobook", campos["modelo"])

    def test_06_missing_and_corrupt_values(self):
        """Gracefully handle None, empty strings, whitespace, numbers, CLEMONT company name."""
        self.assertEqual(extraer_seriales(None), [])
        self.assertEqual(extraer_seriales(""), [])
        self.assertEqual(extraer_seriales("   \t\n   "), [])
        self.assertEqual(extraer_seriales("CLEMONT"), [])
        self.assertEqual(extraer_seriales("Serial: CLEMONT"), [])
        self.assertEqual(extraer_seriales(987654321), ["987654321"])
        self.assertEqual(extraer_seriales("N/A; - ; None"), [])


class TestAdversarialScope3DpiAndPillowMetadata(unittest.TestCase):
    """Scope 3: 300 DPI sticker image sizing and Pillow metadata (info.get('dpi') == (300, 300))."""

    def test_01_sticker_50x25mm_300dpi(self):
        """Format 50x25mm must produce 591x295 pixels and dpi=(300, 300) in Pillow metadata."""
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, "sticker_50x25.png")
            equipo = {
                "id_activo": "LAP-2026-0020",
                "equipo": "LAPTOP-2895F563",
                "modelo": "Lenovo 83A0",
                "serial": "SDK0T76530",
                "responsable": "DANIEL PANESSO",
                "sede": "Administrativa 2do piso",
                "estado": "Activo"
            }
            res_path = generar_etiqueta_individual(equipo, out_path, formato="50x25mm", dpi=300)
            self.assertTrue(os.path.exists(res_path))

            with Image.open(res_path) as img:
                # Pixel sizing check
                self.assertEqual(img.size, (591, 295), f"Expected (591, 295), got {img.size}")
                # 300 DPI metadata check
                dpi_meta = img.info.get("dpi")
                self.assertIsNotNone(dpi_meta, "Pillow image.info must contain 'dpi' metadata chunk")
                self.assertEqual(dpi_meta, (300, 300), f"Expected dpi (300, 300), got {dpi_meta}")

    def test_02_sticker_100x50mm_300dpi(self):
        """Format 100x50mm must produce 1200x600 pixels and dpi=(300, 300)."""
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, "sticker_100x50.png")
            equipo = {
                "id_activo": "SRV-2026-0001",
                "equipo": "SRV-PROD-01",
                "modelo": "PowerEdge R740",
                "serial": "8HG54F2",
                "responsable": "INFRAESTRUCTURA",
                "sede": "Datacenter Central",
                "estado": "Activo"
            }
            res_path = generar_etiqueta_individual(equipo, out_path, formato="100x50mm", dpi=300)
            with Image.open(res_path) as img:
                self.assertEqual(img.size, (1200, 600))
                self.assertEqual(img.info.get("dpi"), (300, 300))

    def test_03_sticker_67x25mm_300dpi(self):
        """Format 67x25mm must produce 788x300 pixels and dpi=(300, 300)."""
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, "sticker_67x25.png")
            equipo = {
                "id_activo": "MON-2026-0005",
                "equipo": "MONITOR-LG-27",
                "modelo": "LG 27UL500",
                "serial": "104NDTG2E123",
                "responsable": "FINANZAS",
                "sede": "Sede Norte",
                "estado": "Activo"
            }
            res_path = generar_etiqueta_individual(equipo, out_path, formato="67x25mm", dpi=300)
            with Image.open(res_path) as img:
                self.assertEqual(img.size, (788, 300))
                self.assertEqual(img.info.get("dpi"), (300, 300))


class TestAdversarialScope4DynamicColumns(unittest.TestCase):
    """Scope 4: Dynamic column adaptation for 5, 8, 12, 15 column sheets."""

    def setUp(self):
        self.root, self.app = create_headless_app()

    def tearDown(self):
        try:
            self.root.destroy()
        except Exception:
            pass

    def _create_test_workbook(self, tmpdir, filename, sheet_name, headers, rows):
        path = os.path.join(tmpdir, filename)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name
        ws.append(headers)
        for r in rows:
            ws.append(r)
        wb.save(path)
        wb.close()
        return path

    def test_01_dynamic_5_column_sheet(self):
        """Adapt GUI Treeview and mapping to a 5-column custom asset sheet."""
        with tempfile.TemporaryDirectory() as tmpdir:
            headers = ["Item", "Sede", "Dispositivo", "Serial", "Estado"]
            rows = [
                [1, "Sede Centro", "DESKTOP-001", "ABC111222", "OK"],
                [2, "Sede Norte", "DESKTOP-002", "XYZ333444", "OK"],
            ]
            path = self._create_test_workbook(tmpdir, "cinco_cols.xlsx", "Activos5", headers, rows)

            self.app.ruta_excel.set(path)
            self.app.nombre_hoja.set("Activos5")
            self.app._auto_detectar_columnas_hoja("Activos5")

            # Check Treeview reconfigured to 5 columns
            cols = list(self.app.tree["columns"])
            self.assertEqual(len(cols), 5, f"Expected 5 columns, got {len(cols)}: {cols}")
            self.assertEqual(cols, headers)

            # Check rows loaded
            children = self.app.tree.get_children()
            self.assertEqual(len(children), 2)
            self.assertEqual(self.app.tree.item(children[0], "values"), ("1", "Sede Centro", "DESKTOP-001", "ABC111222", "OK"))

            # Check detected dropdowns
            self.assertIn("Col D (4)", self.app.col_serial_var.get())
            self.assertIn("Col C (3)", self.app.col_equipo_var.get())

    def test_02_dynamic_8_column_sheet(self):
        """Adapt GUI Treeview and mapping to an 8-column branch audit sheet."""
        with tempfile.TemporaryDirectory() as tmpdir:
            headers = ["ID_Reg", "Sucursal", "Hostname", "Marca_Mod", "Serial_Num", "Usuario", "Estado_Op", "Observaciones"]
            rows = [
                [101, "Medellin", "MED-LAP-01", "Dell Latitude", "8899AABB", "JUAN VALENCIA", "Operativo", "Entregado 2025"],
                [102, "Cali", "CAL-LAP-02", "Lenovo ThinkPad", "CCDDEE11", "MARIA LOPEZ", "Operativo", "Entregado 2026"],
            ]
            path = self._create_test_workbook(tmpdir, "ocho_cols.xlsx", "Auditoria8", headers, rows)

            self.app.ruta_excel.set(path)
            self.app.nombre_hoja.set("Auditoria8")
            self.app._auto_detectar_columnas_hoja("Auditoria8")

            cols = list(self.app.tree["columns"])
            self.assertEqual(len(cols), 8, f"Expected 8 columns, got {len(cols)}")
            self.assertEqual(cols, headers)

            children = self.app.tree.get_children()
            self.assertEqual(len(children), 2)
            self.assertEqual(len(self.app.tree.item(children[0], "values")), 8)

    def test_03_dynamic_12_column_sheet_clemont(self):
        """Adapt GUI Treeview to standard Clemont 12-column layout."""
        with tempfile.TemporaryDirectory() as tmpdir:
            headers = [
                "Item", "Cliente", "Centro Operativo", "Equipo / Servidor",
                "Sistema Operativo", "Tipo de Licencia", "Fabricante", "Modelo",
                "Responsable", "Estado", "Observaciones y Serial", "Verificado"
            ]
            rows = [
                [1, "CLEMONT", "Administrativa", "LAP-01", "Win 11", "OEM", "Lenovo", "83A0", "PEDRO", "Activo", "Serial equipo: 12345678", "OK"]
            ]
            path = self._create_test_workbook(tmpdir, "doce_cols.xlsx", "Inventario", headers, rows)

            self.app.ruta_excel.set(path)
            self.app.nombre_hoja.set("Inventario")
            self.app._auto_detectar_columnas_hoja("Inventario")

            cols = list(self.app.tree["columns"])
            self.assertEqual(len(cols), 12)
            self.assertEqual(cols, headers)

    def test_04_dynamic_15_column_enterprise_sheet(self):
        """Adapt GUI Treeview to wide 15-column enterprise datacenter sheet."""
        with tempfile.TemporaryDirectory() as tmpdir:
            headers = [
                "ID", "Entidad", "Sede", "Piso", "Host", "Tipo", "Fabricante",
                "Modelo", "Serial", "MAC", "IP", "Usuario", "Area", "Estado", "Fecha"
            ]
            rows = [
                ["A1", "CORP", "Sede Central", "Piso 4", "SRV-DC01", "Servidor", "Dell",
                 "PowerEdge", "987654321", "00:1A:2B:3C:4D:5E", "10.0.0.1", "ADMIN", "TI", "Activo", "2026-01-01"]
            ]
            path = self._create_test_workbook(tmpdir, "quince_cols.xlsx", "Datacenter15", headers, rows)

            self.app.ruta_excel.set(path)
            self.app.nombre_hoja.set("Datacenter15")
            self.app._auto_detectar_columnas_hoja("Datacenter15")

            cols = list(self.app.tree["columns"])
            self.assertEqual(len(cols), 15, f"Expected 15 columns, got {len(cols)}")
            self.assertEqual(cols, headers)

            children = self.app.tree.get_children()
            self.assertEqual(len(children), 1)
            vals = self.app.tree.item(children[0], "values")
            self.assertEqual(len(vals), 15)
            self.assertEqual(vals[4], "SRV-DC01")
            self.assertEqual(vals[8], "987654321")


if __name__ == "__main__":
    suite = unittest.TestSuite()
    loader = unittest.TestLoader()
    suite.addTests(loader.loadTestsFromTestCase(TestAdversarialScope1ClipboardPaste))
    suite.addTests(loader.loadTestsFromTestCase(TestAdversarialScope2DelimiterHandling))
    suite.addTests(loader.loadTestsFromTestCase(TestAdversarialScope3DpiAndPillowMetadata))
    suite.addTests(loader.loadTestsFromTestCase(TestAdversarialScope4DynamicColumns))

    runner = unittest.TextTestRunner(verbosity=2)
    res = runner.run(suite)
    sys.exit(0 if res.wasSuccessful() else 1)
