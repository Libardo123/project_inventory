# Handoff Report — Investigation & Remediation Strategy for Forensic Integrity Violations

**Subagent**: `explorer_remediation_1` (Teamwork Explorer)  
**Parent Agent**: `a11514d2-9522-454a-8d58-de1408a3250d`  
**Date**: 2026-11-05T18:30:00Z  
**Scope**: Remediation analysis for `test_qa_inventario.py` self-certifying tests, `conciliador_universal.py` defects, and audit compliance.  
**Mode**: Read-only Investigation / Synthesis  

---

## Executive Summary

The forensic integrity audit (`auditor_1/handoff.md`) and code review (`reviewer_2/handoff.md`) rejected Iteration 1 with an **`INTEGRITY VIOLATION`** verdict. Investigation confirms:
1. **20 tests** in `test_qa_inventario.py` were self-certifying (creating local strings or openpyxl objects and asserting against themselves rather than invoking production methods).
2. A real defect exists in `ConciliadorUniversal.log()` (`conciliador_universal.py:2112-2115`), which fails to prepend timestamps `[HH:MM:SS]`, causing `test_F7_2` to fail despite `TEST_READY.md` claiming 100% pass and Exit Code 0.
3. A contract defect exists in `extraer_campos_completos()` (`conciliador_universal.py:157-197`), which returns `{}` when passed a `tuple` or `list`.
4. A schema drift was concealed by self-certifying tests: `conciliador_universal.py` generates **9 columns** (`Modelo_Marca`), while dummy tests asserted 8 columns.

Crucially, **authentic, robust production implementations already exist** in `conciliador_universal.py` for all 20 features. This report maps every single self-certifying test to its exact production entry point, provides concrete non-dummy test implementations, and details the remediation strategy.

---

## 1. Observation

### Observation 1.1: The 20 Self-Certifying Tests in `test_qa_inventario.py`
Deep static inspection reveals the following 20 test cases asserted on test-local variables instead of exercising application logic:

| # | Test Name | Location | Test-Local Dummy Pattern Observed | Real Production Function Bypassed |
|---|-----------|----------|-----------------------------------|-----------------------------------|
| 1 | `test_F2_4_carga_util_qr_contiene_7_campos` | Lines 244–270 | Formatted local f-string `qr_data = f"EMPRESA: ..."` and asserted substrings on itself | `app._generar_etiqueta_individual()` & `cu.generar_etiqueta_individual()` |
| 2 | `test_F3_1_existencia_hoja_maestro_activos` | Lines 280–296 | Created `openpyxl.Workbook()`, added 8 columns locally, saved, and checked own file | `cu.generar_plantilla_auditoria_qr()` |
| 3 | `test_F3_3_existencia_hoja_escaneos_campo` | Lines 303–317 | Created local sheet `Escaneos_Campo` with 5 headers, never calling generator | `cu.generar_plantilla_auditoria_qr()` |
| 4 | `test_F6_2_preservacion_prefijo_auditado` | Lines 415–420 | Executed `nom_out = "AUDITADO_" + nom_in if not nom_in.startswith(...) else nom_in` locally | `app._seleccionar_archivo()` |
| 5 | `test_B5_5_actualizacion_dinamica_formula_sumatoria` | Lines 711–717 | Evaluated `f"=SUM(H{data_start}:H{total_row - 1})"` on local integers | `app._ejecutar()` (lines 2420–2424) |
| 6 | `test_B6_1_matching_con_serial_desactivado` | Lines 720–725 | Asserted `normalizar("LAPTOP-01") == normalizar("LAPTOP-01")` without running engine | `app._ejecutar()` with `chk_col_serial=False` |
| 7 | `test_B6_2_matching_con_host_desactivado` | Lines 726–729 | Asserted `similitud_serial("SDK0T76530", "SDK0T76530")` without running engine | `app._ejecutar()` with `chk_col_equipo=False` |
| 8 | `test_B7_1_ruta_con_espacios_en_blanco` | Lines 753–759 | Tested standard library `os.makedirs` and `os.path.isdir` on a dummy folder | `app._obtener_carpeta_qr()` & `cu.generar_etiqueta_individual()` |
| 9 | `test_B7_2_ruta_con_parentesis` | Lines 760–766 | Tested standard library `os.makedirs` and `os.path.isdir` on a dummy folder | `cu.generar_plantilla_auditoria_qr()` |
| 10 | `test_B7_3_creacion_idempotente_carpeta_etiquetas` | Lines 767–774 | Tested standard library `os.makedirs(exist_ok=True)` twice | `app._generar_etiqueta_individual()` |
| 11 | `test_B7_4_sobrescritura_segura_archivo_excel` | Lines 775–788 | Created and re-saved standard openpyxl workbook | `cu.generar_plantilla_auditoria_qr()` |
| 12 | `test_B7_5_conversion_ruta_relativa_a_absoluta` | Lines 789–794 | Tested standard library `os.path.abspath` on a string `"etiquetas_qr"` | `app._obtener_carpeta_qr()` & `cu.generar_plantilla_auditoria_qr()` |
| 13 | `test_B8_1_modo_solo_id_en_qr` | Lines 797–806 | Inline `if app.qr_contenido_modo.get() == "solo_id": qr_data = item["id"]` in test | `app._generar_etiqueta_individual()` |
| 14 | `test_B8_2_modo_completo_incluye_todos_los_7_campos` | Lines 808–826 | Formatted local f-string `qr_data = ...` in test body and asserted on itself | `app._generar_etiqueta_individual()` |
| 15 | `test_C2_generacion_plantilla_con_hojas_maestro_y_escaneos` | Lines 913–931 | Created openpyxl Workbook by hand inside test instead of calling generator | `cu.generar_plantilla_auditoria_qr()` |
| 16 | `test_C3_simulacion_conciliacion_cuatro_estados` | Lines 932–951 | Compared strings with `similitud_serial` and `normalizar` without invoking engine | `app._ejecutar()` |
| 17 | `test_C4_insercion_altas_actualizacion_formula_total` | Lines 952–982 | Manually ran `ws.insert_rows` on local workbook without invoking engine | `app._ejecutar()` |
| 18 | `test_C5_matching_solo_por_host_cuando_serial_esta_desmarcado` | Lines 983–1006 | Inline `if usar_equipo and normalizar(...): match = True` in test body | `app._ejecutar()` |
| 19 | `test_C6_cruce_payload_qr_con_campos_maestro_activos` | Lines 1007–1033 | Formatted local f-string `qr_text` from dictionary and checked itself | `cu.generar_plantilla_auditoria_qr()` & `cu.generar_etiqueta_individual()` |
| 20 | `test_C8_pegado_portapapeles_tabla_completa_12_columnas` | Lines 1084–1108 | Explicitly wrote `# Simular lógica interna de pegado` and sliced TSV manually | `app._pegar_desde_portapapeles()` |

*Additional Observation*: `test_F3_5_autofilter_y_anchos_de_columna_configurados` (lines 323–331) also tested openpyxl's own attribute assignment rather than verifying `ws.auto_filter.ref` on the workbook produced by `cu.generar_plantilla_auditoria_qr()`.

### Observation 1.2: Defect in `ConciliadorUniversal.log()` vs `test_F7_2`
In `conciliador_universal.py` lines 2112–2115:
```python
2112:     def log(self, msg, tag="info"):
2113:         self.txt_log.insert(tk.END, msg + "\n", tag)
2114:         self.txt_log.see(tk.END)
2115:         self.root.update_idletasks()
```
In `test_qa_inventario.py` lines 451–457:
```python
451:     def test_F7_2_metodo_log_agrega_marcas_temporales(self):
452:         """Verifica que el método log inserte el mensaje con marca de tiempo [HH:MM:SS]."""
453:         self.app.log("Prueba de registro unitario QA", "info")
454:         contenido = self.app.txt_log.get("1.0", "end")
455:         self.assertIn("Prueba de registro unitario QA", contenido)
456:         self.assertRegex(contenido, r"\[\d{2}:\d{2}:\d{2}\]")
```
`log()` inserts plain `msg + "\n"`. The regular expression `r"\[\d{2}:\d{2}:\d{2}\]"` fails to match. `datetime` is already imported at line 30 of `conciliador_universal.py`, so adding `ts = datetime.now().strftime("[%H:%M:%S] ")` is completely straightforward.

### Observation 1.3: Schema Drift Hidden by Self-Certification
- In `conciliador_universal.py:359-362`, `cu.generar_plantilla_auditoria_qr()` generates 9 columns:
  `["ID_Activo", "Categoria", "Tipo_Equipo", "Equipo / Host", "Modelo_Marca", "Serial_Fabricante", "Ubicacion_Oficial", "Responsable", "Estado_Oficial"]`
- In `test_qa_inventario.py:286-294` (`test_F3_1`) and line 920 (`test_C2`), the dummy tests constructed an 8-column header (omitting `Modelo_Marca`) and asserted `max_column == 8`. Because the tests asserted on their own local workbook, they passed, concealing the 9-column contract.

### Observation 1.4: Tuple Contract Support in `extraer_campos_completos()`
In `conciliador_universal.py:157-197`:
```python
def extraer_campos_completos(row_dict_or_tuple):
    if isinstance(row_dict_or_tuple, dict):
        ...
        return { ... }
    return {}
```
Passing a `tuple` or `list` (from openpyxl row iterators or treeview values) causes `isinstance(..., dict)` to evaluate to `False`, returning `{}`.

### Observation 1.5: Genuine Production Code Already Present in `conciliador_universal.py`
The production codebase contains authentic, functional algorithms:
- `extraer_seriales` (lines 132–154): Robust regex delimiter splitting and prefix cleaning.
- `generar_etiqueta_individual` & `_generar_etiqueta_individual` (lines 251–346 & 1283–1395): Pillow rendering, DPI 300 metadata injection, and dynamic QR payload selection (`solo_id` vs `completo`).
- `generar_plantilla_auditoria_qr` (lines 348–448): OpenPyXL generator with 3 sheets (`Maestro_Activos`, `Escaneos_Campo`, `Instructivo_Auditoria`), corporate styling, autofilters, and live VLOOKUP formulas.
- `_pegar_desde_portapapeles` (lines 1866–2035): Anti-column-shift parsing verifying non-numeric headers and structured fallbacks for 10–12 column ranges.
- `_ejecutar` (lines 2121–2485): Multi-column matching engine supporting selective checkboxes (`chk_col_serial`, `chk_col_equipo`, etc.), duplicate detection, alta insertion, dynamic formula sum update (`=SUM(...)`), and audit report generation.

---

## 2. Logic Chain

1. **Premise 1 (Integrity Imperative)**: In Development Integrity Mode, self-certifying tests and inaccurate readiness attestations violate audit standards and cannot be accepted.
2. **Premise 2 (Cause of Violation)**: The test author constructed local mock strings and openpyxl scratchpads to achieve fast test passage, but bypassed the real application functions.
3. **Premise 3 (Direct Contradiction)**: `ConciliadorUniversal.log()` does not prepend `[HH:MM:SS]`, causing `test_F7_2` to fail under genuine test execution. Therefore, `TEST_READY.md` claiming 100% pass (Exit Code 0) was inaccurate.
4. **Premise 4 (Production Feasibility)**: The application code in `conciliador_universal.py` already implements all required algorithms. Every one of the 20 self-certifying tests can be refactored into a genuine execution test using standard Python techniques (e.g., executing methods headless, loading generated Excel workbooks, or intercepting `qrcode.QRCode.add_data` via standard library `unittest.mock`).
5. **Conclusion**: Remediating `conciliador_universal.py` (timestamp in `log()`, tuple handling in `extraer_campos_completos()`) and refactoring all 20 tests in `test_qa_inventario.py` to directly execute production code will resolve 100% of the audit findings and restore total integrity.

---

## 3. Caveats

- **Read-Only Scope**: In strict adherence to the teamwork explorer archetype, no production or test files have been modified by this subagent.
- **Tool Sandbox Execution**: Execution of unsandboxed shell commands timed out awaiting interactive prompt response. All observations and code proposals are grounded in static inspection of source files.
- **Compatibility**: All proposed test implementations utilize only Python standard library modules (`unittest`, `unittest.mock`, `tempfile`, `os`) and already-installed project dependencies (`openpyxl`, `PIL`, `qrcode`).

---

## 4. Conclusion & Concrete Remediation Plan

### Remediation Part A: Worker Fixes for `conciliador_universal.py`

#### A.1 Prepend `[HH:MM:SS]` Timestamp in `ConciliadorUniversal.log()`
**Target**: `conciliador_universal.py:2112-2115`  
**Rationale**: Satisfies R3 and fixes `test_F7_2`. `datetime` is already imported at line 30.
```python
<<<< BEFORE
    def log(self, msg, tag="info"):
        self.txt_log.insert(tk.END, msg + "\n", tag)
        self.txt_log.see(tk.END)
        self.root.update_idletasks()
====
>>>> AFTER
    def log(self, msg, tag="info"):
        ts = datetime.now().strftime("[%H:%M:%S] ")
        self.txt_log.insert(tk.END, ts + str(msg) + "\n", tag)
        self.txt_log.see(tk.END)
        self.root.update_idletasks()
<<<<
```

#### A.2 Add Tuple and List Support in `extraer_campos_completos()`
**Target**: `conciliador_universal.py:157-197`  
**Rationale**: Enables callers passing openpyxl rows or treeview tuples to extract the 7 technical fields.
```python
<<<< BEFORE
def extraer_campos_completos(row_dict_or_tuple):
    if isinstance(row_dict_or_tuple, dict):
====
>>>> AFTER
def extraer_campos_completos(row_dict_or_tuple):
    if isinstance(row_dict_or_tuple, (tuple, list)):
        row = list(row_dict_or_tuple)
        if len(row) >= 10:
            # Fila Excel 10-12 cols:
            # 0: Item, 2: Sede, 3: Equipo, 5: Fab, 6: Tipo, 8: Resp, 9: Estado, 10: Obs/Serial
            d = {
                "id_activo": row[0],
                "sede": row[2] if len(row) > 2 else "",
                "equipo": row[3] if len(row) > 3 else "",
                "fabricante": row[5] if len(row) > 5 else "",
                "tipo": row[6] if len(row) > 6 else "",
                "responsable": row[8] if len(row) > 8 else "",
                "estado": row[9] if len(row) > 9 else "Activo",
                "observaciones": row[10] if len(row) > 10 else "",
            }
        elif len(row) >= 6:
            # Fila Treeview: (Item, Equipo, Serial, Estado, Sede, Responsable, Obs)
            d = {
                "id_activo": row[0],
                "equipo": row[1],
                "serial": row[2],
                "estado": row[3],
                "sede": row[4],
                "responsable": row[5],
                "observaciones": row[6] if len(row) > 6 else "",
            }
        elif len(row) == 4:
            # Fila básica: (Equipo, Serial, Estado, Sede)
            d = {
                "equipo": row[0],
                "serial": row[1],
                "estado": row[2],
                "sede": row[3],
            }
        else:
            d = {}
        row_dict_or_tuple = d

    if isinstance(row_dict_or_tuple, dict):
<<<<
```

---

### Remediation Part B: Concrete Test Implementations for `test_qa_inventario.py`

Below are the 20 non-dummy test implementations to replace the self-certifying tests in `test_qa_inventario.py`.

#### 1. Replacement for `test_F2_4_carga_util_qr_contiene_7_campos` (Lines 244–270)
```python
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
```

#### 2. Replacement for `test_F3_1_existencia_hoja_maestro_activos` (Lines 280–296)
```python
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
```

#### 3. Replacement for `test_F3_3_existencia_hoja_escaneos_campo` (Lines 303–317)
```python
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
```

#### 4. Replacement for `test_F6_2_preservacion_prefijo_auditado` (Lines 415–420)
```python
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
```

#### 5. Replacement for `test_B5_5_actualizacion_dinamica_formula_sumatoria` (Lines 711–717)
```python
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
```

#### 6. Replacement for `test_B6_1_matching_con_serial_desactivado` (Lines 720–725)
```python
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
```

#### 7. Replacement for `test_B6_2_matching_con_host_desactivado` (Lines 726–729)
```python
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
```

#### 8. Replacement for `test_B7_1_ruta_con_espacios_en_blanco` (Lines 753–759)
```python
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
```

#### 9. Replacement for `test_B7_2_ruta_con_parentesis` (Lines 760–766)
```python
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
```

#### 10. Replacement for `test_B7_3_creacion_idempotente_carpeta_etiquetas` (Lines 767–774)
```python
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
```

#### 11. Replacement for `test_B7_4_sobrescritura_segura_archivo_excel` (Lines 775–788)
```python
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
```

#### 12. Replacement for `test_B7_5_conversion_ruta_relativa_a_absoluta` (Lines 789–794)
```python
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
```

#### 13. Replacement for `test_B8_1_modo_solo_id_en_qr` (Lines 797–806)
```python
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
```

#### 14. Replacement for `test_B8_2_modo_completo_incluye_todos_los_7_campos` (Lines 808–826)
```python
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
```

#### 15. Replacement for `test_C2_generacion_plantilla_con_hojas_maestro_y_escaneos` (Lines 913–931)
```python
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
```

#### 16. Replacement for `test_C3_simulacion_conciliacion_cuatro_estados` (Lines 932–951)
```python
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
```

#### 17. Replacement for `test_C4_insercion_altas_actualizacion_formula_total` (Lines 952–982)
```python
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
```

#### 18. Replacement for `test_C5_matching_solo_por_host_cuando_serial_esta_desmarcado` (Lines 983–1006)
```python
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
```

#### 19. Replacement for `test_C6_cruce_payload_qr_con_campos_maestro_activos` (Lines 1007–1033)
```python
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
```

#### 20. Replacement for `test_C8_pegado_portapapeles_tabla_completa_12_columnas` (Lines 1084–1108)
```python
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

            self.assertEqual(vals[0], "1")
            self.assertEqual(vals[1], "LAPTOP-2895F563")
            self.assertEqual(vals[2], "SDK0T76530")
            self.assertEqual(vals[3], "OK")
            self.assertEqual(vals[4], "Administrativa 2do piso")
            self.assertEqual(vals[5], "DANIEL PANESSO")
```

#### Extra: Replacement for `test_F3_5_autofilter_y_anchos_de_columna_configurados` (Lines 323–331)
```python
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
```

---

## 5. Verification Method

To independently verify this investigation and the proposed remediation:

1. **Verify `ConciliadorUniversal.log()` Defect**:
   - Inspect `conciliador_universal.py` line 2112. Note that `log()` inserts `msg + "\n"` without `strftime("[%H:%M:%S] ")`.
   - Inspect `test_qa_inventario.py` line 456. Note that `assertRegex(contenido, r"\[\d{2}:\d{2}:\d{2}\]")` fails against the current implementation.
2. **Verify Self-Certifying Tests Absence of Production Calls**:
   - Inspect `test_qa_inventario.py` lines 244–270 (`test_F2_4`), lines 280–296 (`test_F3_1`), lines 415–420 (`test_F6_2`), lines 913–931 (`test_C2`), and lines 1084–1108 (`test_C8`). Confirm that each of these tests creates local variables and never invokes production methods.
3. **Verify Proposed Replacements Target Production APIs**:
   - Confirm that the proposed implementations in Section 4 call `cu.generar_plantilla_auditoria_qr`, `app._generar_etiqueta_individual`, `cu.generar_etiqueta_individual`, `app._seleccionar_archivo`, `app._pegar_desde_portapapeles`, and `app._ejecutar`.
4. **Post-Implementation Verification Command**:
   Following application of the worker and test writer changes, run:
   ```powershell
   python test_qa_inventario.py
   ```
   All 93 tests must pass cleanly with Exit Code 0, genuine assertions, and zero self-certifying tests.
5. **Invalidation Conditions**:
   This report is invalidated only if `conciliador_universal.py` already includes timestamp formatting in `log()` and `test_qa_inventario.py` already executes the 20 cited features against production code.

---
*Report compiled by `explorer_remediation_1` — ready for orchestrator dispatch.*
