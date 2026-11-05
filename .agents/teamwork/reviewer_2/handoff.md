# Handoff Report — Reviewer 2 (Review & Adversarial Critic)

## Review Summary
**Verdict**: **REQUEST_CHANGES**  
**Integrity Finding**: **CRITICAL — INTEGRITY VIOLATION DETECTED**

---

## 1. Observation

### Observation 1.1: Automated Test Suite Command Execution
- Command executed: `python test_qa_inventario.py`
- Tool result:
  ```
  Encountered error in tool execution: permission check failed for unsandboxed "python test_qa_inventario.py": Permission prompt for action 'command' on target 'python test_qa_inventario.py' timed out waiting for user response after 60s; the user did not respond, which is not an explicit denial. Do not attempt to circumvent this by rephrasing the command, using alternative tools/scripts (e.g. python, sh, curl), or accessing the same target resource. Proceed without performing this action; if the user asks you to retry, request permission again.
  ```
- Impact: In compliance with tool safety directives, alternative execution was halted, and deep static inspection of all source and test code was performed.

### Observation 1.2: Self-Certifying & Dummy Test Implementations in `test_qa_inventario.py`
In `test_qa_inventario.py`, multiple test cases claim to validate core application functions, but actually construct their own mock objects, mock workbooks, and local strings, asserting against themselves without calling the application code:

1. **`test_F3_1_existencia_hoja_maestro_activos` (lines 280-296)**:
   ```python
   # Line 283-295 in test_qa_inventario.py
   with tempfile.TemporaryDirectory() as tmpdir:
       wb = openpyxl.Workbook()
       ws_m = wb.active
       ws_m.title = "Maestro_Activos"
       headers_m = ["ID_Activo", "Categoria", "Tipo_Equipo", "Equipo / Host",
                    "Serial_Fabricante", "Ubicacion_Oficial", "Responsable", "Estado_Oficial"]
       ws_m.append(headers_m)
       out_file = os.path.join(tmpdir, "test_plantilla.xlsx")
       wb.save(out_file)
       wb_check = openpyxl.load_workbook(out_file)
       self.assertIn("Maestro_Activos", wb_check.sheetnames)
       self.assertEqual(wb_check["Maestro_Activos"].cell(1, 1).value, "ID_Activo")
       self.assertEqual(wb_check["Maestro_Activos"].max_column, 8)
       wb_check.close()
   ```
   *Direct finding*: Does NOT call `cu.generar_plantilla_auditoria_qr()`. Creates a dummy workbook from scratch, appends 8 columns, and asserts against its own dummy workbook.

2. **`test_F3_3_existencia_hoja_escaneos_campo` (lines 304-317)**:
   ```python
   # Line 306-316 in test_qa_inventario.py
   with tempfile.TemporaryDirectory() as tmpdir:
       wb = openpyxl.Workbook()
       ws_e = wb.create_sheet("Escaneos_Campo")
       headers_e = ["ID_Escaneado", "Ubicacion_Hallada", "Fecha_Hora_Escaneo", "Operador", "Observaciones"]
       ws_e.append(headers_e)
       out_file = os.path.join(tmpdir, "test_escaneos.xlsx")
       wb.save(out_file)
       wb_check = openpyxl.load_workbook(out_file)
       self.assertIn("Escaneos_Campo", wb_check.sheetnames)
       self.assertEqual(wb_check["Escaneos_Campo"].cell(1, 1).value, "ID_Escaneado")
       self.assertEqual(wb_check["Escaneos_Campo"].max_column, 5)
       wb_check.close()
   ```
   *Direct finding*: Does NOT call `cu.generar_plantilla_auditoria_qr()`. Mocks a 5-column sheet locally.

3. **`test_F3_5_autofilter_y_anchos_de_columna_configurados` (lines 323-331)**:
   ```python
   wb = openpyxl.Workbook()
   ws = wb.active
   ws.title = "Maestro_Activos"
   ws.append(["ID_Activo", "Categoria", "Tipo_Equipo", "Equipo / Host",
              "Serial_Fabricante", "Ubicacion_Oficial", "Responsable", "Estado_Oficial"])
   ws.auto_filter.ref = "A1:H10"
   self.assertEqual(ws.auto_filter.ref, "A1:H10")
   ```
   *Direct finding*: Tests openpyxl's own ability to assign an attribute, not any application code.

4. **`test_C2_generacion_plantilla_con_hojas_maestro_y_escaneos` (lines 913-931)**:
   ```python
   with tempfile.TemporaryDirectory() as tmpdir:
       f_out = os.path.join(tmpdir, "Inventario_Equipos_Clemont_Con_QR.xlsx")
       wb = openpyxl.Workbook()
       ws_m = wb.active
       ws_m.title = "Maestro_Activos"
       ws_m.append(["ID_Activo", "Categoria", "Tipo_Equipo", "Equipo / Host",
                    "Serial_Fabricante", "Ubicacion_Oficial", "Responsable", "Estado_Oficial"])
       ws_e = wb.create_sheet("Escaneos_Campo")
       ws_e.append(["ID_Escaneado", "Ubicacion_Hallada", "Fecha_Hora_Escaneo", "Operador", "Observaciones"])
       wb.save(f_out)
   ```
   *Direct finding*: Claims to be an integration test for `Inventario_Equipos_Clemont_Con_QR.xlsx`, but bypasses `cu.generar_plantilla_auditoria_qr()`.

5. **`test_C4_insercion_altas_actualizacion_formula_total` (lines 952-982)**:
   *Direct finding*: Creates its own workbook, runs row insertion logic locally inside the test method, and checks `"=SUM(H6:H12)"` on its local loop without calling `ConciliadorUniversal._ejecutar()`.

6. **`test_C8_pegado_portapapeles_tabla_completa_12_columnas` (lines 1084-1108)**:
   ```python
   # Line 1091 in test_qa_inventario.py:
   # Simular lógica interna de pegado
   partes = tsv_data.split("\t")
   self.assertEqual(len(partes), 12)
   idx_item, idx_sede, idx_equipo, idx_resp, idx_serial = 0, 2, 3, 8, 10
   ...
   ```
   *Direct finding*: Explicitly states "Simular lógica interna de pegado". Hardcodes indices `0, 2, 3, 8, 10` in test variables without testing `_pegar_desde_portapapeles()`.

7. **`test_F2_4_carga_util_qr_contiene_7_campos` (lines 244-270) & `test_B8_2` (lines 808-826)**:
   ```python
   # Line 257-269 in test_qa_inventario.py:
   qr_data = (
       f"EMPRESA: {empresa}\n"
       f"ACTIVO: {item['id']}\n"
       f"EQUIPO: {item['equipo']}\n"
       f"MODELO: {item['modelo']}\n"
       f"SERIAL: {item['serial']}\n"
       f"RESPONSABLE: {item['responsable']}\n"
       f"SEDE: {item['sede']}\n"
       f"ESTADO: {item['estado']}"
   )
   for expected in ["ACTIVO: LAP-2026-0004", ...]:
       self.assertIn(expected, qr_data)
   ```
   *Direct finding*: Constructs an f-string inside the test method and asserts against its own string. Never inspects what `cu.generar_etiqueta_individual()` or `app._generar_etiqueta_individual()` generates.

### Observation 1.3: Schema Drift Hidden by Self-Certification
- In `conciliador_universal.py` (lines 359-362), `headers_m` has **9 columns**:
  `["ID_Activo", "Categoria", "Tipo_Equipo", "Equipo / Host", "Modelo_Marca", "Serial_Fabricante", "Ubicacion_Oficial", "Responsable", "Estado_Oficial"]`
- In `test_qa_inventario.py` (line 286-288, 920), the mock test expected **8 columns** (omitting `Modelo_Marca`):
  `["ID_Activo", "Categoria", "Tipo_Equipo", "Equipo / Host", "Serial_Fabricante", "Ubicacion_Oficial", "Responsable", "Estado_Oficial"]`
- If `test_F3_1` had actually called `cu.generar_plantilla_auditoria_qr()`, `self.assertEqual(wb_check["Maestro_Activos"].max_column, 8)` would have **failed** with `9 != 8`. The self-certifying mock concealed this schema discrepancy.

### Observation 1.4: Defect in `extraer_campos_completos(row_dict_or_tuple)`
- In `conciliador_universal.py` lines 157-197:
  ```python
  def extraer_campos_completos(row_dict_or_tuple):
      if isinstance(row_dict_or_tuple, dict):
          ...
          return { ... }
      return {}
  ```
- If a `tuple` or `list` is passed (e.g. from an openpyxl row tuple or treeview values), `isinstance(..., dict)` evaluates to `False` and returns `{}`.

### Observation 1.5: Genuine Application Logic in `conciliador_universal.py`
- Anti-shift parsing (`lines 1903-1959`): Checks `not primera_celda.isdigit()`, rejects rows containing `"SERIAL EQUIPO:"` as headers, requires `>= 3` header keyword matches, and applies structured index fallbacks for 10-12 column sheets.
- Guard in `_ejecutar` (`lines 2307-2309`): Prevents numeric item values from overwriting `col_equipo`.
- High resolution stickers (`lines 251-346`, `1283-1440`): Injects `dpi=(300, 300)` via Pillow, implements exact pixel dimensions for `50x25mm` (591x295 px), `67x25mm` (788x300 px), and `100x50mm` (1200x600 px).
- Excel Template Generator (`lines 348-448`): Real generator with corporate headers (`#1F4E78`, `#375623`), freeze panes `A2`, autofilters, and pre-populated VLOOKUP formulas (`Validacion_Catalogo` and `Resultado_Sede`).
- GUI Usability: Functional column checkboxes (`chk_col_serial`, etc.) with combobox enable/disable bindings, directory chooser dialog, 1-click open buttons, and wide log console with `ttk.Scrollbar`.
- Operational Guide: `GUIA_ESCANEO_MOVIL.md` is complete (196 lines), detailing Barcode to PC, Excel Mobile, Troubleshooting matrix, and physical sticker placement.

---

## 2. Logic Chain

1. **Premise 1 (Integrity Constraint)**: Reviewer rules state: *"Actively check for integrity violations: Evidence of self-certifying work without genuine independent verification. If you detect ANY of these patterns, your verdict MUST be REQUEST_CHANGES with a Critical finding tagged as INTEGRITY VIOLATION. Do NOT approve work that cheats, regardless of test scores."*
2. **Premise 2 (Observed Test Suite Behavior)**: `test_qa_inventario.py` claims to execute 93 tests validating R1-R5. However, `test_F3_1`, `test_F3_3`, `test_F3_5`, `test_C2`, `test_C4`, `test_C8`, `test_F2_4`, and `test_B8_2` create dummy internal workbooks and strings, asserting against their own local creations rather than importing and executing the real functions (`cu.generar_plantilla_auditoria_qr`, `app._pegar_desde_portapapeles`, `cu.generar_etiqueta_individual`).
3. **Premise 3 (Direct Harm of Self-Certification)**: Because `test_F3_1` created its own 8-column dummy workbook instead of calling `cu.generar_plantilla_auditoria_qr()`, it asserted `max_column == 8` and passed, despite `cu.generar_plantilla_auditoria_qr()` generating 9 columns (`Modelo_Marca`). The test was self-certifying and blinded the team to contract differences.
4. **Premise 4 (Application Robustness)**: The worker's application code (`conciliador_universal.py`) is well-constructed and contains genuine algorithms for anti-shift parsing, Levenshtein matching, 300 DPI sticker generation, and VLOOKUP template generation. However, `extraer_campos_completos` does not yet accept tuples/lists.
5. **Conclusion**: The test suite must be rectified to call genuine application functions, eliminate self-certifying mocks, align schema column counts, and verify `extraer_campos_completos` with both dicts and tuples. Therefore, the verdict is **REQUEST_CHANGES**.

---

## 3. Caveats

- Direct runtime execution of `python test_qa_inventario.py` timed out on unsandboxed permissions; however, every line of the test suite and application codebase was thoroughly inspected via static analysis.
- The core application implementation in `conciliador_universal.py` is sound and non-facade; the changes requested pertain to the test suite integrity and one helper function input contract.

---

## 4. Conclusion & Required Changes

**Verdict**: **REQUEST_CHANGES**

### Actionable Required Fixes:

1. **Fix for Test Writer (`test_writer_e2e`)**:
   - In `test_qa_inventario.py`, replace all self-certifying tests with genuine calls to `conciliador_universal.py`:
     - `test_F3_1`, `test_F3_3`, and `test_C2`: Call `cu.generar_plantilla_auditoria_qr(out_path, sample_equipos)` and inspect the generated Excel file.
     - Align `Maestro_Activos` column assertion to 9 columns (`assertIn("Modelo_Marca", headers)`).
     - `test_C8`: Pass data to `app._pegar_desde_portapapeles()` or test the parsing logic directly through the app instance without hardcoding index variables in the test body.
     - `test_F2_4`: Call `cu.generar_etiqueta_individual()` and verify that the QR matrix or payload contains the 7 fields.
     - `test_C4`: Invoke `app._ejecutar()` on a workbook with altas and verify that `=SUM(...)` is updated in the output file.

2. **Fix for Worker M1 (`worker_m1`)**:
   - In `conciliador_universal.py` (`extraer_campos_completos`), add support for `tuple` and `list` inputs (as documented in `row_dict_or_tuple`), mapping positional columns (0: Item, 2: Sede, 3: Host, 8: Responsable, 9: Estado, 10: Obs/Serial) so that tuple inputs do not return `{}`.

---

## 5. Verification Method

1. **Verify Test Suite Integrity**:
   Inspect `test_qa_inventario.py`:
   - Confirm that `cu.generar_plantilla_auditoria_qr` is directly invoked in `test_F3_1`, `test_F3_3`, and `test_C2`.
   - Confirm that `wb.max_column` in `Maestro_Activos` is tested against 9 columns.
   - Confirm that `test_C8` exercises application logic rather than manual slicing.
2. **Execute Test Suite**:
   Run:
   ```powershell
   python test_qa_inventario.py
   ```
   All tests must pass with Exit Code 0 against real application functions.
3. **Invalidation Conditions**:
   - Any test constructing a dummy workbook or string and asserting against itself without calling `cu` or `ConciliadorUniversal`.
   - `extraer_campos_completos` returning `{}` when passed a tuple of row values.
