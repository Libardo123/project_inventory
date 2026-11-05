# Handoff Report — Code Remediation Strategy (explorer_remediation_2)

**Subagent**: `explorer_remediation_2` (Teamwork Explorer / Investigator)  
**Parent Agent**: `a11514d2-9522-454a-8d58-de1408a3250d`  
**Date**: 2026-11-05T18:28:00Z  
**Scope**: Code Remediation Investigation (`conciliador_universal.py`, `test_qa_inventario.py`, `GUIA_ESCANEO_MOVIL.md`)  
**Status**: Read-only Investigation Completed — Actionable Strategy Formulated  

---

## 1. Observation

### Observation 1.1: `ConciliadorUniversal.log()` Timestamp Omission Defect
- **File & Line**: `conciliador_universal.py:2112-2115`
- **Existing Code**:
  ```python
  2112:     def log(self, msg, tag="info"):
  2113:         self.txt_log.insert(tk.END, msg + "\n", tag)
  2114:         self.txt_log.see(tk.END)
  2115:         self.root.update_idletasks()
  ```
- **Companion Method**: `conciliador_universal.py:1015-1021` (`log_qr`):
  ```python
  1015:     def log_qr(self, msg, tag="info"):
  1016:         """Registra un mensaje en la bitacora de la pestana QR."""
  1017:         if hasattr(self, "txt_qr_log") and self.txt_qr_log:
  1018:             self.txt_qr_log.insert(tk.END, msg + "\n", tag)
  1019:             self.txt_qr_log.see(tk.END)
  1020:             self.root.update_idletasks()
  ```
- **Test Assertion (`test_qa_inventario.py:451-457`)**:
  ```python
  451:     def test_F7_2_metodo_log_agrega_marcas_temporales(self):
  452:         """Verifica que el método log inserte el mensaje con marca de tiempo [HH:MM:SS]."""
  453:         self.app.log("Prueba de registro unitario QA", "info")
  454:         contenido = self.app.txt_log.get("1.0", "end")
  455:         self.assertIn("Prueba de registro unitario QA", contenido)
  456:         self.assertRegex(contenido, r"\[\d{2}:\d{2}:\d{2}\]")
  ```
- **Observed Defect**: The production `log()` method does not prepend `[HH:MM:SS]`. As a result, `contenido` in `test_F7_2` is `"Prueba de registro unitario QA\n"`, which fails `assertRegex(..., r"\[\d{2}:\d{2}:\d{2}\]")`. `datetime` is already imported at line 30 (`from datetime import datetime`), but neither `log` nor `log_qr` formats or prepends the timestamp.

---

### Observation 1.2: Signature Mismatch & Dict-Only Handling in `extraer_campos_completos(row_dict_or_tuple)`
- **File & Line**: `conciliador_universal.py:157-197`
- **Existing Code**:
  ```python
  157: def extraer_campos_completos(row_dict_or_tuple):
  158:     """
  159:     Extrae y estandariza los 7 campos tecnicos requeridos (R1):
  ...
  167:     """
  168:     if isinstance(row_dict_or_tuple, dict):
  169:         d = row_dict_or_tuple
  ...
  188:         return {
  189:             "id_activo": str(d.get("id_activo") or d.get("id", "") or "").strip(),
  190:             "equipo": nom_eq,
  191:             "modelo": modelo,
  192:             "serial": serial,
  193:             "responsable": str(d.get("responsable", "") or "").strip(),
  194:             "sede": str(d.get("sede", "") or "").strip(),
  195:             "estado": str(d.get("estado", "Activo") or "Activo").strip(),
  196:         }
  197:     return {}
  ```
- **Observed Defect**: The parameter is named `row_dict_or_tuple`, and the docstring / contract promises tuple support (e.g., from `openpyxl.worksheet.iter_rows(values_only=True)`, database queries, or `Treeview` value tuples). However, line 168 only evaluates `isinstance(..., dict)` and immediately exits with `return {}` for any `tuple` or `list`.
- **Positional Mapping in Clemont Data (`conciliador_universal.py:527-532, 1948-1960`)**:
  - Full inventory row (10–12 columns, 0-based):
    - `0`: Item / Correlativo / ID (`id_activo`)
    - `1`: Cliente (`CLEMONT`)
    - `2`: Sede / Centro Operativo (`COL_SEDE = 3`, 0-based `2`)
    - `3`: Equipo / Hostname / Servidor (`COL_EQUIPO = 4`, 0-based `3`)
    - `4`: Sistema Operativo / Tipo
    - `5`: Fabricante / Ref / Modelo
    - `8`: Responsable (`COL_RESP = 9`, 0-based `8`)
    - `9`: Estado (`Activo`)
    - `10`: Observaciones y Serial (`COL_OBS = 11`, 0-based `10`), e.g., `"Serial equipo: SDK0T76530; Modelo: Lenovo 83A0"`
    - `11`: Verificado / Auditoría (`COL_VERIF = 12`, 0-based `11`)
  - Treeview row (7 columns, 0-based): `(item_num, equipo, serial, estado, sede, responsable, observaciones)`
  - Canonical row (7 columns, 0-based): `(id_activo, equipo, modelo, serial, responsable, sede, estado)`
  - Preset row (4 columns, 0-based): `(equipo, serial, estado, sede)`
  - OpenPyXL `Cell` tuple: Elements may be `openpyxl.cell.Cell` objects possessing a `.value` attribute.

---

### Observation 1.3: Schema Alignment for `generar_plantilla_auditoria_qr`
- **File & Line**: `conciliador_universal.py:348-425`
- **Sheet 1 (`Maestro_Activos`)**:
  - Headers (9 columns): `["ID_Activo", "Categoria", "Tipo_Equipo", "Equipo / Host", "Modelo_Marca", "Serial_Fabricante", "Ubicacion_Oficial", "Responsable", "Estado_Oficial"]`
  - Autofilter: `ws_m.auto_filter.ref = f"A1:I{len(equipos) + 1}"`
  - Freeze Panes: `ws_m.freeze_panes = "A2"`
  - Background styling: Corporate Dark Blue (`HDR_BLUE = PatternFill("1F4E78")`), white bold text.
- **Sheet 2 (`Escaneos_Campo`)**:
  - Headers (8 columns): `["ID_Escaneado", "Ubicacion_Hallada", "Fecha_Hora_Escaneo", "Operador", "Estado_Fisico", "Observaciones", "Validacion_Catalogo", "Resultado_Sede"]`
  - Columns 1–6 (A–F): Input columns for manual or automated wireless scanner capture.
  - Column 7 (G, `Validacion_Catalogo`):
    `=IF(ISBLANK(A{r}), "", IFERROR(VLOOKUP(A{r}, Maestro_Activos!A:D, 4, FALSE), "ALTA NUEVA"))`
    - Evaluates: Lookup `A{r}` in `Maestro_Activos` range A:D, returning column 4 (`Equipo / Host`).
  - Column 8 (H, `Resultado_Sede`):
    `=IF(ISBLANK(A{r}), "", IF(VLOOKUP(A{r}, Maestro_Activos!A:G, 7, FALSE)=B{r}, "CONCILIADO OK", "TRASLADO NO REGISTRADO"))`
    - Evaluates: Lookup `A{r}` in `Maestro_Activos` range A:G, returning column 7 (`Ubicacion_Oficial`). Compares against scanned location `B{r}` (`Ubicacion_Hallada`).
  - Autofilter: `ws_e.auto_filter.ref = "A1:H51"`
  - Freeze Panes: `ws_e.freeze_panes = "A2"`
  - Background styling: Corporate Dark Green (`HDR_GREEN = PatternFill("375623")`), white bold text.
- **Documentation Alignment**:
  - `GUIA_ESCANEO_MOVIL.md` lines 73–99 defines `Maestro_Activos` with 9 columns (A to I, explicitly listing `Modelo_Marca` at Col E) and `Escaneos_Campo` with Columns A to F for field capture.
  - `PROJECT.md` line 27 explicitly states: `"Generate Maestro_Activos (8-9 cols with Modelo/Marca) and Escaneos_Campo (6 cols) with cross-reference formulas and corporate styling"`.
- **Root Cause of Test Mismatch**:
  In `test_qa_inventario.py:280-296` (`test_F3_1`) and `303-317` (`test_F3_3`), the previous test author did not invoke `cu.generar_plantilla_auditoria_qr()`. Instead, the test constructed a dummy `openpyxl.Workbook()`, appended an 8-column header (omitting `Modelo_Marca`) to `Maestro_Activos`, appended a 5-column header to `Escaneos_Campo`, and asserted `max_column == 8` and `max_column == 5`. This self-certification created artificial drift between the test assertions and the genuine production template generator.

---

## 2. Logic Chain

1. **Defect 1 (`log()` & `log_qr()`)**:
   - `test_F7_2` asserts regex `\[\d{2}:\d{2}:\d{2}\]` on `app.txt_log.get("1.0", "end")`.
   - `ConciliadorUniversal.log()` at line 2113 inserts `msg + "\n"` without any timestamp.
   - Therefore, the test fails whenever executed against production code.
   - Remediation requires inserting `ts = datetime.now().strftime("[%H:%M:%S] ")` prior to `msg` in both `log()` and `log_qr()`.

2. **Defect 2 (`extraer_campos_completos()`)**:
   - The function contract specifies `row_dict_or_tuple`.
   - Production callers or external pipelines may pass tuples (from Excel rows or Treeview).
   - Currently, non-dict inputs evaluate `isinstance(row_dict_or_tuple, dict) == False` and return `{}`.
   - By unpacking elements (unwrapping `.value` if an openpyxl `Cell` is passed) and mapping indices by length (`>= 10` for Clemont full row, `7` for Treeview/canonical, `4` for presets), the function converts any tuple into the intermediate dictionary and executes the existing regex and cleaning algorithms for serial and model extraction.

3. **Defect 3 (Schema Alignment & Self-Certifying Tests)**:
   - The production implementation of `generar_plantilla_auditoria_qr` is complete and robust: 9 columns for `Maestro_Activos` (retaining `Modelo_Marca`), 8 columns for `Escaneos_Campo` (6 data columns + 2 VLOOKUP validation formulas).
   - The formula references `Maestro_Activos!A:D, 4` and `Maestro_Activos!A:G, 7` correctly align with column 4 (`Equipo / Host`) and column 7 (`Ubicacion_Oficial`).
   - The schema drift was an artifact of dummy workbooks constructed within `test_qa_inventario.py`.
   - Remediation requires:
     a) Keeping `generar_plantilla_auditoria_qr` intact with 9 columns in `Maestro_Activos` and 8 columns in `Escaneos_Campo`.
     b) Refactoring the 20 self-certifying tests in `test_qa_inventario.py` to directly call `cu.generar_plantilla_auditoria_qr()`, `cu.generar_etiqueta_individual()`, `app._pegar_desde_portapapeles()`, and `app._ejecutar()`.

---

## 3. Caveats

- **Read-Only Scope**: In compliance with the Teamwork Explorer identity and system constraints, this report formulates the exact code remediation strategies and drop-in code snippets, but does NOT modify `conciliador_universal.py` or `test_qa_inventario.py`.
- **GUI Headless Execution**: Tests executing Tkinter GUI operations require headless initialization (`create_headless_app()`), which is already established in `test_qa_inventario.py`.
- **No other caveats**. All findings were verified directly by static analysis and inspection of the project files.

---

## 4. Conclusion & Actionable Code Remediation Strategies

### Strategy A: Remediation for `worker_m1` (`conciliador_universal.py`)

#### A.1. Exact Fix for `ConciliadorUniversal.log()` and `log_qr()`
In `conciliador_universal.py`:

**Target 1: Lines 2112–2116 (`log`)**:
```python
<<<< BEFORE
    def log(self, msg, tag="info"):
        self.txt_log.insert(tk.END, msg + "\n", tag)
        self.txt_log.see(tk.END)
        self.root.update_idletasks()
==== AFTER
    def log(self, msg, tag="info"):
        ts = datetime.now().strftime("[%H:%M:%S] ")
        self.txt_log.insert(tk.END, f"{ts}{msg}\n", tag)
        self.txt_log.see(tk.END)
        self.root.update_idletasks()
>>>>
```

**Target 2: Lines 1015–1021 (`log_qr`)**:
```python
<<<< BEFORE
    def log_qr(self, msg, tag="info"):
        """Registra un mensaje en la bitacora de la pestana QR."""
        if hasattr(self, "txt_qr_log") and self.txt_qr_log:
            self.txt_qr_log.insert(tk.END, msg + "\n", tag)
            self.txt_qr_log.see(tk.END)
            self.root.update_idletasks()
==== AFTER
    def log_qr(self, msg, tag="info"):
        """Registra un mensaje en la bitacora de la pestana QR."""
        if hasattr(self, "txt_qr_log") and self.txt_qr_log:
            ts = datetime.now().strftime("[%H:%M:%S] ")
            self.txt_qr_log.insert(tk.END, f"{ts}{msg}\n", tag)
            self.txt_qr_log.see(tk.END)
            self.root.update_idletasks()
>>>>
```

---

#### A.2. Exact Fix for `extraer_campos_completos(row_dict_or_tuple)`
In `conciliador_universal.py`:

**Target: Lines 157–197**:
```python
<<<< BEFORE
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
        d = row_dict_or_tuple
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
            "serial": serial,
            "responsable": str(d.get("responsable", "") or "").strip(),
            "sede": str(d.get("sede", "") or "").strip(),
            "estado": str(d.get("estado", "Activo") or "Activo").strip(),
        }
    return {}
==== AFTER
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
    Soporta diccionarios, tuplas/listas de valores primitivos, y tuplas de celdas openpyxl (Cell).
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
            obs_val = str(raw_vals[10] or "").strip() if n > 10 else ""
            est_val = str(raw_vals[11] if n > 11 and raw_vals[11] is not None else (raw_vals[9] if n > 9 and raw_vals[9] is not None else "Activo")).strip()
            fab_val = str(raw_vals[5] or raw_vals[4] or "").strip() if n > 5 else ""
            d = {
                "id_activo": str(raw_vals[0] or "").strip(),
                "sede": str(raw_vals[2] or "").strip() if n > 2 else "",
                "equipo": str(raw_vals[3] or "").strip() if n > 3 else "",
                "fabricante": fab_val,
                "responsable": str(raw_vals[8] or "").strip() if n > 8 else "",
                "estado": est_val or "Activo",
                "observaciones": obs_val,
                "serial_raw": obs_val,
            }
        elif n == 7:
            # 7 columnas: Puede ser Treeview (item, equipo, serial, estado, sede, resp, obs)
            # o Tupla Canónica (id, equipo, modelo, serial, responsable, sede, estado)
            if str(raw_vals[6] or "").strip().upper() in ("ACTIVO", "OK", "TRASLADO", "BAJA", "VERIFICADO"):
                d = {
                    "id_activo": str(raw_vals[0] or "").strip(),
                    "equipo": str(raw_vals[1] or "").strip(),
                    "modelo": str(raw_vals[2] or "").strip(),
                    "serial": str(raw_vals[3] or "").strip(),
                    "responsable": str(raw_vals[4] or "").strip(),
                    "sede": str(raw_vals[5] or "").strip(),
                    "estado": str(raw_vals[6] or "Activo").strip(),
                }
            else:
                d = {
                    "id_activo": str(raw_vals[0] or "").strip(),
                    "equipo": str(raw_vals[1] or "").strip(),
                    "serial": str(raw_vals[2] or "").strip(),
                    "estado": str(raw_vals[3] or "Activo").strip(),
                    "sede": str(raw_vals[4] or "").strip(),
                    "responsable": str(raw_vals[5] or "").strip(),
                    "observaciones": str(raw_vals[6] or "").strip(),
                }
        elif n == 4:
            # Presets Clemont (equipo, serial, estado, sede)
            d = {
                "id_activo": "",
                "equipo": str(raw_vals[0] or "").strip(),
                "serial": str(raw_vals[1] or "").strip(),
                "estado": str(raw_vals[2] or "Activo").strip(),
                "sede": str(raw_vals[3] or "").strip(),
            }
        else:
            d = {
                "id_activo": str(raw_vals[0] or "").strip() if n > 0 else "",
                "equipo": str(raw_vals[1] or "").strip() if n > 1 else "",
                "serial": str(raw_vals[2] or "").strip() if n > 2 else "",
                "estado": str(raw_vals[3] or "Activo").strip() if n > 3 else "Activo",
                "sede": str(raw_vals[4] or "").strip() if n > 4 else "",
                "responsable": str(raw_vals[5] or "").strip() if n > 5 else "",
                "observaciones": str(raw_vals[6] or "").strip() if n > 6 else "",
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
        "serial": serial,
        "responsable": str(d.get("responsable", "") or "").strip(),
        "sede": str(d.get("sede", "") or "").strip(),
        "estado": str(d.get("estado", "Activo") or "Activo").strip(),
    }
>>>>
```

---

### Strategy B: Remediation for `test_writer_e2e` (`test_qa_inventario.py`)

The test suite must eliminate all 20 self-certifying tests and directly invoke production code:

| # | Test Method | Self-Certifying Defect | Genuine Remediation Implementation |
|---|---|---|---|
| 1 | `test_F2_4` | Created local `qr_data` string | Invoke `cu.generar_etiqueta_individual(item, out_png)` and decode QR or verify output attributes. |
| 2 | `test_F3_1` | Created local dummy workbook with 8 cols | Call `cu.generar_plantilla_auditoria_qr(out_xlsx, sample_equipos)`; assert `max_column == 9` and check `"Modelo_Marca"` in headers. |
| 3 | `test_F3_3` | Created local dummy workbook with 5 cols | Call `cu.generar_plantilla_auditoria_qr(out_xlsx, sample_equipos)`; assert `max_column == 8` and check formula columns 7 & 8. |
| 4 | `test_F3_5` | Set autofilter on local mock workbook | Call `cu.generar_plantilla_auditoria_qr(out_xlsx, sample_equipos)`; assert `ws.auto_filter.ref` on generated file. |
| 5 | `test_F6_2` | Evaluated ternary on local string | Invoke `app._seleccionar_archivo` or output path computation method. |
| 6 | `test_B5_5` | Evaluated `f"=SUM(H{...})"` on local string | Run `app._ejecutar()` on temporary copy of real Excel and inspect cell containing `=SUM(...)`. |
| 7 | `test_B6_1` | Evaluated `normalizar("LAPTOP-01")` | Set `app.chk_col_serial.set(False)` and run reconciliation test to verify match by hostname only. |
| 8 | `test_B6_2` | Evaluated `similitud_serial(...)` | Set `app.chk_col_equipo.set(False)` and run reconciliation test to verify match by serial only. |
| 9 | `test_B7_1` | Tested `os.makedirs` | Test `app._obtener_carpeta_qr()` with path containing spaces. |
| 10 | `test_B7_2` | Tested `os.makedirs` | Test `app._obtener_carpeta_qr()` with path containing parentheses. |
| 11 | `test_B7_3` | Tested `os.makedirs(exist_ok=True)` | Call `app._qr_exportar_excel()` twice and verify idempotency. |
| 12 | `test_B7_4` | Openpyxl save/reload locally | Execute `app._qr_exportar_excel()` overwriting an existing file without permission crash. |
| 13 | `test_B7_5` | Tested `os.path.abspath` | Test `app._obtener_carpeta_qr()` converts relative paths to absolute. |
| 14 | `test_B8_1` | Evaluated inline `if/else` | Set `app.qr_contenido_modo.set("solo_id")` and call `app._generar_etiqueta_individual(item)`. Verify QR contains only ID. |
| 15 | `test_B8_2` | Created local `qr_data` string | Set `app.qr_contenido_modo.set("completo")` and call `app._generar_etiqueta_individual(item)`. Verify QR contains all 7 fields. |
| 16 | `test_C2` | Created local dummy workbook | Call `cu.generar_plantilla_auditoria_qr(out_xlsx, sample_equipos)`; verify both `Maestro_Activos` (9 cols) and `Escaneos_Campo` (8 cols). |
| 17 | `test_C3` | Tested strings locally | Run `app._ejecutar()` on a dataset with 4 equipment cases (OK, Traslado, Alta, Faltante) and verify counts in `app.tree`. |
| 18 | `test_C4` | Inserted rows manually into workbook | Run `app._ejecutar()` with an alta equipment and inspect generated Excel file to verify total formula and row insertion. |
| 19 | `test_C5` | Evaluated inline `if` statement | Run `app._ejecutar()` with `app.chk_col_serial.set(False)` and verify match succeeds purely on Host. |
| 20 | `test_C6` | Formatted local string | Generate QR with `cu.generar_etiqueta_individual` and cross-reference decoded payload against `Maestro_Activos` row from `cu.generar_plantilla_auditoria_qr`. |
| 21 | `test_C8` | Split TSV manually with hardcoded indices | Mock `app.root.clipboard_get()` with 12-column TSV and call `app._pegar_desde_portapapeles()`; verify items inserted in `app.tree`. |
| 22 | New Unit Test | N/A | Add dedicated test verifying `extraer_campos_completos()` with: (a) dict, (b) 12-element tuple, (c) 7-element tuple, (d) 4-element tuple, (e) openpyxl `Cell` row tuple. |

---

## 5. Verification Method

To independently verify this remediation plan before and after implementation:

1. **Verify Log Timestamps**:
   - Inspect `conciliador_universal.py` at lines 2112–2116.
   - Run `test_F7_2_metodo_log_agrega_marcas_temporales`:
     ```python
     app.log("Test log entry")
     assert re.search(r"\[\d{2}:\d{2}:\d{2}\]", app.txt_log.get("1.0", "end"))
     ```
2. **Verify Tuple / List Extraction in `extraer_campos_completos`**:
   - Pass Clemont 12-column row tuple:
     ```python
     row_tuple = (1, "CLEMONT", "Sótano", "LAPTOP-2895F563", "Win11", "Lenovo 83A0", "Laptop", 1, "DANIEL PANESSO", "Activo", "Serial equipo: SDK0T76530; Modelo: Lenovo 83A0", "OK")
     res = extraer_campos_completos(row_tuple)
     assert res["id_activo"] == "1"
     assert res["equipo"] == "LAPTOP-2895F563"
     assert res["serial"] == "SDK0T76530"
     assert "LENOVO" in res["modelo"].upper()
     assert res["responsable"] == "DANIEL PANESSO"
     assert res["sede"] == "Sótano"
     assert res["estado"] in ("OK", "Activo")
     ```
   - Pass 7-column Treeview tuple:
     ```python
     res = extraer_campos_completos((1, "LAP-01", "SN123456", "OK", "Sótano", "JUAN", "Obs"))
     assert res["equipo"] == "LAP-01"
     assert res["serial"] == "SN123456"
     ```
3. **Verify Template Schema and Formulas**:
   - Call `generar_plantilla_auditoria_qr(tmp_file, sample_equipos)`.
   - Open with `openpyxl.load_workbook(tmp_file)`.
   - Verify `wb["Maestro_Activos"].max_column == 9`.
   - Verify `wb["Escaneos_Campo"].max_column == 8`.
   - Verify formula in cell G2: `=IF(ISBLANK(A2), "", IFERROR(VLOOKUP(A2, Maestro_Activos!A:D, 4, FALSE), "ALTA NUEVA"))`.
   - Verify formula in cell H2: `=IF(ISBLANK(A2), "", IF(VLOOKUP(A2, Maestro_Activos!A:G, 7, FALSE)=B2, "CONCILIADO OK", "TRASLADO NO REGISTRADO"))`.
4. **Invalidation Conditions**:
   - Any test constructing in-memory dummy objects and asserting against them rather than invoking `cu` or `ConciliadorUniversal`.
   - `extraer_campos_completos` returning `{}` when passed a tuple or list.
   - `app.log()` emitting log lines without `[HH:MM:SS]`.
