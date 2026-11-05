# Handoff Report: Codebase Survey & Diagnosis of `conciliador_universal.py`

**Agent**: `explorer_survey_1` (teamwork_explorer)  
**Date**: 2026-11-05T17:40:00Z  
**Target File**: `c:\Users\Analista-TI\Desktop\entrega inicial\Scrip-Inventario\conciliador_universal.py`  
**Working Directory**: `c:\Users\Analista-TI\Desktop\entrega inicial\Scrip-Inventario\.agents\teamwork\explorer_survey_1`  
**Reference Document**: `c:\Users\Analista-TI\Desktop\entrega inicial\Scrip-Inventario\.agents\teamwork\ORIGINAL_REQUEST.md`

---

## 1. Observation

### 1.1 Codebase Layout and File Inventory
Inspection of `c:\Users\Analista-TI\Desktop\entrega inicial\Scrip-Inventario\` revealed:
- `conciliador_universal.py`: Main application script (2,408 lines, 110,290 bytes).
- `scrip1.py`: Historical v1.0 script (835 lines, 25,095 bytes), single-tab prototype.
- `ejecutar.bat`: Windows launcher executing `python conciliador_universal.py`.
- `requirements.txt`: Lists `openpyxl>=3.1.0`, `qrcode[pil]>=8.0.0`, `Pillow>=10.0.0`, `pandas>=2.0.0`.
  - *Observation*: `pandas` is listed in `requirements.txt` but is **not imported or used** anywhere in `conciliador_universal.py`. Only `openpyxl` is used for Excel manipulation.
- Real test Excel files present:
  - `Inventario_Equipos_Clemont (2) (2).xlsx` (43,488 bytes)
  - `AUDITADO_Inventario_Equipos_Clemont (2) (2).xlsx` (44,726 bytes)
- Output directory `etiquetas_qr/`: Contains previously generated artifacts, including corrupt files such as `LAP-2026-0001_Item.png`, `LAP-2026-0002_1.png`, `LAP-2026-0003_2.png`, `LAP-2026-0004_3.png`, and `Inventario_Equipos_Clemont_Con_QR.xlsx`.

### 1.2 Structure and Anatomy of `conciliador_universal.py`
The script is organized into single-file modules:
1. **Lines 1–50**: Module docstrings and conditional imports (`qrcode`, `PIL.Image`, `PIL.ImageDraw`, `PIL.ImageFont`, flag `QR_DISPONIBLE`).
2. **Lines 52–150**: Serial and text normalization routines:
   - `normalizar(texto)`: Alphanumeric uppercase stripping.
   - `normalizar_tolerante(texto)`: Character substitution (`O`→`0`, `Z`→`2`, `S`→`5`, `I`→`1`, `B`→`8`).
   - `similitud_serial(s1, s2)`: Heuristic comparison using Levenshtein distance and substring search.
   - `_levenshtein(a, b)`: Dynamic programming edit distance.
   - `extraer_seriales(texto)`: Regular expression extractor. Splitting regex: `r"[/,\n]|\s{2,}"`.
3. **Lines 152–175**: Styling constants (openpyxl `Font`, `PatternFill`, `Border`, `Side`, `Alignment`).
4. **Lines 177–198**: `_obtener_fuente(size, bold=False)`: PIL `FreeTypeFont` loader with fallback to `ImageFont.load_default()`.
5. **Lines 204–2408**: Class `ConciliadorUniversal`:
   - Class-level constants: `PRESET_CLEMONT` (20 laptop tuples), `PRESET_TABLETS` (2 tuples), `PRESET_MONITORES` (12 tuples), default columns (`COL_SEDE=3`, `COL_EQUIPO=4`, `COL_RESP=9`, `COL_OBS=11`, `COL_VERIF=12`).
   - `__init__`: Initializes Tkinter root (`1140x900`), variables, and invokes GUI builder.
   - GUI Builders:
     - `_construir_gui`: Creates title banner and a 2-tab `ttk.Notebook`.
     - `_construir_tab_conciliacion`: Tab 1 with File selection, Sheet selector, Column mapping with checkboxes, 7-column Treeview for physical audit, action buttons, and bitácora console.
     - `_construir_tab_qr`: Tab 2 with Source selector (Excel vs Treeview vs Preset), Format selector (`50x25mm`, `67x25mm`, `100x50mm`), QR mode ("completo" vs "solo_id"), Action buttons ("Generar 34 Etiquetas QR", "Crear Planillas", "Exportar Base"), live scan tester, and QR bitácora console.
   - Core Processing Methods:
     - `_obtener_lista_equipos_qr`: Collects assets for QR/sticker generation.
     - `_generar_etiqueta_individual`: Generates 300 DPI PNG label.
     - `_qr_crear_planilla`: Arranges labels onto Letter-sized sheet (2550 x 3300 px at 300 DPI).
     - `_qr_exportar_excel`: Creates workbook with `Maestro_Activos`, `Escaneos_Campo`, `Instructivo_Auditoria`.
     - `_pegar_desde_portapapeles`: Parses TSV data from clipboard into Treeview.
     - `_auto_detectar_columnas_hoja`: Scans header row in selected sheet to populate column dropdowns.
     - `_ejecutar`: Performs row reconciliation, duplicate checks, alta insertions, formula fix, and report generation.

### 1.3 Verbatim Findings in Real Excel Files
1. **`Inventario_Equipos_Clemont (2) (2).xlsx`**:
   - Sheets: `['Resumen y Tablas Dinámicas', 'Inventario de Equipos', 'Inventario de Equipos (2)']`.
   - Sheet `Inventario de Equipos` has header at Row 5:
     `['Item', 'Cliente', 'Centro Operativo / Sede', 'Equipo / Servidor', 'Sistema Operativo / Software', 'Fabricante', 'Tipo Licencia', 'Cantidad', 'Responsable / Usuario', 'Estado', 'Observaciones y Serial', 'Verificado']`
   - Data rows: Row 6 to 76 (71 equipment rows). Row 77 is `TOTAL EQUIPOS`.
   - Sample data row (Row 6):
     - Col 1 (A): `20` (Item #)
     - Col 2 (B): `'CLEMONT'` (Cliente)
     - Col 3 (C): `'Administrativa 2do piso'` (Centro Operativo / Sede)
     - Col 4 (D): `'LAPTOP-2895F563'` (Equipo / Servidor)
     - Col 5 (E): `'WINDOWS 11 PRO'` (Sistema Operativo)
     - Col 6 (F): `'Microsoft'` (Fabricante)
     - Col 7 (G): `'Sistema operativo'` (Tipo Licencia)
     - Col 8 (H): `1` (Cantidad)
     - Col 9 (I): `'DANIEL PANESSO'` (Responsable / Usuario)
     - Col 10 (J): `'Activo'` (Estado)
     - Col 11 (K): `'Serial equipo: SDK0T76530; Modelo: LENOVO 83A0'` (Observaciones y Serial)
     - Col 12 (L): `None` (Verificado)
2. **`AUDITADO_Inventario_Equipos_Clemont (2) (2).xlsx` (Existing Corrupted Run)**:
   - Row 6 in `Inventario de Equipos`:
     `[20, 'CLEMONT', 'Administrativa 2do piso', '2', 'WINDOWS 11 PRO', 'Microsoft', 'Sistema operativo', 1, 'DANIEL PANESSO', 'OK', 'Serial equipo: CLEMONT; Modelo: LENOVO 83A0', None]`
     - Col 4 (Equipo / Servidor) was corrupted to `'2'`!
     - Col 11 (Observaciones y Serial) was corrupted to `'Serial equipo: CLEMONT; Modelo: LENOVO 83A0'`!
   - In sheet `Reporte de Auditoria`:
     - Row 12: `[1, '2', 'Serial equipo: CLEMONT; Modelo: LENOVO 83A0', 'CLEMONT', 'OK', 'Administrativa 2do piso', 'Verificado OK']`
     - Row 14: `[3, '1', 'Serial equipo: CLEMONT; Modelo: HP LAPTOP 14 D10XXX', 'CLEMONT', 'OK', 'Administrativa 2do piso', 'Verificado OK']`
     - Row 15: `[4, '3', 'Serial equipo: CLEMONT; Modelo: Vivobook ASUS x515EA', 'CLEMONT', 'OK', 'Administrativa 2do piso', 'Verificado OK']`
     - 55 false "Altas nuevas" generated.
3. **Artifacts in `etiquetas_qr/`**:
   - Files named: `LAP-2026-0001_Item.png`, `LAP-2026-0002_1.png`, `LAP-2026-0003_2.png`, `LAP-2026-0004_3.png` ... `LAP-2026-0072_71.png`.
   - Individual images show: Hostname set to the Item number (e.g., "1"), Sede set to hostname ("LAPTOP-2895F563"), and Estado set to Sede ("Administrativa 2do piso").

---

## 2. Logic Chain

### 2.1 The Root Cause of Column Offsets/Shifts
Tracing the exact execution flow that produced the corrupted audit and labels:

1. **Step 1: Clipboard ingestion via `_pegar_desde_portapapeles()` (lines 1653–1714)**:
   - When a user copies rows from Excel, the clipboard contains TSV lines.
   - Line 1658 checks:
     ```python
     if any("SERIAL" in p or "EQUIPO" in p or "ITEM" in p or "CLIENTE" in p for p in partes_primera):
     ```
   - When copying data rows without headers, `partes_primera` is data row 1 (Row 6).
   - In Row 6, cell K is `'Serial equipo: SDK0T76530; Modelo: LENOVO 83A0'`.
   - Because cell K contains `"SERIAL"`, the condition evaluates to **True**. The code mistakenly assumes that data row 6 is a header row (`Caso A`).
2. **Step 2: Misidentification of Column Indices**:
   - Under `Caso A`, the loop scans `partes_primera`:
     - Index 0 is `'20'` (or `'1'`): does not contain `"ITEM"` or `"#"` when copied without header. `idx_item` remains `None`.
     - Index 1 is `'CLEMONT'`: does not match any keyword.
     - Index 2 is `'ADMINISTRATIVA 2DO PISO'`: does not contain `"SEDE"` or `"UBICACION"`. `idx_sede` remains `None`.
     - Index 3 is `'LAPTOP-2895F563'`: does not contain `"EQUIPO"`, `"HOST"`, `"DISPOSITIVO"`, or `"SERVIDOR"`. `idx_equipo` remains `None`.
     - Index 10 is `'SERIAL EQUIPO: ...'`: matches `"SERIAL"`. `idx_serial` is set to 10.
     - All other indices (`idx_equipo`, `idx_sede`, `idx_resp`, `idx_estado`, `idx_item`) remain `None`.
3. **Step 3: Flawed Default Fallback Indices (lines 1710–1714)**:
   ```python
   if idx_equipo is None: idx_equipo = 0
   if idx_serial is None: idx_serial = 1 if num_cols > 1 else 0
   if idx_estado is None: idx_estado = 2 if num_cols > 2 else None
   if idx_sede is None:   idx_sede   = 3 if num_cols > 3 else None
   ```
   - `idx_equipo` defaults to `0` (which is the **Item number**: 1, 2, 3...).
   - `idx_estado` defaults to `2` (which is **Sede**: "Administrativa 2do piso").
   - `idx_sede` defaults to `3` (which is **Equipo/Host**: "LAPTOP-2895F563").
   - `idx_resp` remains `None` (empty string).
4. **Step 4: Insertion into Treeview (lines 1753–1754)**:
   - Treeview columns are `("#", "Equipo / Host", "Serial Real", "Estado", "Sede", "Responsable", "Observaciones")`.
   - Values inserted:
     - Col 1 (`Equipo / Host`): Item number (`partes[0]`) → `"1"`, `"2"`, `"3"`.
     - Col 2 (`Serial Real`): Serial cell (`partes[10]`).
     - Col 3 (`Estado`): Sede cell (`partes[2]`) → `"Administrativa 2do piso"`.
     - Col 4 (`Sede`): Hostname cell (`partes[3]`) → `"LAPTOP-2895F563"`.
5. **Step 5: Propagation to Label Generation (`_obtener_lista_equipos_qr`, lines 889–940)**:
   - When reading from Treeview (`qr_fuente = "auto"`):
     - `nom_eq = vals[1]` → takes the Item number `"1"`.
     - `sede = vals[4]` → takes the Hostname `"LAPTOP-2895F563"`.
     - `estado = vals[3]` → takes the Sede `"Administrativa 2do piso"`.
   - File name construction: `clean_eq = re.sub(r"[^A-Za-z0-9_-]", "", item["equipo"])[:20]`, producing:
     `LAP-2026-0002_1.png`, `LAP-2026-0003_2.png`, etc.
   - Text printed on sticker and encoded in QR:
     - `Host`: `"1"`
     - `Sede`: `"LAPTOP-2895F563"`
     - `Estado`: `"Administrativa 2do piso"`
6. **Step 6: Propagation to Excel Audit (`_ejecutar`, lines 2040–2051)**:
   - In `_ejecutar`:
     `ws.cell(r, col_equipo).value = coincidencia["equipo"]` → overwrites Column D with `"2"`, `"1"`, `"3"`.
     `ws.cell(r, col_obs).value = f"Serial equipo: {coincidencia['serial']}..."` → overwrites Column K.
   - Result: Permanent corruption of data in the audited spreadsheet.

### 2.2 Semicolon and Serial Parsing Bugs
1. In `extraer_seriales(texto)` (lines 132–147):
   - Separator regex: `re.split(r"[/,\n]|\s{2,}", str(texto))`. It does **not** include `;`.
   - In the cleaning regex: `p = re.sub(r"(?i)modelo\s*[:\-]?\s*.*", "", p).strip()`.
   - If the cell is `'Serial equipo: SDK0T76530; Modelo: LENOVO 83A0'`, stripping `Modelo:...` leaves `'SDK0T76530;'` with a trailing semicolon.
   - The trailing semicolon is preserved and printed onto stickers and QR codes as `'SDK0T76530;'`.
2. In `_obtener_lista_equipos_qr` (lines 842–845):
   - Fallback `serial = sns[0] if sns else raw_obs`.
   - When a row has no serial (e.g., Row 8 `'Modelo: HP LAPTOP 14 D10XXX'`), `sns` is empty, so `serial` becomes the entire raw observation string. The sticker then prints: `S/N: Modelo: HP LAPTOP 14 D10XXX`.

### 2.3 The 7 Required Fields Gap Analysis
Requirement R1 specifies:
1. `ID Activo` (e.g., `LAP-2026-0001`)
2. `Hostname / Equipo` (e.g., `LAPTOP-2895F563`)
3. `Marca / Modelo` (e.g., `Lenovo 83A0`)
4. `Serial S/N` (e.g., `SDK0T76530`)
5. `Responsable` (e.g., `DANIEL PANESSO`)
6. `Centro Operativo / Sede` (e.g., `Administrativa 2do piso`)
7. `Estado del activo` (e.g., `Activo / Verificado`)

Current code behavior:
- **`ID Activo`**: In `_obtener_lista_equipos_qr`, IDs are generated automatically (`cat-2026-XXXX`). In Excel `Maestro_Activos`, `ID_Activo` is exported. However, in the inventory sheet, there is no mapping for an existing asset tag column.
- **`Hostname`**: Read from `col_equipo` (Col D).
- **`Marca / Modelo`**: **Not mapped**. There is no dropdown, checkbox, or column auto-detection for Modelo or Marca. The script only searches regex `Modelo:` inside `Observaciones y Serial`, ignoring Excel Column F ("Fabricante").
- **`Serial S/N`**: Read from `col_obs` (Col K) with trailing semicolon flaws.
- **`Responsable`**: Read from `col_resp` (Col I).
- **`Centro Operativo / Sede`**: Read from `col_sede` (Col C).
- **`Estado del activo`**: Conflated with `col_verif` (Col L) and Column J (10). In the input file, Column J is `Estado` ("Activo") and Column L is `Verificado` (audit check result). The code mixes these two distinct concepts.

### 2.4 Sticker & QR Code Sizing and DPI Analysis
- Libraries: `qrcode` + `Pillow` (PIL).
- Individual sticker resolutions:
  - `100x50mm`: 1,200 x 600 px (300 DPI). Contains all fields.
  - `67x25mm`: 788 x 300 px (300 DPI). Truncates Modelo to 14 chars, Serial to 20 chars.
  - `50x25mm`: 591 x 295 px (300 DPI). **Completely omits `Estado`**, and `Modelo` is conditionally placed and often overflows or collides.
- Sheet grids: Letter size (2,550 x 3,300 px at 300 DPI).
- QR content:
  - Default mode `completo`: Encodes Empresa, Activo, Equipo, Modelo, Serial, Responsable, Sede, Estado.
  - Mode `solo_id`: Encodes only `item['id']`.
  - Discrepancy: `conciliador_universal.py` docstrings (line 991) and `README.md` (line 47) claim the QR encodes *only* the asset ID, whereas R1 strictly requires the QR to encode the complete technical specification.

### 2.5 Excel Export and R2 Synchronization Analysis
- Method `_qr_exportar_excel` (lines 1246–1343):
  - Generates `Inventario_Equipos_Clemont_Con_QR.xlsx` in `etiquetas_qr/`.
  - Creates 3 sheets: `Maestro_Activos`, `Escaneos_Campo`, `Instructivo_Auditoria`.
  - **Column mismatch**:
    - `Maestro_Activos` headers: `['ID_Activo', 'Categoria', 'Tipo_Equipo', 'Equipo / Host', 'Serial_Fabricante', 'Ubicacion_Oficial', 'Responsable', 'Estado_Oficial']`. Missing `Marca / Modelo`!
    - `Escaneos_Campo` headers: `['ID_Escaneado', 'Ubicacion_Hallada', 'Fecha_Hora_Escaneo', 'Operador', 'Observaciones']`.
  - The headers do not match the 7 required fields, and there are no Excel cross-referencing formulas (e.g. `XLOOKUP` / `VLOOKUP`) connecting `Escaneos_Campo` to `Maestro_Activos`.

### 2.6 GUI Status vs Requirement R3
| R3 Requirement | Current Status | Detailed Gap |
|---|---|---|
| **Individual checkboxes (chulos) for columns** | Partial | Implemented for Serial, Host, Sede, Responsable, Auditoría. **Missing**: Checkbox & dropdown for `Marca / Modelo`, `ID Activo`, and separation of `Estado` vs `Verificado`. |
| **Configurable save path selector** | Partial | Implemented in Tab 1 (`ruta_salida` with "Cambiar destino..."). **Missing**: Tab 2 hardcodes output folder to `"etiquetas_qr"`, with no user selector. |
| **1-click direct open buttons** | Partial | Tab 1 has `Abrir Excel Auditado` and `Abrir Carpeta`. Tab 2 has `Abrir Carpeta de Salida`. **Missing**: Tab 2 has no 1-click button to open the exported QR template Excel file. |
| **Wide log console with scrollbar** | Partial | Tab 1 has `txt_log` (Consolas 10, height 14) with vertical scrollbar. **Missing**: Tab 2 `txt_qr_log` has **no scrollbar**; output overflows when generating 71+ items. |

---

## 3. Caveats
- No code in the repository was modified during this survey (read-only compliance).
- Dynamic behavior was verified through AST analysis, static code review, and non-destructive external inspection scripts running against copies of data rows in memory.
- The repository contains an active copy of `AUDITADO_Inventario_Equipos_Clemont (2) (2).xlsx` and `etiquetas_qr/` that reflect the previous buggy run; these serve as clear reproduction evidence.

---

## 4. Conclusion
1. **Primary Bug Identified**: The data offset / shift in stickers, QR codes, and audited spreadsheets is primarily caused by **flawed header detection in `_pegar_desde_portapapeles()`** (lines 1658–1714). The presence of the word `"SERIAL"` in data cells triggers `Caso A`, causing the first data row to be treated as a header. Unmatched columns fall back to invalid defaults (`idx_equipo = 0`), assigning the item index (`1`, `2`, `3`...) to the hostname, which corrupts the Treeview, the audited spreadsheet, and the generated labels.
2. **Secondary Bugs Identified**:
   - `extraer_seriales()` fails to split on semicolons `;`, leaving trailing semicolons in serial numbers.
   - When serials are missing, the full observation text is assigned as the serial number.
   - `50x25mm` sticker format omits `Estado` and frequently clips `Modelo`.
   - `Maestro_Activos` sheet omits the `Marca / Modelo` column.
   - Tab 2 hardcodes `"etiquetas_qr"` output path and "34 activos" in labels, and its log console lacks a scrollbar.
   - `pandas` is an unused dependency in `conciliador_universal.py`.
3. **Actionable Fix Path for Implementer**:
   - Refactor `_pegar_desde_portapapeles()` with strict header validation (requiring at least 3 distinct column name matches, not single keyword substring hits).
   - Add explicit GUI mapping, auto-detection, and checkboxes for `Marca / Modelo` and `ID Activo`.
   - Fix `extraer_seriales()` to clean semicolons, normalize serial delimiters, and handle missing serials cleanly.
   - Update sticker layouts for all 3 formats (`50x25mm`, `67x25mm`, `100x50mm`) to guarantee all 7 fields are legibly displayed at 300 DPI without truncation or omission.
   - Synchronize `Maestro_Activos` and `Escaneos_Campo` with the exact 7 fields, embedding lookup formulas (`XLOOKUP` / `VLOOKUP`) and documenting wireless mobile scanner connectivity (Barcode to PC / Excel Mobile).
   - Add folder selector and 1-click open buttons to Tab 2, and add a scrollbar to `txt_qr_log`.

---

## 5. Verification Method

To independently verify the observations and diagnostic claims made in this report:

1. **Verify Corrupted Audit File**:
   Run in PowerShell:
   ```powershell
   python -c "import openpyxl; wb=openpyxl.load_workbook('AUDITADO_Inventario_Equipos_Clemont (2) (2).xlsx', data_only=True); ws=wb['Inventario de Equipos']; print('Row 6 D (Equipo):', repr(ws.cell(6,4).value)); print('Row 6 K (Serial):', repr(ws.cell(6,11).value))"
   ```
   *Expected output*: `Row 6 D (Equipo): '2'`, `Row 6 K (Serial): 'Serial equipo: CLEMONT; Modelo: LENOVO 83A0'`.

2. **Verify Clipboard Bug Simulation**:
   Execute the clipboard parsing logic against row 6 of `Inventario_Equipos_Clemont (2) (2).xlsx`:
   ```powershell
   python -c "import openpyxl; wb=openpyxl.load_workbook('Inventario_Equipos_Clemont (2) (2).xlsx', data_only=True); ws=wb['Inventario de Equipos (2)']; row_vals=[str(ws.cell(6, c).value or '') for c in range(1, 13)]; line='\t'.join(row_vals); p_upper=[p.strip().upper() for p in line.split('\t')]; print('Matches header check:', any('SERIAL' in p or 'EQUIPO' in p for p in p_upper))"
   ```
   *Expected output*: `Matches header check: True`.

3. **Verify Corrupted Label Filenames in `etiquetas_qr/`**:
   ```powershell
   Get-ChildItem -Path "etiquetas_qr" -Filter "*Item.png"
   Get-ChildItem -Path "etiquetas_qr" -Filter "*_1.png"
   ```
   *Expected output*: Lists `LAP-2026-0001_Item.png`, `LAP-2026-0002_1.png`.

4. **Verify DPI of Generated Stickers**:
   ```powershell
   python -c "from PIL import Image; im=Image.open('etiquetas_qr/LAP-2026-0001_LaptopDellLatitude.png'); print('DPI:', im.info.get('dpi')); print('Size:', im.size)"
   ```
   *Expected output*: `DPI: (300, 300)`.

