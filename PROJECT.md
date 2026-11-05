# Project: Clemont Inventory Reconciliation & QR Asset Tagging System

## Architecture
- **Language / Runtime**: Python 3.10+ (Windows environment)
- **Core Modules & Boundaries**:
  - `conciliador_universal.py` & `GUIA_ESCANEO_MOVIL.md`: Core Application Module (Desktop GUI, Excel ingestion & reconciliation engine, 300 DPI QR/sticker generator, audit template generator, mobile scanning operational guide).
  - `test_qa_inventario.py`: Automated QA & E2E Testing Module (Headless automated test runner validating data integrity against real Excel files, 300 DPI image specs, QR readability, and 0 exit code).
- **Libraries & Dependencies**:
  - `openpyxl`: Excel workbook reading, writing, styling (corporate headers, thin borders, dynamic formulas).
  - `qrcode`: 2D QR matrix generation.
  - `Pillow` (`PIL`): Image creation, layout composition, typography, 300 DPI metadata injection (`dpi=(300, 300)`).
  - `tkinter` & `ttk`: Desktop graphical interface.
- **Data Flow**:
  1. Input: `Inventario_Equipos_Clemont (2) (2).xlsx` (Sheet: `"Inventario de Equipos"`).
  2. Ingestion & Normalization: Row 5 headers, dynamic extraction of the 7 technical fields (ID, Hostname, Marca/Modelo, Serial, Responsable, Sede, Estado), normalization of serials (splitting `;`, stripping whitespace, tolerant character substitution).
  3. Processing & Tagging: Generation of 300 DPI PNG stickers and QR images; generation of `Inventario_Equipos_Clemont_Con_QR.xlsx` with linked sheets `Maestro_Activos` and `Escaneos_Campo`.
  4. Reconciliation: Cross-referencing physical field audit scans against master inventory, reporting: `CONCILIADO OK`, `TRASLADO NO REGISTRADO`, `ALTA / SOBRANTE`, and `FALTANTE`.
  5. Output: `AUDITADO_Inventario_Equipos_Clemont (2) (2).xlsx` (with Col L verification, altas inserted before Total, updated `=SUM()` formula, peripheral sub-tables, and `"Reporte de Auditoria"` sheet).

## Feature Inventory
| # | Feature | Description | Milestone | Source |
|---|---------|-------------|-----------|--------|
| 1 | Fix Column Shift in Clipboard Parsing | Correct `_pegar_desde_portapapeles()` false-positive header detection on rows with "SERIAL" to ensure Host, Sede, and Estado are never shifted to Item # | M1 | Survey / R1 |
| 2 | Robust Serial & Delimiter Extraction | Fix `extraer_seriales()` to split on `;` and strip prefixes without attaching trailing semicolons or swallowing models | M1 | Survey / R1 |
| 3 | 7-Field Technical Spec Extraction | Extract all 7 fields (ID, Host, Modelo/Marca, Serial, Responsable, Sede, Estado) cleanly from Col A-L | M1 | Survey / R1 |
| 4 | 300 DPI QR & Physical Sticker Generation | Render stickers (50x25mm and 100x50mm) and individual QR images at 300 DPI with `dpi=(300, 300)` metadata, legible typography, and all 7 fields in QR payload | M1 | Survey / R1, R4 |
| 5 | Master & Field Audit Excel Templates | Generate `Maestro_Activos` (8-9 cols with Modelo/Marca) and `Escaneos_Campo` (6 cols) with cross-reference formulas and corporate styling | M1 | Survey / R2 |
| 6 | Mobile Scan Integration & Operational Guide | Document and integrate wireless scanning workflows (Barcode to PC, Excel Mobile, Bluetooth 2D) into `GUIA_ESCANEO_MOVIL.md` | M1 | Survey / R2 |
| 7 | Column Checkboxes (Chulos) in GUI | Add individual checkboxes per comparison column to dynamically toggle matching criteria | M1 | Survey / R3 |
| 8 | Output Directory Selector in GUI | Add folder chooser (`askdirectory`) affecting audited Excel and `etiquetas_qr/` | M1 | Survey / R3 |
| 9 | 1-Click Open Buttons in GUI | Add direct buttons to open audited Excel and output folder via `os.startfile` in both tabs | M1 | Survey / R3 |
| 10 | Wide Log Console with Scrollbar | Implement `ttk.Scrollbar` on `txt_log` and `txt_qr_log` with auto-scrolling and color coding | M1 | Survey / R3 |
| 11 | Headless QA Test Runner | Build `test_qa_inventario.py` to execute automated validation against real Excel files | M2 | Survey / R5 |
| 12 | Column Offset & Data Integrity Verification | Automated test validating no cell scrambling between input and audited output | M2 | Survey / R5 |
| 13 | QR Readability & 7-Field Decodability Test | Automated validation decoding generated QR images to ensure 100% readability and presence of all 7 fields | M2 | Survey / R5 |

## Milestones
| # | Name | Scope | Dependencies | Status |
|---|------|-------|-------------|--------|
| M1 | Application Implementation (`conciliador_universal.py` & `GUIA_ESCANEO_MOVIL.md`) | Implement all R1-R4 features: fix column shift, robust serial parsing, 7-field extraction, 300 DPI QR/stickers, `Maestro_Activos`/`Escaneos_Campo` templates, `GUIA_ESCANEO_MOVIL.md`, column checkboxes, output directory chooser, 1-click buttons, scrollbar consoles | none | IN_PROGRESS |
| M2 | Final Milestone: Automated QA Test Suite & E2E Validation (`test_qa_inventario.py`) | Build comprehensive automated test suite and execute headless tests against real Excel files verifying 0 exit code, data integrity, and QR readability (R5) | none (Dual Track) | IN_PROGRESS |

## Interface Contracts
### `conciliador_universal.py` ↔ Data Extraction & Model
- `extraer_campos_completos(row_dict_or_tuple)` -> `dict`:
  - Keys: `'id_activo'`, `'equipo'`, `'modelo'`, `'serial'`, `'responsable'`, `'sede'`, `'estado'`
- `extraer_seriales(texto: str)` -> `list[str]`:
  - Splits on `;`, `,`, newlines, strips `Serial equipo:`, `Modelo:`, validates length >= 6, removes punctuation.
- `generar_etiqueta_individual(datos_equipo: dict, ruta_salida: str, formato: str, dpi: int = 300)` -> `str`:
  - Creates PNG image at `dpi=(300, 300)` with 7 fields visually laid out and QR matrix embedded. Returns absolute file path.
- `generar_plantilla_auditoria_qr(ruta_salida: str, equipos: list[dict])` -> `str`:
  - Creates `Inventario_Equipos_Clemont_Con_QR.xlsx` with sheets `Maestro_Activos` and `Escaneos_Campo`.

### Automated QA Test Contract
- Invocation: `python test_qa_inventario.py`
- Exit Code: `0` on 100% assertions passed; non-zero on any failure.
- Checks:
  - Real Excel files existence and structure.
  - No column scrambling (Host, Sede, Serial, Responsable integrity).
  - 300 DPI image generation and metadata.
  - QR matrix decoding and verification of all 7 fields.

## Code Layout
- `c:\Users\Analista-TI\Desktop\entrega inicial\Scrip-Inventario\`
  - `conciliador_universal.py`: Main application code (Owned exclusively by Worker M1).
  - `GUIA_ESCANEO_MOVIL.md`: Operational guide (Owned exclusively by Worker M1).
  - `test_qa_inventario.py`: QA automated test suite (Owned exclusively by Test Writer E2E).
  - `Inventario_Equipos_Clemont (2) (2).xlsx`: Real baseline input Excel file (Read-only).
  - `AUDITADO_Inventario_Equipos_Clemont (2) (2).xlsx`: Reference audited Excel file (Read-only).
