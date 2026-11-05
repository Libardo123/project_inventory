# E2E Test Infra: Clemont Inventory Reconciliation

## Test Philosophy
- Opaque-box, requirement-driven. No dependency on implementation internals.
- Methodology: Category-Partition + Boundary Value Analysis + Pairwise + Real-World Workload Testing.
- Verification exclusively against real Excel files: `Inventario_Equipos_Clemont (2) (2).xlsx` and `AUDITADO_Inventario_Equipos_Clemont (2) (2).xlsx`.

## Feature Inventory
| # | Feature | Source (requirement) | Tier 1 | Tier 2 | Tier 3 | Tier 4 |
|---|---------|---------------------|:------:|:------:|:------:|:------:|
| 1 | R1: 7 Campos Ficha Técnica sin desfase | ORIGINAL_REQUEST §R1 | 5 | 5 | ✓ | ✓ |
| 2 | R1: Stickers y Códigos QR a 300 DPI | ORIGINAL_REQUEST §R1 | 5 | 5 | ✓ | ✓ |
| 3 | R2: Estructura Maestro_Activos y Escaneos_Campo | ORIGINAL_REQUEST §R2 | 5 | 5 | ✓ | ✓ |
| 4 | R2: Guía Operativa de Escaneo Móvil | ORIGINAL_REQUEST §R2 | 5 | 5 | ✓ | ✓ |
| 5 | R3: Checkboxes (chulos) individuales por columna | ORIGINAL_REQUEST §R3 | 5 | 5 | ✓ | ✓ |
| 6 | R3: Selector de ruta y botones 1-clic | ORIGINAL_REQUEST §R3 | 5 | 5 | ✓ | ✓ |
| 7 | R3: Consola de bitácora con scrollbar | ORIGINAL_REQUEST §R3 | 5 | 5 | ✓ | ✓ |
| 8 | R5: Conciliación integral sin errores y exit code 0 | ORIGINAL_REQUEST §R5 | 5 | 5 | ✓ | ✓ |

## Test Architecture
- Test runner location: `test_qa_inventario.py`
- Invocation: `python test_qa_inventario.py`
- Pass/Fail semantics: Exit code 0 indicates all assertions passed cleanly; non-zero indicates failure.
- Output: Detailed test execution report with timings, checked fields, decoded QR payloads, and image DPI verification.

## Real-World Application Scenarios (Tier 4)
| # | Scenario | Features Exercised | Complexity |
|---|----------|--------------------|------------|
| 1 | Conciliación completa de libro Clemont con 20 laptops, 2 tablets, 12 monitores y altas | F1, F3, F8 | High |
| 2 | Generación masiva de etiquetas 300 DPI y validación decodificada de todos los QR | F1, F2, F8 | High |
| 3 | Generación y enlace de plantilla de auditoría con fórmulas cruzadas | F3, F4, F8 | Medium |
| 4 | Verificación de interfaz GUI (chulos, selector de ruta, botones 1-clic, scrollbar) | F5, F6, F7 | Medium |
| 5 | Prueba de estrés contra datos corruptos, seriales con punto y coma y filas de altas | F1, F2, F8 | High |

## Coverage Thresholds
- Tier 1: ≥5 per feature (Total ≥ 40)
- Tier 2: ≥5 per feature boundary / corrupt input (Total ≥ 40)
- Tier 3: Pairwise coverage of major feature interactions (Total ≥ 8)
- Tier 4: ≥5 realistic application scenarios (Total ≥ 5)
- **Total Minimum Test Cases**: ≥ 93 test assertions
