# Inventario y Control de Activos - Clemont v3.0

Herramienta universal de auditoría, conciliación y rotulación con códigos QR para inventario de equipos tecnológicos (Laptops, Tablets, Monitores y Periféricos).

Desarrollada en **Python 3.x + Tkinter + openpyxl + qrcode + Pillow**.

---

## Estructura del Proyecto

| Archivo / Carpeta | Descripción |
|---|---|
| [`conciliador_universal.py`](file:///c:/Users/Analista-TI/Desktop/entrega%20inicial/Scrip-Inventario/conciliador_universal.py) | Aplicación principal v3.0 con 2 pestañas (Conciliación + Módulo QR) |
| [`ejecutar.bat`](file:///c:/Users/Analista-TI/Desktop/entrega%20inicial/Scrip-Inventario/ejecutar.bat) | Lanzador directo con doble clic para Windows |
| `etiquetas_qr/` | Carpeta automática de salida para PNGs individuales y planillas de impresión |
| [`requirements.txt`](file:///c:/Users/Analista-TI/Desktop/entrega%20inicial/Scrip-Inventario/requirements.txt) | Dependencias de Python (`openpyxl`, `qrcode[pil]`, `Pillow`, `pandas`) |
| `scrip1.py` | Versión v1.0 histórica (conservada como referencia) |

---

## Instalación de Requisitos

```powershell
pip install -r requirements.txt
```

---

## Módulos y Funcionalidades

### 1. Pestaña: Conciliación de Inventario (Auditoría)
- **Selector de Hoja con Dropdown Automático:** Detecta y lista al instante las hojas disponibles del archivo `.xlsx` cargado, preseleccionando automáticamente la hoja de inventario.
- **Pegado Inteligente desde Excel:** Botón `PEGAR DESDE EXCEL (Ctrl+V)` para capturar datos tabulares directos desde el portapapeles.
- **Preset Físico Clemont Integrado:** Botón `Cargar Preset Clemont` con los 20 equipos verificados en sitio.
- **Motor de Similitud Tolerante:** Levenshtein con tolerancia a confusiones tipográficas frecuentes (`O`/`0`, `Z`/`2`, `S`/`5`, `I`/`1`, `B`/`8`), truncamientos y permutaciones.
- **Detección de Filas Duplicadas:** Marca celdas en amarillo e inserta comentarios nativos de Excel explicando el motivo específico (mismo serial o mismo host repetido).
- **Detección de Altas / Sobrantes:** Detecta equipos encontrados en campo que no estaban dados de alta en el Excel.
- **Secciones de Monitores y Tablets:** Incorporación automática con estilos empresariales después de la fila total.
- **Hoja "Reporte de Auditoría":** Resumen ejecutivo con autofiltro y códigos de color corporativos.
- **Fórmulas Reparadas Automáticamente:** Corrección dinámica de la fórmula `=SUM()` en la fila de totales.

### 2. Pestaña: Módulo QR (Rotulación e Impresión)
- **Taxonomía de Asset Tags:** Identificadores únicos inmutables `[PREFIJO]-[AÑO]-[CORRELATIVO]`:
  - `LAP-2026-0001` a `LAP-2026-0020` (Laptops / PCs)
  - `TAB-2026-0001` a `TAB-2026-0002` (Tablets)
  - `MON-2026-0001` a `MON-2026-0012` (Monitores)
- **Seguridad e Inmutabilidad del QR:** El código QR codifica **únicamente** el ID del activo (`LAP-2026-XXXX`). Esto garantiza que traslados de sede o cambios de usuario nunca desactualicen la etiqueta física.
- **Generación Masiva en Alta Resolución:** Genera los 34 archivos PNG individuales a **300 DPI**, listos para impresoras térmicas o corte manual.
- **Planillas de Impresión en Pliegos Tamaño Carta (300 DPI):**
  - Formato `50 x 25 mm`: 30 etiquetas por hoja (3 col x 10 filas) con guías de corte.
  - Formato `67 x 25 mm` (Avery 5160 estándar): 30 etiquetas por hoja.
  - Formato `100 x 50 mm` (Avery 5163): 10 etiquetas grandes por hoja (ideal para laptops).
- **Exportación de Plantilla Excel (`Inventario_Equipos_Clemont_Con_QR.xlsx`):**
  - Hoja `Maestro_Activos`: Catálogo oficial de referencia con todos los IDs asignados.
  - Hoja `Escaneos_Campo`: Plantilla lista para capturar lecturas con lector de código de barras 2D o celular.
  - Hoja `Instructivo_Auditoria`: Guía paso a paso de operación.
- **Simulador / Verificador en Vivo:** Campo de consulta rápida en tiempo real para verificar cualquier ID o serial escaneado.
