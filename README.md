# Inventario de Equipos - Clemont

Herramienta de auditoría y conciliación de inventario de equipos en Python + Tkinter.

## Archivos

| Archivo | Descripción |
|---|---|
| `conciliador_universal.py` | Script principal v2.0 (usar este) |
| `scrip1.py` | Versión anterior (referencia) |
| `Inventario_Equipos_Clemont (2) (2).xlsx` | Excel de inventario Clemont |

## Requisitos

```
pip install openpyxl
```

## Uso

1. Ejecutar: `python conciliador_universal.py`
2. Seleccionar el archivo Excel
3. Elegir la hoja objetivo (`Inventario de Equipos (2)` para pruebas)
4. Copiar el rango de equipos verificados en Excel → **Pegar desde Excel** en la app
5. Ejecutar conciliación
6. Se genera `AUDITADO_*.xlsx` con el resultado

## Cómo pegar equipos verificados

En tu hoja de verificaciones en Excel, selecciona 4 columnas:

```
Equipo/Host  |  Serial Real  |  Estado (OK/Traslado)  |  Sede
```

Haz `Ctrl+C` → en la app clic en **PEGAR DESDE EXCEL**.

