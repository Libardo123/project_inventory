@echo off
title Conciliador de Inventario Clemont v3.0 (Conciliacion + Modulo QR)
cd /d "%~dp0"

echo ========================================================
echo   Iniciando Conciliador de Inventario Clemont v3.0
echo   - Modulo 1: Conciliacion y Auditoria Universal
echo   - Modulo 2: Generador de Etiquetas QR y Planillas
echo ========================================================
echo.

python conciliador_universal.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Ocurrio un error al ejecutar el script.
    pause
)
