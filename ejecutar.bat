@echo off
title Conciliador de Inventario Clemont
cd /d "%~dp0"

echo ========================================================
echo   Iniciando Conciliador de Inventario Clemont v2.0
echo ========================================================
echo.

python conciliador_universal.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Ocurrio un error al ejecutar el script.
    pause
)
