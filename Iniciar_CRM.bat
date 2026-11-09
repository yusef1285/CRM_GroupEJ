@echo off
title CRM Group EJ Morales
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
    echo No se encontro Python. Instale Python 3.10 o superior desde https://www.python.org/downloads/
    echo Importante: marque la casilla "Add Python to PATH" durante la instalacion.
    pause
    exit /b
)
python CRM_GroupEJ.py
if errorlevel 1 pause
