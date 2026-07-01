@echo off
title Lanzador - Dashboard de Auditoría de Médicos Supernumerarios
echo ====================================================================
echo   🏥 Lanzador del Dashboard de Auditoría de Médicos Supernumerarios
echo ====================================================================
echo.

:: Verificar si el lanzador py.exe existe
where py >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] No se pudo encontrar el lanzador 'py' de Python.
    echo Por favor, asegúrate de tener instalado Python en tu equipo.
    echo Puedes descargarlo de: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] Verificando dependencias necesarias de Python...
py -m pip install pandas openpyxl streamlit xlsxwriter python-calamine --quiet
if %errorlevel% neq 0 (
    echo [WARNING] Hubo un problema instalando dependencias. 
    echo Intentando continuar con la ejecución...
) else (
    echo [OK] Dependencias instaladas/verificadas con éxito.
)

echo.
echo [2/3] Iniciando el servidor local de Streamlit...
echo La aplicación web se abrirá automáticamente en tu navegador web predeterminado.
echo Para cerrar la aplicación, simplemente cierra esta ventana negra.
echo.

:: Ejecutar Streamlit
py -m streamlit run app.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] La aplicación web se detuvo de manera inesperada o se presionó Ctrl+C.
    pause
)
