@echo off
setlocal enabledelayedexpansion
title Procesador de Albaranes a CSV

if "%~1"=="" (
    echo No has arrastrado ningun archivo o carpeta.
    echo Procesando la carpeta "muestras" por defecto...
    python extractor.py muestras --csv "albaranes_extraidos.csv" --sep ";"
) else (
    echo Procesando: %1
    python extractor.py "%~1" --csv "albaranes_extraidos.csv" --sep ";"
)

echo.
echo ====================================================
echo Proceso finalizado. El archivo CSV generado es:
echo albaranes_extraidos.csv
echo ====================================================
echo.
pause
