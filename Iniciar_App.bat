@echo off
title Albarán a CSV Pro
echo ====================================================
echo      INICIANDO ALBARAN A CSV PRO
echo ====================================================
echo.
echo Iniciando servidor local en el puerto 8080...
start "" http://localhost:8080
python servidor.py 8080
pause
