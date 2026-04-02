@echo off
title NF-e Destinadas - SEFAZ
echo.
echo  ============================================
echo   NF-e Destinadas - Consulta SEFAZ
echo  ============================================
echo.

python main.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  ERRO ao iniciar. Verifique se o Python esta instalado
    echo  e se as dependencias foram instaladas com:
    echo.
    echo    pip install -r requirements.txt
    echo.
    pause
)
