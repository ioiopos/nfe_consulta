@echo off
chcp 1252 >nul
title Abbas Tecnologia - Instalador NF-e Destinadas
color 0A

echo.
echo  ================================================
echo   ABBAS Tecnologia - NF-e Destinadas
echo   Instalador Automatico
echo  ================================================
echo.

:: Verifica Python
echo  [1/6] Verificando Python...
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    py --version >nul 2>&1
    if %ERRORLEVEL% NEQ 0 (
        echo.
        echo  ERRO: Python nao encontrado!
        echo.
        echo  Instale o Python em: https://www.python.org/downloads/
        echo  IMPORTANTE: marque "Add Python to PATH" durante a instalacao.
        echo.
        pause
        exit /b 1
    )
    set PYTHON=py
) else (
    set PYTHON=python
)
for /f "tokens=*" %%v in ('%PYTHON% --version 2^>^&1') do set PY_VER=%%v
echo  OK - %PY_VER%
echo.

:: Define pasta de instalacao
set INSTALL_DIR=%USERPROFILE%\AppData\Local\AbbasTI\NFe
echo  [2/6] Criando pastas em:
echo  %INSTALL_DIR%

if not exist "%INSTALL_DIR%"             mkdir "%INSTALL_DIR%"
if not exist "%INSTALL_DIR%\src"         mkdir "%INSTALL_DIR%\src"
if not exist "%INSTALL_DIR%\data"        mkdir "%INSTALL_DIR%\data"
if not exist "%INSTALL_DIR%\data\notas"  mkdir "%INSTALL_DIR%\data\notas"
if not exist "%INSTALL_DIR%\logs"        mkdir "%INSTALL_DIR%\logs"
if not exist "%INSTALL_DIR%\xml"         mkdir "%INSTALL_DIR%\xml"
if not exist "%INSTALL_DIR%\certs"       mkdir "%INSTALL_DIR%\certs"
echo  OK
echo.

:: Copia arquivos
echo  [3/6] Copiando arquivos...
set SCRIPT_DIR=%~dp0

xcopy /E /I /Y "%SCRIPT_DIR%src"          "%INSTALL_DIR%\src"  >nul 2>&1
copy  /Y        "%SCRIPT_DIR%main.py"      "%INSTALL_DIR%\main.py" >nul 2>&1
copy  /Y        "%SCRIPT_DIR%requirements.txt" "%INSTALL_DIR%\requirements.txt" >nul 2>&1
echo  OK
echo.

:: Instala dependencias
echo  [4/6] Instalando dependencias Python...
echo  (pode demorar alguns minutos na primeira vez)
echo.
%PYTHON% -m pip install --upgrade pip -q
%PYTHON% -m pip install -r "%INSTALL_DIR%\requirements.txt" -q
if %ERRORLEVEL% NEQ 0 (
    echo  AVISO: Algumas dependencias podem ter falhado.
    echo  Tente manualmente: pip install -r requirements.txt
)
echo  OK
echo.

:: pywin32
echo  [5/6] Configurando suporte a certificados Windows...
%PYTHON% -m pip install pywin32 -q
%PYTHON% -m pywin32_postinstall -install >nul 2>&1
echo  OK
echo.

:: Atalho na area de trabalho
echo  [6/6] Criando atalho na area de trabalho...

set SHORTCUT=%USERPROFILE%\Desktop\NF-e Destinadas (Abbas).lnk
set VBS=%TEMP%\atalho_nfe.vbs

(
echo Set oWS = WScript.CreateObject("WScript.Shell"^)
echo Set oLink = oWS.CreateShortcut("%SHORTCUT%"^)
echo oLink.TargetPath = "%PYTHON%"
echo oLink.Arguments = "main.py"
echo oLink.WorkingDirectory = "%INSTALL_DIR%"
echo oLink.Description = "NF-e Destinadas - Abbas Tecnologia"
echo oLink.WindowStyle = 1
echo oLink.Save
) > "%VBS%"

cscript //nologo "%VBS%" >nul 2>&1
del "%VBS%" >nul 2>&1
echo  OK
echo.

echo  ================================================
echo   Instalacao concluida com sucesso!
echo.
echo   Abrir o sistema:
echo   - Duplo clique em "NF-e Destinadas (Abbas)"
echo     na area de trabalho
echo   OU
echo   - Execute: python main.py
echo     na pasta: %INSTALL_DIR%
echo.
echo   XMLs salvos em: %INSTALL_DIR%\xml
echo  ================================================
echo.

set /p ABRIR=Deseja abrir o sistema agora? (S/N): 
if /i "%ABRIR%"=="S" (
    cd /d "%INSTALL_DIR%"
    start "" %PYTHON% main.py
)

echo.
echo  Obrigado por usar Abbas Tecnologia!
timeout /t 4 >nul
