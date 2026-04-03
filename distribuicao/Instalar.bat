@echo off
chcp 1252 >nul
title Abbas Tecnologia - Instalando NF-e Destinadas
color 0A

echo.
echo  ABBAS Tecnologia - NF-e Destinadas
echo  Instalando...
echo.

:: Pasta de instalacao
set DEST=%LOCALAPPDATA%\AbbasTI\NFe
if not exist "%DEST%" mkdir "%DEST%"
if not exist "%DEST%\xml"        mkdir "%DEST%\xml"
if not exist "%DEST%\data"       mkdir "%DEST%\data"
if not exist "%DEST%\data\notas" mkdir "%DEST%\data\notas"
if not exist "%DEST%\logs"       mkdir "%DEST%\logs"

:: Copia o .exe
copy /Y "%~dp0ConsultaNFe.exe" "%DEST%\ConsultaNFe.exe" >nul
echo  Arquivos copiados para: %DEST%

:: Cria atalho na area de trabalho via VBScript
set VBS=%TEMP%\atalho_nfe.vbs
set LINK=%USERPROFILE%\Desktop\NF-e Destinadas (Abbas).lnk
(
echo Set oWS = WScript.CreateObject("WScript.Shell"^)
echo Set oLink = oWS.CreateShortcut("%LINK%"^)
echo oLink.TargetPath = "%DEST%\ConsultaNFe.exe"
echo oLink.WorkingDirectory = "%DEST%"
echo oLink.Description = "NF-e Destinadas - Abbas Tecnologia"
echo oLink.WindowStyle = 1
echo oLink.Save
) > "%VBS%"
cscript //nologo "%VBS%" >nul 2>&1
del "%VBS%" >nul 2>&1
echo  Atalho criado na area de trabalho.

echo.
echo  Instalacao concluida!
echo  Use o atalho "NF-e Destinadas (Abbas)" na area de trabalho.
echo.

set /p ABRIR=Abrir agora? (S/N): 
if /i "%ABRIR%"=="S" start "" "%DEST%\ConsultaNFe.exe"

timeout /t 3 >nul
