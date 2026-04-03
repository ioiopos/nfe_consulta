@echo off
chcp 1252 >nul
title Abbas Tecnologia - Build NF-e Destinadas
color 0A

echo.
echo  ================================================
echo   ABBAS Tecnologia - Build do Executavel
echo   Gera ConsultaNFe.exe com Python embutido
echo  ================================================
echo.

:: Detecta Python
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 ( set PYTHON=py ) else ( set PYTHON=python )

:: Instala PyInstaller se nao tiver
echo  [1/4] Verificando PyInstaller...
%PYTHON% -m PyInstaller --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo  Instalando PyInstaller...
    %PYTHON% -m pip install pyinstaller -q
)
echo  OK
echo.

:: Instala dependencias do projeto
echo  [2/4] Instalando dependencias...
%PYTHON% -m pip install -r requirements.txt -q
%PYTHON% -m pip install pywin32 -q
echo  OK
echo.

:: Gera o .exe
echo  [3/4] Compilando executavel (pode demorar 2-5 minutos)...
echo.

%PYTHON% -m PyInstaller ^
    --noconsole ^
    --onefile ^
    --name "ConsultaNFe" ^
    --add-data "src;src" ^
    --hidden-import "win32com.client" ^
    --hidden-import "win32com.shell" ^
    --hidden-import "pythoncom" ^
    --hidden-import "pywintypes" ^
    --hidden-import "win32api" ^
    --hidden-import "cryptography" ^
    --hidden-import "cryptography.hazmat.primitives.serialization.pkcs12" ^
    --hidden-import "signxml" ^
    --hidden-import "lxml" ^
    --hidden-import "lxml.etree" ^
    --collect-all "signxml" ^
    --collect-all "lxml" ^
    main.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  ERRO na compilacao! Verifique as mensagens acima.
    pause
    exit /b 1
)

echo.
echo  OK - Executavel gerado em: dist\ConsultaNFe.exe
echo.

:: Cria o instalador final que distribui para clientes
echo  [4/4] Gerando pacote de distribuicao...

if not exist "distribuicao" mkdir distribuicao
copy /Y "dist\ConsultaNFe.exe" "distribuicao\ConsultaNFe.exe" >nul
copy /Y "instalar_cliente.bat"  "distribuicao\Instalar.bat"    >nul 2>&1

echo.
echo  ================================================
echo   BUILD CONCLUIDO!
echo.
echo   Arquivo gerado: dist\ConsultaNFe.exe
echo   Tamanho: 
for %%F in ("dist\ConsultaNFe.exe") do echo    %%~zF bytes
echo.
echo   Para distribuir ao cliente:
echo   Envie apenas: ConsultaNFe.exe
echo   O cliente da duplo clique e pronto!
echo  ================================================
echo.
pause
