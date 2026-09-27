@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Instalador - Validador de Documentos
echo ======================================================
echo    INSTALADOR - VALIDADOR DE DOCUMENTOS
echo ======================================================
echo.

REM ---------- 1. Procurar o Python ----------
REM Prefere versoes estaveis: Python muito novo (ex.: 3.15) ainda nao tem
REM todas as bibliotecas prontas e a instalacao falha.
set "PY="
for %%v in (3.13 3.12 3.14 3.11 3.10) do (
  if not defined PY (
    py -%%v --version >nul 2>&1 && set "PY=py -%%v"
  )
)
if not defined PY python --version >nul 2>&1 && set "PY=python"
if not defined PY (
  echo [ERRO] Python nao encontrado neste computador.
  echo        Instale pelo site python.org marcando "Add Python to PATH"
  echo        e rode este instalador de novo.
  echo.
  pause
  exit /b 1
)
for /f "delims=" %%i in ('%PY% -c "import sys,os;print(os.path.join(os.path.dirname(sys.executable),'pythonw.exe'))"') do set "PYW=%%i"
if not exist "%PYW%" (
  echo [ERRO] Nao encontrei o pythonw.exe do Python.
  pause
  exit /b 1
)
for /f "delims=" %%i in ('%PY% -c "import sys;print(sys.version.split()[0])"') do set "PYVER=%%i"
echo [OK] Usando Python %PYVER%: %PYW%
echo.

REM ---------- 2. Instalar as bibliotecas ----------
echo Instalando as bibliotecas (pode levar alguns minutos)...
%PY% -m pip install --disable-pip-version-check -q -r requirements.txt
if errorlevel 1 (
  echo [AVISO] Algumas bibliotecas nao foram instaladas.
  echo         No computador do trabalho, a rede da empresa pode bloquear o download.
  echo         Se o programa nao abrir, peca ajuda ao TI para liberar o pip.
) else (
  echo [OK] Bibliotecas instaladas.
)
echo.

REM ---------- 3. Conferir componente de janelas ----------
%PY% -c "import tkinter" >nul 2>&1
if errorlevel 1 (
  echo [ERRO] O Python foi instalado sem o componente de janelas.
  echo        Rode o instalador do Python de novo, escolha "Modify"
  echo        e marque "tcl/tk and IDLE".
  echo.
  pause
  exit /b 1
)

REM ---------- 4. Conferir se o programa carrega ----------
%PY% -c "import sys; sys.path.insert(0, r'%~dp0.'); import validador_legibilidade" 2>nul
if errorlevel 1 (
  echo [AVISO] O programa nao conseguiu carregar todas as bibliotecas.
  echo         Rode de novo este instalador. Se continuar, mande um print ao responsavel.
  echo.
)

REM ---------- 5. Criar o icone na Area de Trabalho ----------
set "ALVO=%~dp0Validador.pyw"
set "PASTA=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$w=New-Object -ComObject WScript.Shell; $d=[Environment]::GetFolderPath('Desktop'); $a=$w.CreateShortcut((Join-Path $d 'Validador de Documentos.lnk')); $a.TargetPath=$env:PYW; $a.Arguments=[char]34+$env:ALVO+[char]34; $a.WorkingDirectory=$env:PASTA; $a.IconLocation=$env:SystemRoot+'\System32\imageres.dll,108'; $a.Save()" >nul 2>&1

if not errorlevel 1 goto atalho_ok

REM Plano B: se o PowerShell estiver bloqueado, cria um lancador simples
for /f "delims=" %%d in ('%PY% -c "import os;p=os.path.join(os.environ['USERPROFILE'],'OneDrive','Desktop');print(p if os.path.isdir(p) else os.path.join(os.environ['USERPROFILE'],'Desktop'))"') do set "DESK=%%d"
> "%DESK%\Validador de Documentos.bat" echo @echo off
>> "%DESK%\Validador de Documentos.bat" echo start "" "%PYW%" "%ALVO%"
echo [OK] Icone criado na Area de Trabalho (modo alternativo).
goto fim

:atalho_ok
echo [OK] Icone "Validador de Documentos" criado na Area de Trabalho!

:fim
echo.
echo ======================================================
echo    PRONTO! Abra o icone "Validador de Documentos"
echo    na sua Area de Trabalho.
echo ======================================================
echo.
pause
