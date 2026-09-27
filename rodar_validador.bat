@echo off
chcp 65001 >nul
cd /d "%~dp0"
REM Voce pode arrastar uma pasta em cima deste arquivo para analisar direto.
set "PASTA=%~1"
if "%PASTA%"=="" set /p PASTA="Cole o caminho da pasta com os documentos e aperte Enter: "
python validador_legibilidade.py "%PASTA%"
echo.
pause
