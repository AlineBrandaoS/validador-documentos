@echo off
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$w=New-Object -ComObject WScript.Shell; $a=$w.CreateShortcut([Environment]::GetFolderPath('Desktop')+'\Validador de Documentos.lnk'); $a.TargetPath='%~dp0Validador.pyw'; $a.WorkingDirectory='%~dp0'; $a.IconLocation='%SystemRoot%\System32\imageres.dll,108'; $a.Save()"
if errorlevel 1 (
  echo Nao foi possivel criar o atalho.
) else (
  echo Atalho "Validador de Documentos" criado na Area de Trabalho!
)
pause
