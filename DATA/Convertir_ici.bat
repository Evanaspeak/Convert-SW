@echo off
REM ===================================================================
REM  Convertir_ici.bat
REM  Glissez-deposez des fichiers CAO sur cette icone : la fenetre de
REM  conversion s'ouvre deja prechargee avec ces fichiers.
REM  (Equivalent au glisser-deposer sur Convertisseur_CAO.vbs.)
REM ===================================================================
cd /d "%~dp0"

where pyw >nul 2>nul
if %errorlevel%==0 (
    start "" pyw convertisseur_cao.py %*
) else (
    start "" pythonw convertisseur_cao.py %*
)
