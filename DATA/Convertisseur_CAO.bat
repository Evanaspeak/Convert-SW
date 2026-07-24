@echo off
REM ===================================================================
REM  Convertisseur_CAO.bat
REM  Ouvre l'application AVEC une console visible (utile pour voir une
REM  erreur de demarrage). Pour un lancement propre sans console, utilisez
REM  Convertisseur_CAO.vbs a la racine.
REM ===================================================================
cd /d "%~dp0"

py -3 convertisseur_cao.py %*

echo.
echo (Fenetre laissee ouverte. Fermez-la quand vous avez fini.)
pause
