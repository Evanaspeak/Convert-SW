@echo off
REM ===================================================================
REM  Convertisseur_CAO.bat
REM  Ouvre l'interface graphique du convertisseur (double-clic).
REM  Aucune ligne de commande a taper.
REM ===================================================================
cd /d "%~dp0"

REM pyw / pythonw = Python "fenetre" (sans console noire).
where pyw >nul 2>nul
if %errorlevel%==0 (
    start "" pyw convertisseur_gui.py
    goto :eof
)
where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw convertisseur_gui.py
    goto :eof
)

REM Repli : Python normal (une console s'ouvrira, utile pour voir une erreur).
py -3 convertisseur_gui.py
