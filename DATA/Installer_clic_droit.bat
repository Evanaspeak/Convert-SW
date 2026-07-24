@echo off
setlocal enabledelayedexpansion
REM ===================================================================
REM  Installer_clic_droit.bat
REM  Ajoute l'entree de menu clic droit "Convertir..." pour les fichiers
REM  CAO/DAO, en utilisant AUTOMATIQUEMENT le chemin de ce dossier.
REM  Aucun chemin a modifier, aucun droit administrateur requis
REM  (ecrit dans le registre de l'utilisateur : HKCU).
REM ===================================================================

set "SCRIPT=%~dp0convertisseur_cao.py"

if not exist "%SCRIPT%" (
    echo.
    echo   Introuvable : "%SCRIPT%"
    echo   Ce fichier doit rester dans le dossier DATA, a cote de
    echo   convertisseur_cao.py.
    echo.
    pause
    exit /b 1
)

REM --- Pieces / assemblages -> STEP + STL ---
for %%E in (.sldprt .sldasm .CATPart .CATProduct .ipt .iam .prt .asm) do (
    reg add "HKCU\Software\Classes\SystemFileAssociations\%%E\shell\ConvertirCAO" /ve /d "Convertir (Convert-Rename)" /f >nul
    reg add "HKCU\Software\Classes\SystemFileAssociations\%%E\shell\ConvertirCAO\command" /ve /d "pyw \"%SCRIPT%\" \"%%1\"" /f >nul
)

REM --- Mises en plan -> DWG + DXF + PDF ---
for %%E in (.slddrw .CATDrawing .idw .drw) do (
    reg add "HKCU\Software\Classes\SystemFileAssociations\%%E\shell\ConvertirCAO" /ve /d "Convertir (Convert-Rename)" /f >nul
    reg add "HKCU\Software\Classes\SystemFileAssociations\%%E\shell\ConvertirCAO\command" /ve /d "pyw \"%SCRIPT%\" \"%%1\"" /f >nul
)

echo.
echo   Menu clic droit installe.
echo   Faites un clic droit sur un fichier CAO puis "Convertir...".
echo   (Pour retirer : lancez Desinstaller_clic_droit.bat)
echo.
pause
