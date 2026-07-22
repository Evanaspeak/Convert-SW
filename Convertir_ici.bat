@echo off
REM ===================================================================
REM  Convertir_ici.bat
REM  Lanceur glisser-deposer du convertisseur CAO/DAO par lot.
REM
REM  Utilisation :
REM    - Selectionnez un ou plusieurs fichiers CAO (.sldprt, .CATPart,
REM      .ipt, .prt, mises en plan, etc.)
REM    - Faites-les glisser sur cette icone.
REM  Les exports (STEP+STL ou DWG+DXF) apparaissent dans un sous-dossier
REM  « Export » cree a cote de chaque fichier source.
REM ===================================================================

setlocal

REM Se placer dans le dossier du script (contient convertisseur_cao.py)
cd /d "%~dp0"

if "%~1"=="" (
    echo.
    echo   Aucun fichier fourni.
    echo   Glissez-deposez vos fichiers CAO sur ce .bat, ou lancez :
    echo       Convertir_ici.bat  fichier1.sldprt  fichier2.CATPart ...
    echo.
    pause
    exit /b 1
)

REM On tente « py » (lanceur Python officiel), puis « python » en repli.
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 convertisseur_cao.py %*
) else (
    python convertisseur_cao.py %*
)

echo.
echo   Conversion terminee. Fenetre laissee ouverte pour lecture du journal.
pause
endlocal
