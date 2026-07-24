@echo off
setlocal
REM ===================================================================
REM  Desinstaller_clic_droit.bat
REM  Retire l'entree de menu clic droit "Convertir..." (registre HKCU).
REM ===================================================================

for %%E in (.sldprt .sldasm .slddrw .CATPart .CATProduct .CATDrawing .ipt .iam .idw .prt .asm .drw) do (
    reg delete "HKCU\Software\Classes\SystemFileAssociations\%%E\shell\ConvertirCAO" /f >nul 2>nul
)

echo.
echo   Menu clic droit retire.
echo.
pause
