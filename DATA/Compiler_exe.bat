@echo off
REM ===================================================================
REM  Compiler_exe.bat
REM  Fabrique un executable unique « Convert-Rename.exe » (aucun Python
REM  ni pywin32 requis chez les collegues ; le logiciel CAO reste requis).
REM
REM  A lancer UNE FOIS sur une machine Windows avec Python installe.
REM  Resultat : DATA\dist\Convert-Rename.exe
REM ===================================================================
setlocal
cd /d "%~dp0"

echo Verification de PyInstaller...
py -3 -m PyInstaller --version >nul 2>nul
if errorlevel 1 (
    echo Installation de PyInstaller...
    py -3 -m pip install pyinstaller || (echo Echec de l'installation de PyInstaller. & pause & exit /b 1)
)

echo.
echo Compilation en cours (peut prendre 1 a 2 minutes)...
py -3 -m PyInstaller ^
  --noconfirm --clean --onefile --windowed ^
  --name "Convert-Rename" ^
  --icon "icone.ico" ^
  --hidden-import win32timezone ^
  --hidden-import win32com ^
  --hidden-import win32com.client ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --hidden-import convertisseur_gui ^
  convertisseur_cao.py

if errorlevel 1 (
    echo.
    echo La compilation a echoue. Copiez le message d'erreur ci-dessus.
    pause
    exit /b 1
)

echo.
echo ====================================================================
echo  OK : l'executable est ici -> DATA\dist\Convert-Rename.exe
echo  Vous pouvez le copier ou vous voulez et le distribuer tel quel.
echo ====================================================================
pause
