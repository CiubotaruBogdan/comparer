@echo off
echo ============================================
echo  Document Comparator - PyInstaller Build
echo ============================================
echo.

:: Instaleaza dependentele daca nu sunt
pip install -r requirements.txt

echo.
echo Se construieste executabilul...
echo.

pyinstaller ^
    --noconfirm ^
    --onefile ^
    --windowed ^
    --name "Comparer" ^
    --icon "icon.ico" ^
    --add-data "icon.ico;." ^
    --add-data "icon.png;." ^
    --collect-all fitz ^
    --collect-all docx ^
    --hidden-import PyQt5 ^
    --hidden-import PyQt5.QtWidgets ^
    --hidden-import PyQt5.QtCore ^
    --hidden-import PyQt5.QtGui ^
    --hidden-import PyQt5.QtPrintSupport ^
    --hidden-import difflib ^
    --hidden-import html ^
    --hidden-import pathlib ^
    main.py

echo.
if exist "dist\Comparer.exe" (
    echo  Build reusit!
    echo  Executabilul se afla in: dist\Comparer.exe
) else (
    echo  Build ESUAT. Verifica erorile de mai sus.
)

echo.
pause
