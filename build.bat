@echo off
setlocal
echo ============================================
echo  Document Comparator - PyInstaller Build
echo ============================================
echo.

pip install -r requirements-dev.txt
if errorlevel 1 goto :error

echo.
echo Se construieste executabilul...
echo.

pyinstaller Comparer.spec --noconfirm --clean
if errorlevel 1 goto :error

if not exist "dist\Comparer.exe" goto :missing
echo.
echo  Build reusit!
echo  Executabilul se afla in: dist\Comparer.exe
exit /b 0

:missing
echo.
echo  Build ESUAT. Executabilul dist\Comparer.exe nu a fost creat.
exit /b 1

:error
echo.
echo  Build ESUAT. Verifica erorile de mai sus.
exit /b 1
