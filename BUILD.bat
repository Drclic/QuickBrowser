@echo off
echo ============================================
echo   PDF Quick Browser v3 - Build
echo ============================================
echo.

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERREUR] Python n'est pas installe ou pas dans le PATH.
    pause
    exit /b 1
)

echo [1/3] Creation du venv et installation...
python -m venv venv
call venv\Scripts\activate.bat
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERREUR] pip install a echoue.
    pause
    exit /b 1
)

echo.
echo    Pour tester : venv\Scripts\python pdf_quick_browser.py
echo    Appuyez sur une touche pour builder...
pause

echo.
echo [2/3] Build de l'executable...
pyinstaller pdf_quick_browser.spec --clean
if %errorlevel% neq 0 (
    echo [ERREUR] Le build a echoue.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   BUILD TERMINE !
echo   Executable : dist\PDF-Quick-Browser.exe
echo ============================================
pause
