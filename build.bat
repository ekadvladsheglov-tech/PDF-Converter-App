@echo off
title PDF Setup Builder v2.0
color 0E

echo ==========================================
echo PDF CONVERTER SETUP BUILDER v2.0
echo ==========================================
echo.

:: ПРОВЕРКИ
where python >nul 2>&1
if %errorlevel% neq 0 (
    color 0C & echo ERROR: Python not found! & pause & exit /b
)

set "INNOPATH=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist "%INNOPATH%" set "INNOPATH=C:\Program Files (x86)\Inno Setup 5\ISCC.exe"
if not exist "%INNOPATH%" (
    color 0C & echo ERROR: Inno Setup not found! & pause & exit /b
)

if not exist "setup_script.iss" (
    color 0C & echo ERROR: setup_script.iss not found! & pause & exit /b
)

echo [OK] All checks passed.
echo.

:: ШАГ 1: ОЧИСТКА
echo [Step 1/7] Cleaning...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "venv" rmdir /s /q venv
if exist "temp_setup" rmdir /s /q temp_setup
if exist "Output" rmdir /s /q Output

:: ШАГ 2: VENV
echo [Step 2/7] Creating VENV...
python -m venv venv
call venv\Scripts\activate.bat

:: ШАГ 3: БИБЛИОТЕКИ (добавлены Pillow, PyMuPDF, tkinterdnd2)
echo [Step 3/7] Installing libraries...
python -m pip install --upgrade pip >nul 2>&1
python -m pip install customtkinter pdf2docx python-docx PyPDF2 reportlab Pillow PyMuPDF pyinstaller >nul 2>&1
python -m pip install tkinterdnd2 >nul 2>&1
python -m pip install "numpy<2.0" >nul 2>&1

:: ШАГ 4: СБОРКА EXE
echo [Step 4/7] Building EXE...
pyinstaller --noconsole --onefile --name "PDF_Converter" --collect-all numpy --collect-all tkinterdnd2 --add-data "DejaVuSans.ttf;." pdf_converter.py
if not exist "dist\PDF_Converter.exe" (
    color 0C & echo ERROR: PyInstaller failed! & pause & exit /b
)

:: ШАГ 5: ПОДГОТОВКА ФАЙЛОВ
echo [Step 5/7] Preparing files...
mkdir temp_setup >nul 2>&1
copy /Y "dist\PDF_Converter.exe" "temp_setup\" >nul
if exist "DejaVuSans.ttf" copy /Y "DejaVuSans.ttf" "temp_setup\" >nul

:: ШАГ 6: КОМПИЛЯЦИЯ УСТАНОВЩИКА
echo [Step 6/7] Compiling Setup.exe...
"%INNOPATH%" setup_script.iss
if %errorlevel% neq 0 (
    color 0C
    echo ERROR: Inno Setup failed!
    pause
    exit /b
)

:: ШАГ 7: ПЕРЕНОС НА РАБОЧИЙ СТОЛ
echo [Step 7/7] Moving to Desktop...
if exist "Output\PDF_Converter_Setup_v1.0.0.exe" (
    copy /Y "Output\PDF_Converter_Setup_v1.0.0.exe" "%USERPROFILE%\Desktop\PDF_Converter_Setup.exe" >nul

    rmdir /s /q build
    rmdir /s /q dist
    rmdir /s /q venv
    rmdir /s /q temp_setup

    color 0A
    echo.
    echo ==========================================
    echo          SUCCESS!
    echo Setup.exe is on your Desktop!
    echo ==========================================
) else (
    color 0C & echo ERROR: Output folder is empty!
)

pause
exit /b