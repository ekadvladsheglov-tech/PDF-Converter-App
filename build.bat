@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"
title PDF Setup Builder v3.1
color 0E

echo ==========================================
echo PDF CONVERTER SETUP BUILDER v3.1
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

if not exist "app_web.py" (
    color 0C & echo ERROR: app_web.py not found! & pause & exit /b
)

if not exist "DejaVuSans.ttf" (
    color 0E & echo WARNING: DejaVuSans.ttf not found! Cyrillic in Word-to-PDF may break.
    echo Press any key to continue without it... & pause >nul
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
if exist "pip_log.txt" del /q pip_log.txt
if exist "build_log.txt" del /q build_log.txt
if exist "inno_log.txt" del /q inno_log.txt

:: ШАГ 2: VENV
echo [Step 2/7] Creating VENV...
python -m venv venv
if %errorlevel% neq 0 (
    color 0C & echo ERROR: Failed to create VENV! & pause & exit /b
)
call venv\Scripts\activate.bat

:: ШАГ 3: БИБЛИОТЕКИ + ИКОНКА
echo [Step 3/7] Installing libraries...
python -m pip install --upgrade pip > pip_log.txt 2>&1
python -m pip install -r requirements.txt >> pip_log.txt 2>&1
if %errorlevel% neq 0 (
    color 0C & echo ERROR: pip install failed! See pip_log.txt for details. & pause & exit /b
)
if not exist "pdf.ico" (
    echo [Step 3/7] Generating pdf.ico...
    python make_icon.py
)

:: ШАГ 4: СБОРКА EXE
echo [Step 4/7] Building EXE...
if exist "pdf.ico" (
    pyinstaller --clean --noconsole --onefile --name "PDF_Converter" --collect-all pywebview --add-data "DejaVuSans.ttf;." --add-data "pdf.ico;." --icon "pdf.ico" app_web.py > build_log.txt 2>&1
) else (
    pyinstaller --clean --noconsole --onefile --name "PDF_Converter" --collect-all pywebview --add-data "DejaVuSans.ttf;." app_web.py > build_log.txt 2>&1
)
if %errorlevel% neq 0 (
    color 0C & echo ERROR: PyInstaller failed! See build_log.txt for details. & pause & exit /b
)
if not exist "dist\PDF_Converter.exe" (
    color 0C & echo ERROR: PDF_Converter.exe not created! See build_log.txt & pause & exit /b
)

:: ШАГ 5: ПОДГОТОВКА ФАЙЛОВ
echo [Step 5/7] Preparing files...
mkdir temp_setup >nul 2>&1
copy /Y "dist\PDF_Converter.exe" "temp_setup\" >nul

:: ШАГ 6: КОМПИЛЯЦИЯ УСТАНОВЩИКА
echo [Step 6/7] Compiling Setup.exe...
"%INNOPATH%" setup_script.iss > inno_log.txt 2>&1
if %errorlevel% neq 0 (
    color 0C & echo ERROR: Inno Setup failed! See inno_log.txt & pause & exit /b
)

:: ШАГ 7: ПЕРЕНОС НА РАБОЧИЙ СТОЛ
echo [Step 7/7] Moving to Desktop...
set "SETUPFILE="
for %%f in ("Output\*.exe") do set "SETUPFILE=%%f"

if defined SETUPFILE (
    copy /Y "!SETUPFILE!" "%USERPROFILE%\Desktop\PDF_Converter_Setup.exe" >nul

    rmdir /s /q build
    rmdir /s /q dist
    rmdir /s /q venv
    rmdir /s /q temp_setup
    if exist "pip_log.txt" del /q pip_log.txt
    if exist "build_log.txt" del /q build_log.txt
    if exist "inno_log.txt" del /q inno_log.txt

    color 0A
    echo.
    echo ==========================================
    echo          SUCCESS!
    echo Setup.exe is on your Desktop!
    echo ==========================================
) else (
    color 0C & echo ERROR: No setup file found in Output folder! Check inno_log.txt
)

pause
exit /b