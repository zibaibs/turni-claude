@echo off
cd /d "%~dp0"

if not exist "input_turni.xlsx" (
    echo ERRORE: input_turni.xlsx non trovato nella cartella dello script.
    pause
    exit /b 1
)

if not exist "output" mkdir output

python turnazione_completa.py
if %errorlevel% neq 0 (
    echo.
    echo ERRORE durante l'esecuzione dello script.
    pause
    exit /b %errorlevel%
)

echo.
echo Output generato in: output\turnazione_generata.xlsx
pause
