@echo off
cd /d "%~dp0"

if not exist "input_turni.xlsx" (
    echo ERRORE: input_turni.xlsx non trovato nella cartella.
    echo Copia input_turni_template.xlsx, rinominalo in input_turni.xlsx e compilalo.
    pause
    exit /b 1
)

if not exist "output" mkdir output

if exist "dist\GeneraTurni.exe" (
    dist\GeneraTurni.exe
) else (
    python turnazione_completa.py
)

if %errorlevel% neq 0 (
    echo.
    echo ERRORE durante l'esecuzione.
    pause
    exit /b %errorlevel%
)

echo.
echo Output generato in: output\turnazione_generata.xlsx
pause
